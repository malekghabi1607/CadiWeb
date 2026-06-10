from __future__ import annotations

import asyncio
import os
import time
from datetime import datetime
from pathlib import Path
from uuid import uuid4
from urllib.parse import unquote, urlparse

from fastapi import APIRouter, HTTPException
from pydantic import BaseModel

from vte.core import config
from vte.domain.iris import IRIS, IRIS_natif
from vte.services import iris_dumps_services


# =============================================================================
# Routeur IRIS
# =============================================================================
#
# Ce module est la façade HTTP entre l'interface React et le code métier IRIS
# historique du package `vte`.
#
# Il ne contient volontairement pas la logique métier lourde :
# - la recherche / copie / archivage des dumps natifs est dans
#   `vte.services.iris_dumps_services` ;
# - la consolidation Excel est dans `vte.services.iris_services.IRIS_services`.
#
# Son rôle est de :
# - normaliser les demandes du front ;
# - exposer les fichiers trouvés sous une forme facile à afficher ;
# - transformer les exceptions métier/fichier en réponses API lisibles ;
# - limiter les actions sensibles, notamment l'ouverture de chemins locaux.
#
# Guide rapide du flux:
# 1. GET /dumps lit les dossiers configures et construit le tableau React.
# 2. POST /validate-structure verifie les fichiers avant un traitement long.
# 3. POST /traiter ou /traiter-manuel-chemins cree un job et rend un job_id.
# 4. GET /job/{job_id} permet au front de suivre la progression.
# 5. DELETE /job/{job_id} marque le job comme annule; le callback de progression
#    stoppe le traitement a la prochaine etape controlable.
# 
# Le module reste volontairement une facade HTTP: les regles de selection,
# archivage et consolidation restent dans `vte`.

router = APIRouter(prefix="/iris", tags=["IRIS"])

# Stocke les jobs en cours : { job_id: { status, progress, step, files, message, warning } }
# Reinitialise au redemarrage du serveur.
#
# C'est suffisant pour une API locale CADI: le front garde seulement le job_id
# dans sessionStorage pour reprendre apres refresh, mais on ne cherche pas ici a
# survivre a un redemarrage complet de FastAPI.
_jobs: dict[str, dict] = {}

# Les jobs termines restent consultables quelques minutes pour que le front
# puisse afficher le resultat apres la fin, puis ils sont purges de la memoire.
JOB_RETENTION_SECONDS = 30 * 60

# `asyncio.create_task` ne garde qu'une reference faible a la tache cree : si on
# ne conserve la reference nulle part ailleurs, le ramasse-miettes Python peut
# detruire (et donc annuler) la tache en plein traitement, typiquement quand le
# front navigue vers d'autres pages et declenche d'autres appels API. On garde
# donc une reference forte ici, et on la retire automatiquement a la fin.
_background_tasks: set[asyncio.Task] = set()


def _spawn_job_task(coro) -> asyncio.Task:
    """Lance une coroutine en tache de fond sans risque de garbage-collection prematuree."""
    task = asyncio.create_task(coro)
    _background_tasks.add(task)
    task.add_done_callback(_background_tasks.discard)
    return task


# =============================================================================
# Configuration fonctionnelle IRIS
# =============================================================================

EXPORTS_IRIS = {
    "R04110": "Sessions",
    "R0304": "Formations",
    "R04301": "Ventes",
    "R04500": "Inscriptions",
}

# Seuls ces exports sont recopiés depuis la GED dans le flux existant.
# Les autres restent traités depuis le dossier local configuré.
# Cette distinction vient du processus metier actuel: certains exports sont
# deposables automatiquement en GED, d'autres restent fournis manuellement/local.
GED_EXPORT_CODES = {"R04301", "R04500"}



# =============================================================================
# Modeles d'entree API
# =============================================================================

class UpdateDumpsRequest(BaseModel):
    """Demande de mise à jour des dumps pour une liste de codes IRIS."""

    # Codes choisis par le front. Les codes inconnus seront ignores par
    # `_normaliser_codes`, puis une liste vide deviendra une erreur 400.
    codes: list[str]


class ValidateStructureRequest(BaseModel):
    """Demande de vérification; si `codes` est vide, on vérifie tous les exports."""

    # None ou liste vide = verifier tous les exports connus.
    # Liste renseignee = verifier uniquement la selection utilisateur.
    codes: list[str] | None = None


class RunTreatmentRequest(BaseModel):
    """Paramètres transmis par le front pour lancer la consolidation IRIS."""

    # Codes IRIS selectionnes dans le tableau React.
    exports: list[str]
    # Reserve pour garder une API evolutive; le flux actuel produit du xlsx.
    format: str = "xlsx"
    # `auto` lit les fichiers depuis la configuration; `manual` passe par chemins.
    mode: str = "auto"


class OpenPathRequest(BaseModel):
    """Chemin ou URI `file://` que le backend doit ouvrir via Windows."""

    # Peut etre une URI file:// produite par `_path_to_uri` ou un chemin Windows
    # brut. `_uri_to_path` normalise les deux formats.
    path: str


class SelectManualFilesRequest(BaseModel):
    """Demande d'ouverture du selecteur Python local pour un export IRIS."""

    export: str


class RunManualPathsRequest(BaseModel):
    """Demande de traitement manuel a partir de chemins deja choisis cote back."""

    export: str
    paths: list[str]


# =============================================================================
# Fonctions utilitaires locales
# =============================================================================

def _normaliser_codes(codes: list[str]) -> list[str]:
    """Nettoie les codes envoyés par React et ignore ceux hors périmètre CADI."""
    # Bloc filtre :
    # - trim pour accepter les saisies/API avec espaces ;
    # - upper pour rendre le front insensible à la casse ;
    # - exclusion silencieuse des codes inconnus.
    return [code.strip().upper() for code in codes if code.strip().upper() in EXPORTS_IRIS]


def _format_datetime(timestamp: float | None) -> str:
    """Convertit une date fichier en libelle lisible par React."""
    # Bloc affichage :
    # React attend une chaîne courte et stable, pas un timestamp brut.
    if timestamp is None:
        return "-"
    return datetime.fromtimestamp(timestamp).strftime("%d/%m/%Y %H:%M")


def _path_to_uri(path: Path) -> str:
    """Construit une URI `file://` pour le front.

    Le front ne l'ouvre pas directement dans le navigateur : il la renvoie à
    `/open-path`, car les navigateurs bloquent souvent les liens `file://`.
    L'URI reste utile comme identifiant stable et comme info d'infobulle.
    """
    # Bloc conversion :
    # `Path.as_uri()` sait encoder les espaces et les caractères spéciaux.
    # En cas de chemin non URI-compatible, on renvoie le chemin texte brut.
    try:
        return path.resolve().as_uri()
    except ValueError:
        return str(path)


def _uri_to_path(path_or_uri: str) -> Path:
    """Convertit un chemin brut ou une URI `file://` en `Path` Windows."""
    # Bloc entrée front :
    # Le front peut envoyer soit un chemin Windows brut, soit l'URI générée par
    # `_path_to_uri`. On accepte les deux pour garder l'API robuste.
    if path_or_uri.startswith("file://"):
        parsed = urlparse(path_or_uri)
        decoded_path = unquote(parsed.path)
        if parsed.netloc:
            # Cas UNC : file://serveur/partage/dossier -> //serveur/partage/dossier
            return Path(f"//{parsed.netloc}{decoded_path}")
        # Cas local : file:///C:/... -> C:/...
        return Path(decoded_path.lstrip("/"))
    return Path(path_or_uri)


def _is_relative_to(path: Path, parent: Path) -> bool:
    """Compatibilité simple autour de `Path.relative_to` avec chemins résolus."""
    # Bloc sécurité :
    # On résout les chemins avant comparaison pour éviter les contournements par
    # segments `..` ou chemins équivalents écrits différemment.
    try:
        path.resolve().relative_to(parent.resolve())
        return True
    except ValueError:
        return False


def _chemin_ouvrable(path: Path) -> bool:
    """Autorise l'ouverture seulement dans les dossiers IRIS exposés.

    `os.startfile` ouvre réellement un fichier sur le poste Windows. On évite
    donc de transformer cette route en lanceur arbitraire de chemins système.
    """
    # Bloc whitelist :
    # L'API n'ouvre que les fichiers/dossiers qui appartiennent aux deux zones
    # métier IRIS exposées dans l'interface.
    dossiers_autorises = (
        Path(config.REPERTOIRE_EXTRACT_IRIS_LOCAL),
        Path(config.REPERTOIRE_EXCEL_IRIS_OUTPUT),
    )
    return any(_is_relative_to(path, dossier) for dossier in dossiers_autorises)


def _lister_fichiers_locaux() -> list[Path]:
    """Liste les dumps Excel présents dans le répertoire local configuré.

    Ce répertoire local est la base de travail de l'écran IRIS. La GED sert à
    l'alimenter pour certains exports, mais l'affichage et le traitement lisent
    ensuite les fichiers depuis ce dossier local.
    """
    # Bloc lecture disque :
    # Le service partagé lève des exceptions fichier; on les convertit ici en
    # erreurs HTTP compréhensibles par le front.
    try:
        return iris_dumps_services.lister_fichiers_excel(config.REPERTOIRE_EXTRACT_IRIS_LOCAL)
    except FileNotFoundError:
        raise HTTPException(
            status_code=404,
            detail=f"Repertoire IRIS introuvable : {config.REPERTOIRE_EXTRACT_IRIS_LOCAL}",
        )
    except NotADirectoryError:
        raise HTTPException(
            status_code=500,
            detail=f"Chemin IRIS invalide : {config.REPERTOIRE_EXTRACT_IRIS_LOCAL}",
        )


def _dernier_dump_pour_code(code_export: str, fichiers: list[Path]) -> Path | None:
    """Retourne le dump le plus récent pour un code IRIS.

    Priorité :
    - date contenue dans le nom du fichier, via le service historique ;
    - date de modification système si aucun nom ne contient de date exploitable.
    """
    # Bloc sélection :
    # On commence par isoler les fichiers du code demandé, puis on applique la
    # règle métier existante de recherche du dernier dump.
    fichiers_du_code = iris_dumps_services.filtrer_par_code(fichiers, code_export)
    dernier_par_nom = iris_dumps_services.trouver_dernier_dump(fichiers_du_code)

    # Bloc retour prioritaire : le nom du fichier contient une date exploitable.
    if dernier_par_nom is not None:
        return dernier_par_nom

    # Bloc absence : aucun fichier local ne correspond au code export.
    if not fichiers_du_code:
        return None

    # Bloc fallback : si le nom ne permet pas de trancher, on prend le plus récent
    # selon la date de modification système.
    return max(fichiers_du_code, key=lambda fichier: fichier.stat().st_mtime)


def _construire_dump_row(code_export: str, fichiers: list[Path]) -> dict[str, object]:
    """Transforme un dump détecté en ligne JSON pour le tableau React.

    Fichier cible et date = dernier fichier COMPLET produit (output),
    pour refléter la date du dernier traitement effectué.
    """
    dernier_dump = _dernier_dump_pour_code(code_export, fichiers)
    repertoire_source = Path(config.REPERTOIRE_EXTRACT_IRIS_LOCAL)
    exists = dernier_dump is not None

    # Cherche le dernier fichier COMPLET produit pour ce code dans le dossier output.
    output_fichier: Path | None = None
    output_dir = Path(config.REPERTOIRE_EXCEL_IRIS_OUTPUT)
    if output_dir.exists():
        candidats = sorted(
            [f for f in output_dir.glob(f"{code_export}_*COMPLET*.xlsx") if not f.name.startswith("~$")],
            key=lambda f: f.stat().st_mtime,
            reverse=True,
        )
        output_fichier = candidats[0] if candidats else None

    # Fichier cible = output si disponible, sinon source.
    cible = output_fichier if output_fichier else dernier_dump

    return {
        "code": code_export,
        "type": EXPORTS_IRIS[code_export],
        "source": repertoire_source.name,
        "sourcePath": str(repertoire_source),
        "sourceUrl": _path_to_uri(repertoire_source),
        "fichier": cible.name if cible else "Aucun fichier trouve",
        "targetPath": str(cible) if cible else "",
        "targetUrl": _path_to_uri(cible) if cible else "",
        "updatedAt": _format_datetime(cible.stat().st_mtime if cible else None),
        "statut": "ok" if exists else "error",
        "selected": exists,
    }


def _verifier_dump_row(code_export: str, fichiers: list[Path]) -> dict[str, object]:
    """Vérifie un dump précis avant traitement.

    On contrôle uniquement les prérequis fichier indispensables au lancement :
    existence, type fichier, extension Excel, lisibilité basique et taille non
    nulle. Le contrôle métier du contenu reste dans les classes IRIS.
    """
    dernier_dump = _dernier_dump_pour_code(code_export, fichiers)

    # Bloc présence : aucun fichier exploitable pour ce code.
    if dernier_dump is None:
        return {
            "code": code_export,
            "valid": False,
            "status": "error",
            "message": "Fichier manquant",
        }

    # Bloc existence : le chemin était référencé mais n'existe plus.
    if not dernier_dump.exists():
        return {
            "code": code_export,
            "valid": False,
            "status": "error",
            "message": "Chemin introuvable",
        }

    # Bloc type de chemin : on attend un fichier Excel, pas un dossier.
    if not dernier_dump.is_file():
        return {
            "code": code_export,
            "valid": False,
            "status": "error",
            "message": "Ce chemin n'est pas un fichier",
        }

    # Bloc extension : les dumps IRIS manipulés ici sont des classeurs Excel.
    if dernier_dump.suffix.lower() not in {".xlsx", ".xls"}:
        return {
            "code": code_export,
            "valid": False,
            "status": "error",
            "message": f"Extension non attendue : {dernier_dump.suffix}",
        }

    # Bloc lisibilité : un `stat()` suffit à détecter les erreurs d'accès
    # basiques sans charger tout le classeur.
    try:
        size = dernier_dump.stat().st_size
    except OSError as exc:
        return {
            "code": code_export,
            "valid": False,
            "status": "error",
            "message": f"Lecture impossible : {exc}",
        }

    # Bloc contenu minimal : un fichier vide ne peut pas être traité.
    if size <= 0:
        return {
            "code": code_export,
            "valid": False,
            "status": "error",
            "message": "Fichier vide",
        }

    # Bloc config auto : vérifie que chaque fichier référencé dans la config existe.
    # La vérification précédente confirme qu'il y a "au moins un fichier" dans le
    # dossier, mais le traitement auto lit une liste précise issue de config_extractsIRIS.
    # Si l'un de ces fichiers est absent, le traitement échouera avec [WinError 2].
    type_export = EXPORTS_IRIS.get(code_export)
    if type_export:
        chemins_config = IRIS.DICT_EXPORTS_IRIS.get("chemins_fichiersInput", {}).get(type_export, ())
        repertoire = IRIS_natif.cei(type_export)._input.repertoire
        manquants = [nom for nom in chemins_config if not (repertoire / nom).exists()]
        if manquants:
            liste = ", ".join(manquants[:2]) + ("..." if len(manquants) > 2 else "")
            return {
                "code": code_export,
                "valid": False,
                "status": "error",
                "message": f"{len(manquants)} fichier(s) absent(s) du config : {liste}",
            }

    # Bloc succès : tous les fichiers du config sont présents et lisibles.
    return {
        "code": code_export,
        "valid": True,
        "status": "ok",
        "message": "Fichier trouve et lisible",
        "file": dernier_dump.name,
        "updatedAt": _format_datetime(dernier_dump.stat().st_mtime),
    }


def _lister_resultats_traitement(codes_exports: list[str] | tuple[str, ...] | None = None) -> list[dict[str, str]]:
    """Liste les derniers fichiers générés dans le dossier des extracts complets."""
    # Bloc dossier résultat :
    # Si le dossier n'existe pas, on renvoie une liste vide plutôt qu'une erreur,
    # car l'absence de résultats n'empêche pas d'afficher la page.
    repertoire_output = Path(config.REPERTOIRE_EXCEL_IRIS_OUTPUT)
    if not repertoire_output.exists():
        return []

    # Bloc filtre optionnel :
    # Quand l'utilisateur ne traite que certains exports, le front ne veut voir
    # que les résultats correspondants.
    codes_filtres = set(_normaliser_codes(list(codes_exports or [])))

    # Bloc tri :
    # Les plus récents sont affichés en premier dans la liste de résultats.
    fichiers_output = sorted(
        repertoire_output.glob("*.xlsx"),
        key=lambda fichier: fichier.stat().st_mtime,
        reverse=True,
    )

    rows: list[dict[str, str]] = []
    for fichier in fichiers_output:
        # Ignore les fichiers temporaires créés par Excel lorsqu'un classeur est ouvert.
        if fichier.name.startswith("~$"):
            continue

        # Bloc rattachement métier :
        # On déduit le type d'export depuis le préfixe du nom de fichier.
        code = next((code_export for code_export in EXPORTS_IRIS if fichier.name.startswith(code_export)), "")
        if codes_filtres and code not in codes_filtres:
            continue

        # Bloc ligne résultat :
        # Même logique que les dumps : nom affiché + URI pour ouverture via API.
        label = f"{code} - {EXPORTS_IRIS[code]}" if code else fichier.stem
        rows.append({
            "export": label,
            "fichier": fichier.name,
            "path": str(fichier),
            "url": _path_to_uri(fichier),
            "date": _format_datetime(fichier.stat().st_mtime),
            "statut": "ok",
        })

        # Le front n'a pas besoin de tout l'historique; on garde une liste courte.
        if len(rows) >= 8:
            break

    return rows




def _run_avec_com(func):
    """Initialise COM (requis par xlwings) dans le thread du pool, execute func, puis libere COM."""
    import pythoncom
    pythoncom.CoInitialize()
    try:
        return func()
    finally:
        pythoncom.CoUninitialize()


def _now_ts() -> float:
    """Retourne un timestamp monotone simple pour la gestion des jobs en memoire."""
    return time.time()


def _cleanup_jobs() -> None:
    """Supprime les jobs termines depuis plus de 30 minutes."""
    limite = _now_ts() - JOB_RETENTION_SECONDS
    anciens_jobs = [
        job_id
        for job_id, job in _jobs.items()
        if job.get("status") in {"done", "error", "cancelled"}
        and job.get("finished_at", float("inf")) < limite
    ]
    for job_id in anciens_jobs:
        _jobs.pop(job_id, None)


def _mark_job_finished(job: dict, status: str) -> None:
    """Marque un job comme termine et memorise l'heure de fin pour le nettoyage."""
    job["status"] = status
    job["finished_at"] = _now_ts()


def _valider_fichier_manuel(chemin: Path, code_export: str) -> Path:
    """Valide un fichier choisi manuellement avant de le transmettre au metier."""
    if not chemin.exists():
        raise HTTPException(status_code=400, detail=f"Fichier introuvable : {chemin}")

    if not chemin.is_file():
        raise HTTPException(status_code=400, detail=f"Ce chemin n'est pas un fichier : {chemin}")

    if chemin.suffix.lower() not in {".xlsx", ".xls"}:
        raise HTTPException(
            status_code=400,
            detail=f"Extension non attendue pour {chemin.name} : attendu .xlsx ou .xls.",
        )

    if not chemin.name.startswith(code_export):
        raise HTTPException(
            status_code=400,
            detail=f"Mauvais fichier pour {code_export} : {chemin.name}. Le nom doit commencer par {code_export}.",
        )

    return chemin


def _ligne_fichier_manuel(chemin: Path) -> dict[str, str]:
    """Transforme un chemin choisi par Tkinter en ligne JSON pour React."""
    return {
        "name": chemin.name,
        "path": str(chemin),
        "url": _path_to_uri(chemin),
    }




# =============================================================================
# Endpoints de lecture
# =============================================================================

@router.get("/dumps")
def get_dumps() -> dict[str, object]:
    """Retourne l'état courant des dumps et des derniers résultats.

    Cette route alimente le tableau principal au chargement de la page.
    """
    # Bloc état local :
    # On lit une seule fois les fichiers, puis on construit une ligne par export
    # connu pour garder un tableau stable côté front.
    fichiers = _lister_fichiers_locaux()

    # Bloc réponse :
    # On renvoie aussi les dossiers source/sortie pour les actions d'ouverture,
    # même si le front peut choisir de ne pas les afficher.
    return {
        "dumps": [_construire_dump_row(code_export, fichiers) for code_export in EXPORTS_IRIS],
        "outputFiles": _lister_resultats_traitement(),
        "sourceDirectory": str(config.REPERTOIRE_EXTRACT_IRIS_LOCAL),
        "sourceDirectoryUrl": _path_to_uri(Path(config.REPERTOIRE_EXTRACT_IRIS_LOCAL)),
        "outputDirectory": str(config.REPERTOIRE_EXCEL_IRIS_OUTPUT),
        "outputDirectoryUrl": _path_to_uri(Path(config.REPERTOIRE_EXCEL_IRIS_OUTPUT)),
    }


# =============================================================================
# Endpoints d'action
# =============================================================================

@router.post("/update-dumps")
def update_dumps(payload: UpdateDumpsRequest) -> dict[str, object]:
    """Copie les derniers dumps GED utiles et archive les anciens locaux.

    Le service métier applique la règle actuelle :
    - copie GED uniquement pour `GED_EXPORT_CODES` ;
    - archivage local pour tous les codes demandés.
    """
    # Bloc validation entrée :
    # Les codes inconnus sont retirés; une liste vide devient une erreur client.
    codes = _normaliser_codes(payload.codes)
    if not codes:
        raise HTTPException(status_code=400, detail="Aucun code IRIS valide.")

    # Bloc séparation métier :
    # `codes_ged` pilote la copie depuis la GED; `codes_archives` pilote le
    # nettoyage des anciens dumps locaux.
    codes_ged = tuple(code for code in codes if code in GED_EXPORT_CODES)
    codes_archives = tuple(codes)

    # Bloc action GED/local :
    # La fonction métier copie, archive et renvoie un bilan par code export.
    try:
        bilan = iris_dumps_services.mettre_a_jour_dumps_depuis_ged(
            codes_exports=codes_ged,
            codes_exports_a_archiver=codes_archives,
        )
    except Exception as exc:
        # En environnement dégradé, la GED réseau peut être inaccessible
        # (VPN, droits du processus, sandbox, chemin non monté). On ne bloque pas
        # l'écran : les dumps locaux restent utilisables.
        return {
            "message": f"GED inaccessible : mode local conserve. Detail : {exc}",
            "updatedCodes": codes,
            "gedCodes": list(codes_ged),
            "result": {},
            **get_dumps(),
        }

    # Bloc succès :
    # On renvoie le bilan puis l'état complet recalculé pour que React remplace
    # immédiatement son tableau.
    return {
        "message": "Mise a jour IRIS terminee.",
        "updatedCodes": codes,
        "gedCodes": list(codes_ged),
        "result": {
            code: {
                "copie": str(data.get("copie")) if data.get("copie") else None,
                "archives": [str(path) for path in data.get("archives", [])],
            }
            for code, data in bilan.items()
        },
        **get_dumps(),
    }


@router.post("/validate-structure")
def validate_structure(payload: ValidateStructureRequest | None = None) -> dict[str, object]:
    """Vérifie chaque dump attendu avant traitement."""
    # Bloc périmètre :
    # Sans payload, on contrôle tous les exports; sinon seulement la sélection.
    fichiers = _lister_fichiers_locaux()
    codes = _normaliser_codes(payload.codes) if payload and payload.codes else list(EXPORTS_IRIS)

    # Bloc contrôle par ligne :
    # Chaque export reçoit son propre statut, ce qui permet au tableau de montrer
    # exactement quel fichier bloque.
    checks = [_verifier_dump_row(code_export, fichiers) for code_export in codes]
    invalides = [check for check in checks if not check["valid"]]

    # Bloc réponse avec erreurs :
    # HTTP 200 volontaire : la vérification a réussi techniquement, même si des
    # fichiers métier sont invalides.
    if invalides:
        return {
            "valid": False,
            "message": "Verification terminee avec erreurs.",
            "checks": checks,
        }

    # Bloc réponse succès :
    # Tous les fichiers sélectionnés sont présents et lisibles.
    return {
        "valid": True,
        "message": "Tous les dumps IRIS attendus sont presents et lisibles.",
        "checks": checks,
    }


@router.post("/open-path")
def open_path(payload: OpenPathRequest) -> dict[str, str]:
    """Ouvre un fichier ou dossier IRIS sur le poste Windows.

    Cette route remplace les liens `file://` directs côté React. Cela évite les
    blocages navigateur et donne une UX proche d'un double-clic Explorateur.
    """
    # Bloc conversion :
    # Le front envoie l'URI stockée dans les lignes; on revient au chemin Windows.
    path = _uri_to_path(payload.path)

    # Bloc validation existence :
    # Message clair si le fichier a été déplacé entre l'affichage et le clic.
    if not path.exists():
        raise HTTPException(status_code=404, detail=f"Chemin introuvable : {path}")

    # Bloc sécurité :
    # Même si le front est local, on refuse d'ouvrir un chemin hors périmètre IRIS.
    if not _chemin_ouvrable(path):
        raise HTTPException(status_code=403, detail=f"Ouverture non autorisee : {path}")

    # Bloc ouverture Windows :
    # `os.startfile` ne retourne pas de handle exploitable; si l'appel passe,
    # l'ouverture est considérée déclenchée.
    try:
        # Windows choisit l'application associée : Excel pour .xlsx,
        # Explorateur pour un dossier, etc.
        os.startfile(str(path))  # type: ignore[attr-defined]
    except OSError as exc:
        raise HTTPException(status_code=500, detail=f"Ouverture impossible : {exc}")

    return {"status": "ok", "path": str(path)}


def _run_tkinter_selection(code: str, type_export: str) -> list[Path]:
    """Ouvre la fenetre Tkinter de selection de fichiers IRIS (fonction bloquante sync).

    Tkinter doit tourner dans un thread OS, jamais dans l'event loop asyncio.
    A appeler uniquement via `await asyncio.to_thread(_run_tkinter_selection, ...)`.
    """
    import tkinter as tk
    from tkinter import filedialog, messagebox

    titre = f"Selectionner uniquement des fichiers Extract IRIS {type_export} ({code})"
    chemins_valides: list[Path] = []

    while True:
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        root.update()
        try:
            chemins = filedialog.askopenfilenames(
                parent=root,
                title=titre,
                filetypes=[("Fichiers Excel", "*.xlsx *.xls")],
                initialdir=str(Path(config.REPERTOIRE_EXTRACT_IRIS_LOCAL)),
            )
        finally:
            root.attributes("-topmost", False)
            root.destroy()

        if not chemins:
            break

        # Controle strict : Excel uniquement, nom commencant par le code export.
        refus_extension = [Path(c).name for c in chemins if Path(c).suffix.lower() not in {".xlsx", ".xls"}]
        refus_code = [
            Path(c).name
            for c in chemins
            if Path(c).suffix.lower() in {".xlsx", ".xls"} and not Path(c).name.startswith(code)
        ]
        refus = refus_extension + refus_code
        if refus:
            alerte = tk.Tk()
            alerte.withdraw()
            alerte.attributes("-topmost", True)
            alerte.update()
            details = []
            if refus_extension:
                details.append(
                    "Ces fichiers ne sont pas des fichiers Excel :\n"
                    + "\n".join(f"  • {nom}" for nom in refus_extension)
                )
            if refus_code:
                details.append(
                    f"Ces fichiers ne correspondent pas au code {code} ({type_export}) :\n"
                    + "\n".join(f"  • {nom}" for nom in refus_code)
                )
            messagebox.showwarning(
                title=f"Fichier(s) refusé(s) — {code}",
                message=(
                    "\n\n".join(details)
                    + f"\n\nSélectionnez uniquement des fichiers Excel dont le nom commence par {code}."
                    + "\n\nLa fenêtre de sélection va se rouvrir."
                ),
                parent=alerte,
            )
            alerte.destroy()
            continue

        chemins_valides = [Path(c) for c in chemins]
        break

    return chemins_valides


@router.post("/select-manual-files")
async def select_manual_files(payload: SelectManualFilesRequest) -> dict[str, object]:
    """Ouvre la fenetre Python/Tkinter locale pour choisir des fichiers IRIS.

    Cette route remplace le selecteur navigateur `<input type=file>` pour le mode
    manuel. Tkinter est lance dans un thread OS via `asyncio.to_thread` pour ne
    pas bloquer l'event loop FastAPI pendant que la fenetre de selection est ouverte.
    """
    codes = _normaliser_codes([payload.export])
    if len(codes) != 1:
        raise HTTPException(status_code=400, detail="Selectionnez un seul export IRIS valide.")

    code = codes[0]
    type_export = EXPORTS_IRIS[code]

    chemins_valides = await asyncio.to_thread(_run_tkinter_selection, code, type_export)

    return {
        "export": code,
        "type": type_export,
        "files": [_ligne_fichier_manuel(chemin) for chemin in chemins_valides],
    }


@router.get("/job/{job_id}")
def get_job(job_id: str) -> dict[str, object]:
    """Retourne l'état courant d'un job de traitement IRIS."""
    _cleanup_jobs()
    job = _jobs.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail=f"Job inconnu : {job_id}")
    return job


@router.delete("/job/{job_id}")
def cancel_job(job_id: str) -> dict[str, str]:
    """Demande l'annulation propre d'un job en cours."""
    _cleanup_jobs()
    job = _jobs.get(job_id)
    if job and job.get("status") == "running":
        job["cancelled"] = True
        job["step"] = "Annulation en cours..."
    return {"status": "cancel_requested"}




@router.post("/traiter")
async def traiter(payload: RunTreatmentRequest) -> dict[str, object]:
    """Lance la consolidation IRIS en arrière-plan et retourne un job_id.

    Le front peut ensuite interroger GET /job/{job_id} pour suivre la
    progression même après un rechargement de page.
    """
    _cleanup_jobs()
    codes = _normaliser_codes(payload.exports)
    if not codes:
        raise HTTPException(status_code=400, detail="Aucun export IRIS valide.")

    types_exports = tuple(EXPORTS_IRIS[code] for code in codes)
    job_id = str(uuid4())
    # Etat initial du job lu par le front:
    # - `status` pilote la boucle de polling;
    # - `progress` et `step` alimentent la barre;
    # - `statut_codes` colore chaque ligne du tableau;
    # - `files` sera rempli au fur et a mesure des fichiers produits.
    _jobs[job_id] = {
        "status": "running",
        "progress": 10,
        "step": f"Démarrage — {', '.join(types_exports)}...",
        "files": [],
        "message": "",
        "warning": False,
        "statut_codes": {code: "gray" for code in codes},
        "steps_history": [],
    }

    # Le traitement pandas/openpyxl est long et bloquant. On le lance dans une
    # tache asynchrone afin que l'endpoint rende immediatement le job_id.
    # `_spawn_job_task` garde une reference forte pour eviter qu'elle soit
    # annulee par le garbage collector pendant le traitement.
    _spawn_job_task(_bg_traitement_auto(job_id, codes, types_exports))
    return {"job_id": job_id}


async def _bg_traitement_auto(job_id: str, codes: list[str], types_exports: tuple[str, ...]) -> None:
    """Exécute la consolidation IRIS type par type avec progression détaillée."""
    job = _jobs[job_id]
    job["statut_codes"] = {code: "gray" for code in codes}
    job["steps_history"] = []
    n = len(codes)

    try:
        for i, (code, type_export) in enumerate(zip(codes, types_exports)):
            # Le front affiche le code courant en "en cours".
            job["current_code"] = code
            job["statut_codes"][code] = "blue"

            # Plage de progression réservée à ce type : 5% → 95% répartis équitablement
            base = 5 + int(i / n * 90)
            span = int(90 / n)

            def cb(pct: int, step: str, _base: int = base, _span: int = span, _te: str = type_export) -> None:
                # Callback appele par le code metier IRIS. Il transforme la
                # progression interne d'un export en progression globale du job.
                if job.get("cancelled"):
                    raise InterruptedError("Traitement annulé par l'utilisateur.")
                label = f"{_te} — {step}"
                job["progress"] = _base + int(pct / 100 * _span)
                job["step"] = label
                job["steps_history"].append(label)

            # Relit le dossier au moment du traitement pour inclure les fichiers
            # ajoutés depuis le démarrage du backend (évite les listes figées à l'import).
            noms_fichiers = iris_dumps_services.construire_liste_fichiers_type(
                config.REPERTOIRE_EXTRACT_IRIS_LOCAL, code
            )
            repertoire = IRIS_natif.cei(type_export)._input.repertoire
            chemins = tuple(repertoire / nom for nom in noms_fichiers)

            cb(0, f"Vérification de {len(chemins)} fichier(s)...")

            instance = await asyncio.to_thread(
                _run_avec_com,
                lambda te=type_export, c=chemins, callback=cb: IRIS_natif.avec_traitement(
                    typeExport=te,
                    chemins_fichiersInput=c,
                    progress_callback=callback,
                )
            )

            job["statut_codes"][code] = "ok"

            # Capture uniquement le fichier produit par ce traitement
            try:
                # `IRIS_natif.avec_traitement` renvoie l'instance qui connait son
                # chemin de sortie. On convertit ce chemin en ligne JSON pour React.
                out = instance.chemin
                if out and out.exists():
                    job["files"].append({
                        "export": f"{code} - {type_export}",
                        "fichier": out.name,
                        "path": str(out),
                        "url": _path_to_uri(out),
                        "date": _format_datetime(out.stat().st_mtime),
                        "statut": "ok",
                    })
            except Exception:
                pass

        _mark_job_finished(job, "done")
        job["progress"] = 100
        job["step"] = "Tous les exports traités ✓"
        job["current_code"] = None
        job["message"] = f"Traitement terminé pour {', '.join(types_exports)}."

    except InterruptedError:
        _mark_job_finished(job, "cancelled")
        job["progress"] = 100
        job["step"] = "Annulé par l'utilisateur"
        job["message"] = "Traitement annulé — tu peux relancer."
        job["current_code"] = None
        # Remet les statuts en attente pour permettre un relaunch propre
        job["statut_codes"] = {code: "gray" for code in job.get("statut_codes", {})}

    except Exception as exc:
        _mark_job_finished(job, "error")
        job["progress"] = 100
        job["step"] = "Erreur"
        job["message"] = str(exc)
        job["warning"] = True
        if "current_code" in job and job["current_code"]:
            job["statut_codes"][job["current_code"]] = "error"




@router.post("/traiter-manuel-chemins")
async def traiter_manuel_chemins(payload: RunManualPathsRequest) -> dict[str, object]:
    """Lance un traitement manuel depuis des chemins choisis par le backend."""
    _cleanup_jobs()
    codes = _normaliser_codes([payload.export])
    if len(codes) != 1:
        raise HTTPException(status_code=400, detail="Selectionnez un seul export IRIS valide.")

    code = codes[0]
    type_export = EXPORTS_IRIS[code]

    if not payload.paths:
        raise HTTPException(status_code=400, detail="Aucun fichier fourni.")

    tous_les_chemins = tuple(
        _valider_fichier_manuel(Path(path), code) for path in payload.paths
    )

    job_id = str(uuid4())
    _jobs[job_id] = {
        "status": "running",
        "progress": 25,
        "step": f"{type_export} — lecture et concaténation des fichiers Excel...",
        "files": [],
        "message": "",
        "warning": False,
        "statut_codes": {code: "blue"},
        "steps_history": [],
    }

    _spawn_job_task(_bg_traitement_manuel(job_id, code, type_export, tous_les_chemins))
    return {"job_id": job_id}


async def _bg_traitement_manuel(
    job_id: str, code: str, type_export: str, chemins: tuple[Path, ...]
) -> None:
    """Exécute le traitement manuel IRIS dans un thread séparé avec progression détaillée."""
    job = _jobs[job_id]
    job["statut_codes"] = {code: "blue"}
    job["current_code"] = code
    job["steps_history"] = []

    def cb(pct: int, step: str) -> None:
        # Meme mecanisme que le mode auto, mais la progression couvre presque
        # toute la barre parce qu'un job manuel ne traite qu'un export.
        if job.get("cancelled"):
            raise InterruptedError("Traitement annulé par l'utilisateur.")
        label = f"{type_export} — {step}"
        job["progress"] = 10 + int(pct / 100 * 88)
        job["step"] = label
        job["steps_history"].append(label)

    try:
        instance = await asyncio.to_thread(
            _run_avec_com,
            lambda: IRIS_natif.avec_traitement(
                typeExport=type_export,
                chemins_fichiersInput=chemins,
                progress_callback=cb,
            )
        )

        _mark_job_finished(job, "done")
        job["progress"] = 100
        job["step"] = f"{type_export} — terminé ✓"
        job["current_code"] = None
        job["statut_codes"] = {code: "ok"}
        job["message"] = f"Traitement manuel terminé pour {type_export}."

        try:
            out = instance.chemin
            if out and out.exists():
                job["files"].append({
                    "export": f"{code} - {type_export}",
                    "fichier": out.name,
                    "path": str(out),
                    "url": _path_to_uri(out),
                    "date": _format_datetime(out.stat().st_mtime),
                    "statut": "ok",
                })
        except Exception:
            pass

    except InterruptedError:
        _mark_job_finished(job, "cancelled")
        job["progress"] = 100
        job["step"] = "Annulé par l'utilisateur"
        job["message"] = "Traitement annulé — tu peux relancer."
        job["statut_codes"] = {code: "gray"}
        job["current_code"] = None

    except Exception as exc:
        _mark_job_finished(job, "error")
        job["progress"] = 100
        job["step"] = "Erreur"
        job["message"] = str(exc)
        job["warning"] = True
        job["statut_codes"] = {code: "error"}
