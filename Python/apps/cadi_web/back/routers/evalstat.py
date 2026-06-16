from __future__ import annotations

import os
import re
from datetime import datetime
from pathlib import Path
from urllib.parse import unquote, urlparse

import pandas as pd
from fastapi import APIRouter, File, Form, HTTPException, UploadFile
from pydantic import BaseModel

from vte.core import config
from vte.domain.evalStat import EvalStat_formation
from vte.domain.formation import Formation
from vte.domain.session import Session


# =============================================================================
# Routeur EvalStat
# =============================================================================
#
# Ce module est la facade HTTP entre l'interface React et le traitement des
# bilans d'evaluation de session (EvalStat) du package `vte`.
#
# Methode retenue : l'utilisateur saisit un trigramme de formation (pas de
# code IRIS). CADI affiche la liste des sessions IRIS de cette formation,
# triees par date de debut decroissante, avec le CSV "*-Stagiaires.csv"
# retrouve automatiquement dans FORMATIONS_C a partir du trigramme et du code
# IRIS. L'utilisateur coche les sessions a traiter puis lance le traitement.
#
# Guide rapide du flux :
# 1. Au demarrage du back, CADI precharge en cache les exports IRIS R0304
#    (formations) et R04110 (sessions).
# 2. GET /formations fournit les trigrammes proposes a la recherche depuis R0304.
# 3. GET /sessions?formation=XXX liste les CSV EvalStat trouves dans
#    FORMATIONS_C pour ce trigramme, puis enrichit les lignes avec R04110.
# 4. POST /process traite les sessions cochees (bilans session + formation).
# 5. POST /open-path ouvre le fichier ou dossier demande sur le poste serveur.


router = APIRouter(prefix="/evalstat", tags=["EvalStat"])


# =============================================================================
# Modeles d'entree API
# =============================================================================

class OpenPathRequest(BaseModel):
    """Chemin ou URI `file://` que le backend doit ouvrir via Windows."""

    path: str


class ProcessSessionsRequest(BaseModel):
    """CSV (chemins ou URI) des sessions cochees pour une formation, a traiter en lot."""

    formation: str
    paths: list[str] = []


# =============================================================================
# Constantes
# =============================================================================

PYTHON_ROOT = Path(__file__).resolve().parents[4]

# Les CSV "Stagiaires" peuvent se trouver dans l'un ou l'autre de ces
# sous-dossiers selon l'anciennete de la session.
CSV_EVAL_DIR_NAMES = (
    "rapports-sessions-CSV-evaluations",
    "rapports-sessions-evaluations",
)
CSV_STAGIAIRES_PATTERN = "*-Stagiaires.csv"

IRIS_SESSIONS_PATTERN = "R04110_Sessions-COMPLET*.xlsx"
IRIS_FORMATIONS_PATTERN = "R0304_Formations-COMPLET*.xlsx"

FORMATIONS_C_UNC_ROOT = Path("//instnt/partage/FORMATIONS_C")
FORMATIONS_C_MAPPED_ROOT = Path("P:/FORMATIONS_C")


# =============================================================================
# Fonctions utilitaires locales (chemins, dates, normalisation)
# =============================================================================

def _normaliser_trigramme(trigramme: str) -> str:
    """Nettoie un trigramme de formation et rejette les valeurs vides."""
    trigramme_clean = trigramme.strip().upper()
    if not trigramme_clean:
        raise HTTPException(status_code=400, detail="Le trigramme de formation est obligatoire.")
    return trigramme_clean


def _path_to_uri(path: Path) -> str:
    """Construit une URI `file://` pour le front (utilisee comme identifiant/tooltip)."""
    try:
        return path.resolve().as_uri()
    except ValueError:
        return str(path)


def _open_windows_path(path: Path) -> None:
    """Ouvre un fichier ou un dossier avec l'application Windows associee."""
    os.startfile(path)


def _uri_to_path(path_or_uri: str) -> Path:
    """Convertit un chemin Windows brut ou une URI `file://` (UNC ou local) en `Path`."""
    if path_or_uri.startswith("file://"):
        parsed = urlparse(path_or_uri)
        decoded_path = unquote(parsed.path)
        if parsed.netloc:
            # Cas UNC : file://serveur/partage/dossier -> //serveur/partage/dossier
            return Path(f"//{parsed.netloc}{decoded_path}")
        # Cas local : file:///C:/... -> C:/...
        return Path(decoded_path.lstrip("/"))
    return Path(path_or_uri)


def _format_datetime(timestamp: float | None) -> str:
    """Convertit un timestamp fichier en libelle lisible par React (ou "-" si absent)."""
    if timestamp is None:
        return "-"
    return datetime.fromtimestamp(timestamp).strftime("%d/%m/%Y %H:%M")


def _format_excel_date(value: object) -> str:
    """Convertit une cellule date pandas en chaine ISO "AAAA-MM-JJ" (ou "" si vide)."""
    if value is None or pd.isna(value):
        return ""
    try:
        return pd.Timestamp(value).strftime("%Y-%m-%d")
    except (TypeError, ValueError):
        return ""


def _safe_int(value: object) -> int | None:
    """Convertit une cellule pandas en `int`, ou `None` si vide/non numerique."""
    try:
        if pd.isna(value):
            return None
        return int(float(value))
    except (TypeError, ValueError):
        return None


def _text(value: object) -> str:
    """Convertit une cellule pandas en texte nettoye, ou chaine vide si vide."""
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


# =============================================================================
# Index des sessions IRIS (fichier R04110_Sessions-COMPLET*.xlsx)
# =============================================================================
#
# Cet index croise chaque code IRIS avec sa formation et les informations
# affichees au front (dates, statut, nombre de nommes...). Le fichier source
# est un export complet potentiellement volumineux situe sur un partage
# reseau : sa lecture est mise en cache pour ne pas re-parser tout le classeur
# a chaque appel de /formations ou /sessions.

def _iris_extract_root_candidates() -> tuple[Path, ...]:
    """Dossiers possibles contenant les exports IRIS (sessions, formations), par ordre de priorite."""
    return (
        Path("H:/_Echanges/VTE/Prog/IRIS/Extracts complets"),
        config.REPERTOIRE_EXCEL_IRIS_OUTPUT,
        PYTHON_ROOT / "IRIS" / "Extracts complets",
    )


def _latest_file(root: Path, pattern: str) -> Path | None:
    """Retourne le fichier le plus recent de `root` correspondant a `pattern`."""
    if not root.exists() or not root.is_dir():
        return None
    try:
        files = [path for path in root.glob(pattern) if path.is_file()]
    except OSError:
        return None
    if not files:
        return None
    return max(files, key=lambda path: path.stat().st_mtime)


def _latest_iris_sessions_file() -> Path | None:
    """Localise le dernier export IRIS des sessions parmi les dossiers candidats."""
    for root in _iris_extract_root_candidates():
        latest = _latest_file(root, IRIS_SESSIONS_PATTERN)
        if latest is not None:
            return latest
    return None


# Cache memoire du dernier index IRIS lu, invalide via le chemin + la date de
# modification du fichier source. Tant que l'export n'a pas change, /formations
# et /sessions reutilisent ce resultat au lieu de re-lire le xlsx.
_iris_sessions_cache: dict[str, object] = {"path": None, "mtime": None, "sessions": None}


def _load_iris_sessions_index() -> tuple[dict[int, dict[str, object]], Path | None]:
    """Charge l'index {code IRIS -> infos session} depuis le dernier export, avec cache memoire."""
    iris_file = _latest_iris_sessions_file()
    if iris_file is None:
        return {}, None

    # Bloc cache : un stat() est largement plus rapide qu'un read_excel complet,
    # donc on ne re-parse le classeur que si le fichier source a change.
    try:
        mtime = iris_file.stat().st_mtime
    except OSError:
        mtime = None

    cache = _iris_sessions_cache
    if (
        cache["sessions"] is not None
        and cache["path"] == str(iris_file)
        and cache["mtime"] == mtime
    ):
        return cache["sessions"], iris_file

    # Bloc lecture : dans cet export, l'entete reel commence a la ligne 2 (header=1).
    try:
        df = pd.read_excel(iris_file, header=1)
    except Exception:
        return {}, iris_file

    required_cols = {"Code IRIS", "Trigramme formation", "N° Session"}
    if not required_cols.issubset(set(df.columns)):
        return {}, iris_file

    # Bloc indexation : une entree par code IRIS, avec uniquement les colonnes
    # demandees par le tableau /sessions.
    sessions: dict[int, dict[str, object]] = {}
    for _, row in df.iterrows():
        code_iris = _safe_int(row.get("Code IRIS"))
        if code_iris is None:
            continue
        sessions[code_iris] = {
            "codeIris": code_iris,
            "formation": _text(row.get("Trigramme formation")).upper(),
            "trigrammeRp": _text(row.get("Trigramme RP")),
            "sessionRef": _text(row.get("N° Session")),
            "sessionTitle": _text(row.get("Session")),
            "sessionStatus": _text(row.get("Statut \nSession")),
            "startDate": _format_excel_date(row.get("Date début ses.")),
            "endDate": _format_excel_date(row.get("Date fin ses.")),
            "nbNommes": _safe_int(row.get("Nb. Nommés")),
        }

    cache["path"] = str(iris_file)
    cache["mtime"] = mtime
    cache["sessions"] = sessions
    return sessions, iris_file


def _extraire_code_depuis_chemin_sans_popup(path: Path) -> int | None:
    """Extrait un code IRIS (5 chiffres) depuis un chemin, sans dialogue utilisateur."""
    match = re.search(r"\b\d{5}\b", str(path))
    return int(match.group(0)) if match else None


def _extraire_code_depuis_csv(path: Path) -> int | None:
    """Tente de retrouver un code IRIS dans le contenu d'un CSV mal nomme."""
    try:
        content = path.read_text(encoding="utf-8-sig", errors="ignore")
    except OSError:
        return None

    match = re.search(r"\b\d{5}\b", content)
    return int(match.group(0)) if match else None


# =============================================================================
# Index des formations IRIS (fichier R0304_Formations-COMPLET*.xlsx)
# =============================================================================
#
# Cet index donne l'ensemble des trigrammes de formation connus dans IRIS. Il
# alimente les suggestions du champ de recherche, une fois croise avec les
# dossiers reellement presents sur FORMATIONS_C (cf. _existing_formation_codes).

def _latest_iris_formations_file() -> Path | None:
    """Localise le dernier export IRIS des formations parmi les dossiers candidats."""
    for root in _iris_extract_root_candidates():
        latest = _latest_file(root, IRIS_FORMATIONS_PATTERN)
        if latest is not None:
            return latest
    return None


# Cache memoire du dernier ensemble de trigrammes IRIS, invalide via le chemin
# + la date de modification du fichier source.
_iris_formations_cache: dict[str, object] = {"path": None, "mtime": None, "codes": None}


def _load_iris_formations_codes() -> set[str]:
    """Charge l'ensemble des trigrammes de formation connus dans IRIS, avec cache memoire."""
    iris_file = _latest_iris_formations_file()
    if iris_file is None:
        return set()

    try:
        mtime = iris_file.stat().st_mtime
    except OSError:
        mtime = None

    cache = _iris_formations_cache
    if (
        cache["codes"] is not None
        and cache["path"] == str(iris_file)
        and cache["mtime"] == mtime
    ):
        return cache["codes"]

    # Bloc lecture : dans cet export, l'entete est sur la premiere ligne (header=0).
    try:
        df = pd.read_excel(iris_file, header=0)
    except Exception:
        return set()

    if "Trigramme formation" not in df.columns:
        return set()

    codes = {
        _text(value).upper()
        for value in df["Trigramme formation"]
        if _text(value)
    }

    cache["path"] = str(iris_file)
    cache["mtime"] = mtime
    cache["codes"] = codes
    return codes


def warm_evalstat_cache() -> None:
    """Precharge les exports IRIS EvalStat sans attendre la premiere action utilisateur."""
    try:
        _load_iris_formations_codes()
        _load_iris_sessions_index()
    except Exception:
        # Le prechargement est un confort : une erreur ne doit pas empecher
        # FastAPI de demarrer. Les routes remonteront l'erreur utile ensuite.
        pass


# =============================================================================
# Decouverte des dossiers FORMATIONS_C et des CSV EvalStat
# =============================================================================

def _formation_root_candidates() -> tuple[Path, ...]:
    """Racines possibles du partage FORMATIONS_C (UNC, lecteur mappe, GED), sans doublons."""
    raw_roots = (
        FORMATIONS_C_UNC_ROOT,
        FORMATIONS_C_MAPPED_ROOT,
        config.GED / "FORMATIONS_C",
    )
    roots: list[Path] = []
    seen: set[str] = set()
    for root in raw_roots:
        key = str(root).lower()
        if key in seen:
            continue
        seen.add(key)
        roots.append(root)
    return tuple(roots)


def _existing_formation_codes() -> set[str]:
    """Liste les trigrammes de formation pour lesquels un dossier existe sur le partage."""
    codes: set[str] = set()
    for root in _formation_root_candidates():
        try:
            if not root.exists() or not root.is_dir():
                continue
            codes.update(
                item.name.strip().upper()
                for item in root.iterdir()
                if item.is_dir() and item.name.strip()
            )
        except OSError:
            continue
    return codes


def _safe_is_dir(path: Path) -> bool:
    """Teste un dossier reseau sans laisser remonter les erreurs d'acces Windows."""
    try:
        return path.exists() and path.is_dir()
    except OSError:
        return False


def _safe_is_file(path: Path) -> bool:
    """Teste un fichier reseau sans laisser remonter les erreurs d'acces Windows."""
    try:
        return path.exists() and path.is_file()
    except OSError:
        return False


def _formation_directory_status(trigramme: str) -> tuple[bool, str | None]:
    """Indique si FORMATIONS_C/<trigramme> est lisible depuis le backend."""
    inaccessible_paths: list[str] = []
    for root in _formation_root_candidates():
        formation_dir = root / trigramme
        try:
            if formation_dir.exists() and formation_dir.is_dir():
                return True, None
        except OSError:
            inaccessible_paths.append(str(formation_dir))

    if inaccessible_paths:
        return False, (
            f"Le dossier FORMATIONS_C/{trigramme} n'est pas accessible depuis le backend. "
            "Aucun CSV ne peut etre detecte."
        )

    return False, f"Aucun dossier FORMATIONS_C/{trigramme} n'a ete trouve depuis le backend."


def _nearest_search_directory(trigramme: str) -> Path | None:
    """Retourne le dossier le plus proche ou guider l'utilisateur quand aucun CSV n'est trouve."""
    for root in _formation_root_candidates():
        candidates = [
            root / trigramme / "P07-bilan-sessions-et-bilan-formation" / CSV_EVAL_DIR_NAMES[0],
            root / trigramme / "P07-bilan-sessions-et-bilan-formation" / CSV_EVAL_DIR_NAMES[1],
            root / trigramme / "P07-bilan-sessions-et-bilan-formation",
            root / trigramme,
            root,
        ]
        for candidate in candidates:
            if _safe_is_dir(candidate):
                return candidate
    return FORMATIONS_C_UNC_ROOT / trigramme


def _manual_upload_directory(trigramme: str) -> Path:
    """Dossier local de secours pour les CSV choisis depuis le navigateur."""
    return PYTHON_ROOT / "apps" / "cadi_web" / "back" / ".evalstat_uploads" / trigramme


def _safe_upload_name(filename: str) -> str:
    """Nettoie un nom de fichier upload sans casser les codes IRIS presents."""
    name = Path(filename or "evalstat.csv").name
    safe_name = re.sub(r"[^A-Za-z0-9_.() -]", "_", name).strip(" .")
    return safe_name or "evalstat.csv"


def _csv_eval_dir_candidates(formation: str | None = None) -> tuple[Path, ...]:
    """Dossiers candidats contenant les CSV "Stagiaires" pour une formation (ou racines brutes)."""
    if not formation:
        return _formation_root_candidates()

    trigramme = formation.strip().upper()
    return tuple(
        root
        / trigramme
        / "P07-bilan-sessions-et-bilan-formation"
        / eval_dir_name
        for root in _formation_root_candidates()
        for eval_dir_name in CSV_EVAL_DIR_NAMES
    )


def _find_csvs_for_formation(trigramme: str) -> dict[int, Path]:
    """Indexe les CSV "*-Stagiaires.csv" de `trigramme` par code IRIS detecte dans leur chemin."""
    csv_by_code: dict[int, Path] = {}
    for eval_dir in _csv_eval_dir_candidates(trigramme):
        if not _safe_is_dir(eval_dir):
            continue
        try:
            csv_paths = eval_dir.rglob(CSV_STAGIAIRES_PATTERN)
        except OSError:
            continue

        for csv_path in csv_paths:
            code_iris = _extraire_code_depuis_chemin_sans_popup(csv_path)
            if code_iris is not None and code_iris not in csv_by_code:
                csv_by_code[code_iris] = csv_path

        if csv_by_code:
            return csv_by_code

    return csv_by_code


def _session_row_from_csv(trigramme: str, code_iris: int, csv_path: Path, iris_sessions: dict[int, dict[str, object]]) -> dict[str, object] | None:
    """Construit une ligne de tableau EvalStat depuis un CSV et les infos IRIS."""
    session = iris_sessions.get(code_iris)
    if session is None or session.get("formation") != trigramme:
        return None
    return {
        **session,
        "csvPath": str(csv_path),
        "csvUrl": _path_to_uri(csv_path),
        "csvExists": _safe_is_file(csv_path),
    }


# =============================================================================
# Verification et traitement des CSV EvalStat
# =============================================================================

def _verifier_csv(trigramme: str, path: Path) -> dict[str, object]:
    """Controle un CSV avant traitement (extension, existence, code IRIS dans le nom, taille)."""
    trigramme_clean = _normaliser_trigramme(trigramme)
    code_iris = _extraire_code_depuis_chemin_sans_popup(path) or _extraire_code_depuis_csv(path)

    checks: list[dict[str, object]] = [
        {"label": "Trigramme", "status": "ok", "message": trigramme_clean},
        {
            "label": "Extension",
            "status": "ok" if path.suffix.lower() == ".csv" else "error",
            "message": path.suffix or "Aucune extension",
        },
        {
            "label": "Existence fichier",
            "status": "ok" if path.exists() and path.is_file() else "error",
            "message": str(path),
        },
        {
            "label": "Code IRIS",
            "status": "ok" if code_iris is not None else "warning",
            "message": str(code_iris) if code_iris is not None else "Aucun code a 5 chiffres detecte",
        },
    ]

    if path.exists() and path.is_file():
        try:
            size = path.stat().st_size
        except OSError as exc:
            size = 0
            checks.append({"label": "Lecture taille", "status": "error", "message": str(exc)})

        checks.append({
            "label": "Fichier non vide",
            "status": "ok" if size > 0 else "error",
            "message": f"{size} octets",
        })

    valid = all(check["status"] != "error" for check in checks)
    return {
        "valid": valid,
        "message": "CSV pret pour traitement." if valid else "CSV non conforme.",
        "trigramme": trigramme_clean,
        "codeIris": code_iris,
        "file": {
            "name": path.name,
            "path": str(path),
            "url": _path_to_uri(path),
            "updatedAt": _format_datetime(path.stat().st_mtime if path.exists() else None),
        },
        "checks": checks,
    }


def _statut_est_traite(statut: str) -> bool:
    """Indique si le statut metier correspond a une session traitee avec succes."""
    return statut.strip().lower() in {"traite", "traité", "traitÃ©"}


def _build_process_response(trigramme: str, code_iris: int, csv_path: Path, statut: str) -> dict[str, object]:
    """Formate la reponse de traitement avec les fichiers produits (bilans session et formation)."""
    session_excel = csv_path.with_suffix(".xlsx")
    formation_excel = EvalStat_formation.construire_chemin_eval_formation(trigramme)
    success = _statut_est_traite(statut)

    return {
        "success": success,
        "status": statut,
        "message": "Session traitee." if success else f"Statut EvalStat : {statut}",
        "codeIris": code_iris,
        "trigramme": trigramme,
        "files": {
            "csv": {"path": str(csv_path), "url": _path_to_uri(csv_path), "name": csv_path.name},
            "sessionExcel": {"path": str(session_excel), "url": _path_to_uri(session_excel), "name": session_excel.name},
            "formationExcel": {"path": str(formation_excel), "url": _path_to_uri(formation_excel), "name": formation_excel.name},
        },
    }


def _process_csv_path(trigramme: str, csv_path: Path) -> dict[str, object]:
    """Verifie un CSV puis lance le traitement EvalStat (bilans session + formation)."""
    verification = _verifier_csv(trigramme, csv_path)
    code_iris = verification["codeIris"]

    if not verification["valid"]:
        is_empty_csv = any(
            check.get("label") == "Fichier non vide"
            and check.get("status") == "error"
            and str(check.get("message", "")).startswith("0 ")
            for check in verification["checks"]
        )
        return {
            "success": False,
            "status": "Ignore" if is_empty_csv else "Erreur",
            "message": "CSV vide : fichier ignore." if is_empty_csv else verification["message"],
            "codeIris": code_iris,
            "trigramme": trigramme,
            "files": {
                "csv": {"path": str(csv_path), "url": _path_to_uri(csv_path), "name": csv_path.name},
            },
            "checks": verification["checks"],
        }

    if code_iris is None:
        return {
            "success": False,
            "status": "Exclu - Code IRIS introuvable",
            "message": "Aucun code IRIS detecte dans le nom du CSV.",
            "codeIris": None,
            "trigramme": trigramme,
            "files": {
                "csv": {"path": str(csv_path), "url": _path_to_uri(csv_path), "name": csv_path.name},
            },
        }

    # Bloc traitement metier : ouvre/cree le classeur formation, traite la
    # session, puis referme systematiquement le classeur formation.
    formation: Formation | None = None
    session: Session | None = None
    try:
        formation = Formation.avec_ouverture_ou_creation_evalStat_formation(trigramme)
        session = Session.avec_ouverture_ou_traitement_evalStat(
            formation=formation,
            code_IRIS=code_iris,
            chemin_csv=csv_path,
            ecrire_eval_formation=True,
            ouvrir_dossier=False,
        )
    except Exception as exc:
        if session is not None and session.eval is not None and session.eval.statut:
            return _build_process_response(trigramme, code_iris, csv_path, session.eval.statut)
        return {
            "success": False,
            "status": "Erreur",
            "message": f"Traitement EvalStat impossible : {exc}",
            "codeIris": code_iris,
            "trigramme": trigramme,
            "files": {
                "csv": {"path": str(csv_path), "url": _path_to_uri(csv_path), "name": csv_path.name},
            },
        }
    finally:
        if formation is not None and formation.eval is not None and formation.eval.fe is not None:
            try:
                formation.eval.fe.close()
            except Exception:
                pass

    eval_session = session.eval if session is not None else None
    statut = eval_session.statut if eval_session is not None else "Inconnu"
    return _build_process_response(trigramme, code_iris, csv_path, statut)


# =============================================================================
# Routes API
# =============================================================================

@router.get("/formations")
def get_formations() -> dict[str, object]:
    """Trigrammes de formation proposes a la recherche depuis l'export IRIS R0304."""
    return {"formations": sorted(_load_iris_formations_codes())}


@router.get("/sessions")
def get_sessions(formation: str) -> dict[str, object]:
    """Sessions EvalStat de `formation`, enrichies depuis IRIS et triees par date decroissante."""
    trigramme = _normaliser_trigramme(formation)
    search_directory = _nearest_search_directory(trigramme)
    formation_dir_exists, formation_dir_message = _formation_directory_status(trigramme)
    if not formation_dir_exists:
        return {
            "formation": trigramme,
            "sessions": [],
            "irisFile": None,
            "message": formation_dir_message,
            "searchDirectory": str(search_directory) if search_directory else None,
        }

    iris_sessions, iris_file = _load_iris_sessions_index()
    csv_by_code = _find_csvs_for_formation(trigramme)

    # Bloc croisement : on part des CSV trouves dans FORMATIONS_C, puis on
    # recupere les informations metier de la session dans l'export IRIS R04110.
    # Ainsi l'interface n'affiche pas tout l'historique IRIS, seulement les
    # sessions qui ont vraiment un CSV EvalStat exploitable.
    rows: list[dict[str, object]] = []
    for code_iris, csv_path in csv_by_code.items():
        row = _session_row_from_csv(trigramme, code_iris, csv_path, iris_sessions)
        if row is not None:
            rows.append(row)

    rows.sort(key=lambda row: row["startDate"], reverse=True)

    return {
        "formation": trigramme,
        "sessions": rows,
        "irisFile": str(iris_file) if iris_file else None,
        "message": (
            None
            if rows
            else f"Aucun CSV EvalStat trouve pour {trigramme} dans FORMATIONS_C."
        ),
        "searchDirectory": str(search_directory) if search_directory else None,
    }


@router.post("/process")
def process_sessions(payload: ProcessSessionsRequest) -> dict[str, object]:
    """Traite les CSV des sessions cochees pour `formation` (bilans session + formation)."""
    trigramme = _normaliser_trigramme(payload.formation)
    rows = [_process_csv_path(trigramme, _uri_to_path(path)) for path in payload.paths]
    return {
        "success": all(bool(row.get("success")) for row in rows) if rows else False,
        "rows": rows,
    }


@router.post("/upload-csv")
async def upload_csv(
    formation: str = Form(...),
    file: UploadFile = File(...),
) -> dict[str, object]:
    """Recoit un CSV choisi dans le navigateur puis renvoie la session prete a cocher."""
    trigramme = _normaliser_trigramme(formation)
    original_name = _safe_upload_name(file.filename or "")
    if not original_name.lower().endswith(".csv"):
        raise HTTPException(status_code=400, detail="Le fichier choisi doit etre un CSV.")

    content = await file.read()
    if not content:
        raise HTTPException(status_code=400, detail="Le CSV est vide.")

    upload_dir = _manual_upload_directory(trigramme)
    upload_dir.mkdir(parents=True, exist_ok=True)
    csv_path = upload_dir / original_name
    copy_index = 1
    while csv_path.exists():
        csv_path = upload_dir / f"{Path(original_name).stem}-copie-{copy_index}.csv"
        copy_index += 1
    csv_path.write_bytes(content)

    verification = _verifier_csv(trigramme, csv_path)
    code_iris = verification["codeIris"]
    if code_iris is None:
        raise HTTPException(status_code=400, detail="Aucun code IRIS trouve.")

    iris_sessions, iris_file = _load_iris_sessions_index()
    session = iris_sessions.get(code_iris)
    if session is None:
        raise HTTPException(status_code=400, detail="Code IRIS absent de R04110.")
    if session.get("formation") != trigramme:
        raise HTTPException(status_code=400, detail="Ce CSV appartient a une autre formation.")

    row = _session_row_from_csv(trigramme, code_iris, csv_path, iris_sessions)
    if row is None:
        raise HTTPException(status_code=400, detail="Session IRIS introuvable.")

    return {
        "formation": trigramme,
        "session": row,
        "irisFile": str(iris_file) if iris_file else None,
    }


@router.post("/open-path")
def open_path(payload: OpenPathRequest) -> dict[str, str]:
    """Ouvre un fichier ou dossier dans l'explorateur Windows depuis le poste serveur."""
    path = _uri_to_path(payload.path)

    try:
        _open_windows_path(path)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Ouverture impossible : {exc}")

    return {"status": "ok", "path": str(path)}
