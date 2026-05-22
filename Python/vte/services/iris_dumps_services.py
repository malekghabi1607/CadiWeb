import re
import shutil
from datetime import date
from pathlib import Path

from vte.core import config


# ================================================
# === Liste des fichiers Excel disponibles IRIS ===
# ================================================
def lister_fichiers_excel(repertoire: Path | str) -> list[Path]:
    """
    Liste les fichiers Excel .xlsx dans un dossier.

    Exemple :
        >>> fichiers = lister_fichiers_excel("R:/_Echanges/VTE/Prog/IRIS/Extracts originaux")
        >>> [fichier.name for fichier in fichiers]
        ["R04110_Sessions-2026 au 2026.05.20.xlsx"]

    :param repertoire: chemin du dossier dans lequel chercher les fichiers Excel.
    :type repertoire: Path | str
    :raises FileNotFoundError: si le repertoire demande n'existe pas.
    :raises NotADirectoryError: si le chemin donne existe mais n'est pas un dossier.
    :return: liste triee des chemins des fichiers Excel trouves.
    :rtype: list[Path]
    """
    # Convertit le chemin en Path pour travailler proprement avec les fichiers.
    chemin_repertoire = Path(repertoire)

    # Verifie que le dossier existe.
    if not chemin_repertoire.exists():
        raise FileNotFoundError(f"Repertoire introuvable : {chemin_repertoire}")

    # Verifie que le chemin donne est bien un dossier.
    if not chemin_repertoire.is_dir():
        raise NotADirectoryError(f"Le chemin n'est pas un repertoire : {chemin_repertoire}")

    # Retourne seulement les fichiers Excel, tries par nom.
    return sorted(chemin_repertoire.glob("*.xlsx"), key=lambda fichier: fichier.name.lower())


# ================================================
# === Filtrage des fichiers par code export IRIS ===
# ================================================
def filtrer_par_code(fichiers: list[Path], code_export: str) -> list[Path]:
    """
    Garde seulement les fichiers qui commencent par un code export IRIS.

    Codes principaux :
        - R04110 : sessions ;
        - R0304  : referentiel formations ;
        - R04301 : ventes ;
        - R04500 : inscriptions.

    Exemple :
        >>> fichiers = [
        ...     Path("R04301_Sessions-Ventes-FC2025 FINAL.xlsx"),
        ...     Path("R04500_Sessions-Inscriptions-FC2025 FINAL.xlsx"),
        ... ]
        >>> ventes = filtrer_par_code(fichiers, "R04301")
        >>> [fichier.name for fichier in ventes]
        ["R04301_Sessions-Ventes-FC2025 FINAL.xlsx"]

    :param fichiers: liste des fichiers a filtrer.
    :type fichiers: list[Path]
    :param code_export: code IRIS attendu au debut du nom de fichier.
    :type code_export: str
    :return: fichiers dont le nom commence par le code IRIS demande.
    :rtype: list[Path]
    """
    # Nettoie le code si l'utilisateur a mis des espaces.
    code_normalise = code_export.strip()

    # Garde seulement les fichiers qui commencent par le code demande.
    return [
        fichier
        for fichier in fichiers
        if fichier.name.startswith(code_normalise)
    ]

# ================================================
# === Extraction de date depuis un nom de fichier ===
# ================================================
def extraire_date_depuis_nom(nom_fichier: str) -> date | None:
    """
    Extrait une date depuis un nom de fichier IRIS.

    Formats reconnus :
        - 2026-04-01
        - 2026.05.20

    Exemple :
        >>> extraire_date_depuis_nom("R04301_Sessions-Ventes au 2026-04-01.xlsx")
        datetime.date(2026, 4, 1)

    :param nom_fichier: nom du fichier dans lequel chercher une date.
    :type nom_fichier: str
    :return: date trouvee, ou None si aucune date n'est presente.
    :rtype: date | None
    """
    # Cherche une date au format AAAA-MM-JJ ou AAAA.MM.JJ.
    motif_date = r"(\d{4})[.-](\d{2})[.-](\d{2})"
    dates_trouvees = re.findall(motif_date, nom_fichier)

    if not dates_trouvees:
        return None

    # Si plusieurs dates existent dans le nom, on garde la derniere.
    annee, mois, jour = dates_trouvees[-1]
    return date(int(annee), int(mois), int(jour))


# ================================================
# === Recherche du dump IRIS le plus recent ===
# ================================================
def trouver_dernier_dump(fichiers: list[Path]) -> Path | None:
    """
    Trouve le fichier le plus recent dans une liste de dumps IRIS.

    La date est lue dans le nom du fichier avec `extraire_date_depuis_nom`.
    Les fichiers sans date sont ignores.

    Exemple :
        >>> fichiers = [
        ...     Path("R04301_Sessions-Ventes-filtre sur FC2026 au 2026-02-02 LG.xlsx"),
        ...     Path("R04301_Sessions-Ventes-filtre sur FC2026 au 2026-04-01 LG.xlsx"),
        ... ]
        >>> trouver_dernier_dump(fichiers).name
        "R04301_Sessions-Ventes-filtre sur FC2026 au 2026-04-01 LG.xlsx"

    :param fichiers: liste de fichiers IRIS a comparer.
    :type fichiers: list[Path]
    :return: fichier le plus recent, ou None si aucun fichier n'a de date.
    :rtype: Path | None
    """
    # Stocke les couples (date trouvee, fichier correspondant).
    fichiers_dates: list[tuple[date, Path]] = []

    # Parcourt tous les fichiers pour extraire la date presente dans leur nom.
    for fichier in fichiers:
        date_fichier = extraire_date_depuis_nom(fichier.name)

        # Si aucune date n'est trouvee, le fichier est ignore.
        if date_fichier is not None:
            fichiers_dates.append((date_fichier, fichier))

    # Si aucun fichier ne contient de date, on ne peut pas choisir le plus recent.
    if not fichiers_dates:
        return None

    # Recupere le fichier associe a la date la plus recente.
    return max(fichiers_dates, key=lambda element: element[0])[1]


# ================================================
# === Construction de la liste de fichiers IRIS ===
# ================================================
def construire_liste_fichiers_type(repertoire: Path | str, code_export: str) -> tuple[str, ...]:
    """
    Construit la liste des fichiers a utiliser pour un type d'export IRIS.

    La regle est simple :
        - garder tous les fichiers historiques contenant "FINAL" ;
        - garder seulement le dernier fichier courant date ;
        - retourner les noms des fichiers, comme dans `config_extractsIRIS.py`.
        

    Exemple :
        >>> construire_liste_fichiers_type(
        ...     "R:/_Echanges/VTE/Prog/IRIS/Extracts originaux",
        ...     "R04301",
        ... )
        (
            "R04301_Sessions-Ventes-FC2025 FINAL.xlsx",
            "R04301_Sessions-Ventes-filtre sur FC2026 au 2026-04-01 LG.xlsx",
        )

    :param repertoire: dossier contenant les dumps IRIS originaux.
    :type repertoire: Path | str
    :param code_export: code export IRIS a traiter, par exemple R04301.
    :type code_export: str
    :return: noms des fichiers a utiliser pour ce type d'export.
    :rtype: tuple[str, ...]
    """
    # Liste tous les fichiers Excel du dossier.
    fichiers_excel = lister_fichiers_excel(repertoire)

    # Garde uniquement les fichiers du type demande.
    fichiers_du_type = filtrer_par_code(fichiers_excel, code_export)

    # Les fichiers FINAL sont les historiques a conserver.
    fichiers_final = [
        fichier
        for fichier in fichiers_du_type
        if "FINAL" in fichier.name.upper()
    ]

    # Les fichiers courants sont ceux qui ne sont pas marques FINAL.
    fichiers_courants = [
        fichier
        for fichier in fichiers_du_type
        if "FINAL" not in fichier.name.upper()
    ]

    # Parmi les fichiers courants, on garde uniquement le plus recent.
    dernier_fichier_courant = trouver_dernier_dump(fichiers_courants)

    fichiers_a_utiliser = fichiers_final.copy()
    if dernier_fichier_courant is not None:
        fichiers_a_utiliser.append(dernier_fichier_courant)

    # Retourne seulement les noms, tries par nom pour garder un ordre stable.
    return tuple(
        fichier.name
        for fichier in sorted(fichiers_a_utiliser, key=lambda fichier: fichier.name.lower())
    )


# ================================================
# === Copie du dernier dump depuis la GED IRIS ===
# ================================================
def copier_dernier_dump_depuis_ged(
    code_export: str,
    repertoire_source: Path | str = config.REPERTOIRE_EXTRACT_IRIS_GED,
    repertoire_destination: Path | str = config.REPERTOIRE_EXTRACT_IRIS_LOCAL,
) -> Path | None:
    """
    Copie le dernier dump IRIS disponible depuis la GED vers le dossier local.

    Cette fonction sert surtout pour les exports qui arrivent automatiquement
    dans la GED, par exemple :
        - R04301 : ventes ;
        - R04500 : inscriptions.

    Elle cherche le fichier le plus recent pour le code demande, puis le copie
    dans le dossier local des extracts originaux.

    Exemple :
        >>> copier_dernier_dump_depuis_ged("R04301")
        Path("R:/_Echanges/VTE/Prog/IRIS/Extracts originaux/R04301_...")

    :param code_export: code export IRIS a recuperer, par exemple R04301.
    :type code_export: str
    :param repertoire_source: dossier GED dans lequel chercher le dump.
    :type repertoire_source: Path | str
    :param repertoire_destination: dossier local ou copier le dump.
    :type repertoire_destination: Path | str
    :return: chemin du fichier copie, ou None si aucun dump date n'est trouve.
    :rtype: Path | None
    """
    # Liste les fichiers Excel disponibles dans la source.
    fichiers_source = lister_fichiers_excel(repertoire_source)

    # Garde seulement les fichiers du code demande.
    fichiers_du_type = filtrer_par_code(fichiers_source, code_export)

    # Selectionne le fichier le plus recent.
    dernier_dump = trouver_dernier_dump(fichiers_du_type)
    if dernier_dump is None:
        return None

    chemin_destination = Path(repertoire_destination)

    # Verifie que le dossier de destination existe.
    if not chemin_destination.exists():
        raise FileNotFoundError(f"Repertoire destination introuvable : {chemin_destination}")

    if not chemin_destination.is_dir():
        raise NotADirectoryError(f"La destination n'est pas un repertoire : {chemin_destination}")

    fichier_destination = chemin_destination / dernier_dump.name

    # Copie le fichier en conservant les metadonnees autant que possible.
    shutil.copy2(dernier_dump, fichier_destination)

    return fichier_destination


# ================================================
# === Archivage des anciens dumps IRIS ===
# ================================================
def _chemin_archive_disponible(repertoire_archives: Path, nom_fichier: str) -> Path:
    """
    Retourne un chemin d'archive disponible sans ecraser un fichier existant.
    """
    chemin_archive = repertoire_archives / nom_fichier

    if not chemin_archive.exists():
        return chemin_archive

    compteur = 1
    while True:
        nouveau_nom = f"{chemin_archive.stem}_{compteur}{chemin_archive.suffix}"
        nouveau_chemin = repertoire_archives / nouveau_nom
        if not nouveau_chemin.exists():
            return nouveau_chemin
        compteur += 1


def archiver_anciens_dumps(
    code_export: str,
    repertoire: Path | str = config.REPERTOIRE_EXTRACT_IRIS_LOCAL,
    repertoire_archives: Path | str | None = None,
) -> list[Path]:
    """
    Archive les anciens dumps courants pour un code export IRIS.

    Regle appliquee :
        - les fichiers contenant "FINAL" sont conserves ;
        - le fichier courant le plus recent est conserve ;
        - les autres fichiers courants dates sont deplaces dans Archives ;
        - les fichiers sans date sont laisses en place par securite.

    Exemple :
        >>> archiver_anciens_dumps("R04301")
        [Path("R:/_Echanges/VTE/Prog/IRIS/Extracts originaux/Archives/R04301_...xlsx")]

    :param code_export: code export IRIS a nettoyer, par exemple R04301.
    :type code_export: str
    :param repertoire: dossier contenant les dumps IRIS originaux.
    :type repertoire: Path | str
    :param repertoire_archives: dossier d'archive. Par defaut : repertoire / "Archives".
    :type repertoire_archives: Path | str | None
    :return: liste des chemins des fichiers archives.
    :rtype: list[Path]
    """
    chemin_repertoire = Path(repertoire)

    # Liste les fichiers du type demande.
    fichiers_excel = lister_fichiers_excel(chemin_repertoire)
    fichiers_du_type = filtrer_par_code(fichiers_excel, code_export)

    # On ne nettoie que les fichiers courants, jamais les historiques FINAL.
    fichiers_courants = [
        fichier
        for fichier in fichiers_du_type
        if "FINAL" not in fichier.name.upper()
    ]

    # Le plus recent reste dans le dossier principal.
    dernier_dump = trouver_dernier_dump(fichiers_courants)
    if dernier_dump is None:
        return []

    # Les fichiers sans date sont gardes pour eviter un archivage risqué.
    fichiers_a_archiver = [
        fichier
        for fichier in fichiers_courants
        if fichier != dernier_dump and extraire_date_depuis_nom(fichier.name) is not None
    ]

    if not fichiers_a_archiver:
        return []

    chemin_archives = Path(repertoire_archives) if repertoire_archives else chemin_repertoire / "Archives"
    chemin_archives.mkdir(exist_ok=True)

    fichiers_archives: list[Path] = []
    for fichier in fichiers_a_archiver:
        destination = _chemin_archive_disponible(chemin_archives, fichier.name)
        shutil.move(str(fichier), str(destination))
        fichiers_archives.append(destination)

    return fichiers_archives


# ================================================
# === Mise a jour complete des dumps depuis GED ===
# ================================================
def mettre_a_jour_dumps_depuis_ged(
    codes_exports: tuple[str, ...] = ("R04301", "R04500"),
    repertoire_source: Path | str = config.REPERTOIRE_EXTRACT_IRIS_GED,
    repertoire_destination: Path | str = config.REPERTOIRE_EXTRACT_IRIS_LOCAL,
) -> dict[str, dict[str, Path | list[Path] | None]]:
    """
    Met a jour les dumps IRIS recuperes depuis la GED.

    Pour chaque code export :
        - copie le dernier dump disponible depuis la GED ;
        - archive les anciens dumps courants dans le dossier local.

    Par defaut, cette fonction traite :
        - R04301 : ventes ;
        - R04500 : inscriptions.

    Exemple :
        >>> mettre_a_jour_dumps_depuis_ged()
        {
            "R04301": {"copie": Path(...), "archives": [Path(...)]},
            "R04500": {"copie": Path(...), "archives": []},
        }

    :param codes_exports: codes IRIS a mettre a jour.
    :type codes_exports: tuple[str, ...]
    :param repertoire_source: dossier GED contenant les dumps source.
    :type repertoire_source: Path | str
    :param repertoire_destination: dossier local des extracts originaux.
    :type repertoire_destination: Path | str
    :return: bilan des fichiers copies et archives par code export.
    :rtype: dict[str, dict[str, Path | list[Path] | None]]
    """
    bilan: dict[str, dict[str, Path | list[Path] | None]] = {}

    for code_export in codes_exports:
        # Copie le dernier fichier disponible dans la GED.
        fichier_copie = copier_dernier_dump_depuis_ged(
            code_export=code_export,
            repertoire_source=repertoire_source,
            repertoire_destination=repertoire_destination,
        )

        # Archive les anciens fichiers courants du dossier local.
        fichiers_archives = archiver_anciens_dumps(
            code_export=code_export,
            repertoire=repertoire_destination,
        )

        bilan[code_export] = {
            "copie": fichier_copie,
            "archives": fichiers_archives,
        }

    return bilan
