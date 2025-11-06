from pathlib import Path

### --------------------------------------------------------------------
#  Définitions des constantes de module
### --------------------------------------------------------------------

### -----
# Variables d'environnement
### -----

###
# === Chemins GED ===
###
UNITE = "UEM"

# Répertoires de la GED
GED:Path = Path("//instnt/partage")
REPERTOIRE_FORMATION:Path = GED / "FORMATIONS_C/{trigramme_formation}"

REPERTOIRE_BILANS:Path = REPERTOIRE_FORMATION / "P07-bilan-sessions-et-bilan-formation"
REPERTOIRE_CSV_EVALUATIONS:Path = REPERTOIRE_BILANS / "rapports-sessions-CSV-evaluations"
REPERTOIRE_BILANS_SESSIONS:Path = REPERTOIRE_BILANS / "{annee}"

REPERTOIRE_CONCEPTION:Path = REPERTOIRE_FORMATION / "P05-P06-dossier-conception-referentiel"
REPERTOIRE_FDC:Path = REPERTOIRE_CONCEPTION / "fiche-de-cout-et-code-de-formation"
REPERTOIRE_SPECS:Path = REPERTOIRE_CONCEPTION / "specifications-pedagogiques-et-referentiel"

# Répertoire avec tous les modèles employés par les applications
REPERTOIRES_MODELES = Path("//harmonie/instn/uem/_Echanges/VTE/Prog/Modèles")





###
# === EXTRACT IRIS ===
###
REPERTOIRE_EXTRACT_IRIS_GED:Path = Path("//instnt/HOME/REFERENCE/IRIS - rapports de synthese")
REPERTOIRE_EXTRACT_IRIS_LOCAL:Path = Path("//harmonie/instn/uem/_Echanges/VTE/Prog/IRIS/Extracts originaux")
REPERTOIRE_EXCEL_IRIS_OUTPUT:Path = Path("//harmonie/instn/uem/_Echanges/VTE/Prog/IRIS/Extracts complets")

# Propriétés des exports IRIS
IRIS_SESSIONS_PARAMS = dict(
    nom_typeExport="Sessions",
    codeExport="R04110",

    repertoire_input=REPERTOIRE_EXTRACT_IRIS_LOCAL,  # Pas d'extract sessions dans la GED
    #nom_onglet_input = "Data",  # Si non spécifié, alors nom_onglet_input = "Data"
    nbLignes_avantET_input = 1,  # Si non spécifié, alors nbLignes_avantET_input = 0

    #repertoire_modele = REPERTOIRES_MODELES,  # Si non spécifié, alors repertoire_modele = REPERTOIRES_MODELES
    #nom_fichier_modele = "R04110_Sessions-Modèle.xlsx",  # Si non spécifié, alors nom_fichier_modele = f"{codeExport}_{nom_typeExport}-Modèle.xlsx"

    #repertoire_output = REPERTOIRE_EXCEL_IRIS_OUTPUT,  # Si non spécifié, alors repertoire_output = REPERTOIRE_EXCEL_IRIS_OUTPUT
    #nom_fichier_output = "R04110_Sessions-COMPLET.xlsx"  # Si non spécifié, alors nom_fichier_output = f"{codeExport}_{nom_typeExport}-COMPLET-{date.today():%Y.%m.%d}.xlsx"
)

IRIS_FORMATIONS_PARAMS = dict(
    nom_typeExport="Formations",
    codeExport="R0304",
    repertoire_input=REPERTOIRE_EXTRACT_IRIS_LOCAL  # Pas d'extract sessions dans la GED
)

IRIS_VENTES_PARAMS = dict(
    nom_typeExport="Ventes",
    codeExport="R04301",
    repertoire_input=REPERTOIRE_EXTRACT_IRIS_LOCAL,  # Car il y a des petits bugs sur certains CSV
    nbLignes_avantET_input = 1
)

IRIS_INSCRIPTIONS_PARAMS = dict(
    nom_typeExport="Inscriptions",
    codeExport="R04500",
    repertoire_input=REPERTOIRE_EXTRACT_IRIS_GED,
    nbLignes_avantET_input = 1
)

# Sera rempli après import de instn.py depuis instn.py sinon boucle récursive (méthode .initialiser_PropExportIRIS_de_config() )
IRIS_SESSIONS = None
IRIS_FORMATIONS = None
IRIS_VENTES = None
IRIS_INSCRIPTIONS = None






###
# === EvalStat ===
###
# Modèle qui accueille les CSV des stagiaires 
CHEMIN_MODELE_EXCEL_EVALUATIONS_STAGIAIRES:Path = REPERTOIRES_MODELES / "Evaluation-Stagiaires-Modèle.xlsx"

# Excel d'évaluation des stagiaires pour une formation (i.e. concatène les CSV de retours de toutes les sessions)
CHEMIN_EXCEL_EVALUATIONS_FORMATION:Path = REPERTOIRE_CSV_EVALUATIONS / "Evaluation-Stagiaires-Global-{trigramme_formation}.xlsx"








###
# === Bilan de session ===
###
CHEMIN_MODELE_WORD_BILAN_SESSION:Path = REPERTOIRES_MODELES / "P07-Pr05-F05-Bilan-session-V3.docx"
CHEMIN_WORD_BILAN_SESSION_OUTPUT:Path = REPERTOIRE_BILANS_SESSIONS / "P07-Pr05-F05-Bilan session-{periode}-{unite}.docx"


ADRESSE_MAIL_CHEF_UNITE:str = "florent.lemont@cea.fr"
CORPS_MAIL_CHEF_UNITE:str = """
    <p>Bonjour Florent,</p>
    <p>Est-ce que tu peux signer le bilan de session ci-dessous stp.<br>
    Lien du bilan de session : <a href="{lien_pdf_bilan}">{lien_pdf_bilan}</a></p> 
    <p>Il concerne la formation {formation} : {periode}.</p>
"""









###
# === Bilan de formation ===
###
# Modèle du bilan à remplir
CHEMIN_MODELE_WORD_BILAN_FORMATION:Path = REPERTOIRES_MODELES / "P07-Pr05-F06-Bilan-formation-V3.docx"
CHEMIN_WORD_BILAN_FORMATION_OUTPUT:Path = REPERTOIRE_BILANS / "P07-Pr05-F06-Bilan formation-Année {annee}.docx"

# Bilan en sortie après remplissage
chemin_word_bilan_formation_output = r'C:\Users\vt238770\Documents\_CEA\Prog\Modèles\Bilan formation - output.docx'





###
# === GETTER ===
###
def format_path(path: Path, **kwargs) -> Path:
    """
    Formate un chemin contenant des placeholders avec les valeurs fournies.
    
    Exemple :
        format_path(Path("C:/data/{unite}/{periode}.xlsx"), unite="UGA", periode="2025.11.01")
        → Path("C:/data/UGA/2025.11.01.xlsx")
    """
    return Path(str(path).format(**kwargs))

#def chemin_modele_excel_evaluations(trigramme_formation: str) -> Path:
#    """Retourne le Path complet pour le fichier Excel d’évaluation."""
#    return Path(str(CHEMIN_EXCEL_EVALUATIONS_FORMATION).format(trigramme_formation=trigramme_formation))

#def chemin_word_bilan_session_output(trigramme_formation:str, periode:str, unite:str=config.UNITE) -> Path:
#    return Path(str(CHEMIN_WORD_BILAN_SESSION_OUTPUT).format(trigramme_formation=trigramme_formation, periode=periode, unite=unite))





"""
# ANCIENNES DEFINITIONS

IRIS_SESSIONS:PropExportIRIS = PropExportIRIS(
    nom_typeExport = "Sessions",
    codeExport = "R04110",

    repertoire_input = REPERTOIRE_EXTRACT_IRIS_LOCAL,  # Pas d'extract sessions dans la GED
    nom_onglet_input = "Data",
    nbLignes_avantET_input = 1,

    #repertoire_modele = REPERTOIRES_MODELES,  # Si non spécifié, alors repertoire_modele = REPERTOIRES_MODELES
    #nom_fichier_modele = "R04110_Sessions-Modèle.xlsx",  # Si non spécifié, alors nom_fichier_modele = f"{codeExport}_{nom_typeExport}-Modèle.xlsx"

    #repertoire_output = REPERTOIRE_EXCEL_IRIS_OUTPUT,  # Si non spécifié, alors repertoire_output = REPERTOIRE_EXCEL_IRIS_OUTPUT
    #nom_fichier_output = "R04110_Sessions-COMPLET.xlsx"  # Si non spécifié, alors nom_fichier_output = f"{codeExport}_{nom_typeExport}-COMPLET-{date.today():%Y.%m.%d}.xlsx"
    )

IRIS_FORMATIONS:PropExportIRIS = PropExportIRIS(
    nom_typeExport = "Formations",
    codeExport = "R0304",

    repertoire_input = REPERTOIRE_EXTRACT_IRIS_LOCAL,  # Pas d'extract formations dans la GED
    nom_onglet_input = "Data",
    nbLignes_avantET_input = 0,

    nom_fichier_modele = "R0304_Formations-Modèle.xlsx",

    nom_fichier_output = "R0304_Formations-COMPLET.xlsx"
    )

IRIS_VENTES:PropExportIRIS = PropExportIRIS(
    nom_typeExport = "Ventes",
    codeExport = "R04301",

    repertoire_input = REPERTOIRE_EXTRACT_IRIS_LOCAL,  # Car il y a des petits bugs sur certains CSV
    nom_onglet_input = "Data",
    nbLignes_avantET_input = 1,

    nom_fichier_modele = "R04301_Ventes-Modèle.xlsx",

    nom_fichier_output = "R04301_Ventes-COMPLET.xlsx"
    )

IRIS_INSCRIPTIONS:PropExportIRIS = PropExportIRIS(
    nom_typeExport = "Inscriptions",
    codeExport = "R04500",

    repertoire_input = REPERTOIRE_EXTRACT_IRIS_GED,
    nom_onglet_input = "Data",
    nbLignes_avantET_input = 1,

    nom_fichier_modele = "R04500_Inscriptions-Modèle.xlsx",

    nom_fichier_output = "R04500_Inscriptions-COMPLET.xlsx"
    )

"""