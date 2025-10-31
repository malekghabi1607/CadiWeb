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

# Répertoires de la GED
GED = Path("//instnt/partage")
REPERTOIRE_FORMATION = GED / "FORMATIONS_C/{trigramme_formation}"
REPERTOIRE_CSV_EVALUATIONS = REPERTOIRE_FORMATION / "P07-bilan-sessions-et-bilan-formation/rapports-sessions-CSV-evaluations"

# Répertoire avec tous les modèles employés par les applications
REPERTOIRES_MODELES = Path("//harmonie/instn/uem/_Echanges/VTE/Prog/Modèles")

# Excel d'évaluation des stagiaires pour une formation (i.e. concatène les CSV de retours de toutes les sessions)
NOM_EXCEL_EVALUATIONS_FORMATION = "Evaluation-Stagiaires-Global-{trigramme_formation}.xlsx"
CHEMIN_EXCEL_EVALUATIONS_FORMATION = REPERTOIRE_CSV_EVALUATIONS / NOM_EXCEL_EVALUATIONS_FORMATION

###
# === EXTRACT IRIS ===
###
REPERTOIRE_EXCEL_IRIS_SESSIONS:str = r"\\harmonie\instn\uem\_Echanges\VTE\Prog\IRIS\Extracts complets"

###
# === EvalStat ===
###
# Modèle qui accueille les CSV des stagiaires 
CHEMIN_MODELE_EXCEL_EVALUATIONS_STAGIAIRES = REPERTOIRES_MODELES / "Evaluation-Stagiaires-Modèle.xlsx"


###
# === Bilan de session ===
###
CHEMIN_WORD_BILAN_INPUT = REPERTOIRES_MODELES / "P07-Pr05-F05-Bilan-session-V3.docx"
REPERTOIRE_WORD_BILAN_OUTPUT = REPERTOIRE_FORMATION / "P07-bilan-sessions-et-bilan-formation"

ADRESSE_MAIL_CHEF_UNITE:str = "florent.lemont@cea.fr"
CORPS_MAIL_CHEF_UNITE:str = """
    <p>Bonjour Florent,</p>
    <p>Est-ce que tu peux signer le bilan de session ci-dessous stp.\nLien du bilan de session : <a href="{lien_pdf_bilan}">{lien_pdf_bilan}</a></p> 
    <p>Il concerne la formation {formation} : {periode}.</p>
"""

CRITERES_A_ENLEVER = [  # Critères à ne pas retenir pour le calcul des moyennes < 3
    "Comment avez-vous connu cette formation ?", "Avez-vous d'autres besoins de formation ?", "Commentaires, remarques, suggestions", "Recommanderiez-vous cette formation ?"]



###
# === getter ===
###
def chemin_modele_excel_evaluations(trigramme_formation: str) -> Path:
    """Retourne le Path complet pour le fichier Excel d’évaluation."""
    return Path(str(CHEMIN_EXCEL_EVALUATIONS_FORMATION).format(trigramme_formation=trigramme_formation))



