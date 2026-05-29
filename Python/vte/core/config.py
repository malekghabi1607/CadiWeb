from pathlib import Path

_RACINE_PROJET = Path(__file__).resolve().parents[2]


def _chemin_avec_fallback(chemin_reseau: str, chemin_local: Path) -> Path:
    chemin = Path(chemin_reseau)
    return chemin if chemin.exists() else chemin_local

### --------------------------------------------------------------------
#  Définitions des constantes de module
### --------------------------------------------------------------------


###
# === DIVERS ===
###
UNITE = "UEM"

# Répertoire avec tous les modèles employés par les applications
REPERTOIRES_MODELES = Path("//harmonie/instn/uem/_Echanges/VTE/Prog/Modèles")



###
# === REPERTOIRES GED - PLAN DE CLASSEMENT ===
###

GED:Path = Path("//instnt/partage")
REPERTOIRE_FORMATION:Path = GED / "FORMATIONS_C/{trigramme_formation}"

REPERTOIRE_BILANS:Path = REPERTOIRE_FORMATION / "P07-bilan-sessions-et-bilan-formation"
REPERTOIRE_CSV_EVALUATIONS:Path = REPERTOIRE_BILANS / "rapports-sessions-CSV-evaluations"
REPERTOIRE_BILANS_SESSIONS:Path = REPERTOIRE_BILANS / "{annee}"
REPERTOIRE_BILANS_FORMATIONS:Path = REPERTOIRE_BILANS / "{annee}"

REPERTOIRE_CONCEPTION:Path = REPERTOIRE_FORMATION / "P05-P06-dossier-conception-referentiel"
REPERTOIRE_FDC:Path = REPERTOIRE_CONCEPTION / "fiche-de-cout-et-code-de-formation"
REPERTOIRE_SPECS:Path = REPERTOIRE_CONCEPTION / "specifications-pedagogiques-et-referentiel"


FICHIER_FDC:Path = "Fiche de coûts INSTN - {trigramme_formation} - {unite} - {date_aaaa_mm_jj}.xlsx"
REGEX_FDC:str = "Fiche de coûts INSTN - .*"



###
# === EXTRACT IRIS ===
###
REPERTOIRE_EXTRACT_IRIS_GED:Path = Path("//instnt/HOME/REFERENC/IRIS - rapports de synthese")  # Répertoire extracts originaux
REPERTOIRE_EXTRACT_IRIS_LOCAL:Path = _chemin_avec_fallback(
    "//harmonie/instn/uem/_Echanges/VTE/Prog/IRIS/Extracts originaux",
    _RACINE_PROJET / "IRIS" / "Extracts originaux",
)  # Répertoire local extracts originaux (à causes de 2/3 adaptations d'exports mal fichus/buggés et que je ne pouvais pas mettre moi sur la GED après réparation)
REPERTOIRE_EXCEL_IRIS_OUTPUT:Path = _chemin_avec_fallback(
    "//harmonie/instn/uem/_Echanges/VTE/Prog/IRIS/Extracts complets",
    _RACINE_PROJET / "IRIS" / "Extracts complets",
)  # Répertoire avec les extracts concaténés et dans les modèles

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
    nbLignes_avantET_input = 1,

    ordre_colonnes_modele=[
        "N° Session", 
        "Intitulé Session", 
        "Trigramme formation", 
        "Code IRIS", 
        "Type de formation", 
        "Année", 
        "Trigramme RP", 
        "Trigramme AF", 
        "3ème élément de la référence", 
        "Statut Session", 
        "Lieu Session", 
        "Modalité", 
        "Type", 
        "Resp pédagogique", 
        "Affectation RP", 
        "Organisatrice", 
        "Domaine parent", 
        "Participants MIN", 
        "Participants MAX", 
        "Durée (H) Session", 
        "Durée (J) Session", 
        "Date Début Session", 
        "Année Début Session", 
        "Mois Début Session", 
        "Date Fin Session", 
        "Année Fin Session", 
        "Mois Fin Session", 
        "Gestionnaire Session", 
        "Lieu de formation", 
        "Chef de projet", 
        "Code Formation", 
        "Ref. Formation", 
        "Intitulé Formation", 
        "Spécialité Formation", 
        "Code Domaine", 
        "Ref. Domaine", 
        "Domaine", 
        "Ref. Domaine principal", 
        "Domaine principal", 
        "Ref. Org. Facturation", 
        "Intitulé Org. Facturation", 
        "Dossier N°", 
        "Statut Dossier", 
        "N°Cde", 
        "Statut Cde", 
        "Organisme", 
        "Commercial", 
        "Client", 
        "Secteur d'activité", 
        "Catégorie  client", 
        "Fidélité", 
        "Autre critère", 
        "Contact Client", 
        "Mail contact client", 
        "Fonction contact client", 
        "Civilité stagiaire", 
        "Nom Stagiaire", 
        "Prénom Stagiaire", 
        "Sexe Stagiaire", 
        "Age Stagiaire", 
        "Date de naissance", 
        "Mail Stagiaire", 
        "Nationalité Stagiaire", 
        "Etablissement Stagiaire", 
        "SIRET", 
        "Entité Juridique Stagiaire", 
        "Affect. CEA / Société", 
        "CSP  Stagiaire", 
        "Contrat Stagiaire", 
        "Fonction Stagiaire", 
        "Référence Stagiaire", 
        "Motif annulation", 
        "Statut de la qualif.", 
        "Pédagogie terminée", 
        "Financier terminé", 
        "Planifiée (H)", 
        "Qualifiée (H)", 
        "Réalisée (H)", 
        "Planifiée (J)", 
        "Qualifiée (J)", 
        "Réalisée (J)", 
        "Prévu", 
        "Réalisé", 
        "Facturé", 
        "A facturer", 
        "Réglé", 
    ]
)






###
# === EvalStat ===
###
# Modèle qui accueille les CSV des stagiaires 
CHEMIN_MODELE_EXCEL_EVALUATIONS_STAGIAIRES:Path = REPERTOIRES_MODELES / "Evaluation-Stagiaires-Modèle.xlsx"

# Excel d'évaluation des stagiaires pour une formation (i.e. concatène les CSV de retours de toutes les sessions)
CHEMIN_EXCEL_EVALUATIONS_FORMATION:Path = REPERTOIRE_CSV_EVALUATIONS / "Evaluation-Stagiaires-Global-{trigramme_formation}.xlsx"








###
# === Bilan de sessions ===
###
CHEMIN_MODELE_WORD_BILAN_SESSIONS:Path = REPERTOIRES_MODELES / "P07-Pr05-F05-Bilan-session-V3.docx"
CHEMIN_WORD_BILAN_SESSIONS_OUTPUT:Path = REPERTOIRE_BILANS_SESSIONS / "P07-Pr05-F05-Bilan session-{periode}-{unite}.docx"


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
CHEMIN_MODELE_WORD_BILAN_FORMATION:Path = REPERTOIRES_MODELES / "P07-Pr05-F06-Bilan-formation-V3_VTE.docx"
CHEMIN_WORD_BILAN_FORMATION_OUTPUT:Path = REPERTOIRE_BILANS_FORMATIONS / "P07-Pr05-F06-Bilan formation-Année {annee}.docx"

# Bilan en sortie après remplissage
#chemin_word_bilan_formation_output = r'C:\Users\vt238770\Documents\_CEA\Prog\Modèles\Bilan formation - output.docx'








###
# === REE ===
###
REPERTOIRE_DOCUMENTS_REE:Path = REPERTOIRES_MODELES / "REE"  # Répertoire contenant l'ensemble des documents et modèles pour la gestion des REE
CHEMIN_MAIL_DEMANDE_INFOS_ADMIN_REE:Path = REPERTOIRES_MODELES / "REE" / "Demande des informations administratives.msg"  # Message type à envoyer aux intervenants
CHEMIN_MODELE_EXCEL_FICHE_INTERVENANT:Path = REPERTOIRES_MODELES / "REE" / "P09_Pr01_Ta.E_Grille des critères de qualification des compétences_V1.xlsx"    # Fichier Excel à remplir pour Laetitia Da Mota (RH INSTN qui s'occupe de rentrer les REE dans IRIS)

REPERTOIRE_SAUVEGARDE_FICHIERS_REE:Path = Path("//harmonie/INSTN/UEM/_Documents_communs/Formations/Formateurs/1.Intervenants - Documents administratifs") # Lieu où sauvegarder localement les fichiers de l'intervenant (CV, fiches administrative, autorisation employeur...)

ADRESSE_MAIL_GESTIONNAIRE_REE_INSTN:str = "vacataires.instn@cea.fr"  # Adresse mail du gestionnaire INSTN des REE (pour envoi enregistrement IRIS : Laëtitia Da Mota)
CORPS_MAIL_GESTIONNAIRE_REE_INSTN:str = "<p>Bonjour Laëtitia,</p><p>Je t’ai mis en PJ les documents pour intégrer/mettre à jour la fiche IRIS de {Prenoms} {NOM}.</p><p>Je te remercie, passe une excellente journée :)</p>"  # Corps de mail pour l'envoi au gestionnaire INSTN des REE 
    




###
# === REE ===
###
UNITES = {
    "INSTN Saclay" : "UES",
    "INSTN Marcoule" : "UEM",
    "INSTN Cadarache" : "UEM",
    "INSTN Grenoble" : "UEG",
    "INSTN Cherbourg-Octeville" : "UECC",
    "INSTN Cherbourg-en-Cotentin" : "UECC",

    "Plateforme DOSEO" : "DOSEO",
    "École du sodium" : "École du sodium",
    "Cluses" : "Cluses",
    "St-Etienne" : "St-Etienne",
    "Bureau Formation Cadarache" : "BF CAD",
    "Bourges" : "Bourges",
    "Distanciel" : "Distanciel",
    "SANOFI Chilly-Mazarin" : "SANOFI Chilly-Mazarin",
    "PARIS - Plateforme IHU ICAN - Pitié Salpêtrière" : "Pitié Salpêtrière",
    "PARIS - Plate-forme d'Imagerie du Petit Animal - Hôpital Cochin" : "Hôpital Cochin",
    "E-learning" : "E-learning",
    "PARIS - Plate-forme d'imageries du vivant (PIV) - Hôpital Cochin" : "Hôpital Cochin",
    "SANOFI Vitry" : "SANOFI Vitry",
    "EDF Itech" : "EDF Itech",
}

###
# === GETTER ===
###
def format_path(path: Path, **kwargs) -> Path:
    """
    Formate un chemin contenant des placeholders avec les valeurs fournies.
    
    Exemple :
        >>> config.format_path(
                config.CHEMIN_WORD_BILAN_SESSION_OUTPUT, 
                trigramme_formation=instance._codeFormation, 
                annee=instance._annee, 
                periode=f"{numSession}", 
                unite=config.UNITE)
        >>> format_path(Path("C:/data/{unite}/{periode}.xlsx"), unite="UGA", periode="2025.11.01") → Path("C:/data/UGA/2025.11.01.xlsx")
    """
    return Path(str(path).format(**kwargs))
