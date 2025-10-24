from dataclasses import dataclass
from .instn import *


_chemin_excel_evaluations_defaut:str = r'\\instnt\partage\FORMATIONS_C\###\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\Evaluation-Stagiaires-Global-###.xlsx'
_chemin_modeleExcel_stagiaires:str = r"C:\Users\vt238770\Documents\_CEA\Prog\Python\Modèles\Evaluation-Stagiaires-Modèle.xlsx"
_chemin_excel_sessions:str = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts complets\R04110_Sessions-COMPLET-2025.08.24.xlsx" # TODO faire une méthode pour chercher automatiquement le dernier fichier
_repertoire_excel_sessions:str = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts complets\"


# Initialisation des chemins des répertoires
#rep_gedMiroir = r"\\instnt\partage\FORMATIONS_C"
rep_fdc_defaut = r"\\instnt\partage\FORMATIONS_C\XXX\P05-P06-dossier-conception-referentiel\fiche-de-cout-et-code-de-formation"
rep_specsPedagogiques_defaut = r"\\instnt\partage\FORMATIONS_C\XXX\P05-P06-dossier-conception-referentiel\specifications-pedagogiques-et-referentiel"


# Modèle du bilan à remplir
chemin_word_bilan_input = r'C:\Users\vt238770\Documents\_CEA\Prog\Modèles\Bilan formation.docx'

# Bilan en sortie après remplissage
chemin_word_bilan_output = r'C:\Users\vt238770\Documents\_CEA\Prog\Modèles\Bilan formation - output.docx'

# Fiche de coûts
#chemin_fdc = r'\\instnt\partage\FORMATIONS_C\948\P05-P06-dossier-conception-referentiel\fiche-de-cout-et-code-de-formation\Fiche de coûts - 948 - Elaboration de scénarios de DEM - 2025.01.24.xlsx'

# Specs pédagogiques
#chemin_specsPedagogiques = r'\\instnt\partage\FORMATIONS_C\948\P05-P06-dossier-conception-referentiel\specifications-pedagogiques-et-referentiel\P06-Pr01-F01_Specifications-pedagogiques - 948 - 2025.04.pdf'






# FdC
nomOnglet_fdc = "Fiche de coûts"




# === Traiter_contactsApprentis ===
# Uniquemlent lz catégorie :
@dataclass
class PropEntretien:
    sujet: str
    chemin_modele: str
    duree: timedelta
    date_debut:datetime
    categorie: str = "FI"
    
    def __post_init__(self):
        # Normaliser en datetime
        if isinstance(self.date_debut, date) and not isinstance(self.date_debut, datetime):
            self.date_debut = datetime.combine(self.date_debut, time(9, 0))




# --- UGA ---
annee_scolaire = "2025-2026"
envoyer_mail = False

prefixe_sujet = "Master IN - Suivi d'alternance"
mail_responsables_univ = "master-in-responsables@univ-grenoble-alpes.fr"

chemin_modele_mail_priseContact = r"C:\Users\vt238770\Documents\_CEA\Modèles adaptés\Mails\Master IN\Master IN - Tutorat en entreprise - Prise de contact.msg"

chemin_fichier_etudiants =  fr"\\instnt\partage\FORMATIONS_I\GDRA - Master IN (parcours ADIN-GDRA-SN)\{annee_scolaire}\1-dossier etudiants\Master ADIN - {annee_scolaire.replace('-', '_')}.xlsx"

nom_onglet = "Etudiants"

colonnes_fe_etudiants = [
    "Cursus", 
    "Nom", "Prénom", "Mail apprenti", "Téléphone apprenti", 
    "Entreprise ", "Lieu entreprise", "Nom TE", "Prénom TE", "Mail TE", "Téléphone TE", "Ma fonction de suivi de l'alternant", 
    "Engagement des parties", "Entretien de prise de fonction", "1ère visite en entreprise", "2ème visite en entreprise", "Fiche évaluation 1", "Fiche évaluation 2"]

# Entretiens
prise_de_fonction = cls.PropEntretien(
    sujet = "Entretien de prise de fonction",
    chemin_modele = r"C:\Users\vt238770\Documents\_CEA\Modèles adaptés\Mails\Master IN\Master IN - Suivi d'alternance - Entretien.oft",
    duree = timedelta(hours=0, minutes=45),
    date_debut = RDV_Outlook.get_lundi_depuis_num_semaine(numero_semaine=41) # (Autour du 6 octobre : dernière semaine de la première période en entreprise → A faire avant mi-novembre)
)

premiere_visite = cls.PropEntretien(
    sujet = "1ère visite en entreprise",
    chemin_modele = r"C:\Users\vt238770\Documents\_CEA\Modèles adaptés\Mails\Master IN\Master IN - Suivi d'alternance - Entretien.oft",
    duree = timedelta(hours=1),
    date_debut = RDV_Outlook.get_lundi_depuis_num_semaine(numero_semaine=51) # (Autour du 15 décembre : dernière semaine avant vacances Noël et reprise école → A faire avant mi-janvier)
)

deuxieme_visite = cls.PropEntretien(
    sujet = "2ème visite en entreprise",
    chemin_modele = r"C:\Users\vt238770\Documents\_CEA\Modèles adaptés\Mails\Master IN\Master IN - Suivi d'alternance - Entretien.oft",
    duree = timedelta(hours=1),
    date_debut = RDV_Outlook.get_lundi_depuis_num_semaine(numero_semaine=17) # (Autour du 20 avril : dernière semaine reprise école → A faire avant fin mai)
)

# Fichiers à renvoyer
ficheEvaluation1 = cls.PropFichierARenvoyer(
    sujet = "Fiche évaluation 1",
    chemin_fichier = r"\\instnt\partage\FORMATIONS_I\GDRA - Master IN (parcours ADIN-GDRA-SN)\2024-2025\9-stages\Fiche_Evaluation_alternance_M2_Ingénierie_Nucléaire_janvier.docx",
    periode = "mi-année",
    deadline_retour = RDV_Outlook.get_lundi_depuis_num_semaine(numero_semaine=6) # (Autour du 5 février)
)

ficheEvaluation2 = cls.PropFichierARenvoyer(
    sujet = "Fiche évaluation 2",
    chemin_fichier = r"\\instnt\partage\FORMATIONS_I\GDRA - Master IN (parcours ADIN-GDRA-SN)\2024-2025\9-stages\Fiche_Evaluation_alternance_M2_Ingénierie_Nucléaire_aout.docx",
    periode = "fin d'année",
    deadline_retour = RDV_Outlook.get_lundi_depuis_num_semaine(numero_semaine=35) # (Autour du 25 août)
)


# --- UNIMES ---
# === Initialisations statiques ===
annee_scolaire = "2025-2026"
envoyer_mail = False

prefixe_sujet = "LP3D - Suivi d'alternance"
mail_responsables_univ = "isabelle.techer@unimes.fr"

chemin_modele_mail_priseContact = r"C:\Users\vt238770\Documents\_CEA\Modèles adaptés\Mails\LP3D\LP3D - Tutorat en entreprise - Prise de contact.msg"

chemin_fichier_etudiants =  fr"\\instnt\partage\FORMATIONS_I\LP3D+-démantelement désamiantage dépollution\{annee_scolaire}\1-dossier etudiants\LP3D - {annee_scolaire.replace('-', '_')}.xlsx"

nom_onglet = "Etudiants"

colonnes_fe_etudiants = [
    "Cursus", 
    "Nom", "Prénom", "Mail apprenti", "Téléphone apprenti", 
    "Entreprise ", "Lieu entreprise", "Nom TE", "Prénom TE", "Mail TE", "Téléphone TE", "Ma fonction de suivi de l'alternant", 
    "Première prise de contact", "Documents CFA à viser", "Entretien d'installation", "Entretien d’installation + fin période 1 en entreprise", "2ème entretien : fin période 2 en entreprise", "3ème entretien : milieu période 3 en entreprise"]

# Entretiens
premiere_visite = cls.PropEntretien(
    sujet = "Entretien d’installation + fin période 1 en entreprise",
    chemin_modele = r"C:\Users\vt238770\Documents\_CEA\Modèles adaptés\Mails\LP3D\LP3D - Suivi d'alternance - Entretien.oft",
    duree = timedelta(hours=0, minutes=45),
    date_debut = RDV_Outlook.get_lundi_depuis_num_semaine(numero_semaine=51) # (Autour du 15 décembre : dernière semaine avant vacances Noël et reprise école → A faire avant début janvier)
)

deuxieme_visite = cls.PropEntretien(
    sujet = "2ème entretien : fin période 2 en entreprise",
    chemin_modele = r"C:\Users\vt238770\Documents\_CEA\Modèles adaptés\Mails\LP3D\LP3D - Suivi d'alternance - Entretien.oft",
    duree = timedelta(hours=1),
    date_debut = RDV_Outlook.get_lundi_depuis_num_semaine(numero_semaine=15) # (Autour du 6 avril : dernière semaine avant reprise école → A faire avant 17 avril)
)

troisieme_visite = cls.PropEntretien(
    sujet = "3ème entretien : milieu période 3 en entreprise",
    chemin_modele = r"C:\Users\vt238770\Documents\_CEA\Modèles adaptés\Mails\LP3D\LP3D - Suivi d'alternance - Entretien.oft",
    duree = timedelta(hours=1),
    date_debut = RDV_Outlook.get_lundi_depuis_num_semaine(numero_semaine=28) # (Autour du 6 juillet : avant vacances de chacun → A faire avant fin août)
)




# === REE ===
_repertoire_documents_ree:str = r"\\harmonie\instn\uem\_Documents_communs\Formations\Formateurs\0.Docs à envoyer"  # Répertoire de la GED où sont 
_chemin_mailtype_informationsAdministratives = r"\\harmonie\instn\uem\_Documents_communs\Formations\Formateurs\Mails types\Demande des informations administratives.msg"  # Message type à envoyer aux intervenants
_repertoire_sauvegarde_fichiersREE:str = r"\\harmonie\INSTN\UEM\_Documents_communs\Formations\Formateurs\1.Intervenants - Documents administratifs" # Lieu où sauvegarder les fichiers de l'intervenant
_chemin_modele_excel_ficheIntervenant:str = r"\\harmonie\INSTN\UEM\_Documents_communs\Formations\Formateurs\P09-Pr01-Qualifier les ressources enseignantes\P09_Pr01_Ta.E_Grille des critères de qualification des compétences_V1.xlsx"  # Fichier Excel à remplir pour Laetitia Da Mota (RH INSTN qui s'occupe de rentrer les REE dans IRIS)
_adresse_mail_gestionnaire_ree_INSTN:str = "vacataires.instn@cea.fr"
_corps_html_mail_gestionnaire_ree_INSTN:str = "<p>Bonjour Laëtitia,</p><p>Je t’ai mis en PJ les documents pour intégrer/mettre à jour la fiche IRIS de ###.</p><p>Je te remercie, passe une excellente journée,</p>"

# Association colonnes excel avec command control Word
# La préparation de ce ditionnaire peut être faite avec : ree._generer_dictionnaire_depuis_excel()
_dict_colExcel_cc:Dict[str, str] = {
    "NOM": "Nom",
    "Pr\u00e9nom": "Prenoms",
    "Dipl\u00f4me ou formation/exp\u00e9rience professionnelle": "Diplome",
    "Dur\u00e9e exp\u00e9rience professionnelle": "DureeExperiencePro",
    "Niveau d'expertise permettant une reconnaissance": "NiveauExpertise",
    "Domaine / Sp\u00e9cialit\u00e9 de l'expertise": "DomaineExpertise",
    "ATTRIBUTION Niveau comp\u00e9tences techniques": None,
    "Combien de jours anim\u00e9s, en moyenne par an": "formation_nbJoursAnimes",
    "Combien de jours de formations suivies en p\u00e9dagogie (=animation)": "formation_nbJoursFormationPedagogie",
    "Profils d'apprenants form\u00e9s": "formation_profilApprenants",
    "Taux consolid\u00e9 de la satisfaction des apprenants relativement \u00e0 l'enseignant-formateur consid\u00e9r\u00e9": None,
    "Estimation par le RP de la capacit\u00e9 de l'enseignant-formateur \u00e0 animer (fond de salle)": None,
    "Outils num\u00e9riques utilis\u00e9s durant les animations r\u00e9alis\u00e9es (serious game, blended-learning\u2026)": "formation_outilsNumeriques",
    "Combien de jours pass\u00e9s en conception de s\u00e9quence de formation, en moyenne par an": "IngPedago_nbJoursConception",
    "Combien de jours de formations suivies en ing\u00e9nierie p\u00e9dagogique (=conception de s\u00e9quences de formation)": "IngPedago_nbJoursFormationIngPedago",
    "Estimation par le RP de la conception de la s\u00e9quence en fonction des objectifs p\u00e9dagogiques fournis par le RP (fond de salle, analyse des supports fournis)": None,
    "Estimation par le RP de la pertinence de l'\u00e9valuation des acquis r\u00e9alis\u00e9e par l'enseignant-formateur sur sa s\u00e9quence (analyse de la progression des apprenants : tests avant/apr\u00e8s)": None,
    "Estimation par le RP de l'utilisation des m\u00e9thodes actives (\u00e9tudes de cas, r\u00e9solution de probl\u00e8mes, classes invers\u00e9es, travaux de groupes\u2026)": None,
    "Combien d'ann\u00e9es d'exp\u00e9rience en conception de dispositifs de formations (=cr\u00e9ation et coordination)": "IngFormation_nbJoursConception",
    "Combien de jours de formations suivies en ing\u00e9nierie de formation (=conception de dispositifs de formation)": "IngFormation_nbJoursFormationIngFormation",
    "Estimation par le chef de projet ou le CUE de la complexit\u00e9 des pr\u00e9c\u00e9dents dispositifs de formation con\u00e7us": None,
    "Profil des apprenants des dispositifs de formations prc\u00e9demment con\u00e7us": "IngFormation_profilApprenants",
    "Estimation par le chef de projet ou le CUE de l'\u00e9valuation des acquis r\u00e9alis\u00e9 dans le dispositif de formation (mesure de la progression des apprenants=estimation de la qualit\u00e9 du dispositif de formation)": None,        
    "Combien d'ann\u00e9es d'exp\u00e9rience en tant que tuteur acad\u00e9mique": "IngFormation_nbAnneesTuteur",
    "Combien de r\u00e9f\u00e9rentiels d'activit\u00e9, de comp\u00e9tence et d'\u00e9valuation r\u00e9alis\u00e9s": "IngCompetences_nbReferentiels",
    "Combien de jours de formations suivies en ing\u00e9nierie de comp\u00e9tences": "IngCompetences_nbJoursFormationIngCompetences",
    "Estimation par la cellule p\u00e9dagogique de DPF de la complexit\u00e9 des pr\u00e9c\u00e9dentes r\u00e9alisations de l'ing\u00e9nieur/consultant en ing\u00e9nierie de comp\u00e9tences (complexit\u00e9 du m\u00e9tier et de son environnement : risques, r\u00e9glementation...)": None,
    "ATTRIBUTION Niveau comp\u00e9tences p\u00e9dagogiques": None,
    "Evaluation CECRL ou \u00e9quivalence TOEIC, TOEFL": "ResultatLangue2",
    "ATTRIBUTION Niveau comp\u00e9tences linguistiques": None,
    "Curriculum vitae": None
}

# --- Paramètres utilisateur
# Documents à envoyer / demander
_docsREE:dict[DocREE] ={
    "Fiche administrative" : DocREE(
        nom_fichier=r"Fiche administrative vacataire INSTN.docx", 
        frequence_maj=["Initialisation", "Mise à jour"],
        intervenants=["CEA", "vacataire", "contrat spécifique de collaboration", "auto-entrepreneur"]
    ),
    "CV" : DocREE(
        nom_fichier=r"CV-Type.docx", 
        frequence_maj=["Initialisation", "Mise à jour"],
        intervenants=["CEA", "vacataire", "contrat spécifique de collaboration", "auto-entrepreneur"]
    ),
    "RIB" : DocREE( 
        frequence_maj=["Initialisation", "Mise à jour"],
        intervenants=["vacataire", "auto-entrepreneur"]
    ),
    "Attestation employeur" : DocREE(
        nom_fichier=r"Attestation employeur.docx", 
        frequence_maj=["Initialisation", "Tous les ans"],
        intervenants=["vacataire", "contrat spécifique de collaboration"]
    ),
    "Devis" : DocREE(
        frequence_maj=["Initialisation", "Tous les ans"],
        intervenants=["auto-entrepreneur"]
    ),
    "Bilan pédagogique et financier année n-1" : DocREE(
        frequence_maj=["Initialisation", "Tous les ans"],
        intervenants=["auto-entrepreneur"]
    ),
    "Guide pour l'intervenant" : DocREE(
        nom_fichier=r"Guide pour l'intervenant.pdf",
        intervenants=["CEA", "vacataire", "contrat spécifique de collaboration", "auto-entrepreneur"]
    ),

}

_correspondance_frequence_texte:dict[str] = {
    "Initialisation" : "initialisation du dossier",
    "Mise à jour" : "s'il y a une mise à jour",
    "Tous les ans" : "<strong><u>chaque année civile</u></strong>"
}



# === BILAN SESSION V3 ===
_lieuPrincipal:str = "INSTN Marcoule"

# === IRIS ===
_dictCodesIRIS:dict[str] = {
        "Sessions" : "R04110",
        "Formations" : "R0304",
        "Ventes" : "R04301",
        "Inscriptions" : "R04500",
    }

_sessions:PropExportIRIS = PropExportIRIS(
    nom_typeExport = "Sessions",
    codeExport = "R04110",

    repertoire_input = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts originaux",
    nom_onglet_input = "Data",
    nbLignes_avantET_input = 1,

    repertoire_modele = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Modèles",
    nom_fichier_modele = "R04110_Sessions-Modèle.xlsx",

    repertoire_output = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts complets",
    nom_fichier_output = "R04110_Sessions-COMPLET.xlsx"
    )

_formations:PropExportIRIS = PropExportIRIS(
    nom_typeExport = "Formations",
    codeExport = "R0304",

    repertoire_input = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts originaux",
    nom_onglet_input = "Data",
    nbLignes_avantET_input = 0,

    repertoire_modele = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Modèles",
    nom_fichier_modele = "R0304_Formations-Modèle.xlsx",

    repertoire_output = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts complets",
    nom_fichier_output = "R0304_Formations-COMPLET.xlsx"
    )

_ventes:PropExportIRIS = PropExportIRIS(
    nom_typeExport = "Ventes",
    codeExport = "R04301",

    repertoire_input = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts originaux", #Car il y a des petits bugs sur certains CSV
    nom_onglet_input = "Data",
    nbLignes_avantET_input = 1,

    repertoire_modele = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Modèles",
    nom_fichier_modele = "R04301_Ventes-Modèle.xlsx",

    repertoire_output = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts complets",
    nom_fichier_output = "R04301_Ventes-COMPLET.xlsx"
    )

_inscriptions:PropExportIRIS = PropExportIRIS(
    nom_typeExport = "Inscriptions",
    codeExport = "R04500",

    repertoire_input = r"\\instnt\HOME\REFERENC\IRIS - rapports de synthese",
    nom_onglet_input = "Data",
    nbLignes_avantET_input = 1,

    repertoire_modele = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Modèles",
    nom_fichier_modele = "R04500_Inscriptions-Modèle.xlsx",

    repertoire_output = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts complets",
    nom_fichier_output = "R04500_Inscriptions-COMPLET.xlsx"
    )

_dict_DE_IRIS = {}
_dict_DE_IRIS["CodesExports"] = {
        "Sessions" : "R04110",
        "Formations" : "R0304",
        "Ventes" : "R04301",
        "Inscriptions" : "R04500",
    }
_dict_DE_IRIS["PropExportIRIS"] = {
        "Sessions" : _sessions, 
        "Formations" : _formations,
        "Ventes" : _ventes,
        "Inscriptions" : _inscriptions}

_tSessions = (
    'R04110_Sessions-2011 à 2014 FINAL.xlsx',
    'R04110_Sessions-2015 FINAL.xlsx',
    'R04110_Sessions-2016 FINAL.xlsx',
    'R04110_Sessions-2017 FINAL.xlsx',
    'R04110_Sessions-2018 FINAL.xlsx',
    'R04110_Sessions-2019 FINAL.xlsx',
    'R04110_Sessions-2020 FINAL.xlsx',
    'R04110_Sessions-2021 FINAL.xlsx',
    'R04110_Sessions-2022 FINAL.xlsx',
    'R04110_Sessions-2023 FINAL.xlsx',
    'R04110_Sessions-2024 FINAL.xlsx',
    'R04110_Sessions-2025 au 2025.10.21.xlsx')

_tFormations = (
    "R0304_Ref_Formation-Listedesformations-2025.10.22.xlsx", )

_tVentes = (
    'R04301_Sessions-Ventes-FC2020 FINAL.xlsx',
    'R04301_Sessions-Ventes-FC2021 FINAL.xlsx',
    'R04301_Sessions-Ventes-FC2022 FINAL.xlsx',
    'R04301_Sessions-Ventes-FC2023 FINAL.xlsx',
    'R04301_Sessions-Ventes-FC2024 FINAL.xlsx',
    'R04301_Sessions-Ventes-filtre sur FC-2025 au 2025-02-05 LG.xlsx',
    'R04301_Sessions-Ventes-filtre sur FC-2025 au 2025-03-03 LG.xlsx',
    'R04301_Sessions-Ventes-filtre sur FC-2025 au 2025-04-01 LG.xlsx',
    'R04301_Sessions-Ventes-filtre sur FC-2025 au 2025-05-12 LG.xlsx',
    'R04301_Sessions-Ventes-filtre sur FC-2025 au 2025-06-02 LG.xlsx',
    'R04301_Sessions-Ventes-filtre sur FC-2025 au 2025-07-01 LG.xlsx',
    'R04301_Sessions-Ventes-filtre sur FC-2025 au 2025-08-01 LG.xlsx',
    'R04301_Sessions-Ventes-filtre sur FC-2025 au 2025-09-01 LG.xlsx',
    'R04301_Sessions-Ventes-filtre sur FC-2025 au 2025-10-01 LG.xlsx')
    
_tInscriptions = (
    'R04500_Sessions-Inscriptions-FC2020 FINAL.xlsx',
    'R04500_Sessions-Inscriptions-FC2021 FINAL.xlsx',
    'R04500_Sessions-Inscriptions-FC2022 FINAL.xlsx',
    'R04500_Sessions-Inscriptions-FC2023 FINAL.xlsx',
    'R04500_Sessions-Inscriptions-FC2024 FINAL.xlsx',
    'R04500_Sessions-Inscriptions-filtre sur FC-2025 au 2025-02-05.xlsx',
    'R04500_Sessions-Inscriptions-filtre sur FC-2025 au 2025-03-03.xlsx',
    'R04500_Sessions-Inscriptions-filtre sur FC-2025 au 2025-04-01.xlsx',
    'R04500_Sessions-Inscriptions-filtre sur FC-2025 au 2025-05-12.xlsx',
    'R04500_Sessions-Inscriptions-filtre sur FC-2025 au 2025-06-02.xlsx',
    'R04500_Sessions-Inscriptions-filtre sur FC-2025 au 2025-07-01.xlsx',
    'R04500_Sessions-Inscriptions-filtre sur FC-2025 au 2025-08-01.xlsx',
    'R04500_Sessions-Inscriptions-filtre sur FC-2025 au 2025-09-01.xlsx',
    'R04500_Sessions-Inscriptions-filtre sur FC-2025 au 2025-10-01.xlsx')

_dict_DE_IRIS["Fichiers"] = {
    "Sessions" : _tSessions,
    "Formations" : _tFormations,
    "Ventes" : _tVentes,
    "Inscriptions" : _tInscriptions
}


# === EVALSTAT ===

# === Colonnes du CSV selon traitement à avoir ===
# Colonnes descriptives à recopier
_colonnes_csv_fixes = [
    "Chemin fichier CSV", "Prénom", "Nom", "Entreprise", "Code session"]

# Colonnes avec note/commentaire en binôme
_colonnes_csv_avec_commentaires = [
    "Accueil, organisation et qualité des informations délivrées",
    "Conseils et orientation avant l'inscription",
    "Informations après l'inscription",
    "Accueil à l'arrivée sur site",
    "Prise en compte de vos besoins et attentes",
    "Qualité des animations",
    "Logique d'enchainement des interventions",
    "Qualité des supports de cours utilisés",
    "Qualité des moyens pédagogique",
    "Accès aux outils digitaux",
    "Satisfaction globale",
    "Avez-vous d'autres besoins de formation ?"]

# Colonnes à valeur texte seule
_colonnes_csv_commentaires_seuls = [
    "Comment avez-vous connu cette formation ?",
    "Commentaires, remarques, suggestions"]

# Colonnes note seule (il se trouve que je vais aussi devoir convertir le booléen)
_colonnes_csv_bool = [
    "Recommanderiez-vous cette formation ?"]

# === Colonnes de l'extract IRIS Sessions à récupérer ===
_colonnes_sessions = [
    "N° Session",
    "Formation",
    "Trigramme formation",
    "Code IRIS",
    "Date début ses.",
    "Année début ses.",
    "Type de formation",
    "Trigramme RP",
    "Trigramme AF",
    "Nb. Présents"]


