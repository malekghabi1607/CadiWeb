from vte.instn import *
"""
Mode op

Vous pouvez soit traiter des bilans de session :
   - soit par session [liste_codes_IRIS] en indiquant des codes IRIS à 5 chiffres
   - soit par période (1er semestre, 2nd semestre ou annuel) en indiquant :
        ¤ un trigramme formation en majuscule (ex. : TEL)
        ¤ une année à 4 chiffres (ex. : 2025)
        ¤ une periode = [1er semestre, 2nd semestre, Année] (ex. : 2nd semestre)

Pour chacun de ces cas, vous pouvez :
   - soit les faire un par un ;

"""


# Année en cours
annee_scolaire = "2025-2026"

# Infos mails
prefixe_sujet = "Master IN - Suivi d'alternance"
mail_responsables_univ = "master-in-responsables@univ-grenoble-alpes.fr"
chemin_modele_mail_priseContact = r"\\harmonie\instn\uem\_Echanges\VTE\Prog\Modèles\Tutorat\Master IN\Master IN - Tutorat en entreprise - Prise de contact.msg"
envoyer_mail = False


# Fichier Excel avec les informations des étudiants
chemin_fichier_etudiants =  fr"\\instnt\partage\FORMATIONS_I\GDRA - Master IN (parcours ADIN-GDRA-SN)\{annee_scolaire}\1-dossier etudiants\Master ADIN - {annee_scolaire.replace('-', '_')}.xlsx"
nom_onglet = "Etudiants"

# Colonnes à récupérer du fichier Excel étudiant, hors relances et entretiens
colonnes_fe_etudiants = [
   "Cursus", 
   "Nom", "Prénom", "Mail apprenti", "Téléphone apprenti", 
   "Entreprise ", "Lieu entreprise", "Nom TE", "Prénom TE", "Mail TE", "Téléphone TE", "Ma fonction de suivi de l'alternant", 
   "Engagement des parties"]
   #"Première prise de contact", "Documents CFA à viser", "Entretien d'installation", "Entretien d’installation + fin période 1 en entreprise (autour de  S51)", "2ème entretien : fin période 2 en entreprise (autour de  S15)", "3ème entretien : milieu période 3 en entreprise"]


# Relances mails pour signature docs → Doit avoir la même structure (ordre + nom) que les colonnes Excel qui trace les retours tuteurs et apprentis : relances puis entretiens
relances = [
   "Engagement des parties",
]

# RDV entretiens  → Doit avoir les mêmes noms que les colonnes Excel qui trace les retours tuteurs et apprentis et colonnes_fe_etudiants
entretiens = [
   dict(
      sujet = "Suivi n°1 : prise de fonction",
      chemin_modele = r"\\harmonie\instn\uem\_Echanges\VTE\Prog\Modèles\Tutorat\Master IN\Master IN - Suivi d'alternance - Entretien.oft",
      duree = timedelta(hours=0, minutes=45),
      date_debut = RDV_Outlook.get_lundi_depuis_num_semaine(numero_semaine=41) # (Autour du 6 octobre : dernière semaine de la première période en entreprise → A faire avant mi-novembre)
   ),

   dict(
      sujet = "Suivi n°2",
      chemin_modele = r"\\harmonie\instn\uem\_Echanges\VTE\Prog\Modèles\Tutorat\Master IN\Master IN - Suivi d'alternance - Entretien.oft",
      duree = timedelta(hours=1),
      date_debut = RDV_Outlook.get_lundi_depuis_num_semaine(numero_semaine=51) # (Autour du 15 décembre : dernière semaine avant vacances Noël et reprise école → A faire avant mi-janvier)
   ),

   dict(
      sujet = "Suivi n°3",
      chemin_modele = r"\\harmonie\instn\uem\_Echanges\VTE\Prog\Modèles\Tutorat\Master IN\Master IN - Suivi d'alternance - Entretien.oft",
      duree = timedelta(hours=1),
      date_debut = RDV_Outlook.get_lundi_depuis_num_semaine(numero_semaine=17) # (Autour du 20 avril : dernière semaine reprise école → A faire avant fin mai)
   ),
]


# Mail + RDV de rappel fichiers à renvoyer (fiches d'évaluation)
fichiersARenvoyer = [
   dict(
      sujet = "Fiche évaluation 1",
      chemin_fichier = r"\\instnt\partage\FORMATIONS_I\GDRA - Master IN (parcours ADIN-GDRA-SN)\2025-2026\9-stages\Fiche_Evaluation_alternance_M2_Ingénierie_Nucléaire_janvier.docx",
      periode = "mi-année",
      deadline_retour = datetime(2026, 1, 27) #RDV_Outlook.get_lundi_depuis_num_semaine(numero_semaine=6) # (Autour du 5 février)
   ),

   dict(
      sujet = "Fiche évaluation 2",
      chemin_fichier = r"\\instnt\partage\FORMATIONS_I\GDRA - Master IN (parcours ADIN-GDRA-SN)\2025-2026\9-stages\Fiche_Evaluation_alternance_M2_Ingénierie_Nucléaire_aout.docx",
      periode = "fin d'année",
      deadline_retour = RDV_Outlook.get_lundi_depuis_num_semaine(numero_semaine=35) # (Autour du 25 août)
   ),
]