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
prefixe_sujet = "LP3D - Suivi d'alternance"
mail_responsables_univ = "isabelle.techer@unimes.fr"
chemin_modele_mail_priseContact = r"\\harmonie\instn\uem\_Echanges\VTE\Prog\Modèles\Tutorat\LP3D\LP3D - Tutorat en entreprise - Prise de contact.msg"
envoyer_mail = False


# Fichier Excel avec les informations des étudiants
chemin_fichier_etudiants =  fr"\\instnt\partage\FORMATIONS_I\LP3D+-démantelement désamiantage dépollution\{annee_scolaire}\1-dossier etudiants\LP3D - {annee_scolaire.replace('-', '_')}.xlsx"
nom_onglet = "Etudiants"

# Colonnes à récupérer du fichier Excel étudiant, hors relances et entretiens
colonnes_fe_etudiants = [
   "Cursus", 
   "Nom", "Prénom", "Mail apprenti", "Téléphone apprenti", 
   "Entreprise ", "Lieu entreprise", "Nom TE", "Prénom TE", "Mail TE", "Téléphone TE", "Ma fonction de suivi de l'alternant", 
   "Première prise de contact"]
   #"Première prise de contact", "Documents CFA à viser", "Entretien d'installation", "Entretien d’installation + fin période 1 en entreprise (autour de  S51)", "2ème entretien : fin période 2 en entreprise (autour de  S15)", "3ème entretien : milieu période 3 en entreprise"]



# Relances mail pour signature docs → Doit avoir la même structure (ordre + nom) que les colonnes Excel qui trace les retours tuteurs et apprentis : relances puis entretiens
relances = [
   "Documents CFA à viser",
   "Entretien d'installation"
]
# RDV entretiens  → Doit avoir les mêmes noms que les colonnes Excel qui trace les retours tuteurs et apprentis et colonnes_fe_etudiants
entretiens = [
   dict(
      sujet = "1er entretien : installation + fin période 1 en entreprise",
      chemin_modele = r"\\harmonie\instn\uem\_Echanges\VTE\Prog\Modèles\Tutorat\LP3D\LP3D - Suivi d'alternance - Entretien.oft",
      duree = timedelta(hours=0, minutes=45),
      date_debut = RDV_Outlook.get_lundi_depuis_num_semaine(numero_semaine=51) # (Autour du 15 décembre : dernière semaine avant vacances Noël et reprise école → A faire avant début janvier)
   ),

   dict(
      sujet = "2ème entretien : fin période 2 en entreprise",
      chemin_modele = r"\\harmonie\instn\uem\_Echanges\VTE\Prog\Modèles\Tutorat\LP3D\LP3D - Suivi d'alternance - Entretien.oft",
      duree = timedelta(hours=1),
      date_debut = RDV_Outlook.get_lundi_depuis_num_semaine(numero_semaine=15) # (Autour du 6 avril : dernière semaine avant reprise école → A faire avant 17 avril)
   ),

   dict(
      sujet = "3ème entretien : milieu période 3 en entreprise",
      chemin_modele = r"\\harmonie\instn\uem\_Echanges\VTE\Prog\Modèles\Tutorat\LP3D\LP3D - Suivi d'alternance - Entretien.oft",
      duree = timedelta(hours=1),
      date_debut = RDV_Outlook.get_lundi_depuis_num_semaine(numero_semaine=28) # (Autour du 6 juillet : avant vacances de chacun → A faire avant fin août)
   ),
]

# Mail + RDV de rappel fichiers à renvoyer (fiches d'évaluation)
fichiersARenvoyer = []
