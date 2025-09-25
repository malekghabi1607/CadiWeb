from vte.instn import *

# ==== Initialisation variables utilisateur ====
# Initialisation des chemins des répertoires
rep_gedMiroir = r"\\instnt\partage\FORMATIONS_C"
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







# === Test PropExportIRIS ===
#print(inscriptions)


### --------------------------------------------------------------------
#  Extracts IRIS
### --------------------------------------------------------------------



#=== Tests TravauxFichiersIRIS ===
#env = TravauxFichiersIRIS.avecLecture(sessions, (r'C:/Users/vt238770/Documents/_CEA/_Formations/Extracts IRIS - Faits/Extracts originaux/R04110_Sessions-2021 FINAL.xlsx', r'C:/Users/vt238770/Documents/_CEA/_Formations/Extracts IRIS - Faits/Extracts originaux/R04110_Sessions-2022 FINAL.xlsx'))
#env = TravauxFichiersIRIS.avecEcritureOutputDefaut(sessions, (r'.\R04110_Sessions-2021 FINAL.xlsx', r'.\R04110_Sessions-2022 FINAL.xlsx'))
#env = TravauxFichiersIRIS.avecEcritureOutputDefaut(sessions, (r'C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts originaux\R04110_Sessions-2021 FINAL.xlsx', r'C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts originaux\R04110_Sessions-2022 FINAL.xlsx'))
#env = TravauxFichiersIRIS.avecEcritureOutputDefaut(sessions, r'C:/Users/vt238770/Documents/_CEA/_Formations/Extracts IRIS - Faits/Extracts originaux/R04110_Sessions-2021 FINAL.xlsx')
#env = TravauxFichiersIRIS.avecEcritureOutputDefaut(formations, )
#print(env)



# ==== Initialisation exports IRIS ====
dictCodesIRIS = {
        "Sessions" : "R04110",
        "Formations" : "R0304",
        "Ventes" : "R04301",
        "Inscriptions" : "R04500",
    }

sessions = PropExportIRIS(
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

formations = PropExportIRIS(
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

ventes = PropExportIRIS(
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

inscriptions = PropExportIRIS(
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

dict_DE_IRIS = {}
dict_DE_IRIS["CodesExports"] = {
        "Sessions" : "R04110",
        "Formations" : "R0304",
        "Ventes" : "R04301",
        "Inscriptions" : "R04500",
    }
dict_DE_IRIS["PropExportIRIS"] = {
        "Sessions" : sessions, 
        "Formations" : formations,
        "Ventes" : ventes,
        "Inscriptions" : inscriptions}




# === Lancer TravauxFichiersIRIS pour un seul type d'export ===
#env = TravauxFichiersIRIS.avecEcritureOutputDefaut(sessions, (r'C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts originaux\R04110_Sessions-2021 FINAL.xlsx', r'C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts originaux\R04110_Sessions-2022 FINAL.xlsx'))
#env = TravauxFichiersIRIS.avecEcritureOutputDefaut(formations, )
#print(env)

# === Lancer TravauxFichiersIRIS pour plusieurs types d'export ===
#mettreAJourTousLesExportsIRIS_auto(("Sessions", "Formations", "Ventes", "Inscriptions"))
#mettreAJourTousLesExportsIRIS_auto(("Sessions", "Formations"))
#mettreAJourTousLesExportsIRIS_fileDialog(("Formations", ))










### --------------------------------------------------------------------
#  EvalStat
### --------------------------------------------------------------------

# === Fichier CSV individuel
#es = Traiter_evalStat()
#es = Traiter_evalStat.depuis_chemin_csv_stagiaires(r"\\instnt\partage\FORMATIONS_C\22B\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\S-17606 - 22B - 06-2025\S-17606-FC25-22B-VTE-CAR-Stagiaires.csv")
#es = Traiter_evalStat.depuis_chemin_csv_stagiaires(r"P:\FORMATIONS_C\778\P07-bilan-sessions-et-bilan-formation\2024-Bilans 778\Evaluations 778(2024.11)\S-15715-FC24-778-VMO-VCA-Stagiaires.csv")
#es = Traiter_evalStat.depuis_chemin_csv_stagiaires(r"\\INSTNT\partage\FORMATIONS_C\948\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2024-01-S-15697 UEM\S-15697-FC24-948-VTE-VCA-Stagiaires.csv")

tuple_csv_stagiaires = (
    r"\\INSTNT\partage\FORMATIONS_C\948\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2021-11-S-11369 UEM\EVALSTAT\S-11369-FC21-948-VLE-SNA-Stagiaires.csv",
    r"\\INSTNT\partage\FORMATIONS_C\948\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2022-11-S-12995 UEM\S-12995-FC22-948-VTE-SNA-Stagiaires.csv",
    r"\\INSTNT\partage\FORMATIONS_C\948\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2024-01-S-15697 UEM\S-15697-FC24-948-VTE-VCA-Stagiaires.csv"
    )

#es = Traiter_evalStat.depuis_tuple_csv_stagiaires(tuple_csv_stagiaires=tuple_csv_stagiaires)




### --------------------------------------------------------------------
#  Fiche de coûts
### --------------------------------------------------------------------

#lire_fdc(r"\\instnt\PARTAGE\FORMATIONS_C\948\P05-P06-dossier-conception-referentiel\fiche-de-cout-et-code-de-formation\Fiche de coûts - 948 - Elaboration de scénarios de DEM - 2025.01.24.xlsx")




### --------------------------------------------------------------------
#  Bilans pédagogiques
### --------------------------------------------------------------------

# === Lancer génération du bilan (formation je pense)
#fenetreBilan()

# === bilans de session
#bs = BilanSessionV3("948", 2024, "Année")






### --------------------------------------------------------------------
#   Code pour Fichier EE
### --------------------------------------------------------------------
#fw = FichierWord.depuisFichier(r"C:\Users\vt238770\Documents\_CEA\EXCEL - Tests ou Backup\Lien formulaire Word vers Excel\Fiche admin - ContentControl.docx")
#print(fw)
#mail= Mail()

#mail._creer_mail(
#    destinataires="vacataires.instn@cea.fr",
#    sujet="Documents pour mise à jour IRIS", # Pimper avec le nom de l'intervenant
#    corps_html="<p>Bonjour Laëtitia,</p><p>Je t’ai mis en PJ les documents pour intégrer/mettre à jour la fiche IRIS de ###.</p><p>Je te remercie, passe une excellente journée,</p>", # Pimper avec le nom de l'intervenant
#    pieces_jointes=r"C:\Users\vt238770\Documents\_CEA\EXCEL - Tests ou Backup\Lien formulaire Word vers Excel\Fiche admin - ContentControl.docx", # if None, sélectionner avec fileDialog
#    envoyer_mail=False  # envoie directement sans afficher
#)

#t_ree = Traiter_REE()
#Il me faudrait une classe Traiter_REE :
#   - variables de la classe : _word_ficheAdministrative:FichierWord= None, _excel_ficheIntervenant:FichierExcel = None, _mail_traitement _repertoire_sauvegarde + _mail_gestionnaire_REE +











### --------------------------------------------------------------------
#   Code pour Mails
### --------------------------------------------------------------------
#rdv = RDV_Outlook.depuis_modele_msg(
#    chemin_modele=r"C:\Users\vt238770\Documents\_CEA\Modèles adaptés\Mails\Master IN\Master IN - Suivi d'alternance - Entretien de prise de fonction.msg",
#    participants_obligatoires="vincent.testard@cea.fr",
#    participants_facultatifs="assistant@example.com"
#)






### --------------------------------------------------------------------
#   Code pour traiter le contact des apprentis
### --------------------------------------------------------------------
#ca = Traiter_contactsApprentis.UGA()

#rdv = RDV_Outlook()
r"C:\Users\vt238770\Documents\_CEA\Modèles adaptés\Mails\TEST.msg"




