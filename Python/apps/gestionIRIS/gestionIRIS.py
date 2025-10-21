from vte.instn import *
#from vte.ihm_console import IHM_console
#from vte.ihm_tkinter import IHMTkinter

from mailmerge import MailMerge

def main():
    # === Importer données fichier(s) IRIS dans un seul Dataframe
    #env = IRIS.avecLecture(sessions, (r'C:/Users/vt238770/Documents/_CEA/_Formations/Extracts IRIS - Faits/Extracts originaux/R04110_Sessions-2021 FINAL.xlsx', r'C:/Users/vt238770/Documents/_CEA/_Formations/Extracts IRIS - Faits/Extracts originaux/R04110_Sessions-2022 FINAL.xlsx'))  # OK
    #print(env._df_tableau)
    

    
    # === Ecrire nouveaux extracts IRIS complets (qui concatène plusieurs extracts individuels) ===
    # Procédure :
    # IRIS : 
    #    - faire export sessions pour toute l'année en cours 
    #    - faire export formations
    # COPIER-COLLER ventes depuis référence (GED) sans écraser car petits bugs sur en-têtes de certains fichiers
    # IRIS.mettreAJourTousLesExportsIRIS_fileDialog(("Sessions", "Formations", "Ventes", "Inscriptions"))


    # TODO Mettre à jour extracts IRIS complets
    # Lire les extracts
    # Faire un ls du répertoire cible des extracts
    # Ouvrir extract vomplet et regarder les fichiers inclus
    # Comparer fichiers inclus et ls
    # Supprimer des lign,es (si besoin)
    # Ajouter fichiers nécessaires 

    print("")




def test():
    # === Test PropExportIRIS ===
    #print(ventes)  # OK
    
    #=== Tests TravauxFichiersIRIS avecLecture ===
    #env = IRIS.avecLecture(sessions, r'C:/Users/vt238770/Documents/_CEA/_Formations/Extracts IRIS - Faits/Extracts originaux/R04110_Sessions-2021 FINAL.xlsx')  # OK
    #env = IRIS.avecLecture(sessions, (r'C:/Users/vt238770/Documents/_CEA/_Formations/Extracts IRIS - Faits/Extracts originaux/R04110_Sessions-2021 FINAL.xlsx', r'C:/Users/vt238770/Documents/_CEA/_Formations/Extracts IRIS - Faits/Extracts originaux/R04110_Sessions-2022 FINAL.xlsx'))  # OK
    #print(env._df_tableau)



    # === Lancer TravauxFichiersIRIS pour un seul type d'export ===
    #env = TravauxFichiersIRIS.avecEcritureOutputDefaut(sessions, (r'.\R04110_Sessions-2021 FINAL.xlsx', r'.\R04110_Sessions-2022 FINAL.xlsx'))
    #env = TravauxFichiersIRIS.avecEcritureOutputDefaut(sessions, (r'C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts originaux\R04110_Sessions-2021 FINAL.xlsx', r'C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts originaux\R04110_Sessions-2022 FINAL.xlsx'))
    #env = TravauxFichiersIRIS.avecEcritureOutputDefaut(sessions, r'C:/Users/vt238770/Documents/_CEA/_Formations/Extracts IRIS - Faits/Extracts originaux/R04110_Sessions-2021 FINAL.xlsx')
    #env = IRIS.avecEcritureOutputDefaut(formations, )  # OK
    #print(env)



    # === Lancer TravauxFichiersIRIS pour plusieurs types d'export ===
    #mettreAJourTousLesExportsIRIS_auto(("Sessions", "Formations", "Ventes", "Inscriptions"))
    #mettreAJourTousLesExportsIRIS_auto(("Sessions", "Formations"))
    #IRIS.mettreAJourTousLesExportsIRIS_fileDialog(("Sessions", "Inscriptions"))  # OK

    # Procédure :
    # IRIS : 
    #    - faire export sessions pour toute l'année en cours 
    #    - faire export formations
    # COPIER-COLLER ventes depuis référence (GED) sans écraser car petits bugs sur en-têtes de certains fichiers
    # IRIS.mettreAJourTousLesExportsIRIS_fileDialog(("Sessions", "Inscriptions"))
    print()



    

if __name__ == "__main__":
    test()
    #main()










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








# === Lancer TravauxFichiersIRIS pour un seul type d'export ===
#env = TravauxFichiersIRIS.avecEcritureOutputDefaut(sessions, (r'C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts originaux\R04110_Sessions-2021 FINAL.xlsx', r'C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts originaux\R04110_Sessions-2022 FINAL.xlsx'))
#env = TravauxFichiersIRIS.avecEcritureOutputDefaut(formations, )
#print(env)

# === Lancer TravauxFichiersIRIS pour plusieurs types d'export ===
#mettreAJourTousLesExportsIRIS_auto(("Sessions", "Formations", "Ventes", "Inscriptions"))
#mettreAJourTousLesExportsIRIS_auto(("Sessions", "Formations"))
#mettreAJourTousLesExportsIRIS_fileDialog(("Formations", ))
