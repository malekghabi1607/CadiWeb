from vte.instn import *
#from vte.ihm_console import IHM_console
#from vte.ihm_tkinter import IHMTkinter

    
# === Ecrire nouveaux extracts IRIS complets (qui concatène plusieurs extracts individuels) ===
# Procédure :
# IRIS : 
#    - faire export sessions pour depuis le début de l'année en cours jusqu'à 2030 (05_Planification ; puis R04110)
#    - faire export formations (01_Référentiel puis R0304_listeFormations)
# COPIER-COLLER ventes depuis référence (GED) sans écraser car petits bugs sur en-têtes de certains fichiers
# IRIS.mettreAJourTousLesExportsIRIS_fileDialog(("Sessions", "Formations", "Ventes", "Inscriptions"))

def main():






    print("")


### --------------------------------------------------------------------
#  1 seul type d'export
### --------------------------------------------------------------------
# Test ConfigExportIRIS
def test01():
    print(IRIS._SESSIONS)  # OK

# Concatener_exportsIRIS_typeUnique — chemin = str
def test02():
    IRIS.concatener_exportsIRIS_typeUnique(typeExport="Sessions", chemins_fichiersInput=r"R:\_Echanges\VTE\Prog\IRIS\Extracts originaux\R04110_Sessions-2025 au 2025.11.29.xlsx")

# Concatener_exportsIRIS_typeUnique — Sans fichier input # TODO : faire avec 1 et plusieurs fichiers
def test03():
    IRIS.concatener_exportsIRIS_typeUnique(typeExport="Sessions")
  
# Concaténer avec changement colonnes dans le modèle 
def test04():
    IRIS.concatener_exportsIRIS_typeUnique(typeExport="Inscriptions", chemins_fichiersInput=r"H:\REFERENC\IRIS - rapports de synthese\R04500_Sessions-Inscriptions-filtre sur FC-2025 au 2025-11-03.xlsx")




### --------------------------------------------------------------------
#  Plusieurs types d'export
### --------------------------------------------------------------------

# Cas 1 seul demandé + filedialog
def test05():
    IRIS.concatener_exportsIRIS_plusieursTypes(typesExports="Sessions")

# Cas 2 seul demandé + auto
def test06():
    IRIS.concatener_exportsIRIS_plusieursTypes(typesExports="Sessions", depuis_config=True)

# Cas 3 demandés + filedialog
def test07():
    IRIS.concatener_exportsIRIS_plusieursTypes(typesExports=("Sessions", "Formations"))

    # === Lancer TravauxFichiersIRIS pour plusieurs types d'export ===
    #mettreAJourTousLesExportsIRIS_auto(("Sessions", "Formations", "Ventes", "Inscriptions"))
    #mettreAJourTousLesExportsIRIS_auto(("Sessions", "Formations"))
    #IRIS.mettreAJourTousLesExportsIRIS_fileDialog(("Sessions", "Inscriptions"))  # OK




    

if __name__ == "__main__":
    test07()
    #main()










# === Test PropExportIRIS ===
#print(inscriptions)


### --------------------------------------------------------------------
#  Extracts IRIS
### --------------------------------------------------------------------




# === Lancer TravauxFichiersIRIS pour un seul type d'export ===
#env = TravauxFichiersIRIS.avecEcritureOutputDefaut(sessions, (r'C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts originaux\R04110_Sessions-2021 FINAL.xlsx', r'C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts originaux\R04110_Sessions-2022 FINAL.xlsx'))
#env = TravauxFichiersIRIS.avecEcritureOutputDefaut(formations, )
#print(env)

# === Lancer TravauxFichiersIRIS pour plusieurs types d'export ===
#mettreAJourTousLesExportsIRIS_auto(("Sessions", "Formations", "Ventes", "Inscriptions"))
#mettreAJourTousLesExportsIRIS_auto(("Sessions", "Formations"))
#mettreAJourTousLesExportsIRIS_fileDialog(("Formations", ))
