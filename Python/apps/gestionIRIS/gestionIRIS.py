from vte.instn import *
from vte.ihm_console import IHM_console

from colorama import init
init(autoreset=True)

    
# === Ecrire nouveaux extracts IRIS complets (qui concatène plusieurs extracts individuels) ===
# Procédure :
# IRIS : 
#    - faire export sessions pour depuis le début de l'année en cours jusqu'à 2030 (05_Planification ; puis R04110)
#    - faire export formations (01_Référentiel puis R0304_listeFormations)
# COPIER-COLLER ventes depuis référence (GED) sans écraser car petits bugs sur en-têtes de certains fichiers
# IRIS.mettreAJourTousLesExportsIRIS_fileDialog(("Sessions", "Formations", "Ventes", "Inscriptions"))

MODE = "console"

depuis_config:bool = False
MENUS = {
    "Tout traiter": {
        "action": IRIS.concatener_exportsIRIS_plusieursTypes,
        "kwargs": {"typesExports":("Sessions", "Formations", "Ventes", "Inscriptions")},
        "demander": ["depuis_config"],
        "indications":"Entrez :\n   - 0 pour sélection manuelle des fichiers (filedialog)\n   - 1 pour sélection automatique (emploi de config_extractsIRIS)"
    },
    
    "Traiter sessions (R04110)": {
        "action": IRIS.concatener_exportsIRIS_plusieursTypes,
        "kwargs": {"typesExports":"Sessions"},
        "demander": ["depuis_config"],
        "indications":"Entrez :\n   - 0 pour sélection manuelle des fichiers (filedialog)\n   - 1 pour sélection automatique (emploi de config_extractsIRIS)"
    },
    "Traiter Formations (R0304)": {
        "action": IRIS.concatener_exportsIRIS_plusieursTypes,
        "kwargs": {"typesExports":"Formations"},
        "demander": ["depuis_config"],
        "indications":"Entrez :\n   - 0 pour sélection manuelle des fichiers (filedialog)\n   - 1 pour sélection automatique (emploi de config_extractsIRIS)"
    },
    "Traiter Ventes (R04301)": {
        "action": IRIS.concatener_exportsIRIS_plusieursTypes,
        "kwargs": {"typesExports":"Ventes"},
        "demander": ["depuis_config"],
        "indications":"Entrez :\n   - 0 pour sélection manuelle des fichiers (filedialog)\n   - 1 pour sélection automatique (emploi de config_extractsIRIS)"
    },
    "Traiter Inscriptions (R04500)": {
        "action": IRIS.concatener_exportsIRIS_plusieursTypes,
        "kwargs": {"typesExports":"Inscriptions"},
        "demander": ["depuis_config"],
        "indications":"Entrez :\n   - 0 pour sélection manuelle des fichiers (filedialog)\n   - 1 pour sélection automatique (emploi de config_extractsIRIS)"
    }

    
}

def main():
    vlog.print("Info", "*************\nBienvenue dans le script pour concaténer des fichiers IRIS.\n*************", style=["vert clair"])
    """
    ihm = IHM_console({})  # Pas besoin de menu ici

    # On demande si on veut être en auto ou en filedialog
    depuis_config = ihm.demander_saisie(
    texte="Choisissez le mode de sélection :\n1) Sélection manuelle (filedialog)\n2) Sélection automatique (emploi de config_extractsIRIS)\n",
    type_attendu=int
    )

    # Convertir la réponse en booléen
    if depuis_config == 1:
        depuis_config = False
    elif depuis_config == 2:
        depuis_config = True


    """
    ihm = IHM_console(MENUS)
    ihm.afficher_menu()






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
    #test07()
    main()










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
