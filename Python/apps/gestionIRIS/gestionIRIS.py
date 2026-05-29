from vte.services.iris_services import IRIS_services
from vte.ihm_console import IHM_console
from vte.utils.utils import *

from colorama import init

from vte.services.iris_dumps_services import mettre_a_jour_dumps_depuis_ged
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


def mettre_a_jour_dumps_ged_console() -> None:
    """
    Lance la mise a jour des dumps GED et affiche un bilan lisible en console.
    """
    bilan = mettre_a_jour_dumps_depuis_ged()

    print("\n=== Bilan mise a jour dumps GED ===")
    codes_copie_ged = {"R04301", "R04500"}
    for code_export, infos in bilan.items():
        fichier_copie = infos["copie"]
        fichiers_archives = infos["archives"]

        print(f"\n{code_export}")

        if code_export not in codes_copie_ged:
            print("  Copie    : non concernee")
        elif fichier_copie is None:
            print("  Copie    : aucun fichier recent trouve")
        else:
            print(f"  Copie    : {fichier_copie.name}")

        if not fichiers_archives:
            print("  BAK      : aucun ancien dump archive")
        else:
            print("  BAK      :")
            for fichier_archive in fichiers_archives:
                print(f"    - {fichier_archive.name}")


MENUS = {
    "Tout traiter": {
        "action": IRIS_services.concatener_plusieursTypes,
        "kwargs": {"typesExports":("Sessions", "Formations", "Ventes", "Inscriptions")},
        "demander": ["depuis_config"],
        "indications":"Entrez :\n   - 0 pour sélection manuelle des fichiers (filedialog)\n   - 1 pour sélection automatique (emploi de config_extractsIRIS)"
    },
    
    "Traiter sessions (R04110)": {
        "action": IRIS_services.concatener_plusieursTypes,
        "kwargs": {"typesExports":"Sessions"},
        "demander": ["depuis_config"],
        "indications":"Entrez :\n   - 0 pour sélection manuelle des fichiers (filedialog)\n   - 1 pour sélection automatique (emploi de config_extractsIRIS)"
    },
    "Traiter Formations (R0304)": {
        "action": IRIS_services.concatener_plusieursTypes,
        "kwargs": {"typesExports":"Formations"},
        "demander": ["depuis_config"],
        "indications":"Entrez :\n   - 0 pour sélection manuelle des fichiers (filedialog)\n   - 1 pour sélection automatique (emploi de config_extractsIRIS)"
    },
    "Traiter Ventes (R04301)": {
        "action": IRIS_services.concatener_plusieursTypes,
        "kwargs": {"typesExports":"Ventes"},
        "demander": ["depuis_config"],
        "indications":"Entrez :\n   - 0 pour sélection manuelle des fichiers (filedialog)\n   - 1 pour sélection automatique (emploi de config_extractsIRIS)"
    },
    "Traiter Inscriptions (R04500)": {
        "action": IRIS_services.concatener_plusieursTypes,
        "kwargs": {"typesExports":"Inscriptions"},
        "demander": ["depuis_config"],
        "indications":"Entrez :\n   - 0 pour sélection manuelle des fichiers (filedialog)\n   - 1 pour sélection automatique (emploi de config_extractsIRIS)"
    },



    "Mettre a jour les dumps GED": {
    "action": mettre_a_jour_dumps_ged_console,
    "kwargs": {},
    "demander": [],
    "indications": "Copie les derniers dumps R04301 et R04500 depuis la GED et archive les anciens dumps."
},
    
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
