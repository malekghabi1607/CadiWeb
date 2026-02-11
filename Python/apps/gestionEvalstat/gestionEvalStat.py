from vte.instn import *
from vte.ihm_console import IHM_console

from colorama import init
init(autoreset=True)


MODE = "console"

MENUS = {
    "Un seul EvalStat en sélectionnant un CSV": {
        "action": EvalStat.depuis_chemin_csv_evaluations_stagiaires,
        "kwargs": {},
        "demander": ["trigramme_formation"],
        "indications":"Entrez le trigramme de la formation (ex. : 948)"
    },

    "Un ou plusieurs EvalStat par codes IRIS (manuel)": {
        "action": EvalStat.depuis_liste_codes_IRIS,
        "kwargs": {},
        "demander": [],
        "indications":"Entrez un ou plusieurs codes IRIS séparés par des virgules (ex. : 12995, 15697)"
    },
    "Plusieurs EvalStat par code IRIS (depuis user_config.py)": {
        "action": chargement_config_demander_verif_utilisateur,
        "kwargs": {"nom_variable": "liste_codes_IRIS", 'fonction_execution':EvalStat.depuis_liste_codes_IRIS},
        "demander": []
    },
    "Plusieurs EvalStat par chemins CSV (depuis user_config.py)": {
        "action": chargement_config_demander_verif_utilisateur,
        "kwargs": {"nom_variable": "tuple_csv_stagiaires", 'fonction_execution':EvalStat.depuis_tuple_csv_stagiaires},
        "demander": []
    },
}


def main():
    
    vlog.print("Info", "*************\nBienvenue dans le script pour générer des EvalStat.\n*************", style=["vert clair"])

    ihm = IHM_console(MENUS)
    ihm.afficher_menu()

    

if __name__ == "__main__":
    main()





