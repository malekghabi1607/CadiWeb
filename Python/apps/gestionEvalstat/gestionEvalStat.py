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



def test_ajout_csv_a_evaluationFormation_existant():
    """
    Pour debugger le fait qu'actuellement si l'on rajoute un nouveau csv à un fichier Excel de l'évaluation d'un formation est ajouté, alors les 3 colonnes merdent et sont rajoutés après le tableau.

    Donc je :
      - réinitialise (supprime l'excel évaluation formation ; copie-colle un backup et le renomme) ;
      - relance ma fonction
    """


    fichier_evalStat_formation:Path = chemin_vers_unc(r"P:\FORMATIONS_C\948\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\Evaluation-Stagiaires-Global-948.xlsx", retour_type=Path)
    fichier_evalStat_formation_BAK:Path = chemin_vers_unc(r"P:\FORMATIONS_C\948\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\Evaluation-Stagiaires-Global-948 - Bak sans 17139.xlsx", retour_type=Path)

    chemin_csv_stagiaires:Path = chemin_vers_unc(r"P:\FORMATIONS_C\948\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2025-12-S17139 UEM\S-17139-FC25-948-VTE-VCA-Stagiaires.csv", retour_type=Path)
    chemin_IRIS_sessions:Path = chemin_vers_unc(r"R:\_Echanges\VTE\Prog\IRIS\Extracts complets\TESTS - 948 - R04110_Sessions-COMPLET.xlsx", retour_type=Path)

    # Supprimer l'ancien fichier
    if fichier_evalStat_formation.exists():
        fichier_evalStat_formation.unlink()
    
    # Copier le backup
    if fichier_evalStat_formation_BAK.exists():
        shutil.copy2(fichier_evalStat_formation_BAK, fichier_evalStat_formation)

    # Lancer la fonction
    EvalStat.depuis_chemin_csv_evaluations_stagiaires(chemin_csv_stagiaires=chemin_csv_stagiaires, chemin_IRIS_sessions=chemin_IRIS_sessions)

if __name__ == "__main__":
    test_ajout_csv_a_evaluationFormation_existant()
    #main()




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





