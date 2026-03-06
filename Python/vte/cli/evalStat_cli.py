from vte.domain.formation import Formation
from vte.utils.utils import *

# ==========================================================================================
# CLASSE EVALSTAT_CLI
#
# Command Line Interface
# Pour lancer des fonctions et méthodes en CLI relatives à EvalStat
# ==========================================================================================



class evalStat_cli:
    """
    Services Command Line Interface d'EvalStat.
    """
    tel_csv_existant = {
        "trigramme_formation": "TEL",
        "codes_IRIS": 16411,
        "chemin_eval_formation": Path(r"P:\FORMATIONS_C\TEL\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\Evaluation-Stagiaires-Global-TEL.xlsx"),
        "chemin_csv_session": Path(r"P:\FORMATIONS_C\TEL\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\S-16411-rapports-session-evaluations\S-16411-FC25-TEL-VTE-CAR-Stagiaires.csv"),
        "chemin_eval_session": Path(r"P:\FORMATIONS_C\TEL\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\S-16411-rapports-session-evaluations\S-16411-FC25-TEL-VTE-CAR-Stagiaires.xlsx"),

        "chemin_IRIS_sessions": Path(r"R:\_Echanges\VTE\Prog\IRIS\Extracts complets\TESTS - TEL - R04110_Sessions-COMPLET.xlsx"),

        "resultat_1er_elem_eval_formation": "S-12766-FC22-TEL-JVI-LRA",
    }


    backup_fichier_test(tel_csv_existant["chemin_eval_formation"])
    backup_fichier_test(tel_csv_existant["chemin_eval_session"])

    formation = Formation.avec_ouverture_evalStat("TEL")
    formation.ajout_sessions(codes_IRIS=16411)
    #formation.ajout_evalStat()  # EvalStat formation
    formation.sessions[16411]
    formation.sessions[16411].ajout_evalStat()

    formation.sessions[16411].eval_session.avec_traitement_depuis_chemin_csv(
        session=formation.sessions[16411],
        chemin_csv=tel_csv_existant["chemin_csv_session"],
        chemin_IRIS_sessions=tel_csv_existant["chemin_IRIS_sessions"],
        ouvrirDossier=True)

    
