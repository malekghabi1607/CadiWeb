from pprint import pprint

from vte.core.iris_referentiel import set_iris_chemin_specifique
from vte.domain.evalStat import EvalStat_session
from vte.domain.formation import Formation
from vte.services.evalStat_services import EvalStat_services
from vte.utils.utils import *

from vte.core.user_config_evalStat import *

# ==========================================================================================
# CLASSE EVALSTAT_CLI
#
# Command Line Interface
# Pour lancer des fonctions et méthodes en CLI relatives à EvalStat
# ==========================================================================================



# ======================================================================================
# === TESTS
# ======================================================================================
def verifications():
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

    set_iris_chemin_specifique(typeExport="Sessions", chemin=tel_csv_existant["chemin_IRIS_sessions"])  # On redéfinit le chemin d'IRIS Sessions
    
    formation = Formation.avec_ouverture_evalStat("TEL")
    formation.ajout_sessions(codes_IRIS=16411)
    formation.sessions[16411].eval = EvalStat_session.avec_ouverture_ou_traitement(
            session=formation.sessions[16411],
            chemin_csv=tel_csv_existant["chemin_csv_session"],
            ouvrirDossier=True
            )



# ======================================================================================
# === TRAITEMENT tous les IRIS
# ======================================================================================
def traite_tous_evalStats_depuis_tuple_csv():

    # Pour supprimer les anciens ficheirs globaux
    f_eval_948 = Path(r"P:\FORMATIONS_C\948\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\Evaluation-Stagiaires-Global-948.xlsx")
    f_eval_TEL = Path(r"P:\FORMATIONS_C\TEL\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\Evaluation-Stagiaires-Global-TEL.xlsx")
    f_eval_22B = Path(r"P:\FORMATIONS_C\22B\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\Evaluation-Stagiaires-Global-22B.xlsx")

    f_eval_948.unlink()
    #f_eval_TEL.unlink()
    #f_eval_22B.unlink()

    tuple_csv_stagiaires_employe = tuple_csv_stagiaires

    statuts_csv = EvalStat_services.traiter_evalStat_depuis_iterable_de_csv(
        tuple_csv_stagiaires= tuple_csv_stagiaires_employe,
        ouvrirDossier=False
    )

    pprint(statuts_csv)


# ======================================================================================
# === TESTS
# ======================================================================================

def main():
    traite_tous_evalStats_depuis_tuple_csv()

if __name__ == "__main__":
    #test07()
    main()

