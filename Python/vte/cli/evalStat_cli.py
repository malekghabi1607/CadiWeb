from pprint import pprint

from vte.core.iris_referentiel import set_iris_chemin_specifique
from vte.domain.evalStat import EvalStat_formation, EvalStat_session
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
# === TRAITEMENT tous les EVALSTAT
# ======================================================================================
def traite_tous_evalStats_depuis_tuple_csv(trigramme_formation:str="Tous") -> None:
    """
    Traite tous les evalstat d'une formation à partir des chemins csv donnés dans user_config_evalStat.py

    Si trigramme_formation = "Tous", alors on traite toutes les formations et pas qu'un seul

    :param trigramme_formation: Trigramme de la formation pour laquelle on souhaite générer tous les EvalStat. Défaut = "Tous"
    :type trigramme_formation: str, optional
    """

    match trigramme_formation:
        case "948":
            tuple_csv_stagiaires_employe = tuple_csv_stagiaires_948
        case "TEL":
            tuple_csv_stagiaires_employe = tuple_csv_stagiaires_TEL
        case "22B":
            tuple_csv_stagiaires_employe = tuple_csv_stagiaires_22B
        case "Tous":
            tuple_csv_stagiaires_employe = tuple_csv_stagiaires_Tous

    # Chemin des anciens fichiers eval formation ; 
    f_eval = EvalStat_formation.construire_chemin_eval_formation(trigramme_formation=trigramme_formation)
    #f_eval_948 = Path(r"P:\FORMATIONS_C\948\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\Evaluation-Stagiaires-Global-948.xlsx")
    #f_eval_TEL = Path(r"P:\FORMATIONS_C\TEL\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\Evaluation-Stagiaires-Global-TEL.xlsx")
    #f_eval_22B = Path(r"P:\FORMATIONS_C\22B\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\Evaluation-Stagiaires-Global-22B.xlsx")

    # On supprime les anciens fichiers eval formation
    f_eval.unlink()
    #f_eval_948.unlink()
    #f_eval_TEL.unlink()
    #f_eval_22B.unlink()

    statuts_csv = EvalStat_services.ouvrir_ou_traiter_evalStat_depuis_iterable_de_csv(
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

