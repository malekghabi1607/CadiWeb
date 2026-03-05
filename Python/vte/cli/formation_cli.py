from pathlib import Path

from vte.domain.formation import Formation
from vte.utils.utils import backup_fichier_test, rollback_nom_fichier_test, log_erreur
# ==========================================================================================
# CLASSE EVALSTAT_CLI
#
# Command Line Interface
# Pour lancer des fonctions et méthodes en CLI relatives à EvalStat
# ==========================================================================================



class formation_cli:
    """
    Services Command Line Interface d'EvalStat.
    """
    
    chemin = Path(r"P:\FORMATIONS_C\TEL\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\Evaluation-Stagiaires-Global-TEL.xlsx")

    backup_fichier_test(chemin)

    try:
        formation = Formation("TEL")
        formation.ajout_sessions(codes_IRIS=16411)
        formation.ajout_evalStat()

        print()
        print(formation.eval in locals())
        print(len(formation.eval.df_stagiaires))
        #print(formation.eval.df_stagiaires["N° Session"].loc[0] == "")
        
    except Exception as e:
        log_erreur(f"Erreur lors du test du fichier : {e}")
    
    finally:
        rollback_nom_fichier_test(chemin)
    



