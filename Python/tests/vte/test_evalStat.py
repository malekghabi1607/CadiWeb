import pytest

from vte.domain.formation import Formation
from vte.domain.evalStat import *


#DATA_DIR = Path(__file__).parent / "data"


# ======================================================================================
# FIXTURES
# ======================================================================================
@pytest.fixture
def tel_csv_existant():
    return {
        "trigramme_formation": "TEL",
        "codes_IRIS": 16411,
        "chemin_eval_formation": Path(r"P:\FORMATIONS_C\TEL\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\Evaluation-Stagiaires-Global-TEL.xlsx"),
        "chemin_eval_session": Path(r"P:\FORMATIONS_C\TEL\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\S-16411-rapports-session-evaluations\S-16411-FC25-TEL-VTE-CAR-Stagiaires.csv"),

        "resultat_1er_elem_eval_formation": "S-12766-FC22-TEL-JVI-LRA",
    }


# ======================================================================================
# TESTS CLASSE EVALSTAT FORMATION
# ======================================================================================

def test_creation_nouvel_eval_formation(tel_csv_existant):

    backup_fichier_test(tel_csv_existant["chemin_eval_formation"])

    try:
        formation = Formation(tel_csv_existant["trigramme_formation"])
        formation.ajout_sessions(codes_IRIS=tel_csv_existant["codes_IRIS"])
        formation.ajout_evalStat()
        #print("\n")
        #print(len(formation.eval.df_stagiaires))
        
        assert len(formation.eval.df_stagiaires) == 0  # La variable existe mais la longueur vaut 0
        #assert formation.eval.df_stagiaires["N° Session"].loc[0] == ""
    
    except Exception as e:
        log_erreur(f"Erreur lors du test du fichier : {e}")
    
    finally:
        rollback_nom_fichier_test(tel_csv_existant["chemin_eval_formation"])

def test_ouverture_eval_formation_existant(tel_csv_existant):


    formation = Formation(tel_csv_existant["trigramme_formation"])
    formation.ajout_sessions(codes_IRIS=tel_csv_existant["codes_IRIS"])
    formation.ajout_evalStat()  # lance EvalStat_formation(formation=self)
    #print("\n")
    #print(len(formation.eval.df_stagiaires))
    
    #assert len(formation.eval.df_stagiaires) == 0  # La variable existe mais la longueur vaut 0
    assert formation.eval.df_stagiaires["N° Session"].loc[0] == tel_csv_existant["resultat_1er_elem_eval_formation"]
    
# TODO : sauvegarde eval formation
# TODO : mise à jour suite à nouveau eval session
# TODO : mise à jour suite à traitement plusieurs eval sessions

# Lancer les tests :
#    - pytest : tous les tests depuis la racine du projet
#    - pytest -v : plus verbeux
#    - ciblé : python -m pytest -v tests/vte/test_iris.py
#    - ciblé avec les print: python -m pytest -s -v tests/vte/test_evalStat.py
#    - ciblé et unitaire : python -m pytest -s -v tests/vte/test_evalStat.py::test_creation_nouvel_eval_formation

# python -m pytest -s -v tests/vte/test_evalStat.py::
# python -m pytest -s -v tests/vte/test_evalStat.py::test_creation_nouvel_eval_formation
# python -m pytest -s -v tests/vte/test_evalStat.py::test_ouverture_eval_formation_existant