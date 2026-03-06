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
        "chemin_csv_session": Path(r"P:\FORMATIONS_C\TEL\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\S-16411-rapports-session-evaluations\S-16411-FC25-TEL-VTE-CAR-Stagiaires.csv"),
        "chemin_eval_session": Path(r"P:\FORMATIONS_C\TEL\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\S-16411-rapports-session-evaluations\S-16411-FC25-TEL-VTE-CAR-Stagiaires.xlsx"),

        "chemin_IRIS_sessions": Path(r"R:\_Echanges\VTE\Prog\IRIS\Extracts complets\TESTS - TEL - R04110_Sessions-COMPLET.xlsx"),

        "resultat_1er_elem_eval_formation": "S-12766-FC22-TEL-JVI-LRA",
    }


# ======================================================================================
# TESTS CLASSE EVALSTAT FORMATION
# ======================================================================================

def test_creation_nouvel_eval_formation(tel_csv_existant):
    """
    Pas d'eval foramtion du tout
    """

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


# Création avec eval formation existant ne contenant pas la ref du CSV (on le remplit)
# Création avec eval formation existant mais contenant la ref du CSV (on saute le traitement) instance._statut_csv = "Exclu - CSV déjà dans fichier global" ; instance._fe = None
# Création avec eval formation existant ne contenant pas la ref du CSV mais IRIS sessions ne contient pas le code IRIS instance._statut_csv = "Exclu - Code IRIS pas dans Extract IRIS sessions" ; instance._fe = None
# Pb lecture du CSV : self._statut_csv = "Exclu - Problème lecture CSV" ; instance._fe = None
# csv stagiaire vide : self._statut_csv = "Exclu - CSV vide / Aucun retour" ; instance._fe = None
# Plusieurs cas de CSV : 
# Cas 1 (nouveau format de csv) : supprimer "Date de fin" → Test : r"P:\FORMATIONS_C\54C\P07-bilan-sessions-et-bilan-formation\rapports-sessions-evaluations\2023-06-S14317 UEM\S-14317-FC23-54C-VTE-LRA-Stagiaires.csv"
# Cas 2 (ancien format de csv) : supprimer la 2e et 3e colonne (indices 1 et 2) → Test : r"P:\FORMATIONS_C\22B\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\S-17606 - 22B - 06-2025\S-17606-FC25-22B-VTE-CAR-Stagiaires.csv"
#Quand on màj eval foramtion, vérifier que la liste des CSV est bien à jour

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