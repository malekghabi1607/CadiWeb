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
        "code_IRIS": 16411,
        "chemin_eval_formation": chemin_vers_unc(Path(r"P:\FORMATIONS_C\TEL\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\Evaluation-Stagiaires-Global-TEL.xlsx")),
        "chemin_csv_session": chemin_vers_unc(Path(r"P:\FORMATIONS_C\TEL\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\S-16411-rapports-session-evaluations\S-16411-FC25-TEL-VTE-CAR-Stagiaires.csv")),
        "chemin_eval_session": chemin_vers_unc(Path(r"P:\FORMATIONS_C\TEL\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\S-16411-rapports-session-evaluations\S-16411-FC25-TEL-VTE-CAR-Stagiaires.xlsx")),

        "chemin_IRIS_sessions": chemin_vers_unc(Path(r"R:\_Echanges\VTE\Prog\IRIS\Extracts complets\TESTS - TEL - R04110_Sessions-COMPLET.xlsx")),

        "resultat_1er_elem_eval_formation": "S-12766-FC22-TEL-JVI-LRA",
    }


# ======================================================================================
# TESTS CLASSE EVALSTAT FORMATION
# ======================================================================================
# python -m pytest -s -v tests/vte/test_evalStat.py::test_creation_nouvel_eval_formation
def test_creation_nouvel_eval_formation(tel_csv_existant):
    """
    Pas d'eval formation existant initialement → On en crée un
    """

    backup_fichier_test(tel_csv_existant["chemin_eval_formation"])
    
    # Vérif que le fichier n'est pas présent au départ
    assert not tel_csv_existant["chemin_eval_formation"].is_file()

    formation = Formation.avec_ouverture_evalStat(tel_csv_existant["trigramme_formation"])
    print("\n")
    #print(len(formation.eval.df_stagiaires))
    
    assert formation.eval.fe is not None
    if formation.eval.fe is not None :
        assert (formation.eval.chemin_fe == tel_csv_existant["chemin_eval_formation"])
        assert len(formation.eval.df_stagiaires) == 0

    rollback_nom_fichier_test(tel_csv_existant["chemin_eval_formation"])

# python -m pytest -s -v tests/vte/test_evalStat.py::test_creation_ouvrir_eval_formation_existant
def test_creation_ouvrir_eval_formation_existant(tel_csv_existant):
    """
    Eval formation existant initialement → On l'ouvre
    """
    # Backup et vérif que le fichier est présent au départ
    backup_fichier_test(tel_csv_existant["chemin_eval_formation"], deplacement=False)
    assert tel_csv_existant["chemin_eval_formation"].is_file()

    formation = Formation.avec_ouverture_evalStat(tel_csv_existant["trigramme_formation"])
    print("\n")
    #print(len(formation.eval.df_stagiaires))
    
    assert formation.eval.fe is not None
    if formation.eval.fe is not None :
        assert (formation.eval.chemin_fe == tel_csv_existant["chemin_eval_formation"])
        assert len(formation.eval.df_stagiaires) > 0

    rollback_nom_fichier_test(tel_csv_existant["chemin_eval_formation"])

# Création avec eval formation existant ne contenant pas la ref du CSV (on le remplit)
def test_creation_eval_session_refCSV_inexistant_dans_eval_formation(tel_csv_existant):

    # Backup et vérif que le fichier est présent au départ
    backup_fichier_test(tel_csv_existant["chemin_eval_formation"], deplacement=False)
    assert tel_csv_existant["chemin_eval_formation"].is_file()

    # On crée la formation
    formation = Formation.avec_creation_sessions(
        trigramme_formation=tel_csv_existant["trigramme_formation"],
        codes_IRIS=tel_csv_existant["code_IRIS"]
    )
    print("\n")
    #print(len(formation.eval.df_stagiaires))
    
    assert formation.eval.fe is not None
    if formation.eval.fe is not None :
        assert (formation.eval.chemin_fe == tel_csv_existant["chemin_eval_formation"])
        assert len(formation.eval.df_stagiaires) > 0
    assert formation.sessions[tel_csv_existant["code_IRIS"]].code_IRIS == tel_csv_existant["code_IRIS"]


    rollback_nom_fichier_test(tel_csv_existant["chemin_eval_formation"])

# Création avec eval formation existant ne contenant pas la ref du CSV (on le remplit)
def test_eval_session_refCSV_inexistant_dans_eval_formation(tel_csv_existant):

    # Backup et vérif que le fichier est présent au départ
    backup_fichier_test(tel_csv_existant["chemin_eval_formation"], deplacement=False)
    assert tel_csv_existant["chemin_eval_formation"].is_file()

    # On crée la formation
    formation = Formation.avec_creation_sessions(
        trigramme_formation=tel_csv_existant["trigramme_formation"],
        codes_IRIS=tel_csv_existant["code_IRIS"]
    )
    print("\n")
    #print(len(formation.eval.df_stagiaires))
    
    assert formation.eval.fe is not None
    if formation.eval.fe is not None :
        assert (formation.eval.chemin_fe == tel_csv_existant["chemin_eval_formation"])
        assert len(formation.eval.df_stagiaires) > 0
    assert formation.sessions[tel_csv_existant["code_IRIS"]].code_IRIS == tel_csv_existant["code_IRIS"]
    

    rollback_nom_fichier_test(tel_csv_existant["chemin_eval_formation"])

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