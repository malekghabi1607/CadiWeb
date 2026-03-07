from unittest.mock import Mock, patch

import pytest

from vte.domain.formation import Formation
from vte.domain.evalStat import *
from vte.domain.session import Session


#DATA_DIR = Path(__file__).parent / "data"


# ======================================================================================
# FIXTURES
# ======================================================================================
@pytest.fixture
def tel():
    return {
        "trigramme_formation": "TEL",
        "code_IRIS": 16411,
        "chemin_eval_formation": chemin_vers_unc(Path(r"P:\FORMATIONS_C\TEL\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\Evaluation-Stagiaires-Global-TEL.xlsx")),
        "chemin_csv_session": chemin_vers_unc(Path(r"P:\FORMATIONS_C\TEL\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\S-16411-rapports-session-evaluations\S-16411-FC25-TEL-VTE-CAR-Stagiaires.csv")),
        "chemin_csv_session_non_present_eval_formation": chemin_vers_unc(Path(r"P:\FORMATIONS_C\948\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2022-11-S-12995 UEM\S-12995-FC22-948-VTE-SNA-Stagiaires.csv")),
        "chemin_eval_session": chemin_vers_unc(Path(r"P:\FORMATIONS_C\TEL\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\S-16411-rapports-session-evaluations\S-16411-FC25-TEL-VTE-CAR-Stagiaires.xlsx")),

        "chemin_IRIS_sessions": chemin_vers_unc(Path(r"R:\_Echanges\VTE\Prog\IRIS\Extracts complets\TESTS - TEL - R04110_Sessions-COMPLET.xlsx")),

        "resultat_1er_elem_eval_formation": "S-12766-FC22-TEL-JVI-LRA",
    }


@pytest.fixture
def formation(tel):
    """Formation réelle avec son EvalStat ouvert"""
    return Formation.avec_ouverture_evalStat(tel["trigramme_formation"])


@pytest.fixture
def session(formation, tel):
    """Session réelle attachée à la formation"""
    formation.ajout_sessions(codes_IRIS=tel["code_IRIS"])
    session = formation.sessions[tel["code_IRIS"]]
    return session




# ======================================================================================
# TESTS CLASSE EVALSTAT
# ======================================================================================
# ----------------------------------------------------------------------
# Test de __init__
# ----------------------------------------------------------------------
# python -m pytest -s -v tests/vte/test_evalStat.py::test_init
def test_init(tel):
    evalStat = EvalStat()
    assert isinstance(evalStat, EvalStat)
    assert evalStat.fe is None




# ----------------------------------------------------------------------
# Test des constructeurs alternatifs
# ----------------------------------------------------------------------





# ----------------------------------------------------------------------
# Test des méthodes internes
# ----------------------------------------------------------------------




# ----------------------------------------------------------------------
# Test des méthodes externes
# ----------------------------------------------------------------------






# ======================================================================================
# TESTS CLASSE EVALSTAT SESSION
# ======================================================================================
# ----------------------------------------------------------------------
# Test de __init__
# ----------------------------------------------------------------------
# python -m pytest -s -v tests/vte/test_evalStat.py::test_init_evalStat_session
def test_init_evalStat_session(session):

    es = EvalStat_session(session)

    assert es._session == session
    assert es.fe is None



# ----------------------------------------------------------------------
# Test des constructeurs alternatifs
# ----------------------------------------------------------------------
@patch("vte.domain.evalStat.EvalStat_session.traiter_eval")
def test_avec_traitement_evalStat_session(mock_traiter, session):

    es = EvalStat_session.avec_traitement(session)

    assert isinstance(es, EvalStat_session)

    mock_traiter.assert_called_once()






# ----------------------------------------------------------------------
# Test des méthodes internes
# ----------------------------------------------------------------------
# TODO : faire mieux
@patch("vte.domain.evalStat.trouve_encodage_csv")
@patch("pandas.read_csv")
def test_charger_csv_stagiaire(mock_read_csv, mock_encodage, session, tel):

    df = pd.DataFrame({
        "Nom": ["Dupont"],
        "Prénom": ["Jean"]
    })

    mock_encodage.return_value = "utf-8"
    mock_read_csv.return_value = df

    es = EvalStat_session(session)

    result = es._charger_csv_stagiaire(tel["chemin_csv_session"])

    assert isinstance(result, pd.DataFrame)
    assert not result.empty

# TODO : bug
"""
def test_traiter_df_csv(session):

    es = EvalStat_session(session)

    df = pd.DataFrame({
        "Nom": ["Dupont"],
        "Prénom": ["Jean"],
        "Date": ["01/01/2024"]
    })

    es.df_csv = df

    es._fe = Mock()
    es._fe.get_df_tableau.return_value = df
    es._fe.set_df_tableau = Mock()

    es._traiter_df_csv()

    assert "Code session" in es.df_csv.columns
"""

# TODO : bug
"""
def test_maj_dataframes_eval_formation(session):

    es = EvalStat_session(session)

    es.df_csv = pd.DataFrame({"A": [1]})
    es.df_stagiaires = pd.DataFrame({"B": [2]})

    session.eval_formation.df_csv = None
    session.eval_formation.df_stagiaires = None

    es._maj_dataframes_eval_formation()

    assert session.eval_formation.df_csv is not None
    assert session.eval_formation.df_stagiaires is not None
"""



# ----------------------------------------------------------------------
# Test des méthodes externes
# ----------------------------------------------------------------------





# ======================================================================================
# TESTS CLASSE EVALSTAT FORMATION
# ======================================================================================
# ----------------------------------------------------------------------
# Test de __init__
# ----------------------------------------------------------------------
def test_init_evalStat_formation(formation):

    es = EvalStat_formation(formation)

    assert es.trigramme_formation == formation.trigramme_formation


# ----------------------------------------------------------------------
# Test des constructeurs alternatifs
# ----------------------------------------------------------------------
@patch("vte.domain.evalStat.EvalStat_formation._ouvrir_ou_creer_eval_formation")
def test_avec_ouverture_evalStat_formation(mock_ouvrir, formation):

    es = EvalStat_formation.avec_ouverture(formation)

    assert isinstance(es, EvalStat_formation)

    mock_ouvrir.assert_called_once()




# ----------------------------------------------------------------------
# Test des méthodes internes
# ----------------------------------------------------------------------






# ----------------------------------------------------------------------
# Test des méthodes externes
# ----------------------------------------------------------------------





# TODO : bug
"""
@patch("vte.domain.evalStat.hash_df")
def test_ecritdf_et_sauve_siModif(mock_hash, formation):

    es = EvalStat_formation(formation)

    es._fe = Mock()

    mock_hash.side_effect = ["hash1", "hash2"]

    es._df_initial_hash = "hash1"

    es.ecritdf_et_sauve_siModif()

    es._fe.get_tableau.assert_called()

"""


#Mes anciens tests
"""
# python -m pytest -s -v tests/vte/test_evalStat.py::test_creation_nouvel_eval_formation
def test_creation_nouvel_eval_formation(tel_csv_existant):
    # Backup et vérif que le fichier n'est pas présent au départ
    backup_fichier_test(tel_csv_existant["chemin_eval_formation"])
    assert not tel_csv_existant["chemin_eval_formation"].is_file()

    try :
        formation = Formation.avec_ouverture_evalStat(tel_csv_existant["trigramme_formation"])
        #print("\n")
        #print(len(formation.eval.df_stagiaires))
        
        assert formation.eval.fe is not None
        if formation.eval.fe is not None :
            assert (formation.eval.chemin_fe == tel_csv_existant["chemin_eval_formation"])
            assert len(formation.eval.df_stagiaires) == 0

    finally:
        rollback_nom_fichier_test(tel_csv_existant["chemin_eval_formation"])

# python -m pytest -s -v tests/vte/test_evalStat.py::test_creation_ouvrir_eval_formation_existant
def test_creation_ouvrir_eval_formation_existant(tel_csv_existant):
    # Backup et vérif que le fichier est présent au départ
    backup_fichier_test(tel_csv_existant["chemin_eval_formation"], deplacement=False)
    assert tel_csv_existant["chemin_eval_formation"].is_file()

    try:
        formation = Formation.avec_ouverture_evalStat(tel_csv_existant["trigramme_formation"])
        #print("\n")
        #print(len(formation.eval.df_stagiaires))
        
        assert formation.eval.fe is not None
        if formation.eval.fe is not None :
            assert (formation.eval.chemin_fe == tel_csv_existant["chemin_eval_formation"])
            assert len(formation.eval.df_stagiaires) > 0

    finally:
        rollback_nom_fichier_test(tel_csv_existant["chemin_eval_formation"])

# Création avec eval formation existant ne contenant pas la ref du CSV (on le remplit)
def test_eval_session_refCSV_inexistant_dans_eval_formation(tel_csv_existant):

    # Backup et vérif que le fichier est présent au départ
    backup_fichier_test(tel_csv_existant["chemin_eval_formation"], deplacement=False)
    assert tel_csv_existant["chemin_eval_formation"].is_file()

    try:
        # On crée la formation
        formation = Formation.avec_creation_sessions(
            trigramme_formation=tel_csv_existant["trigramme_formation"],
            codes_IRIS=tel_csv_existant["code_IRIS"]
        )
        # chemin_csv_session_non_present_eval_formation
        print("\n")
        #print(len(formation.eval.df_stagiaires))
        
        assert formation.eval.fe is not None
        if formation.eval.fe is not None :
            assert (formation.eval.chemin_fe == tel_csv_existant["chemin_eval_formation"])
            assert len(formation.eval.df_stagiaires) > 0
        assert formation.sessions[tel_csv_existant["code_IRIS"]].code_IRIS == tel_csv_existant["code_IRIS"]
        
    finally:
        rollback_nom_fichier_test(tel_csv_existant["chemin_eval_formation"])

"""


# Création avec eval formation existant mais contenant la ref du CSV (on saute le traitement) instance._statut_csv = "Exclu - CSV déjà dans fichier global" ; instance._fe = None
# Création avec eval formation existant ne contenant pas la ref du CSV mais IRIS sessions ne contient pas le code IRIS instance._statut_csv = "Exclu - Code IRIS pas dans Extract IRIS sessions" ; instance._fe = None
# Pb lecture du CSV : self._statut_csv = "Exclu - Problème lecture CSV" ; instance._fe = None
# csv stagiaire vide : self._statut_csv = "Exclu - CSV vide / Aucun retour" ; instance._fe = None
# Plusieurs cas de CSV : 
#    Cas 1 (nouveau format de csv) : supprimer "Date de fin" → Test : r"P:\FORMATIONS_C\54C\P07-bilan-sessions-et-bilan-formation\rapports-sessions-evaluations\2023-06-S14317 UEM\S-14317-FC23-54C-VTE-LRA-Stagiaires.csv"
#    Cas 2 (ancien format de csv) : supprimer la 2e et 3e colonne (indices 1 et 2) → Test : r"P:\FORMATIONS_C\22B\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\S-17606 - 22B - 06-2025\S-17606-FC25-22B-VTE-CAR-Stagiaires.csv"
# Quand on màj eval formation, vérifier que la liste des CSV est bien à jour

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