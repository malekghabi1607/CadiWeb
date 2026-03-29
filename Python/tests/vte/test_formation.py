from unittest.mock import Mock, patch

import pytest

from tests.conftest import tel_data
from vte.domain.evalStat import EvalStat_formation
from vte.domain.fdc import FdC
from vte.domain.formation import Formation
from vte.domain.session import Session


# ======================================================================================
# TESTS CLASSE FORMATION
# ======================================================================================

tel = tel_data()

DATASETS_SESSIONS = {
    "Une seule session":{
        "trigramme_formation": tel["trigramme_formation"],
        "codes_IRIS": tel["code_IRIS"],
    },
    "Plusieurs sessions":{
        "trigramme_formation": tel["trigramme_formation"],
        "codes_IRIS": tel["codes_IRIS"],
    }
}

@pytest.fixture
def datasets_sessions():
    return DATASETS_SESSIONS

# ----------------------------------------------------------------------
# Test de __init__
# ----------------------------------------------------------------------
def test_init(tel):

    formation = Formation(tel["trigramme_formation"])

    assert formation.trigramme_formation == tel["trigramme_formation"]
    assert formation._eval is None
    assert formation.eval is not None
    assert formation._fdc is None
    assert formation.sessions == []
    assert formation._bilans_sessions == {}


# ----------------------------------------------------------------------
# Test des constructeurs alternatifs
# ----------------------------------------------------------------------
@patch("vte.domain.formation.EvalStat_formation.avec_ouverture_ou_creation")
def test_avec_ouverture_ou_creation_evalStat_formation(mock, tel):

    mock.return_value = Mock(spec=EvalStat_formation)

    formation = Formation.avec_ouverture_ou_creation_evalStat_formation(tel["trigramme_formation"])

    # Ce qui n'est pas sensé avoir bougé
    assert formation._fdc is None
    assert formation.sessions == []
    assert formation._bilans_sessions == {}
    
    # Ce qui est sensé avoir été affecté
    assert formation.trigramme_formation == tel["trigramme_formation"]
    assert formation.eval is mock.return_value

# Une seule session / Plusieurs sessions
@pytest.mark.parametrize("nom_cas, dataset", DATASETS_SESSIONS.items())
def test_avec_avec_ajout_sessions(nom_cas, dataset):
    formation = Formation.avec_ajout_sessions(
        trigramme_formation=dataset["trigramme_formation"],
        codes_IRIS=dataset["codes_IRIS"]
        )

    # Ce qui n'est pas sensé avoir bougé
    assert formation._eval is None
    assert formation.eval is not None
    assert formation._fdc is None
    assert formation._bilans_sessions == {}

    # Ce qui est sensé avoir été affecté
    assert formation.trigramme_formation == dataset["trigramme_formation"]

    if isinstance(dataset["codes_IRIS"], int):
        dataset["codes_IRIS"] = [dataset["codes_IRIS"]]

    assert len(formation.sessions) == len(dataset["codes_IRIS"])
    for code_IRIS in dataset["codes_IRIS"] :
        session = formation.get_session_par_codeIRIS(code_IRIS)
        assert isinstance(session, Session)
        assert session.code_IRIS == code_IRIS


# ----------------------------------------------------------------------
# Test des méthodes externes
# ----------------------------------------------------------------------
def test_ajout_sessions(tel):

    formation = Formation(tel["trigramme_formation"])

    formation.ajout_sessions(tel["codes_IRIS"])

    for session in formation.sessions:
        assert isinstance(session, Session)
        assert session.code_IRIS in tel["codes_IRIS"]

# ----------------------------------------------------------------------
# Test des méthodes externes
# ----------------------------------------------------------------------
def test_ouvrir_fdc(tel):
    formation = Formation.avec_ajout_sessions(
        trigramme_formation=tel["trigramme_formation"],
        codes_IRIS=tel["code_IRIS"]
        )
    formation.ouvrir_fdc(tel["chemin_fdc"])

    # Ce qui n'est pas sensé avoir bougé
    assert formation._eval is None
    assert formation.eval is not None
    assert formation._bilans_sessions == {}
    
    # Ce qui est sensé avoir été affecté
    assert formation.trigramme_formation == tel["trigramme_formation"]
    assert isinstance(formation._fdc, FdC)
    assert formation.fdc._formation.trigramme_formation == tel["trigramme_formation"]
    assert formation.fdc._fe.chemin_fichier == tel["chemin_fdc"]
    assert len(formation.sessions) == 1


# python -m pytest -s -v tests/vte/test_formation.py::test_ajout_sessions
@patch("vte.domain.formation.EvalStat_formation.avec_ouverture_ou_creation")
@patch("vte.domain.formation.Session.eval", new_callable=Mock)
def test_ouvrir_ou_traiter_eval_sessions(mock_session, mock_formation, tel):

    mock_formation.return_value = Mock(spec=EvalStat_formation)

    formation = Formation(tel["trigramme_formation"])
    formation.ajout_sessions(tel["codes_IRIS"])

    for session in formation.sessions:
        session.eval = Mock()

    #formation.eval = Mock()

    formation.ouvrir_ou_traiter_eval_sessions(
        ouvrirDossier=False
    )



    # Ce qui n'est pas sensé avoir bougé
    assert formation._fdc is None
    assert formation._bilans_sessions == {}

    # Ce qui est sensé avoir été affecté
    assert formation.trigramme_formation == tel["trigramme_formation"]
    assert formation.eval is mock_formation.return_value
    formation.eval.ecrit_et_sauve_df_siModif.assert_called_once()
    assert len(formation.sessions) == len(tel["codes_IRIS"])
    for session in formation.sessions:
        assert isinstance(session, Session)
        assert session.code_IRIS in tel["codes_IRIS"]
        session.eval.ouvrir_ou_traiter_eval.assert_called_once()




"""
def test_creation_avec_ouverture_evalStat_existant(tel):
    # Backup et vérif que le fichier est présent au départ
    backup_fichier_test(tel["chemin_eval_formation"], deplacement=False)
    assert tel["chemin_eval_formation"].is_file()

    try:
        formation = Formation.avec_ouverture_evalStat(tel["trigramme_formation"])
        #print("\n")
        #print(len(formation.eval.df_stagiaires))
        
        assert formation.trigramme_formation == tel["trigramme_formation"]
        assert formation.eval.fe is not None
        if formation.eval.fe is not None :
            assert (formation.eval.chemin_fe == tel["chemin_eval_formation"])
            assert len(formation.eval.df_stagiaires) > 0

    finally:
        rollback_nom_fichier_test(tel["chemin_eval_formation"])

def test_creation_avec_ouverture_evalStat_Nonexistant(tel):
    # Backup et vérif que le fichier est présent au départ
    backup_fichier_test(tel["chemin_eval_formation"], deplacement=True)
    assert not tel["chemin_eval_formation"].is_file()

    try:
        formation = Formation.avec_ouverture_evalStat(tel["trigramme_formation"])
        #print("\n")
        #print(len(formation.eval.df_stagiaires))
        
        assert formation.trigramme_formation == tel["trigramme_formation"]
        assert formation.eval.fe is not None
        if formation.eval.fe is not None :
            assert (formation.eval.chemin_fe == tel["chemin_eval_formation"])
            assert len(formation.eval.df_stagiaires) == 0
    finally:
        rollback_nom_fichier_test(tel["chemin_eval_formation"])

def test_creation_avec_creation_une_session(tel):

    # On crée la formation
    formation = Formation.avec_creation_sessions(
        trigramme_formation=tel["trigramme_formation"],
        codes_IRIS=tel["code_IRIS"]
    )
    #print("\n")
    #print(len(formation.eval.df_stagiaires))
    
    assert formation.trigramme_formation == tel["trigramme_formation"]
    assert len(formation.sessions) == 1
    assert formation.sessions[tel["code_IRIS"]].code_IRIS == tel["code_IRIS"]

def test_creation_avec_creation_plusieurs_sessions(tel):

    # On crée la formation
    formation = Formation.avec_creation_sessions(
        trigramme_formation=tel["trigramme_formation"],
        codes_IRIS=tel["codes_IRIS"]
    )
    #print("\n")
    #print(len(formation.eval.df_stagiaires))
    
    assert formation.trigramme_formation == tel["trigramme_formation"]
    assert len(formation.sessions) == len(tel["codes_IRIS"])
    assert formation.sessions[tel["codes_IRIS"][-1]].code_IRIS == tel["codes_IRIS"][-1]



    # Lancer les tests :

"""

#    - pytest : tous les tests depuis la racine du projet
#    - pytest -v : plus verbeux
#    - ciblé : python -m pytest -v tests/vte/test_formation.py
#    - ciblé avec les print: python -m pytest -s -v tests/vte/test_evalStat.py
#    - ciblé et unitaire : python -m pytest -s -v tests/vte/test_evalStat.py::test_creation_nouvel_eval_formation

# python -m pytest -s -v tests/vte/test_formation.py::

# python -m pytest -v tests/vte/test_formation.py