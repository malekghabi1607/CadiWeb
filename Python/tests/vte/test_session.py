from pathlib import Path
from unittest.mock import Mock, patch

import pytest

from vte.domain.evalStat import EvalStat_session
from vte.domain.formation import Formation
from vte.domain.session import Session
from vte.utils.utils import backup_fichier_test, chemin_vers_unc, rollback_nom_fichier_test


# ======================================================================================
# FIXTURES
# ======================================================================================
@pytest.fixture
def tel():
    return {
        "trigramme_formation": "TEL",
        "code_IRIS": 16411,
        "codes_IRIS": (12766, 13414, 15942, 16161, 16411),

        "chemin_eval_formation": chemin_vers_unc(Path(r"P:\FORMATIONS_C\TEL\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\Evaluation-Stagiaires-Global-TEL.xlsx")),
    }

@pytest.fixture
def formation(tel):
    return Formation(tel["trigramme_formation"])

# ======================================================================================
# TESTS CLASSE EVALSTAT SESSION
# ======================================================================================

# ----------------------------------------------------------------------
# Test de __init__
# ----------------------------------------------------------------------
# python -m pytest -s -v tests/vte/test_session.py::test_init
def test_init(formation, tel):
    session = Session(formation=formation, code_IRIS=tel["code_IRIS"])
    assert isinstance(session, Session)
    assert session.trigramme_formation == tel["trigramme_formation"]
    assert session.code_IRIS == tel["code_IRIS"]
    assert session.eval is None

# ----------------------------------------------------------------------
# Test des méthodes internes
# ----------------------------------------------------------------------
# python -m pytest -s -v tests/vte/test_session.py::test_ajout_evalStat
@patch("vte.domain.session.EvalStat_session")
def test_ajout_evalStat(mock, formation, tel):
    session = Session(formation=formation, code_IRIS=tel["code_IRIS"])
    session._ajout_evalStat()
    
    mock.assert_called_once_with(session=session) # On teste que l'appel a bien été fait avec les bons arguments
    assert session.eval == mock.return_value  # On teste la valeur retournée

    #assert isinstance(session.eval, EvalStat_session)

# python -m pytest -s -v tests/vte/test_session.py::test_ajout_evalStat_avec_traitement
@patch("vte.domain.session.EvalStat_session.avec_traitement")
def test_ajout_evalStat_avec_traitement(mock, formation, tel):
    session = Session(formation=formation, code_IRIS=tel["code_IRIS"])
    session._ajout_evalStat_avec_traitement(
        chemin_csv=None,
        chemin_IRIS_sessions=None,
        ecrire_eval_formation=False,
        ouvrirDossier=False
    )

    mock.assert_called_once_with(
        session=session,
        chemin_csv=None,
        chemin_IRIS_sessions=None,
        ecrire_eval_formation=False,
        ouvrirDossier=False

    )
    assert session.eval == mock.return_value
    

# ----------------------------------------------------------------------
# Test des constructeurs alternatifs
# ----------------------------------------------------------------------
# python -m pytest -s -v tests/vte/test_session.py::test_avec_traitement_evalStat
@patch("vte.domain.session.Session._ajout_evalStat_avec_traitement")
def test_avec_traitement_evalStat(mock, formation, tel):
    session = Session.avec_traitement_evalStat(
        formation=formation, 
        code_IRIS=tel["code_IRIS"],
        chemin_csv=None,
        chemin_IRIS_sessions=None,
        ecrire_eval_formation=False,
        ouvrirDossier=False
        )
    

    assert session.trigramme_formation == tel["trigramme_formation"]
    assert session.code_IRIS == tel["code_IRIS"]
    mock.assert_called_once_with(
        session=session,
        chemin_csv=None,
        chemin_IRIS_sessions=None,
        ecrire_eval_formation=False,
        ouvrirDossier=False

    )










"""
def test_creation_avec_creation_une_session(tel):

    # On crée la formation
    formation = Formation(trigramme_formation=tel["trigramme_formation"])
    formation.
    
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
"""


    # Lancer les tests :
#    - pytest : tous les tests depuis la racine du projet
#    - pytest -v : plus verbeux
#    - ciblé : python -m pytest -v tests/vte/test_formation.py
#    - ciblé avec les print: python -m pytest -s -v tests/vte/test_evalStat.py
#    - ciblé et unitaire : python -m pytest -s -v tests/vte/test_evalStat.py::test_creation_nouvel_eval_formation

# python -m pytest -s -v tests/vte/test_session.py::
# python -m pytest -s -v tests/vte/test_session.py::test_creation_nouvel_eval_formation
# python -m pytest -s -v tests/vte/test_session.py::test_ouverture_eval_formation_existant


#python -m pytest -v tests/vte/test_formation.py tests/vte/test_session.py
#python -m pytest -v tests/vte/test_session.py