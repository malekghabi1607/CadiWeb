from unittest.mock import Mock, patch

from vte.domain.evalStat import EvalStat_formation
from vte.domain.formation import Formation
from vte.domain.session import Session


# ======================================================================================
# TESTS CLASSE FORMATION
# ======================================================================================

# ----------------------------------------------------------------------
# Test de __init__
# ----------------------------------------------------------------------
# python -m pytest -s -v tests/vte/test_formation.py::test_init
def test_init(tel):

    formation = Formation(tel["trigramme_formation"])

    assert formation.trigramme_formation == tel["trigramme_formation"]
    assert formation.eval is None
    assert formation.sessions == {}


# ----------------------------------------------------------------------
# Test des constructeurs alternatifs
# ----------------------------------------------------------------------
# python -m pytest -s -v tests/vte/test_formation.py::test_avec_ouverture_evalStat
@patch("vte.domain.formation.EvalStat_formation.avec_ouverture")
def test_avec_ouverture_evalStat(mock, tel):

    mock.return_value = Mock(spec=EvalStat_formation)

    formation = Formation.avec_ouverture_evalStat(tel["trigramme_formation"])

    assert formation.eval is mock.return_value

# Une seule session
# python -m pytest -s -v tests/vte/test_formation.py::test_avec_creation_session
@patch("vte.domain.formation.EvalStat_formation.avec_ouverture")
def test_avec_creation_session(mock, tel):
    formation = Formation.avec_creation_sessions(
        trigramme_formation=tel["trigramme_formation"],
        codes_IRIS=tel["code_IRIS"]
        )

    assert formation.trigramme_formation == tel["trigramme_formation"]
    assert formation.eval is None
    assert len(formation.sessions) == 1
    assert formation.sessions[tel["code_IRIS"]].code_IRIS == tel["code_IRIS"]

# Plusieurs sessions
# python -m pytest -s -v tests/vte/test_formation.py::test_avec_creation_sessions
@patch("vte.domain.formation.EvalStat_formation.avec_ouverture")
def test_avec_creation_sessions(mock, tel):
    formation = Formation.avec_creation_sessions(
        trigramme_formation=tel["trigramme_formation"],
        codes_IRIS=tel["codes_IRIS"]
        )

    assert formation.trigramme_formation == tel["trigramme_formation"]
    assert formation.eval is None
    assert len(formation.sessions) == len(tel["codes_IRIS"])
    for code_IRIS in tel["codes_IRIS"] :
        assert formation.sessions[code_IRIS].code_IRIS == code_IRIS

# ----------------------------------------------------------------------
# Test des méthodes internes
# ----------------------------------------------------------------------
# python -m pytest -s -v tests/vte/test_formation.py::test_avec_ouverture_evalStat
# def test_ajout_sessions(tel):

    formation = Formation(tel["trigramme_formation"])

    formation.ajout_sessions(tel["codes_IRIS"])

    for code_IRIS in tel["codes_IRIS"]:
        assert isinstance(formation.sessions[code_IRIS], Session)


# ----------------------------------------------------------------------
# Test des méthodes externes
# ----------------------------------------------------------------------
# python -m pytest -s -v tests/vte/test_formation.py::test_ajout_sessions
@patch("vte.domain.formation.Session.eval", new_callable=Mock)
def test_traiter_eval_sessions(mock, tel):

    formation = Formation(tel["trigramme_formation"])

    formation.ajout_sessions(tel["codes_IRIS"])

    for session in formation.sessions.values():
        session.eval = Mock()

    formation.eval = Mock()

    formation.traiter_eval_sessions(
        chemin_csv=None,
        ouvrirDossier=False
    )

    for session in formation.sessions.values():
        session.eval.traiter_eval.assert_called_once()

    formation.eval.ecritdf_et_sauve_siModif.assert_called_once()


#TODO terster _ajout_FdC_avec_ouverture


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