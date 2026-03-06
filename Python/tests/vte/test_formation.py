from pathlib import Path

import pytest

from vte.domain.formation import Formation
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



# ======================================================================================
# TESTS CLASSE EVALSTAT FORMATION
# ======================================================================================
# Création formation
def test_creation_avec_trigramme(tel):
    formation = Formation(tel["trigramme_formation"])
    assert formation.trigramme_formation == tel["trigramme_formation"]

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
#    - pytest : tous les tests depuis la racine du projet
#    - pytest -v : plus verbeux
#    - ciblé : python -m pytest -v tests/vte/test_formation.py
#    - ciblé avec les print: python -m pytest -s -v tests/vte/test_evalStat.py
#    - ciblé et unitaire : python -m pytest -s -v tests/vte/test_evalStat.py::test_creation_nouvel_eval_formation

# python -m pytest -s -v tests/vte/test_evalStat.py::
# python -m pytest -s -v tests/vte/test_evalStat.py::test_creation_nouvel_eval_formation
# python -m pytest -s -v tests/vte/test_evalStat.py::test_ouverture_eval_formation_existant

#python -m pytest -v tests/vte/test_formation.py