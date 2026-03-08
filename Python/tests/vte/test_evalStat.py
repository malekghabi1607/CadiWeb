import pandas as pd
from unittest.mock import patch

import pytest

from vte.domain.formation import Formation
from vte.domain.evalStat import EvalStat, EvalStat_session, EvalStat_formation

from vte.domain.session import Session
from vte.utils.utils import backup_fichier_test, restore_nom_fichier_test

@pytest.fixture
def csv_datasets(tel):
    """
    Liste de datasets CSV à tester.
    Chaque dataset représente un cas réel rencontré.
    """

    return [

        {
            "nom": "csv_valide_standard",
            "csv": tel["chemin_csv_session"],
            "iris": tel["chemin_IRIS_sessions"],
            "statut": "Traité",
        },

        {
            "nom": "csv_deja_dans_eval_formation",
            "csv": tel["csv_deja_dans_eval_formation"],
            "iris": tel["chemin_IRIS_sessions"],
            "statut": "Exclu - CSV déjà dans fichier global",
        },

        {
            "nom": "code_IRIS_absent_extract",
            "csv": tel["chemin_csv_session_non_present_eval_formation"],
            "iris": tel["iris_sessions_sans_code"],
            "statut": "Exclu - Code IRIS pas dans Extract IRIS sessions",
        },

        {
            "nom": "csv_vide",
            "csv": tel["csv_vide"],
            "iris": tel["chemin_IRIS_sessions"],
            "statut": "Exclu - CSV vide / Aucun retour",
        },

        {
            "nom": "csv_probleme_lecture",
            "csv": tel["csv_probleme_lecture"],
            "iris": tel["chemin_IRIS_sessions"],
            "statut": "Exclu - Problème lecture CSV",
        },

        {
            "nom": "csv_nouveau_format",
            "csv": tel["csv_nouveau_format"],
            "iris": tel["chemin_IRIS_sessions"],
            "statut": "Traité",
        },

        {
            "nom": "csv_ancien_format",
            "csv": tel["csv_ancien_format"],
            "iris": tel["chemin_IRIS_sessions"],
            "statut": "Traité",
        },

    ]



# ======================================================================================
# TESTS CLASSE EVALSTAT
# ======================================================================================

# ----------------------------------------------------------------------
# Test de __init__
# ----------------------------------------------------------------------
# python -m pytest -s -v tests/vte/test_evalStat.py::test_init
def test_init():

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
# python -m pytest -s -v tests/vte/test_evalStat.py::test_avec_traitement_evalStat_session
@patch("vte.domain.evalStat.EvalStat_session.traiter_eval")
def test_avec_traitement_evalStat_session(mock_traiter, session):

    es = EvalStat_session.avec_traitement(session)

    assert isinstance(es, EvalStat_session)

    mock_traiter.assert_called_once()




# ======================================================================================
# TESTS CLASSE EVALSTAT FORMATION
# ======================================================================================
# ----------------------------------------------------------------------
# Test de __init__
# ----------------------------------------------------------------------
# python -m pytest -s -v tests/vte/test_evalStat.py::test_init_evalStat_formation
def test_init_evalStat_formation(formation):

    es = EvalStat_formation(formation)

    assert es.trigramme_formation == formation.trigramme_formation


# ----------------------------------------------------------------------
# Test des constructeurs alternatifs
# ----------------------------------------------------------------------
# python -m pytest -s -v tests/vte/test_evalStat.py::test_avec_ouverture_evalStat_formation
@patch("vte.domain.evalStat.EvalStat_formation._ouvrir_ou_creer_eval_formation")
def test_avec_ouverture_evalStat_formation(mock_ouvrir, formation):

    es = EvalStat_formation.avec_ouverture(formation)

    assert isinstance(es, EvalStat_formation)

    mock_ouvrir.assert_called_once()








# ======================================================================================
# TESTS MÉTIER (PIPELINE COMPLET)
# ======================================================================================
# ----------------------------------------------------------------------
# CREATION / OUVERTURE EVAL FORMATION
# ----------------------------------------------------------------------

# === CREATION NOUVEL EVAL FORMATION ===
# python -m pytest -v tests/vte/test_evalStat.py::test_creation_nouvel_eval_formation
def test_creation_nouvel_eval_formation(tel):

    backup_fichier_test(tel["chemin_eval_formation"])
    assert not tel["chemin_eval_formation"].is_file()

    try:

        formation = Formation.avec_ouverture_evalStat(
            tel["trigramme_formation"]
        )
        #print("\n")
        #print(len(formation.eval.df_stagiaires))

        # On fait nos tests
        assert formation.eval.fe is not None

        assert formation.eval.chemin_fe == tel["chemin_eval_formation"]
        assert len(formation.eval.df_stagiaires) == 0

    except Exception as e:
        raise e

    finally:
        try:
            formation.eval.fe.close()
        finally:
            pass
        restore_nom_fichier_test(tel["chemin_eval_formation"])


# === OUVERTURE EVAL FORMATION EXISTANT ===
# python -m pytest -v tests/vte/test_evalStat.py::test_creation_ouvrir_eval_formation_existant
def test_creation_ouvrir_eval_formation_existant(tel):

    # from vte.utils.utils import backup_fichier_test, rollback_nom_fichier_test
    # from vte.domain.formation import Formation

    backup_fichier_test(tel["chemin_eval_formation"], deplacement=False)
    assert tel["chemin_eval_formation"].is_file()

    try:

        formation = Formation.avec_ouverture_evalStat(
            tel["trigramme_formation"]
        )
        #print("\n")
        #print(len(formation.eval.df_stagiaires))

        assert formation.eval.fe is not None
        assert formation.eval.chemin_fe == tel["chemin_eval_formation"]
        assert len(formation.eval.df_stagiaires) > 0

    except Exception as e:
        raise e

    finally:
        try:
            formation.eval.fe.close()
        finally:
            pass
        restore_nom_fichier_test(tel["chemin_eval_formation"])


# ----------------------------------------------------------------------
# PIPELINE FONCTIONNEL
# ----------------------------------------------------------------------
def test_traitement_eval_session_fonctionnel(tel):
    
   
    # Backups de mon environnement de travail
    backup_fichier_test(tel["chemin_eval_formation"])  # Eval formation
    backup_fichier_test(tel["chemin_eval_session"])  # Eval session

    # TODO : je dois préparer un evalFormation qui ne contient pas la session que je vais traiter

    try:
        

        formation = Formation(tel["trigramme_formation"])
        session = Session.avec_traitement_evalStat(
            formation=formation,
            code_IRIS=tel["code_IRIS"],
            chemin_csv=tel["chemin_csv_session"],
            chemin_IRIS_sessions=tel["chemin_IRIS_sessions"],
            # ecrire_eval_formation=True,  # Valeur par défaut = True
            ouvrirDossier=True
        )

        # Vérification de formation
        assert formation.eval.fe is not None
        assert formation.eval.chemin_fe == tel["chemin_eval_formation"]
        assert len(formation.eval.df_stagiaires) > 0

    finally:
        try:
            formation.eval.fe.close()
        finally:
            pass
        
        # Restauration de mon environnement de travail
        restore_nom_fichier_test(tel["chemin_eval_formation"])
        restore_nom_fichier_test(tel["chemin_eval_session"])




# si chemin_csv is None, alors self._filedialog_csv appelé


# ----------------------------------------------------------------------
# Création avec eval formation existant ne contenant pas le code IRIS de la session
# (on peut traiter l'EvalStat session)
# ----------------------------------------------------------------------
# TODO : peut-être à finir
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
        #print("\n")
        #print(len(formation.eval.df_stagiaires))
        
        assert formation.eval.fe is not None
        if formation.eval.fe is not None :
            assert (formation.eval.chemin_fe == tel_csv_existant["chemin_eval_formation"])
            assert len(formation.eval.df_stagiaires) > 0
        assert formation.sessions[tel_csv_existant["code_IRIS"]].code_IRIS == tel_csv_existant["code_IRIS"]
        
    finally:
        restore_nom_fichier_test(tel_csv_existant["chemin_eval_formation"])



#ChatGPT - Probablement inclure tests traitement EvalStat session complet, et vérifier retours evalstat session + evalstat formation
def test_maj_eval_formation_apres_traitement(
    formation_tel_complete,
    tel
):

    formation_tel_complete.ajout_session_avec_evalStat(
        code_IRIS=tel["code_IRIS"],
        chemin_csv=tel["chemin_csv_session"],
        chemin_IRIS_sessions=tel["chemin_IRIS_sessions"]
    )

    eval_formation = formation_tel_complete.eval

    assert len(eval_formation.df) == tel["nb_lignes_eval_formation_apres_maj"]

    assert eval_formation.df.iloc[-1,0] == tel["dernier_code_session_eval_formation"]

#ChatGPT
def test_generation_excel_eval_session(
    formation_tel_complete,
    tel
):

    s = formation_tel_complete.ajout_session_avec_evalStat(
        code_IRIS=tel["code_IRIS"],
        chemin_csv=tel["chemin_csv_session"],
        chemin_IRIS_sessions=tel["chemin_IRIS_sessions"]
    )

    evalstat = s.eval

    assert evalstat.chemin_excel == tel["chemin_eval_session"]


# À partir de là c'est ChatGPT
# ======================================================================================
# TESTS DES CAS CSV QUI EXCLUENT LE TRAITEMENT EVALSTAT SESSION
# ======================================================================================
@pytest.mark.parametrize("dataset_index", range(7))
def test_pipeline_csv(
    formation_tel_complete,
    tel,
    csv_datasets,
    dataset_index
):

    dataset = csv_datasets[dataset_index]

    s = formation_tel_complete.ajout_session_avec_evalStat(
        code_IRIS=tel["code_IRIS"],
        chemin_csv=dataset["csv"],
        chemin_IRIS_sessions=dataset["iris"]
    )

    evalstat = s.eval

    assert evalstat._statut_csv == dataset["statut"]




def test_csv_deja_present_dans_eval_formation(
    formation_tel_complete,
    tel
):

    s = formation_tel_complete.ajout_session_avec_evalStat(
        code_IRIS=tel["code_IRIS"],
        chemin_csv=tel["csv_deja_dans_eval_formation"],
        chemin_IRIS_sessions=tel["chemin_IRIS_sessions"]
    )

    evalstat = s.eval

    assert evalstat._statut_csv == "Exclu - CSV déjà dans fichier global"
    assert evalstat._fe is None

def test_code_IRIS_absent_extract_sessions(
    formation_tel_complete,
    tel
):

    s = formation_tel_complete.ajout_session_avec_evalStat(
        code_IRIS=tel["code_IRIS"],
        chemin_csv=tel["chemin_csv_session_non_present_eval_formation"],
        chemin_IRIS_sessions=tel["iris_sessions_sans_code"]
    )

    evalstat = s.eval

    assert evalstat._statut_csv == "Exclu - Code IRIS pas dans Extract IRIS sessions"
    assert evalstat._fe is None

def test_csv_probleme_lecture(
    formation_tel_complete,
    tel
):

    s = formation_tel_complete.ajout_session_avec_evalStat(
        code_IRIS=tel["code_IRIS"],
        chemin_csv=tel["csv_probleme_lecture"],
        chemin_IRIS_sessions=tel["chemin_IRIS_sessions"]
    )

    evalstat = s.eval

    assert evalstat._statut_csv == "Exclu - Problème lecture CSV"
    assert evalstat._fe is None

def test_csv_vide(
    formation_tel_complete,
    tel
):

    s = formation_tel_complete.ajout_session_avec_evalStat(
        code_IRIS=tel["code_IRIS"],
        chemin_csv=tel["csv_vide"],
        chemin_IRIS_sessions=tel["chemin_IRIS_sessions"]
    )

    evalstat = s.eval

    assert evalstat._statut_csv == "Exclu - CSV vide / Aucun retour"
    assert evalstat._fe is None


# ======================================================================================
# TESTS DES VERSIONS DES CSV
# ======================================================================================
# Cas 1 déjà traité normalement
def test_csv_nouveau_format(
    formation_tel_complete,
    tel
):

    s = formation_tel_complete.ajout_session_avec_evalStat(
        code_IRIS=tel["code_IRIS"],
        chemin_csv=tel["csv_nouveau_format"],
        chemin_IRIS_sessions=tel["chemin_IRIS_sessions"]
    )

    evalstat = s.eval

    assert evalstat._statut_csv == "Traité"
    assert evalstat._fe is not None

# Cas 2 – Ancien format
def test_csv_ancien_format(
    formation_tel_complete,
    tel
):

    s = formation_tel_complete.ajout_session_avec_evalStat(
        code_IRIS=tel["code_IRIS"],
        chemin_csv=tel["csv_ancien_format"],
        chemin_IRIS_sessions=tel["chemin_IRIS_sessions"]
    )

    evalstat = s.eval

    assert evalstat._statut_csv == "Traité"
    assert evalstat._fe is not None


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