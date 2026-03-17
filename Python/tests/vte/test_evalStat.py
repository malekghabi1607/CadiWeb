from pathlib import Path
import shutil

from unittest.mock import patch

import pytest

from vte.core.iris_referentiel import set_iris_chemin_specifique
from vte.domain.formation import Formation
from vte.domain.session import Session
from vte.domain.evalStat import EvalStat, EvalStat_session, EvalStat_formation

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

    assert es.fe is None
    assert es._session == session
    assert es._chemin_csv is None
    assert es._statut is None


# ----------------------------------------------------------------------
# Test des constructeurs alternatifs
# ----------------------------------------------------------------------
def test_avec_ouverture_evalStat_session(session, tel):
    # TODO : je ne teste pas avec ouvrir_fe=True
    es = EvalStat_session.avec_ouverture(session=session, ouvrir_fe=False)

    assert es.fe is None
    assert es._session == session
    assert es._chemin_csv == tel["chemin_csv_session"]
    assert es._statut == "Traité"


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
"""
Permet de tester toute la procédure métier de création d'un evalStat :
    - chemin_csv :
        ¤ chemin_csv = bon chemin → On continue (défaut) [test_traitement_eval_session_fonctionnel]
        ¤ chemin_csv is None → Ouverture filedialog → self._filedialog_csv(trigramme_formation=self.trigramme_formation) is called

    - Excel eval formation déjà existant :
        ¤ Oui → On ouvre l'eval formation (défaut) [test_creation_ouvrir_eval_formation_existant]
        ¤ Non → On crée l'éval formation [test_creation_nouvel_eval_formation]

    - chemin_csv :
        ¤ non présent dans eval formation → on continue (défaut) [test_traitement_eval_session_fonctionnel]
        ¤ déjà dans eval formation → self._statut_csv = "Exclu - CSV déjà dans fichier global" [test_eval_session_multi_csv]
        ¤ a un problème lors de la lecture (ex. : pas un vrai .csv) → self._statut_csv = "Exclu - Problème lecture CSV" [test_eval_session_multi_csv]
        ¤ est vide (présent mais aucune ligne de données) → self._statut_csv = "Exclu - CSV vide / Aucun retour" [test_eval_session_multi_csv]
        ¤ CSV ancien format → Traitement va au bout [test_eval_session_multi_csv]

    - code_IRIS :
        ¤ présent dans IRIS sessions → On continue (défaut) [test_traitement_eval_session_fonctionnel]
        ¤ non présent dans IRIS sessions → self._statut_csv = "Exclu - Code IRIS pas dans Extract IRIS sessions" [test_eval_session_multi_csv]

Fin si ok : 
    - self._statut_csv = "Traité"
    - excel eval session créé avec bonnes valeurs
    - excel eval formation créé/màj avec bonnes valeurs

Fonction testées :
    - Eval_session._filedialog_csv
    - Eval_formation._ouvrir_ou_creer_eval_formation
    - Eval_session._charger_csv_stagiaire
    - Eval_session._traiter_df_csv
    - Eval_session._traiter_df_stagiaires
    - Eval_session._ecrit_df_et_sauve
    - Eval_session._maj_dataframes_eval_formation
    - Eval_formation.ecritdf_et_sauve_siModif
"""
# ----------------------------------------------------------------------
# CREATION / OUVERTURE EVAL FORMATION
# ----------------------------------------------------------------------

# === CREATION NOUVEL EVAL FORMATION ===
# python -m pytest -v tests/vte/test_evalStat.py::test_creation_nouvel_eval_formation
def test_creation_nouvel_eval_formation(tel):
    """
    Création nouvel evalStat formation (emploi de avec_ouverture_evalStat avec fichier Excel eval formation inexistant)
    """

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
    """
    Ouverture nouvel evalStat formation (emploi de avec_ouverture_evalStat avec fichier Excel eval formation inexistant)
    """
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
# python -m pytest -vv -s tests/vte/test_evalStat.py::test_traitement_eval_session_fonctionnel

@patch("vte.domain.evalStat.FichierExcel.actualiser_TCD")
@patch("vte.domain.evalStat.FichierExcel._TableauExcel.copierFormat_tableauStructure_xlwings")
@patch("vte.domain.evalStat.FichierExcel._TableauExcel.maj_references_misesEnFormeConditionnelles")
def test_traitement_eval_session_fonctionnel(
    mock_mfc,
    mock_copier,
    mock_tcd,
    tel
    ):
    """
    Cas par défaut
    Permet de tester toute la procédure métier de création d'un evalStat :
        - chemin_csv :
            ¤ chemin_csv = bon chemin → On continue (défaut)

        - Excel eval formation déjà existant :
            ¤ Oui → On ouvre l'eval formation (défaut)

        - chemin_csv :
            ¤ non présent dans eval formation → on continue (défaut)

        - code_IRIS :
            ¤ présent dans IRIS sessions → On continue (défaut)

    Fin si ok : 
        - self._statut_csv = "Traité"
        - excel eval session créé avec bonnes valeurs
        - excel eval formation créé/màj avec bonnes valeurs
    """    
   
    # Backups de mon environnement de travail
    backup_fichier_test(tel["chemin_eval_formation"])  # Eval formation
    backup_fichier_test(tel["chemin_eval_session"])  # Eval session


    # On copie l'eval formation de test 
    shutil.copy(str(tel["chemin_eval_formation_sans16411_sans12766_sans11090"]), str(tel["chemin_eval_formation"]))

    try:
        set_iris_chemin_specifique(typeExport="Sessions", chemin=tel["chemin_IRIS_sessions"])  # On redéfinit le chemin d'IRIS Sessions
        formation = Formation.avec_ouverture_evalStat(tel["trigramme_formation"])
        session = Session.avec_traitement_evalStat(
            formation=formation,
            code_IRIS=tel["code_IRIS"],
            chemin_csv=tel["chemin_csv_session"],
            # ecrire_eval_formation=True,  # Valeur par défaut = True
            ouvrirDossier=True
        )
        # Vérifications de session
        assert session.eval.fe is not None
        assert session.eval.statut == "Traité"
        assert session.eval.chemin_fe == tel["chemin_eval_session"]
        assert len(session.eval.df_stagiaires) == tel["resultat_apresTraitement_evalSession_nbLignes"]

        assert Path(session.eval.df_stagiaires.iloc[-1]["Chemin fichier CSV"]) == tel["chemin_csv_session"]
        assert formation.eval.df_stagiaires.iloc[-1]["NOM Prénom"] == tel["resultat_apresTraitement_derniereLigne_NOMPrenom"]
        assert formation.eval.df_stagiaires.iloc[-1]["Critère"] == tel["resultat_apresTraitement_derniereLigne_Critere"]
        assert formation.eval.df_stagiaires.iloc[-1]["Note"] == tel["resultat_apresTraitement_derniereLigne_Note"]


        # Vérification de formation
        assert formation.eval.fe is not None
        assert formation.eval.chemin_fe == tel["chemin_eval_formation"]
        assert len(formation.eval.df_stagiaires) == tel["resultat_apresTraitement_evalFormation_nbLignes"]

        assert Path(formation.eval.df_stagiaires.iloc[-1]["Chemin fichier CSV"]) == tel["chemin_csv_session"]
        assert formation.eval.df_stagiaires.iloc[-1]["NOM Prénom"] == tel["resultat_apresTraitement_derniereLigne_NOMPrenom"]
        assert formation.eval.df_stagiaires.iloc[-1]["Critère"] == tel["resultat_apresTraitement_derniereLigne_Critere"]
        assert formation.eval.df_stagiaires.iloc[-1]["Note"] == tel["resultat_apresTraitement_derniereLigne_Note"]

    finally:
        try:
            formation.eval.fe.close()
        finally:
            pass
        
        # Restauration de mon environnement de travail
        restore_nom_fichier_test(tel["chemin_eval_formation"])
        restore_nom_fichier_test(tel["chemin_eval_session"])







# ======================================================================================
# TESTS DES CAS CSV QUI EXCLUENT LE TRAITEMENT EVALSTAT SESSION
# ======================================================================================
# python -m pytest -v tests/vte/test_evalStat.py::test_eval_session_multi_csv
@pytest.mark.parametrize("dataset_index", range(5))
@patch("vte.domain.evalStat.FichierExcel.actualiser_TCD")
@patch("vte.domain.evalStat.FichierExcel._TableauExcel.copierFormat_tableauStructure_xlwings")
@patch("vte.domain.evalStat.FichierExcel._TableauExcel.maj_references_misesEnFormeConditionnelles")
def test_eval_session_multi_csv(
    mock_mfc,
    mock_copier,
    mock_tcd,
    formation,
    tel,
    evalstat_csv_datasets,
    dataset_index
):
    # Backups de mon environnement de travail
    backup_fichier_test(tel["chemin_eval_formation"], deplacement=False)  # Eval formation
    backup_fichier_test(tel["chemin_eval_session"])  # Eval session

    # On copie l'eval formation de test 
    shutil.copy(str(tel["chemin_eval_formation_sans16411_sans12766_sans11090"]), str(tel["chemin_eval_formation"]))


    try:
        dataset = evalstat_csv_datasets[dataset_index]
        set_iris_chemin_specifique(typeExport="Sessions", chemin=tel["chemin_IRIS_sessions"])  # On redéfinit le chemin d'IRIS Sessions

        formation = Formation.avec_ouverture_evalStat(tel["trigramme_formation"])
        session = Session.avec_traitement_evalStat(
            formation=formation,
            code_IRIS=dataset["code_IRIS"],
            chemin_csv=dataset["csv"],
            # ecrire_eval_formation=True,  # Valeur par défaut = True
            ouvrirDossier=False
        )




        # Vérifications de session
        if dataset["statut"] == "Traité":
            assert session.eval.fe is not None
        else:
            assert session.eval.fe is None
        assert session.eval.statut == dataset["statut"]


        # Vérification de formation
        assert formation.eval.fe is not None
        assert formation.eval.chemin_fe == tel["chemin_eval_formation"]
        assert len(formation.eval.df_stagiaires) == dataset["nbLignes_evalFormation"]

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
# Le code IRIS n'est pas présent dans IRIS Sessions → DEJA DANS MULTI
# ----------------------------------------------------------------------
# python -m pytest -v tests/vte/test_evalStat.py::test_eval_session_codeIRIS_inexistant_dans_IRISsessions
"""
def test_eval_session_codeIRIS_inexistant_dans_IRISsessions(tel):

    # Backups de mon environnement de travail
    backup_fichier_test(tel["chemin_eval_formation"], deplacement=False)  # Eval formation
    backup_fichier_test(tel["chemin_eval_session"])  # Eval session

    # On copie l'eval formation de test 
    shutil.copy(str(tel["chemin_eval_formation_sans16411_sans12766_sans11090"]), str(tel["chemin_eval_formation"]))


    try:
        formation = Formation.avec_ouverture_evalStat(tel["trigramme_formation"])
        session = Session.avec_traitement_evalStat(
            formation=formation,
            code_IRIS=tel["code_IRIS_non_present"],
            chemin_csv=tel["chemin_csv_session"],
            chemin_IRIS_sessions=tel["chemin_IRIS_sessions"],
            # ecrire_eval_formation=True,  # Valeur par défaut = True
            ouvrirDossier=True
        )
        # Vérifications de session
        assert session.eval.fe is None
        assert session.eval.statut_csv == "Exclu - Code IRIS pas dans Extract IRIS sessions"


        # Vérification de formation
        assert formation.eval.fe is not None
        assert formation.eval.chemin_fe == tel["chemin_eval_formation"]
        assert len(formation.eval.df_stagiaires) == tel["resultat_sansTraitement_evalFormation_nbLignes"]

        assert formation.eval.df_stagiaires.iloc[-1]["NOM Prénom"] == tel["resultat_sansTraitement_derniereLigne_NOMPrenom"]
        assert formation.eval.df_stagiaires.iloc[-1]["Critère"] == tel["resultat_sansTraitement_derniereLigne_Critere"]
        assert formation.eval.df_stagiaires.iloc[-1]["Commentaires"] == tel["resultat_sansTraitement_derniereLigne_Commentaires"]

    finally:
        try:
            formation.eval.fe.close()
        finally:
            pass
        
        # Restauration de mon environnement de travail
        restore_nom_fichier_test(tel["chemin_eval_formation"])
        restore_nom_fichier_test(tel["chemin_eval_session"])
"""






# Lancer les tests :
#    - pytest : tous les tests depuis la racine du projet
#    - pytest -v : plus verbeux
#    - ciblé : python -m pytest -v tests/vte/test_iris.py
#    - ciblé avec les print: python -m pytest -s -v tests/vte/test_evalStat.py
#    - ciblé et unitaire : python -m pytest -s -v tests/vte/test_evalStat.py::test_creation_nouvel_eval_formation

# python -m pytest -s -v tests/vte/test_evalStat.py::
# python -m pytest -s -v tests/vte/test_evalStat.py::test_creation_nouvel_eval_formation
# python -m pytest -s -v tests/vte/test_evalStat.py::test_ouverture_eval_formation_existant