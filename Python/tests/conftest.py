from pathlib import Path
import pytest

from vte.domain.formation import Formation
from vte.utils.utils import chemin_vers_unc


# ======================================================================================
# FIXTURES COMMUNES
# ======================================================================================
r"""
@pytest.fixture
def tel():
    return {
        "chemin_IRIS_sessions": chemin_vers_unc(Path(r"R:\_Echanges\VTE\Prog\IRIS\Extracts complets\TESTS - TEL - R04110_Sessions-COMPLET.xlsx")),

        "trigramme_formation": "TEL",
        "code_IRIS": 16411,
        "codes_IRIS": (12766, 13414, 15942, 16161, 16411),

        "chemin_eval_formation": chemin_vers_unc(Path(r"P:\FORMATIONS_C\TEL\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\Evaluation-Stagiaires-Global-TEL.xlsx")),
        "chemin_csv_session": chemin_vers_unc(Path(r"P:\FORMATIONS_C\TEL\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\S-16411-rapports-session-evaluations\S-16411-FC25-TEL-VTE-CAR-Stagiaires.csv")),
        "chemin_eval_session": chemin_vers_unc(Path(r"P:\FORMATIONS_C\TEL\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\S-16411-rapports-session-evaluations\S-16411-FC25-TEL-VTE-CAR-Stagiaires.xlsx")),

        "chemin_csv_session_non_present_eval_formation": chemin_vers_unc(Path(r"P:\FORMATIONS_C\948\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2022-11-S-12995 UEM\S-12995-FC22-948-VTE-SNA-Stagiaires.csv")),

        "resultat_1er_elem_eval_formation": "S-12766-FC22-TEL-JVI-LRA",
    }
"""
@pytest.fixture
def tel():
    return {
        # ----------------------------
        # PROPRIETES FONCTIONNELLES
        # ----------------------------
        # Propriétés de la formation
        "trigramme_formation": "TEL",
        "chemin_eval_formation": chemin_vers_unc(Path(r"P:\FORMATIONS_C\TEL\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\Evaluation-Stagiaires-Global-TEL.xlsx")),  # chemin Excel Eval formation existant

        # Propriétés de la session
        "code_IRIS": 16411,

        "chemin_csv_session": chemin_vers_unc(Path(r"P:\FORMATIONS_C\TEL\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\S-16411-rapports-session-evaluations\S-16411-FC25-TEL-VTE-CAR-Stagiaires.csv")),  # CSV session valide
        "chemin_eval_session": chemin_vers_unc(Path(r"P:\FORMATIONS_C\TEL\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\S-16411-rapports-session-evaluations\S-16411-FC25-TEL-VTE-CAR-Stagiaires.xlsx")),  # Excel EvalStat session attendu
        
        # Propriétés transverses
        "chemin_IRIS_sessions": chemin_vers_unc(Path(r"R:\_Echanges\VTE\Prog\IRIS\Extracts complets\TESTS - TEL - R04110_Sessions-COMPLET.xlsx")),  # Extract IRIS sessions contenant le code IRIS        


        # ----------------------------
        # PROPRIETES PLUSIEURS SESSIONS 
        # ----------------------------         
        "codes_IRIS": (12766, 13414, 15942, 16161, 16411),


        # ----------------------------
        # PROPRIETES AVEC 
        # ----------------------------        


        "resultat_1er_elem_eval_formation": "S-12766-FC22-TEL-JVI-LRA",  # première cellule attendue après traitement


        # ----------------------------
        # CAS PROBLEMES TRAITEMENT EVALSTAT
        # ----------------------------
        # CSV dont la session n'est pas dans l'eval formation
        "chemin_csv_session_non_present_eval_formation": chemin_vers_unc(Path(r"P:\FORMATIONS_C\948\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2022-11-S-12995 UEM\S-12995-FC22-948-VTE-SNA-Stagiaires.csv")), 
        
        # CSV dont la référence est déjà présente dans l'eval formation
        "csv_deja_dans_eval_formation": "",
        
        # extract IRIS sessions ne contenant PAS le code IRIS
        "iris_sessions_sans_code": "",
        
        # CSV volontairement corrompu
        "csv_probleme_lecture": "",
        
        # CSV stagiaires vide
        "csv_vide": "",
        


        # ----------------------------
        # CAS FORMAT CSV
        # ----------------------------

        "csv_nouveau_format": "",
        # CSV nouveau format contenant colonne "Date de fin"

        "csv_ancien_format": "",
        # CSV ancien format avec 2 colonnes inutiles


        # ----------------------------
        # RESULTATS ATTENDUS
        # ----------------------------

        "nb_lignes_eval_formation_apres_maj": 0,
        # nombre de lignes attendu dans eval formation après ajout session

        "dernier_code_session_eval_formation": "",
        # code session attendu dernière ligne

    }

@pytest.fixture
def formation(tel):
    """Formation réelle avec juste trigramme formation"""
    return Formation(tel["trigramme_formation"])

@pytest.fixture
def session(formation, tel):
    """Session réelle attachée à la formation"""
    formation.ajout_sessions(codes_IRIS=(tel["code_IRIS"],))
    return formation.sessions[tel["code_IRIS"]]

"""
@pytest.fixture
def formation(tel):
    return Formation.avec_ouverture_evalStat(tel["trigramme_formation"])
"""
