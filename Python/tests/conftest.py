from pathlib import Path
from typing import Any
import pytest

from vte.domain.formation import Formation
from vte.utils.utils import chemin_vers_unc


# ======================================================================================
# FIXTURES COMMUNES
# ======================================================================================
def tel_data() -> dict[str, Any]:
    return {
        # ----------------------------
        # PROPRIETES FONCTIONNELLES
        # ----------------------------
        # Propriétés de la formation
        "trigramme_formation": "TEL",
        "chemin_eval_formation": chemin_vers_unc(Path(r"P:\FORMATIONS_C\TEL\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\Evaluation-Stagiaires-Global-TEL.xlsx")),  # chemin Excel Eval formation existant
        "chemin_eval_formation_sans16411_sans12766_sans11090": chemin_vers_unc(Path(r"C:\Users\vt238770\Documents\_CEA\Prog\Python\tests\vte\data\Evaluation-Stagiaires-Global-TEL-Sans16411_Sans12766_Sans11090.xlsx")),  # chemin Excel Eval formation ne contenant pas la session 16411 (nouveau format de CSV) ni la 12766 (ancien format de CSV) ni 11090 (fichier CSV avec en-tête mais sans données)
        "chemin_eval_formation_avec16411": chemin_vers_unc(Path(r"C:\Users\vt238770\Documents\_CEA\Prog\Python\tests\vte\data\Evaluation-Stagiaires-Global-TEL.xlsx")),  # chemin Excel Eval formation contenant la session 16411 (et toutes les autres)

        # Propriétés de la session
        "code_IRIS": 16411,

        "chemin_csv_session": chemin_vers_unc(Path(r"P:\FORMATIONS_C\TEL\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\S-16411-rapports-session-evaluations\S-16411-FC25-TEL-VTE-CAR-Stagiaires.csv")),  # CSV session valide
        "chemin_eval_session": chemin_vers_unc(Path(r"P:\FORMATIONS_C\TEL\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\S-16411-rapports-session-evaluations\S-16411-FC25-TEL-VTE-CAR-Stagiaires.xlsx")),  # Excel EvalStat session attendu
        "chemin_bilanSessions": chemin_vers_unc(Path(r"P:\FORMATIONS_C\TEL\P07-bilan-sessions-et-bilan-formation\2025\P07-Pr05-F05-Bilan session-S-16411-FC25-TEL-VTE-CAR (octobre 2025)-UEM.docx")),
        
        # Propriétés transverses
        "chemin_IRIS_sessions": chemin_vers_unc(Path(r"R:\_Echanges\VTE\Prog\IRIS\Extracts complets\TESTS - TEL - R04110_Sessions-COMPLET.xlsx")),  # Extract IRIS sessions contenant le code IRIS        



        # ----------------------------
        # PROPRIETES PLUSIEURS SESSIONS 
        # ----------------------------         
        "codes_IRIS": (16161, 16411),  # Ce sont les 2 des formats CSV "nouveaux"


        # ----------------------------
        # RESULTATS
        # ----------------------------      
        # ---- Si le traitement réussit ----
        # # - Résultats communs eval session et formation  
        "resultat_apresTraitement_derniereLigne_NOMPrenom": "CAMBE Justine",
        "resultat_apresTraitement_derniereLigne_Critere": "Avez-vous d'autres besoins de formation ?",
        "resultat_apresTraitement_derniereLigne_Note": 0,

        # - Résultats communs eval session et formation
        "resultat_apresTraitement_evalSession_nbLignes": 72,
        "resultat_apresTraitement_evalFormation_nbLignes": 174, # 229-55,

       
        # ---- Si le traitement échoue ----
        # - Résultats communs eval session et formation
        "resultat_sansTraitement_evalFormation_nbLignes": 102, # 157-55,
        "resultat_sansTraitement_derniereLigne_NOMPrenom": "PASCAL Thomas",
        "resultat_sansTraitement_derniereLigne_Critere": "Comment avez-vous connu cette formation ?",
        "resultat_sansTraitement_derniereLigne_Commentaires": "Service formation",
        

    }


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
    return tel_data()

@pytest.fixture
def formation(tel):
    """Formation réelle avec juste trigramme formation"""
    return Formation(tel["trigramme_formation"])

@pytest.fixture
def session(formation, tel):
    """Session réelle attachée à la formation"""
    formation.ajout_sessions(codes_IRIS=(tel["code_IRIS"],))
    return formation.get_session_par_codeIRIS(tel["code_IRIS"])

"""
@pytest.fixture
def formation(tel):
    return Formation.avec_ouverture_evalStat(tel["trigramme_formation"])
"""
