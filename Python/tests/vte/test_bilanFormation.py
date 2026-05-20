import pytest
import sys
from pathlib import Path
from pprint import *

from vte.core.iris_referentiel import set_iris_chemin_specifique
from vte.domain.evalStat import EvalStat_session
from vte.domain.formation import Formation
from vte.domain.session import Session
from vte.domain.bilanFormation import BilanFormation
from vte.domain.iris import IRIS_traite, IRIS_sessions, IRIS_ventes
from vte.utils.utils import backup_fichier_test, chemin_vers_unc, restore_nom_fichier_test

sys.path.append(str(Path(__file__).resolve().parents[1]))
from conftest import tel_data

tel = tel_data()

# Guide mock
"""
Cas	Où mocker
from X import Y	module_qui_utilise.Y
import X	module_qui_utilise.X.Y
méthode de classe	module.Classe.methode
"""

# Mettre le dataset comme ça est une astuce pour pouvoir appeler tous les cas dans le parametrize.
"""
En effet, on ne peut pas accéder à une fixture dans le décorateur @pytest...
En déclarant avant, on a l'info pour faire le .keys() et on l'emploie pour définir le fixture
"""


# TODO : les bilans sortent avec les mêmes stats sur année n et n-1. A priori c'est à cause du mock mock_demande_sessions_a_exclure car je lui dis les codes à employer et ce sont les mêmes pour les années n et n-1 alors que ces codes devraient être différents

#2024 → Stats en 2024
#2024 → Pas de stats en 2023
DATASETS_ANNEES = {
    "Bilan 2025": {
        "chemin_IRIS_sessions": tel["chemin_IRIS_sessions"],  # Chemin spécifique IRIS sessions pour tests
        "trigramme_formation": tel["trigramme_formation"],
        "annee": 2025,  # Année du bilan de formation à générer
        #"codes_IRIS": [17343, 16411],  # Liste des codes IRIS à inclure dans le bilan
        "chemin_bilanFormation": Path(chemin_vers_unc(r"P:\FORMATIONS_C\TEL\P07-bilan-sessions-et-bilan-formation\2025\P07-Pr05-F06-Bilan formation-Année 2025.docx")),  #  Chemin du bilan de formation à Backup au cas où

        # mocks
        "mock_verifier_existance_fichier": True,  # True = on continue
        "mock_demande_sessions_a_exclure": [17343, 16411],  # Codes qui seront finalement retenus pour le bilan de formation
        #"mock_statut_eval": "Traité",

        # attendu
        #"expected_statut": "OK",


        },
    
    "Bilan 2024": {
        "chemin_IRIS_sessions": tel["chemin_IRIS_sessions"],  # Chemin spécifique IRIS sessions pour tests
        "trigramme_formation": tel["trigramme_formation"],
        "annee": 2024,  # Année du bilan de formation à générer
        #"codes_IRIS": [16934, 16298, 16161, 15942, 15830, 15610],  # Liste des codes IRIS à inclure dans le bilan
        "chemin_bilanFormation": Path(chemin_vers_unc(r"P:\FORMATIONS_C\TEL\P07-bilan-sessions-et-bilan-formation\2024\P07-Pr05-F05-Bilan session-Année 2024-UEM.docx")),  #  Chemin du bilan de formation à Backup au cas où

        # mocks
        "mock_verifier_existance_fichier": True,  # True = on continue
        "mock_demande_sessions_a_exclure": [16934, 16298, 16161, 15942, 15830, 15610],  # Codes qui seront finalement retenus pour le bilan de formation
        #"mock_statut_eval": "Traité",

        # attendu
        #"expected_statut": "OK",


        },
        
}





DATASETS_PERIODE = {
    "Bilan TEL annee 2025 (1 sessions, 1 CSV) ; année 2024 (3 sessions, 2 CSV seulement)": {
        "chemin_IRIS_sessions": tel["chemin_IRIS_sessions"],
        "trigramme_formation": tel["trigramme_formation"],
        "annee": 2025,
        "periode": "Année",
        "chemin_bilanFormation": chemin_vers_unc(Path(r"P:\FORMATIONS_C\TEL\P07-bilan-sessions-et-bilan-formation\2025\P07-Pr05-F06-Bilan-formation-TEL-2025.docx")),
        
        # mocks
        "mock_verifier_existance_fichier": True,  # True = on continue
        "mock_demande_sessions_a_retenir": [15830, 15942, 16161],
        #"mock_statut_eval": "Exclu - Aucun CSV fourni",

        # attendu
        #"expected_statut": "Exclu - Aucun CSV fourni",
        },

}

STATUTS_EVAL_NON_TRAITE = {
    # Année 2024
    16934 : "Exclu - Aucun CSV fourni",
    16298 : "Exclu - Aucun CSV fourni",
    # 16161 : "Traité",  # → Pas besoin de le stipuler ici. Ici, uniquement les non-traités
    # 15942 : "Traité",  # → Pas besoin de le stipuler ici. Ici, uniquement les non-traités
    15830 : "Exclu - Aucun CSV fourni",
    15610 : "Exclu - Aucun CSV fourni",
    
    # Année 2025
    17343 : "Exclu - Aucun CSV fourni",  # Bug QR code
    # 16411 : "Traité",  # → Pas besoin de le stipuler ici. Ici, uniquement les non-traités
}


@pytest.fixture
def datasets_codesIRIS():
    """
    Liste de datasets CSV à tester.
    Chaque dataset représente un cas réel rencontré.
    Pour utilisation avec depuis_codesIRIS
    """
    return DATASETS_CODES_IRIS

@pytest.fixture
def datasets_periode():
    """
    Liste de datasets CSV à tester.
    Chaque dataset représente un cas réel rencontré.
    Pour utilisation avec depuis_periode
    """
    return DATASETS_PERIODE



def apply_mocks(monkeypatch, dataset):
    # --- mock fichier ---
    # Vérification que le bilan existe déjà → On veut que le retour (continuer) = True
    monkeypatch.setattr(
        "vte.domain.bilanFormation.verifier_existance_fichier",  # Fonction à remplacer
        lambda path: dataset["mock_verifier_existance_fichier"]  # Fonction de remplacement
    )



    # --- mock EvalStat ---
    # Retour que fait EvalStat_session.avec_ouverture_ou_traitement dans EvalStat_session.statut → Il faut lister en amont les codes IRIS qui n'ont pas de QRCode pour ne pas avoir les popup de recherche des CSV
    def fake_evalstat(*args, **kwargs):
        session:Session = kwargs.get("session") or args[0]
        return EvalStat_session(
            session=session,
            statut=STATUTS_EVAL_NON_TRAITE[session.code_IRIS] if session.code_IRIS in STATUTS_EVAL_NON_TRAITE.keys() else "Traité"
        )

    monkeypatch.setattr(
        "vte.domain.evalStat.EvalStat_session.avec_ouverture_ou_traitement",
        fake_evalstat
    )

    # --- mock demande_sessions_a_retenir ---
    def fake_demande_sessions_a_exclure(*args, **kwargs):
        return dataset["mock_demande_sessions_a_exclure"]

    monkeypatch.setattr(
        IRIS_sessions,
        "demande_sessions_a_exclure",
        fake_demande_sessions_a_exclure
    )



# Pour choisir mes cas
"""
@pytest.mark.parametrize("nom_cas", [
    "Bilan unique par code IRIS. CSV dans eval formation",
    "Bilan unique par code IRIS. CSV introuvable",
])
"""



@pytest.mark.parametrize("nom_cas, dataset", DATASETS_ANNEES.items())
def test_bilanFormation_depuis_codesIRIS(nom_cas, dataset, monkeypatch):
    # MODE PYTEST
    apply_mocks(monkeypatch, dataset)  # Appliquer les mocks
    
    # Backups de mon environnement de travail
    backup_fichier_test(dataset["chemin_bilanFormation"])  # Bilan sessions word
    backup_fichier_test(dataset["chemin_bilanFormation"].with_suffix(".pdf"))  # Bilan sessions pdf

    set_iris_chemin_specifique(typeExport="Sessions", chemin=dataset["chemin_IRIS_sessions"])  # On redéfinit le chemin d'IRIS Sessions

    formation = Formation.pour_traitement_bilanFormation_depuis_annee(
        trigramme_formation = dataset["trigramme_formation"],
        annee = dataset["annee"]
        )
    
    # Restauration de mon environnement de travail
    restore_nom_fichier_test(dataset["chemin_bilanFormation"])  # Bilan sessions word
    restore_nom_fichier_test(dataset["chemin_bilanFormation"].with_suffix(".pdf"))  # Bilan sessions pdf


