import pytest
import sys
from pathlib import Path
from pprint import *

from vte.core.iris_referentiel import set_iris_chemin_specifique
from vte.domain.formation import Formation
from vte.domain.bilanSessions import BilanSessions
from vte.utils.utils import backup_fichier_test, chemin_vers_unc, restore_nom_fichier_test

sys.path.append(str(Path(__file__).resolve().parents[1]))
from conftest import tel_data

tel = tel_data()

# Mettre le dataset comme ça est une astuce pour pouvoir appeler tous les cas dans le parametrize.
"""
En effet, on ne peut pas accéder à une fixture dans le décorateur @pytest...
En déclarant avant, on a l'info pour faire le .keys() et on l'emploie pour définir le fixture
"""
DATASETS_CODES_IRIS = {
        "Bilan unique par code IRIS. CSV dans eval formation": {
            "chemin_IRIS_sessions": tel["chemin_IRIS_sessions"],
            "trigramme_formation": tel["trigramme_formation"],
            "codes_IRIS": tel["code_IRIS"],
            "chemin_bilanSessions": tel["chemin_bilanSessions"]
            },


        "Bilan unique par code IRIS. CSV introuvable": {
            "chemin_IRIS_sessions": tel["chemin_IRIS_sessions"],
            "trigramme_formation": tel["trigramme_formation"],
            "codes_IRIS": 17343,
            "chemin_bilanSessions": chemin_vers_unc(Path(r"P:\FORMATIONS_C\TEL\P07-bilan-sessions-et-bilan-formation\2025\P07-Pr05-F05-Bilan session-S-17343-FC25-TEL-VTE-CAR-UEM.docx"))
            },

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
def datasets_periode(tel):
    """
    Liste de datasets CSV à tester.
    Chaque dataset représente un cas réel rencontré.
    Pour utilisation avec depuis_periode
    """

    return {

        "Bilan TEL année 2024 (3 sessions, 2 CSV seulement)": {
            "chemin_IRIS_sessions": tel["chemin_IRIS_sessions"],
            "trigramme_formation": tel["trigramme_formation"],
            "annee": 2024,
            "periode": "Année",
            "chemin_bilanSessions": chemin_vers_unc(Path(r"P:\FORMATIONS_C\TEL\P07-bilan-sessions-et-bilan-formation\2024\P07-Pr05-F05-Bilan session-Année 2024-UEM.docx")),
            },


    }

# Pour choisir mes cas
"""
@pytest.mark.parametrize("nom_cas", [
    "Bilan unique par code IRIS. CSV dans eval formation",
    "Bilan unique par code IRIS. CSV introuvable",
])
"""
# Pour faire tous les cas
#@pytest.mark.parametrize("nom_cas", DATASETS_CODES_IRIS.keys())
#def test_bilanSession_depuis_codesIRIS(nom_cas, datasets_codesIRIS):
def verif_bilanSession_depuis_codesIRIS():

    # Sélection du dataset - Mode fonction
    dataset = DATASETS_CODES_IRIS["Bilan unique par code IRIS. CSV introuvable"]
    
    # Sélection du dataset - Mode pytest
    #dataset = datasets_codesIRIS[nom_cas]

    # Backups de mon environnement de travail
    backup_fichier_test(dataset["chemin_bilanSessions"])  # Bilan sessions word
    backup_fichier_test(dataset["chemin_bilanSessions"].with_suffix(".pdf"))  # Bilan sessions pdf

    set_iris_chemin_specifique(typeExport="Sessions", chemin=dataset["chemin_IRIS_sessions"])  # On redéfinit le chemin d'IRIS Sessions

    formation = Formation.avec_ajout_sessions(trigramme_formation=dataset["trigramme_formation"], codes_IRIS=dataset["codes_IRIS"])
    BilanSessions.depuis_codesIRIS(formation=formation, codes_IRIS=dataset["codes_IRIS"])


    # Restauration de mon environnement de travail
    restore_nom_fichier_test(dataset["chemin_bilanSessions"])  # Bilan sessions word
    restore_nom_fichier_test(dataset["chemin_bilanSessions"].with_suffix(".pdf"))  # Bilan sessions pdf



def verif_bilanSession_periode(datasets_periode): # 16411
    
    # Sélection du dataset
    dataset = datasets_periode["Bilan TEL année 2024 (3 sessions, 2 CSV seulement)"]
    
    # Backups de mon environnement de travail
    backup_fichier_test(dataset["chemin_bilanSessions"])  # Bilan sessions word
    backup_fichier_test(dataset["chemin_bilanSessions"].with_suffix(".pdf"))  # Bilan sessions pdf

    set_iris_chemin_specifique(typeExport="Sessions", chemin=dataset["chemin_IRIS_sessions"])  # On redéfinit le chemin d'IRIS Sessions

    formation = Formation.avec_ajout_sessions(trigramme_formation=dataset["trigramme_formation"], codes_IRIS=dataset["codes_IRIS"])
    BilanSessions.depuis_codesIRIS(formation=formation, codes_IRIS=dataset["codes_IRIS"])

    
    # Restauration de mon environnement de travail
    restore_nom_fichier_test(dataset["chemin_bilanSessions"])  # Bilan sessions word
    restore_nom_fichier_test(dataset["chemin_bilanSessions"].with_suffix(".pdf"))  # Bilan sessions pdf


def main():
    verif_bilanSession_depuis_codesIRIS()

if __name__ == "__main__":
    main()

