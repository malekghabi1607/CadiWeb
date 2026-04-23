import pytest
import sys
from pathlib import Path
from pprint import *

from vte.core.iris_referentiel import set_iris_chemin_specifique
from vte.domain.evalStat import EvalStat_session
from vte.domain.formation import Formation
from vte.domain.session import Session
from vte.domain.bilanSessions import BilanSessions
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
DATASETS_CODES_IRIS = {
    "Bilan unique par code IRIS. CSV dans eval formation": {
        "chemin_IRIS_sessions": tel["chemin_IRIS_sessions"],
        "trigramme_formation": tel["trigramme_formation"],
        "codes_IRIS": tel["code_IRIS"],
        "chemin_bilanSessions": tel["chemin_bilanSessions"],

        # mocks
        "mock_verifier_existance_fichier": True,  # True = on continue
        #"mock_statut_eval": "Traité",

        # attendu
        #"expected_statut": "OK",


        },


    "Bilan unique par code IRIS. CSV introuvable": {
        "chemin_IRIS_sessions": tel["chemin_IRIS_sessions"],
        "trigramme_formation": tel["trigramme_formation"],
        "codes_IRIS": 17343,
        "chemin_bilanSessions": chemin_vers_unc(Path(r"P:\FORMATIONS_C\TEL\P07-bilan-sessions-et-bilan-formation\2025\P07-Pr05-F05-Bilan session-S-17343-FC25-TEL-VTE-CAR-UEM.docx")),

        # mocks
        "mock_verifier_existance_fichier": True,  # True = on continue
        #"mock_statut_eval": "Exclu - Aucun CSV fourni",

        # attendu
        #"expected_statut": "Exclu - Aucun CSV fourni",
        },

}

DATASETS_PERIODE = {
    "Bilan TEL annee 2024 (3 sessions, 2 CSV seulement)": {
        "chemin_IRIS_sessions": tel["chemin_IRIS_sessions"],
        "trigramme_formation": tel["trigramme_formation"],
        "annee": 2024,
        "periode": "Année",
        "chemin_bilanSessions": chemin_vers_unc(Path(r"P:\FORMATIONS_C\TEL\P07-bilan-sessions-et-bilan-formation\2024\P07-Pr05-F05-Bilan session-Année 2024-UEM.docx")),
        
        # mocks
        "mock_verifier_existance_fichier": True,  # True = on continue
        "mock_demande_sessions_a_retenir": [15830, 15942, 16161],
        #"mock_statut_eval": "Exclu - Aucun CSV fourni",

        # attendu
        #"expected_statut": "Exclu - Aucun CSV fourni",
        },

}

STATUTS_EVAL_NON_TRAITE = {
    15830 : "Exclu - Aucun CSV fourni",
    11090 : "Exclu - CSV vide / Aucun retour",
    12768 : "Exclu - Aucun CSV fourni",  # On n'a que les pdf
    17343 : "Exclu - Aucun CSV fourni",
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
    monkeypatch.setattr(
        "vte.domain.bilanSessions.verifier_existance_fichier",  # Fonction à remplacer
        lambda path: dataset["mock_verifier_existance_fichier"]  # Fonction de remplacement
    )

    # --- mock EvalStat ---
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
    def fake_demande_sessions_a_retenir(*args, **kwargs):
        return dataset["mock_demande_sessions_a_retenir"]

    monkeypatch.setattr(
        IRIS_sessions,
        "demande_sessions_a_retenir",
        fake_demande_sessions_a_retenir
    )

    # --- mock envoyer_mail_chef_unite ---
    def fake_envoyer_mail_chef_unite(*args, **kwargs):
        pass

    monkeypatch.setattr(
        "vte.domain.bilanSessions.BilanSessions._envoyer_mail_chef_unite",
        fake_envoyer_mail_chef_unite
    )


# Pour choisir mes cas
"""
@pytest.mark.parametrize("nom_cas", [
    "Bilan unique par code IRIS. CSV dans eval formation",
    "Bilan unique par code IRIS. CSV introuvable",
])
"""
# Pour faire tous les cas
@pytest.mark.parametrize("nom_cas, dataset", DATASETS_CODES_IRIS.items())
def test_bilanSession_depuis_codesIRIS(nom_cas, dataset, monkeypatch):
#def verif_bilanSession_depuis_codesIRIS():

    # MODE FONCTION
    #dataset = DATASETS_CODES_IRIS["Bilan unique par code IRIS. CSV introuvable"]
    
    # MODE PYTEST
    apply_mocks(monkeypatch, dataset)  # Appliquer les mocks



    # Backups de mon environnement de travail
    backup_fichier_test(dataset["chemin_bilanSessions"])  # Bilan sessions word
    backup_fichier_test(dataset["chemin_bilanSessions"].with_suffix(".pdf"))  # Bilan sessions pdf

    set_iris_chemin_specifique(typeExport="Sessions", chemin=dataset["chemin_IRIS_sessions"])  # On redéfinit le chemin d'IRIS Sessions

    formation = Formation.pour_traitement_bilanSessions_depuis_codesIRIS(codes_IRIS=dataset["codes_IRIS"])
    #formation = Formation.avec_ajout_sessions(trigramme_formation=dataset["trigramme_formation"], codes_IRIS=dataset["codes_IRIS"])
    #BilanSessions.depuis_codesIRIS(formation=formation, codes_IRIS=dataset["codes_IRIS"])


    # Restauration de mon environnement de travail
    restore_nom_fichier_test(dataset["chemin_bilanSessions"])  # Bilan sessions word
    restore_nom_fichier_test(dataset["chemin_bilanSessions"].with_suffix(".pdf"))  # Bilan sessions pdf

    #input(f"Fin du test de {nom_cas}, appuer sur une touche")


@pytest.mark.parametrize("nom_cas, dataset", DATASETS_PERIODE.items())
def test_bilanSession_depuis_periode(nom_cas, dataset, monkeypatch):
#def verif_bilanSession_periode():
    # MODE FONCTION
    #dataset = DATASETS_PERIODE["Bilan TEL année 2024 (3 sessions, 2 CSV seulement)"]
    
    # MODE PYTEST
    apply_mocks(monkeypatch, dataset)  # Appliquer les mocks
    
    # Backups de mon environnement de travail
    backup_fichier_test(dataset["chemin_bilanSessions"])  # Bilan sessions word
    backup_fichier_test(dataset["chemin_bilanSessions"].with_suffix(".pdf"))  # Bilan sessions pdf

    set_iris_chemin_specifique(typeExport="Sessions", chemin=dataset["chemin_IRIS_sessions"])  # On redéfinit le chemin d'IRIS Sessions

    formation = Formation.pour_traitement_bilanSessions_depuis_periode(
        dataset["trigramme_formation"],
        annee=dataset["annee"],
        periode=dataset["periode"]
        )
    #formation = Formation(trigramme_formation=dataset["trigramme_formation"])
    #BilanSessions.depuis_periode(
    #    formation=formation,
    #    annee=dataset["annee"],
    #    periode=dataset["periode"]
    #    )

    
    # Restauration de mon environnement de travail
    restore_nom_fichier_test(dataset["chemin_bilanSessions"])  # Bilan sessions word
    restore_nom_fichier_test(dataset["chemin_bilanSessions"].with_suffix(".pdf"))  # Bilan sessions pdf


def main():
    #verif_bilanSession_periode()
    print("CAS DISPONIBLES:")
    for i, k in enumerate(DATASETS_CODES_IRIS.keys()):
        print(i, k)

    idx = int(input("Choisir un cas : "))

    nom_cas = list(DATASETS_CODES_IRIS.keys())[idx]
    dataset = DATASETS_CODES_IRIS[nom_cas]

    monkeypatch = pytest.MonkeyPatch()

    test_bilanSession_depuis_codesIRIS(
        nom_cas,
        dataset,
        monkeypatch
    )

if __name__ == "__main__":
    main()

