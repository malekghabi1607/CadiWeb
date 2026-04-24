from vte.utils.utils import *
from vte.core.iris_referentiel import set_iris_chemin_specifique
from vte.domain.formation import Formation

sys.path.append(str(Path(__file__).resolve().parents[1]))
from conftest import tel_data


    
tel = tel_data()
#2025 → Stats en 2024
#2024 → Pas de stats en 2023
DATASETS_PERIODE = {
    "Bilan TEL annee 2025 (1 sessions, 1 CSV) ; année 2024 (3 sessions, 2 CSV seulement)": {
    "chemin_IRIS_sessions": tel["chemin_IRIS_sessions"],
    "trigramme_formation": tel["trigramme_formation"],
    "annee": 2025,
    "chemin_bilanFormation": chemin_vers_unc(Path(r"P:\FORMATIONS_C\TEL\P07-bilan-sessions-et-bilan-formation\2025\P07-Pr05-F06-Bilan-formation-TEL-2025.docx")),
    
    # mocks
    #"mock_verifier_existance_fichier": True,  # True = on continue
    #"mock_demande_sessions_a_retenir": [15830, 15942, 16161],
    #"mock_statut_eval": "Exclu - Aucun CSV fourni",

    # attendu
    #"expected_statut": "Exclu - Aucun CSV fourni",
    },

}

dataset = DATASETS_PERIODE["Bilan TEL annee 2025 (1 sessions, 1 CSV) ; année 2024 (3 sessions, 2 CSV seulement)"]

# Backups de mon environnement de travail
backup_fichier_test(dataset["chemin_bilanFormation"])  # Bilan sessions word
backup_fichier_test(dataset["chemin_bilanFormation"].with_suffix(".pdf"))  # Bilan sessions pdf

set_iris_chemin_specifique(typeExport="Sessions", chemin=dataset["chemin_IRIS_sessions"])  # On redéfinit le chemin d'IRIS Sessions

formation = Formation.pour_traitement_bilanFormation_depuis_annee(
    trigramme_formation=dataset["trigramme_formation"],
    annee=dataset["annee"]
    )
#formation = Formation(trigramme_formation=dataset["trigramme_formation"])
#BilanSessions.depuis_periode(
#    formation=formation,
#    annee=dataset["annee"],
#    periode=dataset["periode"]
#    )


# Restauration de mon environnement de travail
restore_nom_fichier_test(dataset["chemin_bilanFormation"])  # Bilan sessions word
restore_nom_fichier_test(dataset["chemin_bilanFormation"].with_suffix(".pdf"))  # Bilan sessions pdf

