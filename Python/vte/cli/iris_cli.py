
from vte.services.iris_services import IRISServices

from vte.core.iris_referentiel import IRISReferentiel
from vte.services.iris_services import choisir_fichier_IRIS








"""
# Par ChatGPT
def charger_sessions_IRIS():
    chemin = choisir_fichier_IRIS("Sessions")
    iris = IRISReferentiel.get("Sessions", chemin)
    iris.charger()
    return iris


def maj_exports_IRIS(typesExports, depuis_config=False):
    IRISServices.concatener_plusieurs_types(typesExports, depuis_config)
"""