from __future__ import annotations
from typing import Dict, Tuple

from vte.domain.iris import IRIS, IRIS_traite



# ======================================================================================
# REFERENTIEL IRIS (chargé une seule fois)
# Référentiel partagé, chargé une seule fois (configs, constantes lourdes)
# ======================================================================================


# ======================================================================================
# CACHE INTERNE
# ======================================================================================

_cache: Dict[str, IRIS] = {}



# ======================================================================================
# API PUBLIQUE
# ======================================================================================

def charge_iris(typeExport: str) -> IRIS_traite:
    """
    Retourne une instance IRIS chargée pour un type donné.
    Si elle n'existe pas encore, elle est créée et mise en cache.
    """

    if typeExport not in _cache:
        iris = IRIS_traite(typeExport=typeExport)
        _cache[typeExport] = iris

    return _cache[typeExport]


def vider_cache() -> None:
    """
    Vide le cache (utile pour tests).
    """
    _cache.clear()