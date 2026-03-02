from __future__ import annotations
from typing import Dict, Tuple

from vte.domain import IRIS



# ======================================================================================
# REFERENTIEL IRIS (chargé une seule fois)
# Référentiel partagé, chargé une seule fois (configs, constantes lourdes)
# ======================================================================================


class IRISReferentiel:
    """
    Référentiel applicatif des fichiers IRIS déjà chargés.
    Évite de relire plusieurs fois les mêmes fichiers Excel (coûteux).
    """

    _cache: Dict[Tuple[str, str], IRIS] = {}

    @classmethod
    def get(cls, type_export: str, chemin: str) -> IRIS:
        key = (type_export, chemin)

        if key not in cls._cache:
            cls._cache[key] = IRIS(
                type_export=type_export,
                chemin_fichier=chemin
            )

        return cls._cache[key]

    @classmethod
    def clear(cls) -> None:
        cls._cache.clear()