from __future__ import annotations
from typing import Dict, Tuple

from vte.domain.fichiers.iris import IRIS

from __future__ import annotations

from vte.core import config_extractsIRIS, config
from vte.domain.fichiers.iris import ConfigExportIRIS


# ======================================================================================
# REFERENTIEL IRIS (chargé une seule fois)
# Référentiel partagé, chargé une seule fois (configs, constantes lourdes)
# ======================================================================================

_SESSIONS = ConfigExportIRIS(**config.IRIS_SESSIONS_PARAMS)
_FORMATIONS = ConfigExportIRIS(**config.IRIS_FORMATIONS_PARAMS)
_VENTES = ConfigExportIRIS(**config.IRIS_VENTES_PARAMS)
_INSCRIPTIONS = ConfigExportIRIS(**config.IRIS_INSCRIPTIONS_PARAMS)

DICT_EXPORTS_IRIS = {
    """
    Dictionnaire (clefs = [ ; ]) :
       - "ConfigExportIRIS" : propriétés des Exports IRIS (chemins, modèles, nb lignes avant tableau...). Ces données sont renseignées pour :
            - les inputs (natifs IRIS) ;
            - les modèles à employer (qui sont peuplés par les inputs) ;
            - les output (fichiers traités / concaténés).
       - "chemins_fichiersInput" : liste des chemins des fichiers natifs IRIS à concaténer pour obtenir les fichiers output
       - "colonnes_modele" : je ne suis plus sûr : soit liste soit ordre soit nouveau nom pour le modèle versus l'input
    """

    "ConfigExportIRIS": {
        "Sessions": _SESSIONS,
        "Formations": _FORMATIONS,
        "Ventes": _VENTES,
        "Inscriptions": _INSCRIPTIONS,
    },


    "chemins_fichiersInput": {
        "Sessions": config_extractsIRIS._tSessions,
        "Formations": config_extractsIRIS._tFormations,
        "Ventes": config_extractsIRIS._tVentes,
        "Inscriptions": config_extractsIRIS._tInscriptions,
    },

    "colonnes_modele": {
        "Sessions": None,
        "Formations": None,
        "Ventes": None,
        "Inscriptions": config_extractsIRIS._colonnes_modele_inscriptions,
    },
}

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