from __future__ import annotations
from pathlib import Path
from typing import Optional

import pandas as pd

from vte.utils.office import FichierExcel
from vte.services.iris_services import (
    ConfigExportIRIS,
    charger_df_extractsIRIS_originaux
)


class IRIS:
    """
    Représente un fichier IRIS individuel (sessions, formations, ventes, inscriptions).
    """

    def __init__(
        self,
        type_export: str,
        chemin_fichier: Optional[Path] = None
    ):
        self.type_export = type_export
        self.chemin_fichier = chemin_fichier

        self._config = ConfigExportIRIS(type_export)
        self._fe: Optional[FichierExcel] = None
        self._df: Optional[pd.DataFrame] = None

    # =========================
    # === CHARGEMENT FICHIER ==
    # =========================
    def charger(self) -> None:
        """
        Charge le fichier Excel IRIS et son DataFrame associé.
        """
        self._fe = FichierExcel(
            chemin_fichier=self.chemin_fichier,
            nom_onglet=self._config.nom_onglet
        )
        self._df = charger_df_extractsIRIS_originaux(self._fe, self._config)

    @property
    def df(self) -> pd.DataFrame:
        if self._df is None:
            self.charger()
        return self._df