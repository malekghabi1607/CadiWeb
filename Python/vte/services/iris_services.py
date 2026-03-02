from __future__ import annotations

from pathlib import Path
from typing import Optional, Tuple

from vte.domain.iris import IRIS, IRIS_natif
from vte.utils.utils import convertir_tuple_path, convertir_tuple_str


class IRISServices:
    """
    Services métier autour des exports IRIS.
    """


    @staticmethod
    def concatener_plusieursTypes(typesExports:str|Tuple[str], depuis_config:bool=False) -> None:
        """
        Génère un ou plusieurs fichier IRIS en concaténant plusieurs extracts originaux.

        Les fichiers à traiter sont :
           - soit spécifiés dans config_extractsIRIS.py (depuis_config = True) ;
           - soit on ouvre un filedialog (depuis_config = False = valeur par défaut).
        
        Gère le traitement de plusieurs types.

        :param typesExports: Types des exports IRIS à traiter, un ou plusieurs de ceux-ci ("Formations", "Sessions", "Ventes", "Inscriptions").
        :type typesExports: str|Tuple[str]
        :param depuis_config: Si True, on exploite les fichiers spécifiés dans config_extractsIRIS.py. Sinon (défaut) on ouvre un filedialog.
        :type depuis_config: bool
        
        :return: Ne retourne rien
        :rtype: None

        :example:

        >>> concatener_plusieursTypes(
                typesExports = ("Sessions", "Formations", "Ventes", "Inscriptions"), 
                depuis_config = True
                )
        >>> concatener_plusieursTypes(typesExports = "Sessions")


        .. seealso:: Rien du tout
        .. warning:: Rien du tout.
        .. note:: Remplace complètement l'ancien fichier (pas de mise à jour incrémentale)
        .. todo:: Rien du tout.      
        """
        def f_dict_chemins(depuis_config:bool,typeExport:str) -> Tuple[Path]:
            """
            Fonction à employer en fonction du choix utilisateur depuis_config/filedialog.
            
            Si depuis_config = True  : on charge les chemins depuis config_extractsIRIS.
            Si depuis config = False : on ouvre un filedialog pour choix utilisateur.

            :return: Un tuple de chemins d'extracts IRIS originaux (Path)
            :rtype: Tuple[Path]
            """
            if depuis_config:
                # Cas depuis_config : on crée un tuple de tous les chemins pour le type export donné en argument
                return tuple((IRIS_natif.cei(typeExport)._input.repertoire / nom) for nom in IRIS.DICT_EXPORTS_IRIS["chemins_fichiersInput"][typeExport])
            else :
                # Cas avec ouverture filedialog
                return IRIS_natif.choisir_fichiers(typeExport)

        # Si typesExports est un str, alors on convertit en tuple
        typesExports = convertir_tuple_str(typesExports)

        # On définit le dictionnaire des chemins à employer pour chaque type d'export
        # (On boucle sur typesExports et pour chacun d'eux, soit on prend les données de la config soit filedialog.)
        dict_chemins_fichiersInpout = {
            typeExport: f_dict_chemins(depuis_config, typeExport) for typeExport in typesExports
            }

        # On traite les fichiers par type d'export
        for typeExport in typesExports:
            IRIS_natif.avec_traitement(
                typeExport=typeExport,
                chemins_fichiersInput=dict_chemins_fichiersInpout[typeExport]
                )

