from __future__ import annotations

from functools import cached_property
import math
from pathlib import Path
from typing import Optional, Protocol
from datetime import date, datetime

from vte.core import config
from vte.utils.utils import *
from vte.utils.office import FichierExcel
from vte.utils.utils_instn import construire_chemin_config, recupere_trig_formation_depuis_chemin

# ======================================================================================
# PROTOCOLES
# (pour faire passer les informations des objets parents sans ref circulaires)
# ======================================================================================
class Formation_protocol(Protocol):
    """
    Protocol de Formation : permet de simuler une formation en évitant les références circulaires
    """
    @property
    def trigramme_formation(self) -> str: ...


# ======================================================================================
# CLASSE Specs pédago
# ======================================================================================
class Specs:
    """
    Classe qui gère l'ouverture d'une spec pédagogique INSTN et l'accès à ses différentes valeurs
    """

    
    # === VARIABLES DE CLASSE COMMUNES À TOUTES LES INSTANCES ===




    # ====================
    # === CONSTRUCTEUR ===
    # ====================
    def __init__(self, formation: Optional[Formation_protocol] = None, chemin: Optional[Path] = None) -> None:
        """
        Initialise une instance Specs.

        :param formation: l'instance de Formation à laquelle appartiennent les specs (nécessaire uniquement pour faciliter la sélection du fichier de specs (pré-sélection répertoire))
        :type formation: Optional[Formation_protocol], optional
        :param chemin: Chemin des specs
        :type chemin: Optional[Path], optional
        """
        # Variables propres aux specs
        self._formation:Optional[Formation_protocol] = formation  # Trigramme de la formation ; nécessaire uniquement pour facilite la sélection du fichier de FdC (pré-sélection répertoire)
        self._chemin: Optional[Path] = chemin  # Chemin des specs

    @classmethod
    def depuis_chemin(cls, formation: Optional[Formation_protocol] = None, chemin: Optional[Path|str] = None) -> Specs:
        """
        Initialise une instance Specs à partir du chemin des specs.

        Si aucun chemin n'est donné, on ouvre un filedialog.

        :param formation: l'instance de Formation pour la fiche de coûts (nécessaire uniquement pour facilite la sélection du fichier de FdC (pré-sélection répertoire))
        :type formation: Optional[Formation_protocol], optional
        :param chemin: chemin des specs. Défaut = None
        :type chemin: Optional[Path|str], optional
        """
        instance = Specs(formation=formation, chemin=chemin)

        if instance._chemin is None:
            instance._chemin = instance._choisir_specs()

        return instance
        
    # =========================
    # === METHODES INTERNES ===
    # =========================
    def _choisir_specs(self) -> Path:
        """
        Ouvre un filedialog pour demander à l'utilisateur de sélectionner les specs d'une formation.
        On pointe au mieux sur le répertoire des specs pour la boîte de dialogue.
        
        :return: le chemin des specs
        :rtype: Path | None
        """

        return choisir_fichier(titre=f"Sélectionner les dernières specs pédagogiques de la formation",
                        types_fichiers=[("Fichiers PDF", "*.pdf"), ("Documents Word", "*.docx")],
                        dossier_initial=self.dossier_plan_classement,
                        texte_bouton_choisir=f"Choisir specs à nouveau"
                        )

    
        

    


    # =========================
    # === METHODES EXTERNES ===
    # =========================
    @staticmethod
    def construire_chemin_repertoire_specs(trigramme_formation:Optional[str]=None) -> Path:
        """
        Construit le chemin du répertoire de la fiche des specs de la formation (à partir des données de la config REPERTOIRE_FDC) :
            - le chemin est transformé en unc (s'il y a un raccourci lecteur réseau sur le poste de l'utilisateur on transforme en chemin réseau complet) ;
            - l'utilisateur peut optimiser le chemin (chemin le plus long entre l'attendu et ce qui existe).

        :param trigramme_formation: Trigramme de la formation. Défaut = None
        :type trigramme_formation: Optional[str], optional

        :return: Le chemin de sortie du bilan de sessions.
        :rtype: Path
        """
        return construire_chemin_config(
            chemin_a_completer = config.REPERTOIRE_SPECS,
            trigramme_formation = trigramme_formation
        )



    # =========================
    # === GETTERS / SETTERS ===
    # =========================
    @cached_property
    def trigramme_formation(self) -> str|None:
        """
        Renvoie le trigramme de la formation.
        On la lit soit :
           - depuis l'instance Formation ;
           - en l'extrayant depuis le chemin du fichier Excel de la fiche de coûts
        
        Si _formation et _fe sont None les 2, alors on renvoie None

        :return: le trigramme de la formation
        :rtype: int|None
        """
        if self._formation is not None:
            return self._formation.trigramme_formation
        elif self._fe is not None:
            return recupere_trig_formation_depuis_chemin(self._fe.chemin_fichier)
        else :
            return None
    
    @property
    def chemin(self) -> Path:
        """
        Renvoie le chemin des specs.
        
        :return: le chemin des specs
        :rtype: Path
        """
        return self._chemin


    @property
    def date_derniere_modification(self) -> datetime:
        """
        Extrait la date de la fiche de coût à partir de la dernière date de modif (valeur système windows)
        
        :return: la date de la fiche de coût à partir de la dernière date de modif (valeur système windows)
        :rtype: datetime
        """
        return datetime.fromtimestamp(self.chemin.stat().st_mtime)
    
    @property
    def date_derniere_modification_str(self) -> str:
        """
        Extrait la date de la fiche de coût à partir de la dernière date de modif (valeur système windows)
        
        :return: la date de la fiche de coût à partir de la dernière date de modif (valeur système windows)
        :rtype: str
        """
        return self.date_derniere_modification.strftime("%d/%m/%Y")

    @cached_property
    def dossier_plan_classement(self) -> Path:
        return Specs.construire_chemin_repertoire_specs(trigramme_formation=self.trigramme_formation)


    # ========================================================
    # === GETTERS données FdC (lié à la version de la FdC) ===
    # ========================================================
    
