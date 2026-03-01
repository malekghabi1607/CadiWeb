from typing import Optional

from fdc import *
from EvalStat import *
from bilanFormation import *
from bilanSession import *
from vte.session import *

# ======================================================================================
# CLASSE FORMATION
# ======================================================================================
class Formation:
    def __init__(self, trigramme_formation: str):
        self._trigramme_formation:str = trigramme_formation

        # Une formation a une fiche de coûts
        self._fdc:Optional[FdC] = None
        
        # Une formation a une évaluation stagiaire de la formation (regroupement de toutes les évaluations stagiaires de toutes les sessions)
        #self._eval:Optional[EvalStat] = None

        # Une formation a un ou plusieurs bilans de formation (annuel)
        #self.bilans_formation:Optional[dict[int, BilanFormation]] = {}  # bilans_formation[2025] : Index = année du bilan

        # Une formation a un ou plusieurs bilans de session par année (soit 1 par semestre, soit annuel s'il n'y a qu'une session annuellement)
        #self.bilans_session:Optional[dict[int, dict[int, BilanSession]]] = {}  #bilans_session[2025][0] : Index1 = année du bilan ; Index2 = période du bilan (1 = 1er semestre ; 2 = 2nd semestre ; 0 = annuel)

        # Une formation a une ou plusieurs sessions
        #self.sessions:Optional[dict[int, Session]] = {}  # Index = code IRIS de la formation
