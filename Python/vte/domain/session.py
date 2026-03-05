from __future__ import annotations
from typing import Optional, Protocol

#from vte.domain.formation import Formation
from vte.domain.evalStat import EvalStat_session, EvalStat_formation

class Formation_protocol(Protocol):
    @property
    def trigramme_formation(self) -> str: ...

    @property
    def eval(self) -> EvalStat_formation: ...


# ======================================================================================
# CLASSE SESSION
# ======================================================================================


class Session:
    def __init__(self, formation:Formation_protocol, code_IRIS: int):
        
        self._formation = formation  # Protocol pour éviter les références circulaires
        self._code_IRIS = code_IRIS
        
        # Une session a une évaluation stagiaire de la session /!\ Faire distinction entre EvalStat natif et le mien → On va prendre le mien
        self._eval:Optional[EvalStat_session] = None

    @classmethod
    def avec_ajout_evalStat(cls, formation:Formation_protocol, code_IRIS: int) -> Session:
        instance = cls(formation = formation, code_IRIS = code_IRIS)
        instance.ajout_evalStat()
        return instance

    # =========================
    # === METHODES EXTERNES ===
    # =========================
    def ajout_evalStat(self) -> None:
        self._eval = EvalStat_session(session=self)

    # ==================================================================================
    # GETTERS / SETTERS
    # ==================================================================================    
    @property
    def code_IRIS(self) -> int|None:
        return self._code_IRIS
    
    @property
    def trigramme_formation(self) -> str|None:
        return self._formation.trigramme_formation
    
    @property
    def eval_session(self) -> EvalStat_session:
        return self._eval
    
    @property
    def eval_formation(self) -> EvalStat_formation:
        return self._formation.eval