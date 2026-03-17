from __future__ import annotations
from pathlib import Path
from typing import Optional, Protocol

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
    def __init__(
            self, 
            formation:Formation_protocol, 
            code_IRIS: int
            ):
        """
        Crée une instance de Session a minima.

        Contient le code IRIS de la session.
        Fait référence à sa formation parente.

        :param formation: Formation parente de la session
        :type formation: Formation_protocol
        :param code_IRIS: Code IRIS de la session
        :type code_IRIS: int
        """
        
        self._formation = formation  # Protocol pour éviter les références circulaires
        self._code_IRIS = code_IRIS
        
        # Une session a une évaluation stagiaire de la session /!\ Faire distinction entre EvalStat natif et le mien → On va prendre le mien
        self._eval:Optional[EvalStat_session] = None

    @classmethod
    def avec_ouverture_ou_traitement_evalStat(
        cls, 
        formation:Formation_protocol, 
        code_IRIS: int, 
        chemin_csv:Optional[Path|str]=None, 
        ecrire_eval_formation:bool=True, 
        ouvrirDossier:bool=False
        ) -> Session:
        """
        Crée une l'instance de Session en traitant son EvalStat.

        Crée l'instance EvalStat d'une session et traite cet EvalStat.

        Crée l'Excel EvalStat d'une session à partir d'un CSV (s'il n'est pas donné, on ouvre un filedialog)
        L'évaluation de la formation est mise à jour avec ces nouvelles données et est sauvée en fin de traitement selon le critère ecrire_eval_formation.

        :param formation: Formation parente de la session
        :type formation: Formation_protocol
        :param code_IRIS: Code IRIS de la session
        :type code_IRIS: int
        :param chemin_csv: chemin du CSV à traiter. S'il est None, on ouvre un filedialog
        :type chemin_csv: Optional[Path|str]
        :param ecrire_eval_formation: Pour écrire physiquement l'Excel eval formation en fin de traitement. Si False, il devra être écrit ailleurs (à l'endroit où il y a la boucle pour du multi-traitement typiquement). Défaut = True.
        :type ecrire_eval_formation: bool
        :param ouvrirDossier: Ouvre le répertoire de l'EvalStat généré. Defaut = False.
        :type ouvrirDossier: bool
        """
        instance = cls(formation = formation, code_IRIS = code_IRIS)
        
        instance._eval = EvalStat_session.avec_ouverture_ou_traitement(
            session=instance,
            chemin_csv=chemin_csv,
            ecrire_eval_formation=ecrire_eval_formation,
            ouvrirDossier=ouvrirDossier
            )
        return instance

    # =========================
    # === METHODES INTERNES ===
    # =========================


    # =========================
    # === METHODES EXTERNES ===
    # =========================






    # ==================================================================================
    # GETTERS / SETTERS
    # ==================================================================================    
    @property
    def code_IRIS(self) -> int:
        return self._code_IRIS
    
    @property
    def trigramme_formation(self) -> str:
        return self._formation.trigramme_formation
    
    @property
    def eval(self) -> Optional[EvalStat_session]:
        return self._eval
    
    @eval.setter
    def eval(self, valeur:Optional[EvalStat_session]) -> None:
        self._eval = valeur

    @property
    def eval_formation(self) -> EvalStat_formation:
        return self._formation.eval
    