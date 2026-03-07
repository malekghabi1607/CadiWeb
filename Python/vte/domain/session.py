from __future__ import annotations
from pathlib import Path
from typing import Optional, Protocol

#from vte.domain.formation import Formation
from vte.domain.evalStat import EvalStat_session, EvalStat_formation
from vte.domain.evalStat import *

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
        """
        Crée une instance de Session a minima.

        Contient le code IRIS de la session.

        Fait référence à sa formation parente.

        Args:
            formation (Formation_protocol): Formation parente de la session
            code_IRIS (int): Code IRIS de la session
        """
        
        self._formation = formation  # Protocol pour éviter les références circulaires
        self._code_IRIS = code_IRIS
        
        # Une session a une évaluation stagiaire de la session /!\ Faire distinction entre EvalStat natif et le mien → On va prendre le mien
        self._eval:Optional[EvalStat_session] = None
   
    @classmethod
    def avec_traitement_evalStat(cls, formation:Formation_protocol, code_IRIS: int, chemin_csv:Optional[Path|str]=None, chemin_IRIS_sessions:Optional[Path]=None, ecrire_eval_formation:bool=False, ouvrirDossier:bool=False) -> Session:
        """
        Crée une l'instance de Session en traitant son EvalStat.
        
        Crée l'instance EvalStat d'une session et traite cet EvalStat.
        
        Crée l'Excel EvalStat d'une session à partir d'un CSV (s'il n'est pas donné, on ouvre un filedialog)

        L'évaluation de la formation est mise à jour avec ces nouvelles données et est sauvée en fin de traitement selon le critère ecrire_eval_formation.

        Args:
            formation (Formation_protocol): Formation parente de la session
            code_IRIS (int): Code IRIS de la session
            chemin_csv (Optional[Path|str], optional): chemin du CSV à traiter. S'il est None, on ouvre un filedialog
            chemin_IRIS_sessions (Optional[Path], optional): Chemin du fichier IRIS sessions à employer si l'utilisateur ne veut pas celui par défaut. Defaults = None = Fichier généré le plus récent dans le répertoire donné en config.
            ecrire_eval_formation (bool, optional): Pour écrire physiquement l'Excel eval formation en fin de traitement. Si False, il devra être écrit ailleurs (à l'endroit où il y a la boucle pour du multi-traitement typiquement). Défaut = True.
            ouvrirDossier (bool, optional): Ouvre le répertoire de l'EvalStat généré. Defaut = False.
        """
        instance = cls(formation = formation, code_IRIS = code_IRIS)
        instance._ajout_evalStat_avec_traitement(
            session=instance,
            chemin_csv=chemin_csv,
            chemin_IRIS_sessions=chemin_IRIS_sessions,
            ecrire_eval_formation=ecrire_eval_formation,
            ouvrirDossier=ouvrirDossier
            )
        return instance

    # =========================
    # === METHODES INTERNES ===
    # =========================
    def _ajout_evalStat(self) -> None:
        """
        Crée un objet EvalStat vide d'une session.
        """
        self._eval = EvalStat_session(session=self)

    def _ajout_evalStat_avec_traitement(self, chemin_csv:Optional[Path|str]=None, chemin_IRIS_sessions:Optional[Path]=None, ecrire_eval_formation:bool=False, ouvrirDossier:bool=False) -> None:
        """
        Crée l'Excel EvalStat d'une session.
         
        Si CSV non donné, alors on ouvre un filedialog.

        L'évaluation de la formation est mise à jour avec ces nouvelles données.

        L'évaluation de la formation est sauvée en fin de traitement. Ca pourrait être fait ailleurs si boucle de traitement de plusieurs EvalStat de sessions d'une même formation.

        Args:
            session (Session_protocol): La session à laquelle est affectée l'EvalStat
            chemin_csv (Optional[Path|str], optional): chemin du CSV à traiter. S'il est None, on ouvre un filedialog
            chemin_IRIS_sessions (Optional[Path], optional): Chemin du fichier IRIS sessions à employer si l'utilisateur ne veut pas celui par défaut. Defaults = None = Fichier généré le plus récent dans le répertoire donné en config.
            ouvrirDossier (bool, optional): Ouvre le répertoire de l'EvalStat généré. Defaut = False.
        """
        self._eval = EvalStat_session.avec_traitement(
            session=self,
            chemin_csv=chemin_csv,
            chemin_IRIS_sessions=chemin_IRIS_sessions,
            ecrire_eval_formation=ecrire_eval_formation,
            ouvrirDossier=ouvrirDossier
            )
        






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
    def eval(self) -> EvalStat_session:
        return self._eval
    
    @property
    def eval_formation(self) -> EvalStat_formation:
        return self._formation.eval
    