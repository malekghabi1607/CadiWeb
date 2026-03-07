from __future__ import annotations
from pathlib import Path
from typing import Optional
from collections.abc import Iterable

from vte.domain.evalStat import EvalStat_formation
from vte.fdc import FdC
from vte.domain.session import Session
from vte.utils.utils import convertir_collection

# ======================================================================================
# CLASSE FORMATION
# ======================================================================================
class Formation:
    def __init__(self, trigramme_formation: str):
        self._trigramme_formation:str = trigramme_formation

        # Une formation a une fiche de coûts
        #self._fdc:Optional[FdC] = None
        
        # Une formation a une évaluation stagiaire de la formation (regroupement de toutes les évaluations stagiaires de toutes les sessions)
        self._eval:Optional[EvalStat_formation] = None

        # Une formation a un ou plusieurs bilans de formation (annuel)
        #self.bilans_formation:Optional[dict[int, BilanFormation]] = {}  # bilans_formation[2025] : Index = année du bilan

        # Une formation a un ou plusieurs bilans de session par année (soit 1 par semestre, soit annuel s'il n'y a qu'une session annuellement)
        #self.bilans_session:Optional[dict[int, dict[int, BilanSession]]] = {}  #bilans_session[2025][0] : Index1 = année du bilan ; Index2 = période du bilan (1 = 1er semestre ; 2 = 2nd semestre ; 0 = annuel)

        # Une formation a une ou plusieurs sessions
        self._sessions:dict[int, Session] = {}  # Index = code IRIS de la formation

    @classmethod
    def avec_ouverture_evalStat(cls, trigramme_formation: str) -> Formation:
        """
        Crée une Formation en créant un evalStat formation (ou en l'ouvrant s'il existe déjà)
        
        Args:
            trigramme_formation (str): trigramme de la formation
        """
        instance = cls(trigramme_formation=trigramme_formation)
        instance._ajout_evalStat_avec_ouverture()
        return instance

    @classmethod
    def avec_creation_sessions(cls, trigramme_formation: str, codes_IRIS:int|Iterable[int]) -> Formation:
        """
        Initialise une formation en créant une ou plusieurs sessions

        Args:
            trigramme_formation (str): trigramme de la formation
            codes_IRIS (int): codes IRIS des sessions à créer

        Returns:
            Formation: Instance de Formation
        """
        instance = cls(trigramme_formation=trigramme_formation)
        instance.ajout_sessions(codes_IRIS=codes_IRIS)
        return instance
    
    @classmethod
    def avec_creation_sessions_et_traitement_EvalStat(cls, trigramme_formation: str, codes_IRIS:int|Iterable[int], chemin_csv:Optional[Path|str]=None, chemin_IRIS_sessions:Optional[Path]=None, ouvrirDossier:bool=False) -> Formation:
        """
        Initialise une formation en créant une ou plusieurs sessions
        
        Crée l'Excel EvalStat d'une session à partir d'un CSV (s'il n'est pas donné, on ouvre un filedialog)
        L'évaluation de la formation est mise à jour avec ces nouvelles données.

        L'évaluation de la formation est sauvée en fin de traitement. Ca pourrait être fait ailleurs si boucle de traitement de plusieurs EvalStat de sessions d'une même formation.

        Args:
            trigramme_formation (str): trigramme de la formation
            codes_IRIS (int): codes IRIS des sessions à créer
            chemin_csv (Optional[Path|str], optional): chemin du CSV à traiter. S'il est None, on ouvre un filedialog
            chemin_IRIS_sessions (Optional[Path], optional): Chemin du fichier IRIS sessions à employer si l'utilisateur ne veut pas celui par défaut. Defaults = None = Fichier généré le plus récent dans le répertoire donné en config.
            ouvrirDossier (bool, optional): Ouvre le répertoire de l'EvalStat généré. Defaut = False.

        Returns:
            Formation: Instance de Formation
        """
        instance = cls.avec_ouverture_evalStat(trigramme_formation=trigramme_formation)
        instance.ajout_sessions(codes_IRIS=codes_IRIS)

        instance.traiter_eval_sessions(
            chemin_csv=chemin_csv,
            chemin_IRIS_sessions=chemin_IRIS_sessions,
            ouvrirDossier=ouvrirDossier
        )
        
        return instance



    # =========================
    # === METHODES INTERNES ===
    # =========================  
    def ajout_evalStat_objetVierge(self) -> None:
        if self._eval is None:
            self._eval = EvalStat_formation(formation=self)

    def _ajout_evalStat_avec_ouverture(self) -> None:
        if self._eval is None:
            self._eval = EvalStat_formation.avec_ouverture(formation=self)



    # =========================
    # === METHODES EXTERNES ===
    # =========================
    def ajout_sessions(self, codes_IRIS:int|Iterable[int]) -> None:
        """
        Ajouter une ou plusieurs sessions au dictionnaire de la formation

        Args:
            codes_IRIS (int | Iterable[int]): code IRIS (int) ou codes IRIS (Iterable : peut être une liste ou un tuple)
        """
        codes = convertir_collection(codes_IRIS)

        for code_IRIS in codes:
            self._sessions[code_IRIS] = Session(formation=self, code_IRIS=code_IRIS)

    def traiter_eval_sessions(
        self,
        chemin_csv: Optional[Path|str],
        chemin_IRIS_sessions: Optional[Path],
        ouvrirDossier: bool
    ):
        """
        Traite les EvalStat de plusieurs sessions. Pour chacune d'elle :

        Crée l'Excel EvalStat d'une session à partir d'un CSV (s'il n'est pas donné, on ouvre un filedialog)

        L'évaluation de la formation est mise à jour avec ces nouvelles données et est sauvée en fin de traitement de la boucle.

        Args:
            chemin_csv (Optional[Path|str], optional): chemin du CSV à traiter. S'il est None, on ouvre un filedialog
            chemin_IRIS_sessions (Optional[Path], optional): Chemin du fichier IRIS sessions à employer si l'utilisateur ne veut pas celui par défaut. Defaults = None = Fichier généré le plus récent dans le répertoire donné en config.
            ouvrirDossier (bool, optional): Ouvre le répertoire de l'EvalStat généré. Defaut = False.
        """
        # On boucle sur Session pour traiter_eval en spécifiant ecrire_eval_formation=False
        for session in self.sessions.values():
            session.eval.traiter_eval(
                chemin_csv=chemin_csv,
                chemin_IRIS_sessions=chemin_IRIS_sessions,
                ecrire_eval_formation=False,  # On sauvegardera en fin de boucle
                ouvrirDossier=ouvrirDossier
            )
        
        # Sauvegarde de l'eval formation
        self.eval.ecritdf_et_sauve_siModif()

    # =========================
    # === GETTERS / SETTERS ===
    # =========================
    @property
    def trigramme_formation(self) -> str|None:
        """
        Renvoie le trigramme de la formation.
        
        :return: le trigramme de la formation
        :rtype: int
        """
        return self._trigramme_formation
    
    @property
    def sessions(self) -> dict[int, Session]|None:
        """
        Renvoie le dictionnaire des sessions de la formation.
        
        :return: le dictionnaire des sessions de la formation.
        :rtype: dict[int, Session]
        """
        return self._sessions
    
    @property
    def eval(self) -> EvalStat_formation|None:
        return self._eval
    
    @eval.setter
    def eval(self, valeur:EvalStat_formation) -> None:
        self._eval = valeur