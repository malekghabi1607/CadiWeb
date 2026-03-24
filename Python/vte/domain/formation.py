from __future__ import annotations
from pathlib import Path
from typing import Optional
from collections.abc import Iterable

from vte.domain.bilanSessions import BilanSessions
from vte.domain.evalStat import EvalStat_formation
from vte.domain.fdc import FdC
from vte.domain.session import Session
from vte.utils.utils import convertir_collection

# ======================================================================================
# CLASSE FORMATION
# ======================================================================================
class Formation:
    def __init__(self, trigramme_formation: str):
        """
        Initialise une instance de Formation a minima.

        :param trigramme_formation: Le trigramme de la formation.
        :type trigramme_formation: str
        """
        self._trigramme_formation:str = trigramme_formation

        # Une formation a une fiche de coûts
        self._fdc:Optional[FdC] = None
        
        # Une formation a une évaluation stagiaire de la formation (regroupement de toutes les évaluations stagiaires de toutes les sessions)
        self._eval:Optional[EvalStat_formation] = None

        # Une formation a un ou plusieurs bilans de formation (annuel)
        #self.bilans_formation:Optional[dict[int, BilanFormation]] = {}  # bilans_formation[2025] : Index = année du bilan

        # Une formation a un ou plusieurs bilans de session par année (soit 1 par semestre, soit annuel s'il n'y a qu'une session annuellement)
        self.bilans_sessions:Optional[dict[int, dict[int, BilanSessions]]] = {}  #bilans_session[2025][0] : Index1 = année du bilan ; Index2 = période du bilan (1 = 1er semestre ; 2 = 2nd semestre ; 0 = annuel)

        # Une formation a une ou plusieurs sessions
        self._sessions:list[Session] = []

    @classmethod
    def avec_ouverture_ou_creation_evalStat(cls, trigramme_formation: str) -> Formation:
        """
        Crée une instance de Formation en créant un evalStat formation (ou en l'ouvrant s'il existe déjà)

        :param trigramme_formation: Le trigramme de la formation
        :type trigramme_formation: str
        :return: Une instance de formation
        :rtype: Formation
        """
        instance = cls(trigramme_formation=trigramme_formation)
        instance.ouvrir_ou_creer_evalStat()
        return instance

    @classmethod
    def avec_ajout_sessions(cls, trigramme_formation: str, codes_IRIS:int|Iterable[int]) -> Formation:
        """
        Initialise une instance de Formation contenant une ou plusieurs instances de Session.

        :param trigramme_formation: Le trigramme de la formation
        :type trigramme_formation: str
        :param codes_IRIS: Un itérable de codes IRIS
        :type codes_IRIS: int | Iterable[int]
        :return: Une instance de formation
        :rtype: Formation
        """
        instance = cls(trigramme_formation=trigramme_formation)
        instance.ajout_sessions(codes_IRIS=codes_IRIS)
        return instance










    @classmethod
    def avec_creation_session_et_ouverture_ou_traitement_EvalStat(
        cls, 
        trigramme_formation: str, 
        code_IRIS:int, 
        chemin_csv:Optional[Path|str]=None, 
        ecrire_eval_formation:bool=True, 
        ouvrirDossier:bool=False
        ) -> Formation:
        """
        Initialise une instance de Formation contenant une instance de Session.
        
        Crée l'Excel EvalStat d'une session à partir d'un CSV (s'il n'est pas donné, on ouvre un filedialog)
        L'évaluation de la formation est mise à jour avec ces nouvelles données.

        L'évaluation de la formation est sauvée en fin de traitement.

        :param trigramme_formation: Le trigramme de la formation
        :type trigramme_formation: str
        :param code_IRIS: Le code IRIS de la session
        :type code_IRIS: int
        :param chemin_csv: _description_, defaults to None
        :type chemin_csv: Optional[Path | str], optional
        :param ecrire_eval_formation: Pour écrire physiquement l'Excel eval formation en fin de traitement. Si False, il devra être écrit ailleurs (à l'endroit où il y a la boucle pour du multi-traitement typiquement). Défaut = True.
        :type ecrire_eval_formation: bool, optional
        :param ouvrirDossier: Ouvre le répertoire de l'EvalStat généré. Defaut = False.
        :type ouvrirDossier: bool, optional
        :return: Une instance de formation
        :rtype: Formation
        """

        #On crée l'instance avec ouverture de l'EvalStat formation
        instance = cls.avec_ouverture_ou_creation_evalStat(trigramme_formation=trigramme_formation)

        # On ajoute la session 
        instance.ajout_sessions(codes_IRIS=code_IRIS)

        instance.traiter_eval_sessions(
            chemin_csv=chemin_csv,
            ecrire_eval_formation=ecrire_eval_formation,
            ouvrirDossier=ouvrirDossier
        )
        
        return instance



    # =========================
    # === METHODES INTERNES ===
    # =========================
    def _ajout_FdC_avec_ouverture(self, chemin_fdc:Optional[Path|str] = None) -> None:
        """
        Ajoute la fiche de coûts de la formation à l'instance de Formation (i.e. renseigne self._fdc)

        La fiche de coût est ouverte si elle existe.
        """
        self._fdc = FdC.depuis_chemin(formation=self, chemin_fdc=chemin_fdc)



    # =========================
    # === METHODES EXTERNES ===
    # =========================
    def get_session_par_codeIRIS(self, code_IRIS:int) -> Optional[Session]:
        """
        Renvoie la session de la liste self._sessions avec le code_IRIS.

        Renvoie None si non trouvé.

        :param code_IRIS: Code IRIS de la session à retourner
        :type code_IRIS: int
        :return: L'objet Session de self._sessions avec ce code IRIS. None si non trouvé.
        :rtype: Optional[Session]
        """
        return next((session for session in self.sessions if session.code_IRIS == code_IRIS), None)

    def ajout_sessions(self, codes_IRIS:int|Iterable[int]) -> None:
        """
        Ajouter une ou plusieurs sessions au dictionnaire de la formation

        :param codes_IRIS: Codes IRIS à ajouter
        :type codes_IRIS: int | Iterable[int]
        """
        codes = convertir_collection(codes_IRIS)

        for code_IRIS in codes:
            self._sessions.append(Session(formation=self, code_IRIS=code_IRIS))

    def ouvrir_ou_creer_evalStat(self) -> None:
        """
        Définit self._eval en créant ou ouvrant l'évaluation de la formation à l'instance de Formation.

        L'évaluation est créée si elle n'existe pas ou est ouverte si elle existe.
        """
        if self._eval is None:
            self._eval = EvalStat_formation.avec_ouverture_ou_creation(formation=self)


    def traiter_eval_sessions(
        self,
        chemin_csv: Optional[Path|str],
        ecrire_eval_formation: Optional[bool] = False, 
        ouvrirDossier: Optional[bool] = False
        ) -> None:
        """
        Traite les EvalStat de plusieurs sessions. Pour chacune d'elle :

        Crée l'Excel EvalStat d'une session à partir d'un CSV (s'il n'est pas donné, on ouvre un filedialog)

        L'évaluation de la formation est mise à jour avec ces nouvelles données et est sauvée en fin de traitement.

        :param chemin_csv: _description_, defaults to None
        :type chemin_csv: Optional[Path | str], optional
        :param ecrire_eval_formation: Pour écrire physiquement l'Excel eval formation en fin de traitement. Si False, il devra être écrit ailleurs (à l'endroit où il y a la boucle pour du multi-traitement typiquement). Défaut = False car ic on peut traiter plusiseurs sessions d'une même formation.
        :type ecrire_eval_formation: bool, optional
        :param ouvrirDossier: Ouvre le répertoire de l'EvalStat généré. Defaut = False.
        :type ouvrirDossier: bool, optional
        """


        # On boucle sur Session pour traiter_eval en spécifiant ecrire_eval_formation=False
        for session in self.sessions.values():
            session.eval.traiter_eval(
                chemin_csv=chemin_csv,
                ecrire_eval_formation=ecrire_eval_formation,  # On sauvegardera en fin de boucle
                ouvrirDossier=ouvrirDossier
            )
        
        # Sauvegarde de l'eval formation
        self.eval.ecrit_et_sauve_df_siModif()

    # =========================
    # === GETTERS / SETTERS ===
    # =========================
    @property
    def trigramme_formation(self) -> Optional[str]:
        """
        Renvoie le trigramme de la formation.
        
        :return: le trigramme de la formation
        :rtype: int
        """
        return self._trigramme_formation
    
    @property
    def sessions(self) -> list[Session]: #Optional[dict[int, Session]]:
        """
        Renvoie le dictionnaire des sessions de la formation.
        
        :return: le dictionnaire des sessions de la formation.
        :rtype: dict[int, Session]
        """
        return self._sessions
    
    @property
    def tuple_codesIRIS_de_sessions(self) -> tuple[int]:
        """
        Renvoie le tuple des codes IRIS de la liste self._sessions

        :return: le tuple des code_IRIS de la liste self._sessions
        :rtype: tuple[int]
        """
        return (session.code_IRIS for session in self.sessions)

    @property
    def eval(self) -> Optional[EvalStat_formation]:
        """
        Renvoie self._eval (objet EvalStat de la formation).

        Si EvalStat est None (jamais ouvert/créé), alors on l'ouvre/on le crée avec EvalStat_formation.avec_ouverture

        :return: _description_
        :rtype: Optional[EvalStat_formation]
        """
        if self._eval is None:
            self._eval = EvalStat_formation.avec_ouverture_ou_creation(formation=self)
        return self._eval
    
    @eval.setter
    def eval(self, valeur:EvalStat_formation) -> None:
        self._eval = valeur

    @property
    def fdc(self) -> Optional[FdC]:
        return self._fdc
    
    @fdc.setter
    def fdc(self, valeur:FdC) -> None:
        self._fdc = valeur