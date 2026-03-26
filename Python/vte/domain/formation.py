from __future__ import annotations
from pathlib import Path
from typing import Optional
from collections.abc import Iterable

from vte.core.iris_referentiel import get_iris
from vte.domain.bilanSessions import BilanSessions
from vte.domain.evalStat import EvalStat_formation
from vte.domain.fdc import FdC
from vte.domain.session import Session
from vte.utils.utils import convertir_collection, vlog

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

        # Une formation a une évaluation stagiaire de la formation (regroupement de toutes les évaluations stagiaires de toutes les sessions)
        self._eval:Optional[EvalStat_formation] = None

        # Une formation a une fiche de coûts
        self._fdc:Optional[FdC] = None
        
        # Une formation a une ou plusieurs sessions
        self._sessions:list[Session] = []

        # Une formation a un ou plusieurs bilans de session par année (soit 1 par semestre, soit annuel s'il n'y a qu'une session annuellement)
        self._bilans_sessions:dict[int, dict[str, BilanSessions]] = {}  #bilans_session[2025]["Année"] : Index1 = année du bilan ; Index2 = période du bilan (1er semestre ; 2nd semestre ; Annuel)

        # Une formation a un ou plusieurs bilans de formation (annuel)
        #self.bilans_formation:Optional[dict[int, BilanFormation]] = {}  # bilans_formation[2025] : Index = année du bilan

    # TODO : voir si ce n'est pas un constructeur de traitement finalement (déplacer plus bas et renommer)
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
        instance.ouvrir_ou_creer_eval_formation()
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









    # ======================================
    # === CONSTRUCTEURS POUR TRAITEMENTS ===
    # ======================================
    
    # Méthodes de traitement : on fera toutes les étapes intemédiaires si besoin (créer formation, ajouter sessions, ajouter FdC...)

    @classmethod
    def pour_traitement_bilanSessions_depuis_codesIRIS(cls, codes_IRIS:int|Iterable[int], annee:Optional[int]=None, periode:str="Année") -> Formation:
        """
        Génère un bilan de sessions à partir d'un ou plusieurs codes IRIS (un bilan pour une session ou pour plusieurs sessions (période)).

        Si annee est donnée, alors on ne fait pas de contrôle. Sinon on la détermine avec la période grâce aux codes IRIS et à IRIS sessions.

        :param codes_IRIS: Codes IRIS des sessions pour lesquels on souhaite faire le bilan
        :type codes_IRIS: int | Iterable[int]
        :param annee: Année du bilan (s'il n'est pas donné on l'obtiendra d'IRIS sessions)
        :type annee: Optional[int], optional
        :param periode: période du bilan (appartient à ["Année", "1er semestre", "2nd semestre"]), defaut = "Année"
        :type periode: str, optional
        :return: le bilan de sessions est traité et l'objet formation est bien créé avec les sessions correspondantes.
        :rtype: Formation
        """
        # On évalue le trigramme de la formation
        trigramme_formation = get_iris(typeExport="Sessions").get_champ_depuis_codesIRIS(champ="Trigramme formation", codes_IRIS=codes_IRIS, valeurUnique=True)

        # On définit la formation et on ajoute les sessions
        instance = cls.avec_ajout_sessions(trigramme_formation=trigramme_formation, codes_IRIS=codes_IRIS)

        # On ajoute le bilan de sessions au dictionnaire et on le traite
        instance.ajout_bilan_sessions_avec_traitement(
            codes_IRIS=codes_IRIS,
            annee=annee,
            periode=periode
        )
        
        return instance 

    @classmethod
    def pour_traitement_bilanSessions_depuis_periode(cls, trigramme_formation:str, annee:int, periode:str="Année") -> Formation:
        """
        Permet de générer un bilan de sessions selon une année et une période qui est l'un de ces éléments : ["1er semestre", "2nd semestre", "Année"]

        Ex : BilanSessions.bilanUnique_parPeriode("948", 2024, "Année")

        :param annee: Année du bilan
        :type annee: int
        :param periode: Période du bilan (appartient à ["Année", "1er semestre", "2nd semestre"]), defaut = "Année"
        :type periode: str, optional
        :return: le bilan de sessions est traité et l'objet formation est bien créé avec les sessions correspondantes.
        :rtype: Formation
        """
        # On définit la formation 
        instance = cls(trigramme_formation=trigramme_formation)

        # On ajoute le bilan de sessions au dictionnaire et on le traite
        instance.ajout_bilan_sessions_avec_traitement(
            annee=annee,
            periode=periode
        )
        
        # On ajoute les sessions à formation à partir des codes_IRIS ssi on a des codes IRIS qui ont été traités
        if annee in instance._bilans_sessions.keys():
            if periode in instance._bilans_sessions[annee].keys():
                if len(instance._bilans_sessions[annee][periode].codes_IRIS) > 0:
                    instance.ajout_sessions(codes_IRIS=instance._bilans_sessions[annee][periode].codes_IRIS)

        return instance

    
    # TODO : non employé, vérifier besoin et adapter si besoin
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

        instance.ouvrir_ou_traiter_eval_sessions(
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
    def ouvrir_ou_creer_eval_formation(self) -> None:
        """
        Définit self._eval en créant ou ouvrant l'évaluation de la formation à l'instance de Formation.

        L'évaluation est créée si elle n'existe pas ou est ouverte si elle existe.
        """
        if self._eval is None:
            self._eval = EvalStat_formation.avec_ouverture_ou_creation(formation=self)

    def ouvrir_fdc(self, chemin_fdc:Optional[Path|str] = None) -> None:
        """
        Ouvre la fiche de coûts de la formation (i.e. définit self._fdc).

        Si le chemin n'est pas donné, alors on ouvre un filedialog.

        :param chemin_fdc: Chemin de la fiche de coûts. Défaut = None
        :type chemin_fdc: Optional[Path | str], optional
        """
        self._fdc = FdC.depuis_chemin(formation=self, chemin_fdc=chemin_fdc)

    def ajout_sessions(self, codes_IRIS:int|Iterable[int]) -> None:
        """
        Ajouter une ou plusieurs sessions au dictionnaire de la formation.

        On vérifie qu'aucune session avec ce code IRIS n'appartient à cette liste au préalable.

        :param codes_IRIS: Codes IRIS à ajouter
        :type codes_IRIS: int | Iterable[int]
        """
        codes = convertir_collection(codes_IRIS)

        for code_IRIS in codes:
            if code_IRIS not in [session.code_IRIS for session in self._sessions]:
                self._sessions.append(Session(formation=self, code_IRIS=code_IRIS))

    def ouvrir_ou_traiter_eval_sessions(
        self,
        ecrire_eval_formation: Optional[bool] = False, 
        ouvrirDossier: Optional[bool] = False
        ) -> None:
        """
        Ouvre ou traite les EvalStat de toutes les sessions. Pour chacune d'elle :

        Ouvre ou crée l'Excel EvalStat d'une session à partir d'un filedialog

        L'évaluation de la formation est mise à jour avec ces nouvelles données et est sauvée en fin de traitement.

        :param ecrire_eval_formation: Pour écrire physiquement l'Excel eval formation en fin de traitement. Si False, il devra être écrit ailleurs (à l'endroit où il y a la boucle pour du multi-traitement typiquement). Défaut = False car ic on peut traiter plusiseurs sessions d'une même formation.
        :type ecrire_eval_formation: bool, optional
        :param ouvrirDossier: Ouvre le répertoire de l'EvalStat généré. Defaut = False.
        :type ouvrirDossier: bool, optional
        """


        # On boucle sur Session pour traiter_eval en spécifiant ecrire_eval_formation=False
        for session in self.sessions:
            session.eval.ouvrir_ou_traiter_eval(
                ecrire_eval_formation=ecrire_eval_formation,  # On sauvegardera en fin de boucle
                ouvrirDossier=ouvrirDossier
            )
        
        # Sauvegarde de l'eval formation
        self.eval.ecrit_et_sauve_df_siModif()

    def ajout_bilan_sessions_avec_traitement(self, codes_IRIS:Optional[int|Iterable[int]]=None, annee:Optional[int]=None, periode:str="Année") -> None:
        """
        On ajoute un nouveau bilan de sessions à self._bilans_sessions[annee][periode].

        On traite ce bilan de sessions (création du word et envoi du mail pour signature).

        Si codes_IRIS fourni, alors on traite avec BilanSessions.depuis_codesIRIS (et dans ce cas, si annee est donnée alors on ne fait pas de contrôle ; sinon on la détermine avec la période grâce aux codes IRIS et à IRIS sessions).
        Sinin il faut fournir annee et on traite avec BilanSessions.depuis_periode.

        :param codes_IRIS: Codes IRIS des sessions pour lesquels on souhaite faire le bilan. Optionnel si annee fourni.
        :type codes_IRIS: Optional[int|Iterable[int]], optional
        :param annee: Année du bilan. Ooptionnel si codes_IRIS fourni
        :type annee: Optional[int], optional
        :param periode: période du bilan (appartient à ["Année", "1er semestre", "2nd semestre"]), defaut = "Année"
        :type periode: str, optional
        """
        # Vérif qu'on peut employer soit BilanSessions.depuis_codesIRIS soit BilanSessions.depuis_periode
        if (codes_IRIS is None) and (annee is None) :
            vlog.log_erreur("code_IRIS et annee sont None tous les deux : on ne peut pas lancer le traitement de BilanSessions")
        
        # On traite le bilan
        if codes_IRIS is not None:
            # Alors on traite à partir des codes IRIS
            bilan = BilanSessions.depuis_codesIRIS(formation=self, codes_IRIS=codes_IRIS, annee=annee, periode=periode)
        else:
            # Alors on traite à partir de la période
            bilan = BilanSessions.depuis_periode(formation=self, annee=annee, periode=periode)

        # Les champs année et période peuvent être complétés/définis lors du traitement de BilanSessions, donc je ne peut affcter self._bilans_sessions[bilan.annee][bilan.periode] que maintenant
        # On initialise le dictionnaire de 2nd niveau si non déjà fait
        if bilan.annee not in self._bilans_sessions.keys():
            self._bilans_sessions[bilan.annee] = {}
        
        # On ajoute l'élément
        self._bilans_sessions[bilan.annee][bilan.periode] = bilan

    # ====================
    # === METHODES GET ===
    # ====================
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