from __future__ import annotations
from abc import ABC, abstractmethod
from pathlib import Path
from pprint import pprint
from typing import Optional, Protocol
from functools import cached_property  # Décorateur générique pour mettre en cache des données lourdes que je recalculais pleinde fois en @property

from pandas import DataFrame
from mailmerge import MailMerge

from vte.core import config
from vte.core.iris_referentiel import get_iris
from vte.domain.evalStat import EvalStat_formation
from vte.domain.session import Session
from vte.domain.iris import IRIS, IRIS_traite
from vte.services.evalStat_services import EvalStat_services
from vte.utils.office import FichierExcel, FichierWord, Mail
from vte.utils.utils import *

# TODO : Pour l'instant c'est une classe de traitement. Le jour où j'ai besoin d'ouvrir un BilanSession pour le lire uniquement, prendre modèle sur IRIS avec des classes de lecture et de traitement

# Test à faire sur TEL période : 2024

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

    @property
    def eval(self) -> EvalStat_formation|None: ...
    
    def get_session_par_codeIRIS(self, code_IRIS:int) -> Optional[Session]: ...

    #def ajout_sessions(self, codes_IRIS:int|Iterable[int]) -> None: ...

# ======================================================================================
# CLASSE BilanSessions
# ======================================================================================

class BilanSessions:
    """
    C'est la classe qui contient tous les éléments de ma formation pour mon bilan de sessions V3
    """
    # === VARIABLES PARTAGÉES ENTRE TOUTES LES INSTANCES


    _CRITERES_A_ENLEVER:list[str] = [  # Critères à ne pas retenir pour le calcul des moyennes < 3
        "Comment avez-vous connu cette formation ?", "Avez-vous d'autres besoins de formation ?", "Commentaires, remarques, suggestions", "Recommanderiez-vous cette formation ?"]

    # Dictionnaire pour mapper les statuts aux clés de self._statuts ["Exploités pour les évaluations (CSV présents)", "Exploités pour les évaluations (CSV présents)", "Exclus des évaluations (problème traitement CSV)", "Exclus des évaluations (CSV manquants)", "Exclus des évaluations (CSV vide / aucun retour)", "Exclus entièrement du bilan (non présent dans IRIS / mauvais code)"]
    _mapping_statuts:dict[str, str] = {
        "Traité": "Exploités pour les évaluations (CSV présents et non vides)",                                  # Exploités pour stats initiales → Dans _demande_sessions_a_exclure
        #  "Traité - CSV déjà dans fichier global": "Exploités pour les évaluations (CSV présents et non vides)",   # Exploités pour les stats stagiaires → Dans _maj_evalstat_formation
        "Exclu - Aucun CSV fourni": "Exclus des évaluations (CSV manquant)",                   # Exclus des évaluations car problème au traitement des CSV → Dans _maj_evalstat_formation
        "Exclu - Code IRIS pas dans Extract IRIS sessions": "Exclus entièrement du bilan (non présent dans IRIS / mauvais code)",  # Exclus entièrement du bilan car sessions non réalisées ou mauvais RP (exclus par l'utilisateur) → Dans _maj_evalstat_formation
        "Exclu - Problème lecture CSV": "Exclus des évaluations (problème traitement CSV)",         # Exclus des évaluations car CSV stagiaires manquants → Dans _maj_evalstat_formation
        "Exclu - CSV vide / Aucun retour": "Exclus des évaluations (CSV vide / aucun retour)",      # Exclus des évaluations car le CSV est vide (i.e. aucun retour d'utilisateur)
    }

    # Dictionnaire pour mapper les statuts qui nécessitent de supprimer le code IRIS des stats générales
    _statut_exclus_entierement:tuple[str] = (
        "Exclu - Code IRIS pas dans Extract IRIS sessions",
    )

    # =====================
    # === CONSTRUCTEURS ===
    # =====================
    def __init__(self, formation:Formation_protocol, annee:int, periode:str="Année", codes_IRIS:Optional[int|Iterable[int]]=[]) -> None:
        """
        Initialise un bilan de sessions a minima
        
        Un bilan de sessions est accolé à une formation car on peut avoir plusieurs sessions dans un bilan de sessions.

        :param formation: objet Formation associé à ce bilan
        :type formation: Formation_protocol
        :param annee: Année du bilan
        :type annee: int
        :param periode: période du bilan (appartient à ["Année", "1er semestre", "2nd semestre"]), defaut = "Année"
        :type periode: str, optional
        :param codes_IRIS: Codes IRIS des sessions pour lesquels on souhaite faire le bilan. Défaut = []
        :type codes_IRIS: Optional[int|Iterable[int]], optional
        """
        # --- Variables qui caractérisent le bilan de sessions
        self._formation = formation
        self._annee: int = annee
        self._periode: str = periode  # ["Année", "1er semestre", "2nd semestre"]
        self._periode_pour_titre:str = ""  # f"Session {numSession} uniquement ({moisSession} {instance._annee})", f"{self._periode} {self._annee}"
        self._codes_IRIS:list[int] = []
        self.codes_IRIS = codes_IRIS

        # --- Variables de traitement ---
        self._statuts:dict[str, list] = {clef: [] for clef in self._mapping_statuts.values()}  # Dictionnaire qui liste les codes IRIS selon chaque statut
        self._stats_stagiaires: dict[str, dict[str, int|float|str|None]] = {}  # Dictionnaire des stats des CSV
        """
        dictionnaire de la forme :
        {
            "Nom du critère": {
                "Nombre": ...,
                "Moyenne": ...,
                "Commentaires": ...
            },
            ...
        }
        """

    @classmethod   
    def depuis_codesIRIS(cls, formation:Formation_protocol, codes_IRIS:int|Iterable[int], annee:Optional[int]=None, periode:str="Année") -> BilanSessions:
        """
        Génère un bilan de sessions à partir d'un ou plusieurs codes IRIS (un bilan pour une session ou pour plusieurs sessions (période)).

        Si annee est donnée, alors on ne fait pas de contrôle. Sinon on la détermine avec la période grâce aux codes IRIS et à IRIS sessions.

        :param formation: Objet formation associé à ce bilan
        :type formation: Formation_protocol
        :param codes_IRIS: Codes IRIS des sessions pour lesquels on souhaite faire le bilan
        :type codes_IRIS: int|Iterable[int]
        :param annee: Année du bilan (s'il n'est pas donné on l'obtiendra d'IRIS sessions)
        :type annee: Optional[int], optional
        :param periode: période du bilan (appartient à ["Année", "1er semestre", "2nd semestre"]), defaut = "Année"
        :type periode: str, optional
        :return: Un objet bilan de sessions de ce ou ces code(s) IRIS
        :rtype: BilanSessions
        """
        # On évalue l'année et la période à partir des codes IRIS si besoin (sinon aucune vérification : on fait confiance à l'utilisateur)
        if annee is None:
            annee, periode = get_iris(typeExport="Sessions").get_periode_depuis_codes_IRIS(codes_IRIS=codes_IRIS)
        
        # On initialise l'instance 
        instance = cls(formation=formation, annee=annee, periode=periode, codes_IRIS=codes_IRIS)
        #instance.codes_IRIS = codes_IRIS

        # Si l'année n'est pas donnée, je la récupère d'IRIS sessions
        #instance._annee = annee if annee is not None else int(instance.df_sessions_filtre_codesIRIS["Année début ses."].iloc[0])

        # === ON FAIT LES VERIFICATIONS QUI ANNULERAIENT LE TRAITEMENT ===
        continuer = instance.verifier_traitement_bilan()
        if not continuer:
            return

        # === TRAITEMENT DU BILAN DE SESSIONS ===
        instance._traiter()

        return instance

    @classmethod   
    def depuis_periode(cls, formation:Formation_protocol, annee:int, periode:str="Année") -> Optional[BilanSessions]:
        """
        Permet de générer un bilan de sessions selon une année et une période qui est l'un de ces éléments : ["1er semestre", "2nd semestre", "Année"]

        Ex : BilanSessions.bilanUnique_parPeriode("948", 2024, "Année")

        :param formation: Objet formation associé à ce bilan
        :type formation: Formation_protocol
        :param annee: Année du bilan
        :type annee: int
        :param periode: Période du bilan (appartient à ["Année", "1er semestre", "2nd semestre"]), defaut = "Année"
        :type periode: str, optional
        :return: Un objet bilan de sessions des codes IRIS sélectionnés par l'utilisateur
        :rtype: Optional[BilanSessions]
        """
        # On demande à l'utilisateur les sessions qu'il souhaite exclure
        codes_IRIS = get_iris(typeExport="Sessions").demande_sessions_a_retenir(
            trigramme_formation=formation.trigramme_formation,
            annee=annee,
            periode=periode
        )

        # On traite le bilan de session
        if len(codes_IRIS) > 0:  # Si on a des codes IRIS à traiter, on traite le bilan de sessions
            return BilanSessions.depuis_codesIRIS(
                formation=formation,
                codes_IRIS=codes_IRIS,
                annee=annee,
                periode=periode
            )
        else:  # Si aucun code IRIS n'a été sélectionné, on print un warning
            vlog.print("Info", f"⚠️  Toutes les sessions sont exclues : il n'y a plus de raison de faire le bilan de sessions.")
            return None








    # =========================
    # === METHODES INTERNES === 
    # =========================

    # === Méthodes get ===
    def _get_codesIRIS_par_statut_evalStat(self, statut_evalStat: str) -> list[int]:
        """
        Retourne une liste des codes IRIS pour un statut donné.

        :param statut_evalStat: statut EvalStat (["Traité", "Traité - Code IRIS déjà dans l'évaluation de la formation", "Traité - CSV déjà dans l'évaluation de la formation", "Exclu - Aucun CSV fourni", "Exclu - Code IRIS pas dans Extract IRIS sessions", "Exclu - Problème lecture CSV", "Exclu - CSV vide / Aucun retour"])
        :type statut_evalStat: str
        :return: une liste des codes IRIS avec ce statut
        :rtype: list[int]
        """
        codes_IRIS = []
        for code_IRIS in self._codes_IRIS:
            session = self._formation.get_session_par_codeIRIS(code_IRIS)
            if session.eval.statut == statut_evalStat:
                codes_IRIS.append(code_IRIS)
        return codes_IRIS

    def _get_codesIRIS_par_statut_pourBilan(self, statut_pourBilan: str) -> list[int]:
        """
        Retourne une liste des codes IRIS pour une clé de mapping donnée (statut pour bilan, i.e. mieux nommés).

        :param statut_pourBilan: Statut "pour bilan" (i.e. mieux nommés) ["Exploités pour les évaluations (CSV présents)", "Exploités pour les évaluations (CSV présents)", "Exclus des évaluations (problème traitement CSV)", "Exclus des évaluations (CSV manquants)", "Exclus des évaluations (CSV vide / aucun retour)", "Exclus entièrement du bilan (non présent dans IRIS / mauvais code)"] 
        :type statut_pourBilan: str
        :return: une liste des codes IRIS avec ce statut
        :rtype: list[int]
        """
        #pprint(self._mapping_statuts)

        statut_evalStat = next((clef for clef, valeur in self._mapping_statuts.items() if valeur == statut_pourBilan), None)
        if statut_evalStat is None:
            return []
        return self._get_codesIRIS_par_statut_evalStat(statut_evalStat)

    def _get_generateur_word(self, version: str) -> BilanSessions_generateur_word:
        """
        Récupère la bonne version du générateur Word (si jamais on a plusieurs versions)

        :param version: Version du bilan de sessions à employer
        :type version: str
        :raises ValueError: Si problème
        :return: le générateur du bilan de sessions
        :rtype: BilanSessions_generateur_word
        """
        if version == "V3":
            return Bilan_V3(self)
        else:
            raise ValueError(f"Version inconnue : {version}")


    # === Pipeline métier ===
    def _traiter(self) -> None:
        """
        Traite le bilan de sessions :
            - mise à jour des evalStat ;
            - calcul des stats des csv (moyennes, concaténation commentaires...) ;
            - création du word (calcul et fusion des champs de fusion) ;
            - ouverture word (pour adaptations par utilisateur) ;
            - préparation mail au n+1.
        """
        # On met à jour l'Excel evalstat de la formation si la session demandée par l'utilisateur ne s'y trouve pas
        self._maj_evalstat()

        # On calcule les stats
        self._calculer_stats_criteres()
        #pprint(instance._stats_stagiaires)

        # On construit le bilan de sessions
        self._construire_word(version="V3")

        # On ouvre le word
        FichierWord.depuisFichier(chemin_fichier=self.chemin_word_bilan_output, charger_contentControl=False, afficherWord=True)

        # On envoie un mail au chef d'unité pour la signature du pdf
        self._envoyer_mail_chef_unite()


    # === Détail pipeline ===
    def _maj_evalstat(self) -> None:
        """
        Crée ou ouvre les EvalStat de la/les sessions demandées et met à jour le fichier EvalStat de la formation
        Ne s'applique que si des sessions demandées par l'utilisateur ne s'y trouvent pas.
        (on regarde les CSV qui ne sont pas dans le fichier Excel global à partir de la liste df_sessions_filtre['Code IRIS'])
        """
        # On ouvre ou on traite les EvalStats non déjà créés
        EvalStat_services.ouvrir_ou_traiter_evalStat_depuis_liste_codes_IRIS(codes_IRIS=self._codes_IRIS, formation=self._formation)

    def _calculer_stats_criteres(self) -> None: # dict[str, dict[str, int|float|str|None]]:
        """
        Calcule les statistiques (nombre de retours, moyenne retours, agrégation des commentaires) de tous les critères.

        Fait ce traitement pour tous les éléments dont nous avons des CSV (i.e. appartenant à liste_codesIRIS_pour_statsCSV)
        
        Retourne un dictionnaire de la forme :
        {
            "Nom du critère": {
                "Nombre": ...,
                "Moyenne": ...,
                "Commentaires": ...
            },
            ...
        }

        :return: Un dictionnaire de tous les critères
        :rtype: dict
        """

        # S'il n'y a pas de CSV disponibles pour les stats, alors ce n'est pas la peine de faire les stats
        if len(self.liste_codesIRIS_pour_statsCSV) > 0 :
            # On récupère la iste des critères
            liste_criteres = self.df_stagiaires_filtre_statsCSV['Critère'].dropna().unique()


            # On fait les stats pour chaque critère
            for critere in liste_criteres:
                # Dataframe filtré sur ce critère
                df_filtre = self.df_stagiaires_filtre_statsCSV[self.df_stagiaires_filtre_statsCSV['Critère'] == critere]
                
                # Nombre d'éléments avec ce critère
                nb = len(df_filtre)
                # Moyenne de ce critère
                moyenne = float(df_filtre['Note'].mean()) if nb > 0 else None

                # On concatère les commentaires associés
                commentaires_concat = "\n".join(
                    "• " + c.strip()
                    for c in df_filtre['Commentaires'].dropna().astype(str)
                    if c.strip() != ""
                )

                # On met toutes ces données en forme dans _stats_stagiaires
                self._stats_stagiaires[critere] = {
                    "Nombre": nb,
                    "Moyenne": moyenne,
                    "Commentaires": commentaires_concat
                }
        #else:
            #vlog.print("Info", f"⚠️  Aucun CSV disponible pour ce bilan : les statistiques des stagiaires ne seront pas évaluées.")
            #self._commentairesBilan += f"\n⚠️  Aucun CSV disponible pour ce bilan : les statistiques des stagiaires ne seront pas évaluées.\n"

        #return self._stats_stagiaires

    def _construire_word(self, version: str = "V3") -> None:
        """
        Construit le Word (construit/calcule les champs puis les fusionne dans le Word)

        :param version: Version du bilan de sessions. Défaut = "V3"
        :type version: str, optional
        """
        builder = self._get_generateur_word(version)
        builder.construire()

    def _envoyer_mail_chef_unite(self, pj:Optional[list[str]] = None) -> None:
        """
        Envoie un mail au chef d'unité avec en lien le PDF à signer
        """       
        


        chemin_pdf_bilan_output = self.chemin_word_bilan_output.with_suffix(".pdf")
        corps_html = remplacer_champs(config.CORPS_MAIL_CHEF_UNITE, [
            ["lien_pdf_bilan", chemin_pdf_bilan_output],
            ["formation", f"{self.intitule_formation} ({self.trigramme_formation})"],
            ["periode", self._periode_pour_titre],
        ])

        Mail.creer_mail(
            destinataires=config.ADRESSE_MAIL_CHEF_UNITE,
            sujet=f"Signature bilan de sessions {self.intitule_formation} ({self.trigramme_formation}) : {chemin_pdf_bilan_output.name}",
            corps_html=corps_html,
            pieces_jointes=pj,
            envoyer_mail=False  # envoie directement sans afficher
        ) 



    # =========================
    # === METHODES EXTERNES ===
    # =========================
    def verifier_traitement_bilan(self) -> bool:
        """
        Vérifie si l'on doit traiter un bilan de sessions à partir :
            - le bilan Word n'existe pas déjà ;
            - des codes IRIS (bien à 5 caractères + bien présent dans IRIS sessions).

        Pour l'instant on coupe le programme en cas d'échec d'une vérification (la sortie bool=False ne sert à rien car non employée)

        :return: Un bool pour dire si le traitement doit continuer :
            - True on peut traiter l'EvalStat (aucun souci détecté) ; 
            - False il ne faut pas traiter l'EvalStat (code IRIS ou CSV non détectés dans EvalStat + code IRIS présent dans IRIS sessions).
        :rtype: bool
        """
        # Vérification 1 : on vérifie la pré-existance du bilan Word ; si il existe déjà, alors on arrête le traitement
        continuer = verifier_existance_fichier(self.chemin_word_bilan_output)
        if not continuer:
            vlog.log_erreur("❌  Bilan de sessions déjà existant → Arrêt du traitement du bilan par l'utilisateur")
            
        
        
        # Vérifications 2 : liés à code_IRIS
        for code_IRIS in self._codes_IRIS:
            if code_IRIS is not None:
                # 2.1 : On vérifie que code_iris est bien un entier à 5 chiffres
                est_code_IRIS_valide, _ = IRIS.verifier_code_IRIS(code_IRIS, int)
                if not est_code_IRIS_valide:
                    vlog.log_erreur(f"❌  Code IRIS renseigné non valide : {code_IRIS}")
                    #statut = "Exclu - Code IRIS renseigné non valide"
                    #return False, statut


                # 2.2 : On vérifie si le code_IRIS est bien existant dans l'extract IRIS sessions
                if code_IRIS not in self.df_sessions_filtre_periode["Code IRIS"].values:
                    vlog.log_erreur(f"❌  Code IRIS {code_IRIS} non trouvé dans l’extract IRIS.")
                    #print(f"⚠️  Code IRIS {code_IRIS} non trouvé dans l’extract IRIS.")
                    #statut = "Exclu - Code IRIS pas dans Extract IRIS sessions"
                    #return False, statut
 

        # Si rien n'a arrêté les vérifications, alors tout est OK
        return True



    # =========================
    # === GETTERS / SETTERS === 
    # =========================
    # Getters / setters liés à BilanSessions
    @property
    def annee(self) -> int:
        return self._annee
    
    @property
    def periode(self) -> int:
        return self._periode
    
    @property
    def chemin_word_bilan_output(self) -> Path:
        """
        Renvoie le chemin de sortie du bilan de formation.
        Cette donnée est stockée dans le fichier de config (valeur par défaut).
        
        :return: Chemin de sortie du bilan de formation.
        :rtype: Path
        """
        return config.format_path(config.CHEMIN_WORD_BILAN_SESSION_OUTPUT, trigramme_formation=self.trigramme_formation, annee=self._annee, periode=self.periode_pour_titre, unite=config.UNITE)

    @property
    def periode_pour_titre(self) -> str:
        """
        Période de la session dans le cadre d'une session unique (len(codes_IRIS)=1).

        Construction :
            - si un seul code IRIS : f"Session {self.numero_session} uniquement ({self.mois_session} {self._annee})"
            - si plusieurs codes IRIS : f"{self._periode} {self._annee}"
        """
        if self._periode_pour_titre == "":
            if len(self._codes_IRIS) == 1:  # Cas code IRIS unique
                self._periode_pour_titre = f"{self.numero_session} ({self.mois_session} {self._annee})"
            elif len(self._codes_IRIS) > 1:
                self._periode_pour_titre = f"{self._periode} {self._annee}"
            else:
                vlog.log_erreur("J'appelle periodeSessionsEvaluees alors que len(self.codes_IRIS)<=0")
        
        return self._periode_pour_titre

    @property
    def codes_IRIS(self) -> Iterable[int]:
        return self._codes_IRIS
    
    @codes_IRIS.setter
    def codes_IRIS(self, valeur:int|Iterable[int]) -> None:
        if isinstance(valeur, int):
            self._codes_IRIS = [valeur]
        elif isinstance(valeur, Iterable) and not isinstance(valeur, str):  # Iterable[int] n’est pas valide dans isinstance
            self._codes_IRIS = valeur
        else:
            vlog.log_erreur("La valeur n'est ni un int ni un Iterable de int (codes_IRIS.setter)")

    @property
    def code_IRIS(self) -> int:
        """
        Le 1er code IRIS (code_IRIS au singulier donc on considère qu'on est sur un bilan contenant un seul code IRIS unique)
        """
        return self._codes_IRIS[0]

    @cached_property
    def liste_codesIRIS_exclus_totalement(self) -> list[int]:
        """
        Retourne la liste des codes IRIS qui seront complètement exclus du bilan de sessions (en-tête et CSV).

        Ca correspond aux codes IRIS qui ont été exclus par l'utilisateur ou qui ne sont pas dans IRIS sessions.

        :return: La liste des codes IRIS qui seront complètement exclus du bilan de sessions (en-tête et CSV).
        :rtype: list[int]
        """
        return [
            code_IRIS for code_IRIS in self._codes_IRIS
            if self._formation.get_session_par_codeIRIS(code_IRIS).eval.statut in self._statut_exclus_entierement
        ]

    @cached_property
    def liste_codesIRIS_pour_enTete(self) -> list[int]:
        """
        Retourne la liste des codes IRIS qui seront dans l'en-tête du bilan de sessions.

        Ca correspond aux codes IRIS qui n'ont pas été exclus par l'utilisateur et qui sont dans IRIS sessions.

        :return: La liste des codes IRIS qui seront dans l'en-tête du bilan de sessions.
        :rtype: list[int]
        """
        return [
            code_IRIS for code_IRIS in self._codes_IRIS
            if code_IRIS not in self.liste_codesIRIS_exclus_totalement
        ]
    
    @cached_property
    def liste_codesIRIS_pour_statsCSV(self) -> list[int]:
        """
        Retourne la liste des codes IRIS pour lesquels on peut calculer les stats à partir des CSV.

        Ca correspond aux codes IRIS pour lesquels le CSV est fonctionnel et non vide.

        :return: La liste des codes IRIS pour lesquels on peut calculer les stats à partir des CSV.
        :rtype: list[int]
        """
        return self._get_codesIRIS_par_statut_pourBilan("Exploités pour les évaluations (CSV présents et non vides)")

    @cached_property
    def liste_codesIRIS_avec_pb_CSV(self) -> list[int]:
        """
        Retourne la liste des codes IRIS pour lesquels il y a un problème de CSV.

        Ca correspond aux codes IRIS de l'en-tête "moins" les codes IRIS pour lesquels on a des CSV fonctionnels.

        :return: La liste des codes IRIS pour lesquels il y a un problème de CSV.
        :rtype: list[int]
        """
        return [
            code_IRIS for code_IRIS in self.liste_codesIRIS_pour_enTete
            if code_IRIS not in self.liste_codesIRIS_pour_statsCSV
        ]



    # Liens avec Formation
    @property
    def trigramme_formation(self) -> str:
        return self._formation.trigramme_formation
    
    @property
    def eval_formation(self) -> Optional[EvalStat_formation]:
        return self._formation.eval

    @property
    def eval_fe(self) -> Optional[FichierExcel]:
        return self._formation.eval.fe
     
    @cached_property
    def df_stagiaires(self) -> Optional[DataFrame]:
        return self.eval_formation.fe.get_df_tableau("Stagiaires")
    
    @cached_property
    def df_stagiaires_filtre_statsCSV(self) -> DataFrame:
        """
        df_stagiaires (eval formation) filtré sur les codes IRIS pour lesquels on peut calculer les stats à partir des CSV.

        Ca correspond aux codes IRIS pour lesquels le CSV est fonctionnel et non vide.

        :return: df_stagiaires (eval formation) filtré sur les codes IRIS pour lesquels on peut calculer les stats à partir des CSV.
        :rtype: DataFrame
        """
        return self.df_stagiaires[self.df_stagiaires['Code IRIS'].isin(self.liste_codesIRIS_pour_statsCSV)]
    
    @cached_property
    def df_stagiaires_filtre_statsCSV_1ligne_par_session(self) -> DataFrame:
        return self.df_stagiaires_filtre_statsCSV.drop_duplicates(subset=['N° Session'])



    # Liens avec IRIS 
    @cached_property
    def iris_sessions(self) -> IRIS_traite:
        return get_iris(typeExport="Sessions")
  
    @cached_property
    def intitule_formation(self) -> str:
        """
        Intitulé de la formation (prend le nom de la dernière ligne pour avoir la dernière mise à jour)
        """
        return self.df_sessions_filtre_periode["Session"].iloc[-1]

    @cached_property
    def mois_session(self) -> str:
        """
        Mois de la session dans le cadre d'une session unique (len(codes_IRIS)=1)
        """
        return mois_fr_depuis_date(self.df_sessions_filtre_codesIRIS["Date début ses."].iloc[0])

    @cached_property
    def numero_session(self) -> str:
        """
        Numéro de la session dans le cadre d'une session unique (len(codes_IRIS)=1)
        """
        return self.df_sessions_filtre_codesIRIS["N° Session"].iloc[0]


    @cached_property
    def df_sessions_filtre_periode(self) -> DataFrame:
        """
        Renvoie le dataframe de l'extract IRIS filtré VTE selon plusieurs critères :
            - Statut Session != "Annulée" (toujours) ;
            - Nb. Nommés != 0 (toujours) ;
            - l'année de la session (année n du bilan de formation) ;
            - le trigramme de la formation traité par le bilan (self.trigramme_formation).
        
        :return: extract IRIS VTE filtré
        :rtype: DataFrame
        """
        return self.iris_sessions.df_filtre_periode(trigramme_formation=self.trigramme_formation, annee=self._annee, periode=self._periode)

    @cached_property
    def df_sessions_filtre_codesIRIS(self) -> DataFrame:
        """
        Renvoie le dataframe de l'extract IRIS filtré VTE selon plusieurs critères :
            - Code IRIS in self._codes_IRIS
            - Statut Session != "Annulée" (toujours) ;
            - Nb. Nommés != 0 (toujours) ;
            - l'année de la session (année n du bilan de formation) ;
            - le trigramme de la formation traité par le bilan (self.trigramme_formation).
        
        :return: extract IRIS VTE filtré sur codes_IRIS
        :rtype: DataFrame
        """     
        return self.df_sessions_filtre_periode[self.df_sessions_filtre_periode['Code IRIS'].isin(self._codes_IRIS)]

    @cached_property
    def df_sessions_filtre_exclus_totalement(self) -> DataFrame: #_df_stagiaires_final → df_sessions_filtre_stats_generales
        """
        Renvoie le dataframe de l'extract IRIS filtré sur les codes IRIS qui seront complètement exclus du bilan de sessions (en-tête et CSV).

        Ca correspond aux codes IRIS qui ont été exclus par l'utilisateur ou qui ne sont pas dans IRIS sessions.

        :return: dataframe de l'extract IRIS filtré sur les codes IRIS qui seront complètement exclus du bilan de sessions (en-tête et CSV).
        :rtype: DataFrame
        """
        return self.df_sessions_filtre_codesIRIS[self.df_sessions_filtre_codesIRIS['Code IRIS'].isin(self.liste_codesIRIS_pour_enTete)]

    @cached_property
    def df_sessions_filtre_enTete(self) -> DataFrame: #_df_stagiaires_final → df_sessions_filtre_stats_generales
        """
        Renvoie le dataframe de l'extract IRIS filtré sur les codes IRIS exploités pour le bilan (en-tête)

        Ca correspond aux codes IRIS qui n'ont pas été exclus par l'utilisateur et qui sont dans IRIS sessions.

        :return: dataframe de l'extract IRIS filtré sur les codes IRIS exploités pour le bilan (en-tête)
        :rtype: DataFrame
        """
        return self.df_sessions_filtre_codesIRIS[self.df_sessions_filtre_codesIRIS['Code IRIS'].isin(self.liste_codesIRIS_pour_enTete)]

    @cached_property
    def df_sessions_filtre_statsCSV(self) -> DataFrame:
        """
        dataframe de l'extract IRIS filtré sur les codes IRIS pour lesquels on peut calculer les stats à partir des CSV.

        Ca correspond aux codes IRIS pour lesquels le CSV est fonctionnel et non vide.

        :return: dataframe de l'extract IRIS filtré sur les codes IRIS pour lesquels on peut calculer les stats à partir des CSV.
        :rtype: DataFrame
        """
        return self.df_sessions_filtre_codesIRIS[self.df_sessions_filtre_codesIRIS['Code IRIS'].isin(self.liste_codesIRIS_pour_statsCSV)]

    @cached_property
    def df_sessions_filtre_avec_pb_CSV(self) -> DataFrame:
        """
        dataframe de l'extract IRIS filtré sur les codes IRIS pour lesquels il y a un problème de CSV.

        Ca correspond aux codes IRIS de l'en-tête "moins" les codes IRIS pour lesquels on a des CSV fonctionnels.

        :return: dataframe de l'extract IRIS filtré sur les codes IRIS pour lesquels il y a un problème de CSV.
        :rtype: DataFrame
        """
        return self.df_sessions_filtre_codesIRIS[self.df_sessions_filtre_codesIRIS['Code IRIS'].isin(self.liste_codesIRIS_avec_pb_CSV)]






# ======================================================================================
# CLASSE BilanSessions_generateur_word
# ======================================================================================
class BilanSessions_generateur_word(ABC):
    """
    Classe abstraite pour construire un bilan (toutes versions confondues)

    NE S'APPELLE PAS DIRECTEMENT : UNIQUEMENT VIA SES FILLES
    """

    CHEMIN_MODELE: Path = None  # à surcharger

    # =====================
    # === CONSTRUCTEUR ===
    # =====================
    def __init__(self, bilanSessions:BilanSessions):
        self._bilanSessions:BilanSessions = bilanSessions
        self._champs: dict[str, str] = {}

    # =========================
    # === METHODES EXTERNES ===
    # =========================
    def construire(self):
        """
        Méthode principale appelée par BilanSessions.

        Calcule/construit les champs puis le fusionne dans le Word
        """
        self._construire_champs()
        self._fusionner_word()
        #self._post_traitement() # → Dans BilanSessions

    # =========================
    # === METHODES INTERNES === 
    # =========================
    @abstractmethod
    def _construire_champs(self) -> None:
        """
        Remplit self._champs
        Doit être surchargée dans les classes filles héritées
        """
        pass

    def _fusionner_word(self) -> None:
        """
        Tout le process pour écrire les champs de fusion dans le bilan de sessions Word :
            - ouvrir le word à partir du modèle ;
            - fusionne les champs de fusion ;
            - crée le répertoire pour le bilan si besoin ;
            - sauvegarde le word.
        """
        # Ouvre le document Word à partir du modèle
        document = MailMerge(self.CHEMIN_MODELE)
        #print(document.get_merge_fields())

        # Fusionne les champs de fusion
        document.merge(**self._champs)

        # Chemin de sauvegarde
        chemin = self._bilanSessions.chemin_word_bilan_output

        # On crée le répertoire pour les bilans de session de cette année s'il n'existe pas
        chemin.parent.mkdir(parents=True, exist_ok=True)

        # On écrit le fichier
        document.write(chemin)

    """
    def _post_traitement(self) -> None:
        "
        Peut être surchargé (ouvrir word, mail…)
        "
        # Afficher le word
        FichierWord.depuisFichier(
            chemin_fichier=self._BilanSessions.chemin_word_bilan_output,
            charger_contentControl=False,
            afficherWord=True
        )

        # Préparer le mail pour le chef d'unité
        self._BilanSessions._envoyer_mail_chef_unite()
    """




# ======================================================================================
# PROTOCOLES
# ======================================================================================
class Bilan_V3(BilanSessions_generateur_word):

    CHEMIN_MODELE = config.CHEMIN_MODELE_WORD_BILAN_SESSION

    # =======================
    # === PIPELINE METIER === 
    # =======================
    def _construire_champs(self) -> None:
        """
        Pipeline pour calculer et construire les champs de fusion du Word
        """
        # === Données transverses ===
        # Nombre de stagiaires qui ont formulé des retours (provient de eval formation)
        self._nb_stagiaires_retours = self._bilanSessions.df_stagiaires_filtre_statsCSV['NOM Prénom'].nunique()
        # Nombre d'apprenants sur les sessions dont on peut faire les stats CSV (peut provenir de IRIS session ou de eval formation, on prend de df_stagiaire)
        self._nb_apprenants = int(self._bilanSessions.df_stagiaires_filtre_statsCSV_1ligne_par_session['Nb présents'].sum())

        # === On génère le word ===
        self._construire_entete()
        self._construire_commentaires()
        self._construire_stats()


    # =============================
    # === SOUS-PARTIES DU BILAN === 
    # =============================
    def _construire_entete(self):
        """
        Construit la partie en-tête du bilan de sessions
        """
        # On met ici toutes les sessions de la période non excclues par l'utilisateur et qui est dans IRIS sessions

        self._champs["titreFormation"] = self._bilanSessions.intitule_formation
        self._champs["codeFormation"] = self._bilanSessions.trigramme_formation
        self._champs["periodeSessionsEvaluees"] = self._bilanSessions.periode_pour_titre
        self._champs["nbSessionsEvaluees"] = f"{len(self._bilanSessions.df_sessions_filtre_enTete)} session" + ("s" if len(self._bilanSessions.df_sessions_filtre_enTete) > 1 else "")  # Valeur toutes les données  
        self._champs["numerosSessions"] = "\n".join(self._bilanSessions.df_sessions_filtre_enTete["N° Session"].dropna().astype(str).unique())
        self._champs["nbApprenants"] = f"{self._bilanSessions.df_sessions_filtre_enTete['Nb. Nommés'].sum()} apprenant" + ("s" if self._bilanSessions.df_sessions_filtre_enTete['Nb. Nommés'].sum() > 1 else "")  # Valeur toutes les données
        self._champs["rp"] = ", ".join(self._bilanSessions.df_sessions_filtre_enTete["Nom responsable pédag."].dropna().astype(str).unique() + " " + self._bilanSessions.df_sessions_filtre_enTete["Prénom responsable pédag."].dropna().astype(str).unique())  # Valeur toutes les données
        self._champs["af"] = ", ".join(self._bilanSessions.df_sessions_filtre_enTete["Créée par"].dropna().astype(str).unique())  # Valeur toutes les données

    def _construire_commentaires(self):
        """
        Construit la partie commentaires du bilan de sessions
        """
        commentaires = ""

        # Cas avec aucun CSV dispo pour les stats
        if len(self._bilanSessions.liste_codesIRIS_pour_statsCSV) == 0 :
            vlog.print("Info", f"⚠️  Aucun CSV disponible pour ce bilan : les statistiques des stagiaires ne seront pas évaluées.")
            commentaires = f"\n⚠️  Aucun CSV disponible pour ce bilan : les statistiques des stagiaires ne seront pas évaluées.\n"
        
        # Cas avec certains CSV non dispo pour les stats mais pas tous
        elif len(self._bilanSessions.liste_codesIRIS_pour_enTete) != len(self._bilanSessions.liste_codesIRIS_pour_statsCSV) :

            print(self._bilanSessions.df_stagiaires_filtre_statsCSV_1ligne_par_session)
            print()
            print(self._bilanSessions.df_stagiaires_filtre_statsCSV_1ligne_par_session['N° Session'].tolist())

            commentaires = "\n⚠️  Certaines sessions n'ont pas de CSV exploitables pour les statistiques (cf. liste ci-dessous)."
            commentaires += "\n\nDonnées employées pour les statistiques :"
            commentaires += "\n   • Sessions évaluées : "+"".join(f"\n       - {session}" for session in self._bilanSessions.df_stagiaires_filtre_statsCSV_1ligne_par_session['N° Session'].tolist())
            commentaires += "\n   • Nombre d'apprenants sur ces sessions : " + f"{self._nb_apprenants:.0f}"
            commentaires += "\n   • Nombre de stagiaires ayant formulé des retours : " + f"{self._nb_stagiaires_retours:.0f}"

        # Affichage sessions avec pb CSV
        if len(self._bilanSessions.liste_codesIRIS_avec_pb_CSV) > 0:
            commentaires += "\n\nListe des sessions dont les statistiques n'ont pas pu être évaluées :"
            l_codes_IRIS = self._bilanSessions.liste_codesIRIS_avec_pb_CSV

            for statut_evalStat, statut_pourBilan in self._bilanSessions._mapping_statuts.items():
                if (statut_evalStat not in ["Traité", "Exclu - Code IRIS pas dans Extract IRIS sessions"]) :
                    # On filtre par statut
                    codes_statut = self._bilanSessions._get_codesIRIS_par_statut_evalStat(statut_evalStat)

                    # On fait l'intersection avec ceux qui ont pb CSV
                    codes_finaux = [c for c in codes_statut if c in l_codes_IRIS]

                    l_sessions = (
                        self._bilanSessions.df_sessions_filtre_avec_pb_CSV[self._bilanSessions.df_sessions_filtre_avec_pb_CSV['Code IRIS'].isin(codes_finaux)]
                        .dropna(subset=["N° Session"])
                        .drop_duplicates(subset=['Code IRIS'])
                        ["N° Session"]
                        .tolist()
                        )
                    if len(l_sessions) > 0:
                        commentaires += f"\n   • {statut_pourBilan} :" + "".join(f"\n       - {isession}" for isession in l_sessions)

        # Affichage sessions exclues
        if len(self._bilanSessions.liste_codesIRIS_exclus_totalement) > 0:
            commentaires += "\n\nListe des sessions de la période entièrement exclues du bilan :"
            l_codes_IRIS = self._bilanSessions.liste_codesIRIS_exclus_totalement
            l_sessions = (
                self._bilanSessions.df_sessions_filtre_exclus_totalement[self._bilanSessions.df_sessions_filtre_exclus_totalement['Code IRIS'].isin(l_codes_IRIS)]
                .dropna()
                .drop_duplicates(subset=['Code IRIS'])
                ["N° Session"]
                .tolist()
                )
            if len(l_sessions) > 0:
                commentaires += "".join(f"\n       - {isession}" for isession in l_sessions)

        #vlog.print("Info", f"\n{commentaires}")     

        self._champs["commentairesBilan"] = commentaires

    def _construire_stats(self):
        """
        Construit la partie statistiques du bilan de sessions
        """
        # On n'affecte les champs suivants que si des CSV sont disponibles pour les stats
        if len(self._bilanSessions.liste_codesIRIS_pour_statsCSV) > 0 :
            # Satisfaction globale
            self._champs["satisfactionGlobale_moy"] = self._get_stat_avec_format(
                "Satisfaction globale",
                "Moyenne",
                lambda v: f"{v:.1f}/5"
            )
            self._champs["satisfactionGlobale_com"] = self._get_stat_avec_format(
                "Satisfaction globale",
                "Commentaires",
                lambda v: v.replace("_x000D_", "\n"),
                default=""
            )
            """
            try:
                self._satisfactionGlobale_moy = f'{self._bilanSession._stats_stagiaires["Satisfaction globale"]["Moyenne"]:.1f}/5'
            except:
                self._satisfactionGlobale_moy = "Pas de donnée"      
            try:
                self._satisfactionGlobale_com = self._bilanSession._stats_stagiaires["Satisfaction globale"]["Commentaires"].replace("_x000D_", "\n")
            except:
                pass
            """

            # Recommanderiez-vous + commentaires remarques suggestions
            self._champs["recommandation_moy"] = self._get_stat_avec_format(
                "Recommanderiez-vous cette formation ?",
                "Moyenne",
                lambda v: f"{v/5*100:.0f}%"  # (on divise par 5 car on a un booléen stcké sous forme de note sur 5 : 0 = False, 5 = True)
            )
            self._champs["commentairesRemarquesSuggestions_com"] = self._get_stat_avec_format(
                "Commentaires, remarques, suggestions",
                "Commentaires",
                lambda v: v.replace("_x000D_", "\n"),
                default=""
            )
            """
            try:
                self._recommandation_moy = f'{self._bilanSession._stats_stagiaires["Recommanderiez-vous cette formation ?"]["Moyenne"]/5*100:.0f}%'  # (on divise par 5 car on a un booléen stcké sous forme de note sur 5 : 0 = False, 5 = True)
            except:
                self._recommandation_moy = "Pas de donnée"        
            try:
                self._commentairesRemarquesSuggestions_com = self._bilanSession._stats_stagiaires["Commentaires, remarques, suggestions"]["Commentaires"].replace("_x000D_", "\n")
            except:
                pass
            """

            # Notes inférieures à 3
            stats_sous_3 = {  # Dictionnaire pour les critères dont la moyenne est inférieure à 3 et non exclus (critères dans la liste self._CRITERES_A_ENLEVER)
                critere: valeurs
                for critere, valeurs in self._bilanSessions._stats_stagiaires.items()
                if (
                    critere not in self._bilanSessions._CRITERES_A_ENLEVER
                    and valeurs["Moyenne"] is not None
                    and valeurs["Moyenne"] < 3
                )
            }
            self._champs["evalInf3_val"] = f"{len(stats_sous_3)}"
            self._champs["evalInf3_com"] = "\n".join(
                f"• {clef} ({valeurs['Moyenne']:.1f}) :{valeurs['Commentaires'].replace('•', '\n   -').replace('\n\n', '\n')}"
                for clef, valeurs in stats_sous_3.items()
                )
            """
            try:
                self._evalInf3_val = f"{len(stats_sous_3)}"
            except:
                self._evalInf3_val = "Pas de donnée"
            try:
                self._evalInf3_com = "\n".join(f"• {clef} ({valeurs['Moyenne']:.1f}) :{valeurs['Commentaires'].replace('•', '\n   -').replace('\n\n', '\n')}"
                    for clef, valeurs in stats_sous_3.items()
                )

                #self._evalInf3_com = "\n".join(
                #    valeurs["Commentaires"]
                #    for valeurs in stats_sous_3.values()
                #    if valeurs["Commentaires"]
                #).replace("_x000D_", "\n")
            except:
                pass
            """

            # Taux de retour
            self._champs["tauxRetours_val"] = f"{(self._nb_stagiaires_retours/self._nb_apprenants)*100:.0f}%"
            """
            try:
                self._tauxRetours_val = f"{(self._nb_stagiaires_retours/self._nb_apprenants)*100:.0f}%"
            except:
                self._tauxRetours_val = "Pas de donnée"
            """

        else: # Cas : aucun CSV
            defaut = "Aucun CSV dispo"
            self._champs["satisfactionGlobale_moy"] = defaut
            self._champs["satisfactionGlobale_com"] = defaut
            
            self._champs["recommandation_moy"] = defaut
            self._champs["commentairesRemarquesSuggestions_com"] = defaut
            
            self._champs["evalInf3_val"] = defaut
            self._champs["evalInf3_com"] = defaut
            
            self._champs["tauxRetours_val"] = defaut


    # ====================
    # === MÉTHODES GET === 
    # ====================
    def _get_stat(self, critere: str, champ: str, default=None) -> int|float|str|None:
        """
        Récupère une stat depuis le dictionnaire _stats_stagiaires

        :param critere: Critère de la stat à récupérer
        :type critere: str
        :param champ: Champ de ce critère à récupérer (["Nombre", "Moyenne", "Commentaires"])
        :type champ: str
        :param default: Valeur retournée si ce critère n'existe pas. Défaut = None
        :type default: _type_, optional
        :return: la valeur du champ de ce critère (ex. _stats_stagiaires["Satisfaction globale"]["Moyenne"])
        :rtype: int|float|str|None
        """
        stats = self._bilanSessions._stats_stagiaires.get(critere)
        if not stats:
            return default

        valeur = stats.get(champ)
        return valeur if valeur is not None else default

    def _get_stat_avec_format(self, critere: str, champ: str, format:Callable[[Any], str], default:str="Pas de donnée") -> str:
        """
        Récupère et formatte une stat depuis le dictionnaire _stats_stagiaires

        :param critere: Critère de la stat à récupérer
        :type critere: str
        :param champ: _descChamp de ce critère à récupérer (["Nombre", "Moyenne", "Commentaires"])ription_
        :type champ: str
        :param format: format à appliquer (ex. lambda v: f"{v:.1f}/5" ou lambda v: f"{v/5*100:.0f}%" ou lambda v: v.replace("_x000D_", "\n"))
        :type format: Callable[[Any], str]
        :param default: Valeur renvoyée si aucune donnée. Défaut = "Pas de donnée"
        :type default: str, optional
        :return: une statistique formatée en str
        :rtype: str
        """
        valeur = self._get_stat(critere, champ)

        if valeur is None:
            return default

        try:
            return format(valeur)
        except Exception:
            return default









# Compréhension décorateur générique pour mettre en cache des données lourdes que je recalculais pleinde fois en @property
"""
# from functools import cached_property fait déjà ça nativement, mais sinon :

def cached_property(func):
    attr_name = f"_cache_{func.__name__}"

    @property
    def wrapper(self):
        if not hasattr(self, attr_name):
            setattr(self, attr_name, func(self))
        return getattr(self, attr_name)

    return wrapper


Permet de mettre en cache une valeur plutôt que de la réévaluer plein de fois.

Ex. avant je faisais ceci :
@property
def df_sessions_filtre_periode(self) -> DataFrame:
    return self.iris_sessions.df_filtre_periode(trigramme_formation=self.trigramme_formation, annee=self._annee, periode=self._periode)

Maintenant je peux faire :
@cached_property
def df_sessions_filtre_periode(self) -> DataFrame:
    return self.iris_sessions.df_filtre_periode(trigramme_formation=self.trigramme_formation, annee=self._annee, periode=self._periode)
"""