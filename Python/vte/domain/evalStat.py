from __future__ import annotations
from functools import cached_property
from pathlib import Path
from typing import Optional, Protocol

import pandas as pd
from pandas import DataFrame

from vte.domain.iris import IRIS_sessions
from vte.core import config
from vte.core.iris_referentiel import *
from vte.utils.office import FichierExcel
from vte.utils.utils import *
from vte.utils.utils import ouvrir_dossier as utils_ouvrir_dossier
from vte.utils.utils_instn import construire_chemin_config, recupere_trig_formation_depuis_chemin

# TODO : Pour l'instant c'est une classe de traitement. Le jour où j'ai besoin d'ouvrir un EvalStat pour le lire uniquement, prendre modèle sur IRIS avec des classes de lecture et de traitement

# ======================================================================================
# PROTOCOLES
# (pour faire passer les informations des objets parents sans ref circulaires)
# ======================================================================================
class Formation_protocol(Protocol):
    """
    Protocol de Formation : permet de simuler une formation en évitant les références circulaires
    """
    def get_session_par_codeIRIS(self, code_IRIS:int) -> Optional[Session_protocol]: ...

    @staticmethod
    def get_codesIRIS_par_statutBilan(formation:Formation_protocol, statut_pourBilan: str, codes_IRIS:Optional[int|Iterable[int]]=[]) -> list[int]:...

    @property
    def trigramme_formation(self) -> str: ...
    
    @property
    def sessions(self) -> list[Session_protocol]: ...
    
    @property
    def eval(self) -> EvalStat_formation|None: ...

class Session_protocol(Protocol):
    """
    Protocol de Session : permet de simuler une session en évitant les références circulaires
    """
    @property
    def code_IRIS(self) -> int|None: ...
        
    @property
    def trigramme_formation(self) -> str: ...

    @property
    def eval(self) -> Optional[EvalStat_session]: ...

    @eval.setter
    def eval(self, valeur:Optional[EvalStat_session]) -> None: ...

    @property
    def eval_formation(self) -> EvalStat_formation: ...




# ======================================================================================
# CLASSE EVALSTAT
# Objet fichier EvalStat + logique directement liée au fichier
# ======================================================================================
class EvalStat:
    """
    Classe mère pour le traitement des évaluations stagiaires individuelles.

    Gère :
       - la lecture des fichiers CSV stagiaires, 
       - la création des fichiers d'évaluation au format xlsx,
       - la création/mise à jour du fichier d'évaluation de la formation au format xlsx.
    """
    # ===========================
    # === VARIABLES DE CLASSE ===
    # ===========================
        
    # === dico_colonnes[critère d'évaluation][type d'info] ===
    _dico_colonnes:dict[str, dict[str, str]] = {
        "Comment avez-vous connu cette formation ?": {"Type colonne": "Commentaires seuls", "Groupe critère": "Question ouverte"},
        "Accueil, organisation et qualité des informations délivrées": {"Type colonne": "Avec commentaire", "Groupe critère": "Accueil & conseils"},
        "Conseils et orientation avant l'inscription": {"Type colonne": "Avec commentaire", "Groupe critère": "Accueil & conseils"},
        "Informations après l'inscription": {"Type colonne": "Avec commentaire", "Groupe critère": "Accueil & conseils"},
        "Accueil à l'arrivée sur site": {"Type colonne": "Avec commentaire", "Groupe critère": "Accueil & conseils"},
        "Prise en compte de vos besoins et attentes": {"Type colonne": "Avec commentaire", "Groupe critère": "Pédagogie"},
        "Qualité des animations": {"Type colonne": "Avec commentaire", "Groupe critère": "Interventions"},
        "Logique d'enchainement des interventions": {"Type colonne": "Avec commentaire", "Groupe critère": "Pédagogie"},
        "Qualité des supports de cours utilisés": {"Type colonne": "Avec commentaire", "Groupe critère": "Interventions"},
        "Qualité des moyens pédagogique": {"Type colonne": "Avec commentaire", "Groupe critère": "Pédagogie"},
        "Accès aux outils digitaux": {"Type colonne": "Avec commentaire", "Groupe critère": "Pédagogie"},
        "Satisfaction globale": {"Type colonne": "Avec commentaire", "Groupe critère": "Satisfaction"},
        "Recommanderiez-vous cette formation ?": {"Type colonne": "Note seule", "Groupe critère": "Satisfaction"},
        "Avez-vous d'autres besoins de formation ?": {"Type colonne": "Avec commentaire", "Groupe critère": "Commercial"},
        "Commentaires, remarques, suggestions": {"Type colonne": "Commentaires seuls", "Groupe critère": "Question ouverte"},
    }

    # === Critères qui n'ont pas de valeur "classique" pour les stats (i.e. ce sont des bool ou des str), ils ne sont pas à retenir pour le calcul des moyennes < 3 ===
    _criteres_sans_note_standard:list[str] = [  
        "Comment avez-vous connu cette formation ?", "Avez-vous d'autres besoins de formation ?", "Commentaires, remarques, suggestions", "Recommanderiez-vous cette formation ?"]

    # === Colonnes du CSV avec les noms qu'il faudrait (sans espaces en trop ou trucs bizares) ===
    _colonnes_csv_bonsNoms = [
        "Chemin fichier CSV",
        "Date",
        "Prénom",
        "Nom",
        "Entreprise",
        "Code session",
        "Comment avez-vous connu cette formation ?",
        "Accueil, organisation et qualité des informations délivrées",
        "Commentaires",
        "Conseils et orientation avant l'inscription",
        "Commentaires ",
        "Informations après l'inscription",
        "Commentaires  ",
        "Accueil à l'arrivée sur site",
        "Commentaires   ",
        "Prise en compte de vos besoins et attentes",
        "Commentaires    ",
        "Qualité des animations",
        "Commentaires     ",
        "Logique d'enchainement des interventions",
        "Commentaires      ",
        "Qualité des supports de cours utilisés",
        "Commentaires       ",
        "Qualité des moyens pédagogique",
        "Commentaires        ",
        "Accès aux outils digitaux",
        "Commentaires         ",
        "Satisfaction globale",
        "Commentaires          ",
        "Recommanderiez-vous cette formation ?",
        "Avez-vous d'autres besoins de formation ?",
        "Lesquels ?",
        "Commentaires, remarques, suggestions"]
    
    # === Colonnes descriptives à recopier ===
    _colonnes_csv_fixes = [
        "Chemin fichier CSV", "Prénom", "Nom", "Entreprise", "Code session"]

    # === Colonnes de l'extract IRIS Sessions à récupérer ===
    _colonnes_sessions = [
        "N° Session",
        "Formation",
        "Trigramme formation",
        "Code IRIS",
        "Date début ses.",
        "Année début ses.",
        "Type de formation",
        "Trigramme RP",
        "Trigramme AF",
        "Nb. Nommés"]



    # ====================
    # === CONSTRUCTEUR ===
    # ====================
    def __init__(self) -> None:
        """
        Crée une instance d'EvalStat à minima (self_fe = None)
        """
        self._fe:Optional[FichierExcel] = None  # Objet Excel contenant les données EvalStat stagiaire individuel



    # =========================
    # === METHODES INTERNES ===
    # =========================
    def _ecrit_df_et_sauve(self, nouveau_chemin_fichier:Optional[Path] = None) -> None:
        """
        Ecrit les dataframes csv et stagiaires dans le fichier excel.

        Sauve et ferme.

        Actualise les TCD.

        Si nouveau_chemin est None, alors on écrit au même endroit que l'ancien fichier (sauvegarde simple)

        :param nouveau_chemin_fichier: Chemin du fichier de sortie. Défaut = None.
        :type nouveau_chemin_fichier: Optional[Path]
        """

        # On écrit et on sauve (pour l'instant on remplace tout le dataframe sans optimiser)
        self.fe.get_tableau("CSV_stagiaires").ecrit_dataFrame_dans_tableauStructure(supprimeDonneesEtRemplace = True)
        self.fe.get_tableau("Stagiaires").ecrit_dataFrame_dans_tableauStructure(supprimeDonneesEtRemplace = True)

        self._fe.save(nouveau_chemin_fichier)
        self._fe.close()

        # Actualiser TCD
        self._fe.actualiser_TCD()





    # =========================
    # === METHODES EXTERNES ===
    # =========================




    # =========================
    # === GETTERS / SETTERS ===
    # =========================
    @property
    def criteres_sans_note_standard(self) -> list[str]:
        return self._criteres_sans_note_standard

    @property
    def fe(self) -> FichierExcel|None:
        return self._fe
     
    @property
    def df_csv(self) -> DataFrame|None:
        return self._fe.get_df_tableau("CSV_stagiaires")

    @df_csv.setter
    def df_csv(self, valeur:DataFrame) -> None:
        self._fe.set_df_tableau("CSV_stagiaires", df=valeur)

    @property
    def df_stagiaires(self) -> DataFrame|None:
        return self._fe.get_df_tableau("Stagiaires")

    @df_stagiaires.setter
    def df_stagiaires(self, valeur:DataFrame) -> None:
        self._fe.set_df_tableau("Stagiaires", df=valeur)

    @property
    def chemin_fe(self) -> Path|None:
        return self._fe.chemin_fichier

    # Liens avec IRIS sessions
    @property
    def iris_sessions(self) -> IRIS_sessions:
        return get_iris(typeExport="Sessions")
    
    @property
    def df_IRIS_sessions(self) -> DataFrame:
        return self.iris_sessions.df





# ======================================================================================
# CLASSE EVALSTAT_FORMATION
# ======================================================================================
class EvalStat_formation(EvalStat): 
    """
    Classe evalStat Formation employée en parllèle du traitement des évaluations stagiaires individuelles.

    Gère la création ou la mise à jour du fichier d'évaluation de la formation au format xlsx.
    """

    # =====================
    # === CONSTRUCTEURS ===
    # =====================
    def __init__(self, formation:Formation_protocol):
        """
        Initialisation d'un EvalStat formation a minima
        """
        # On initialise la classe mère
        super().__init__()

        self._formation = formation  # Protocol pour éviter les références circulaires
        #self._fe → classe mère

        self._df_initial_hash:Optional[str] = None  # hash du df initial pour savoir s'il a été modifié, auquel cas on sauvegardera à la fin
        self._supprimeEtRemplace_donneesEval:Optional[bool] = None

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
    def avec_ouverture_ou_creation(cls, formation:Formation_protocol) -> EvalStat_formation:
        """
        Initialisation d'un EvalStat formation avec (ouverture ou création) du fichier d'évaluation de la formation
        """
        instance = cls(formation)

        instance._ouvrir_ou_creer_eval_formation()

        return instance



    # =========================
    # === METHODES INTERNES ===
    # =========================
    def _ouvrir_ou_creer_eval_formation(self) -> None:  # -> Tuple[FichierExcel, bool]:
        """
        Ouvre ou crée le fichier Excel d'évaluations d'une formation.

        Construit le chemin vers le fichier d'évaluations correspondant au trigramme de la formation.

        Si ce fichier existe, il est ouvert et les données des stagiaires sont chargées dans un DataFrame.

        Sinon, un nouveau fichier est créé à partir d'un modèle, et les données seront à initialiser.

        :param trigramme_formation: Trigramme de la formation.
        :type trigramme_formation: str
        """
        timer.debut(f"Ouverture ou création du fichier Excel de la formation {self.trigramme_formation}")
        
        # On regarde si Evaluation-Stagiaires-Global-XXX.xlsx existe
        # Auquel cas on l'ouvre et on ne supprimera pas les données existantes
        if self.chemin_eval_formation.is_file():
            # Alors on l'ouvre
            self._fe = FichierExcel.depuis_fichier(chemin_fichier=self.chemin_eval_formation)
            vlog.ajouter_message("Ouverture EvalStat Global formation", self.chemin_fe, style=["vert"])

            # On récupère la liste des sessions déjà intégrées (liste des codes et des chemins) → Ce sera pour écrire dans le tkinter
            """
            dico_sessionsDejaTraitees = dict(
                self.df_stagiaires[self.df_stagiaires["Trigramme formation"] == self.trigramme_formation]   # 1. filtre sur le trigramme
                .drop_duplicates(subset=["Code IRIS", "Chemin fichier CSV"])                                # 2. élimine les doublons
                [["Code IRIS", "Chemin fichier CSV"]]                                                       # 3. sélection des colonnes
                .values                                                                                     # 4. valeurs du DF
                )  
            """

            # Il ne faudra pas supprimer les anciennes données de fe_evaluations_formation
            self._supprimeEtRemplace_donneesEval = False

            
       
        # Sinon on crée l'évaluation depuis le modèle et on supprimera les anciennes données
        else:
            #print("Création nouvel EvalStat formation")

            # On vérifie que le répertoire dédié existe sinon on le créée : \\instnt\partage\FORMATIONS_C\###\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\
            self.chemin_eval_formation.parent.mkdir(parents=True, exist_ok=True)

            # On créée le fichier excel à partir du modèle
            self._fe = FichierExcel.depuis_modele(
                chemin_modele = config.CHEMIN_MODELE_EXCEL_EVALUATIONS_STAGIAIRES, 
                chemin_fichier_sauv = self.chemin_eval_formation
                )
            vlog.ajouter_message("Création EvalStat Global formation", self.chemin_fe, style=["vert"])
            
            # Il faudra supprimer les anciennes données de fe_evaluations_formation
            self._supprimeEtRemplace_donneesEval = True

        # On fait un hash du df pour savoir si, en fin de traitement il aura été modifié, auquel cas on le sauvegardera
        self._df_initial_hash = hash_df(self.df_stagiaires)
        timer.fin()

        #return self._fe_evaluations_formation, self._df_evaluations_formation, supprimeDonneesEtRemplace



    # =========================
    # === METHODES EXTERNES ===
    # =========================
    def ecrit_et_sauve_df_siModif(self):
        """
        Sauvegarde et fermeture du fichier d'évaluation de la formation.

        On ne l'exécute que si le DataFrame fichier d'évaluation a été modifié.
        """
        # TODO : dans _mettre_a_jour_evaluations_formation() je mets à jour le df de l'excel évaluation formation. Il sera écrit physiquement à la sortie du contexte formation.
        # TODO : df nouveau df comprend l'ancien (i.e. évaluation formation existant) + le nouveau que l'on traite.
        # TODO : pour l'instant je réécris tout ce df mais pour être optimal on ne pourrait écrire que le nouveau


        print(f"\n\n{Style.BRIGHT}{Fore.RED}Écriture du fichier global des évaluations de la formation {self.trigramme_formation}")

        if self._fe is not None:
            df_final_hash = hash_df(self.df_stagiaires)

            if df_final_hash != self._df_initial_hash:
                # Le DataFrame a changé → on sauvegarde
                self._ecrit_df_et_sauve()

            else:
                vlog.ajouter_message(
                    "Aucune modification détectée dans l'évaluation formation → pas de sauvegarde",
                    self.chemin_fe,
                    style=["jaune"]
                )

    def verifier_presence_code_IRIS_dans_eval(self, code_IRIS:int) -> bool:
        """
        On regarde dans le DataFrame de l'evaluation si un code IRIS est déjà présent

        :param code_IRIS: code IRIS à tester
        :type code_IRIS: int
        :return: si True, alors le code IRIS est déjà présent, sinon False
        :rtype: bool
        """
        if code_IRIS in self.df_stagiaires["Code IRIS"].values:
            #print("✅ Code IRIS présent")
            return True
        else:
            return False

    def verifier_traitement_evalStat_session(self, code_IRIS:Optional[int]=None, chemin_csv:Optional[Path]=None) -> Tuple[bool, str]:
        """
        Vérifie si l'on doit traiter un EvalStat session à partir de :
            - la présence dans eval formation (code_IRIS ou chemin_csv) ;
            - la présence de code_IRIS dans IRIS_sessions.

        :param code_IRIS: Code IRIS de la session, défaut = None
        :type code_IRIS: Optional[int], optional
        :param chemin_csv: chemin du csv de l'évaluation de la session, défaut = None
        :type chemin_csv: Optional[Path], optional
        :return:
        - Un bool pour dire si le traitement doit continuer 
            - True il faut traiter l'EvalStat (code IRIS ou CSV non détectés dans EvalStat + code IRIS présent dans IRIS sessions) ; 
            - False il ne faut pas traiter l'EvalStat.
        - un string avec le statut pour connaitre l'exclusion de traitement le cas échéant (pertinent ssi False ; si True on renvoie chaine vide).
        :rtype: Tuple[bool, str]
        """
        # Si code_IRIS==None et chemin_csv==None alors on ne peut rien vérifier
        if (code_IRIS is None) and (chemin_csv is None):
            vlog.log_erreur("On ne peut pas traiter la vérification car on n'a ni code IRIS ni chemin CSV.")
        
        # On vérifie que l'eval formation est ouvert en mémoire sinon c'est le premier appel → on ouvre ou crée
        if self._fe is None:
            self._ouvrir_ou_creer_eval_formation()

        # Vérifications 1 : liés à code_IRIS
        if code_IRIS is not None:
            # 2.1 : On vérifie que code_iris est bien un entier à 5 chiffres
            est_code_IRIS_valide, code_IRIS = IRIS.verifier_code_IRIS(code_IRIS, int)
            if not est_code_IRIS_valide:
                vlog.log_erreur(f"❌  Code IRIS {code_IRIS} non valide : EvalStat {code_IRIS} non traité.", continuer=True)
                statut = "Exclu - Code IRIS renseigné non valide"
                return False, statut
        
            # 2.2 : On vérifie que le code IRIS n'est pas déjà dans le fichier global de la session sinon on n'a pas besoin de traiter (déjà fait)
            if self.verifier_presence_code_IRIS_dans_eval(code_IRIS):  
                print(f"✅  Code IRIS {code_IRIS} déjà présent dans fichier eval de la formation : EvalStat {code_IRIS} exclu du traitement.")
                statut = "Traité"  # "Traité - Code IRIS déjà dans l'évaluation de la formation" (je ne peux pas mettre plusieurs statuts traités car après dans bilanSession, dans la matrice "statut evalStat" "statut pour bilan" il me fait une correspondance 1 pour 1)
                return False, statut

            # 2.3 : On vérifie si le code_IRIS est bien existant dans l'extract IRIS
            if code_IRIS not in self.df_IRIS_sessions["Code IRIS"].values:
                print(f"⚠️  Code IRIS {code_IRIS} non trouvé dans l’extract IRIS  : EvalStat {code_IRIS} exclu du traitement.")
                statut = "Exclu - Code IRIS pas dans Extract IRIS sessions"
                return False, statut
        




        # Vérifications 2 : liés au chemin du CSV
        if chemin_csv is not None and code_IRIS is not None:  # Ca ne sert à rien de faire une double vérification de la présence dans eval formation
            # Si le chemin est avec un raccourci réseau alors on récupère le chemin en entier + on convertit en Path 
            chemin_csv = chemin_vers_unc(Path(chemin_csv))

            # on vérifie que chemin_csv_session n'est pas déjà dans le fichier évaluations des formations
            if str(chemin_csv) in self.df_stagiaires["Chemin fichier CSV"].drop_duplicates().tolist():  
                print(f"✅  CSV déjà présent dans fichier eval de la formation : EvalStat {chemin_csv} exclu du traitement.")
                #pprint(self._df_evaluations_formation["Chemin fichier CSV"])
                statut = "Traité"  # "Traité - CSV déjà dans l'évaluation de la formation" (je ne peux pas mettre plusieurs statuts traités car après dans bilanSession, dans la matrice "statut evalStat" "statut pour bilan" il me fait une correspondance 1 pour 1)
                return False, statut

        # Si aucun Test n'est vérifié
        return True, "A traiter"

    def chemin_csv_depuis_code_IRIS(self, code_IRIS:int) -> Optional[Path] :
        """
        Récupère le chemin CSV d'un code IRIS tel que stocké dans l'évaluation formation

        :param code_IRIS: Code IRIS dont il faut récupérer le chemin du CSV
        :type code_IRIS: int
        :return: Chemin du CSV si trouvé, sinon None
        :rtype: Optional[Path]
        """
        df = self.df_stagiaires[self.df_stagiaires['Code IRIS'] == code_IRIS]

        if df.empty:
            return None
        
        return Path(df["Chemin fichier CSV"].iloc[0])




    def liste_codesIRIS_avec_CSV_presents_et_non_vides(self, codes_IRIS:Optional[int|Iterable[int]]=[]) -> list[int]:
        """
        Retourne, parmi une liste donnée, la liste des codes IRIS pour lesquels on peut calculer les stats à partir des CSV (i.e. codes IRIS pour lesquels le CSV est fonctionnel et non vide).

        On emploie uniquement les codes_IRIS (sous-partie) parmi toutes les sessions comprises dans formation.sessions. Si None, alors on fait toutes les sessions.

        :param codes_IRIS: Sous-partie des codes IRIS à traiter parmi toutes les sessions étant dans formation.sessions. Si None, alors on fait toutes les sessions de formation. Défaut = [].
        :type codes_IRIS: Optional[int|Iterable[int]]
        :return: La liste des codes IRIS pour lesquels on peut calculer les stats à partir des CSV.
        :rtype: list[int]
        """
        # Je passe par self._formation et pas Formation_protocol car la méthode statique d'un protocole n'est pas appelée et reste dans le protocole.
        # En effet les méthodes statiques sont des méthodes de classe or les contrats fais avec les protocoles se font sur des méthodes instanciées.
        return self._formation.get_codesIRIS_par_statutBilan(
            formation=self._formation,
            codes_IRIS=codes_IRIS,
            statut_pourBilan="Exploités pour les évaluations (CSV présents et non vides)"
            )

    def calculer_stats_criteres(self, codes_IRIS:Optional[int|Iterable[int]]=[]) -> None:  #dict[str, dict[str, int|float|str|None]]:
        """
        Calcule, pour une liste code_IRIS donnée, les statistiques (nombre de retours, moyenne retours, agrégation des commentaires) de tous les critères.

        Si codes_IRIS=None, on fait les stats pour toutes les sessions dans formation.sessions (sinon on prend les formations.session qui sont dans codes_IRIS uniquement).

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

        
        :param codes_IRIS: Sous-partie des codes IRIS à traiter parmi toutes les sessions étant dans formation.sessions. Si None, alors on fait toutes les sessions de formation. Défaut = [].
        :type codes_IRIS: Optional[int|Iterable[int]]
        :return: Un dictionnaire de tous les critères
        :rtype: dict
        """
        # Je ne veux faire les stats que des codes_IRIS qui sont dans formations.sessions
        # S'il y a des stats, alors les codes IRIS seront dans l'eval formation

        # S'il n'y a pas de CSV disponibles pour les stats, alors ce n'est pas la peine de faire les stats
        if len(self.liste_codesIRIS_avec_CSV_presents_et_non_vides(codes_IRIS = codes_IRIS)) > 0 :
            # On récupère la iste des critères
            liste_criteres = self.df_stagiaires_filtre_statsCSV(codes_IRIS = codes_IRIS)['Critère'].dropna().unique()


            # On fait les stats pour chaque critère
            for critere in liste_criteres:
                # Dataframe filtré sur ce critère
                df_filtre = self.df_stagiaires_filtre_statsCSV(codes_IRIS = codes_IRIS)[self.df_stagiaires_filtre_statsCSV(codes_IRIS = codes_IRIS)['Critère'] == critere]
                
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
  
    def calculer_stats_criteres_BAK(self, codes_IRIS:Optional[int|Iterable[int]]=None) -> dict[str, dict[str, int|float|str|None]]:
        """
        Calcule, pour une liste code_IRIS donnée, les statistiques (nombre de retours, moyenne retours, agrégation des commentaires) de tous les critères.

        Si codes_IRIS=None, on fait les stats pour toutes les sessions dans formation.sessions (sinon on prend les formations.session qui sont dans codes_IRIS uniquement).

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

        
        :param codes_IRIS: Sous-partie des codes IRIS à traiter parmi toutes les sessions étant dans formation.sessions. Si None, alors on fait toutes les sessions de formation. Défaut = None.
        :type codes_IRIS: Optional[int|Iterable[int]], optional
        :return: Un dictionnaire de tous les critères
        :rtype: dict
        """

        # S'il n'y a pas de CSV disponibles pour les stats, alors ce n'est pas la peine de faire les stats
        if len(self.liste_codesIRIS_avec_CSV_presents_et_non_vides()) > 0 :
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

    def df_stagiaires_filtre_statsCSV(self, codes_IRIS:Optional[int|Iterable[int]]=[]) -> DataFrame:
        """
        df_stagiaires (eval formation) filtré sur les codes IRIS pour lesquels on peut calculer les stats à partir des CSV.

        Ca correspond aux codes IRIS pour lesquels le CSV est fonctionnel et non vide.

        On emploie uniquement les codes_IRIS (sous-partie) parmi toutes les sessions comprises dans formation.sessions. Si None, alors on fait toutes les sessions.

        :param codes_IRIS: Sous-partie des codes IRIS à traiter parmi toutes les sessions étant dans formation.sessions. Si None, alors on fait toutes les sessions de formation. Défaut = None.
        :type codes_IRIS: Optional[int|Iterable[int]], optional
        :return: df_stagiaires (eval formation) filtré sur les codes IRIS pour lesquels on peut calculer les stats à partir des CSV.
        :rtype: DataFrame
        """
        return self.df_stagiaires[self.df_stagiaires['Code IRIS'].isin(self.liste_codesIRIS_avec_CSV_presents_et_non_vides(codes_IRIS = codes_IRIS))]
    

    @staticmethod
    def construire_chemin_eval_formation(trigramme_formation:Optional[str]=None) -> Path:
        """
        Construit le chemin de l'évaluation formation à partir du trigramme de la formation (à partir des données de la config CHEMIN_EXCEL_EVALUATIONS_FORMATION) :
            - le chemin est transformé en unc (s'il y a un raccourci lecteur réseau sur le poste de l'utilisateur on transforme en chemin réseau complet) ;
            - l'utilisateur peut optimiser le chemin (chemin le plus long entre l'attendu et ce qui existe).

        :param trigramme_formation: Trigramme de la formation. Défaut = None
        :type trigramme_formation: Optional[str], optional
        
        :return: Le chemin de l'évaluation de la formation
        :rtype: Path
        """
        return construire_chemin_config(
            chemin_a_completer = config.CHEMIN_EXCEL_EVALUATIONS_FORMATION,
            trigramme_formation = trigramme_formation
        )
    
    

    # ====================
    # === MÉTHODES GET === 
    # ====================
    def get_stat(self, critere: str, champ: str, default=None) -> int|float|str|None:
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
        stats = self._stats_stagiaires.get(critere)
        if not stats:
            return default

        valeur = stats.get(champ)
        return valeur if valeur is not None else default

    def get_stat_avec_format(self, critere: str, champ: str, format:Callable[[Any], str], default:str="Pas de donnée") -> str:
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
        valeur = self.get_stat(critere, champ)

        if valeur is None:
            return default

        try:
            return format(valeur)
        except Exception:
            return default



    # =========================
    # === GETTERS / SETTERS ===
    # =========================

    @property
    def trigramme_formation(self) -> str:
        return self._formation.trigramme_formation
    
    @property
    def stats_stagiaires(self) -> dict[str, dict[str, int|float|str|None]]:
        return self._stats_stagiaires

    @cached_property
    def chemin_eval_formation(self) -> Path:
        return EvalStat_formation.construire_chemin_eval_formation(trigramme_formation=self.trigramme_formation)




# ======================================================================================
# CLASSE EVALSTAT_SESSION
# ======================================================================================
class EvalStat_session(EvalStat):
    """
    Classe evalStat Session pour le traitement des évaluations stagiaires individuelles.

    Gère :
       - la lecture des fichiers CSV stagiaires, 
       - la création des fichiers d'évaluation au format xlsx,
       - l'appel à EvalStat_formation pour la création/mise à jour du fichier d'évaluation de la formation au format xlsx.
    """
    
    # Dictionnaire pour mapper les statuts aux clés de self._statuts ["Exploités pour les évaluations (CSV présents)", "Exploités pour les évaluations (CSV présents)", "Exclus des évaluations (problème traitement CSV)", "Exclus des évaluations (CSV manquants)", "Exclus des évaluations (CSV vide / aucun retour)", "Exclus entièrement du bilan (non présent dans IRIS / mauvais code)"]
    _mapping_statuts:dict[str, str] = {
        "Traité": "Exploités pour les évaluations (CSV présents et non vides)",                                  # Exploités pour stats initiales → Dans _demande_sessions_a_exclure
        #  "Traité - CSV déjà dans fichier global": "Exploités pour les évaluations (CSV présents et non vides)",   # Exploités pour les stats stagiaires → Dans _maj_evalstat_formation
        "Exclu - Aucun CSV fourni": "Exclus des évaluations (CSV manquant)",                   # Exclus des évaluations car problème au traitement des CSV → Dans _maj_evalstat_formation
        "Exclu - Code IRIS pas dans Extract IRIS sessions": "Exclus entièrement du bilan (non présent dans IRIS / mauvais code)",  # Exclus entièrement du bilan car sessions non réalisées ou mauvais RP (exclus par l'utilisateur) → Dans _maj_evalstat_formation
        "Exclu - Problème lecture CSV": "Exclus des évaluations (problème traitement CSV)",         # Exclus des évaluations car CSV stagiaires manquants → Dans _maj_evalstat_formation
        "Exclu - CSV vide / Aucun retour": "Exclus des évaluations (CSV vide / aucun retour)",      # Exclus des évaluations car le CSV est vide (i.e. aucun retour d'utilisateur)
    }

    # Propriété de classe → @property ne marche que pour une instance → On fait une propriété de classe publique
    mapping_statuts:dict[str, str] = _mapping_statuts
    
    
    # =====================
    # === CONSTRUCTEURS ===
    # =====================
    def __init__(self, session:Session_protocol, statut:Optional[str]=None):
        """
        Crée l'instance EvalStat d'un session a minima.

        :param session: La session à laquelle est affectée l'EvalStat
        :type session: Session_protocol
        """
        # On initialise la classe mère
        super().__init__()

        self._session:Session_protocol = session  # C'est un protocol pour éviter les références circulaires
        self._chemin_csv: Optional[Path] = None
        self._statut: Optional[str] = statut  # ex: "A traiter", "Traité", "Exclu - Aucun CSV fourni", "Exclu - Code IRIS pas dans Extract IRIS sessions", "Exclu - Problème lecture CSV", "Exclu - CSV vide / Aucun retour"

    @classmethod
    def avec_ouverture_ou_traitement(cls, session:Session_protocol, chemin_csv:Optional[Path|str]=None, ecrire_eval_formation:bool=True, ouvrir_dossier:bool=False, ouvrir_fe:bool=False) -> EvalStat_session:
        """
        Initialisation d'un EvalStat session avec ouverture ou création du fichier Excel EvalStat.

        Si création :
            - crée l'instance EvalStat d'une session et traite cet EvalStat ;
            - crée l'Excel EvalStat d'une session à partir d'un CSV (s'il n'est pas donné, on ouvre un filedialog) ;
            - l'évaluation de la formation est mise à jour avec ces nouvelles données et est sauvée en fin de traitement selon le critère ecrire_eval_formation.

        :param session: La session à laquelle est affectée l'EvalStat
        :type session: Session_protocol
        :param chemin_csv: chemin du CSV à traiter. S'il est None, on ouvre un filedialog
        :type chemin_csv: Optional[Path|str]
        :param ecrire_eval_formation: Pour écrire physiquement l'Excel eval formation en fin de traitement. Si False, il devra être écrit ailleurs (à l'endroit où il y a la boucle pour du multi-traitement typiquement). Défaut = True.
        :type ecrire_eval_formation: bool
        :param ouvrir_dossier: Ouvre le répertoire de l'EvalStat généré. Defaut = False.
        :type ouvrir_dossier: bool
        :param ouvrir_fe: Si True, ouvre et charge l'Objet FichierExcel dans fe (i.e. si les données de eval formation ne suffisent pas)
        :type ouvrir_fe: bool, optional
        :return: l'EvalStat session
        :rtype: EvalStat_session
        """
        instance = cls(session)
        instance.ouvrir_ou_traiter_eval(
            chemin_csv=chemin_csv,
            ecrire_eval_formation=ecrire_eval_formation,
            ouvrir_dossier=ouvrir_dossier,
            ouvrir_fe=ouvrir_fe
        )
        return instance


    # TODO : probablement à virer
    @classmethod
    def avec_traitement(cls, session:Session_protocol, chemin_csv:Optional[Path|str]=None, ecrire_eval_formation:bool=True, ouvrir_dossier:bool=False) -> EvalStat_session:
        """
        Crée l'instance EvalStat d'une session et traite cet EvalStat.

        Crée l'Excel EvalStat d'une session à partir d'un CSV (s'il n'est pas donné, on ouvre un filedialog).

        L'évaluation de la formation est mise à jour avec ces nouvelles données et est sauvée en fin de traitement selon le critère ecrire_eval_formation.

        :param session: La session à laquelle est affectée l'EvalStat
        :type session: Session_protocol
        :param chemin_csv: chemin du CSV à traiter. S'il est None, on ouvre un filedialog
        :type chemin_csv: Optional[Path|str]
        :param ecrire_eval_formation: Pour écrire physiquement l'Excel eval formation en fin de traitement. Si False, il devra être écrit ailleurs (à l'endroit où il y a la boucle pour du multi-traitement typiquement). Défaut = True.
        :type ecrire_eval_formation: bool
        :param ouvrir_dossier: Ouvre le répertoire de l'EvalStat généré. Defaut = False.
        :type ouvrir_dossier: bool
        :return: _description_
        :rtype: EvalStat_session
        """
        instance = cls(session)
        instance.traiter_eval(
            chemin_csv=chemin_csv,
            ecrire_eval_formation=ecrire_eval_formation,
            ouvrir_dossier=ouvrir_dossier
        )
        return instance

        

    # ===================================================
    # === MÉTHODES D’INSTANCE - TRAITEMENT INDIVIDUEL ===
    # ===================================================
    def _charger_csv_stagiaire(self) -> Optional[DataFrame] :
        """
        Permet de stocker un CSV dans un dataframe en employant le bon encodage
        Si le return est None c'est qu'il y a eu un problème ou que le dataframe est vide (csv présent avec en-têtes mais sans ligne).
        Dans ce cas on écrit le self._statut_csv pour tracer la raison exclusion.
        chemin_csv est le chemin du CSV à aller récupérer (à ce stade il est connu)

        :param chemin_csv: chemin du csv à charger
        :type chemin_csv: Path
        :return: le dataframe du csv. Si None, c'est qu'il y a eu un problème ou que le dataframe est vide (csv présent avec en-têtes mais sans ligne).
        :rtype: Optional[DataFrame]
        """
        codage_csv = trouve_encodage_csv(self._chemin_csv)  # On récupère l'encodage et on importe le CSV dans un DataFrame
        try:
            df_csv_stagiaires = pd.read_csv(self._chemin_csv, sep=';', encoding=codage_csv)  # Ouverture du CSV et mise dans un DataFrame
        except Exception as e:
            print(f"❌ Erreur lors de la lecture du CSV {self._chemin_csv} : {e}")
            self._statut = "Exclu - Problème lecture CSV"
            return None

        if df_csv_stagiaires.empty:
            print("⚠️  CSV vide → fichier ignoré.")
            self._statut = "Exclu - CSV vide / Aucun retour"
            return None
        
        # Prise en compte qu'on a plusieurs formats de CSV : on doit traiter des colonnes en + ou - en conséquences
        if "Date de fin" in df_csv_stagiaires.columns:
            # Cas 1 (nouveau format de csv) : supprimer "Date de fin" → Test : r"\\instnt\PARTAGE\FORMATIONS_C\TEL\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\S-16411-rapports-session-evaluations\S-16411-FC25-TEL-VTE-CAR-Stagiaires.csv"
            df_csv_stagiaires = df_csv_stagiaires.drop(columns=["Date de fin"])
        else:
            # Cas 2 (ancien format de csv) : supprimer la 2e et 3e colonne (indices 1 et 2) → Test : r"\\instnt\PARTAGE\FORMATIONS_C\TEL\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\S-12766-rapports-session-evaluations\S-12766-FC22-TEL-JVI-LRA-Stagiaires.csv"
            df_csv_stagiaires = df_csv_stagiaires.drop(df_csv_stagiaires.columns[[1, 2]], axis=1)

        # On rajoute le chemin du CSV en première colonne
        df_csv_stagiaires.insert(0, "Chemin fichier CSV", str(self._chemin_csv))

        # Je renomme à la main toutes les colonnes à la main car les CSV c'est le bordel avec des espaces qui trainent et des caractères spéciaux
        mapping = dict(zip(df_csv_stagiaires.columns, self._colonnes_csv_bonsNoms))  # On fait un dictionnaire de mapping anciens noms/nouveaux noms
        df_csv_stagiaires = df_csv_stagiaires.rename(columns=mapping)  # On renomme les colonnes

        # Mise au format jj/mm/aaaa de la colonne "Date" (si elle existe)
        try:
            df_csv_stagiaires["Date"] = pd.to_datetime(df_csv_stagiaires["Date"], dayfirst=True, errors="coerce").dt.strftime("%d/%m/%Y")  # dayfirst=True indique que le premier nombre correspond au jour (format jj/mm/aaaa)
        except Exception as e:
            print(f"Erreur de conversion de la colonne Date : {e}")

        return df_csv_stagiaires

    def _traiter_df_csv(self) -> None:
        """
        On génère le DataFrame du CSV et on l'affecte à self.df_csv.
        """
        # Met à jour ou crée la colonne "Code session" avec self._codeIRIS
        self.df_csv["Code session"] = self.code_IRIS

    def _traiter_df_stagiaires(self) ->None:
        """
        Crée un dataframe à partir du CSV de sorte qu'on puisse l'exploiter par un TCD (regroupement par critères).
        On l'affecte à self.df_stagiaires
        """
        # Étape 1 — Préparation initiale de df_stagiaires à partir de df_csv_stagiaires

        # Nouveau DataFrame à remplir (au début c'est une liste, on convertira ensuite en dataframe)
        df_long = []

        # Parcours des lignes de CSV_stagiaires
        #print(self._df_csv_stagiaires.columns.tolist())
        for _, row in self.df_csv.iterrows():
            #print("Ligne en cours : ")
            #print(row)
            base = {col: row[col] for col in self._colonnes_csv_fixes}  # Création des colonnes qui seront répétées à chaque fois
            base["NOM Prénom"] = f"{str(row['Nom']).upper()} {row['Prénom']}".strip()  # Création du champ "NOM Prénom"

            for critere, meta in self._dico_colonnes.items():
                if critere not in row:
                    continue
                type_col = meta["Type colonne"]
                groupe_critere = meta["Groupe critère"]

                # --- Cas 1 : Note + commentaire ---
                if type_col == "Avec commentaire":
                    val = row[critere]

                    # conversion booléenne spéciale
                    if critere == "Avez-vous d'autres besoins de formation ?":
                        val_str = str(val).strip().lower()
                        val = 5 if val_str == "oui" else (0 if val_str == "non" else None)

                    commentaire_col = row.index[row.index.get_loc(critere) + 1]
                    commentaire = row.get(commentaire_col, None)

                    if pd.notna(val) or pd.notna(commentaire):
                        df_long.append({
                            **base,
                            "Groupe critère": groupe_critere,
                            "Critère": critere,
                            "Note": val,
                            "Commentaires": commentaire
                        })

                # --- Cas 2 : commentaires seuls ---
                elif type_col == "Commentaires seuls":

                    commentaire = row[critere]

                    if pd.notna(commentaire):
                        df_long.append({
                            **base,
                            "Groupe critère": groupe_critere,
                            "Critère": critere,
                            "Note": None,
                            "Commentaires": commentaire
                        })

                # --- Cas 3 : note seule ---
                elif type_col == "Note seule":

                    val = str(row[critere]).strip().lower()

                    note = 5 if val == "oui" else (0 if val == "non" else None)

                    if pd.notna(row[critere]):
                        df_long.append({
                            **base,
                            "Groupe critère": groupe_critere,
                            "Critère": critere,
                            "Note": note,
                            "Commentaires": None
                        })

        # Construction du DataFrame final
        self.df_stagiaires = DataFrame(df_long)

        # Réorganise les colonnes pour placer "NOM Prénom" juste après "Nom"
        colonnes = list(self.df_stagiaires.columns)
        if "NOM Prénom" in colonnes and "Nom" in colonnes:
            colonnes.remove("NOM Prénom")
            index_nom = colonnes.index("Nom")
            colonnes.insert(index_nom + 1, "NOM Prénom")
            self.df_stagiaires = self.df_stagiaires[colonnes]

        # Supprime les colonnes "Prénom" et "Nom" devenues inutiles
        self.df_stagiaires.drop(columns=["Prénom", "Nom"], inplace=True)



        # Étape 2 — On fait la jointure entre df_stagiaires et les données qui proviennent de l'extract IRIS Sessions
       
        # On récupère uniquement les colonnes souhaitées dans une vue pour faciliter le codage (c'est un alias)
        df_sessions_filtre = get_iris("Sessions").df[self._colonnes_sessions]
        
        # On fait la jointure entre df_stagiaires et df_sessions_filtre
        self.df_stagiaires = self.df_stagiaires.merge(
            df_sessions_filtre,
            left_on="Code session",
            right_on="Code IRIS",
            how="left"
        )

        # On réorganise les colonnes : d'abord celles de df_sessions puis celles de df_stagiaires
        colonnes_resultat = (
            df_sessions_filtre.columns.tolist() +  # colonnes de _df_sessions
            [col for col in self.df_stagiaires.columns if col not in df_sessions_filtre.columns]  # le reste (i.e. celles de df_stagiaires)
        )
        self.df_stagiaires = self.df_stagiaires[colonnes_resultat]
        #print("\nTypes de données de df_stagiaires dans _traiter_onglet_stagiaires :")
        #print(df_stagiaires.dtypes)
        #print(df_stagiaires)
        
        # On vire "Code session" qui est redondante avec "Code IRIS"
        self.df_stagiaires.drop(columns=["Code session"], inplace=True)
        
    def _maj_dataframes_eval_formation(self) -> None:
        """
        Met à jour le DataFrame partagé des évaluations de la formation courante 
        (self.eval_formation.df_csv et self.eval_formation.df_stagiaires) avec les données du CSV actuel.

        On écrira le fichier Excel ailleurs (fin du traitement de l'EvalStat ou d'une boucle si plusieurs)
        """

        # Si df_formation_csv est vide, il faut l'initialiser avec le premier df de la session ; sinon on concatène
        if (self.eval_formation.df_csv is None) or (self.eval_formation.df_csv.empty):
            self.eval_formation.df_csv = self.df_csv.copy()  # self._fe_evaluations_stagiaires._tableaux["CSV_stagiaires"]._df.copy()
            self.eval_formation.df_stagiaires = self.df_stagiaires.copy()  # self._fe_evaluations_stagiaires._tableaux["Stagiaires"]._df.copy()

        # df_formation_csv contient déjà des choses → Il faut concaténer les nouvelles données de cette session avec les anciennes
        else:
            if (not self.df_csv.empty) and (self.df_csv is not None): # Evite un future wanring de concaténer avec un df vide
                #self.df_csv = adapter_colonnes_dataframe_selon_modele(df_modele=self.eval_formation.df_csv, df_a_modifier=self.df_csv)
                self.eval_formation.df_csv = pd.concat([self.eval_formation.df_csv, self.df_csv], ignore_index=True)
            
            if (not self.df_stagiaires.empty) and (self.df_stagiaires is not None): # Evite un future wanring de concaténer avec un df vide
                #self.df_stagiaires = adapter_colonnes_dataframe_selon_modele(df_modele=self.eval_formation.df_stagiaires, df_a_modifier=self.df_stagiaires)
                self.eval_formation.df_stagiaires = pd.concat([self.eval_formation.df_stagiaires, self.df_stagiaires], ignore_index=True)



    # =================================================
    # === MÉTHODES EXTERNES - TRAITEMENT INDIVIDUEL ===
    # =================================================
    def verification_traitement_eval(self) -> Tuple[bool, str]:
        """
        Vérifie si l'on doit traiter l'EvalStat session à partir de :
            - la présence dans eval formation (code_IRIS ou chemin_csv) ;
            - la présence de code_IRIS dans IRIS_sessions.

        :param chemin_csv: chemin du csv de l'évaluation de la session, défaut = None
        :type chemin_csv: Optional[Path], optional
        :return:
        - Un bool pour dire si le traitement doit continuer 
            - True il faut traiter l'EvalStat (code IRIS ou CSV non détectés dans EvalStat + code IRIS présent dans IRIS sessions) ; 
            - False il ne faut pas traiter l'EvalStat.
        - un string avec le statut pour connaitre l'exclusion de traitement le cas échéant (pertinent ssi False ; si True on renvoie chaine vide).
        :rtype: Tuple[bool, str]
        """        
        return self.eval_formation.verifier_traitement_evalStat_session(code_IRIS=self.code_IRIS, chemin_csv=self._chemin_csv)

    def ouvrir_eval(self, session:Session_protocol, ouvrir_fe:bool=False) -> None:
        """
        Ouvre un EvalStat session déjà traité (i.e. déjà existant dans EvalStat formation).

        On stocke l'EvalStat dans session.eval

        Si l'objet EvalStat_session de la session n'est pas défini (None), alors on le crée et on l'affecte.

        :param session: La session à laquelle est affectée l'EvalStat
        :type session: Session_protocol
        :param ouvrir_fe: Si True, ouvre et charge l'Objet FichierExcel dans fe (i.e. si les données de eval formation ne suffisent pas)
        :type ouvrir_fe: bool, optional
        :return: l'EvalStat session
        :rtype: EvalStat_session
        """
        
        # On récupère le chemin du CSV depuis l'eval de la formation
        self._chemin_csv = self.eval_formation.chemin_csv_depuis_code_IRIS(self.code_IRIS)
        
        # On défini le statut
        self._statut = "Traité"  #"Traité - Code IRIS déjà dans l'évaluation de la formation" (je ne peux pas mettre plusieurs statuts traités car après dans bilanSession, dans la matrice "statut evalStat" "statut pour bilan" il me fait une correspondance 1 pour 1)
        
        # On ouvre le fichier si demandé
        if ouvrir_fe:
            self._fe = FichierExcel.depuis_fichier(chemin_fichier=self.chemin_excel)
        
        # Si session n'a pas encore son EvalStat de créé, alors on le crée
        if session.eval is None:
            session.eval = self   

    def traiter_eval(self, chemin_csv:Optional[Path|str]=None, ecrire_eval_formation:bool=True, ouvrir_dossier:bool=False) -> None:
        """
        Crée l'Excel EvalStat d'une session à partir d'un CSV (s'il n'est pas donné, on ouvre un filedialog).

        L'évaluation de la formation est mise à jour avec ces nouvelles données et est sauvée en fin de traitement selon le critère ecrire_eval_formation.

        :param chemin_csv: chemin du CSV à traiter. S'il est None, on ouvre un filedialog
        :type chemin_csv: Optional[Path|str]
        :param ecrire_eval_formation: Pour écrire physiquement l'Excel eval formation en fin de traitement. Si False, il devra être écrit ailleurs (à l'endroit où il y a la boucle pour du multi-traitement typiquement). Défaut = True.
        :type ecrire_eval_formation: bool
        :param ouvrir_dossier: Ouvre le répertoire de l'EvalStat généré. Defaut = False.
        :type ouvrir_dossier: bool
        """
        
        # On vérifie s'il est perinent de faire le traitement de l'EvalStat (on vérifie notemment sa déjà présence dans l'EvalStat formation)
        continuer, self._statut = self.verification_traitement_eval()
        if not continuer:
            return

        # Si l'argument est à None, on récupère la valeur existante dans l'objet (i.e. priorité à l'argument devant self)
        if chemin_csv is None :
            self._chemin_csv = EvalStat_session.filedialog_csv(
                    trigramme_formation=self.trigramme_formation, 
                    code_IRIS=self.code_IRIS)
            
            #vlog.log_erreur("Aucun chemin CSV fourni pour le traitement.", continuer=True)
            print(f"⚠️  Aucun chemin CSV fourni pour le traitement.")
            self._statut = "Exclu - Aucun CSV fourni"
            return 
        else:
            self._chemin_csv = chemin_vers_unc(chemin_csv)


        print("\n")
        timer.debut(f"{Style.BRIGHT}{Fore.YELLOW}Gestion du CSV {self._chemin_csv.name}") 

        # Étape 1 — On charge le CSV dans df_csv_stagiaires
        df_csv_stagiaires = self._charger_csv_stagiaire()
        # S'il y a eu un problème dans _charger_csv_stagiaire / le dataframe est vide (csv présent avec en-têtes mais sans ligne) → on saute la fin du traitement
        if (df_csv_stagiaires is None) or (df_csv_stagiaires.empty):
            return
        


        # Étape 2 — Générer l'excel des évaluations des stagiaire (_fe_evaluations_stagiaires : 2 onglets + TCD)
        # 2.1 : On crée le fichier Excel modèle depuis le modèle puis on affecte df_csv_stagiaire
        self._fe = FichierExcel.depuis_modele(chemin_modele=config.CHEMIN_MODELE_EXCEL_EVALUATIONS_STAGIAIRES, chemin_fichier_sauv=self.chemin_excel)
        self.df_csv = df_csv_stagiaires

        # 2.2 : On crée les dataframes
        self._traiter_df_csv()  # Traitement onglet CSV_stagiaires (import "direct" du CSV avec quelques traitements mineurs)
        self._traiter_df_stagiaires()  # Traitement seconde partie du dataframe du CSV

        # 2.3 : On écrit les dataframes dans les 2 onglets
        self._ecrit_df_et_sauve()
        self._statut = "Traité"



        # Étape 3 — On met à jour le DataFrame de formation partagé
        # 3.1 : On met à jour les DataFrame d'evalFormation
        self._maj_dataframes_eval_formation()

        # 3.2 : On sauvegade selon argument utilisateur ; normalement si et seulement si nous ne faisons pas un traitement en boucle (sinon on le fait en fin de traitement de boucle)
        if ecrire_eval_formation:
            self.eval_formation.ecrit_et_sauve_df_siModif()
    
    
    
        # Ouverture du dossier à la fin si demandé
        if ouvrir_dossier:
            utils_ouvrir_dossier(self.chemin_fe.parent)

        timer.fin()       

    def ouvrir_ou_traiter_eval(self, chemin_csv:Optional[Path|str]=None, ecrire_eval_formation:bool=True, ouvrir_dossier:bool=False, ouvrir_fe:bool=False) -> None:
        """
        Ouverture ou création du fichier Excel EvalStat.

        Si création :
            - crée l'instance EvalStat d'une session et traite cet EvalStat ;
            - crée l'Excel EvalStat d'une session à partir d'un CSV (s'il n'est pas donné, on ouvre un filedialog) ;
            - l'évaluation de la formation est mise à jour avec ces nouvelles données et est sauvée en fin de traitement selon le critère ecrire_eval_formation.

        :param chemin_csv: chemin du CSV à traiter. S'il est None, on ouvre un filedialog
        :type chemin_csv: Optional[Path|str]
        :param ecrire_eval_formation: Pour écrire physiquement l'Excel eval formation en fin de traitement. Si False, il devra être écrit ailleurs (à l'endroit où il y a la boucle pour du multi-traitement typiquement). Défaut = True.
        :type ecrire_eval_formation: bool
        :param ouvrir_dossier: Ouvre le répertoire de l'EvalStat généré. Defaut = False.
        :type ouvrir_dossier: bool
        :param ouvrir_fe: Si True, ouvre et charge l'Objet FichierExcel dans fe (i.e. si les données de eval formation ne suffisent pas)
        :type ouvrir_fe: bool, optional
        """
        # On vérifie la présence du code_session dans l'éval formation
        code_IRIS_present_eval_formation = self.eval_formation.verifier_presence_code_IRIS_dans_eval(self.code_IRIS)
        
        if code_IRIS_present_eval_formation :
            # === CAS avec code IRIS déjà traité (i.e. qui est dans eval formation) ===
            self.ouvrir_eval(session=self._session, ouvrir_fe=ouvrir_fe)
        
        else:
            # === CAS avec code IRIS non traité (i.e. qui n'est pas dans eval formation) ===
            self.traiter_eval(
                chemin_csv=chemin_csv,
                ecrire_eval_formation=ecrire_eval_formation,
                ouvrir_dossier=ouvrir_dossier
            )



    # =========================
    # === MÉTHODES EXTERNES ===
    # =========================
    
    @staticmethod
    def construire_dictionnaire_trigrammeFormation_codeIRIS_cheminsCSV_depuis_iterableCSV(chemins_csv: Iterable[Union[str, Path]], formation:Optional[Formation_protocol]=None) -> Dict[str, Dict[int, Path]]:
        """
        Construit un dictionnaire imbriqué de la forme {trigramme_formation: {code_IRIS: chemin_csv}} à partir de chemins CSV.

        On n'y inscrit que les sessions qui sont à traiter (on vérifie si non-présent dans eval formation).

        Ce dictionnaire s'emploie pour faciliter l'appel d'une boucle de traitement d'EvalStat

        :param chemins_csv: Un itérable de chemins CSV qui sont de type str ou Path.
        :type chemins_csv: Iterable[Union[str, Path]]
        :param formation: l'evalFormation pour vérifier si on doit traiter l'EvalStat + le trigramme de la formation qui sert à mieux pointer le répertoire pour sélectrionner le CSV.
        :type formation: Optional[Formation_protocol], Optional
        :return: Un dictionnaire imbriqué où les clés de premier niveau sont les trigrammes de formation,
            et les clés de second niveau sont les codes IRIS.
        :rtype: Dict[str, Dict[int, Path]]
        """
        dictionnaire = {}
        continuer = True
        statut = ""  # Non employé # TODO : je n'emploie pas la donnée de statut, mais en première réflexion ici ce n'est pas pertinent (uniquement pertinent lorsque l'on fait les bilans de sessions)

        for chemin in chemins_csv:
            # Si formation est donné, alors je peux faire des vérifications de pertinence de traiter l'EvalStat session en regardant s'il est déjà présent dans l'eval formation
            if formation is not None:
                continuer, statut = formation.eval.verifier_traitement_evalStat_session(chemin_csv=chemin)
                trigramme_formation = formation.trigramme_formation
            else:
                trigramme_formation = None

            if continuer:
                # Convertir le chemin en objet Path si ce n'est pas déjà fait
                chemin_path = Path(chemin) if isinstance(chemin, str) else chemin

                # Récupérer le trigramme de la formation
                if trigramme_formation is None:
                    trigramme_formation = recupere_trig_formation_depuis_chemin(chemin_path)

                # Récupérer le code IRIS
                code_IRIS = IRIS.extraire_code_IRIS_depuis_chemin(chemin_path)

                # Initialiser le sous-dictionnaire pour le trigramme si nécessaire
                if trigramme_formation not in dictionnaire:
                    dictionnaire[trigramme_formation] = {}

                # Ajouter l'entrée au sous-dictionnaire
                dictionnaire[trigramme_formation][code_IRIS] = chemin_path

        return dictionnaire

    # TODO : ce n'est plus employé
    @staticmethod
    def construire_dictionnaire_trigrammeFormation_codeIRIS_cheminsCSV_depuis_iterableCodesIRIS(codes_IRIS: Iterable[int], formation:Optional[Formation_protocol]=None) -> Dict[str, Dict[int, Path]]:
        """
        Pour une liste de codes IRIS :
            - Soit les EvalStat sont déjà traités (présents dans EvalStat foramtion) → On ouvre l'evalStat de la session (on remplit son statut) ;
            - Soit les EvalStat 

        Construit un dictionnaire imbriqué de la forme {trigramme_formation: {code_IRIS: chemin_csv}} à partir de codes IRIS.

        On n'y inscrit que les sessions qui sont à traiter (on vérifie si non-présent dans eval formation).

        Ce dictionnaire s'emploie pour faciliter l'appel d'une boucle de traitement d'EvalStat.

        :param codes_IRIS: Un itérable de codes IRIS de type int
        :type codes_IRIS: Iterable[int]
        :param formation: l'evalFormation pour vérifier si on doit traiter l'EvalStat + le trigramme de la formation qui sert à mieux pointer le répertoire pour sélectrionner le CSV.
        :type formation: Optional[Formation_protocol], Optional
        :return: Un dictionnaire imbriqué où les clés de premier niveau sont les trigrammes de formation,
            et les clés de second niveau sont les codes IRIS.
        :rtype: Dict[str, Dict[int, Path]]
        """
        dictionnaire = {}
        continuer = True
        statut = ""  # Non employé # TODO : je n'emploie pas la donnée de statut, mais en première réflexion ici ce n'est pas pertinent (uniquement pertinent lorsque l'on fait les bilans de sessions)

        
        # TODO : si l'utilisateur exclut un fichier, alors on doit le tracer qq par et le traiter eficacement (ex : si chemin =="", alors exclut par utilisateur) ; dans traitement eval stat vérifier effets de bord

        for code_IRIS in codes_IRIS:
            # Il me faut le trigramme de la formation pour vérifier la pertinence de traiter l'EvalStat session en regardant s'il est déjà présent dans l'eval formation
            # Si formation n'est pas donné, je récupère le trigramme depuis IRIS sessions
            #if formation is None:
            #    trigramme_formation = 


            # Si formation est donné, alors je peux faire des vérifications de pertinence de traiter l'EvalStat session en regardant s'il est déjà présent dans l'eval formation
            if formation is not None:
                continuer, statut = formation.eval.verifier_traitement_evalStat_session(code_IRIS=code_IRIS)
                trigramme_formation = formation.trigramme_formation
            else:
                trigramme_formation = None

            if continuer:
                # Convertir le chemin en objet Path si ce n'est pas déjà fait
                chemin_path = EvalStat_session.filedialog_csv(
                    trigramme_formation=trigramme_formation, 
                    code_IRIS=code_IRIS)

                # Récupérer le trigramme de la formation si non donné en argument
                if trigramme_formation is None:
                    trigramme_formation = recupere_trig_formation_depuis_chemin(chemin_path)

                # Initialiser le sous-dictionnaire pour le trigramme si nécessaire
                if trigramme_formation not in dictionnaire:
                    dictionnaire[trigramme_formation] = {}

                # Ajouter l'entrée au sous-dictionnaire
                dictionnaire[trigramme_formation][code_IRIS] = chemin_path

        return dictionnaire


    @staticmethod
    def construire_chemin_repertoire_csv(trigramme_formation:Optional[str]=None) -> Path:
        """
        Construit le chemin du répertoire des CSV d'évaluation des sessions de cette formation (à partir des données de la config REPERTOIRE_CSV_EVALUATIONS) :
            - le chemin est transformé en unc (s'il y a un raccourci lecteur réseau sur le poste de l'utilisateur on transforme en chemin réseau complet) ;
            - l'utilisateur peut optimiser le chemin (chemin le plus long entre l'attendu et ce qui existe).

        :param trigramme_formation: Trigramme de la formation. Défaut = None
        :type trigramme_formation: Optional[str], optional
        
        :return: Le chemin du répertoire des CSV d'évaluation des sessions de cette formation
        :rtype: Path
        """
        return construire_chemin_config(
            chemin_a_completer = config.REPERTOIRE_CSV_EVALUATIONS,
            trigramme_formation = trigramme_formation
        )
    
    

    # =============
    # === POPUP ===
    # =============
    def _filedialog_csv(self) -> Path | None:
        """
        Ouvre un filedialog pour demander à l'utilisateur de sélectionner un CSV.
        On pointe au mieux sur le répertoire des CSV de cette formation pour la boîte de dialogue.
        """
        return self.filedialog_csv(self.trigramme_formation, self.code_IRIS)

    @staticmethod
    def filedialog_csv(trigramme_formation:Optional[str]=None, code_IRIS:Optional[int]=None, unc:bool=True) -> Path | None:
        """
        Ouvre un filedialog pour demander à l'utilisateur de sélectionner un CSV.

        On pointe au mieux sur le répertoire des CSV de cette formation pour la boîte de dialogue.

        Le trigramme permet d'optimiser le répertoire de recherche (non obligatoire).

        Le code IRIS permet d'être spécifié dans l'en-tête du filedialog (non obligatoire).

        :param trigramme_formation: Trigramme de la formation. Permet d'optimiser le répertoire de recherche (non obligatoire).
        :type trigramme_formation: Optional[str]
        :param code_IRIS: Code IRIS de la session. Permet d'être spécifié dans l'en-tête du filedialog (non obligatoire).
        :type code_IRIS: Optional[int]
        :param unc: Si True, le chemin renvoyé sera avec le chemin réseau complet et pas le raccourci utilisateur lecteur réseau
        :type unc: bool, optional
        :return: Description
        :rtype: Path | None
        """

        # Si on a un trigramme de formation, alors on est en mesure de trouver un chemin optimisé
        #chemin_repertoire_csv = EvalStat_formation.construire_chemin_eval_formation(trigramme_formation=trigramme_formation)
        #if trigramme_formation:
        #    chemin_repertoire_csv = optimiseCheminRepertoire(
        #        config.format_path(config.REPERTOIRE_CSV_EVALUATIONS, trigramme_formation=trigramme_formation)
        #        )
        #else:
        #    chemin_repertoire_csv = optimiseCheminRepertoire(config.REPERTOIRE_FORMATION.parent)  #Path.cwd()  
            
        # Adaptation de l'intitulé de l'en-tête de la popup
        fin_titre = f"de la session {code_IRIS}" if code_IRIS is not None else "désiré"

        # Ouverture popup
        return choisir_fichier(titre=f"Sélectionner le fichier EvalStat stagiaire {fin_titre}",
                        types_fichiers=[("Fichiers CSV", "*.csv")],
                        #dossier_initial=chemin_repertoire_csv,
                        dossier_initial=EvalStat_session.construire_chemin_repertoire_csv(trigramme_formation=trigramme_formation),
                        obligatoire=False,
                        texte_bouton_choisir="Choisir CSV à nouveau",
                        texte_bouton_aucun="Pas de CSV pour cette session",
                        unc=unc
                        )
        


    # =========================
    # === GETTERS / SETTERS ===
    # =========================
    # Lié à EvalStat_session
    """
    # Propriété de classe → @property ne marche que pour une instance → Mettre en propriété de classe
    @property
    def mapping_statuts(self) -> dict[str, str]:
        return self._mapping_statuts
    """
    @property
    def chemin_csv(self) -> Path:
        return self._chemin_csv

    @property
    def chemin_excel(self) -> Path:
        return self._chemin_csv.with_suffix(".xlsx") 

    @property
    def statut(self) -> str:
        return self._statut
   
    @cached_property
    def chemin_repertoire_csv(self) -> Path:
        return EvalStat_session.construire_chemin_repertoire_csv(trigramme_formation=self.trigramme_formation)



    # Lié à Session
    @property
    def trigramme_formation(self) -> str|None:
        # On essaye d'abord à partir de formation
        if self._session.trigramme_formation is not None :
            return self._session.trigramme_formation
        
        # Sinon tente de le récupérer depuis le chemin de l'évaluation stagiaire 
        elif self.chemin_fe is not None:
            return recupere_trig_formation_depuis_chemin(self.chemin_fe)
        
        # Sinon tente de le récupérer depuis le chemin de l'évaluation formation 
        elif self.eval_formation.chemin_fe is not None:
            return recupere_trig_formation_depuis_chemin(self.eval_formation.chemin_fe)
    
    @property
    def code_IRIS(self) -> int|None:
        return self._session.code_IRIS if self._session is not None else None

    @property
    def eval_formation(self) -> EvalStat_formation:
        return self._session.eval_formation




