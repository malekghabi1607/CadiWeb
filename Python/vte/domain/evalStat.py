from __future__ import annotations
from pathlib import Path
from typing import Optional, Protocol

import pandas as pd

from vte.core import config
#from vte.domain.formation import Formation → Ref circulaire
#from vte.domain.session import Session → Ref circulaire
from vte.core.iris_referentiel import *
from vte.utils.office import FichierExcel
from vte.utils.utils import *
from vte.utils.utils_instn import recupere_trig_formation_depuis_chemin


class Formation_protocol(Protocol):
    @property
    def trigramme_formation(self) -> str: ...

    @property
    def eval(self) -> EvalStat_formation|None: ...
    
    #@property
    #def chemin_dossier_eval(self) -> Path: ...


class Session_protocol(Protocol):
    @property
    def code_IRIS(self) -> int|None: ...
        
    @property
    def trigramme_formation(self) -> str: ...

    @property
    def eval_formation(self) -> EvalStat_formation: ...



# TODO : j'ai du retype de code IRIS en str : self.df_stagiaires["Code IRIS"] = self.df_stagiaires["Code IRIS"].astype(str)  # Retype "Code IRIS"
# ======================================================================================
# CLASSE EVALSTAT
# Objet fichier EvalStat + logique directement liée au fichier
# ======================================================================================
class EvalStat:
    """
    Classe principale pour le traitement des évaluations stagiaires individuelles.
    Gère la lecture des fichiers CSV stagiaires, la mise à jour du fichier Excel
    de la formation, et les interactions éventuelles avec l'extract IRIS.
    """

    # === VARIABLES DE CLASSE COMMUNES À TOUTES LES INSTANCES ===

    
    # → Viendra de iris_sessions = get_iris("Sessions")
    #
    # Extract IRIS Sessions (R04110)
    #_chemin_IRIS_sessions:Path=None
    #_fe_IRIS_sessions: Optional[FichierExcel] = None  # Fichier Excel contenant l'extract IRIS Sessions (ou celles de la période en cours)
    #_df_IRIS_sessions: Optional[pd.DataFrame] = None  # Alias du dataframe



    # → Viendra de Formation

    # Fichier Excel d'évaluation de formation (partagé pendant un contexte Contexte_formation)
    #_CHEMIN_MODELE_EXCEL_EVALUATIONS_STAGIAIRES: Path
    #_trigramme_formation: Optional[str] = None
    #_fe_evaluations_formation: Optional[FichierExcel] = None
    #_df_evaluations_formation: Optional[pd.DataFrame] = None
    #_supprimeDonneesEtRemplace_evaluations_formation: Optional[bool]  = None






    # === Colonnes du CSV selon traitement à avoir ===
    # Colonnes descriptives à recopier
    _colonnes_csv_fixes = [
        "Chemin fichier CSV", "Prénom", "Nom", "Entreprise", "Code session"]

    # Colonnes avec note/commentaire en binôme
    _colonnes_csv_avec_commentaires = [
        "Accueil, organisation et qualité des informations délivrées",
        "Conseils et orientation avant l'inscription",
        "Informations après l'inscription",
        "Accueil à l'arrivée sur site",
        "Prise en compte de vos besoins et attentes",
        "Qualité des animations",
        "Logique d'enchainement des interventions",
        "Qualité des supports de cours utilisés",
        "Qualité des moyens pédagogique",
        "Accès aux outils digitaux",
        "Satisfaction globale",
        "Avez-vous d'autres besoins de formation ?"]

    # Colonnes à valeur texte seule
    _colonnes_csv_commentaires_seuls = [
        "Comment avez-vous connu cette formation ?",
        "Commentaires, remarques, suggestions"]

    # Colonnes note seule (il se trouve que je vais aussi devoir convertir le booléen)
    _colonnes_csv_bool = [
        "Recommanderiez-vous cette formation ?"]
    
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

    # ==================================================================================
    # CONSTRUCTEUR
    # ==================================================================================
    def __init__(self) -> None:
        self._fe:Optional[FichierExcel] = None  # Objet Excel contenant les données EvalStat stagiaire individuel


    # ==================================================================================
    # GETTERS / SETTERS
    # ==================================================================================
    @property
    def fe(self) -> FichierExcel|None:
        return self._fe
    
    @property
    def df_stagiaires(self) -> pd.DataFrame|None:
        return self._fe.get_df_tableau("Stagiaires")
    
    @property
    def df_csv(self) -> pd.DataFrame|None:
        return self._fe.get_df_tableau("CSV_stagiaires")

    @property
    def chemin_fe(self) -> Path|None:
        return self._fe.chemin_fichier
    


    #@chemin_fe.setter
    #def chemin_fe(self, valeur:Path):
    #    self._fe.chemin_fichier = valeur



# ======================================================================================
# CLASSE EVALSTAT_SESSION
# ======================================================================================
class EvalStat_session(EvalStat):

    # =====================
    # === CONSTRUCTEURS ===
    # =====================
    def __init__(self, session:Session_protocol):
        # On initialise la classe mère
        super().__init__()

        self._session = session  # C'est un protocol pour éviter les références circulaires


    @classmethod
    def depuis_chemin_csv(cls, session:Session_protocol, chemin_csv:Path|str, chemin_IRIS_sessions:Optional[Path]=None, ouvrirDossier:bool=False) -> EvalStat_session:
        """
        Crée et traite une instance d'EvalStat à partir d'un fichier CSV.

        Peut être utilisée directement dans un contexte `ContexteFormation`.

        Exemple :
            with ContexteFormation("ABC"):
                EvalStat.depuis_chemin_csv_evaluations_stagiaires(Path("eval_R04110.csv"))
        """
        # On créée l'instance
        instance = cls(session)
        
 





        
        return instance



    # TODO : a priori : plus besoin
    @classmethod
    def depuis_chemin_csv_newOLD(cls, session:Session_protocol, chemin_csv:Optional[Path|str] = None, chemin_IRIS_sessions:Optional[Path]=None, ouvrirDossier:bool=False) -> EvalStat_session:
        """
        Permet de traiter un CSV stagiaire d'une session.
        Intègre le CSV natif dans un modèle Excel plus user-friendly
        
        :param chemin_csv: Chemin du CSV que l'on souhaite traiter. Si non présent, l'utilisateur le pointera avec une filedialog.
        :type chemin_csv: Optional[Path | str]
        :param chemin_IRIS_sessions: Permet de forcer un chemin pour IRIS_sessions plutôt que de prendre celui par défaut
        :type chemin_IRIS_sessions: Optional[Path]
        :param ouvrirDossier: Permet d'ouvrir le répertoire à l'utilisateur en fin de traitement (jamais exploité)
        :type ouvrirDossier: bool
        :return: Un EvalStat_session
        :rtype: EvalStat_session
        """    
        instance = cls(session)
        
        # Soit on a déjà un chemin input CSV, soit on va le pointer manuellement
        if chemin_csv is None:
            chemin_csv = instance._filedialog_csv()
            if chemin_csv is None:
                vlog.log_erreur(f"Pas de fichier CSV pour la session {instance.code_IRIS}", continuer=True)
                return

        # Si un chemin_IRIS_sessions est donné en argument (i.e. not Null), alors c'est pour employer un fichier qui n'est pas celui par défaut (plus récent dans le répertoire idoine)
        # TODO pas besoin de faire le get ici, mais peut-être conserver en mémoire chemin_IRIS_sessions
        #iris_sessions = iris_referentiel.get_iris("Sessions", chemin_IRIS_sessions)



    # ==================================================================================
    # MÉTHODES D’INSTANCE - TRAITEMENT INDIVIDUEL
    # ==================================================================================
    








    # ==================================================================================
    # POPUP
    # ==================================================================================
    def _filedialog_csv(self) -> Path | None:
        """
        Ouvre un filedialog pour demander à l'utilisateur de sélectionner un CSV.
        On pointe au mieux sur le répertoire des CSV de cette formation pour la boîte de dialogue.
        """

        # Si on a un trigramme de formation, alors on est en mesure de trouver un chemin optimisé
        if self.trigramme_formation:
            chemin_repertoire_csv = optimiseCheminRepertoire(
                config.format_path(config.REPERTOIRE_CSV_EVALUATIONS, trigramme_formation=self.trigramme_formation)
                )
        else:
            chemin_repertoire_csv = optimiseCheminRepertoire(config.REPERTOIRE_FORMATION.parent)  #Path.cwd()  
            
        # Adaptation de l'intitulé de l'en-tête de la popup
        fin_titre = f"de la session {self.code_IRIS}" if self.code_IRIS is not None else "désiré"

        # Ouverture popup
        return choisir_fichier(titre=f"Sélectionner le fichier EvalStat stagiaire {fin_titre}",
                        types_fichiers=[("Fichiers CSV", "*.csv")],
                        dossier_initial=chemin_repertoire_csv,
                        obligatoire=False,
                        texte_bouton_choisir="Choisir CSV à nouveau",
                        texte_bouton_aucun="Pas de CSV pour cette session"
                        )

    # ==================================================================================
    # GETTERS / SETTERS
    # ==================================================================================
    


    @property
    def trigramme_formation(self) -> str|None:
        # On essaye d'abord à partir de formation
        if self.trigramme_formation is not None :
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
        return Session_protocol.eval_formation






# ======================================================================================
# CLASSE EVALSTAT_FORMATION
# ======================================================================================
class EvalStat_formation(EvalStat):
    """
    Il doit pouvoir :
        - Contenir un EvalStat_formation déjà traité (fe ; fait une ref à formation)
        - Créer un EvalStat_formation :
           - vérifier si déjà existant
    """

    # =====================
    # === CONSTRUCTEURS ===
    # =====================
    def __init__(self, formation:Formation_protocol):
        # On initialise la classe mère
        super().__init__()

        #self._fe → classe mère
        self._formation = formation  # Protocol pour éviter les références circulaires

        self._df_initial_hash:Optional[str] = None  # hash du df initial pour savoir s'il a été modifié, auquel cas on sauvegardera à la fin
        self._supprimeEtRemplace_donneesEval:Optional[bool] = None



        timer.debut(f"Ouverture ou création du fichier Excel de la formation {self.trigramme_formation}")

        self._ouvrir_ou_creer_evaluations_formation()

        # On retype en str la colonne Code IRIS
        # TODO : vraiment ???
        #self._df_evaluations_formation = self._fe_evaluations_formation._tableaux["Stagiaires"].df  # Alias
        #self._df_evaluations_formation["Code IRIS"] = self._df_evaluations_formation["Code IRIS"].astype(str)
        self.df_stagiaires["Code IRIS"] = self.df_stagiaires["Code IRIS"].astype(str)

        # On fait un hash du df pour savoir s'il a été modifié, auquel cas on sauvegardera à la fin
        self._df_initial_hash = hash_df(self.df_stagiaires)
        timer.fin()





    def _ouvrir_ou_creer_evaluations_formation(self):  # -> Tuple[FichierExcel, bool]:
        """
        Ouvre ou crée le fichier Excel d'évaluations d'une formation.

        Construit le chemin vers le fichier d'évaluations correspondant au trigramme de la formation.
        Si ce fichier existe, il est ouvert et les données des stagiaires sont chargées dans un DataFrame.
        Sinon, un nouveau fichier est créé à partir d'un modèle, et les données seront à initialiser.
        """
        """
        Args:
            trigramme_formation (str): Trigramme de la formation.
        Returns:
            - FichierExcel: Objet FichierExcel ouvert.
            - Un booléen indiquant si les anciennes données doivent être supprimées et remplacées 
                (`True` si nouveau fichier créé, `False` sinon).
        """
        
        # On définit le chemin de destination du fichier excel d'évaluation depuis le fichier config
        chemin_excel_eval = config.format_path(config.CHEMIN_EXCEL_EVALUATIONS_FORMATION, trigramme_formation=self.trigramme_formation)
        #)
        
        # On regarde si Evaluation-Stagiaires-Global-XXX.xlsx existe
        # Auquel cas on l'ouvre et on ne supprimera pas les données existantes
        if chemin_excel_eval.is_file():
            print("Ouverture EvalStat formation existant")
            # Alors on l'ouvre
            self._fe = FichierExcel.depuis_fichier(chemin_fichier=chemin_excel_eval)
            #self._fe_evaluations_formation._tableaux["CSV_stagiaires"].charge_df()
            #self._fe_evaluations_formation._tableaux["Stagiaires"]._df['Trigramme formation'] = self._fe_evaluations_formation._tableaux["Stagiaires"]._df['Trigramme formation'].astype(str)

            # On récupère la liste des sessions déjà intégrées (liste des codes et des chemins) → Ce sera pour écrire dans le tkinter
            dico_sessionsDejaTraitees = dict(
                self.df_stagiaires[self.df_stagiaires["Trigramme formation"] == self.trigramme_formation]   # 1. filtre sur le trigramme
                .drop_duplicates(subset=["Code IRIS", "Chemin fichier CSV"])                                # 2. élimine les doublons
                [["Code IRIS", "Chemin fichier CSV"]]                                                       # 3. sélection des colonnes
                .values                                                                                     # 4. valeurs du DF
                )        

            # Il ne faudra pas supprimer les anciennes données de fe_evaluations_formation
            self._supprimeEtRemplace_donneesEval = False

            vlog.ajouter_message("Ouverture EvalStat Global formation", self.chemin_fe, style=["vert"])
       
        # Sinon on crée l'évaluation depuis le modèle et on supprimera les anciennes données
        else:
            print("Création nouvel EvalStat formation")
            # On vérifie que le répertoire dédié existe sinon on le créée : \\instnt\partage\FORMATIONS_C\###\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\
            chemin_excel_eval.parent.mkdir(parents=True, exist_ok=True)

            # On créée le fichier excel à partir du modèle
            self._fe = FichierExcel.depuis_modele(
                chemin_modele = config.CHEMIN_MODELE_EXCEL_EVALUATIONS_STAGIAIRES, 
                chemin_fichier_sauv = chemin_excel_eval
                )
            
            # Il faudra supprimer les anciennes données de fe_evaluations_formation
            self._supprimeEtRemplace_donneesEval = True

            vlog.ajouter_message("Création EvalStat Global formation", self.chemin_fe, style=["vert"])

        #return self._fe_evaluations_formation, self._df_evaluations_formation, supprimeDonneesEtRemplace
        #return fe, supprimeDonneesEtRemplace


    # ==================================================================================
    # GETTERS / SETTERS
    # ==================================================================================

    @property
    def trigramme_formation(self) -> str:
        return self._formation.trigramme_formation




class A_employer():
    def _sauver_excel_evaluations_formation(self, nouveau_chemin_fichier:Optional[Path] = None, fermer_fichier:Optional[bool]=True) -> None:
        """
        Sauvegarde le fichier Excel des évaluations de la formation.

        :Note: Il faut sauvegarder à la sortie du contexte (car fin du traitement de la ou des sessions d'une même formation)
        """
        print(f"\n\n{Style.BRIGHT}{Fore.RED}Écriture du fichier global des évaluations de la formation {self._trigramme_formation}")

        # On écrit et on sauve
        self._fe_evaluations_formation._tableaux["CSV_stagiaires"].ecrit_dataFrame_dans_tableauStructure(supprimeDonneesEtRemplace = self._supprimeDonneesEtRemplace_evaluations_formation)
        self._fe_evaluations_formation._tableaux["Stagiaires"].ecrit_dataFrame_dans_tableauStructure(supprimeDonneesEtRemplace = self._supprimeDonneesEtRemplace_evaluations_formation)
        #self._fe_evaluations_formation._tableaux["CSV_stagiaires"].charge_df()
        #self._fe_evaluations_formation._tableaux["Stagiaires"].charge_df()

        #self._df_evaluations_formation = instance._fe_evaluations_formation._tableaux["Stagiaires"]._df  # Alias

        self._fe_evaluations_formation.save(nouveau_chemin_fichier)

        # Actualiser TCD
        self._fe_evaluations_formation.actualiser_TCD()

        # On ferme si demandé
        if fermer_fichier:
            self._fe_evaluations_formation.close()
