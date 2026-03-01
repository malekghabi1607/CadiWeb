from __future__ import annotations

from collections import defaultdict
from pathlib import Path
from typing import Optional

from vte import config
from vte.utils import *
from vte.utils_instn import recupere_trig_formation_depuis_chemin, demander_code
from vte.office import FichierExcel
from vte.iris import IRIS


# ======================================================================================
# CLASSE EVALSTAT
# ======================================================================================
class EvalStat:
    """
    Classe principale pour le traitement des évaluations stagiaires individuelles.
    Gère la lecture des fichiers CSV stagiaires, la mise à jour du fichier Excel
    de la formation, et les interactions éventuelles avec l'extract IRIS.
    """

    # === VARIABLES DE CLASSE COMMUNES À TOUTES LES INSTANCES ===

    # Extract IRIS Sessions (R04110)
    _chemin_IRIS_sessions:Path=None
    _fe_IRIS_sessions: Optional[FichierExcel] = None  # Fichier Excel contenant l'extract IRIS Sessions (ou celles de la période en cours)
    _df_IRIS_sessions: Optional[pd.DataFrame] = None  # Alias du dataframe

    # Fichier Excel d'évaluation de formation (partagé pendant un contexte Contexte_formation)
    _CHEMIN_MODELE_EXCEL_EVALUATIONS_STAGIAIRES: Path
    _trigramme_formation: Optional[str] = None
    _fe_evaluations_formation: Optional[FichierExcel] = None
    _df_evaluations_formation: Optional[pd.DataFrame] = None
    _supprimeDonneesEtRemplace_evaluations_formation: Optional[bool]  = None

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
    def __init__(self, codeIRIS: Optional[str] = None, chemin_csv: Optional[Path] = None) -> None:
        """
        Initialise une instance EvalStat individuelle (pour un stagiaire/session).

        Args:
            codeIRIS (str | None): Code IRIS de la session.
            chemin_csv (Path | None): Chemin vers le CSV d'évaluation du stagiaire.
        """
        # Variables propres à un EvalStat individuel
        self._codeIRIS: Optional[str] = codeIRIS
        self._chemin_csv_evaluations_stagiaires: Optional[Path] = chemin_csv  # Fichier CSV d'entrée
        self._fe_evaluations_stagiaires: Optional[FichierExcel] = None  # Objet Excel contenant les données EvalStat stagiaire individuel

        # === Résultat du traitement ===
        self._statut_csv: Optional[str] = None  # ex: "Traité", "Exclu - CSV déjà dans fichier global", "Exclu - Code IRIS pas dans Extract IRIS sessions", "Exclu - Problème lecture CSV", "Exclu - CSV vide / Aucun retour"

    # ==================================================================================
    # CONSTRUCTEURS ALTERNATIFS
    # ==================================================================================   
    @classmethod
    def depuis_chemin_csv_evaluations_stagiaires(cls, chemin_csv_stagiaires:Optional[Path|str] = None, trigramme_formation:Optional[str] = None, chemin_IRIS_sessions:Optional[Path]=None, ouvrirDossier:bool=False) -> EvalStat:
        """
        Permet de traiter un CSV stagiaire d'une session.
        Intègre le CSV natif dans un modèle Excel plus user-friendly
        
        :param cls: Description
        :param chemin_csv_stagiaires: Chemin du CSV que l'on souhaite traiter. Si non présent, l'utilisateur le pointera avec une filedialog. Dans ce cas il faut un trigramme formation.
        :type chemin_csv_stagiaires: Optional[Path | str]
        :param trigramme_formation: Trigramme du CSV que l'on souhaite traiter. Si CSV présent, alors on essayera de le déduire du chemin du CSV.
        :type trigramme_formation: Optional[str]
        :param chemin_IRIS_sessions: Permet de forcer un chemin pour IRIS_sessions plutôt que de demander à l'utilisateur de le pointer avec un fileDialog
        :type chemin_IRIS_sessions: Optional[Path]
        :param ouvrirDossier: Permet d'ouvrir le répertoire à l'utilisateur en fin de traitement (jamais exploité)
        :type ouvrirDossier: bool
        :return: Un EvalStat
        :rtype: EvalStat
        """

        # Soit on a déjà un chemin, soit on va pointer le csv manuellement
        if chemin_csv_stagiaires is None:
            chemin_csv_stagiaires = cls._filedialog_csv(trigramme_formation=trigramme_formation)
            if chemin_csv_stagiaires is None:
                vlog.log_erreur("Pas de fichier CSV", continuer=True)
                return

        # On convertit en Path   
        if isinstance(chemin_csv_stagiaires, str):
            chemin_csv_stagiaires = Path(chemin_csv_stagiaires)

        # On récupère le trigramme de la formation depuis le chemin du CSV
        if trigramme_formation is None:
            trigramme_formation = recupere_trig_formation_depuis_chemin(chemin_csv_stagiaires)

        # Si on force un chemin pour chemin_IRIS_sessions, alors on renseigne la valeur
        if chemin_IRIS_sessions is not None:
            cls._chemin_IRIS_sessions = chemin_IRIS_sessions

        with Contexte_formation(trigramme_formation):
            EvalStat.depuis_chemin_csv_evaluations_stagiaires_avec_contexte(chemin_csv_stagiaires=Path(chemin_csv_stagiaires),
                                                                            ouvrirDossier=ouvrirDossier
                                                                            )
    
    @classmethod
    def depuis_chemin_csv_evaluations_stagiaires_avec_contexte(cls, chemin_csv_stagiaires:Path|str, ouvrirDossier:bool=False) -> EvalStat:
        """
        Crée et traite une instance d'EvalStat à partir d'un fichier CSV.

        Peut être utilisée directement dans un contexte `ContexteFormation`.

        Exemple :
            with ContexteFormation("ABC"):
                EvalStat.depuis_chemin_csv_evaluations_stagiaires(Path("eval_R04110.csv"))
        """
        if cls._fe_evaluations_formation is None:
            raise RuntimeError(
                "Le fichier Excel de la formation n'est pas ouvert.\n"
                "Utilisez un bloc 'with ContexteFormation(trigramme):' avant d'appeler cette méthode."
            )
        
        # On créée l'instance
        instance = cls()

        # si le chemin est avec un raccourci réseau alors on récupère le chemin en entier
        instance._chemin_csv_evaluations_stagiaires = chemin_vers_unc(Path(chemin_csv_stagiaires))   

        # On génère le fichier Excel du CSV à partir du modèle
        instance.traiter(ouvrirDossier=ouvrirDossier)
        
        return instance

    @classmethod
    def depuis_tuple_csv_stagiaires(cls, tuple_csv_stagiaires:Tuple[Path|str], ouvrirDossier:bool=False) -> None:
        """
        A partir d'un tuple de chemins de CSV stagiaire (il peut il y avoir plusieurs trigrammes de formations différents)
        Permet de générer :
           - le fichier excel stagiaires de chaque session (via le CSV)
           - le fichier excel stagiaires de chaque formation (celui qui concatène tous les CSV d'une session) [Il est créé ou on l'append avec les nouvelles valeurs]
        
        On exclue du traitement les chemin_csv_session qui sont déjà dans le FE formation (on considère que le CSV a déjà été traité)
        """

        # On convertit le tuple de strings en dictionnaire avec les trigrammes formation en clef
        dico_chemins_csv_session = defaultdict(list)  #Dictionnaire spécial : lorsqu’on accède à une clé qui n’existe pas encore, il va automatiquement créer une nouvelle entrée avec une valeur par défaut, ici une liste vide (list())
        for chemin in tuple_csv_stagiaires:
            chemin = Path(chemin)
            trigramme_formation = recupere_trig_formation_depuis_chemin(chemin)
            dico_chemins_csv_session[trigramme_formation].append(chemin)
        dico_chemins_csv_session = dict(dico_chemins_csv_session)  # Optionnel : conversion en dict normal
        
        EvalStat.depuis_dico_csv_stagiaires(dico_chemins_csv_session=dico_chemins_csv_session, ouvrirDossier=ouvrirDossier)

    @classmethod
    def depuis_dico_csv_stagiaires(cls, dico_chemins_csv_session: dict[str, list[Path]], ouvrirDossier:bool=False) -> Tuple[Optional[FichierExcel], Optional[dict[str, dict[str, str]]]]:
        """
        Traite un ensemble de fichiers CSV groupés par formation.

        Args :
            dico_chemins_csv_session (dict[str, list[Path]]):
                Dictionnaire {trigramme_formation: [liste_de_csv]}.

        Return :
            Une liste de dictionnaires {
                es._codeIRIS,
                    {
                    "fichier": chemin_csv.name,
                    "statut": es._statut_csv
                    }      
                }
            Je pourrai accéder à la valeur par nom_dico[codeIRIS]["statut"]  
        Exemple :
            EvalStat.depuis_tuple_csv_stagiaires({
                "ABC": [Path("R04110.csv"), Path("R04112.csv")],
                "DEF": [Path("R04201.csv")]
            })
        """

        # Cas avec dictionnaire non-vide
        if dico_chemins_csv_session :
            statuts_csv:dict[str, dict[str, str]] = {}
            for trigramme_formation, chemins_csv in dico_chemins_csv_session.items():
                print(f"\n\n{Style.BRIGHT}{Fore.RED}Gestion des formations {trigramme_formation}")
                with Contexte_formation(trigramme_formation) as ctx:
                    for chemin_csv in chemins_csv:
                        #cls.depuis_chemin_csv_evaluations_stagiaires(chemin_csv)
                        es = EvalStat.depuis_chemin_csv_evaluations_stagiaires_avec_contexte(chemin_csv_stagiaires=Path(chemin_csv),
                                                                                ouvrirDossier=ouvrirDossier
                                                                                )
                        statuts_csv[es._codeIRIS] = {
                            "fichier": chemin_csv.name,
                            "statut": es._statut_csv
                            }
                        print(f"Fin traitement : {es._codeIRIS}\t{chemin_csv.name}\t{es._statut_csv}")
            return ctx._fe_evaluations_formation, statuts_csv
        
        # Cas vec dictionnaire vide
        else:
            return None, None


        

    @classmethod
    def depuis_liste_codes_IRIS(cls, liste_codes_IRIS:list[int]|int, fe_IRIS_sessions:FichierExcel=None, ouvrirDossier:bool=False) -> Tuple[Optional[FichierExcel], Optional[dict[str, dict[str, str]]]]:
        """
        A partir d'une liste de codes IRIS (il peut il y avoir plusieurs trigrammes de formations différents)
        Permet de générer :
           - le fichier excel stagiaires de chaque session (via le CSV)
           - le fichier excel stagiaires de chaque formation (celui qui concatène tous les CSV d'une session) [Il est créé ou on l'append avec les nouvelles valeurs]
        
        On exclue du traitement les chemin_csv_session qui sont déjà dans le FE formation (on considère que le CSV a déjà été traité)

        
        Return :
            Une liste de dictionnaires {
                es._codeIRIS,
                    {
                    "fichier": chemin_csv.name,
                    "statut": es._statut_csv
                    }      
                }
            Je pourrai accéder à la valeur par nom_dico[codeIRIS]["statut"]  
        """
        if fe_IRIS_sessions is not None :
            cls._fe_IRIS_sessions = fe_IRIS_sessions
            # Conversion du code IRIS en string pour jointure future
            cls._df_IRIS_sessions = cls._fe_IRIS_sessions._tableaux["Sessions"]._df
            cls._df_IRIS_sessions["Code IRIS"] = cls._df_IRIS_sessions["Code IRIS"].astype(str)
        
        # Dictionnaire de retour
        statuts_csv:Optional[dict[str, dict[str, str]]] = {}

        # Si en entrée on a un entier, alors on convertit en liste
        if isinstance(liste_codes_IRIS, int):
            liste_codes_IRIS = [liste_codes_IRIS]

        # On lit le fichier extract IRIS
        cls._charger_IRIS_sessions()


        


        # On crée le dictionnaire des csv stagiaires
        dico_chemins_csv_session = defaultdict(list)  #Dictionnaire spécial : lorsqu’on accède à une clé qui n’existe pas encore, il va automatiquement créer une nouvelle entrée avec une valeur par défaut, ici une liste vide (list())
        for code_IRIS in liste_codes_IRIS:
            #vlog.print("Info", f"Recherche infos code IRIS {code_IRIS}")

            # On récupère le trigramme formation depuis le code IRIS (extract IRIS sessions)
            #res = cls._df_IRIS_sessions.loc[cls._df_IRIS_sessions["Code IRIS"] == str(code_IRIS), "Trigramme formation"]
            #trigramme_formation = res.iloc[0] if not res.empty else demander_code(typeCode="Trigramme formation", info=code_IRIS)
            trigramme_formation = rechercheX_dataframe(
                dataframe=cls._df_IRIS_sessions, 
                colonne_recherche="Code IRIS", 
                valeur_recherche=str(code_IRIS), 
                colonne_souhaitee="Trigramme formation", 
                fallback=lambda: demander_code(typeCode="Trigramme formation", info=code_IRIS)
            )

            # On récupère le CSV depuis le fichier Excel global de la formation
            with Contexte_formation(trigramme_formation):
                #res = cls._df_evaluations_formation.loc[cls._df_evaluations_formation["Code IRIS"] == str(code_IRIS), "Chemin fichier CSV"]
                #if not res.empty:
                #    chemin_csv = res.iloc[0]
                #else:
                #    # Il faut que l'utilisateur pointe le CSV de la session correspondante
                #    chemin_csv = cls._filedialog_csv(code_IRIS=code_IRIS, trigramme_formation=trigramme_formation)
                chemin_csv = rechercheX_dataframe(
                    cls._df_evaluations_formation, 
                    "Code IRIS", 
                    str(code_IRIS), 
                    "Chemin fichier CSV",
                    fallback=lambda: cls._filedialog_csv(code_IRIS=code_IRIS, trigramme_formation=trigramme_formation)
                    )

            if chemin_csv is not None:
                chemin_csv = Path(chemin_csv)
                dico_chemins_csv_session[trigramme_formation].append(chemin_csv)
            else:
                statuts_csv[trigramme_formation] = {
                    "fichier": "Fichier non existant",
                    "statut": "Exclu - Fichier non existant"
                    }
        dico_chemins_csv_session = dict(dico_chemins_csv_session)  # Optionnel : conversion en dict normal

        # On lance le traitement des EvalStat 
        fe_evaluations_formation, statuts_csv = EvalStat.depuis_dico_csv_stagiaires(dico_chemins_csv_session=dico_chemins_csv_session, ouvrirDossier=ouvrirDossier)

        return fe_evaluations_formation, statuts_csv

    # ==================================================================================
    # MÉTHODES DE CLASSE - CHARGEMENT DES FICHIERS COMMUNS
    # ==================================================================================
    @classmethod
    def _charger_IRIS_sessions(cls) -> None:
        """
        Retourne l’extract IRIS, en le chargeant si nécessaire.
        (Chargé une seule fois, partagé entre toutes les instances.)
        """
        # Vérifie existance de fe_IRIS sinon on le charge
        cls._fe_IRIS_sessions = IRIS.charger_excel_IRIS_sessions(fe_IRIS_sessions=cls._fe_IRIS_sessions, chemin_IRIS_sessions=cls._chemin_IRIS_sessions)

        # Conversion du code IRIS en string pour jointure future
        cls._df_IRIS_sessions = cls._fe_IRIS_sessions._tableaux["Sessions"]._df
        cls._df_IRIS_sessions["Code IRIS"] = cls._df_IRIS_sessions["Code IRIS"].astype(str)


    @classmethod
    def _sauver_excel_evaluations_formation(cls, nouveau_chemin_fichier:Optional[Path] = None, fermer_fichier:Optional[bool]=True) -> None:
        """
        Sauvegarde le fichier Excel des évaluations de la formation.

        :Note: Il faut sauvegarder à la sortie du contexte (car fin du traitement de la ou des sessions d'une même formation)
        """
        print(f"\n\n{Style.BRIGHT}{Fore.RED}Écriture du fichier global des évaluations de la formation {cls._trigramme_formation}")

        # On écrit et on sauve
        cls._fe_evaluations_formation._tableaux["CSV_stagiaires"].ecrit_dataFrame_dans_tableauStructure(supprimeDonneesEtRemplace = cls._supprimeDonneesEtRemplace_evaluations_formation)
        cls._fe_evaluations_formation._tableaux["Stagiaires"].ecrit_dataFrame_dans_tableauStructure(supprimeDonneesEtRemplace = cls._supprimeDonneesEtRemplace_evaluations_formation)
        #self._fe_evaluations_formation._tableaux["CSV_stagiaires"].charge_df()
        #self._fe_evaluations_formation._tableaux["Stagiaires"].charge_df()

        #self._df_evaluations_formation = instance._fe_evaluations_formation._tableaux["Stagiaires"]._df  # Alias

        cls._fe_evaluations_formation.save(nouveau_chemin_fichier)

        # Actualiser TCD
        cls._fe_evaluations_formation.actualiser_TCD()

        # On ferme si demandé
        if fermer_fichier:
            cls._fe_evaluations_formation.close()

    # ==================================================================================
    # MÉTHODES D’INSTANCE - TRAITEMENT INDIVIDUEL
    # ==================================================================================
    def traiter(self, ouvrirDossier:bool=False) -> None:
        """
        Traitement principal du CSV EvalStat stagiaire individuel.

        Inclut la mise à jour du fichier Excel de la formation courante.
        """
        if not self._chemin_csv_evaluations_stagiaires:
            raise ValueError("Aucun chemin CSV fourni pour le traitement.")

        #timer.debut(f"Traitement du fichier CSV : {self._chemin_csv_evaluations_stagiaires.name}")
        print("\n")
        timer.debut(f"{Style.BRIGHT}{Fore.YELLOW}Gestion du CSV {self._chemin_csv_evaluations_stagiaires.name}") 

        # On vérifie que chemin_csv_session n'est pas déjà dans le fichier évaluations des formations pour savoir si on l'exclue du traitement
        if str(self._chemin_csv_evaluations_stagiaires) in self._df_evaluations_formation["Chemin fichier CSV"].drop_duplicates().tolist():  
            print(f"⚠️  Exclusion car csv déjà dans le fichier global : {self._chemin_csv_evaluations_stagiaires}")
            #pprint(self._df_evaluations_formation["Chemin fichier CSV"])
            self._codeIRIS = str(rechercheX_dataframe(
                self._df_evaluations_formation,
                "Chemin fichier CSV",
                str(self._chemin_csv_evaluations_stagiaires),
                "Code IRIS"
            ))
            self._statut_csv = "Exclu - CSV déjà dans fichier global"
            return

        # On récupère IRIS sessions, seulement si nécessaire (je le fais ici car si besoin action utilisateur ça évite de couper le traitement de la boucle)
        self._charger_IRIS_sessions()

        # Définition self._codeIRIS. Si non existant, on récupère le numéro IRIS depuis le CSV (c'est le plus sur), sinon popup pour demander
        if self._codeIRIS is None:
            self._codeIRIS = IRIS.extraire_code_IRIS_depuis_chemin(self._chemin_csv_evaluations_stagiaires)

    
        # Étape 1 — Charger le CSV
        df_csv_stagiaires = self._charger_csv_stagiaire()
        if df_csv_stagiaires is None:
            return
        if df_csv_stagiaires.empty:
            return
        
        # Étape 2 — Vérifier la cohérence avec l'Extract IRIS (i.e. s'il est bien existant dans l'extract)
        if self._codeIRIS not in self._df_IRIS_sessions["Code IRIS"].values:
            print(f"⚠️  Code IRIS {self._codeIRIS} non trouvé dans l’extract IRIS.")
            self._statut_csv = "Exclu - Code IRIS pas dans Extract IRIS sessions"
            return

        # Étape 3 — Générer l'excel des évaluations des stagiaire (_fe_evaluations_stagiaires : 2 onglets + TCD)
        df_csv_stagiaires = self._traiter_onglet_csv_stagiaires(df_csv_stagiaires, remplace_df=True)  # Traitement onglet CSV_stagiaires (import "direct" du CSV avec quelques traitements mineurs)
        df_stagiaires = self._traiter_onglet_stagiaires(df_csv_stagiaires, remplace_df=True)  # Traitement seconde partie du dataframe du CSV
        self._fe_evaluations_stagiaires.actualiser_TCD()  # Mise à jour TCD
        self._statut_csv = "Traité"

        #print("\nTypes de données de df_stagiaires dans méthode traiter :")
        #print(df_stagiaires.dtypes)

        # Étape 4 — Mettre à jour le DataFrame de formation partagé
        self._mettre_a_jour_evaluations_formation()

        # Ouverture du dossier à la fin si demandé
        if ouvrirDossier:
            ouvrir_dossier(self._chemin_excel_evaluations_stagiaires.parent)

        timer.fin()

    # ==================================================================================
    # MÉTHODES INTERNES
    # ==================================================================================
    def _charger_csv_stagiaire(self) -> Optional[pd.DataFrame] :
        """
        Permet de stocker un CSV dans un dataframe en employant le bon encodage
        si le return est None c'est qu'il y a eu un problème ou que le dataframe est vide (csv présent avec en-têtes mais sans ligne)
        """
        codage_csv = trouve_encodage_csv(self._chemin_csv_evaluations_stagiaires)  # On récupère l'encodage et on importe le CSV dans un DataFrame
        try:
            df_csv_stagiaires = pd.read_csv(self._chemin_csv_evaluations_stagiaires, sep=';', encoding=codage_csv)  # Ouverture du CSV et mise dans un DataFrame
        except Exception as e:
            print(f"❌ Erreur lors de la lecture du CSV {self._chemin_csv_evaluations_stagiaires} : {e}")
            self._statut_csv = "Exclu - Problème lecture CSV"
            return None

        if df_csv_stagiaires.empty:
            print("⚠️  CSV vide → fichier ignoré.")
            self._statut_csv = "Exclu - CSV vide / Aucun retour"
            return None
        
        return df_csv_stagiaires

    def _traiter_onglet_csv_stagiaires(self, df_csv_stagiaires:pd.DataFrame, remplace_df:bool) -> pd.DataFrame:
        """
        Importe un csv stagiaires (eval & go) dans l'onglet "CSV_stagiaires" du modèle Excel EvalStat.
        On fait aussi quelques traitements (ajout nom du CSV, enlever date de fin, gestion de 2 types de csv stagiaires...)

        En fin de traitement :
            - on colle les infos dans le modèle excel stagiaire
            - on sauvegarde cet excel au bon endroit dans la GED
            - on ferme
            - on reload
        En effet, à cause des espaces à la con qui trainent dans les noms des colonnes des CSV, je dois reload le dataframe depuis l'excel que je viens de créer car les colonnes du modèle sont bien nommées        

        return : dataframe équivalent à df_csv_stagiaires

        """
        # Prise en compte qu'on a plusieurs formats de CSV : on doit traiter des colonnes en + ou - en conséquences
        if "Date de fin" in df_csv_stagiaires.columns:
            # Cas 1 (nouveau format de csv) : supprimer "Date de fin" → Test : r"P:\FORMATIONS_C\54C\P07-bilan-sessions-et-bilan-formation\rapports-sessions-evaluations\2023-06-S14317 UEM\S-14317-FC23-54C-VTE-LRA-Stagiaires.csv"
            df_csv_stagiaires = df_csv_stagiaires.drop(columns=["Date de fin"])
        else:
            # Cas 2 (ancien format de csv) : supprimer la 2e et 3e colonne (indices 1 et 2) → Test : r"P:\FORMATIONS_C\22B\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\S-17606 - 22B - 06-2025\S-17606-FC25-22B-VTE-CAR-Stagiaires.csv"
            df_csv_stagiaires = df_csv_stagiaires.drop(df_csv_stagiaires.columns[[1, 2]], axis=1)
        
        # On rajoute le chemin du CSV en première colonne
        df_csv_stagiaires.insert(0, "Chemin fichier CSV", str(self._chemin_csv_evaluations_stagiaires))

        # Met à jour ou crée la colonne "Code session" avec self._codeIRIS
        df_csv_stagiaires["Code session"] = self._codeIRIS

        # Mise au format jj/mm/aaaa de la colonne "Date" (si elle existe)
        if "Date" in df_csv_stagiaires.columns:
            try:
                df_csv_stagiaires["Date"] = pd.to_datetime(df_csv_stagiaires["Date"], dayfirst=True, errors="coerce").dt.strftime("%d/%m/%Y")  # dayfirst=True indique que le premier nombre correspond au jour (format jj/mm/aaaa)
            except Exception as e:
                print(f"Erreur de conversion de la colonne Date : {e}")


        # Save / reload Excel
        # A cause des espaces à la con qui trainent dans les noms des colonnes des CSV, je vais reload le dataframe depuis l'excel que je viens de créer car les colonnes du modèle sont bien nommées
        # Ainsi on sauve ici plutôt qu'à la fin et on reload le DataFrame
        
        # On colle le dataframe dans le modèle et on sauve
        self._fe_evaluations_stagiaires = FichierExcel.depuis_modele(chemin_modele=self._CHEMIN_MODELE_EXCEL_EVALUATIONS_STAGIAIRES, chemin_fichier_sauv=self._chemin_excel_evaluations_stagiaires)
        self._sauver_excel_stagiaire(nom_tableau="CSV_stagiaires",
                                    df_stagiaires=df_csv_stagiaires,
                                    nouveau_chemin_fichier=self._chemin_excel_evaluations_stagiaires,
                                    remplace_df=remplace_df,
                                    fermer_fichier=True)
        
        # On recharge le modèle
        self._fe_evaluations_stagiaires = FichierExcel.depuis_fichier(self._chemin_excel_evaluations_stagiaires)
        #self._fe_evaluations_stagiaires._tableaux["CSV_stagiaires"].charge_df()
        df_csv_stagiaires = self._fe_evaluations_stagiaires._tableaux["CSV_stagiaires"]._df  # Alias

        return df_csv_stagiaires

    def _traiter_onglet_stagiaires(self, df_csv_stagiaires:pd.DataFrame, remplace_df:bool) -> pd.DataFrame:
        """
        Traite le dataframe initial du CSV (df_csv_stagiaires) de sorte qu'on puisse l'exploiter par un TCD (regroupement par critères).
        Il est collé / sauvegardé dans l'excel des évaluations des stagiaires (qui est déjà créé auparavant pour df_csv_stagiaires)
        À la fin du traitement on ne ferme pas le fichier Excel car on devra mettre à jour les TCD

        En fin de traitement :
            - on colle les infos dans l'excel stagiaire (déjà créé auparavant pour df_csv_stagiaires) [onglet Stagiaires]
            - on sauvegarde
            - on ne ferme pas le fichier Excel car on devra mettre à jour les TCD
        
        return : dataframe équivalent à df_stagiaires

        """
        # Étape 1 — Préparation initiale de df_stagiaires à partir de df_csv_stagiaires

        # Nouveau DataFrame à remplir (au début c'est une liste, on convertira ensuite en dataframe)
        df_long = []

        # Parcours des lignes de CSV_stagiaires
        for _, row in df_csv_stagiaires.iterrows():
            #print("Ligne en cours : ")
            #print(row)
            base = {col: row[col] for col in self._colonnes_csv_fixes}  # Création des colonnes qui seront répétées à chaque fois
            base["NOM Prénom"] = f"{str(row['Nom']).upper()} {row['Prénom']}".strip()  # Création du champ "NOM Prénom"

            # Cas 1 : colonnes avec note + commentaire associé
            for critere in self._colonnes_csv_avec_commentaires:
                if critere in row:
                    #print(f"'{critere}'")
                    if(critere == "Avez-vous d'autres besoins de formation ?"):  # Il faut changer le booléen en 0 ou 5
                        val = str(row[critere]).strip().lower()
                        row[critere] = 5 if val == "oui" else (0 if val == "non" else None)
                    commentaire_col = row.index[row.index.get_loc(critere) + 1]
                    if(pd.notna(row[critere]) or pd.notna(row.get(commentaire_col, None))) :
                        df_long.append({
                            **base,
                            "Critère": critere,
                            "Note": row[critere],
                            "Commentaires": row.get(commentaire_col, None)
                        })

            # Cas 2 : colonnes texte seul
            for critere in self._colonnes_csv_commentaires_seuls:
                if critere in row:
                    if(pd.notna(row[critere])) :
                        df_long.append({
                            **base,
                            "Critère": critere,
                            "Note": None,
                            "Commentaires": row[critere]
                        })

            # Cas 3 : Note seule + booléens convertis
            for critere in self._colonnes_csv_bool:
                if critere in row:
                    if(pd.notna(row[critere])) :
                        val = str(row[critere]).strip().lower()
                        note = 5 if val == "oui" else (0 if val == "non" else None)
                        df_long.append({
                            **base,
                            "Critère": critere,
                            "Note": note,
                            "Commentaires": None
                        })

        # Construction du DataFrame final
        df_stagiaires = pd.DataFrame(df_long)

        # Réorganise les colonnes pour placer "NOM Prénom" juste après "Nom"
        colonnes = list(df_stagiaires.columns)
        if "NOM Prénom" in colonnes and "Nom" in colonnes:
            colonnes.remove("NOM Prénom")
            index_nom = colonnes.index("Nom")
            colonnes.insert(index_nom + 1, "NOM Prénom")
            df_stagiaires = df_stagiaires[colonnes]

        # Supprime les colonnes "Prénom" et "Nom" devenues inutiles
        df_stagiaires.drop(columns=["Prénom", "Nom"], inplace=True)



        # Étape 2 — On fait la jointure entre df_stagiaires et les données qui proviennent de l'extract IRIS Sessions

        # Pour faire le merge, il faut que les colonnes soient de même type (là "Code session" est de type int64 et "Code IRIS" est de type object (souvent des chaînes de caractères)).
        # Comme je ne peux être sûr que tous les "Code IRIS" issu des CSV soient bien convertibles en int (c’est-à-dire pas de chaînes vides, NaN, ou autres caractères non numériques), alors je passe par des strings
        df_stagiaires["Code session"] = df_stagiaires["Code session"].astype(str)
        
        # On récupère uniquement les colonnes souhaitées dans une vue pour faciliter le codage (c'est un alias)
        df_sessions_filtre = self._fe_IRIS_sessions._tableaux["Sessions"]._df[self._colonnes_sessions]
        
        # On fait la jointure entre df_stagiaires et df_sessions_filtre
        df_stagiaires = df_stagiaires.merge(
            df_sessions_filtre,
            left_on="Code session",
            right_on="Code IRIS",
            how="left"
        )

        # On réorganise les colonnes : d'abord celles de df_sessions puis celles de df_stagiaires
        colonnes_resultat = (
            df_sessions_filtre.columns.tolist() +  # colonnes de _df_sessions
            [col for col in df_stagiaires.columns if col not in df_sessions_filtre.columns]  # le reste (i.e. celles de df_stagiaires)
        )
        df_stagiaires = df_stagiaires[colonnes_resultat]
        #print("\nTypes de données de df_stagiaires dans _traiter_onglet_stagiaires :")
        #print(df_stagiaires.dtypes)
        #print(df_stagiaires)
        
        # On vire "Code session" qui est redondante avec "Code IRIS"
        df_stagiaires.drop(columns=["Code session"], inplace=True)
        
        # On renomme les colonnes
        #df_sessions_filtre.rename(columns={"Date début ses.": "Date"}, inplace=True)
        #df_sessions_filtre.rename(columns={"Année début ses.": "Année"}, inplace=True)



        # Étape 3 — Sauvegarder le fichier Excel individuel
        self._sauver_excel_stagiaire(nom_tableau="Stagiaires",
                                    df_stagiaires=df_stagiaires,
                                    remplace_df=remplace_df,
                                    fermer_fichier=False)
        
        return df_stagiaires

    def _sauver_excel_stagiaire(self, nom_tableau:str, df_stagiaires:pd.DataFrame, nouveau_chemin_fichier:Optional[Path] = None, remplace_df:bool=False, fermer_fichier:Optional[bool]=False) -> None:
        """
        Sauvegarde le fichier Excel individuel correspondant à ce CSV.
        """

        # On remplace le DataFrame existant par le nouveau
        if remplace_df:
            self._fe_evaluations_stagiaires._tableaux[nom_tableau].remplace_df(df_stagiaires)
        
        # Ancienne version chat
        #chemin_excel = self._chemin_csv_evaluations_stagiaires.with_suffix(".xlsx")
        #print(f"📄 Création du fichier Excel : {chemin_excel}")
        #FichierExcel.creer_depuis_dataframe(df_csv, chemin_excel)

        # On écrit et on sauve
        self._fe_evaluations_stagiaires._tableaux[nom_tableau].ecrit_dataFrame_dans_tableauStructure(df_stagiaires, supprimeDonneesEtRemplace=True)
        self._fe_evaluations_stagiaires.save(nouveau_chemin_fichier)

        # On ferme si demandé
        if fermer_fichier:
            self._fe_evaluations_stagiaires.close()




    def _mettre_a_jour_evaluations_formation(self) -> None:
        """
        Met à jour le DataFrame partagé des évaluations de la formation courante
        avec les données du CSV actuel.
        """

        if EvalStat._df_evaluations_formation is None:
            print("⚠️  Aucun DataFrame de formation ouvert, impossible de mettre à jour.")
            return
        
        #print("\n")
        #timer.debut(f"{Style.BRIGHT}{Fore.RED}Mise à jour de {self._fe_evaluations_formation.chemin_fichier.name}")

        # Nota : le fichier est déjà ouvert / créé à l'initialisation du contexte.

        df_formation_csv = self._fe_evaluations_formation._tableaux["CSV_stagiaires"]._df  # Alias
        df_formation_stagiaires = self._fe_evaluations_formation._tableaux["Stagiaires"]._df  # Alias

        # TODO : ici je mets à jour le df de l'excel évaluation formation. Il sera écrit physiquement à la sortie du contexte formation.
        # TODO : df nouveau df comprend l'ancien (i.e. évaluation formation existant) + le nouveau que l'on traite.
        # TODO : pour l'instant je réécrit tout ce df mais pour être optimal on ne pourrait écrire que le nouveau

        # Si df_formation_csv est vide, il faut l'initialiser avec le premier df sinon on concatène
        if (df_formation_csv is None) or (df_formation_csv.empty):
            df_formation_csv = self._fe_evaluations_stagiaires._tableaux["CSV_stagiaires"]._df.copy()
            df_formation_stagiaires = self._fe_evaluations_stagiaires._tableaux["Stagiaires"]._df.copy()
            
        else:
            df_csv_stagiaires = self._fe_evaluations_stagiaires._tableaux["CSV_stagiaires"]._df # Alias
            if (not df_csv_stagiaires.empty) and (df_csv_stagiaires is not None): # Evite un future wanring de concaténer avec un df vide
                df_csv_stagiaires = adapter_colonnes_dataframe_selon_modele(df_modele=df_formation_csv, df_a_modifier=df_csv_stagiaires)
                df_formation_csv = pd.concat([df_formation_csv, df_csv_stagiaires], ignore_index=True)
            
            df_stagiaires = self._fe_evaluations_stagiaires._tableaux["Stagiaires"]._df # Alias
            if (not df_stagiaires.empty) and (df_stagiaires is not None): # Evite un future wanring de concaténer avec un df vide
                
                df_stagiaires = adapter_colonnes_dataframe_selon_modele(df_modele=df_formation_stagiaires, df_a_modifier=df_stagiaires)
                
                # Tester colonnes
                #print("Colonnes de df_formation_stagiaires :", df_formation_stagiaires.columns)
                #print("Colonnes de df :", df.columns)

                #print("\nTypes de données de df_formation_stagiaires :")
                #print(df_formation_stagiaires.dtypes)

                #print("\nTypes de données de df_stagiaires :")
                #print(df_stagiaires.dtypes)

                # Tester index
                #print("Index de df_formation_stagiaires :", df_formation_stagiaires.index)
                #print("Index de df :", df_stagiaires.index)


                df_formation_stagiaires = pd.concat([df_formation_stagiaires, df_stagiaires], ignore_index=True)

        self._fe_evaluations_formation._tableaux["CSV_stagiaires"]._df = df_formation_csv
        self._fe_evaluations_formation._tableaux["Stagiaires"]._df = df_formation_stagiaires

        #timer.fin()


    def _NON_EMPLOYE_extraire_code_IRIS_depuis_csv(self, df_csv:pd.DataFrame) -> str:
        """
        JAMAIS TESTE NON-EMPLOYE
        Extrait le code IRIS depuis le contenu ou le nom du CSV.
        """
        # Exemple : suppose que le code IRIS est contenu dans une colonne "Code_IRIS"
        if "Code_IRIS" in df_csv.columns:
            return str(df_csv["Code_IRIS"].iloc[0])
        # Sinon, on tente de le déduire depuis le nom du fichier
        return Path(self._chemin_csv_evaluations_stagiaires).stem.split("_")[0]

    # ==================================================================================
    # GETTERS / SETTERS
    # ==================================================================================
    @property
    def _chemin_excel_evaluations_stagiaires(self) -> Path:
        """
        Retourne le chemin Excel correspondant au CSV courant.
        Pour le nom de l'Excel output : on reprend le nom du csv et on remplace par xlsx.
        """
        return self._chemin_csv_evaluations_stagiaires.with_suffix(".xlsx")

    @property
    def _chemin_excel_evaluations_formation(self) -> Path:
        """
        Retourne le chemin du fichier Excel des évaluations des formations depuis la valeur en config.
        """
        return config.format_path(config.CHEMIN_EXCEL_EVALUATIONS_FORMATION, trigramme_formation=self._trigramme_formation)

    # ==================================================================================
    # POPUP
    # ==================================================================================
    @classmethod
    def _filedialog_csv(cls, code_IRIS: Optional[int] = None, trigramme_formation: Optional[str] = None) -> str | None:
        """
        Ouvre un filedialog pour demander à l'utilisateur de sélectionner un CSV.
        On pointe au mieux sur le répertoire des CSV de cette formation pour la boîte de dialogue.
        """

        # Si on n'a pas de trigramme de foramtion, alors on tente de le récupérer depuis _fe_evaluations_formation s'il existe et de trouver un chemin optimisé
        if (trigramme_formation is None) and (cls._fe_evaluations_formation is not None):
            if cls._fe_evaluations_formation.chemin_fichier is not None:
                trigramme_formation = recupere_trig_formation_depuis_chemin(cls._fe_evaluations_formation.chemin_fichier)

        # Si on a un trigramme de formation, alors on est en mesure de trouver un chemin optimis
        if trigramme_formation:
            chemin_repertoire_csv = optimiseCheminRepertoire(
                config.format_path(config.REPERTOIRE_CSV_EVALUATIONS, trigramme_formation=trigramme_formation)
                )
        else:
            chemin_repertoire_csv = optimiseCheminRepertoire(config.REPERTOIRE_FORMATION.parent)  #Path.cwd()  
            


        if code_IRIS is None :
            fin_titre = "désiré"
        else :
            fin_titre = f"de la session {code_IRIS}"


        return choisir_fichier(titre=f"Sélectionner le fichier EvalStat stagiaire {fin_titre}",
                        types_fichiers=[("Fichiers CSV", "*.csv")],
                        dossier_initial=chemin_repertoire_csv,
                        obligatoire=False,
                        texte_bouton_choisir="Choisir CSV à nouveau",
                        texte_bouton_aucun="Pas de CSV pour cette session"
                        )


