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

        # J'ai besoin de plusieurs variables qui seront transverses à plusieurs méthodes lors du traitement
        # === Pour traitement ===
        self._chemin_csv: Optional[Path] = None
        self._chemin_excel: Optional[Path] = None
        self._df_csv_stagiaires: Optional[pd.DataFrame] = None
        self._df_stagiaires: Optional[pd.DataFrame] = None
        # === Résultat du traitement ===
        self._statut_csv: Optional[str] = None  # ex: "Traité", "Exclu - CSV déjà dans fichier global", "Exclu - Code IRIS pas dans Extract IRIS sessions", "Exclu - Problème lecture CSV", "Exclu - CSV vide / Aucun retour"


    @classmethod
    def traitement_depuis_chemin_csv(cls, session:Session_protocol, chemin_csv:Path|str, chemin_IRIS_sessions:Optional[Path]=None, ouvrirDossier:bool=False) -> EvalStat_session:
        """
        Crée et traite une instance d'EvalStat à partir d'un fichier CSV.

        Peut être utilisée directement dans un contexte `ContexteFormation`.

        Exemple :
            with ContexteFormation("ABC"):
                EvalStat.depuis_chemin_csv_evaluations_stagiaires(Path("eval_R04110.csv"))
        """

        # On créée l'instance
        instance = cls(session)

        # Début ancienne méthode depuis_chemin_csv_evaluations_stagiaires
        

        
        # Soit on a déjà un chemin, soit on va pointer le csv manuellement
        if chemin_csv is None:
            instance._chemin_csv = cls._filedialog_csv(trigramme_formation=instance.trigramme_formation)
            if instance._chemin_csv is None:
                vlog.log_erreur("Pas de fichier CSV", continuer=True)
                return 

        # Si le chemin est avec un raccourci réseau alors on récupère le chemin en entier + on convertit en Path 
        instance._chemin_csv = chemin_vers_unc(Path(chemin_csv)) 
        instance._chemin_excel = instance._chemin_csv.with_suffix(".xlsx")

        # On récupère le trigramme de la formation depuis le chemin du CSV
        # → PAs besoin : trigramme dispo depuis session → Formation via protocol
        #if trigramme_formation is None:
        #    trigramme_formation = recupere_trig_formation_depuis_chemin(chemin_csv)

        # Si l'utilisateur force un chemin pour chemin_IRIS_sessions, alors on renseigne la valeur sinon → IRIS_traite sait gérer ça déjà










        # Début de traiter
        if not instance._chemin_csv:
            log_erreur("Aucun chemin CSV fourni pour le traitement.")

        #timer.debut(f"Traitement du fichier CSV : {self._chemin_csv_evaluations_stagiaires.name}")
        print("\n")
        timer.debut(f"{Style.BRIGHT}{Fore.YELLOW}Gestion du CSV {instance._chemin_csv.name}") 



        # Etape interactions avec eval formation (ouverture + vérif si csv déjà existant dedans)
        # On vérifie que l'eval formation est ouvert en mémoire sinon on ouvre ou crée
        # TODO : pas sûr, à voir. Si oui, revoir emplacement ?
        if instance.eval_formation.fe is None:
            # Le fichier Excel de la formation n'est pas ouvert, c'est donc le premier appel → On l'ouvre
            instance.eval_formation._ouvrir_ou_creer_eval_formation()
            #log_erreur("Le fichier Excel de la formation n'est pas ouvert.")

        # On vérifie que chemin_csv_session n'est pas déjà dans le fichier évaluations des formations pour savoir si on l'exclue du traitement
        # instance.eval_formation.df_stagiaires
        if str(instance._chemin_csv) in instance.eval_formation.df_stagiaires["Chemin fichier CSV"].drop_duplicates().tolist():  
            print(f"⚠️  Exclusion car csv déjà dans le fichier global : {instance._chemin_csv}")
            #pprint(self._df_evaluations_formation["Chemin fichier CSV"])

            # _code_IRIS est déjà connu à travers la session : instance.code_IRIS
            # Todo : mettre cette partie ailleurs dans une fonction ailleurs au cas où pour plus tard
            """
            self._codeIRIS = str(rechercheX_dataframe(
                instance.eval_formation.df_stagiaires,
                "Chemin fichier CSV",
                str(chemin_csv),
                "Code IRIS"
            ))
            """

            instance._statut_csv = "Exclu - CSV déjà dans fichier global"
            return






        # On récupère IRIS sessions, seulement si nécessaire (je le fais ici car si besoin action utilisateur ça évite de couper le traitement de la boucle)
        # TODO : quelle boucle ? Mettre ailleurs ?
        #_ = get_iris(typeExport="Sessions", chemin=chemin_IRIS_sessions)
        #self._charger_IRIS_sessions()

        # Définition self._codeIRIS. Si non existant, on récupère le numéro IRIS depuis le CSV (c'est le plus sur), sinon popup pour demander
        # _code_IRIS est déjà connu à travers la session : instance.code_IRIS
        #if self._codeIRIS is None:
        #    self._codeIRIS = IRIS.extraire_code_IRIS_depuis_chemin(chemin_csv)

    



        # Étape 1 — Charger le CSV dans df_csv_stagiaires
        #df_csv_stagiaires = instance._charger_csv_stagiaire()
        instance._charger_csv_stagiaire()
        if instance._df_csv_stagiaires is None:
            # il y a eu un problème dans _charger_csv_stagiaire → on saute la fin du traitement
            return
        if instance._df_csv_stagiaires.empty:
            # le dataframe est vide (csv présent avec en-têtes mais sans ligne) → on saute la fin du traitement
            return
        




        # Étape 2 — Vérifier la cohérence avec l'Extract IRIS (i.e. si le code_IRIS est bien existant dans l'extract)
        df_IRIS_sessions = get_iris(typeExport="Sessions", chemin=chemin_IRIS_sessions).df  # On charge IRIS
        if instance.code_IRIS not in df_IRIS_sessions["Code IRIS"].values:
            print(f"⚠️  Code IRIS {instance.code_IRIS} non trouvé dans l’extract IRIS.")
            instance._statut_csv = "Exclu - Code IRIS pas dans Extract IRIS sessions"
            return




        # Étape 3 — Générer l'excel des évaluations des stagiaire (_fe_evaluations_stagiaires : 2 onglets + TCD)
        # 3.1 - On crée les dataframes
        instance._traiter_df_csv_stagiaires()  # Traitement onglet CSV_stagiaires (import "direct" du CSV avec quelques traitements mineurs)
        # TODO : Attention, avant j'avais le sauve / reload à l'étape d'avant or juste après j'emploie df_csv_stagiaires → Bug à venir ?
        instance._traiter_df_stagiaires()  # Traitement seconde partie du dataframe du CSV
        
        # 3.2 - On fait la manip save / reload l'Excel à partir du modèle à cause des espaces des noms des colonnes
        # TODO : voir si on ne peut pas faire une manip pour récupérer les noms des colonnes du modèles et les affecter aux colonnes du df.
        # On crée le fichier Excel depuis le modèle
        instance._fe = FichierExcel.depuis_modele(chemin_modele=config.CHEMIN_MODELE_EXCEL_EVALUATIONS_STAGIAIRES, chemin_fichier_sauv=instance._chemin_excel)

        # On ecrit_dataFrame_dans_tableauStructure pour les 2 onglets
        # self._fe_evaluations_stagiaires._tableaux[nom_tableau].ecrit_dataFrame_dans_tableauStructure(df_stagiaires, supprimeDonneesEtRemplace=True)
        instance._fe.get_tableau("CSV_stagiaires").ecrit_dataFrame_dans_tableauStructure(instance._df_csv_stagiaires, supprimeDonneesEtRemplace=True)
        instance._fe.get_tableau("Stagiaires").ecrit_dataFrame_dans_tableauStructure(instance._df_stagiaires, supprimeDonneesEtRemplace=True)

        # On sauve et on ferme
        instance._fe.save(instance._chemin_excel)
        instance._fe.close()


        # On recharge l'Excel
        instance._fe = FichierExcel.depuis_fichier(instance._chemin_excel)

        # On réaffecte les 2 df (ils auront les bonnes en-têtes de colonnes)
        instance._df_csv_stagiaires = instance._fe.get_df_tableau("CSV_stagiaires")
        instance._df_stagiaires = instance._fe.get_df_tableau("Stagiaires")

        # 3.3 - On actualise les TCD et on sauve
        #instance._fe_evaluations_stagiaires.actualiser_TCD()  # Mise à jour TCD
        instance.fe.actualiser_TCD()
        instance._statut_csv = "Traité"

        #print("\nTypes de données de df_stagiaires dans méthode traiter :")
        #print(df_stagiaires.dtypes)

        # Étape 4 — Mettre à jour le DataFrame de formation partagé
        #TODO : instance._mettre_a_jour_evaluations_formation()

        # Ouverture du dossier à la fin si demandé
        if ouvrirDossier:
            #ouvrir_dossier(self._chemin_excel_evaluations_stagiaires.parent)
            ouvrir_dossier(instance.chemin_fe.parent)

        timer.fin()










        
        return instance

        




    # ==================================================================================
    # MÉTHODES D’INSTANCE - TRAITEMENT INDIVIDUEL
    # ==================================================================================
    def _charger_csv_stagiaire(self) -> Optional[pd.DataFrame] :
        """
        Permet de stocker un CSV dans un dataframe en employant le bon encodage
        si le return est None c'est qu'il y a eu un problème ou que le dataframe est vide (csv présent avec en-têtes mais sans ligne)

        chemin_csv est le chemin du CSV à aller récupérer (à ce stade il est connu)
        """
        codage_csv = trouve_encodage_csv(self._chemin_csv)  # On récupère l'encodage et on importe le CSV dans un DataFrame
        try:
            df_csv_stagiaires = pd.read_csv(self._chemin_csv, sep=';', encoding=codage_csv)  # Ouverture du CSV et mise dans un DataFrame
        except Exception as e:
            print(f"❌ Erreur lors de la lecture du CSV {self._chemin_csv} : {e}")
            self._statut_csv = "Exclu - Problème lecture CSV"
            return None

        if df_csv_stagiaires.empty:
            print("⚠️  CSV vide → fichier ignoré.")
            self._statut_csv = "Exclu - CSV vide / Aucun retour"
            return None
        
        # On nettoie les espaces en début et fin des noms d'en-tête
        df_csv_stagiaires.columns = [col.strip() for col in df_csv_stagiaires.columns]

        self._df_csv_stagiaires = df_csv_stagiaires

        return df_csv_stagiaires

    def _traiter_df_csv_stagiaires(self) -> pd.DataFrame:
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
        if "Date de fin" in self._df_csv_stagiaires.columns:
            # Cas 1 (nouveau format de csv) : supprimer "Date de fin" → Test : r"P:\FORMATIONS_C\54C\P07-bilan-sessions-et-bilan-formation\rapports-sessions-evaluations\2023-06-S14317 UEM\S-14317-FC23-54C-VTE-LRA-Stagiaires.csv"
            self._df_csv_stagiaires = self._df_csv_stagiaires.drop(columns=["Date de fin"])
        else:
            # Cas 2 (ancien format de csv) : supprimer la 2e et 3e colonne (indices 1 et 2) → Test : r"P:\FORMATIONS_C\22B\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\S-17606 - 22B - 06-2025\S-17606-FC25-22B-VTE-CAR-Stagiaires.csv"
            self._df_csv_stagiaires = self._df_csv_stagiaires.drop(self._df_csv_stagiaires.columns[[1, 2]], axis=1)
        
        # On rajoute le chemin du CSV en première colonne
        self._df_csv_stagiaires.insert(0, "Chemin fichier CSV", str(self._chemin_csv))

        # Met à jour ou crée la colonne "Code session" avec self._codeIRIS
        self._df_csv_stagiaires["Code session"] = self.code_IRIS

        # Mise au format jj/mm/aaaa de la colonne "Date" (si elle existe)
        if "Date" in self._df_csv_stagiaires.columns:
            try:
                self._df_csv_stagiaires["Date"] = pd.to_datetime(self._df_csv_stagiaires["Date"], dayfirst=True, errors="coerce").dt.strftime("%d/%m/%Y")  # dayfirst=True indique que le premier nombre correspond au jour (format jj/mm/aaaa)
            except Exception as e:
                print(f"Erreur de conversion de la colonne Date : {e}")


        # Save / reload Excel
        # A cause des espaces à la con qui trainent dans les noms des colonnes des CSV, je vais reload le dataframe depuis l'excel que je viens de créer car les colonnes du modèle sont bien nommées
        # Ainsi on sauve ici plutôt qu'à la fin et on reload le DataFrame
        
        # On colle le dataframe dans le modèle et on sauve
        """
        self._fe_evaluations_stagiaires = FichierExcel.depuis_modele(chemin_modele=config.CHEMIN_MODELE_EXCEL_EVALUATIONS_STAGIAIRES, chemin_fichier_sauv=self._chemin_excel)
        self._sauver_excel_stagiaire(nom_tableau="CSV_stagiaires",
                                    df_stagiaires=df_csv_stagiaires,
                                    nouveau_chemin_fichier=self._chemin_excel,
                                    remplace_df=remplace_df,
                                    fermer_fichier=True)
        
        # On recharge le modèle
        self._fe_evaluations_stagiaires = FichierExcel.depuis_fichier(self._chemin_excel)
        #self._fe_evaluations_stagiaires._tableaux["CSV_stagiaires"].charge_df()
        df_csv_stagiaires = self._fe_evaluations_stagiaires._tableaux["CSV_stagiaires"]._df  # Alias
        """

        return self._df_csv_stagiaires

    def _traiter_df_stagiaires(self) -> pd.DataFrame:
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
        #print(self._df_csv_stagiaires.columns.tolist())
        for _, row in self._df_csv_stagiaires.iterrows():
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
        self._df_stagiaires = pd.DataFrame(df_long)

        # Réorganise les colonnes pour placer "NOM Prénom" juste après "Nom"
        colonnes = list(self._df_stagiaires.columns)
        if "NOM Prénom" in colonnes and "Nom" in colonnes:
            colonnes.remove("NOM Prénom")
            index_nom = colonnes.index("Nom")
            colonnes.insert(index_nom + 1, "NOM Prénom")
            self._df_stagiaires = self._df_stagiaires[colonnes]

        # Supprime les colonnes "Prénom" et "Nom" devenues inutiles
        self._df_stagiaires.drop(columns=["Prénom", "Nom"], inplace=True)



        # Étape 2 — On fait la jointure entre df_stagiaires et les données qui proviennent de l'extract IRIS Sessions

        # Pour faire le merge, il faut que les colonnes soient de même type (là "Code session" est de type int64 et "Code IRIS" est de type object (souvent des chaînes de caractères)).
        # Comme je ne peux être sûr que tous les "Code IRIS" issu des CSV soient bien convertibles en int (c’est-à-dire pas de chaînes vides, NaN, ou autres caractères non numériques), alors je passe par des strings
        # TODO à remettre si jamais ça ne marche pas en int self._df_stagiaires["Code session"] = self._df_stagiaires["Code session"].astype(str)
        
        # On récupère uniquement les colonnes souhaitées dans une vue pour faciliter le codage (c'est un alias)
        #self._df_sessions_filtre = self._fe_IRIS_sessions._tableaux["Sessions"]._df[self._colonnes_sessions]
        self._df_sessions_filtre = get_iris("Sessions").df[self._colonnes_sessions]
        
        # On fait la jointure entre df_stagiaires et df_sessions_filtre
        self._df_stagiaires = self._df_stagiaires.merge(
            self._df_sessions_filtre,
            left_on="Code session",
            right_on="Code IRIS",
            how="left"
        )

        # On réorganise les colonnes : d'abord celles de df_sessions puis celles de df_stagiaires
        colonnes_resultat = (
            self._df_sessions_filtre.columns.tolist() +  # colonnes de _df_sessions
            [col for col in self._df_stagiaires.columns if col not in self._df_sessions_filtre.columns]  # le reste (i.e. celles de df_stagiaires)
        )
        self._df_stagiaires = self._df_stagiaires[colonnes_resultat]
        #print("\nTypes de données de df_stagiaires dans _traiter_onglet_stagiaires :")
        #print(df_stagiaires.dtypes)
        #print(df_stagiaires)
        
        # On vire "Code session" qui est redondante avec "Code IRIS"
        self._df_stagiaires.drop(columns=["Code session"], inplace=True)
        
        # On renomme les colonnes
        #df_sessions_filtre.rename(columns={"Date début ses.": "Date"}, inplace=True)
        #df_sessions_filtre.rename(columns={"Année début ses.": "Année"}, inplace=True)



        # Étape 3 — Sauvegarder le fichier Excel individuel
        """
        self._sauver_excel_stagiaire(nom_tableau="Stagiaires",
                                    df_stagiaires=df_stagiaires,
                                    remplace_df=remplace_df,
                                    fermer_fichier=False)
        """
        return self._df_stagiaires



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
        """
        Initialisation d'un EvalStat formation a minima
        """
        # On initialise la classe mère
        super().__init__()

        #self._fe → classe mère
        self._formation = formation  # Protocol pour éviter les références circulaires

        self._df_initial_hash:Optional[str] = None  # hash du df initial pour savoir s'il a été modifié, auquel cas on sauvegardera à la fin
        self._supprimeEtRemplace_donneesEval:Optional[bool] = None


    @classmethod
    def avec_ouverture(cls, formation:Formation_protocol) -> EvalStat_formation:
        """
        Initialisation d'un EvalStat formation avec ouverture ou création du fichier Excel
        """
        instance = cls(formation)

        instance._ouvrir_ou_creer_eval_formation()

        return instance



    def _ouvrir_ou_creer_eval_formation(self):  # -> Tuple[FichierExcel, bool]:
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
        timer.debut(f"Ouverture ou création du fichier Excel de la formation {self.trigramme_formation}")



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


        # On retype en str la colonne Code IRIS
        # TODO : vraiment ???
        #self._df_evaluations_formation = self._fe_evaluations_formation._tableaux["Stagiaires"].df  # Alias
        #self._df_evaluations_formation["Code IRIS"] = self._df_evaluations_formation["Code IRIS"].astype(str)
        self.df_stagiaires["Code IRIS"] = self.df_stagiaires["Code IRIS"].astype(str)

        # On fait un hash du df pour savoir si, en fin de traitement il aura été modifié, auquel cas on le sauvegardera
        self._df_initial_hash = hash_df(self.df_stagiaires)
        timer.fin()

        #return self._fe_evaluations_formation, self._df_evaluations_formation, supprimeDonneesEtRemplace
        #return fe, supprimeDonneesEtRemplace


    def _sauver_eval_formation_siModif(self):
        """
        Sauvegarde et fermeture du fichier Excel de la formation à la sortie du contexte.

        Ce bloc est exécuté même en cas d'erreur dans le traitement des sessions. → Non ça c'était quand c'était dans un contexte
        """
        # TODO : dans _mettre_a_jour_evaluations_formation() je mets à jour le df de l'excel évaluation formation. Il sera écrit physiquement à la sortie du contexte formation.
        # TODO : df nouveau df comprend l'ancien (i.e. évaluation formation existant) + le nouveau que l'on traite.
        # TODO : pour l'instant je réécrit tout ce df mais pour être optimal on ne pourrait écrire que le nouveau


        if self._fe is not None:

            df_final_hash = hash_df(self.df_stagiaires)

            if df_final_hash != self._df_initial_hash:
                #TODO : pour l'instant je force à tout réécrire et pas seulement faire les mises à jour
                # Forçage réécriture en entier du df
                self._supprimeDonneesEtRemplace_evaluations_formation = True

                # Le DataFrame a changé → on sauvegarde
                self._sauver_excel_evaluations_formation(fermer_fichier=False)

            else:
                vlog.ajouter_message(
                    "Aucune modification détectée dans l'évaluation formation → pas de sauvegarde",
                    self.chemin_fe,  # self._fe_evaluations_formation.chemin_fichier
                    style=["jaune"]
                )


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
