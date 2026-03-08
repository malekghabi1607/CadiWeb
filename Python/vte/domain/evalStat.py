from __future__ import annotations
from collections import defaultdict
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

# TODO : Pour l'instant c'est une classe de traitemnt. Le jour où j'ai besoin d'ouvrir un EvalStat pour le lire uniquement, prendre modèle sur IRIS avec des classes de lecture et de traitement

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

class Session_protocol(Protocol):
    """
    Protocol de Session : permet de simuler une session en évitant les références circulaires
    """
    @property
    def code_IRIS(self) -> int|None: ...
        
    @property
    def trigramme_formation(self) -> str: ...

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
    # === Colonnes du modèle de l'evalStat
    _colonnes_modele_evalStat = [
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
        """
        Crée une instance d'EvalStat à minima (self_fe = None)
        """
        self._fe:Optional[FichierExcel] = None  # Objet Excel contenant les données EvalStat stagiaire individuel


    # ==================================================================================
    # METHODES INTERNES
    # ==================================================================================
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


    # ==================================================================================
    # GETTERS / SETTERS
    # ==================================================================================
    @property
    def fe(self) -> FichierExcel|None:
        return self._fe
     
    @property
    def df_csv(self) -> pd.DataFrame|None:
        return self._fe.get_df_tableau("CSV_stagiaires")

    @df_csv.setter
    def df_csv(self, valeur:pd.DataFrame) -> None:
        self._fe.set_df_tableau("CSV_stagiaires", df=valeur)

    @property
    def df_stagiaires(self) -> pd.DataFrame|None:
        return self._fe.get_df_tableau("Stagiaires")

    @df_stagiaires.setter
    def df_stagiaires(self, valeur:pd.DataFrame) -> None:
        self._fe.set_df_tableau("Stagiaires", df=valeur)

    @property
    def chemin_fe(self) -> Path|None:
        return self._fe.chemin_fichier
    

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
    # =====================
    # === CONSTRUCTEURS ===
    # =====================
    def __init__(self, session:Session_protocol):
        """
        Crée l'instance EvalStat d'un session a minima.

        :param session: La session à laquelle est affectée l'EvalStat
        :type session: Session_protocol
        """
        # On initialise la classe mère
        super().__init__()

        self._session = session  # C'est un protocol pour éviter les références circulaires

        # === Résultat du traitement ===
        self._statut_csv: Optional[str] = None  # ex: "Traité", "Exclu - CSV déjà dans fichier global", "Exclu - Code IRIS pas dans Extract IRIS sessions", "Exclu - Problème lecture CSV", "Exclu - CSV vide / Aucun retour"


    @classmethod
    def avec_traitement(cls, session:Session_protocol, chemin_csv:Optional[Path|str]=None, chemin_IRIS_sessions:Optional[Path]=None, ecrire_eval_formation:bool=True, ouvrirDossier:bool=False) -> EvalStat_session:
        """
        Crée l'instance EvalStat d'une session et traite cet EvalStat.

        Crée l'Excel EvalStat d'une session à partir d'un CSV (s'il n'est pas donné, on ouvre un filedialog).

        L'évaluation de la formation est mise à jour avec ces nouvelles données et est sauvée en fin de traitement selon le critère ecrire_eval_formation.

        :param session: La session à laquelle est affectée l'EvalStat
        :type session: Session_protocol
        :param chemin_csv: chemin du CSV à traiter. S'il est None, on ouvre un filedialog
        :type chemin_csv: Optional[Path|str]
        :param chemin_IRIS_sessions: Chemin du fichier IRIS sessions à employer si l'utilisateur ne veut pas celui par défaut. Defaults = None = Fichier généré le plus récent dans le répertoire donné en config.
        :type chemin_IRIS_sessions: Optional[Path]
        :param ecrire_eval_formation: Pour écrire physiquement l'Excel eval formation en fin de traitement. Si False, il devra être écrit ailleurs (à l'endroit où il y a la boucle pour du multi-traitement typiquement). Défaut = True.
        :type ecrire_eval_formation: bool
        :param ouvrirDossier: Ouvre le répertoire de l'EvalStat généré. Defaut = False.
        :type ouvrirDossier: bool
        """
        instance = cls(session)
        instance.traiter_eval(
            chemin_csv=chemin_csv,
            chemin_IRIS_sessions=chemin_IRIS_sessions,
            ecrire_eval_formation=ecrire_eval_formation,
            ouvrirDossier=ouvrirDossier
        )
        return instance

        




    # ==================================================================================
    # MÉTHODES D’INSTANCE - TRAITEMENT INDIVIDUEL
    # ==================================================================================
    def _charger_csv_stagiaire(self, chemin_csv:Path) -> Optional[pd.DataFrame] :
        """
        Permet de stocker un CSV dans un dataframe en employant le bon encodage
        Si le return est None c'est qu'il y a eu un problème ou que le dataframe est vide (csv présent avec en-têtes mais sans ligne).
        Dans ce cas on écrit le self._statut_csv pour tracer la raison exclusion.
        chemin_csv est le chemin du CSV à aller récupérer (à ce stade il est connu)

        :param chemin_csv: chemin du csv à charger
        :type chemin_csv: Path
        :return: le dataframe du csv. Si None, c'est qu'il y a eu un problème ou que le dataframe est vide (csv présent avec en-têtes mais sans ligne).
        :rtype: Optional[pd.DataFrame]
        """
        codage_csv = trouve_encodage_csv(chemin_csv)  # On récupère l'encodage et on importe le CSV dans un DataFrame
        try:
            df_csv_stagiaires = pd.read_csv(chemin_csv, sep=';', encoding=codage_csv)  # Ouverture du CSV et mise dans un DataFrame
        except Exception as e:
            print(f"❌ Erreur lors de la lecture du CSV {chemin_csv} : {e}")
            self._statut_csv = "Exclu - Problème lecture CSV"
            return None

        if df_csv_stagiaires.empty:
            print("⚠️  CSV vide → fichier ignoré.")
            self._statut_csv = "Exclu - CSV vide / Aucun retour"
            return None
        
        # On traite les csv selon leur type / renomme les colonnes / ...
        # Prise en compte qu'on a plusieurs formats de CSV : on doit traiter des colonnes en + ou - en conséquences
        
        if "Date de fin" in df_csv_stagiaires.columns:
            # Cas 1 (nouveau format de csv) : supprimer "Date de fin" → Test : r"P:\FORMATIONS_C\22B\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\S-17606 - 22B - 06-2025\S-17606-FC25-22B-VTE-CAR-Stagiaires.csv"
            df_csv_stagiaires = df_csv_stagiaires.drop(columns=["Date de fin"])
            #df_csv_stagiaires = df_csv_stagiaires.rename(columns={df_csv_stagiaires.columns[0]: "Date"})
        else:
            # Cas 2 (ancien format de csv) : supprimer la 2e et 3e colonne (indices 1 et 2) → Test : r"P:\FORMATIONS_C\54C\P07-bilan-sessions-et-bilan-formation\rapports-sessions-evaluations\2023-06-S14317 UEM\S-14317-FC23-54C-VTE-LRA-Stagiaires.csv"
            df_csv_stagiaires = df_csv_stagiaires.drop(df_csv_stagiaires.columns[[1, 2]], axis=1)
            #df_csv_stagiaires = df_csv_stagiaires.rename(columns={df_csv_stagiaires.columns[1]: "Prénom"})  
            #df_csv_stagiaires = df_csv_stagiaires.rename(columns={df_csv_stagiaires.columns[2]: "Nom"})
        #print(df_csv_stagiaires)

        # On rajoute le chemin du CSV en première colonne
        df_csv_stagiaires.insert(0, "Chemin fichier CSV", str(chemin_csv))

        # On nettoie les espaces en début et fin des noms d'en-tête
        #df_csv_stagiaires.columns = [col.strip() for col in df_csv_stagiaires.columns]

        # Je renomme à la main toutes les colonnes à la main car les CSV c'est le bordel avec des espaces qui trainent et des caractères spéciaux
        mapping = dict(zip(df_csv_stagiaires.columns, self._colonnes_modele_evalStat))  # On fait un dictionnaire de mapping anciens noms/nouveaux noms
        df_csv_stagiaires = df_csv_stagiaires.rename(columns=mapping)  # On renomme les colonnes


    
        # Mise au format jj/mm/aaaa de la colonne "Date" (si elle existe)
        try:
            df_csv_stagiaires["Date"] = pd.to_datetime(df_csv_stagiaires["Date"], dayfirst=True, errors="coerce").dt.strftime("%d/%m/%Y")  # dayfirst=True indique que le premier nombre correspond au jour (format jj/mm/aaaa)
        except Exception as e:
            print(f"Erreur de conversion de la colonne Date : {e}")


        #print(df_csv_stagiaires)
        return df_csv_stagiaires

    def _traiter_df_csv(self) -> None:
        """
        On génère le DataFrame du CSV et on l'affecte à self.df_csv.
        On fait quelques traitements (ajout nom du CSV, enlever date de fin, gestion de 2 types de csv stagiaires...)
        """
        # On recrée le chemin du CSV à partir du chemin di fichier Excel
        chemin_csv = self.chemin_fe.with_suffix(".csv")
        
        #print(self.df_csv)
        
        # On colle le chemin du CSV en première colonne
        #self.df_csv.insert(0, "Chemin fichier CSV", str(chemin_csv))
        #self.df_csv["Chemin fichier CSV"] = str(chemin_csv)

        # Met à jour ou crée la colonne "Code session" avec self._codeIRIS
        self.df_csv["Code session"] = self.code_IRIS



        #print(self.df_csv)

        #return self.df_csv

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
        self.df_stagiaires = pd.DataFrame(df_long)

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



    # ==================================================================================
    # MÉTHODES EXTERNES - TRAITEMENT INDIVIDUEL
    # ==================================================================================
    def traiter_eval(self, chemin_csv:Optional[Path|str]=None, chemin_IRIS_sessions:Optional[Path]=None, ecrire_eval_formation:bool=True, ouvrirDossier:bool=False) -> None:
        """
        Crée l'Excel EvalStat d'une session à partir d'un CSV (s'il n'est pas donné, on ouvre un filedialog).

        L'évaluation de la formation est mise à jour avec ces nouvelles données et est sauvée en fin de traitement selon le critère ecrire_eval_formation.

        :param chemin_csv: chemin du CSV à traiter. S'il est None, on ouvre un filedialog
        :type chemin_csv: Optional[Path|str]
        :param chemin_IRIS_sessions: Chemin du fichier IRIS sessions à employer si l'utilisateur ne veut pas celui par défaut. Defaults = None = Fichier généré le plus récent dans le répertoire donné en config.
        :type chemin_IRIS_sessions: Optional[Path]
        :param ecrire_eval_formation: Pour écrire physiquement l'Excel eval formation en fin de traitement. Si False, il devra être écrit ailleurs (à l'endroit où il y a la boucle pour du multi-traitement typiquement). Défaut = True.
        :type ecrire_eval_formation: bool
        :param ouvrirDossier: Ouvre le répertoire de l'EvalStat généré. Defaut = False.
        :type ouvrirDossier: bool
        """
        print("\n")
        timer.debut(f"{Style.BRIGHT}{Fore.YELLOW}Gestion du CSV {chemin_csv.name}") 

        # Soit on a déjà un chemin, soit on va pointer le csv manuellement
        if chemin_csv is None:
            chemin_csv = self._filedialog_csv(trigramme_formation=self.trigramme_formation)
            if chemin_csv is None:
                vlog.log_erreur("Aucun chemin CSV fourni pour le traitement.", continuer=True)
                return 

        # Si le chemin est avec un raccourci réseau alors on récupère le chemin en entier + on convertit en Path 
        chemin_csv = chemin_vers_unc(Path(chemin_csv)) 
        chemin_excel = chemin_csv.with_suffix(".xlsx")






        # Étape 0 — Vérifications
        # 0.1 : on vérifie que l'eval formation est ouvert en mémoire sinon c'est le premier appel → on ouvre ou crée
        if self.eval_formation.fe is None:
            self.eval_formation._ouvrir_ou_creer_eval_formation()

        # 0.2  on vérifie que chemin_csv_session n'est pas déjà dans le fichier évaluations des formations pour savoir si on l'exclue du traitement
        if str(chemin_csv) in self.eval_formation.df_stagiaires["Chemin fichier CSV"].drop_duplicates().tolist():  
            print(f"⚠️  Exclusion car csv déjà dans le fichier global : {chemin_csv}")
            #pprint(self._df_evaluations_formation["Chemin fichier CSV"])
            self._statut_csv = "Exclu - CSV déjà dans fichier global"
            return

        # 0.3 : on vérifie si le code_IRIS est bien existant dans l'extract IRIS
        df_IRIS_sessions = get_iris(typeExport="Sessions", chemin=chemin_IRIS_sessions).df  # On charge IRIS
        if self.code_IRIS not in df_IRIS_sessions["Code IRIS"].values:
            print(f"⚠️  Code IRIS {self.code_IRIS} non trouvé dans l’extract IRIS.")
            self._statut_csv = "Exclu - Code IRIS pas dans Extract IRIS sessions"
            return






        # Étape 1 — On charge le CSV dans df_csv_stagiaires
        df_csv_stagiaires = self._charger_csv_stagiaire(chemin_csv=chemin_csv)
        # S'il y a eu un problème dans _charger_csv_stagiaire / le dataframe est vide (csv présent avec en-têtes mais sans ligne) → on saute la fin du traitement
        if (df_csv_stagiaires is None) or (df_csv_stagiaires.empty):
            return
        


        # Étape 2 — Générer l'excel des évaluations des stagiaire (_fe_evaluations_stagiaires : 2 onglets + TCD)
        # 2.1 : On crée le fichier Excel modèle depuis le modèle puis on affecte df_csv_stagiaire
        self._fe = FichierExcel.depuis_modele(chemin_modele=config.CHEMIN_MODELE_EXCEL_EVALUATIONS_STAGIAIRES, chemin_fichier_sauv=chemin_excel)
        self.df_csv = df_csv_stagiaires

        # 2.2 : On crée les dataframes
        self._traiter_df_csv()  # Traitement onglet CSV_stagiaires (import "direct" du CSV avec quelques traitements mineurs)
        self._traiter_df_stagiaires()  # Traitement seconde partie du dataframe du CSV

        # 2.3 : On écrit les dataframes dans les 2 onglets
        self._ecrit_df_et_sauve()
        self._statut_csv = "Traité"



        # Étape 3 — On met à jour le DataFrame de formation partagé
        # 3.1 : On met à jour les DataFrame d'evalFormation
        self._maj_dataframes_eval_formation()

        # 3.2 : On sauvegade selon argument utilisateur ; normalement si et seulement si nous ne faisons pas un traitement en boucle (sinon on le fait en fin de traitement de boucle)
        if ecrire_eval_formation:
            self.eval_formation.ecritdf_et_sauve_siModif()
    
    
    
        # Ouverture du dossier à la fin si demandé
        if ouvrirDossier:
            ouvrir_dossier(self.chemin_fe.parent)

        timer.fin()       

    # ==================================================================================
    # MÉTHODES EXTERNES
    # ==================================================================================
    @staticmethod
    def construire_dictionnaire_trigrammeFormation_codeIRIS_cheminsCSV_depuis_iterableCSV(chemins_csv: Iterable[Union[str, Path]]) -> Dict[str, Dict[int, Path]]:
        """
        Construit un dictionnaire imbriqué de la forme {trigramme_formation: {code_IRIS: chemin_csv}} à partir de chemins CSV.

        Ce dictionnaire s'emploie pour faciliter l'appel d'une boucle de traitement d'EvalStat

        :param chemins_csv: Un itérable de chemins CSV qui sont de type str ou Path.
        :type chemins_csv: Iterable[Union[str, Path]]
        :return: Un dictionnaire imbriqué où les clés de premier niveau sont les trigrammes de formation,
            et les clés de second niveau sont les codes IRIS.
        :rtype: Dict[str, Dict[int, Path]]
        """
        dictionnaire = {}

        for chemin in chemins_csv:
            # Convertir le chemin en objet Path si ce n'est pas déjà fait
            chemin_path = Path(chemin) if isinstance(chemin, str) else chemin

            # Récupérer le trigramme de la formation
            trigramme_formation = recupere_trig_formation_depuis_chemin(chemin_path)

            # Récupérer le code IRIS
            code_IRIS = IRIS.extraire_code_IRIS_depuis_chemin(chemin_path)

            # Initialiser le sous-dictionnaire pour le trigramme si nécessaire
            if trigramme_formation not in dictionnaire:
                dictionnaire[trigramme_formation] = {}

            # Ajouter l'entrée au sous-dictionnaire
            dictionnaire[trigramme_formation][code_IRIS] = chemin_path

        return dictionnaire


    @staticmethod
    def construire_dictionnaire_trigrammeFormation_codeIRIS_cheminsCSV_depuis_iterableCodesIRIS(codes_IRIS: Iterable[int]) -> Dict[str, Dict[int, Path]]:
        """
        Construit un dictionnaire imbriqué de la forme {trigramme_formation: {code_IRIS: chemin_csv}} à partir de codes IRIS.

        Ce dictionnaire s'emploie pour faciliter l'appel d'une boucle de traitement d'EvalStat.

        :param codes_IRIS: Un itérable de codes IRIS de type int
        :type codes_IRIS: Iterable[int]
        :return: Un dictionnaire imbriqué où les clés de premier niveau sont les trigrammes de formation,
            et les clés de second niveau sont les codes IRIS.
        :rtype: Dict[str, Dict[int, Path]]
        """
        dictionnaire = {}

        for code_IRIS in codes_IRIS:
            # Convertir le chemin en objet Path si ce n'est pas déjà fait
            chemin_path = EvalStat_session.filedialog_csv(code_IRIS=code_IRIS)

            # Récupérer le trigramme de la formation
            trigramme_formation = recupere_trig_formation_depuis_chemin(chemin_path)

            # Initialiser le sous-dictionnaire pour le trigramme si nécessaire
            if trigramme_formation not in dictionnaire:
                dictionnaire[trigramme_formation] = {}

            # Ajouter l'entrée au sous-dictionnaire
            dictionnaire[trigramme_formation][code_IRIS] = chemin_path

        return dictionnaire

    # ==================================================================================
    # POPUP
    # ==================================================================================
    def _filedialog_csv(self) -> Path | None:
        """
        Ouvre un filedialog pour demander à l'utilisateur de sélectionner un CSV.
        On pointe au mieux sur le répertoire des CSV de cette formation pour la boîte de dialogue.
        """
        return self.filedialog_csv(self.trigramme_formation, self.code_IRIS)


    @staticmethod
    def filedialog_csv(trigramme_formation:Optional[str]=None, code_IRIS:Optional[int]=None) -> Path | None:
        """
        Ouvre un filedialog pour demander à l'utilisateur de sélectionner un CSV.

        On pointe au mieux sur le répertoire des CSV de cette formation pour la boîte de dialogue.

        Le trigramme permet d'optimiser le répertoire de recherche (non obligatoire).

        Le code IRIS permet d'être spécifié dans l'en-tête du filedialog (non obligatoire).

        :param trigramme_formation: Trigramme de la formation. Permet d'optimiser le répertoire de recherche (non obligatoire).
        :type trigramme_formation: Optional[str]
        :param code_IRIS: Code IRIS de la session. Permet d'être spécifié dans l'en-tête du filedialog (non obligatoire).
        :type code_IRIS: Optional[int]
        :return: Description
        :rtype: Path | None
        """

        # Si on a un trigramme de formation, alors on est en mesure de trouver un chemin optimisé
        if trigramme_formation:
            chemin_repertoire_csv = optimiseCheminRepertoire(
                config.format_path(config.REPERTOIRE_CSV_EVALUATIONS, trigramme_formation=trigramme_formation)
                )
        else:
            chemin_repertoire_csv = optimiseCheminRepertoire(config.REPERTOIRE_FORMATION.parent)  #Path.cwd()  
            
        # Adaptation de l'intitulé de l'en-tête de la popup
        fin_titre = f"de la session {code_IRIS}" if code_IRIS is not None else "désiré"

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

    @property
    def statut_csv(self) -> str:
        return self._statut_csv




# ======================================================================================
# CLASSE EVALSTAT_FORMATION
# ======================================================================================
class EvalStat_formation(EvalStat): 
    """
    Classe evalStat Formation employée en parllèle du traitement des évaluations stagiaires individuelles.

    Gère la création ou la mise à jour du fichier d'évaluation de la formation au format xlsx.
    """

    # ==================================================================================
    # CONSTRUCTEURS
    # ==================================================================================
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
        Initialisation d'un EvalStat formation avec (ouverture ou création) du fichier d'évaluation de la formation
        """
        instance = cls(formation)

        instance._ouvrir_ou_creer_eval_formation()

        return instance


    # ==================================================================================
    # METHODES INTERNES
    # ==================================================================================
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

        # On définit le chemin de destination du fichier excel d'évaluation depuis le fichier config
        chemin_excel_eval = config.format_path(config.CHEMIN_EXCEL_EVALUATIONS_FORMATION, trigramme_formation=self.trigramme_formation)
        
        # On regarde si Evaluation-Stagiaires-Global-XXX.xlsx existe
        # Auquel cas on l'ouvre et on ne supprimera pas les données existantes
        if chemin_excel_eval.is_file():
            #print("Ouverture EvalStat formation existant")

            # Alors on l'ouvre
            self._fe = FichierExcel.depuis_fichier(chemin_fichier=chemin_excel_eval)
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
            chemin_excel_eval.parent.mkdir(parents=True, exist_ok=True)

            # On créée le fichier excel à partir du modèle
            self._fe = FichierExcel.depuis_modele(
                chemin_modele = config.CHEMIN_MODELE_EXCEL_EVALUATIONS_STAGIAIRES, 
                chemin_fichier_sauv = chemin_excel_eval
                )
            vlog.ajouter_message("Création EvalStat Global formation", self.chemin_fe, style=["vert"])
            
            # Il faudra supprimer les anciennes données de fe_evaluations_formation
            self._supprimeEtRemplace_donneesEval = True

        # On fait un hash du df pour savoir si, en fin de traitement il aura été modifié, auquel cas on le sauvegardera
        self._df_initial_hash = hash_df(self.df_stagiaires)
        timer.fin()

        #return self._fe_evaluations_formation, self._df_evaluations_formation, supprimeDonneesEtRemplace


    # ==================================================================================
    # METHODES EXTERNES
    # ==================================================================================
    def ecritdf_et_sauve_siModif(self):
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

    # ==================================================================================
    # GETTERS / SETTERS
    # ==================================================================================

    @property
    def trigramme_formation(self) -> str:
        return self._formation.trigramme_formation

