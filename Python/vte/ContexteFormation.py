from __future__ import annotations

from vte.core.config import UNITE
from .utils.office import *

import math
from babel.dates import format_date

from dataclasses import dataclass
from collections import defaultdict
from tabulate import tabulate

from mailmerge import MailMerge
from tkinter import ttk, messagebox

from pprint import pprint

from contextlib import AbstractContextManager

from typing import TYPE_CHECKING
if TYPE_CHECKING:
    import vte.core.config as ConfigType  # pour que Pylance ait une base d’autocomplétion

config, config_extractsIRIS, user_config = charger_config()
config:ConfigType  # type hint explicite


# ======================================================================================
# CLASSE CONTEXTE FORMATION
# ======================================================================================
class Contexte_formation(AbstractContextManager):
    """
    Gestionnaire de contexte pour une formation donnée.

    Ce contexte permet :
    - d'ouvrir ou créer une fois pour toutes le fichier Excel des évaluations d'une formation
      (commune à plusieurs sessions ou stagiaires) ;
    - de garantir sa sauvegarde et sa fermeture automatique à la fin du traitement,
      que celui-ci concerne un ou plusieurs fichiers CSV.

    Utilisation typique :
        with Contexte_formation("ABC"):
            EvalStat.depuis_chemin_csv_evaluations_stagiaires(chemin_csv)

    ou bien, pour un traitement multiple :
        with Contexte_formation("ABC"):
            for chemin_csv in liste_csv:
                EvalStat.depuis_chemin_csv_evaluations_stagiaires(chemin_csv)
    """

    # === VARIABLES DE CLASSE COMMUNES À TOUS LES CONTEXTES ===

    # Modèle Excel évaluations stagiaires (commun à toutes les formations)
    _CHEMIN_MODELE_EXCEL_EVALUATIONS_STAGIAIRES: Path = config.CHEMIN_MODELE_EXCEL_EVALUATIONS_STAGIAIRES

    # === CONSTRUCTEUR ===
    def __init__(self, trigramme_formation: str):
        """
        Initialise le contexte pour une formation spécifique.

        Args:
            trigramme_formation (str): Trigramme identifiant la formation.
        """
        self._trigramme_formation:str = trigramme_formation
        self._fe_evaluations_formation: Optional[FichierExcel] = None
        self._df_evaluations_formation: Optional[pd.DataFrame] = None
        self._supprimeDonneesEtRemplace_evaluations_formation: Optional[bool] = None

    # === ENTRÉE DANS LE CONTEXTE ===
    def __enter__(self):
        """
        Ouvre ou crée le fichier Excel des évaluations de la formation.

        Le fichier est partagé entre toutes les instances de `EvalStat`
        créées pendant ce contexte.
        """
        timer.debut(f"[Contexte_formation] Ouverture ou création du fichier Excel de la formation {self._trigramme_formation}")

        # Ouverture ou création du fichier Excel de la formation
        self._fe_evaluations_formation, self._supprimeDonneesEtRemplace_evaluations_formation = self._ouvrir_ou_creer_evaluations_formation(self._trigramme_formation)
        self._df_evaluations_formation = self._fe_evaluations_formation._tableaux["Stagiaires"].df  # Alias
        self._df_evaluations_formation["Code IRIS"] = self._df_evaluations_formation["Code IRIS"].astype(str)
        self._df_initial_hash = hash_df(self._fe_evaluations_formation._tableaux["Stagiaires"]._df)
        timer.fin()

        # Propagation dans EvalStat pour partage entre instances
        EvalStat._CHEMIN_MODELE_EXCEL_EVALUATIONS_STAGIAIRES = self._CHEMIN_MODELE_EXCEL_EVALUATIONS_STAGIAIRES
        EvalStat._trigramme_formation = self._trigramme_formation
        EvalStat._fe_evaluations_formation = self._fe_evaluations_formation
        EvalStat._df_evaluations_formation = self._df_evaluations_formation
        EvalStat._supprimeDonneesEtRemplace_evaluations_formation = self._supprimeDonneesEtRemplace_evaluations_formation

        return self

    # === SORTIE DU CONTEXTE ===
    def __exit__(self, exc_type, exc_value, traceback):
        """
        Sauvegarde et fermeture du fichier Excel de la formation à la sortie du contexte.

        Ce bloc est exécuté même en cas d'erreur dans le traitement des sessions.
        """
        # TODO : dans _mettre_a_jour_evaluations_formation() je mets à jour le df de l'excel évaluation formation. Il sera écrit physiquement à la sortie du contexte formation.
        # TODO : df nouveau df comprend l'ancien (i.e. évaluation formation existant) + le nouveau que l'on traite.
        # TODO : pour l'instant je réécrit tout ce df mais pour être optimal on ne pourrait écrire que le nouveau


        if self._fe_evaluations_formation is not None:

            df_final_hash = hash_df(self._fe_evaluations_formation._tableaux["Stagiaires"]._df)

            if df_final_hash != self._df_initial_hash:
                #TODO : pour l'instant je force à tout réécrire et pas seulement faire les mises à jour
                # Forçage réécriture en entier du df
                EvalStat._supprimeDonneesEtRemplace_evaluations_formation = True



                # Le DataFrame a changé → on sauvegarde
                EvalStat._sauver_excel_evaluations_formation(fermer_fichier=False)
            else:
                vlog.ajouter_message(
                    "Aucune modification détectée → pas de sauvegarde",
                    self._fe_evaluations_formation.chemin_fichier,
                    style=["jaune"]
                )

        # Nettoyage des références dans EvalStat
        EvalStat._CHEMIN_MODELE_EXCEL_EVALUATIONS_STAGIAIRES = None
        EvalStat._trigramme_formation = None
        EvalStat._fe_evaluations_formation = None
        EvalStat._df_evaluations_formation = None
        EvalStat._supprimeDonneesEtRemplace_evaluations_formation = None

        return False  # Ne supprime pas d’éventuelles exceptions

    # === MÉTHODE INTERNE ===
    @staticmethod
    def _ouvrir_ou_creer_evaluations_formation(trigramme_formation: str) -> Tuple[FichierExcel, bool]:
        """
        Ouvre ou crée le fichier Excel d'évaluations d'une formation.

        Cette méthode construit le chemin vers le fichier d'évaluations correspondant au 
        trigramme de la formation. Si ce fichier existe, il est ouvert et les données 
        des stagiaires sont chargées dans un DataFrame. Sinon, un nouveau fichier est 
        créé à partir d'un modèle, et les données seront à initialiser.
        
        Args:
            trigramme_formation (str): Trigramme de la formation.
        Returns:
            - FichierExcel: Objet FichierExcel ouvert.
            - Un booléen indiquant si les anciennes données doivent être supprimées et remplacées 
                (`True` si nouveau fichier créé, `False` sinon).
        """
        # On définit le chemin du répertoire depuis le fichier config
        chemin_excel_evaluations_formation = config.format_path(config.CHEMIN_EXCEL_EVALUATIONS_FORMATION, trigramme_formation=trigramme_formation)

        # On vérifie que le répertoire dédié existe sinon on le créée : \\instnt\partage\FORMATIONS_C\###\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\AAAA
        chemin_excel_evaluations_formation.parent.mkdir(parents=True, exist_ok=True)
        
        # On regarde si Evaluation-Stagiaires-Global-XXX.xlsx existe
        if chemin_excel_evaluations_formation.is_file():
            # Alors on l'ouvre
            fe = FichierExcel.depuis_fichier(chemin_fichier=chemin_excel_evaluations_formation)
            #self._fe_evaluations_formation._tableaux["CSV_stagiaires"].charge_df()
            #self._fe_evaluations_formation._tableaux["Stagiaires"]._df['Trigramme formation'] = self._fe_evaluations_formation._tableaux["Stagiaires"]._df['Trigramme formation'].astype(str)

            
            # On récupère la liste des sessions déjà intégrées (liste des codes et des chemins) → Ce sera pour écrire dans le tkinter
            #dico_sessionsDejaTraitees = dict(
            #df_formation_stagiaires[df_formation_stagiaires["Trigramme formation"] == trigramme]     # 1. filtre sur le trigramme
            #.drop_duplicates(subset=["Code IRIS", "Chemin fichier CSV"])  # 2. élimine les doublons
            #[["Code IRIS", "Chemin fichier CSV"]]            # 3. sélection des colonnes
            #.values                               # 4. valeurs du DF
            #    )        

            # Il ne faudra pas supprimer les anciennes données de fe_evaluations_formation
            supprimeDonneesEtRemplace = False

            vlog.ajouter_message("Ouverture EvalStat Global formation", fe.chemin_fichier, style=["vert"])
        else:
            # On créée le fichier excel à partir du modèle
            fe = FichierExcel.depuis_modele(
                chemin_modele = Contexte_formation._CHEMIN_MODELE_EXCEL_EVALUATIONS_STAGIAIRES, 
                chemin_fichier_sauv = chemin_excel_evaluations_formation
                )
            
            # Il faudra supprimer les anciennes données de fe_evaluations_formation
            supprimeDonneesEtRemplace = True

            vlog.ajouter_message("Création EvalStat Global formation", fe.chemin_fichier, style=["vert"])

        #return self._fe_evaluations_formation, self._df_evaluations_formation, supprimeDonneesEtRemplace
        return fe, supprimeDonneesEtRemplace














