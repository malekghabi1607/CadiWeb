from __future__ import annotations
from pathlib import Path
from typing import Optional

import pandas as pd

from vte.utils.office import FichierExcel

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
