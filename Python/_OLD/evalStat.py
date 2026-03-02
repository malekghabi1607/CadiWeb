from __future__ import annotations

from collections import defaultdict
from pathlib import Path
from typing import Optional

from vte.core import config
from vte.utils.utils import *
from vte.utils.utils_instn import *
from vte.utils.office import FichierExcel
from vte.domain.session import Session
from vte.iris import *


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
    _IRIS_sessions:Optional[IRIS_traite] = None


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
    def __init__(self, session: Optional[Session] = None, chemin_csv: Optional[Path] = None) -> None:
        """
        Initialise une instance EvalStat individuelle (pour un stagiaire/session).

        :param session: objet de type session pour laquelle on va faire l'EvalStat
        :type session: Optional[Session]
        :param chemin_csv: Chemin vers le CSV d'évaluation des stagiaires.
        :type chemin_csv: Optional[Path]
        """
       

        # Variables propres à un EvalStat individuel
        self._session:Optional[Session] = session
        self._chemin_csv: Optional[Path] = chemin_csv  # Fichier CSV d'entrée (CSV EvalStat)

        self._fe: Optional[FichierExcel] = None  # Objet Excel contenant les données EvalStat stagiaire individuel


        # === Résultat du traitement ===
        self._statut_csv: Optional[str] = None  # ex: "Traité", "Exclu - CSV déjà dans fichier global", "Exclu - Code IRIS pas dans Extract IRIS sessions", "Exclu - Problème lecture CSV", "Exclu - CSV vide / Aucun retour"





    """
    self._IRIS_sessions.chemin
    self._IRIS_sessions.fe
    self._IRIS_sessions.df

        _chemin_IRIS_sessions:Path=None
        _fe_IRIS_sessions: Optional[FichierExcel] = None  # Fichier Excel contenant l'extract IRIS Sessions (ou celles de la période en cours)
        _df_IRIS_sessions: Optional[pd.DataFrame] = None  # Alias du dataframe
    """


    # =========================
    # === GETTERS / SETTERS ===
    # =========================
    def IRIS_sessions(self) -> IRIS_traite:
        """
        Renvoie l'instance d'IRIS 

        Returns:
            FichierExcel: l'instance d'IRIS 
        """
        return self._IRIS_sessions
    
    def code_IRIS(self) -> int:
        """
        Renvoie le code IRIS de l'EvalStat.
        (Pointe vers l'objet session.)

        Returns:
            int: le code IRIS de l'EvalStat
        """

