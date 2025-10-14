# Exploitation Word
#https://pbpython.com/python-word-template.html

# Optimisation vitesse
#https://progresser-en-maths.com/la-maniere-la-plus-efficace-diterer-sur-des-lignes-dans-un-dataframe-pandas/
#https://pandas.pydata.org/docs/user_guide/enhancingperf.html
#https://moncoachdata.com/blog/7-techniques-doptimisation-de-la-memoire-avec-pandas/

#Tableaux structurés
#https://openpyxl.readthedocs.io/en/3.1/api/openpyxl.worksheet.table.html#module-openpyxl.worksheet.table
#https://openpyxl-readthedocs-io.translate.goog/en/3.1/worksheet_tables.html?_x_tr_sl=en&_x_tr_tl=fr&_x_tr_hl=fr&_x_tr_pto=sc




from __future__ import annotations

from .utils import *

from typing import Dict, List, Tuple, Optional, Union

import pandas as pd
import pandas as DataFrame

from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.worksheet.table import Table
from openpyxl.utils.cell import range_boundaries, get_column_letter
from openpyxl.worksheet.cell_range import CellRange

import xlwings as xw

from docx import Document
from lxml import etree

import win32com.client
import extract_msg
import os
import shutil
import sys
import time
import copy
import json

from datetime import date, datetime, timedelta, time
from zoneinfo import ZoneInfo
from io import StringIO

from tqdm import tqdm
import colorama
from colorama import Fore, Style

from tkinter import filedialog, Tk

### --------------------------------------------------------------------
#  Définitions classes et fonctions génériques
### --------------------------------------------------------------------

# fonction utilitaire pour nettoyage (présente dans ton code)
"""def nettoyer_nom_colonne(col_name: str) -> str:
    # exemple simple (laisser tel quel si tu as ta propre version)
    if isinstance(col_name, str):
        return col_name.strip()
    return col_name"""


class FichierExcel:
    """
    Classe pour manipuler un fichier Excel avec openpyxl.
    Responsable de l’ouverture, de la gestion du workbook et du suivi des tableaux Excel contenus dans le fichier.

    Usage (exemples):

    >>> # ouvrir un fichier et charger tous les tableaux structurés
    >>> fx = FichierExcel.depuis_fichier("C:/mon/chemin/fichier.xlsx")

    >>> # ouvrir un fichier, mais charger un seul tableau (onglet "Feuil1")
    >>> fx = FichierExcel.depuis_fichier("C:/mon/chemin/fichier.xlsx", nom_onglet="Feuil1")

    >>> # récupérer un tableau par nom et écrire un DataFrame dedans
    >>> tab = fx.get_tableau("Sessions")
    >>> tab.ecrit_dataFrame_dans_tableauStructure(df)

    Notes:
    - L'objet _TableauExcel est semi-privé: ses méthodes d'intérêt sont accessibles, mais on évite d'instancier
      des tableaux directement hors de FichierExcel.
    - Les responsabilités sont séparées : FichierExcel = gestion fichier + workbook ;
      _TableauExcel = manipulation de données tabulaires dans un onglet.

    # - J'ai supprimé weakref (inutile dès lors que la relation parent/enfant est directe).
    """
# === Classe imbriquée semi-privée : _TableauExcel ===
    class _TableauExcel:
        """
        Représente un tableau contenu dans un fichier Excel, pouvant être un tableau structuré (Excel Table)
        ou un tableau "normal" (données tabulaires sans table Excel explicite).

        IMPORTANT: Cette classe est imbriquée dans FichierExcel et doit être créée via les méthodes de fabrique
        de FichierExcel (ex: FichierExcel.charger_tableau, FichierExcel.ajouter_tableau_normal, etc.).

        Exemple d'utilisation (via FichierExcel) :

        >>> fx = FichierExcel.depuis_fichier("chemin.xlsx")
        >>> fx.charger_tableau("Feuil1", nom_tableau="Sessions")
        >>> tab = fx.get_tableau("Sessions")
        >>> tab.charge_df()  # charge le DataFrame depuis le fichier
        >>> tab.ecrit_dataFrame_dans_tableauStructure(df)
        """

        # --- Attributs internes (semi-privés) ---
        _parent: FichierExcel  # référence vers FichierExcel (pas de weakref nécessaire)
        _ws: Optional[Worksheet] = None
        _nom_onglet: Optional[str] = None
        _est_tableau_structure: Optional[bool] = None
        _df: Optional[pd.DataFrame] = None

        # si tableau structuré
        _nom_tableau: Optional[str] = None
        _table: Optional[Table] = None

        # si tableau normal
        _nbLignes_avantET: int = 0

        # dimensions / ref
        _ref_tableau: Optional[str] = None  # Pour mémoriser la référence (ex: A1:D12)
        _min_row: Optional[int] = None
        _max_row: Optional[int] = None
        _min_col: Optional[int] = None
        _max_col: Optional[int] = None
        _nb_lignes_tableau: Optional[int] = None
        _nb_lignesData_tableau: Optional[int] = None

        # === Constructeur principal (privé) ===
        def __init__(self,
            parent: FichierExcel,
            ws: Optional[Worksheet] = None,
            nom_onglet: Optional[str] = None,
            nom_tableau: Optional[str] = None,
            nbLignes_avantET: int = 0) -> None:
            """
            Ne pas appeler directement: utiliser FichierExcel.depuis_fichier / charger_tableau / ajouter_tableau_normal
            """
            # Informations du fichierExcel parent
            self._parent = parent

            # Informations de la classe
            self._ws = ws
            self._nom_onglet = nom_onglet

            # Informations si tableau structuré (peuvent être être None si on a un tableau normal)
            self._nom_tableau = nom_tableau

            # Informations si tableau normal (peuvent être être None si on a un tableau structuré)
            self._nbLignes_avantET = nbLignes_avantET

        # === Constructeurs alternatifs (factory methods) ===
        @classmethod
        def depuis_FichierExcel_et_nomOnglet(cls,
            parent: FichierExcel,
            nom_onglet: str,
            nom_tableau: Optional[str] = None,
            nbLignes_avantET: int = 0,
            charger_df: bool = True) -> FichierExcel._TableauExcel:
            """
            Crée un _TableauExcel à partir d'un FichierExcel déjà chargé (parent._wb doit exister).
            Charge les dimensions selon qu'il s'agisse d'un tableau structuré ou non.
            """
            instance = cls(parent=parent,
                           ws=parent._wb[nom_onglet],
                           nom_onglet=nom_onglet,
                           nom_tableau=nom_tableau,
                           nbLignes_avantET=nbLignes_avantET)

            # Si pas de nom_tableau fourni, on prend le nom de l'onglet comme convention
            instance._nom_tableau = nom_tableau or instance._nom_onglet

            instance._est_tableau_structure = instance.est_tableau_structure()
            if instance._est_tableau_structure:
                instance._table = instance._ws.tables.get(instance._nom_tableau)  # bak = instance._table = instance._ws.tables[instance._nom_tableau]
                if instance._table is None:
                    # dans certaines versions, ws.tables peut être un dict-like ; on essaie une recherche
                    found = next((t for t in instance._ws._tables.values() if getattr(t, "name", None) == instance._nom_tableau), None)
                    instance._table = found
                instance._initialiser_dimensions_structured_table()
            else:
                instance._nbLignes_avantET = nbLignes_avantET
                instance._initialiser_dimensions_tableau_normal()

            if charger_df:
                instance.charge_df()

            return instance

        # === Initialisation dimensions ===
        def _initialiser_dimensions_structured_table(self) -> None:
            self._ref_tableau = self._table.ref
            self._min_col, self._min_row, self._max_col, self._max_row = range_boundaries(self._ref_tableau)
            self._nbLignes_avantET = self._min_row - 1  # nombre de lignes avant l’en-tête du tableau
            # On calcule nb_lignesData_tableau ici mais attention : self._table.headerRowCount existe
            self._nb_lignes_tableau = self._max_row - self._min_row + 1
            # headerRowCount peut ne pas exister selon ta version ; on défensive
            header_count = getattr(self._table, "headerRowCount", 1)  # Bak = self._table.headerRowCount
            self._nb_lignesData_tableau = self._nb_lignes_tableau - header_count

        def _initialiser_dimensions_tableau_normal(self) -> None:
            # Première ligne de données du tableau (juste après les lignes d'entête)
            self._min_row = self._nbLignes_avantET + 1
            self._max_row = self._ws.max_row

            min_col, max_col = None, 0

            # Parcours de la zone supposée du tableau pour trouver les vraies bornes
            for row in self._ws.iter_rows(min_row=self._min_row, max_row=self._max_row):
                for cell in row:
                    if cell.value not in (None, ""):
                        if min_col is None or cell.column < min_col:
                            min_col = cell.column
                        if cell.column > max_col:
                            max_col = cell.column

            # Si aucun contenu trouvé → on garde 1 par défaut
            self._min_col = min_col if min_col is not None else 1
            self._max_col = max_col if max_col > 0 else 1

            # Définir la référence Excel (par ex. "A5:AU20")
            self._ref_tableau = self.definir_rangeExcel()

            # Nombre de lignes total et de lignes de données
            self._nb_lignes_tableau = self._max_row - self._min_row + 1
            self._nb_lignesData_tableau = max(0, self._nb_lignes_tableau - 1)  # 1 ligne header approx.

        # === Méthodes publiques : chargements / réinitisation ===
        def reinitialiser_workbook_openpyxl(self, recalculer_dimensions: bool) -> None:
            """
            Réinitialise les objets liés au Workbook après rechargement.

            :param recalculer_dimensions: Si True, recalcule les dimensions du tableau.
            """
            if not self._parent or not self._parent._wb:
                vlog.log_erreur("Fichier Excel ou Workbook non disponible.")

            if self._nom_onglet is None:
                vlog.log_erreur("Nom d'onglet non défini dans le tableau.")

            # Réassignation du worksheet
            self._ws = self._parent._wb[self._nom_onglet]

            if self._nom_tableau:  # Tableau structuré
                self._est_tableau_structure = True
                self._table = next(
                    (t for t in self._ws._tables.values() if t.name == self._nom_tableau),
                    None
                )

                if self._table is None:
                    vlog.log_erreur(f"Tableau structuré '{self._nom_tableau}' non trouvé dans l’onglet '{self._nom_onglet}'.")

                if recalculer_dimensions:
                    self._initialiser_dimensions_structured_table()

            else:  # Tableau non structuré
                self._est_tableau_structure = False

                if recalculer_dimensions:
                    self._initialiser_dimensions_tableau_normal()

        def charger_workbook(self, wb: Optional[Workbook] = None) -> None:
            """
            Méthode à appeler manuellement si le workbook est connu plus tard.
            Assigne le workbook parent et initialise la feuille et dimensions.
            """
            if wb:
                self._parent._wb = wb

            if self._nom_onglet is None:
                vlog.log_erreur("Le nom de l’onglet doit être défini avant de charger le workbook")
            self._ws = self._parent._wb[self._nom_onglet]

            if self._nom_tableau:
                self._table = self._ws.tables[self._nom_tableau]
                self._initialiser_dimensions_structured_table()
            else:
                self._initialiser_dimensions_tableau_normal()

        def charge_df(self) -> None:
            """
            Charge le DataFrame depuis le fichier Excel (pandas.read_excel) en sautant les lignes
            avant l'en-tête du tableau.
            """
            if not self._parent._chemin_fichier or not self._nom_onglet:
                vlog.log_erreur("Chemin fichier et nom onglet doivent être définis pour charger le DataFrame")

            skiprows = (self._min_row - 1) if self._min_row else 0
            self._df = pd.read_excel(self._parent._chemin_fichier, sheet_name=self._nom_onglet, skiprows=skiprows)
            self._df.columns = [nettoyer_nom_colonne(col) for col in self._df.columns]
            #print([repr(col) for col in self._df.columns])

        def remplace_df(self, df: DataFrame) -> None:
            """Remplace le DataFrame interne par une copie du df fourni."""
            self._df = df.copy()

        # === Méthodes publiques : écriture dans Excel ===
        def ecrit_dataFrame_dans_tableauStructure(self,
                                                 df: DataFrame,
                                                 supprimeDonneesEtRemplace: bool = False,
                                                 remplace_df_par_nouveau: bool = False) -> None:
            """
            Écrit un DataFrame dans un tableau structuré d'une feuille de calcul.

            :param df: DataFrame à intégrer.
            :param supprimeDonneesEtRemplace: True pour supprimer les anciennes données et les remplacer.
            :param remplace_df_par_nouveau: True pour remplacer l'attribut _df interne par df.
            :Example:
            >>> fichier = FichierExcel.depuis_fichier("chemin.xlsx")
            >>> fichier.charger_tableau("Feuil1", nom_tableau="Sessions")
            >>> tab = fichier.get_tableau("Sessions")
            >>> tab.ecrit_dataFrame_dans_tableauStructure(df, supprimeDonneesEtRemplace=True)
            
            .. seealso:: Rien du tout.
            .. warning:: Rien du tout.
            .. note:: Pour rajouter des lignes, on le fait à la suite de la feuille (worksheet) . Il en résulte qu'on gruge un peu : 1) on vire le tableau structuré, 2) on colle toutes les nouvelles lignes, 3) on supprime la ligne 1 du tableau, 4) on redéfinit les dimensions du tableau structuré (car l'ajout de nouvbelles lignes ne l'étend pas automatiquement)
            .. todo:: Rien du tout.
            """
            if remplace_df_par_nouveau:
                self.remplace_df(df)

            # On écrit toutes les autres lignes une par une (on garde les lignes initiales pour garder le format qu'on copiera)
            with tqdm(total=len(df), unit=' ligne', desc=Fore.CYAN + f"Écriture des lignes dans l'output {self.nom_tableau}" + Style.RESET_ALL) as pbar:
                for i, ligne_brute in enumerate(df.itertuples(index=False), 1):
                    pbar.set_postfix(progress=f"{i}/{len(df)}")
                    ligne = [val if pd.notna(val) else None for val in ligne_brute]  #Je dois rajouter cette ligne car il faut tester si je n'ai pas de valeurs <NA> qu'il faut retravailler sinon ça plante
                    self._ws.append(ligne)
                    pbar.update(1)

            # On redéfinit le dimensionnement du tableau (tableau initial + nb lignes de df)
            self.change_references_tableauStructure_par_deltaNbLignes(len(df), avec_initialiation=False)  # Il ne faut pas réinitialiser les dimensions du tableau car pour ma copie des formats j'ai besoin des anciennes références

            # On copie le format sur toutes les nouvelles lignes du tableau à partir de la 1ère ligne, i.e. self._min_row + 1
            self.copierFormat_tableauStructure_xlwings(indexLigneSourceFormat=self._min_row + 1,
                                                      indexLigneDebutCopie=self._max_row + 1,
                                                      indexLigneFinCopie=self._max_row + len(df))

            # Si désiré par l'utilisateur, alors on supprime les anciennes lignes de la feuille Excel
            if supprimeDonneesEtRemplace:
                self._ws.delete_rows(idx=self._min_row + 1, amount=self._nb_lignesData_tableau)  # ça garde la dimension initiale du tableau structuré
                self.change_references_tableauStructure_par_nbLignes(len(df))

            # Vérifier si la première ligne du DataFrame est entièrement vide (tous les éléments NaN)
            # J'ai vir& car a priori pas besoin si on fait supprimeDonneesEtRemplace
            ##if self.ligne_vide_ou_formules(self._min_row + 1): #df.iloc[0].isna().all():
            #    print("La première ligne est entièrement vide")
            #    # Suppression de la 1ère ligne
            #    self._ws.delete_rows(idx=self._min_row + 1, amount=1)
            #
            #    # On redéfinit les dimensions du tableau structuré
            #    self._table.ref = f"{self._ws.cell(row=self._min_row, column=self._min_col).coordinate}:{self._ws.cell(row=self._min_row + len(df) - 1, column=self._max_col).coordinate}"  


            # Mettre à jour les références des mises en formes conditionnelles
            self.maj_references_misesEnFormeConditionnelles()

        # === Copie formats / xlwings / openpyxl ===
        def copieFormat_tableauStructure_openpyxl(self,
                                                  indexLigneSourceFormat: int = 1,
                                                  indexLigneDebutCopie: Optional[int] = None,
                                                  indexLigneFinCopie: Optional[int] = None) -> None:
            """
            Recopie le format d'une ligne d'un tableau structuré à une plage du tableau structuré (cellule par cellule).
            Note :  moins efficace que la version xlwings (COM), mais portable.
                    On met à jour les styles des cellules (openpywl ne sait pas insérer de lignes en conservant les formats ; par ailleurs on ne sait pas appliquer ça ligne par ligne ou colonne par colonne : on va donc le faire cellule par cellule)
                    On n'agrandit pas automatiquement le tableau structuré si l'on rajoute une cellule
                    Quand on agrandit un tableau en redéfinissant le ref, on ne colle pas le format
            """
            if not indexLigneDebutCopie:
                indexLigneDebutCopie = indexLigneSourceFormat + 1
            if not indexLigneFinCopie:
                indexLigneFinCopie = self._max_row

            # On récupère les formats de chaque cellule de la première ligne du tableau structuré 
            formats = []
            for icol in range(self._min_col, self._max_col + 1):
                formats.append(self._ws.cell(indexLigneSourceFormat, icol))
                #print(formats[icol-1].number_format)
            # print(formats)

            # On copie colle les formats par colonne
            with tqdm(total=self._max_col - self._min_col + 1, unit=' colonnes', desc=Fore.CYAN + "Copie des formats" + Style.RESET_ALL, ncols=150) as pbar:
                for icol in range(self._min_col, self._max_col + 1):  # TODO : pas sûr du +1
                    pbar.set_postfix(progress=f"{icol}/{self._max_col}")
                    source = formats[icol - self._min_col]  # BAK : source = formats[icol-1] #Je prends un index de liste et pas un numéro de colonne, donc -1
                    #print(source.number_format, source.number_format == "General")
                    
                    for il in range(indexLigneDebutCopie, indexLigneFinCopie + 1):  # BAK : je n'avais pas le +1 à indexLigneFinCopie
                        #print(il, icol, ws.cell(row=il, column=icol).value, source.number_format)
                        target = self._ws.cell(row=il, column=icol)
                        #target.font = copy.copy(source.font)
                        #target.border = copy.copy(source.border)
                        target.fill = copy.copy(source.fill)
                        #target.alignment = copy.copy(source.alignment)
                        #target.protection = copy.copy(source.protection)
                        target.number_format = copy.copy(source.number_format)
                    pbar.update(1)

        def copierFormat_tableauStructure_xlwings(self,
                                                 indexLigneSourceFormat: int = 1,
                                                 indexLigneDebutCopie: Optional[int] = None,
                                                 indexLigneFinCopie: Optional[int] = None,
                                                 conserverMiseEnFormeConditionnelleCopiee: Optional[bool] = False) -> None:
            """
            Copie rapidement le format d'une plage source vers une plage cible en utilisant xlwings (COM, équivalent au collage spécial > formats dans Excel).
            Cette méthode ouvre Excel en arrière plan, colle les formats, sauvegarde et ferme Excel.
            
            :param indexLigneSourceFormat: Ligne contenant le format modèle à copier (ex: 2 pour la ligne 2).
            :type indexLigneSourceFormat: int
            :param indexLigneDebutCopie: Première ligne de la plage cible (ex: 3 pour copier à partir de la ligne 3).
            :type indexLigneDebutCopie: int, optional
            :param indexLigneFinCopie: Dernière ligne de la plage cible.
            :type indexLigneFinCopie: int, optional
            :param conserverMiseEnFormeConditionnelleCopiee: Si True, conserve les règles de mise en forme conditionnelle copiées. Si False, ne garde que celles déjà présentes au début.
            :type conserverMiseEnFormeConditionnelleCopiee: bool, optional

            :return: Aucun. Le fichier Excel est modifié et enregistré.
            :rtype: None

            :example:
            >>> copier_format_rapide(
                    indexLigneSourceFormat=2,
                    indexLigneDebutCopie=3,
                    indexLigneFinCopie=100,
                    conserverMiseEnFormeConditionnelleCopiee=False
                )

            .. note::
                Cette fonction nécessite Microsoft Excel installé sur votre machine (Windows uniquement).

            .. warning::
                Seuls les formats (style, bordures, police, etc.) sont copiés. Les valeurs ne sont pas modifiées.
                Les règles de mise en forme conditionnelle peuvent être dupliquées par Excel lors de la copie ;
                ce comportement peut être contrôlé via le paramètre `conserverMiseEnFormeConditionnelleCopiee`.
            """
            #print("Nombre de règles avant copierFormat_tableauStructure_xlwings:", len(self._ws.conditional_formatting._cf_rules))
            
            # Avant copie : on regarde les mises en forme conditionnelles existantes
            cles_avant = set(self._ws.conditional_formatting._cf_rules.keys())

            # TODO : ? Faire fonction
            recharger_wb = False
            if self._parent._wb:
                recharger_wb = True
                # On ferme openpyxl
                self._parent.save()
                self._parent.close()

            # Ouvrir xlwings
            #self._parent._xw_app = xw.App(visible=False)
            #self._parent._xw_wb = self._parent._xw_app.books.open(self._parent._chemin_fichier)
            #self._parent._xw_ws = self._parent._xw_wb.sheets[self._nom_onglet]
            #self._parent._xw_ws.activate()
            self._parent.open_xlwings(nom_onglet=self._nom_onglet, visible=False)

            if not indexLigneDebutCopie:
                indexLigneDebutCopie = indexLigneSourceFormat + 1
            if not indexLigneFinCopie:
                indexLigneFinCopie = self._max_row

            range_modele = self.definir_rangeExcel(min_row=indexLigneSourceFormat, max_row=indexLigneSourceFormat)
            range_cible = self.definir_rangeExcel(min_row=indexLigneDebutCopie, max_row=indexLigneFinCopie)

            # Copie de la ligne du modèle (format uniquement)
            self._parent._xw_ws.range(range_modele).api.Copy()
            # Collage spécial des formats uniquement
            self._parent._xw_ws.range(range_cible).api.PasteSpecial(Paste=-4122)  # -4122 = xlPasteFormats

            # Sauve et ferme xlwings
            self._parent.close_xlwings(sauver=True)

            # On recharge le wb openpyxl
            if recharger_wb:
                self._parent.recharger_workbook_openpyxl(recalculerDimensionsTableau=False)

            # Après copie : on compare et on supprime les mises en forme conditionnelle rajoutées si nécessaire
            cles_apres = set(self._ws.conditional_formatting._cf_rules.keys())
            nouvelles_cles = cles_apres - cles_avant
            plage_cible_cr = CellRange(range_cible)

            if not conserverMiseEnFormeConditionnelleCopiee:
                for cle in nouvelles_cles:
                    if self.plages_chevauchent(cle, plage_cible_cr):
                        del self._ws.conditional_formatting._cf_rules[cle]

        def maj_references_misesEnFormeConditionnelles(self, nouvelle_plage: Optional[str] = None, nouvelle_ligne_max: Optional[int] = None) -> None:
            """
            Met à jour les plages des règles de mise en forme conditionnelle appliquées au tableau structuré.
            Voir docstring originelle pour le détail.
            """
            # Met à jour les dimensions du tableau
            if self._est_tableau_structure:
                self._initialiser_dimensions_structured_table()
            else:
                self._initialiser_dimensions_tableau_normal()

            min_col_t, min_row_t, max_col_t, max_row_t = (self._min_col, self._min_row, self._max_col, self._max_row)

            cf_rules = self._ws.conditional_formatting._cf_rules  # Récupérer toutes les règles CF (clé = plage, valeur = liste des règles)
            regles_a_rajouter = []  # Liste pour stocker les règles à ré-appliquer [(nouvelle_plage, règle), ...] (en effet comme on va ajouter des règles on va perturber le dictionnaire qui fait la boucle et ça peut créer des effets de bord)
            a_supprimer = []  # Liste des clés (objets cf_obj) à supprimer

            # On fait une copie de la liste des plages à traiter (pour pouvoir supprimer sans erreur)
            for cf_obj, regles in list(cf_rules.items()):  # C'est une boucle sur les plages et on obtient les règles appliquées à cette plage ; cf_obj est un objet ConditionalFormatting
                # Récupérer les coordonnées de l'ancienne plage
                multi_range = cf_obj.sqref  # c'est un type MultiCellRange

                for cell_range in multi_range:
                    ancienne_plage = cell_range.coord  # c’est la plage Excel sous forme de string, ex "A1:AG2"
                    #print(f"Traitement de la plage : {ancienne_plage}")
                    min_col_r, min_row_r, max_col_r, max_row_r = range_boundaries(ancienne_plage)

                    if not (min_col_t <= min_col_r <= max_col_r <= max_col_t and
                            min_row_t <= min_row_r <= max_row_r + 1 <= max_row_t + 1):
                        log_erreur(f"une règle est en dehors du tableau dont on ne l'applique pas.\nAncienne plage : {ancienne_plage}\nNouvelle plage : {self._ref_tableau}", continuer=True)
                        continue

                    # Déterminer la nouvelle plage
                    np_ = nouvelle_plage
                    if np_ is None:
                        np_ = FichierExcel.definir_rangeExcel(min_row_r, min_col_r, max_row_t + 1, max_col_r)  # BAK : f"{get_column_letter(min_col_r)}{min_row_r}:{get_column_letter(max_col_r)}{max_row_t + 1}"
                    if nouvelle_ligne_max is not None:
                        np_ = FichierExcel.definir_rangeExcel(min_row_r, min_col_r, nouvelle_ligne_max, max_col_r)  # BAK : f"{get_column_letter(min_col_r)}{min_row_r}:{get_column_letter(max_col_r)}{nouvelle_ligne_max}"

                    # Stocker toutes les règles à ré-appliquer
                    for regle in regles:
                        regles_a_rajouter.append((np_, regle))

                    # Marquer l'ancienne règle pour suppression
                    a_supprimer.append(cf_obj)

            # Supprimer les anciennes règles
            for cf_obj in a_supprimer:
                del self._ws.conditional_formatting._cf_rules[cf_obj]

            #print("Nombre de règles après suppression:", len(self._ws.conditional_formatting._cf_rules))

            # Ré-appliquer toutes les règles sur les nouvelles plages
            for np_, regle in regles_a_rajouter:
                self._ws.conditional_formatting.add(np_, regle)

        def recopier_formules_colonnes(self, indexLigneSourceFormat: int, indexLigneDebutCopie: int, indexLigneFinCopie: int) -> None:
            """
            Recopie les formules Excel des colonnes qui en ont dans la ligne source vers toutes les lignes
            entre indexLigneDebutCopie et indexLigneFinCopie (inclus).
            """
            for col in range(self._min_col, self._max_col + 1):
                cellule_modele = self._ws.cell(row=indexLigneSourceFormat, column=col)
                formule_modele = cellule_modele.value
                if isinstance(formule_modele, str) and formule_modele.startswith('='):
                    # Recopie de la formule dans les nouvelles lignes en ajustant la référence relative
                    # openpyxl ne propose pas de recalcul automatique des formules, donc on fait simple :
                    # On copie exactement la même formule (Excel la recalculera automatiquement)
                    for ligne in range(indexLigneDebutCopie, indexLigneFinCopie + 1):
                        self._ws.cell(row=ligne, column=col).value = formule_modele

        def supprimer_premiere_ligne_tableau_structure(self) -> None:
            """
            Supprime proprement la première ligne de données du tableau structuré,
            en ajustant la plage du tableau pour éviter les erreurs à l'ouverture dans Excel.
            """
            ligne_a_supprimer = self._min_row + 1
            self._ws.delete_rows(idx=ligne_a_supprimer, amount=1)

            nb_lignes_données = len(self._df) - 1  # -1 car on supprime une ligne
            #BAK = nouvelle_ref = f"{self._ws.cell(row=self._min_row, column=self._min_col).coordinate}:" \
            #              f"{self._ws.cell(row=self._min_row + nb_lignes_données, column=self._max_col).coordinate}"
            nouvelle_ref = self.definir_rangeExcel(max_row=self._min_row + nb_lignes_données) # J'essaye par rapport au BAK

            self._table.ref = nouvelle_ref  # Appliquer la nouvelle plage au tableau
            if self._table.autoFilter:  # Mettre à jour aussi le filtre automatique (sinon Excel râle)
                self._table.autoFilter.ref = self._table.ref

        def supprimer_ligne_excel_via_xlwings(self, ligne_excel: int) -> None:
            """
            Supprime une ligne via xlwings (COM) pour préserver la structure du tableau structuré.
            """
            if not os.path.exists(self._parent._chemin_fichier):
                #raise FileNotFoundError(f"Fichier Excel introuvable : {self._parent._chemin_fichier}")
                vlog.log_erreur(f"Fichier Excel introuvable : {self._parent._chemin_fichier}")
                
            # BAK
            #app = xw.App(visible=True)
            #app.display_alerts = True
            #app.screen_updating = True
            #wb = app.books.open(self._parent._chemin_fichier)
            #ws = wb.sheets[self._ws.title]
            #self._xw_ws.range(f"{ligne_excel}:{ligne_excel}").api.Delete()
            #wb.save(self._parent._chemin_fichier)
            #wb.close()
            #app.quit()

            self._parent.open_xlwings(nom_onglet=self._ws.title, visible=False)
            self._xw_ws.range(f"{ligne_excel}:{ligne_excel}").api.Delete()
            self._parent.close_xlwings(sauver=True)

            
        # === Méthodes publiques : fonctions de test ===
        def est_tableau_structure(self) -> bool:
            """
            Vérifie si un tableau structuré (ListObject) portant le nom donné existe bien dans la feuille Excel.
            """
            if not self._ws or not self._nom_tableau:
                vlog.log_erreur("Manque le worksheet ou le nom du tableau")
            return self._nom_tableau in self._ws.tables

        def ligne_contient_formules(self, ligne: int) -> bool:
            """Renvoie True si une des cellules de la ligne contient une formule Excel."""
            for col in range(self._min_col, self._max_col + 1):
                cell = self._ws.cell(row=ligne, column=col)
                if cell.data_type == 'f' or (cell.value is not None and isinstance(cell.value, str) and cell.value.startswith("=")):
                    return True
            return False

        def ligne_vide_ou_formules(self, ligne: int) -> bool:
            """Retourne True si toutes les cellules de la ligne sont vides ou contiennent une formule."""
            for col in range(self._min_col, self._max_col + 1):
                cell = self._ws.cell(row=ligne, column=col)
                if cell.value is not None:
                    if isinstance(cell.value, str) and cell.value.startswith("="):
                        continue
                    else:
                        return False
            return True

        # === Méthodes publiques : range et références ===
        def definir_rangeExcel(self, min_row: int = None, min_col: int = None, max_row: int = None, max_col: int = None) -> str:
            """
            Retourne un range Excel "A1:B13" à partir de numéros (lignes / colonnes)
            Si indices non renseignés, on prend les valeurs stockées dans l'instance.
            """
            if min_row is None:
                min_row = self._min_row
            if min_col is None:
                min_col = self._min_col
            if max_row is None:
                max_row = self._max_row
            if max_col is None:
                max_col = self._max_col
            return FichierExcel.definir_rangeExcel(min_row, min_col, max_row, max_col)

        def change_references_tableauStructure(self, min_row: int, min_col: int, max_row: int, max_col: int, avec_initialiation: bool = True) -> None:
            self._table.ref = self.definir_rangeExcel(min_row, min_col, max_row, max_col)
            if avec_initialiation:
                self._initialiser_dimensions_structured_table()

        def change_references_tableauStructure_par_deltaNbLignes(self, deltaLignes: int, avec_initialiation: bool = True) -> None:
            """
            On rajoute deltaLignes à la référence actuelle du tableau (self._max_row + deltaLignes)
            """
            self.change_references_tableauStructure(self._min_row, self._min_col, self._max_row + deltaLignes, self._max_col, avec_initialiation)

        def change_references_tableauStructure_par_nbLignes(self, nbLignes: int, avec_initialiation: bool = True) -> None:
            """
            Le tableau fera nbLignes en tout (self._min_row + nbLignes)
            """
            self.change_references_tableauStructure(self._min_row, self._min_col, self._min_row + nbLignes, self._max_col, avec_initialiation)

        # === Méthodes publiques : divers ===
        def generer_dictionnaire_depuis_excel(self) -> None:
            """
            Récupère les noms des colonnes et prépare un dictionnaire avec ces noms de colonne en clé pour faciliter une déclaration.
            Imprime la structure Python pour inspection.
            """
            noms_colonnes = list(self._df.columns)
            print("mon_dictionnaire = {")
            for i, cle in enumerate(noms_colonnes):
                virgule = "," if i < len(noms_colonnes) - 1 else ""
                print(f"    {json.dumps(cle)}: \"\"{virgule}")
            print("}")

        # === Méthodes statiques / utilitaires ===
        @staticmethod
        def extraire_plage_str(cle):
            if hasattr(cle, "sqref"):
                return str(cle.sqref)
            elif hasattr(cle, "coord"):
                return str(cle.coord)
            else:
                return str(cle)

        @staticmethod
        def plages_chevauchent(plage1, plage2):
            min_col1, min_row1, max_col1, max_row1 = range_boundaries(FichierExcel._TableauExcel.extraire_plage_str(plage1))
            min_col2, min_row2, max_col2, max_row2 = range_boundaries(FichierExcel._TableauExcel.extraire_plage_str(plage2))
            return not (max_col1 < min_col2 or max_col2 < min_col1 or max_row1 < min_row2 or max_row2 < min_row1)

        # === Propriétés (accès contrôlé) ===
        @property
        def nom_onglet(self):
            return self._nom_onglet

        @property
        def nom_tableau(self):
            return self._nom_tableau

        @property
        def df(self):
            return self._df

        @property
        def nbLignes_avantET(self):
            return self._nbLignes_avantET

        @property
        def ws(self): # S'appelait feuille
            return self._parent._wb[self._nom_onglet] if (self._parent and self._parent._wb and self._nom_onglet) else None

        @property
        def ref_tableau(self):
            return self._ref_tableau

        @property
        def min_row(self):
            return self._min_row

        @property
        def max_row(self):
            return self._max_row

        @property
        def min_col(self):
            return self._min_col

        @property
        def max_col(self):
            return self._max_col

        # === Affichage ===
        def __str__(self) -> str:
            if self._ref_tableau is None:
                description = (
                    f"TableauExcel sans référence de tableau : c'est un répertoire IRIS export par défaut\n"
                    f"  Onglet : {self._nom_onglet}\n"
                    f"  Tableau : {self._nom_tableau or 'non structuré'}\n"
                    f"  Lignes avant en-tête : {self._nbLignes_avantET}"
                )
            else:
                description = (
                    f"TableauExcel\n"
                    f"  Onglet : {self._nom_onglet}\n"
                    f"  Tableau : {self._nom_tableau or 'non structuré'}\n"
                    f"  Réf : {self._ref_tableau or 'non définie'}\n"
                    f"  Lignes avant en-tête : {self._nbLignes_avantET}\n"
                    f"  Dimensions : lignes {self._min_row}-{self._max_row}, colonnes {self._min_col}-{self._max_col}"
                )

            if self._df is not None:
                buffer = StringIO()
                stdout_original = sys.stdout
                try:
                    sys.stdout = buffer
                    print(self._df)
                finally:
                    sys.stdout = stdout_original
                contenu_df = buffer.getvalue()
                return f"{description}\n\n{contenu_df.rstrip()}"
            else:
                return description


    


    # ==================================
    # === Retour classe FichierExcel ===
    # ==================================


    # Attributs "privés" / semi-privés
    _repertoire: Optional[str] = None
    _chemin_fichier: Optional[str] = None
    _tableaux: Dict[str, FichierExcel._TableauExcel] = {}

    # openpyxl
    _wb: Optional[Workbook] = None

    # xlwings
    _xw_app: Optional[xw.App] = None
    _xw_wb: Optional[xw.Book] = None
    _xw_ws: Optional[xw.Sheet] = None


    # === Constructeur de FichierExcel ===
    def __init__(self, chemin_fichier: Optional[str] = None) -> None:
        self._chemin_fichier = chemin_fichier
        self._tableaux = {}
        if chemin_fichier:
            self._repertoire = os.path.dirname(chemin_fichier)

    # === Constructeurs alternatifs de FichierExcel ===
    @classmethod
    def depuis_repertoire(cls, repertoire: str) -> FichierExcel:
        instance = cls()
        instance._repertoire = repertoire
        return instance

    @classmethod
    def depuis_fichier(cls,
                      chemin_fichier: Optional[str] = None,
                      avec_ouverture_wb: bool = True,
                      charger_tableau: bool = True,
                      nom_onglet: Optional[str] = None,
                      nom_tableau: Optional[str] = None,
                      nbLignes_avantET: int = 0,
                      charger_df: bool = True) -> FichierExcel:
        """
        Fabrique un FichierExcel à partir d'un chemin. Par défaut ouvre le workbook et 
        charge les tableaux structurés (ou un seul tableau, si nom_onglet fourni).
            Si pas de nom_onglet : on charge tous les tableaux structurés
            Si nom_onglet : on charge un seul tableau (structuré ou non) qui est dans cet onglet (il faut alors un nom_onglet)
            Si pas de nom_tableau, alors on récupère le nom de l'onglet
        """
        # Reprise de comportement (si on ne demande pas l'ouverture (chargement wb), alors pas de raison de charger le tableau. Idem avec charger_tableau et charger_df)
        if not avec_ouverture_wb:
            charger_tableau = False
        if not charger_tableau:
            charger_df = False

        # Si aucun fichier d'entrée, alors l'utilisateur le pointe avec filedialog
        if not chemin_fichier:
            chemin_fichier = cls.choisirFichiers_filedialog()

        instance = cls.depuis_repertoire(os.path.dirname(chemin_fichier))
        instance._chemin_fichier = chemin_fichier

        if avec_ouverture_wb:
            instance.charger_wb()

        if charger_tableau:
            if nom_onglet:
                instance.charger_tableau(nom_onglet, nom_tableau, nbLignes_avantET, charger_df)
            else:
                instance.charger_tableaux(charger_df=charger_df)

        return instance

    @classmethod
    def depuis_modele(cls,
                      chemin_modele: Optional[str] = None,
                      chemin_fichier_sauv: Optional[str] = None,
                      avec_ouverture_wb: bool = True,
                      charger_tableau: bool = True,
                      nom_onglet: Optional[str] = None,
                      nom_tableau: Optional[str] = None,
                      nbLignes_avantET: int = 0,
                      charger_df: bool = True) -> FichierExcel:
        """
        Charge un modèle Excel (et soit tous ses tableaux structurés, soit un tableau non structuré dans un onglet à donner).
        Si chemin_fichier_sauv est fourni, copie physiquement le modèle vers ce chemin, puis ouvre la copie.
        """
        if chemin_modele is None:
            vlog.log_erreur("Le chemin du modèle doit être fourni")

        instance = cls.depuis_fichier(
            chemin_fichier=chemin_modele,
            avec_ouverture_wb=avec_ouverture_wb,
            charger_tableau=charger_tableau,
            nom_onglet=nom_onglet,
            nom_tableau=nom_tableau,
            nbLignes_avantET=nbLignes_avantET,
            charger_df=charger_df
        )

        if chemin_fichier_sauv:
            instance.save_copie_physique(chemin_fichier_sauv)

        return instance

    # === Méthodes utilitaires : workbook ===
    def charger_wb(self) -> None:
        if not self._chemin_fichier:
            vlog.log_erreur("Le chemin du fichier Excel n'est pas défini")
        self._wb = load_workbook(self._chemin_fichier)

    def recharger_workbook_openpyxl(self, recalculerDimensionsTableau: bool = True) -> None:
        """
        Recharge le fichier Excel avec openpyxl et met à jour toutes les références 
        aux objets Worksheet/Table/TableauExcel.
        """
        self.charger_wb()
        for tableau in self._tableaux.values():
            tableau.reinitialiser_workbook_openpyxl(recalculer_dimensions=recalculerDimensionsTableau)

    def actualiser_TCD(self, save: bool = True, quitter: bool = True) -> None:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Open(self._chemin_fichier)
        for sheet in wb.Sheets:
            for pivot in sheet.PivotTables():
                pivot.RefreshTable()
        if save:
            wb.Save()
        if quitter:
            wb.Close(SaveChanges=False)
            excel.Quit()

    def _sauver_fichier(self, chemin_cible: str, copier: bool = False) -> None:
        """
        Méthode interne factorisée pour sauvegarder le fichier Excel.
        
        Args:
            chemin_cible (str): Chemin de sortie (sauvegarde ou copie).
            copier (bool): 
                - False → enregistre avec openpyxl (save classique)
                - True → copie physique du fichier original
        """
        if not chemin_cible:
            vlog.log_erreur("Chemin cible non défini pour la sauvegarde")
            return

        os.makedirs(os.path.dirname(chemin_cible), exist_ok=True)

        if copier:
            if not self._chemin_fichier:
                vlog.log_erreur("Chemin fichier original non défini, impossible de copier")
                return
            try:
                shutil.copy2(self._chemin_fichier, chemin_cible)
                self._chemin_fichier = chemin_cible
                self._repertoire = os.path.dirname(chemin_cible)
                self.recharger_workbook_openpyxl()
            except Exception as e:
                vlog.log_erreur(f"Erreur lors de la copie du fichier vers {chemin_cible} : {e}")
        else:
            if not self._wb:
                vlog.log_erreur("Le workbook n'est pas chargé")
                return
            try:
                self._wb.save(chemin_cible)
            except Exception as e:
                vlog.log_erreur(f"Erreur lors de l'enregistrement du fichier {chemin_cible} : {e}")

    def save(self, nouveau_chemin_fichier: Optional[str] = None) -> None:
        """Sauvegarde le fichier avec openpyxl (écriture du workbook en mémoire)."""
        chemin = nouveau_chemin_fichier or self._chemin_fichier
        self._sauver_fichier(chemin, copier=False)

    def save_copie_physique(self, nouveau_chemin: str) -> None:
        """Crée une copie physique du fichier Excel sur disque et recharge le workbook."""
        self._sauver_fichier(nouveau_chemin, copier=True)

    def close(self) -> None:
        """Ferme le workbook openpyxl (facultatif)."""
        if self._wb:
            self._wb.close()
        else:
            vlog.log_erreur("Le workbook n'est pas chargé")


    # === Méthodes utilisataires xlwings ou com
    def open_xlwings(self, nom_onglet:Optional[str], visible:bool=False) -> None:
        """ Ouvrir l'App xlwings proprement """
        self._xw_app = xw.App(visible=visible)
        if visible:
            self._xw_app.display_alerts = True
            self._xw_app.screen_updating = True

        self._xw_wb = self._xw_app.books.open(self._chemin_fichier)

        if nom_onglet:
            self._xw_ws = self._xw_wb.sheets[nom_onglet]
            self._xw_ws.activate()

    def close_xlwings(self, sauver:bool=True) -> None:
        """Sauve (optionel) et ferme xlwings"""
        if sauver:
            self._xw_wb.save(self._chemin_fichier)
        self._xw_wb.close()
        self._xw_app.quit()
        self._xw_ws = None
        self._xw_wb = None
        self._xw_app = None


    # === Méthodes utilitaires : chargement / identification de tableaux ===
    def charger_tableaux(
        self,
        definitions: Optional[list[tuple[str, Optional[str], int]]] = None,
        charger_df: bool = True
    ) -> dict[str, FichierExcel._TableauExcel]:
        """
        Charge un ou plusieurs tableaux dans le fichier Excel.

        Args:
            definitions:
                - None → charge tous les tableaux structurés présents dans le fichier
                - liste de tuples (nom_onglet, nom_tableau, nbLignes_avantET)
            charger_df: si True, charge également les DataFrames associés

        Usage (exemples):

        >>> # Pour charger un seul tableau :
        >>> f.charger_tableaux([("Onglet1", "TableauA", 2)])

        >>> # Pour charger tous les tableaux structurés :
        >>> f.charger_tableaux()

        Returns:
            dict[str, TableauExcel] : dictionnaire {nom_tableau: instance TableauExcel}
        """
        if not self._wb:
            vlog.log_erreur("Workbook non chargé, impossible de récupérer les tableaux")
            return {}

        # Si aucune définition n’est fournie → on parcourt tout le classeur
        if definitions is None:
            definitions = []
            for nom_onglet in self._wb.sheetnames:
                for nom_tableau in self._wb[nom_onglet].tables:
                    definitions.append((nom_onglet, nom_tableau, 0))

        for nom_onglet, nom_tableau, nbLignes_avantET in definitions:
            try:
                t = self._TableauExcel.depuis_FichierExcel_et_nomOnglet(
                    self,
                    nom_onglet=nom_onglet,
                    nom_tableau=nom_tableau,
                    nbLignes_avantET=nbLignes_avantET,
                    charger_df=charger_df
                )
                self.ajouter_tableau(t)
            except Exception as e:
                vlog.log_erreur(
                    f"Échec du chargement du tableau {nom_tableau} "
                    f"(onglet {nom_onglet}) : {e}"
                )

        return self._tableaux
    
    def charger_tableau(
        self,
        nom_onglet: str,
        nom_tableau: Optional[str] = None,
        nbLignes_avantET: int = 0,
        charger_df: bool = True,
    ) -> None:
        """Alias de charger_tableaux pour un seul tableau"""
        self.charger_tableaux([(nom_onglet, nom_tableau, nbLignes_avantET)], charger_df)

    def ajouter_tableau(self, tableau: FichierExcel._TableauExcel) -> None:
        """
        Ajoute un objet _TableauExcel à la collection interne. (Méthode 'low-level')
        """
        if not tableau._nom_tableau:
            vlog.log_erreur("Le tableau doit avoir un nom pour être ajouté")
        self._tableaux[tableau._nom_tableau] = tableau

    def ajouter_tableau_normal(self, nom_onglet: str, nb_lignes_avant_et: int, nom_tableau: Optional[str] = None, charger_df: bool = True) -> FichierExcel._TableauExcel:
        """
        Crée et ajoute un tableau 'normal' (non structuré) basé sur un onglet.
        Si nom_tableau non donné, alors on lui donne le nom de l'onglet (on considère qu'il y a un seul tableau par onglet)
        Retourne l'objet _TableauExcel ajouté.

        Exemple:
        >>> fx = FichierExcel.depuis_fichier("chemin.xlsx")
        >>> tab = fx.ajouter_tableau_normal("Feuil1", nb_lignes_avant_et=2, nom_tableau="Sessions")
        """
        if not self._wb:
            self.charger_wb()
        
        # On crée une instance _TableauExcel sans table struct (nom_tableau optionnel)
        nom_tableau = nom_tableau or nom_onglet
        tableau = FichierExcel._TableauExcel(parent=self, ws=self._wb[nom_onglet], nom_onglet=nom_onglet, nom_tableau=nom_tableau, nbLignes_avantET=nb_lignes_avant_et)
        tableau._est_tableau_structure = False
        tableau._initialiser_dimensions_tableau_normal()
        if charger_df:
            tableau.charge_df()
        self.ajouter_tableau(tableau)
        return tableau

    # === Méthodes publiques : accès aux tableaux (API propre) ===
    def get_tableau(self, nom: str) -> Optional[FichierExcel._TableauExcel]:
        """
        Récupère un tableau par nom (ou None si absent).
        Usage recommandé : utiliser ce tableau retourné pour appeler ses méthodes de manipulation.
        """
        return self._tableaux.get(nom)

    @property
    def noms_tableaux(self) -> List[str]:
        """Retourne la liste des noms de tableaux disponibles dans ce fichier."""
        return list(self._tableaux.keys())

    @property
    def tableaux(self) -> Dict[str, FichierExcel._TableauExcel]:
        """
        Accès contrôlé à la collection de tableaux.
        (Garde un dictionnaire en lecture — éviter modifications directes)
        """
        return dict(self._tableaux)

    # === Méthodes utilitaires statiques ===
    @staticmethod
    def choisirFichiers_filedialog(initialdir: Optional[str] = None) -> str:
        """
        Lister/sélectionner les documents à concaténer
        """
        # Je ne peux pas faire appel à self._input.repertoire comme initial dir car j'ai un appel avant création de mon instance (i.e. : pas de self)
        chemin_fichier = filedialog.askopenfilename(
            title="Sélectionner le fichier à charger",
            filetype=[("Fichiers Excel", "*.xlsx")],
            initialdir=initialdir or os.getcwd())
        if not chemin_fichier:
            vlog.log_erreur("click sur cancel du filedialog → Pas de chemins de fichier")
        return chemin_fichier

    @staticmethod
    def definir_rangeExcel(min_row: int, min_col: int, max_row: int, max_col: int) -> str:
        """
        Retourne un range Excel "A1:B13" à partir de numéros (lignes / colonnes)
        """
        lettre_col_min = get_column_letter(min_col)
        lettre_col_max = get_column_letter(max_col)
        return f"{lettre_col_min}{min_row}:{lettre_col_max}{max_row}"

    # === Propriétés ===
    @property
    def chemin_fichier(self) -> Optional[str]:
        return self._chemin_fichier

    @chemin_fichier.setter
    def chemin_fichier(self, nouveau_chemin: Optional[str]) -> None:
        if nouveau_chemin == self._chemin_fichier:
            return  # pas besoin de recharger si même chemin

        self._chemin_fichier = nouveau_chemin
        self._repertoire = os.path.dirname(nouveau_chemin) if nouveau_chemin else None

        # Recharge automatiquement le fichier si il existe
        if nouveau_chemin and os.path.isfile(nouveau_chemin):
            try:
                self._wb = load_workbook(nouveau_chemin)
            except Exception as e:
                vlog.log_erreur(f"Erreur lors du chargement du fichier {nouveau_chemin} : {e}")
                self._wb = None
        else:
            self._wb = None

    @property
    def repertoire(self) -> Optional[str]:
        return os.path.dirname(self._chemin_fichier) if self._chemin_fichier else None

    @property
    def nom_fichier(self) -> Optional[str]:
        if self._chemin_fichier is None:
            return None
        return os.path.basename(self._chemin_fichier)

    @property
    def wb(self) -> Optional[Workbook]:
        return self._wb

    # === Affichage ===
    def __str__(self) -> str:
        header = (
            f"📁 Propriétés du fichier Excel\n"
            f"  Fichier : {self.nom_fichier or 'non défini'}\n"
            f"  Répertoire : {self._repertoire or 'non défini'}\n"
            f"  Nombre de tableaux : {len(self._tableaux)}\n"
        )

        if not self._tableaux:
            return header + "\n  Aucun tableau chargé."

        contenu = [header]
        for nom_tableau, tableau in self._tableaux.items():
            titre = f"{Fore.YELLOW + Style.BRIGHT}{nom_tableau}{Style.RESET_ALL}"
            meta = (
                f"  ─ Onglet : {tableau.nom_onglet}\n"
                f"  ─ Référence : {tableau.ref_tableau or 'non définie'}\n"
                f"  ─ Lignes avant en-tête : {tableau.nbLignes_avantET}\n"
                f"  ─ Dimensions : lignes {tableau.min_row}-{tableau.max_row}, colonnes {tableau.min_col}-{tableau.max_col}"
            )
            if tableau._df is not None:
                buffer = StringIO()
                stdout_original = sys.stdout
                try:
                    sys.stdout = buffer
                    print(tableau._df)
                finally:
                    sys.stdout = stdout_original
                contenu_df = buffer.getvalue().rstrip()
            else:
                contenu_df = "  🔸 DataFrame non chargé"
            bloc = f"\n🟨 Tableau : {titre}\n{meta}\n\n{contenu_df}"
            contenu.append(bloc)

        return "\n\n".join(contenu)

# -------------------------
# 3) Notes / TODO
# -------------------------
# - Tests unitaires conseillés : coverer l'ouverture, ajout tableau normal, écriture df, copie formats (xlwings),
#   recharger_workbook_openpyxl et maj_references_misesEnFormeConditionnelles.
# - Certains appels à des attributs internes (ex: _ws._tables) dépendent de la version d'openpyxl ;
#   si tu as rencontré des différences, on peut ajouter des helpers robustes.
#
# Si tu veux, je peux maintenant :
#  - a) ajouter des tests unitaires minimalistes (pytest) couvrant l'API publique,
#  - d) rendre certains comportements configurables (ex: utiliser openpyxl vs xlwings).
#


class FichierWord:
    """
    Classe pour gérer un fichier Word (.docx) avec possibilité de lister les Content Controls.

    Attributs :
        _repertoire (Optional[str]) : Répertoire du fichier.
        _chemin_fichier (Optional[str]) : Chemin complet du fichier.
        _cc (Optional[Dict[str, str]]) : Dictionnaire {nom_content_control: valeur_texte}.

    Exemple d'utilisation :

    ```python
    # Créer une instance en ouvrant un fichier via dialogue
    fw = FichierWord.depuisFichier()

    # Afficher les Content Controls
    print(fw._cc)

    # Afficher résumé
    print(fw)

    # Sauvegarder sous un nouveau nom
    fw.save("nouveau_chemin.docx")

    # Fermer (utile si affichage Word ouvert)
    fw.close()
    ```
    """

    # === Variables de classe ===
    _repertoire: Optional[str] = None # TODO Dangereux car redondant avec _chemin_fichier
    _chemin_fichier: Optional[str] = None

    _cc: Optional[Dict[str, str]] = None #_cc[nomCC, ValeurCC] (content control de Word)
    _doc: Optional[Document] = None

    _com_word_app: Optional[win32com.client.CDispatch] = None  # COM Word Application (si utilisé)
    _com_word_doc: Optional[win32com.client.CDispatch] = None  # COM Word Document (si utilisé)


    def __init__(self,
        chemin_fichier: Optional[str] = None,
        repertoire: Optional[str] = None,
        content_controls: Optional[Dict[str, str]] = None,
        doc_obj: Optional[Document] = None,
        com_word_app: Optional[win32com.client.CDispatch] = None,
        com_word_doc: Optional[win32com.client.CDispatch] = None
    ) -> None:
        self._repertoire = repertoire
        self._chemin_fichier = chemin_fichier

        self._cc = content_controls #_cc[nomCC, ValeurCC]
        self._doc = doc_obj

        self._com_word_app = com_word_app  # COM Word Application (si utilisé)
        self._com_word_doc = com_word_doc  # COM Word Document (si utilisé)

    @classmethod
    def depuisFichier(cls,
        chemin_fichier: Optional[str] = None,
        charger_contentControl: bool = True,
        afficherWord: bool = False
    ) -> FichierWord:
        """
        Crée une instance de FichierWord depuis un fichier.

        Args:
            chemin_fichier (Optional[str]): Chemin vers le fichier Word. Si None, ouvre un dialogue fichier.
            charger_contentControl (bool): Si True, liste et charge les Content Controls dans _cc.
            afficherWord (bool): Si True, ouvre Word en interface COM visible.

        Returns:
            FichierWord: instance créée.

        Exemple :
        ```python
        fw = FichierWord.depuisFichier(afficherWord=True)
        print(fw._cc)
        ```
        """
        if chemin_fichier is None:
            root = Tk()
            root.withdraw()
            chemin_fichier = filedialog.askopenfilename(
                filetypes=[("Fichiers Word", "*.docx *.doc")],
                title="Sélectionner un fichier Word"
            )
            root.destroy()
            if not chemin_fichier:
                raise ValueError("Aucun fichier sélectionné")

        repertoire = os.path.dirname(chemin_fichier)

        content_controls = None
        doc_obj = None
        com_word_app = None
        com_word_doc = None

        if afficherWord:
            # Ouvrir Word via COM visible
            com_word_app = win32com.client.Dispatch("Word.Application")
            com_word_app.Visible = True
            com_word_doc = com_word_app.Documents.Open(chemin_fichier)
        else:
            # Charger docx via python-docx
            doc_obj = Document(chemin_fichier)

        if charger_contentControl:
            try:
                content_controls = cls._lister_content_controls(doc_obj)
            except Exception as e:
                content_controls = {}
                print(f"⚠️ Aucun Content Control trouvée ou erreur : {e}")

        return cls(chemin_fichier=chemin_fichier,
                   repertoire=repertoire,
                   content_controls=content_controls,
                   doc_obj=doc_obj,
                   com_word_app=com_word_app,
                   com_word_doc=com_word_doc)

    @staticmethod
    def _lister_content_controls(doc_obj: Document) -> Dict[str, str]:
        """
        Extrait les Content Controls d'un Document python-docx.

        Args:
            doc_obj (Document): Objet Document python-docx.

        Returns:
            Dict[str, str]: Dictionnaire {titre: texte} des Content Controls.

        Exemple :
        ```python
        cc = FichierWord._lister_content_controls(doc)
        print(cc)
        ```
        """
        if doc_obj is None:
            return {}

        NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'} # Cette URI sert uniquement d’identifiant unique, elle ne nécessite pas d’accès Internet pour fonctionner
        xml = doc_obj.part._element.xml
        root = etree.fromstring(xml.encode('utf-8'))

        sdt_elements = root.findall('.//w:sdt', namespaces=NS)
        dico_cc = {}

        for sdt in sdt_elements:
            texte = ''.join(sdt.xpath('.//w:t/text()', namespaces=NS)).strip()

            alias_elem = sdt.find('.//w:sdtPr/w:alias', namespaces=NS)
            tag_elem = sdt.find('.//w:sdtPr/w:tag', namespaces=NS)

            titre = None
            if alias_elem is not None:
                titre = alias_elem.attrib.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') # Cette URI sert uniquement d’identifiant unique, elle ne nécessite pas d’accès Internet pour fonctionner
            if titre is None and tag_elem is not None:
                titre = tag_elem.attrib.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') # Cette URI sert uniquement d’identifiant unique, elle ne nécessite pas d’accès Internet pour fonctionner

            if titre is None or titre == "":
                titre = f"(sans titre {len(dico_cc) + 1})"

            dico_cc[titre] = texte

        return dico_cc

    def save(self, nouveau_chemin_fichier: Optional[str] = None) -> None:
        """
        Sauvegarde le document.

        Args:
            nouveau_chemin_fichier (Optional[str]): Chemin de sauvegarde. Si None, sauvegarde au chemin initial.

        Exemple :
        ```python
        fw.save("nouveau_nom.docx")
        ```
        """
        chemin = nouveau_chemin_fichier if nouveau_chemin_fichier is not None else self._chemin_fichier
        if self._doc and chemin:
            self._doc.save(chemin)
        elif self._com_word_doc and chemin:
            self._com_word_doc.SaveAs(chemin)
        else:
            raise RuntimeError("Pas de document chargé ou chemin de sauvegarde invalide.")

    def close(self) -> None:
        """
        Ferme proprement le document et l'application Word (si ouverts via COM).

        Exemple :
        ```python
        fw.close()
        ```
        """
        if self._com_word_doc:
            self._com_word_doc.Close(False)
            self._com_word_doc = None
        if self._com_word_app:
            self._com_word_app.Quit()
            self._com_word_app = None
        self._doc = None

    @property
    def chemin_fichier(self) -> Optional[str]:
        """Chemin complet du fichier."""
        return self._chemin_fichier

    @chemin_fichier.setter
    def chemin_fichier(self, valeur: str) -> None:
        self._chemin_fichier = valeur

    @property
    def repertoire(self) -> Optional[str]:
        """Répertoire du fichier."""
        return self._repertoire

    @repertoire.setter
    def repertoire(self, valeur: str) -> None:
        self._repertoire = valeur

    @property
    def nom_fichier(self) -> Optional[str]:
        """Nom du fichier (sans le chemin)."""
        if self._chemin_fichier:
            return os.path.basename(self._chemin_fichier)
        return None

    @property
    def repertoire(self) -> Optional[str]:
        return self._repertoire

    @property
    def cc(self) -> Optional[Dict[str, str]]:
        return self._cc
    
    def __str__(self) -> str:
        """
        Résumé du fichier et ses Content Controls.

        Exemple :
        ```python
        print(fw)
        ```
        """
        nb_cc = len(self._cc) if self._cc else 0
        lignes = [
            f"📁 Propriétés du fichier Word",
            f"  Fichier : {self.nom_fichier or 'non défini'}",
            f"  Répertoire : {self._repertoire or 'non défini'}",
            f"  Nombre de Content Controls : {nb_cc}",
        ]
        if self._cc:
            for nom_cc, valeur_cc in self._cc.items():
                lignes.append(f"    • {nom_cc} : {valeur_cc}")
        return "\n".join(lignes)

class Mail:
    """
    Classe pour préparer et envoyer un mail via Outlook installé localement.

    Variables d'instance :
        _destinataires (Optional[List[str]]): Liste des destinataires principaux.
        _copies (Optional[List[str]]): Liste des destinataires en copie.
        _copies_cachees (Optional[List[str]]): Liste des destinataires en copie cachée.
        _sujet (Optional[str]): Sujet du mail.
        _corps (Optional[str]): Corps du mail (texte brut ou HTML).
        _pieces_jointes (Optional[List[str]]): Liste des chemins vers les fichiers à joindre.

    Exemple :
    --------
    >>> mail = Mail()
    >>> mail._destinataires = ["exemple@domaine.com"]
    >>> mail._sujet = "Test via Outlook"
    >>> mail.definir_corps_message("<p>Bonjour, ceci est un mail préparé via Python.</p>")
    >>> mail.selectionner_pj()
    >>> mail.creer_mail()  # Ouvre la fenêtre Outlook
    >>> mail.envoyer_mail()  # Envoie directement
    """

    # -------------------------------------------------------------------------
    # Variables d'instance
    # -------------------------------------------------------------------------
    _destinataires: Optional[List[str]]
    _copies: Optional[List[str]]
    _copies_cachees: Optional[List[str]]
    _sujet: Optional[str]
    _corps: Optional[str]
    _pieces_jointes: Optional[List[str]]

    # -------------------------------------------------------------------------
    # Initialisation
    # -------------------------------------------------------------------------
    def __init__(self) -> None:
        self._destinataires = None
        self._copies = None
        self._copies_cachees = None
        self._sujet = None
        self._corps = None
        self._pieces_jointes = None

    # -------------------------------------------------------------------------
    # Fonction interne factorisée pour créer un mail
    # -------------------------------------------------------------------------
    @classmethod
    def _creer_mail(
        cls,
        mail_obj: Optional[win32com.client.CDispatch] = None,
        destinataires: Optional[Union[str, List[str], pd.Series]] = None,
        copies: Optional[Union[str, List[str], pd.Series]] = None,
        copies_cachees: Optional[Union[str, List[str], pd.Series]] = None,
        sujet: Optional[str] = None,
        corps_html: Optional[str] = None,
        pieces_jointes: Optional[Union[str, List[str], pd.Series]] = None,
        envoyer_mail: bool = False
    ) -> None:
        """
        Crée ou complète un mail Outlook existant (ex: depuis modèle) ou un mail vierge.

        Args:
            mail_obj: objet mail Outlook existant (None pour un mail vierge)
            destinataires, copies, copies_cachees, sujet, corps_html, pieces_jointes
            envoyer_mail: si True, envoie directement le mail

        Exemple :
        --------
        >>> Mail._creer_mail(
        >>>     destinataires="user@example.com",
        >>>     copies=["copie@example.com"],
        >>>     copies_cachees=None,
        >>>     sujet="Sujet test",
        >>>     corps_html="<p>Bonjour</p>",
        >>>     pieces_jointes=["C:/fichier.pdf"],
        >>>     envoyer_mail=False
        >>> )
        """
        # Convertir tous les champs en listes
        to_list = convertir_en_liste(destinataires)
        cc_list = convertir_en_liste(copies)
        bcc_list = convertir_en_liste(copies_cachees)
        pj_list = convertir_en_liste(pieces_jointes)

        # Créer mail si pas fourni
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = mail_obj or outlook.CreateItem(0)

        # Destinataires, copies, copies cachées
        if to_list:
            mail.To = ";".join(to_list)
        if cc_list:
            mail.CC = ";".join(cc_list)
        if bcc_list:
            mail.BCC = ";".join(bcc_list)

        if sujet:
            mail.Subject = sujet

        # Affichage pour forcer la signature
        mail.Display()

        # Ajouter corps HTML
        signature = mail.HTMLBody
        mail.HTMLBody = (corps_html or "") + signature

        # Ajouter pièces jointes
        for pj in pj_list:
            mail.Attachments.Add(pj)

        # Envoyer ou laisser affiché
        if envoyer_mail:
            mail.Send()

        # Nettoyage COM
        if mail_obj is None:  # si on a créé l'objet ici
            del mail
            del outlook

    # -------------------------------------------------------------------------
    # Méthode publique pour créer un mail vierge
    # -------------------------------------------------------------------------
    @classmethod
    def creer_mail(
        cls,
        destinataires: Optional[Union[str, List[str], pd.Series]] = None,
        copies: Optional[Union[str, List[str], pd.Series]] = None,
        copies_cachees: Optional[Union[str, List[str], pd.Series]] = None,
        sujet: Optional[str] = None,
        corps_html: Optional[str] = None,
        pieces_jointes: Optional[Union[str, List[str], pd.Series]] = None,
        envoyer_mail: bool = False
    ) -> None:
        """
        Crée un mail vierge Outlook et l'affiche ou l'envoie.

        Exemple :
        --------
        >>> Mail.creer_mail(
        >>>     destinataires="user@example.com",
        >>>     copies=["cc@example.com"],
        >>>     sujet="Test",
        >>>     corps_html="<p>Bonjour</p>",
        >>>     pieces_jointes=["C:/fichier.pdf"],
        >>>     envoyer_mail=False
        >>> )
        """
        cls._creer_mail(
            mail_obj=None,
            destinataires=destinataires,
            copies=copies,
            copies_cachees=copies_cachees,
            sujet=sujet,
            corps_html=corps_html,
            pieces_jointes=pieces_jointes,
            envoyer_mail=envoyer_mail
        )

    # -------------------------------------------------------------------------
    # Méthode publique pour créer un mail depuis un modèle .msg
    # -------------------------------------------------------------------------
    @classmethod
    def depuis_modele(
        cls,
        chemin_modele: str,
        destinataires: Optional[Union[str, List[str], pd.Series]] = None,
        copies: Optional[Union[str, List[str], pd.Series]] = None,
        copies_cachees: Optional[Union[str, List[str], pd.Series]] = None,
        sujet: Optional[str] = None,
        corps_html: Optional[str] = None,
        pieces_jointes: Optional[Union[str, List[str], pd.Series]] = None,
        envoyer_mail: bool = False,
        remplaceBalises: Optional[List[List[str]]] = None,  #liste de couples [[texte_a_remplacer, texte_de_remplacement], ...]
    ) -> None:
        """
        Crée un mail Outlook à partir d'un modèle .msg et ajoute éventuellement
        destinataires, sujet, corps et pièces jointes.

        Exemple :
        --------
        >>> Mail.depuis_modele(
        >>>     chemin_modele="C:/Modeles/modele.msg",
        >>>     destinataires="user@example.com",
        >>>     copies=["cc@example.com"],
        >>>     sujet="Sujet test",
        >>>     corps_html="<p>Bonjour</p>",
        >>>     pieces_jointes=["C:/fichier.pdf"],
        >>>     adaptations=[["###statut###", "Vacataire"], ["###annee###", "2025"]],
        >>>     envoyer_mail=False
        >>> )
        """
        if not os.path.isfile(chemin_modele):
            raise FileNotFoundError(f"Fichier modèle non trouvé : {chemin_modele}")

        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItemFromTemplate(os.path.abspath(chemin_modele))

        # ✅ Si des adaptations sont fournies, on les applique dans le corps du mail
        if remplaceBalises:
            corps = mail.HTMLBody
            for ancien, nouveau in remplaceBalises:
                corps = corps.replace(ancien, nouveau)
            mail.HTMLBody = corps

        # Utilise la méthode factorisée
        cls._creer_mail(
            mail_obj=mail,
            destinataires=destinataires,
            copies=copies,
            copies_cachees=copies_cachees,
            sujet=sujet,
            corps_html=corps_html,
            pieces_jointes=pieces_jointes,
            envoyer_mail=envoyer_mail
        )

        # Nettoyage COM
        del mail
        del outlook
    
    # -------------------------------------------------------------------------
    # Méthode pour sélectionner des pièces jointes via un dialogue
    # -------------------------------------------------------------------------
    def selectionner_pj(self) -> None:
        """
        Ouvre une boîte de dialogue pour sélectionner une ou plusieurs pièces jointes.

        Exemple :
        ```python
        mail = Mail()
        mail.selectionner_pj()
        print(mail._pieces_jointes)
        ```
        """
        root = Tk()
        root.withdraw()  # Cacher la fenêtre principale
        fichiers = filedialog.askopenfilenames(
            title="Sélectionner les pièces jointes",
            filetypes=[("Tous les fichiers", "*.*")]
        )
        root.destroy()
        if fichiers:
            self._pieces_jointes = list(fichiers)
        else:
            self._pieces_jointes = []


    def __str__(self) -> str:
        """
        Retourne une description lisible du mail préparé.

        Exemple :
        ```python
        print(mail)
        ```
        """
        lignes = [
            "✉️ Propriétés du mail",
            f"  Destinataires : {', '.join(self._destinataires) if self._destinataires else 'non définis'}",
            f"  Copies : {', '.join(self._copies) if self._copies else 'non définies'}",
            f"  Copies cachées : {', '.join(self._copies_cachees) if self._copies_cachees else 'non définies'}",
            f"  Sujet : {self._sujet or 'non défini'}",
            f"  Corps : {self._corps[:50] + '...' if self._corps and len(self._corps) > 50 else self._corps or 'vide'}",
            f"  Nombre de pièces jointes : {len(self._pieces_jointes) if self._pieces_jointes else 0}",
        ]
        if self._pieces_jointes:
            for pj in self._pieces_jointes:
                lignes.append(f"    • {pj}")
        return "\n".join(lignes)

class RDV_Outlook:
    """
    Classe représentant un rendez-vous Outlook.

    Cette classe permet de créer, manipuler et importer des rendez-vous Outlook
    directement depuis Python en utilisant COM.

    Fonctionnalités principales :
      - Création d’un RDV classique (méthode `creer_rdv_outlook`)
      - Création depuis un modèle Outlook `.oft` (méthode `depuis_modele_oft`)
      - Création rapide depuis texte brut (méthode `depuis_texte`)
      - Création rapide depuis HTML (méthode `depuis_html`)
      - Import d’un rendez-vous déjà existant dans Outlook (méthode `depuis_rdv_outlook`)

    Gestion Skype :
      Si l’argument `rdv_skype=True` est passé, l’objet Outlook sera configuré
      comme une réunion Skype (avec génération automatique du lien et adaptation
      du champ lieu).

    Exemple minimal :
    ----------------
    >>> rdv = RDV_Outlook(
    ...     sujet="Réunion projet",
    ...     lieu="Salle A",
    ...     date_debut=datetime.datetime(2025, 9, 25, 14, 0),
    ...     date_fin=datetime.datetime(2025, 9, 25, 15, 0),
    ...     participants_obligatoires=["alice@example.com"],
    ...     rdv_skype=True
    ... )
    >>> rdv.creer_rdv_outlook(envoyer=False)  # Affiche le rendez-vous dans Outlook
    """
    DEFAULT_TZ = ZoneInfo("Europe/Paris")


    def __init__(
        self,
        sujet: str,
        lieu: Optional[str],
        date_debut: datetime.datetime,
        date_fin: datetime.datetime,
        evenement_journee_entiere: bool = False,
        categorie: Optional[str] = None,
        disponibilite: Optional[str] = "Occupé",  # Disponibilités possibles : "Libre"=0, "Provisoire"=1, "Occupé"=2, "Absent"=3, "Travaille en dehors du bureau"=4
        rappel_active: bool = False,
        rappel_minutes: Optional[int] = None,
        importance: Optional[int] = None,  # 0 = faible ; 1 = normal, 2 = haute
        sensibilite: Optional[int] = None,
        reponse_demande: bool = True,
        texte_rdv: Optional[str] = None,
        html_rdv: Optional[str] = None,
        participants_obligatoires: Optional[Union[List[str], str, pd.Series]] = None,
        participants_facultatifs: Optional[Union[List[str], str, pd.Series]] = None,
        pieces_jointes: Optional[Union[List[str], str, pd.Series]] = None,
        rdv_skype: bool = False
    ) -> None:

        # Champs principaux
        self._sujet = sujet
        self._lieu = lieu
        self._date_debut = self._normaliser_datetime(date_debut)
        self._date_fin = self._normaliser_datetime(date_fin)
        self._evenement_journee_entiere = evenement_journee_entiere
        self._categorie = categorie
        self._disponibilite = disponibilite
        self._rappel_active = rappel_active
        self._rappel_minutes = rappel_minutes
        self._importance = importance
        self._sensibilite = sensibilite
        self._reponse_demande = reponse_demande
        self._texte_rdv = texte_rdv
        self._html_rdv = html_rdv

        # Participants (conversion string / Series / list)
        self._participants_obligatoires = convertir_en_liste(participants_obligatoires)
        self._participants_facultatifs = convertir_en_liste(participants_facultatifs)

        # Pièces jointes
        self._pieces_jointes = convertir_en_liste(pieces_jointes)

        # Option Skype
        self._rdv_skype = rdv_skype

    # -------------------------------------------------------------------------
    # Méthode privée : appliquer les données Python à un AppointmentItem Outlook
    # -------------------------------------------------------------------------
    def _remplir_rdv_outlook(self, rdv_outlook):
        """Applique les propriétés de l'objet Python sur un AppointmentItem Outlook."""

        # Sujet et lieu
        rdv_outlook.Subject = self._sujet
        rdv_outlook.Location = self._lieu or ""

        # Dates
        if self._evenement_journee_entiere:
            rdv_outlook.AllDayEvent = True
            rdv_outlook.Start = self._date_debut
            rdv_outlook.End = (self._date_fin + datetime.timedelta(days=1)) if self._date_fin else self._date_debut
        else:
            rdv_outlook.Start = self._date_debut
            rdv_outlook.End = self._date_fin

        # Rappel
        if self._rappel_active and self._rappel_minutes:
            rdv_outlook.ReminderSet = True
            rdv_outlook.ReminderMinutesBeforeStart = self._rappel_minutes
        else:
            rdv_outlook.ReminderSet = False

        # Catégorie
        if self._categorie:
            rdv_outlook.Categories = self._categorie

        # Disponibilité (BusyStatus)
        disponibilites_map = {
            "Libre": 0,
            "Provisoire": 1,
            "Occupé": 2,
            "Absent": 3,
            "Travaille en dehors du bureau": 4,
        }
        if self._disponibilite:
            rdv_outlook.BusyStatus = disponibilites_map.get(self._disponibilite, 2)

        # Participants
        if self._participants_obligatoires or self._participants_facultatifs:
            rdv_outlook.MeetingStatus = 1  # réunion
        for email in self._participants_obligatoires:
            rdv_outlook.Recipients.Add(email).Type = 1
        for email in self._participants_facultatifs:
            rdv_outlook.Recipients.Add(email).Type = 2
        rdv_outlook.Recipients.ResolveAll()

        # Pièces jointes
        for pj in self._pieces_jointes:
            if os.path.exists(pj):
                rdv_outlook.Attachments.Add(os.path.abspath(pj))

        # Réunion Skype
        if self._rdv_skype:
            rdv_outlook.MeetingStatus = 1  # obligatoire pour activer Skype
            rdv_outlook.IsOnlineMeeting = True
            rdv_outlook.OnlineMeetingProvider = 1  # 1 = Skype
            if not self._lieu:
                rdv_outlook.Location = "Réunion Skype"

        # Corps
        # Ouvre le rendez-vous pour forcer Outlook à injecter la signature
        #rdv_outlook.Display()

        if self._texte_rdv:
            rdv_outlook.Body = self._texte_rdv
        elif self._html_rdv:
            # Fallback : convertir le HTML en texte brut utilisable
            rdv_outlook.Body = html_vers_texte(self._html_rdv)

            
    # -------------------------------------------------------------------------
    # Méthode privée : finalisation (envoi ou affichage + nettoyage)
    # -------------------------------------------------------------------------
    @staticmethod
    def _finaliser_rdv_outlook(rdv, outlook, envoyer: bool) -> None:
        """Affiche ou envoie le rendez-vous, puis libère les ressources COM."""
        if envoyer:
            rdv.Send()
        else:
            rdv.Display()
        del rdv
        del outlook

    # -------------------------------------------------------------------------
    # Création classique
    # -------------------------------------------------------------------------
    def creer_rdv_outlook(self, envoyer: bool = False) -> None:
        """
        Crée un rendez-vous Outlook à partir des attributs Python.

        Exemple :
        ---------
        >>> rdv = RDV_Outlook("Test", "Salle X", dt(2025, 1, 1, 10), dt(2025, 1, 1, 11))
        >>> rdv.creer_rdv_outlook(envoyer=False)  # affiche le rendez-vous
        """
        outlook = win32com.client.Dispatch("Outlook.Application")
        rdv_outlook = outlook.CreateItem(1)  # olAppointmentItem
        self._remplir_rdv_outlook(rdv_outlook)
        self._finaliser_rdv_outlook(rdv_outlook, outlook, envoyer)

    # -------------------------------------------------------------------------
    # Création depuis modèle .oft
    # -------------------------------------------------------------------------
    @classmethod
    def depuis_modele_oft(
        cls,
        chemin_modele: str,
        sujet: Optional[str] = None,
        lieu: Optional[str] = None,
        participants_obligatoires: Optional[Union[List[str], str, pd.Series]] = None,
        participants_facultatifs: Optional[Union[List[str], str, pd.Series]] = None,
        date_debut: Optional[datetime] = None,
        duree: Optional[timedelta] = None,
        categorie: Optional[str] = None,
        rdv_skype: bool = False,
        pieces_jointes: Optional[Union[List[str], str, pd.Series]] = None,
        envoyer: bool = False
    ) -> RDV_Outlook:
        """
        Crée un rendez-vous Outlook à partir d'un modèle `.oft`.

        Exemple :
        ---------
        >>> rdv = RDV_Outlook.depuis_modele_oft("modele.oft", sujet="Réunion", rdv_skype=True)
        """
        outlook = win32com.client.Dispatch("Outlook.Application")
        rdv = outlook.CreateItemFromTemplate(os.path.abspath(chemin_modele))

        instance = cls(
            sujet=sujet or rdv.Subject,
            lieu=lieu or rdv.Location,
            date_debut=date_debut or rdv.Start,
            date_fin=(date_debut + duree) if (date_debut and duree) else rdv.End,
            evenement_journee_entiere=rdv.AllDayEvent,
            categorie=categorie or rdv.Categories,
            disponibilite=rdv.BusyStatus,
            rappel_active=rdv.ReminderSet,
            rappel_minutes=rdv.ReminderMinutesBeforeStart if rdv.ReminderSet else None,
            importance=rdv.Importance,
            sensibilite=rdv.Sensitivity,
            reponse_demande=rdv.ResponseRequested,
            participants_obligatoires=convertir_en_liste(participants_obligatoires),
            participants_facultatifs=convertir_en_liste(participants_facultatifs),
            pieces_jointes = convertir_en_liste(pieces_jointes),
            rdv_skype=rdv_skype
        )

        instance._remplir_rdv_outlook(rdv)
        cls._finaliser_rdv_outlook(rdv, outlook, envoyer)
        return instance

    # -------------------------------------------------------------------------
    # Création directe depuis HTML (avec tous les autres paramètres optionnels)
    # -------------------------------------------------------------------------
    @classmethod
    def depuis_html(
        cls,
        sujet: str,
        html: str,
        date_debut: datetime,
        date_fin: datetime,
        *,
        lieu: Optional[str] = None,
        evenement_journee_entiere: bool = False,
        categorie: Optional[str] = None,
        disponibilite: Optional[str] = "Occupé",
        rappel_active: bool = False,
        rappel_minutes: Optional[int] = None,
        importance: Optional[int] = None,
        sensibilite: Optional[int] = None,
        reponse_demande: bool = True,
        participants_obligatoires: Optional[Union[List[str], str, pd.Series]] = None,
        participants_facultatifs: Optional[Union[List[str], str, pd.Series]] = None,
        pieces_jointes: Optional[Union[List[str], str, pd.Series]] = None,
        rdv_skype: bool = False,
        envoyer: bool = False
    ) -> RDV_Outlook:
        """
        Crée un rendez-vous directement depuis un contenu HTML.

        Tous les autres arguments de l'init peuvent être passés pour personnaliser le RDV.

        Exemple :
        ---------
        >>> RDV_Outlook.depuis_html(
        ...     sujet="Réunion",
        ...     html="<b>Bonjour</b>",
        ...     date_debut=dt.now(),
        ...     date_fin=dt.now() + timedelta(hours=1),
        ...     participants_obligatoires=["alice@example.com"],
        ...     rdv_skype=True
        ... )
        """
        instance = cls(
            sujet=sujet,
            lieu=lieu,
            date_debut=date_debut,
            date_fin=date_fin,
            evenement_journee_entiere=evenement_journee_entiere,
            categorie=categorie,
            disponibilite=disponibilite,
            rappel_active=rappel_active,
            rappel_minutes=rappel_minutes,
            importance=importance,
            sensibilite=sensibilite,
            reponse_demande=reponse_demande,
            html_rdv=html,
            participants_obligatoires=convertir_en_liste(participants_obligatoires),
            participants_facultatifs=convertir_en_liste(participants_facultatifs),
            pieces_jointes=convertir_en_liste(pieces_jointes),
            rdv_skype=rdv_skype
        )
        instance.creer_rdv_outlook(envoyer=envoyer)
        return instance

    # -------------------------------------------------------------------------
    # Création directe depuis texte brut (avec tous les autres paramètres optionnels)
    # -------------------------------------------------------------------------
    @classmethod
    def depuis_texte(
        cls,
        sujet: str,
        texte: str,
        date_debut: datetime,
        date_fin: datetime,
        *,
        lieu: Optional[str] = None,
        evenement_journee_entiere: bool = False,
        categorie: Optional[str] = None,
        disponibilite: Optional[str] = "Occupé",
        rappel_active: bool = False,
        rappel_minutes: Optional[int] = None,
        importance: Optional[int] = None,
        sensibilite: Optional[int] = None,
        reponse_demande: bool = True,
        participants_obligatoires: Optional[Union[List[str], str, pd.Series]] = None,
        participants_facultatifs: Optional[Union[List[str], str, pd.Series]] = None,
        pieces_jointes: Optional[Union[List[str], str, pd.Series]] = None,
        rdv_skype: bool = False,
        envoyer: bool = False
    ) -> RDV_Outlook:
        """
        Crée un rendez-vous directement depuis un contenu texte.

        Tous les autres arguments de l'init peuvent être passés pour personnaliser le RDV.

        Exemple :
        ---------
        >>> RDV_Outlook.depuis_texte(
        ...     sujet="Briefing",
        ...     texte="Rappel: réunion importante",
        ...     date_debut=dt.now(),
        ...     date_fin=dt.now() + timedelta(hours=2),
        ...     participants_obligatoires=["bob@example.com"],
        ...     rdv_skype=True
        ... )
        """
        instance = cls(
            sujet=sujet,
            lieu=lieu,
            date_debut=date_debut,
            date_fin=date_fin,
            evenement_journee_entiere=evenement_journee_entiere,
            categorie=categorie,
            disponibilite=disponibilite,
            rappel_active=rappel_active,
            rappel_minutes=rappel_minutes,
            importance=importance,
            sensibilite=sensibilite,
            reponse_demande=reponse_demande,
            texte_rdv=texte,
            participants_obligatoires=convertir_en_liste(participants_obligatoires),
            participants_facultatifs=convertir_en_liste(participants_facultatifs),
            pieces_jointes=convertir_en_liste(pieces_jointes),
            rdv_skype=rdv_skype
        )
        instance.creer_rdv_outlook(envoyer=envoyer)
        return instance

    # -------------------------------------------------------------------------
    # Import d’un rendez-vous Outlook existant
    # -------------------------------------------------------------------------
    @classmethod
    def depuis_rdv_outlook(cls, rdv) -> RDV_Outlook:
        """
        Construit un objet Python RDV_Outlook à partir d'un AppointmentItem Outlook déjà existant.

        Exemple :
        ---------
        >>> outlook = win32com.client.Dispatch("Outlook.Application")
        >>> rdv_outlook = outlook.CreateItem(1)
        >>> rdv_obj = RDV_Outlook.depuis_rdv_outlook(rdv_outlook)
        """
        return cls(
            sujet=rdv.Subject,
            lieu=rdv.Location,
            date_debut=rdv.Start,
            date_fin=rdv.End,
            evenement_journee_entiere=rdv.AllDayEvent,
            categorie=rdv.Categories,
            disponibilite=rdv.BusyStatus,
            rappel_active=rdv.ReminderSet,
            rappel_minutes=rdv.ReminderMinutesBeforeStart if rdv.ReminderSet else None,
            importance=rdv.Importance,
            sensibilite=rdv.Sensitivity,
            reponse_demande=rdv.ResponseRequested,
            texte_rdv=rdv.Body if hasattr(rdv, "Body") else None,
            html_rdv=getattr(rdv, "HTMLBody", None),
            participants_obligatoires=[],
            participants_facultatifs=[],
            rdv_skype=getattr(rdv, "IsOnlineMeeting", False),
        )




    def convertir_date_heure_en_datetime(
        date_: datetime.date,
        heure_: Optional[datetime.time] = None,
        est_journee_entiere: bool = False
    ) -> datetime.datetime:
        """
        Convertit une date et une heure (optionnelle) en un objet datetime.

        Args:
            date_ (datetime.date) : La date.
            heure_ (Optional[datetime.time], optional) : L'heure (peut être None).
            est_journee_entiere (bool, optional) : True si événement journée entière (heure = minuit). Défaut False.

        Returns:
            datetime.datetime : Objet datetime combiné.

        Exemple:
            >>> convertir_date_heure_en_datetime(datetime.date(2025, 9, 20), datetime.time(14, 30))
            datetime.datetime(2025, 9, 20, 14, 30)

            >>> convertir_date_heure_en_datetime(datetime.date(2025, 9, 20), est_journee_entiere=True)
            datetime.datetime(2025, 9, 20, 0, 0)
        """
        if est_journee_entiere:
            # Pour journée entière, on met l'heure à minuit
            return datetime.datetime.combine(date_, datetime.time(0, 0))

        if heure_ is None:
            # Si pas d'heure fournie et pas journée entière, on met minuit par défaut
            return datetime.datetime.combine(date_, datetime.time(0, 0))

        # Sinon, combine date et heure normalement
        return datetime.datetime.combine(date_, heure_)



    @classmethod
    def depuis_dataframe(cls, df: pd.DataFrame) -> List[RDV_Outlook]:
        """
        Crée une liste d'objets RDV_Outlook à partir d'un DataFrame pandas.
        Ne retient que les lignes où 'Fait' est vide et 'À planifier' = "Oui" si cette colonne existe.

        Args:
            df (pd.DataFrame): DataFrame contenant les données des rendez-vous.

        Returns:
            List[RDV_Outlook]: Liste des rendez-vous à créer.

        Exemple:
            >>> rdvs = RDV_Outlook.depuis_dataframe(mon_dataframe)
        """
        rdvs: List[RDV_Outlook] = []

        colonne_a_planifier_existante = "À planifier" in df.columns

        for _, ligne in df.iterrows():
            fait = str(ligne.get("Fait", "")).strip()
            a_planifier = str(ligne.get("À planifier", "")).strip() if colonne_a_planifier_existante else None

            if fait == "" and (not colonne_a_planifier_existante or a_planifier.lower() == "oui"):

                sujet = str(ligne.get("Intitulé du RDV", ""))
                lieu = ligne.get("Lieu", None)

                date_debut_raw = ligne.get("Date début")
                date_fin_raw = ligne.get("Date fin")
                heure_debut_raw = ligne.get("Heure début")
                duree_raw = ligne.get("Durée [min]", None)
                rappel_raw = ligne.get("Rappel", None)

                # TODO si pb heure UTC : appliquer self._normaliser_datetime(date_fin)
                date_debut = cls._convertir_en_date(date_debut_raw)
                date_fin = cls._convertir_en_date(date_fin_raw) if pd.notna(date_fin_raw) else None
                heure_debut = cls._convertir_en_heure(heure_debut_raw) if pd.notna(heure_debut_raw) else None

                duree_minutes = int(duree_raw) if pd.notna(duree_raw) else None
                rappel_minutes = int(rappel_raw) if pd.notna(rappel_raw) else 0

                obligatoires = cls._separer_emails(ligne.get("Obligatoires"))
                facultatifs = cls._separer_emails(ligne.get("Facultatifs"))
                disponibilite = str(ligne.get("Disponibilité", "Occupé"))
                categorie = ligne.get("Catégorie de calendrier", None)
                texte = ligne.get("Texte RDV", None)

                evenement_journee_entiere = date_fin is not None

                rdv = cls(
                    sujet=sujet,
                    lieu=lieu,
                    date_debut=date_debut,
                    date_fin=date_fin,
                    heure_debut=heure_debut,
                    duree_minutes=duree_minutes,
                    rappel_minutes=rappel_minutes,
                    participants_obligatoires=obligatoires,
                    participants_facultatifs=facultatifs,
                    disponibilite=disponibilite,
                    categorie=categorie,
                    texte_rdv=texte,
                    evenement_journee_entiere=evenement_journee_entiere,
                )
                rdvs.append(rdv)

        return rdvs

    @staticmethod
    def extraire_html_msg(chemin_msg: str) -> str:
        msg = extract_msg.Message(chemin_msg)
        msg.load()  # charge le fichier .msg
        html_body = msg.htmlBody  # Contenu HTML complet, ou None si absent
        return html_body or ""
    
    @staticmethod
    def calculer_rappel_minutes(
        jours: int = 0,
        heures: int = 0,
        minutes: int = 0
    ) -> int:
        """
        Convertit un délai en minutes pour le rappel Outlook.

        Args:
            jours (int, optional) : Nombre de jours avant le RDV. Défaut 0.
            heures (int, optional) : Nombre d’heures avant le RDV. Défaut 0.
            minutes (int, optional) : Nombre de minutes avant le RDV. Défaut 0.

        Returns:
            int : Nombre total de minutes à passer à ReminderMinutesBeforeStart.

        Exemple :
        --------
        >>> RDV_Outlook.calculer_rappel_minutes(jours=7)
        10080

        >>> RDV_Outlook.calculer_rappel_minutes(jours=1, heures=2, minutes=30)
        1590
        """
        total_minutes = jours * 24 * 60 + heures * 60 + minutes
        return total_minutes

    @staticmethod
    def _convertir_en_date(valeur) -> Optional[datetime.date]:
        """
        Convertit une valeur en datetime.date si possible.

        Args:
            valeur: valeur brute potentiellement convertible.

        Returns:
            Optional[datetime.date]: date convertie ou None si échec.
        """
        if pd.isna(valeur):
            return None
        if isinstance(valeur, datetime.date):
            return valeur
        if isinstance(valeur, datetime.datetime):
            return valeur.date()
        try:
            dt = pd.to_datetime(valeur)
            return dt.date()
        except Exception:
            return None

    @staticmethod
    def _convertir_en_heure(valeur) -> Optional[datetime.time]:
        """
        Convertit une valeur en datetime.time si possible.

        Args:
            valeur: valeur brute potentiellement convertible.

        Returns:
            Optional[datetime.time]: heure convertie ou None si échec.
        """
        if pd.isna(valeur):
            return None
        if isinstance(valeur, datetime.time):
            return valeur
        if isinstance(valeur, datetime.datetime):
            return valeur.time()
        try:
            dt = pd.to_datetime(valeur)
            return dt.time()
        except Exception:
            return None

    @staticmethod
    def _normaliser_datetime(dt: datetime, tz: ZoneInfo = None) -> datetime:
        """
        Normalise un datetime pour Outlook :
        - Si naïf → ajoute le fuseau (Europe/Paris par défaut).
        - Si déjà tz-aware → le conserve tel quel.

        Exemple :
            >>> RDV_Outlook._normaliser_datetime(datetime(2025, 2, 5, 8, 0))
            datetime(2025, 2, 5, 8, 0, tzinfo=ZoneInfo("Europe/Paris"))
        """
        if tz is None:
            tz = RDV_Outlook.DEFAULT_TZ
        if dt.tzinfo is None:
            return dt.replace(tzinfo=tz)
        return dt

    @staticmethod
    def _separer_emails(valeur: Optional[Union[str, List[str]]]) -> List[str]:
        """
        Transforme une chaîne d'emails séparés par ';' ou ',' en liste d'emails.
        Si entrée déjà liste, la retourne telle quelle.

        Args:
            valeur (Optional[Union[str, List[str]]]): chaîne ou liste d'emails.

        Returns:
            List[str]: liste d'emails nettoyés.
        """
        if not valeur:
            return []
        if isinstance(valeur, list):
            return valeur
        emails = [email.strip() for email in str(valeur).replace(",", ";").split(";") if email.strip()]
        return emails



    @staticmethod
    def get_lundi_depuis_num_semaine(numero_semaine: int) -> date:
        """Retourne le lundi de la semaine ISO donnée.
        Utilise l'année en cours, sauf si cette semaine est déjà passée, alors prend l'année suivante.
        """
        today = date.today()
        annee = today.year

        # Calcule le lundi de cette semaine dans l'année actuelle
        lundi = datetime.fromisocalendar(annee, numero_semaine, 1).date()

        # Si cette date est déjà passée, on utilise l'année suivante
        if lundi < today:
            annee += 1
            lundi = datetime.fromisocalendar(annee, numero_semaine, 1).date()

        return lundi

    @property
    def heure_debut(self) -> time:
        """
        Retourne l'heure de début (heure + minute) de _date_debut.
        """
        return self._date_debut.time()

    @heure_debut.setter
    def heure_debut(self, valeur: time) -> None:
        """
        Modifie l'heure (heure + minute) de _date_debut en gardant la même date.
        
        Args:
            valeur (time) : Nouvelle heure à appliquer.
        """
        self._date_debut = datetime.combine(self._date_debut.date(), valeur)

    @property
    def heure_fin(self) -> time:
        """
        Retourne l'heure de fin (heure + minute) de _date_fin.
        """
        return self._date_fin.time()

    @heure_fin.setter
    def heure_fin(self, valeur: time) -> None:
        """
        Modifie l'heure (heure + minute) de _date_fin en gardant la même date.
        
        Args:
            valeur (time) : Nouvelle heure à appliquer.
        """
        self._date_fin = datetime.combine(self._date_fin.date(), valeur)

    @property
    def duree(self) -> timedelta:
        """
        Retourne la durée du RDV sous forme de timedelta (_date_fin - _date_debut).
        """
        return self._date_fin - self._date_debut

    @duree.setter
    def duree(self, valeur: timedelta) -> None:
        """
        Modifie la date de fin en fonction de la date de début et de la durée donnée.
        
        Args:
            valeur (timedelta) : Durée à appliquer.
        """
        self._date_fin = self._date_debut + valeur

    # -------------------------------------------------------------------------
    # Représentation texte (__str__)
    # -------------------------------------------------------------------------
    def __str__(self) -> str:
        """
        Retourne une représentation textuelle lisible de l’instance.

        Exemple :
        ---------
        >>> print(rdv)
        📅 Rendez-vous Outlook
          Sujet : Réunion projet
          Lieu : Salle A
          Date début : 2025-09-25 14:00:00
          Date fin : 2025-09-25 15:00:00
        """
        header = "📅 Rendez-vous Outlook\n"
        lignes = []
        for k, v in vars(self).items():
            if v is not None and v != []:
                lignes.append(f"  {k[1:]} : {v}")
        return header + "\n".join(lignes)

### --------------------------------------------------------------------
#  Initialisations variables globales communes
### --------------------------------------------------------------------

# Pour couleur barres de progression
colorama.init(autoreset=True)

# Pour chrono des fonctions
timer = Timer()

# Pour message de sortie applis externes
vlog = Vlog()
