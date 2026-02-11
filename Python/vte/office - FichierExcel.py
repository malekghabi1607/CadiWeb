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
from win32com.client import CDispatch
import extract_msg
import os
import shutil
import sys
import time as time_module     # pour time_module.sleep()
import copy
import json

from datetime import date, datetime, timedelta, time
from zoneinfo import ZoneInfo
from io import StringIO

from tqdm import tqdm
import colorama
from colorama import Fore, Style

from tkinter import filedialog, Tk

from screeninfo import get_monitors
import ctypes

### --------------------------------------------------------------------
#  Définitions classes et fonctions génériques
### --------------------------------------------------------------------


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
                                                 df: Optional[pd.DataFrame] = None,
                                                 supprimeDonneesEtRemplace: bool = False,
                                                 remplace_df_par_nouveau: bool = False,
                                                 copie_formules: bool = False) -> None:
            """
            Écrit un DataFrame dans un tableau structuré d'une feuille de calcul.

            :param df: DataFrame à intégrer. Si non renseigné, on prend le dataframe _df de la classe
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
            
            if df is None:
                df = self.df
            else:
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

            # Si désiré par l'utilisateur, on recopie les formules
            if copie_formules:
                self.recopier_formules_colonnes(indexLigneSourceFormules=self._min_row + 1,
                                                indexLigneDebutCopie=self._max_row + 1,
                                                indexLigneFinCopie=self._max_row + len(df))

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

        def recopier_formules_colonnes(self, indexLigneSourceFormules: int, indexLigneDebutCopie: int, indexLigneFinCopie: int) -> None:
            """
            Recopie les formules Excel des colonnes qui en ont dans la ligne source vers toutes les lignes
            entre indexLigneDebutCopie et indexLigneFinCopie (inclus).
            """
            for col in range(self._min_col, self._max_col + 1):
                cellule_modele = self._ws.cell(row=indexLigneSourceFormules, column=col)
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
    _repertoire: Optional[Path] = None
    _chemin_fichier: Optional[Path] = None
    _tableaux: Dict[str, FichierExcel._TableauExcel] = {}

    # openpyxl
    _wb: Optional[Workbook] = None

    # xlwings
    _xw_app: Optional[xw.App] = None
    _xw_wb: Optional[xw.Book] = None
    _xw_ws: Optional[xw.Sheet] = None


    # === Constructeur de FichierExcel ===
    def __init__(self, chemin_fichier: Optional[Path] = None) -> None:
        self._chemin_fichier = chemin_fichier
        self._tableaux = {}
        if chemin_fichier:
            self._repertoire = chemin_fichier.parent

    # === Constructeurs alternatifs de FichierExcel ===
    @classmethod
    def depuis_repertoire(cls, repertoire: Path) -> FichierExcel:
        instance = cls()
        instance._repertoire = repertoire
        return instance

    @classmethod
    def depuis_fichier(cls,
                      chemin_fichier: Optional[Path|str] = None,
                      avec_ouverture_wb: bool = True,
                      charger_tableau: bool = True,
                      nom_onglet: Optional[str] = None,
                      nom_tableau: Optional[str] = None,
                      nbLignes_avantET: int = 0,
                      charger_df: bool = True,
                      repertoire_recherche_ini:Optional[Path|str] = None) -> FichierExcel:
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
        if chemin_fichier:
            if isinstance(chemin_fichier, str):
                chemin_fichier = Path(chemin_fichier)
        else :
            chemin_fichier = cls.choisirFichiers_filedialog(initialdir=repertoire_recherche_ini if isinstance(repertoire_recherche_ini, Path) else Path(repertoire_recherche_ini))

        instance = cls.depuis_repertoire(chemin_fichier.parent)
        instance._chemin_fichier = chemin_fichier

        #timer.debut(f"Lecture de {instance._chemin_fichier}")

        if avec_ouverture_wb:
            instance.charger_wb()

        if charger_tableau:
            if nom_onglet:
                instance.charger_tableau(nom_onglet, nom_tableau, nbLignes_avantET, charger_df)
            else:
                instance.charger_tableaux(charger_df=charger_df)

        vlog.ajouter_message("OK", f"✅ Lecture de {instance._chemin_fichier}")
        #timer.fin()

        return instance

    @classmethod
    def depuis_modele(cls,
                      chemin_modele: Optional[Path] = None,
                      chemin_fichier_sauv: Optional[Path] = None,
                      avec_ouverture_wb: bool = True,
                      charger_tableau: bool = True,
                      nom_onglet: Optional[str] = None,
                      nom_tableau: Optional[str] = None,
                      nbLignes_avantET: int = 0,
                      charger_df: bool = True,
                      repertoire_recherche_ini:Optional[Path] = None) -> FichierExcel:
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
            charger_df=charger_df,
            repertoire_recherche_ini=repertoire_recherche_ini
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

    def _sauver_fichier(self, chemin_cible: Path, copier: bool = False) -> None:
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

        chemin_cible.parent.mkdir(parents=True, exist_ok=True)

        if copier:
            if not self._chemin_fichier:
                vlog.log_erreur("Chemin fichier original non défini, impossible de copier")
                return
            try:
                shutil.copy2(self._chemin_fichier, chemin_cible)
                self._chemin_fichier = chemin_cible
                self._repertoire = chemin_cible.parent
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

    def save(self, nouveau_chemin_fichier: Optional[Path] = None) -> None:
        """Sauvegarde le fichier avec openpyxl (écriture du workbook en mémoire)."""
        chemin = nouveau_chemin_fichier or self._chemin_fichier
        self._sauver_fichier(chemin, copier=False)

    def save_copie_physique(self, nouveau_chemin: Path) -> None:
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
    def choisirFichiers_filedialog(initialdir: Optional[Path] = None) -> Path:
        return choisirFichiers_filedialog_excel(initialdir=initialdir)

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
    def chemin_fichier(self) -> Optional[Path]:
        return self._chemin_fichier

    @chemin_fichier.setter
    def chemin_fichier(self, nouveau_chemin: Optional[str|Path]) -> None:
        if nouveau_chemin == self._chemin_fichier:
            return  # pas besoin de recharger si même chemin

        if isinstance(nouveau_chemin, str):
            nouveau_chemin = Path(nouveau_chemin)

        self._chemin_fichier = nouveau_chemin
        self._repertoire = nouveau_chemin.parent if nouveau_chemin else None

        # Recharge automatiquement le fichier si il existe
        if nouveau_chemin and nouveau_chemin.is_file():
            try:
                self._wb = load_workbook(nouveau_chemin)
            except Exception as e:
                vlog.log_erreur(f"Erreur lors du chargement du fichier {nouveau_chemin} : {e}")
                self._wb = None
        else:
            self._wb = None

    @property
    def repertoire(self) -> Optional[Path]:
        return self._chemin_fichier.parent if self._chemin_fichier else None

    @property
    def nom_fichier(self) -> Optional[Path]:
        if self._chemin_fichier is None:
            return None
        return self._chemin_fichier.name

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




### --------------------------------------------------------------------
#  Initialisations variables globales communes
### --------------------------------------------------------------------

# Pour couleur barres de progression
colorama.init(autoreset=True)

# Pour chrono des fonctions
timer = Timer()

# Pour message de sortie applis externes
vlog = Vlog()
