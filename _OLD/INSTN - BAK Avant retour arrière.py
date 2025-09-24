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
from importlib.resources import path
from posixpath import basename

import warnings
from xmlrpc.client import Boolean
warnings.filterwarnings("ignore", message="Slicer List extension is not supported and will be removed")

from fileinput import filename
from typing import Dict, List, Tuple, Optional, Union

import pandas as pd
import pandas as DataFrame

from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils.cell import range_boundaries, get_column_letter

import xlwings as xw

from docx import Document
from lxml import etree

import win32com.client

import inspect
import os
import shutil
import sys
import platform
import subprocess
import re
import time
import math
import copy
import json
import ctypes
from ctypes import wintypes

from datetime import date, datetime
from io import StringIO
from dataclasses import dataclass
from collections import defaultdict

from tqdm import tqdm
import colorama
from colorama import Fore, Style

from mailmerge import MailMerge

import tkinter as tk
from tkinter import filedialog, ttk, messagebox, Tk, font

### --------------------------------------------------------------------
#  Définitions classes et fonctions génériques
### --------------------------------------------------------------------

class TableauExcel():
    """
    Représente un tableau contenu dans un fichier Excel, pouvant être un tableau structuré (Excel Table)
    ou un tableau "normal" (données tabulaires sans table Excel explicite).

    Deux modes de création sont proposés via des constructeurs de classe :
    - `depuis_tableau_structure(...)` : pour un tableau structuré avec nom de table Excel
    - `depuis_tableau_normal(...)` : pour un tableau non structuré, où l'utilisateur indique 
      le nombre de lignes avant l'en-tête du tableau

    :param repertoire: Répertoire contenant le fichier Excel
    :type repertoire: str
    :param nom_onglet: Nom de l’onglet contenant le tableau
    :type nom_onglet: str
    :param nom_tableau: (Optionnel) Nom du tableau structuré Excel (si applicable)
    :type nom_tableau: str or None
    :param nbLignes_avantET: (Optionnel) Nombre de lignes avant l'en-tête du tableau (si tableau normal)
    :type nbLignes_avantET: int or None

    :Example:

    >>> # Cas d’un tableau structuré
    >>> tab = TableauExcel.depuis_tableau_structure("chemin/fichier.xlsx", "Feuil1", "Tableau1")

    >>> # Cas d’un tableau non structuré
    >>> tab = TableauExcel.depuis_tableau_normal("chemin/fichier.xlsx", "Feuil2", nbLignes_avantET=2)

    :ivar _ref_tableau: Référence Excel du tableau (ex: "A3:D15")
    :ivar _min_row: Ligne de début du tableau
    :ivar _max_row: Ligne de fin du tableau
    :ivar _min_col: Colonne de début du tableau
    :ivar _max_col: Colonne de fin du tableau
    :ivar _nbLignes_avantET: Nombre de lignes avant l'en-tête (calculé automatiquement si tableau structuré)
    """
    # === Constructeurs ===
    def __init__(self, fichierExcel:FichierExcel=None, ws:Worksheet=None, nom_onglet:str=None, nom_tableau:str=None, nbLignes_avantET:int=0):
        """
        Créer des instances de TableauExportIRIS sans charger immédiatement un fichier Excel
        """
        # Informations du fichierExcel
        self._chemin_fichier = fichierExcel._chemin_fichier if fichierExcel else None
        self._wb = fichierExcel._wb if fichierExcel else None

        # Informations de la classe
        self._ws = ws
        self._nom_onglet = nom_onglet
        self._est_tableau_structure:bool = None
        self._df:DataFrame = None
        
        # Informations si tableau structuré (peuvent être être None si on a un tableau normal)
        self._nom_tableau = nom_tableau
        self._table = None  # Ne contient que les méta-données du tableau, pas les données elles-mêmes

        # Informations si tableau normal (peuvent être être None si on a un tableau structuré)
        self._nbLignes_avantET = nbLignes_avantET

        # Dimensions et références tableau
        self._ref_tableau = None  # Pour mémoriser la référence (ex: A1:D12)
        self._min_row = None
        self._max_row = None
        self._min_col = None
        self._max_col = None
        self._nb_lignes_tableau = None
        self._nb_lignesData_tableau = None

    @classmethod
    def depuis_FichierExcel_et_nomOnglet(cls, fichierExcel:FichierExcel, nom_onglet:str, nom_tableau:str=None, nbLignes_avantET:int=0, charger_df:bool=True):
        """Un seul tableau importé par onglet"""
        instance = cls(fichierExcel=fichierExcel, nom_onglet=nom_onglet, nom_tableau=nom_tableau, nbLignes_avantET=nbLignes_avantET)
        
        # TODO besoin de close _wb ? → Il semble que oui car j'ai dû charger mon wb
        #fichierExcel.charger_wb() → Ca crée d'id différents à tester mais je laisse en commentaire car à un moment j'en avais eu besoin
        #print("BBB")
        #print(fichierExcel._wb)
        instance._ws = fichierExcel._wb[nom_onglet]
        instance._nom_onglet = nom_onglet
        instance._nom_tableau=nom_tableau or instance._nom_onglet  # Si nom_tableau n'est pas donné, alors nom_tableau = nom_onglet (convention de nommage des fichiers Excel)
        instance._est_tableau_structure = instance.est_tableau_structure()

        if instance._est_tableau_structure :
            instance._table = instance._ws.tables[instance._nom_tableau]
            instance._initialiser_dimensions_structured_table()
        else :
            instance._nbLignes_avantET = nbLignes_avantET
            instance._initialiser_dimensions_tableau_normal()
        instance._nb_lignes_tableau = instance._max_row - instance._min_row + 1
        instance._nb_lignesData_tableau = instance._nb_lignes_tableau - instance._table.headerRowCount

        if charger_df:
            instance.charge_df()
        
        return(instance)
    
    @classmethod
    def depuis_tableau_structure(cls, nom_onglet:str, nom_tableau:str = None, wb:Workbook=None):
        """Si nom_tableau n'est pas donné, alors nom_tableau = nom_onglet (convention de nommage des fichiers Excel)"""
        instance = cls(nom_onglet=nom_onglet, nom_tableau=nom_tableau or nom_onglet)
        if wb: instance.charger_workbook(wb)
        return instance

    @classmethod
    def depuis_tableau_normal(cls, nom_onglet:str, nbLignes_avantET:int=0, wb:Workbook=None):
        instance = cls(nom_onglet=nom_onglet, nbLignes_avantET=nbLignes_avantET)
        if wb: instance.charger_workbook(wb)
        return instance

    # === Méthodes internes ===
    def _initialiser_dimensions_structured_table(self):
        self._ref_tableau = self._table.ref
        self._min_col, self._min_row, self._max_col, self._max_row = range_boundaries(self._ref_tableau)
        self._nbLignes_avantET = self._min_row - 1  # nombre de lignes avant l’en-tête du tableau

    def _initialiser_dimensions_tableau_normal(self):
        # Suppose que le tableau commence juste après nbLignes_avantET
        self._min_row = self._nbLignes_avantET + 1
        self._max_row = self._ws.max_row
        self._min_col = 1
        self._max_col = self._ws.max_column
        self._ref_tableau = f"{self._ws.cell(row=self._min_row, column=self._min_col).coordinate}:{self._ws.cell(row=self._max_row, column=self._max_col).coordinate}"

    # === Méthodes externes ===
    def charger_workbook(self, wb:Workbook):
        """Méthode à appeler manuellement si le fichier est connu plus tard"""
        self._ws = wb[self._nom_onglet]

        if self._nom_tableau:
            self._table = self._ws.tables[self._nom_tableau]
            self._initialiser_dimensions_structured_table()
        else:
            self._initialiser_dimensions_tableau_normal()

    def est_tableau_structure(self) -> bool:
        """
        Vérifie si un tableau structuré (ListObject) portant le nom donné
        existe bien dans la feuille Excel donnée.

        Args:
            ws (Worksheet): Feuille Excel (openpyxl)
            nom_tableau (str): Nom du tableau structuré recherché

        Returns:
            bool: True si le tableau existe, False sinon
        """
        # openpyxl stocke les tableaux dans l'attribut 'tables' du Worksheet
        return self._nom_tableau in self._ws.tables

    def charge_df(self):
        #df = pd.read_excel(cheminFichier, skiprows=self.__nbLignes_avantET_exportIRIS)
        #self.__df_tableau = pd.concat(df_list, ignore_index=True) 
        if self._chemin_fichier:
            self._df = pd.read_excel(self._chemin_fichier, self._nom_onglet, skiprows=self._min_row-1)
        else:
            #print("Dans TableauExcel.charge_df(), pas de chemin vers le fichier\n=== Exit() ===")
            #exit()
            log_erreur("pas de chemin vers le fichier")

    def remplace_df(self, df:DataFrame):
        self._df = df.copy()

    def ecrit_dataFrame_dans_tableauStructure(self, df:DataFrame, supprimeDonneesEtRemplace:Boolean=False, remplace_df_par_nouveau:Boolean=False):
        """
        Ecrit un DataFrame dans un tableau structure d'une feuille de calcul  

        :param supprimeDonneesEtRemplace: Pour savoir si l'on ajoute les données du DataFrame à l'existant (False) ou si l'on supprime les données existantes et qu'on les remplace avec celles du DataFrame
        :type supprimeDonneesEtRemplace: Boolean
        :param remplace_df_par_nouveau: Pour remplacer (mise à jour) le dataframe du tableau structuré (i.e. de la classe) par celui en entrée de la fonction
        :type remplace_df_par_nouveau: Boolean
        :return: rien (on écrit/sauve un fichier excel)
        :rtype: None

        :Example:

        >>> ecrit_dataFrame_dans_tableauStructure(df_output, supprimeDonneesEtRemplace = False, remplace_df_par_nouveau:Boolean=True)


        .. seealso:: Rien du tout.
        .. warning:: Rien du tout.
        .. note:: Pour rajouter des lignes, on le fait à la suite de la feuille (worksheet) . Il en résulte qu'on gruge un peu : 1) on vire le tableau structuré, 2) on colle toutes les nouvelles lignes, 3) on supprime la ligne 1 du tableau, 4) on redéfinit les dimensions du tableau structuré (car l'ajout de nouvbelles lignes ne l'étend pas automatiquement)
        .. todo:: Rien du tout.
        """


        if remplace_df_par_nouveau:
            self.remplace_df(df)
        
        # On écrit toutes les autres lignes une par une (on garde les lignes initiales pour garder le format qu'on copiera)
        # Méthode 1 qui marche
        with tqdm(total=len(self._df), unit=' ligne', desc=Fore.CYAN + "Écriture des lignes dans l'output" + Style.RESET_ALL) as pbar:
            for i, il in enumerate(self._df.itertuples(), 1):
                print(il)
                pbar.set_postfix(progress=f"{i}/{len(self._df)}")
                row = [val if pd.notna(val) else None for val in il[1:]] #Je dois rajouter cette ligne car il faut tester si je n'ai pas de valeurs <NA> qu'il faut retravailler sinon ça plante
                self._ws.append(row)
                pbar.update(1)

        # On redéfinit le dimensionnement du tableau (tableau initial + nb lignes de df)
        self._table.ref = f"{self._ws.cell(row=self._min_row, column=self._min_col).coordinate}:{self._ws.cell(row=self._max_row + len(df), column=self._max_col).coordinate}"

        # On copie le format sur toutes les nouvelles lignes du tableau
        self.copieFormat_tableauStructure_openpyxl(indexLigneSourceFormat = self._min_row + 1, indexLigneDebutCopie = self._max_row + 1, indexLigneFinCopie = self._max_row + len(df) +1)  # TODO : j'ai viré le +1 sur indexLigneFinCopie = self._max_row + len(df)

        # Si désiré par l'utilisateur, alors on supprime les anciennes lignes de la feuille Excel (ça garde la dimension initiale du tableau structuré)
        if supprimeDonneesEtRemplace:
            # Suppression des anciennes lignes
            self._ws.delete_rows(idx=self._min_row + 1, amount=self._nb_lignesData_tableau)

            # On redéfinit les dimensions du tableau structuré
            self._table.ref = f"{self._ws.cell(row=self._min_row, column=self._min_col).coordinate}:{self._ws.cell(row=self._min_row + len(df), column=self._max_col).coordinate}"  

        # Vérifier si la première ligne du DataFrame est entièrement vide (tous les éléments NaN)
        if df.iloc[0].isna().all():
            print("La première ligne est entièrement vide")
            # Suppression de la 1ère ligne
            self._ws.delete_rows(idx=self._min_row + 1, amount=1)

            # On redéfinit les dimensions du tableau structuré
            self._table.ref = f"{self._ws.cell(row=self._min_row, column=self._min_col).coordinate}:{self._ws.cell(row=self._min_row + len(df) - 1, column=self._max_col).coordinate}"  

        # On met à jour les références des mises en formes conditionnelles
        self.maj_references_misesEnFormeConditionnelles()        


    def copieFormat_tableauStructure_openpyxl(self, indexLigneSourceFormat:int=1, indexLigneDebutCopie:int=None, indexLigneFinCopie:int=None) :
        """
        Recopie le format d'une ligne d'un tableau structure a une plage du tableau structure

        :param ws: feuille dans laquelle est le tableau structure
        :type ws: openpyxl.worksheet
        :param table: tableau structure
        :type table: openpyxl.Table
        :param indexLigneSourceFormat: indice de la ligne qui est a recopier (indice absolu dans la feuille excel, i.e. pas #ligne dans le tableau structure). Defaut = 1.
        :type indexLigneSourceFormat: int
        :param indexLigneDebutCopie: indice de la 1ere ligne ou il faut copier le format (indice absolu dans la feuille excel, i.e. pas #ligne dans le tableau structure). Defaut = indexLigneSourceFormat + 1
        :type indexLigneDebutCopie: int
        :param indexLigneFintCopie: indice de la derniere ligne ou il faut copier le format (indice absolu dans la feuille excel, i.e. pas #ligne dans le tableau structure). Defaut = derniere ligne tableau structure
        :type indexLigneFinCopie: int
        :return: rien (on a copie les format dans le worksheet)
        :rtype: None

        :Example:

        >>> copieFormatTableauStructure(ws, table, indexLigneSourceFormat = min_row + 1, indexLigneDebutCopie = max_row + 1, indexLigneFinCopie = max_row + len(df) +1)


        .. seealso:: Rien du tout.
        .. warning:: Rien du tout.
        .. note:: Pour rajouter des lignes, on le fait à la suite de la feuille (worksheet) . Il en résulte qu'on gruge un peu : 1) on vire le tableau structuré, 2) on colle toutes les nouvelles lignes, 3) on supprime la ligne 1 du tableau, 4) on redéfinit les dimensions du tableau structuré (car l'ajout de nouvbelles lignes ne l'étend pas automatiquement)
        .. todo:: Rien du tout.
        """

        # On met à jour les styles des cellules (openpywl ne sait pas insérer de lignes en conservant les formats ; par ailleurs on ne sait pas appliquer ça ligne par ligne ou colonne par colonne : on va donc le faire cellule par cellule)
        # On n'agrandit pas automatiquement le tableau structuré si l'on rajoute une cellule
        # Quand on agrandit un tableau en redéfinissant le ref, on ne colle pas le format



        # Infos de longueurs de mon tableau
        #min_col, min_row, max_col, max_row = range_boundaries(table.ref)
        
        # Initialisattion paramètres non renseignés
        if not indexLigneDebutCopie:
            indexLigneDebutCopie = indexLigneSourceFormat + 1
        if not indexLigneFinCopie:
            indexLigneFinCopie = self._max_row


        # On récupère les formats de chaque cellule de la première ligne du tableau structuré 
        formats=[]
        for icol in range(self._min_col, self._max_col+1) :
            formats.append(self._ws.cell(indexLigneSourceFormat, icol))
            #print(formats[icol-1].number_format)
        #print(formats)

        # On copie colle les formats avec ces cellules
        with tqdm(total=self._max_col, unit=' colonnes', desc=Fore.CYAN + "Copie des formats" + Style.RESET_ALL, ncols=150) as pbar:
            for icol in range(self._min_col, self._max_col+1) :  # TODO : pas sûr du +1
                pbar.set_postfix(progress=f"{icol}/{self._max_col}")
                source = formats[icol-1] #Je prends un index de liste et pas un numéro de colonne, donc -1
                #print(source.number_format, source.number_format == "General")

                # Optimisation : on ne fait les copies que si le format est différent de General
                #if source.number_format != "General":
                for il in range(indexLigneDebutCopie, indexLigneFinCopie) :
                    #print(il, icol, ws.cell(row=il, column=icol).value, source.number_format)
                    target = self._ws.cell(row=il, column=icol)
                    #target.font = copy.copy(source.font)
                    #target.border = copy.copy(source.border)
                    target.fill = copy.copy(source.fill)
                    #target.alignment = copy.copy(source.alignment)
                    #target.protection = copy.copy(source.protection)
                    target.number_format = copy.copy(source.number_format)

                pbar.update(1)

    def maj_references_misesEnFormeConditionnelles(self, nouvelle_plage:str=None, nouvelle_ligne_max:int=None):
        """
        Met à jour les plages des règles de mise en forme conditionnelle appliquées au tableau structuré.

        Cette méthode parcourt toutes les règles de mise en forme conditionnelle du worksheet et
        réapplique celles qui concernent uniquement le tableau courant à une nouvelle plage définie.

        Selon les arguments fournis, le comportement est le suivant :
        - Si aucun argument n'est donné, la nouvelle plage conserve les colonnes d'origine 
        et s'étend jusqu'à la dernière ligne du tableau (incluant la ligne Totaux si présente).
        - Si `nouvelle_plage` est fourni, les règles sont déplacées vers cette plage spécifique.
        - Si `nouvelle_ligne_max` est fourni, la plage est recalculée pour conserver les colonnes 
        d'origine et aller jusqu'à la ligne indiquée par `nouvelle_ligne_max`.

        Les anciennes plages de règles sont supprimées après leur réaffectation.

        :param nouvelle_plage: Nouvelle plage Excel (ex: "B4:F20"). Si fourni, ce paramètre prévaut sur `nouvelle_ligne_max` et la logique automatique.
        :type nouvelle_plage: str or None
        :param nouvelle_ligne_max: Nouvelle ligne maximale jusqu’où étendre la plage sur les colonnes d'origine.
        :type nouvelle_ligne_max: int or None

        :Example:

        >>> # Étendre automatiquement toutes les règles sur la hauteur actuelle du tableau
        >>> obj.maj_references_misesEnFormeConditionnelles()

        >>> # Appliquer toutes les règles à une plage précise
        >>> obj.maj_references_misesEnFormeConditionnelles(nouvelle_plage="D5:H50")

        >>> # Étendre toutes les règles aux mêmes colonnes, mais jusqu'à la ligne 100
        >>> obj.maj_references_misesEnFormeConditionnelles(nouvelle_ligne_max=100)
        """
        # Met à jour les dimensions du tableau
        self._initialiser_dimensions_structured_table()

        # Récupérer toutes les règles CF (clé = plage, valeur = liste des règles)
        cf_rules = self._ws.conditional_formatting._cf_rules
        min_col_t, min_row_t, max_col_t, max_row_t = (self._min_col, self._min_row, self._max_col, self._max_row)

        # On fait une copie de la liste des plages à traiter (pour pouvoir supprimer sans erreur)
        for cf_obj, regles in list(cf_rules.items()): # C'est une boucle sur les plages et on obtient les règles appliquées à cette plage ; cf_obj est un objet ConditionalFormatting
            # Récupérer les coordonnées de l'ancienne plage
            #ancienne_plage = cf_obj.sqref  # c’est la plage Excel sous forme de string, ex "A1:AG2"
            multi_range = cf_obj.sqref  # c'est un MultiCellRange

            for cell_range in multi_range:
                ancienne_plage = cell_range.coord # c’est la plage Excel sous forme de string, ex "A1:AG2"
                #print(f"Traitement de la plage : {ancienne_plage}")

                min_col_r, min_row_r, max_col_r, max_row_r = range_boundaries(ancienne_plage)

                # Vérifie si la règle s'applique entièrement **à l'intérieur** du tableau
                if not (
                    #min_col_t <= min_col_r <= max_col_r <= max_col_t and
                    #min_row_t <= min_row_r <= max_row_r <= max_row_t

                    # Cas adapté
                    min_col_t <= min_col_r <= max_col_r <= max_col_t and
                    min_row_t <= min_row_r <= max_row_r+1 <= max_row_t+1  # On met les +1 pour prendre éventuellement la ligne total
                ):
                    log_erreur(f"une règle est en dehors du tableau dont on ne l'applique pas :{min_col_r}, {min_row_r}, {max_col_r}, {max_row_r} → On ignore la règle.", continuer=True)
                    continue  # On ignore cette règle car elle n'est pas dans le tableau

                # Déterminer la nouvelle plage selon l’argument
                if nouvelle_plage is None:
                    nouvelle_plage = f"{get_column_letter(min_col_r)}{min_row_r}:{get_column_letter(max_col_r)}{max_row_t+1}" #+1 pour prendre éventuellement une ligne total
                if nouvelle_ligne_max is not None:
                    nouvelle_plage = f"{get_column_letter(min_col_r)}{min_row_r}:{get_column_letter(max_col_r)}{nouvelle_ligne_max}"

                
                # Appliquer toutes les règles à la nouvelle plage
                for regle in regles:
                    self._ws.conditional_formatting.add(nouvelle_plage, regle)

                # Supprimer l'ancienne règle (elle a été déplacée)
                del self._ws.conditional_formatting._cf_rules[cf_obj]

    def generer_dictionnaire_depuis_excel(self) -> None:
        """
        Récupère les noms des colonnes et prépare un dictionnaire avec ces noms de colonne en clef pour faciliter une déclaration
        """
        # Charge le fichier Excel
        #if self._df is None:
        #    df = pd.read_excel(self._chemin_fichier, sheet_name=nom_feuille)

        # Nettoie les noms de colonnes
        #noms_colonnes = [nettoyer_nom_colonne(col) for col in self._df.columns]

        # Récupération brute des noms de colonnes
        noms_colonnes = list(self._df.columns)

        # Impression avec noms de colonnes échappés (affichage sûr en syntaxe Python)
        print("mon_dictionnaire = {")
        for i, cle in enumerate(noms_colonnes):
            virgule = "," if i < len(noms_colonnes) - 1 else ""
            # json.dumps gère l'échappement des caractères spéciaux (ex: \n, \r, \t, etc.)
            print(f"    {json.dumps(cle)}: \"\"{virgule}")
        print("}")

    def ligne_contient_formules(self, ligne: int) -> bool:
        """
        Renvoie True si une des cellules de la ligne contient une formule Excel
        """
        for col in range(self._min_col, self._max_col + 1):
            cell = self._ws.cell(row=ligne, column=col)
            if cell.data_type == 'f' or (cell.value is not None and isinstance(cell.value, str) and cell.value.startswith("=")):
                return True
        return False

    def ligne_vide_ou_formules(self, ligne:int) -> bool:
        """
        Retourne True si toutes les cellules de la ligne sont vides ou contiennent une formule.
        """
        for col in range(self._min_col, self._max_col + 1):
            cell = self._ws.cell(row=ligne, column=col)
            if cell.value is not None:
                # Si la cellule a une formule
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    continue  # cellule avec formule, OK
                # Si la cellule a une valeur non vide (pas formule)
                else:
                    return False
        return True

    def recopier_formules_colonnes(self, indexLigneSourceFormat:int, indexLigneDebutCopie:int, indexLigneFinCopie:int):
        """
        Recopie les formules Excel des colonnes qui en ont dans la ligne source vers toutes les lignes entre indexLigneDebutCopie et indexLigneFinCopie (inclus).
        indexLigneSourceFormat : ligne qui sert de modèle (par ex. ligne 2)
        indexLigneDebutCopie, indexLigneFinCopie : plage de lignes où recopier les formules
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

    def supprimer_premiere_ligne_tableau_structuré(self):
        """
        Supprime proprement la première ligne de données du tableau structuré,
        en ajustant la plage du tableau pour éviter les erreurs à l'ouverture dans Excel.
        """

        ligne_a_supprimer = self._min_row + 1

        # 1. Supprimer physiquement la ligne
        self._ws.delete_rows(idx=ligne_a_supprimer, amount=1)

        # 2. Recalculer la nouvelle plage du tableau
        nb_lignes_données = len(self._df) - 1  # -1 car on supprime une ligne
        nouvelle_ref = f"{self._ws.cell(row=self._min_row, column=self._min_col).coordinate}:" \
                    f"{self._ws.cell(row=self._min_row + nb_lignes_données, column=self._max_col).coordinate}"

        # 3. Appliquer la nouvelle plage au tableau
        self._table.ref = nouvelle_ref

        # 4. Mettre à jour aussi le filtre automatique (sinon Excel râle)
        if self._table.autoFilter:
            self._table.autoFilter.ref = self._table.ref
    
    def supprimer_ligne_excel_via_xlwings(self, ligne_excel: int):
        """
        Supprime une ligne dans le fichier Excel via xlwings pour préserver la structure du tableau structuré,
        en évitant les erreurs à l'ouverture dues aux formules ou filtres.
        """

        print(f"Suppression de la ligne {ligne_excel} via xlwings...")
        print(self._chemin_fichier)
        # self._chemin_fichier = \\harmonie\INSTN\UEM\_Documents_communs\Formations\Formateurs\P09-Pr01-Qualifier les ressources enseignantes\P09_Pr01_Ta.E_Grille des critères de qualification des compétences_V1.xlsx

        if not os.path.exists(self._chemin_fichier):
            raise FileNotFoundError(f"Fichier Excel introuvable : {self._chemin_fichier}")

        app = xw.App(visible=True)
        app.display_alerts = True
        app.screen_updating = True

        wb = app.books.open(self._chemin_fichier)

        ws = wb.sheets[self._ws.title]
        ws.range(f"{ligne_excel}:{ligne_excel}").api.Delete()

        wb.save(self._chemin_fichier)
        print("Ligne supprimée proprement via xlwings.")
        wb.close()
        app.quit()

        #try:
        #    wb = app.books.open(self._chemin_fichier)

        #    ws = wb.sheets[self._ws.title]
        #    ws.range(f"{ligne_excel}:{ligne_excel}").api.Delete()
        #    input()

        #    wb.save(self._chemin_fichier)
        #    print("Ligne supprimée proprement via xlwings.")
        #finally:
        #    wb.close()
        #    app.quit()

    # === Propriétés ===
    @property
    def nom_onglet(self):
        return self._nom_onglet

    @property
    def nom_tableau(self):
        return self._nom_tableau

    @property
    def nbLignes_avantET(self):
        return self._nbLignes_avantET

    @property
    def feuille(self):
        return self._wb[self._nom_onglet]

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
    def __str__(self):
        # Partie 1 : texte descriptif
        # S'il n'y a pas de référence tableau, alors notre instance est un répertoire export IRIS par défaut
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

        # Partie 2 : affichage du DataFrame s’il existe
        if self._df is not None:
            buffer = StringIO()
            stdout_original = sys.stdout
            try:
                sys.stdout = buffer
                print(self._df)  # exactement ce que print afficherait
            finally:
                sys.stdout = stdout_original
            contenu_df = buffer.getvalue()
            return f"{description}\n\n{contenu_df.rstrip()}"
        else:
            return description

class FichierExcel:
    """
    Classe pour manipuler un fichier Excel avec openpyxl.
    Responsable de l’ouverture, de la gestion du workbook et du suivi des tableaux Excel contenus dans le fichier.
    """

    # === Constructeurs ===
    def __init__(self):
        self._repertoire = None
        self._chemin_fichier = None
        self._wb = None
        self._tableaux:Dict[str, TableauExcel] = {}  # Dict[str, TableauExcel]

    @classmethod
    def depuis_repertoire(cls, repertoire:str):
        instance = cls()
        instance._repertoire = repertoire
        return instance

    @classmethod
    def depuis_fichier(cls, chemin_fichier:str=None, avec_ouverture_wb:bool=True, charger_tableau:bool=True, nom_onglet:str=None, nom_tableau:str=None, nbLignes_avantET:int=0, charger_df:bool=True):
        """
        Par défaut ou ouvre le workbook, mais on peut spécifier que non avec avec_ouverture_wb=False
        Si charger_tableau = True
            Si pas de nom_onglet : on charge tous les tableaux structurés
            Si nom_onglet : on charge un seul tableau (structuré ou non) qui est dans cet onglet (il faut alors un nom_onglet)
        """
        #Reprise de comportement (si on ne demande pas l'ouverture (chargement wb), alors pas de raison de charger le tableau. Idem avec charger_tableau et charger_df)
        if not avec_ouverture_wb: charger_tableau=False
        if not charger_tableau: charger_df=False

        # Si aucun fichier d'entrée, alors l'utilisateur le pointe avec filedialog
        if not chemin_fichier:
            chemin_fichier = cls.choisirFichiers_filedialog()
        
        # Initialisation de l'instance
        instance = cls.depuis_repertoire(os.path.dirname(chemin_fichier))
        #instance._chemin_fichier = chemin_fichier
        instance.chemin_fichier = chemin_fichier

        # Chargement Workbook
        if avec_ouverture_wb:
            instance.charger_wb()
        
        # Chargement des tableaux
        if charger_tableau:
            # Si nom_onglet : on charge un tableau (structuré ou non) dans nom_onglet ; sinon on charge tous les tableaux structurés
            if nom_onglet :
                instance.charger_tableau(nom_onglet, nom_tableau, nbLignes_avantET, charger_df)
            else:
                instance.charger_tousTableauxStructures(charger_df=charger_df)

        return instance


    @classmethod
    def depuis_modele(cls, chemin_modele:str=None, chemin_fichier_sauv:str=None, avec_ouverture_wb:bool=True, charger_tableau:bool=True, nom_onglet:str=None, nom_tableau:str=None, nbLignes_avantET:int=0, charger_df:bool=True):
        """
        Charge un modèle Excel (et soit tous ses tableaux structuré, soit un tableau non structuré dans un onglet à donner)
        On définit le chemin où devra être sauvegardé ce modèle une fois rempli
        """
        instance = cls.depuis_fichier(chemin_modele, avec_ouverture_wb, charger_tableau, nom_onglet, nom_tableau, nbLignes_avantET, charger_df)
        #instance._chemin_fichier = chemin_fichier_sauv
        instance.chemin_fichier = chemin_fichier_sauv

        return instance


    # === Méthodes utilitaires ===
    def charger_wb(self):
        self._wb = load_workbook(self._chemin_fichier)
    
    def charger_tableau(self, nom_onglet:str, nom_tableau:str=None, nbLignes_avantET:int=0, charger_df:bool=True):
        self.ajouter_tableau(TableauExcel.depuis_FichierExcel_et_nomOnglet(self, nom_onglet=nom_onglet, nom_tableau=nom_tableau, nbLignes_avantET=nbLignes_avantET, charger_df=charger_df))
    
    def charger_tousTableauxStructures(self, charger_df:bool=True):
        for nom_onglet in self._wb.sheetnames:
            for nom_tableau in self._wb[nom_onglet].tables:
                self.ajouter_tableau(TableauExcel.depuis_FichierExcel_et_nomOnglet(self, nom_onglet=nom_onglet, nom_tableau=nom_tableau, charger_df=charger_df))

    def ajouter_tableau(self, tableau:TableauExcel):
        self._tableaux[tableau._nom_tableau] = tableau
    
    def actualiser_TCD(self, save: bool = True, quitter: bool = True):
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False  # ou True pour voir ce qu'il fait

            wb = excel.Workbooks.Open(self._chemin_fichier)

            for sheet in wb.Sheets:
                for pivot in sheet.PivotTables():
                    pivot.RefreshTable()  # Actualise le TCD

            if save:
                wb.Save()

            if quitter:
                wb.Close(SaveChanges=False)
                excel.Quit()

    def save(self, nouveau_chemin_fichier:str=None):
        """Sauve le fichier"""
        self._wb.save(nouveau_chemin_fichier if nouveau_chemin_fichier is not None else self._chemin_fichier)

    def close(self):
        """Ferme le fichier (facultatif avec openpyxl mais pratique)."""
        self._wb.close()

    def choisirFichiers_filedialog(self):
        # Lister/sélectionner les documents à concaténer
        chemin_fichier = filedialog.askopenfilename(title="Sélectionner le fichier à charger", filetype=[("Fichiers Excel", "*.xlsx")], initialdir=self._input.repertoire)
        
        # Gestion du cas où il y a non-sélection de fichiers
        if not chemin_fichier:
            log_erreur("click sur cancel du filedialog → Pas de chemins de fichier")

        #self._repertoire = os.path.dirname(chemin_fichier)
        #self._chemin_fichier = chemin_fichier

        return chemin_fichier

    # === Propriétés ===
    @property
    def chemin_fichier(self):
        return self._chemin_fichier

    @chemin_fichier.setter
    def chemin_fichier(self, nouveau_chemin:str):
        maj_wb = False
        self._chemin_fichier = nouveau_chemin
        self._repertoire = os.path.dirname(nouveau_chemin)
        if os.path.isfile(nouveau_chemin):
            self._wb = load_workbook(nouveau_chemin)  # Recharge automatiquement le fichier
            maj_wb = True

        for _, iTableau in self._tableaux.items():
            iTableau._chemin_fichier = nouveau_chemin

            if maj_wb:
                iTableau._wb = self._wb

    @property
    def repertoire(self):
        return os.path.dirname(self._chemin_fichier)

    @property
    def nom_fichier(self):
        if self._chemin_fichier is None:
            return None
        return os.path.basename(self._chemin_fichier)
        
    @property
    def workbook(self):
        return self._wb


    # === Affichage ===
    def __str__(self):
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
            # Affichage du nom du tableau avec couleur
            titre = f"{Fore.YELLOW + Style.BRIGHT}{nom_tableau}{Style.RESET_ALL}"

            # Métadonnées
            meta = (
                f"  ─ Onglet : {tableau.nom_onglet}\n"
                f"  ─ Référence : {tableau.ref_tableau or 'non définie'}\n"
                f"  ─ Lignes avant en-tête : {tableau.nbLignes_avantET}\n"
                f"  ─ Dimensions : lignes {tableau.min_row}-{tableau.max_row}, colonnes {tableau.min_col}-{tableau.max_col}"
            )

            # Capture exacte de print(tableau._df)
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

    def __init__(self,
                 chemin_fichier: Optional[str] = None,
                 repertoire: Optional[str] = None,
                 content_controls: Optional[Dict[str, str]] = None,
                 doc_obj: Optional[Document] = None,
                 com_word_app: Optional[win32com.client.CDispatch] = None,
                 com_word_doc: Optional[win32com.client.CDispatch] = None) -> None:
        self._repertoire: Optional[str] = repertoire
        self._chemin_fichier: Optional[str] = chemin_fichier
        self._cc: Optional[Dict[str, str]] = content_controls #_cc[nomCC, ValeurCC]
        self._doc: Optional[Document] = doc_obj
        self._com_word_app = com_word_app  # COM Word Application (si utilisé)
        self._com_word_doc = com_word_doc  # COM Word Document (si utilisé)

    @classmethod
    def depuisFichier(cls,
                     chemin_fichier: Optional[str] = None,
                     charger_contentControl: bool = True,
                     afficherWord: bool = False) -> FichierWord:
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
                print(f"⚠️ Aucune Content Control trouvée ou erreur : {e}")

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
        _copies_cachees (Optional[List[str]]): Liste des destinataires en copie cachée.
        _sujet (Optional[str]): Sujet du mail.
        _corps (Optional[str]): Corps du mail (texte brut ou HTML).
        _pieces_jointes (Optional[List[str]]): Liste des chemins vers les fichiers à joindre.

    Exemple d'utilisation :
    ```python
    mail = Mail()
    mail._destinataires = ["exemple@domaine.com"]
    mail._sujet = "Test via Outlook"
    mail.definir_corps_message("<p>Bonjour, ceci est un mail préparé via Python.</p>")
    mail.selectionner_pj()
    mail.preparer_mail()  # Ouvre la fenêtre Outlook
    # ou mail.envoyer_mail() pour envoyer directement
    print(mail)
    ```
    """

    def __init__(self) -> None:
        self._destinataires: Optional[List[str]] = None
        self._copies_cachees: Optional[List[str]] = None
        self._sujet: Optional[str] = None
        self._corps: Optional[str] = None  # Corps HTML
        self._pieces_jointes: Optional[List[str]] = None

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
        root = tk.Tk()
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

    def _creer_mail(
        self,
        destinataires: Optional[Union[List[str], str]] = None,
        copies_cachees: Optional[Union[List[str], str]] = None,
        sujet: Optional[str] = None,
        corps_html: Optional[str] = None,
        pieces_jointes: Optional[List[str]] = None,
        envoyer_directement: bool = False
    ) -> None:
        """
        Méthode interne : crée, remplit et affiche/envoie le mail.

        Args:
            destinataires: liste ou str de destinataires
            copies_cachees: liste ou str de copies cachées
            sujet: sujet du mail
            corps_html: corps HTML du mail
            pieces_jointes: liste des chemins vers fichiers joints
            envoyer_directement: si True, envoie sans afficher

        :Example:

        >>> mail = Mail()
        >>> mail._creer_mail(
        >>>     destinataires=["user1@example.com", "user2@example.com", "user3@example.com"],
        >>>     copies_cachees=["cachee1@example.com", "cachee2@example.com"],
        >>>     sujet="Sujet multiple destinataires",
        >>>     corps_html="<p>Bonjour à tous, plusieurs destinataires.</p>",
        >>>     pieces_jointes=["C:\\chemin\\fichier1.pdf", "C:\\chemin\\fichier2.jpg"],
        >>>     envoyer_directement=True  # envoie directement sans afficher
        >>> )

        """
        # Mise à jour des variables d'instance si arguments fournis
        if destinataires is not None:
            if isinstance(destinataires, str):
                self._destinataires = [destinataires]
            else:
                self._destinataires = destinataires

        if copies_cachees is not None:
            if isinstance(copies_cachees, str):
                self._copies_cachees = [copies_cachees]
            else:
                self._copies_cachees = copies_cachees

        if sujet is not None:
            self._sujet = sujet

        if corps_html is not None:
            self._corps = corps_html

        if pieces_jointes is not None:
            self._pieces_jointes = pieces_jointes

        # Création Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)

        # Affecter destinataires, copie cachée, sujet
        if self._destinataires:
            mail.To = ";".join(self._destinataires)
        if self._copies_cachees:
            mail.BCC = ";".join(self._copies_cachees)
        if self._sujet:
            mail.Subject = self._sujet

        # Affiche la fenêtre pour forcer la signature à charger
        mail.Display()

        # Récupérer la signature à partir du mail créé
        signature = mail.HTMLBody

        # Construire le corps complet
        corps_complet = ""
        if self._corps:
            corps_complet = self._corps + signature
        else:
            corps_complet = signature

        mail.HTMLBody = corps_complet

        # Ajouter pièces jointes
        if self._pieces_jointes:
            if isinstance(self._pieces_jointes, str):
                pieces_jointes = [self._pieces_jointes]  # Convertit string en liste
            for pj in pieces_jointes:
                mail.Attachments.Add(pj)

        # Afficher ou envoyer
        if envoyer_directement:
            mail.Send()
        else:
            # La fenêtre est déjà affichée (Display appelé plus haut)
            pass



    def preparer_mail(
        self,
        destinataires: Optional[Union[List[str], str]] = None,
        copies_cachees: Optional[Union[List[str], str]] = None,
        sujet: Optional[str] = None,
        corps_html: Optional[str] = None,
        pieces_jointes: Optional[List[str]] = None
    ) -> None:
        """
        Prépare et affiche la fenêtre Outlook pour composer le mail.

        Args:
            destinataires: destinataires principaux (optionnel)
            copies_cachees: destinataires en copie cachée (optionnel)
            sujet: sujet du mail (optionnel)
            corps_html: corps HTML du mail (optionnel)
            pieces_jointes: liste de chemins vers fichiers joints (optionnel)

        Exemple :
        ```python
        mail = Mail()
        mail.preparer_mail(
            destinataires="user@example.com",
            sujet="Bonjour",
            corps_html="<p>Message ici</p>"
        )
        ```
        """
        self._creer_mail(destinataires, copies_cachees, sujet, corps_html, pieces_jointes, envoyer_directement=False)

    def envoyer_mail(
        self,
        destinataires: Optional[Union[List[str], str]] = None,
        copies_cachees: Optional[Union[List[str], str]] = None,
        sujet: Optional[str] = None,
        corps_html: Optional[str] = None,
        pieces_jointes: Optional[List[str]] = None
    ) -> None:
        """
        Prépare et envoie directement un mail via Outlook sans affichage.

        Args:
            destinataires: destinataires principaux (optionnel)
            copies_cachees: destinataires en copie cachée (optionnel)
            sujet: sujet du mail (optionnel)
            corps_html: corps HTML du mail (optionnel)
            pieces_jointes: liste de chemins vers fichiers joints (optionnel)

        Exemple :
        ```python
        mail = Mail()
        mail.envoyer_mail(
            destinataires="user@example.com",
            sujet="Envoi automatique",
            corps_html="<p>Voici un mail envoyé automatiquement.</p>"
        )
        ```
        """
        self._creer_mail(destinataires, copies_cachees, sujet, corps_html, pieces_jointes, envoyer_directement=True)

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
            f"  Copies cachées : {', '.join(self._copies_cachees) if self._copies_cachees else 'non définies'}",
            f"  Sujet : {self._sujet or 'non défini'}",
            f"  Corps : {self._corps[:50] + '...' if self._corps and len(self._corps) > 50 else self._corps or 'vide'}",
            f"  Nombre de pièces jointes : {len(self._pieces_jointes) if self._pieces_jointes else 0}",
        ]
        if self._pieces_jointes:
            for pj in self._pieces_jointes:
                lignes.append(f"    • {pj}")
        return "\n".join(lignes)

class Timer:
    """
    Classe Timer simple pour mesurer et afficher la durée de traitements dans un script.

    Fonctionnalités :
    -----------------
    - Affiche un message au début d'un traitement : "⏳ Traitement de <description>..."
    - Remplace ce message à la fin (automatique ou manuelle) par : "✅ <description> terminé en X min Y s."
    - Gère automatiquement la fin du chrono précédent à chaque nouvel appel de `debut(...)`.

    Utilisation :
    ------------
    >>> timer = Timer()
    >>> timer.debut("Chargement des données")
    >>> # ... traitement ...
    >>> timer.debut("Traitement des résultats")
    >>> # ... autre traitement ...
    >>> timer.fin()  # Optionnel si on veut terminer explicitement le dernier chrono

    Méthodes :
    ----------
    - debut(description: str): démarre un nouveau chronomètre et affiche un message.
                                Termine automatiquement le précédent s'il est en cours.
    - fin(): termine le chronomètre en cours et affiche la durée du traitement.

    Remarques :
    -----------
    - Aucun module externe requis (comme tqdm).
    - L'affichage est propre dans la console grâce à l'effacement dynamique de la ligne.
    - Conçu pour les scripts où l'on veut chronométrer plusieurs étapes sans se répéter.
    """
    def __init__(self, description=""):
        self.__debut = None
        self.__fin = None
        self.__duree = 0
        self.__description = description
        self.__timer_en_cours = False
        self.__last_message = ""

    def debut(self, description=""):
        # Si un timer est déjà en cours, on le termine proprement
        if self.__timer_en_cours:
            self.fin()

        self.__description = description
        self.__debut = time.time()
        self.__timer_en_cours = True

        self.__last_message = f"⏳ Traitement de {self.__description}..."
        print(self.__last_message, end='', flush=True)

    def fin(self):
        if not self.__timer_en_cours:
            return  # Rien à terminer

        self.__fin = time.time()
        self.__duree = self.__fin - self.__debut
        minutes, secondes = divmod(int(self.__duree), 60)

        message_final = f"✅ {self.__description} terminé en {minutes} min {secondes} s."

        # Nettoyer la ligne précédente et afficher le nouveau message
        clean_line = '\r' + ' ' * len(self.__last_message) + '\r'
        print(clean_line + message_final)

        self.__timer_en_cours = False

class Vlog:
    """
    Classe de journalisation visuelle (Visual Logger).

    Permet de centraliser les messages d'information, d'erreur, d'exclusion, etc.
    Affichage possible en console, en popup Tkinter, et export vers un fichier.

    Exemple :
        vlog = Vlog()
        vlog.ajouter_message("Infos", "Traitement terminé", style=["vert"])
        vlog.log_erreur("Impossible de lire le fichier")
        print(vlog)
        vlog.afficher_popup("Résultat du traitement")
    """

    def __init__(self):
        self._dict_messages: Dict[str, List[Tuple[str, List[str]]]] = {}

        # Définition des styles disponibles pour le widget Tkinter
        self._styles: Dict[str, Dict] = {
            "normal": {"font": ("TkDefaultFont", 10)},
            "gras": {"font": ("TkDefaultFont", 10, "bold")},
            "italique": {"font": ("TkDefaultFont", 10, "italic")},
            "souligne": {"underline": True},
            "rouge": {"foreground": "#cc0000"},
            "rouge clair": {"foreground": "#ff6666"},
            "bleu": {"foreground": "#0000cc"},
            "bleu clair": {"foreground": "#66b3ff"},
            "vert": {"foreground": "#009933"},
            "vert clair": {"foreground": "#66ff99"},
            "jaune": {"foreground": "#e6b800"},
            "orange": {"foreground": "#ff9933"},
        }

    def ajouter_message(self, categorie: str, texte: str, style: Union[str, List[str]] = "normal") -> None:
        """
        Ajoute un message dans une catégorie avec un style donné.

        :param categorie: Nom de la catégorie (ex. : "Erreurs", "Exclusions")
        :param texte: Contenu du message
        :param style: Style (ou liste de styles) appliqué au texte

        :Example:

        >>> # Cas avec un seul style
        >>> vlog.ajouter_message("CSV traités", "fichier_X.csv", style=["vert"])

        >>> # Cas avec styles combinés
        >>> vlog.ajouter_message("Traitement", "Le fichier a été ajouté avec succès", style=["gras", "vert"])


        """
        if isinstance(style, str):
            style = [style]
        if categorie not in self._dict_messages:
            self._dict_messages[categorie] = []
        self._dict_messages[categorie].append((texte, style))

    def reinitialiser_messages(self) -> None:
        """
        Vide complètement les messages enregistrés.
        """
        self._dict_messages.clear()

    def log_erreur(self, message: str, continuer: bool = False) -> None:
        """
        Affiche une erreur enrichie en console et l'enregistre dans les messages.

        :param message: Le message d'erreur
        :param continuer: Si False, appelle `exit()` après affichage

        :Example:

        >>> vlog.log_erreur("Une erreur est survenue.")
        """
        #import colorama
        #from colorama import Fore, Style
        #colorama.init()
        colorama.init(autoreset=True)

        stack = inspect.stack()
        frame = stack[1].frame
        nom_fonction = stack[1].function
        cls_name = None
        if 'self' in frame.f_locals:
            cls_name = type(frame.f_locals['self']).__name__
        location = f"{cls_name + '.' if cls_name else ''}{nom_fonction}()"

        prefix = f"{Fore.YELLOW}❌ "
        suffix = " → On ignore la règle et on continue." if continuer else ""
        message_console = f"{prefix}Erreur dans {location} : {message}{suffix}{Style.RESET_ALL}"
        print(message_console)

        # Ajout dans le logger visuel
        self.ajouter_message("Erreurs", f"{location} : {message}", style=["gras", "rouge clair"])

        if not continuer:
            print("=== exit() ===")
            exit()

    def copier_dans_presse_papiers(self) -> None:
        """
        Copie le contenu du log (version texte brut) dans le presse-papiers.
        """
        r = tk.Tk()
        r.withdraw()
        r.clipboard_clear()
        r.clipboard_append(str(self))
        r.update()
        r.destroy()

    def sauvegarder_vers_fichier(self, chemin: Optional[str] = None) -> None:
        """
        Sauvegarde le contenu du log dans un fichier texte.

        :param chemin: Chemin du fichier. Si None, un nom par défaut est généré.
        """
        if chemin is None:
            timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            chemin = f"log_{timestamp}.txt"
        with open(chemin, "w", encoding="utf-8") as f:
            f.write(str(self))

    def afficher_popup(self, titre: str = "Information", reinitialiser_messages=True) -> None:
        """
        Affiche une popup enrichie avec les messages du log.

        :param titre: Titre de la fenêtre popup

        :Example:

        >>> vlog.afficher_popup("Rapport d'import")
        """
        popup = tk.Toplevel()
        popup.title(titre)
        popup.resizable(True, True)

        # Calcul de la largeur optimale selon le message le plus long
        largeur_min = 60
        largeur_max = 400
        largeur_calculee = largeur_min

        for categorie, messages in self.dict_messages.items():
            for texte, _ in messages:
                ligne = f"{categorie} :\n\t• {texte}"
                largeur_calculee = max(largeur_calculee, len(ligne))

        # Ajuster pour ne pas dépasser une largeur raisonnable
        print(largeur_calculee)
        largeur_calculee = min(largeur_calculee, largeur_max)
        print(largeur_calculee)

        text_widget = tk.Text(popup, wrap="word", height=25, width=largeur_calculee)
        text_widget.pack(expand=True, fill="both", padx=10, pady=10)

        # Configurer tous les styles
        for style_nom, style_conf in self._styles.items():
            text_widget.tag_configure(style_nom, **style_conf)

        # Insérer les messages formatés
        for categorie, messages in self._dict_messages.items():
            text_widget.insert("end", f"\n{categorie} :\n", ("gras", "souligne"))
            for texte, styles in messages:
                if not styles:
                    styles = ["normal"]
                text_widget.insert("end", "\t• ", tuple(styles))
                text_widget.insert("end", texte + "\n", tuple(styles))

        text_widget.config(state="disabled")

        bouton_frame = tk.Frame(popup)
        bouton_frame.pack(pady=(0, 10))

        tk.Button(bouton_frame, text="Copier dans presse-papiers", command=self.copier_dans_presse_papiers).pack(side="left", padx=5)
        tk.Button(bouton_frame, text="Fermer", command=popup.destroy).pack(side="right", padx=5)

        # Centrer la popup
        popup.update_idletasks()
        w = popup.winfo_width()
        h = popup.winfo_height()
        x = (popup.winfo_screenwidth() // 2) - (w // 2)
        y = (popup.winfo_screenheight() // 2) - (h // 2)
        popup.geometry(f"+{x}+{y}")

        # Si demandé, on réinitialise les messages
        if reinitialiser_messages:
            self.reinitialiser_messages()

    @property
    def dict_messages(self) -> Dict[str, List[Tuple[str, List[str]]]]:
        """
        Retourne le dictionnaire complet des messages.
        """
        return self._dict_messages

    def __str__(self) -> str:
        """
        Affichage texte brut pour la console, structuré par catégorie.
        """
        sortie = ""
        for categorie, messages in self._dict_messages.items():
            sortie += f"{categorie} :\n"
            for texte, _ in messages:
                sortie += f"\t• {texte}\n"
        return sortie

def nettoyer_nom_colonne(nom:str) -> str:
    # Supprime les retours à la ligne, espaces au début/fin et caractères spéciaux invisibles
    nom = nom.replace('\n', ' ').replace('\r', ' ')
    nom = re.sub(r'\s+', ' ', nom)  # remplace plusieurs espaces par un seul
    nom = nom.strip()
    return nom






### --------------------------------------------------------------------
#  Définitions classes et fonctions spécifiques INSTN
### --------------------------------------------------------------------

@dataclass
class infosExportsIRIS:
    repertoire: str|None
    nom_fichier: str|None
    chemin_fichier: str|None
    nom_onglet: str|None
    nbLignes_avantET: int|None

class PropExportIRIS:
    """C'est une fabrique. Contient toutes les propriétés des fichiers Excel issus des Exports IRIS (structure de configuration d'un export IRIS)"""
    # === Constructeurs ===
    def __init__(self, nom_typeExport:str, codeExport:str, repertoire_input:str, nom_onglet_input:str, nbLignes_avantET_input:str, repertoire_modele:str, nom_fichier_modele:str, repertoire_output:str, nom_fichier_output:str):
        # Type d'export
        self._nom_typeExport = nom_typeExport
        self._codeExport = codeExport

        # Informations input données
        self._input = infosExportsIRIS(
            repertoire=repertoire_input,
            nom_fichier=None,
            chemin_fichier=None,
            nom_onglet=nom_onglet_input,
            nbLignes_avantET=nbLignes_avantET_input
            )

        # Informations sur le modèle Excel à employer pour remplir l'output
        self._modele = infosExportsIRIS(
            repertoire=repertoire_modele,
            nom_fichier=nom_fichier_modele,
            chemin_fichier=os.path.join(repertoire_modele, nom_fichier_modele),
            nom_onglet=nom_typeExport,
            nbLignes_avantET=None
            )

        # Informations output
        self._output = infosExportsIRIS(
            repertoire=repertoire_output,
            nom_fichier=nom_fichier_output,
            chemin_fichier=None,
            nom_onglet=nom_typeExport,
            nbLignes_avantET=None
            )

    # === Affichage ===
    def __str__(self):
        def afficher_infos(nom_section, obj):
            lignes = [f"  {nom_section} :"]
            for champ, valeur in vars(obj).items():
                if valeur is not None:
                    # Appliquer gris clair uniquement sur les détails
                    lignes.append(f"    {Style.DIM}{Fore.LIGHTWHITE_EX}{champ:<17}: {valeur}{Style.RESET_ALL}")
            return "\n".join(lignes)

        type_export_str = f"  Type export : {self._nom_typeExport} ({self._codeExport})"

        return (
            f"PropExportIRIS\n"
            f"{type_export_str}\n"
            f"{afficher_infos('Input', self._input)}\n"
            f"{afficher_infos('Modèle', self._modele)}\n"
            f"{afficher_infos('Output', self._output)}"
        )

class TravauxFichiersIRIS:
    "C'est la classe qui contient l'environnement pour bosser sur des fichiers Exports IRIS"
    def __init__(self, prop:PropExportIRIS, chemins_fichiersInput:str|tuple[str, ...]=None):
        
        # Type d'export
        self._nom_typeExport = prop._nom_typeExport
        self._codeExport = prop._codeExport

        # Informations génériques sur les exports, modèles et output (dépend du type d'export)
        self._input = prop._input  # Informations input données
        self._modele = prop._modele  # Informations sur le modèle Excel à employer pour remplir l'output
        self._output = prop._output  # Informations output

        self._df_tableau = None
        self._df_chemins = None

        # S'il n'y a pas de chemin_fichiersInput de donné, c'est qu'il faut les sélectionner manuellement
        if not chemins_fichiersInput:
            self.choisirFichiers_filedialog()
        else:
            self._chemins_fichiersInput = chemins_fichiersInput

        # self._chemins_fichiersInput doit être un tuple de strings. Si c'est un string c'est qu'un seul fichier a été donné. Alors on convertit en tuple
        if isinstance(self._chemins_fichiersInput, str):
            self._chemins_fichiersInput = (chemins_fichiersInput,)


    @classmethod
    def avecLecture(cls, propExportIRIS:PropExportIRIS, chemins_fichiersInput:str|tuple[str, ...]=None):
        instance = cls(prop=propExportIRIS, chemins_fichiersInput=chemins_fichiersInput)
        # On lit le/les extract IRIS et on stocke dans self.__df_tableau
        instance.lire_extractIRIS()
        return instance

    @classmethod
    def avecEcritureOutputDefaut(cls, propExportIRIS:PropExportIRIS, chemins_fichiersInput:str|tuple[str, ...]=None):
        
        instance = cls.avecLecture(propExportIRIS, chemins_fichiersInput)

        # Création des paramètres pour l'ouverture du modèle
        chemin_fichier = instance._modele.chemin_fichier
        #nom_onglet = instance._output.nom_onglet

        # Création des paramètres pour l'output
        chemin_fichier_output = os.path.join(instance._output.repertoire, instance._output.nom_fichier[:-5] + "-" + date.today().strftime("%Y.%m.%d") + ".xlsx") #ou f"{datetime.now():%Y.%m.%d}")

        # On ouvre le modèle et tous ses tableaux structurés
        fe_modele = FichierExcel.depuis_fichier(chemin_fichier=chemin_fichier)

        # On copie le DataFrame avec les nouvelles données dans le modèle
        fe_modele._tableaux[instance._nom_typeExport].ecrit_dataFrame_dans_tableauStructure(instance._df_tableau, supprimeDonneesEtRemplace=True)
        
        # On écrit les références des fichiers copiés dans le tableau structuré "Imports"
        instance._df_chemins = pd.DataFrame(instance._chemins_fichiersInput, columns=['Chemin fichier'])
        fe_modele._tableaux["Imports"].ecrit_dataFrame_dans_tableauStructure(df=instance._df_chemins, supprimeDonneesEtRemplace=True)
        
        #On enregistre et on ferme
        fe_modele.save(chemin_fichier_output)

        #On fermele workbook    
        fe_modele.close()

        return instance
 
     # === Méthodes ===
    def choisirFichiers_filedialog(self):
        # Lister/sélectionner les documents à concaténer
        cheminsExcel = filedialog.askopenfilenames(title="Sélectionner les fichiers " + self._nom_typeExport + " (" + self._codeExport + ") Excel à concaténer", filetype=[("Fichiers Excel", "*.xlsx")], initialdir=self._input.repertoire)
        
        # Gestion du cas où il y a non-sélection de fichiers
        if not cheminsExcel:
            log_erreur("click sur cancel du filedialog → Pas de chemins de fichier")
        self._chemins_fichiersInput = cheminsExcel
    
    def lire_extractIRIS(self):
        """
        Crée le DataFrame pour l'export IRIS. On le stocke dans self.__df_tableau
        On selectionne la bonne methode en fonction du type d'export

        :Example:
        >>> self.lire_extractIRIS()


        .. seealso:: Rien du tout.
        .. warning:: Rien du tout.
        .. note:: Rien du tout.
        .. todo:: Rien du tout.
        """

        # Initialisation : on crée un DataFrame vide pour recevoir (peut-être) des infos que l'on traitera et qui nécessitera d'adjoindre des colonnes à self.__df_tableau
        df_colonnes_sup = None

        # On parcourt le tuple des fichiers à lire
        df_list = [] # Liste des DataFrame qui contiendra chaque fichier Excel séparément
        taille_totale = sum(os.path.getsize(fichier) for fichier in self._chemins_fichiersInput) # Calcul taille totale pour barre de progression

        with tqdm(total=taille_totale, unit='o', unit_scale=True, desc=Fore.CYAN+"Lecture des fichiers Excel" + Style.RESET_ALL) as pbar:
            #for i, ifichier in enumerate((os.path.basename(chemin) for chemin in self._chemins_fichiersInput), 1):
            for i, chemin in enumerate(self._chemins_fichiersInput, 1):
                # Données pour tqdm
                fichier = os.path.basename(chemin)
                taille = os.path.getsize(chemin)  
                pbar.set_postfix(file=fichier, progress=f"{i}/{len(self._chemins_fichiersInput)}")  # Affichage dynamique dans la barre
                
                df = pd.read_excel(chemin, skiprows=self._input.nbLignes_avantET)
                df_list.append(df)  # On ajoute le DataFrame à notre liste de DataFrame
                
                # Mise à jour de la barre avec la taille du fichier
                pbar.update(taille)

        # Concaténation finale (note : toute la fin de la méthode se fait quasi-instantanément)
        self._df_tableau = pd.concat(df_list, ignore_index=True) 

        # Selon le type d'export à traiter, on va faire des traitements spécifiques (extraction d'info des colonnes référence formation ou n° Iris)
        match self._codeExport:
            # Cas Sessions ou Inscriptions ou Ventes (sensiblement comme 'Inscription R04500' mais groupé par Client (pas de détail de chaque stagiaire))
            case "R04110" | "R04500" | "R04301":
                # On extrait / retravaille les informations de la colonne 'N° Session'
                df_colonnes_sup = self._df_tableau['N° Session'].apply(self.extraire_infos_numSessionIRIS)
                # A cause de certains éléments None ou NaN, Pandas type la colonne en float. Je la retype en Int64 qui permet de stocker des NaN avec des entiers, contrairement au type int standard.
                df_colonnes_sup['Année'] = df_colonnes_sup['Année'].astype('Int64')
                #print(df_colonnes_sup)

            # Cas Formations
            case "R0304":
                # On extrait / retravaille les informations de la colonne 'Référence'
                df_colonnes_sup = self._df_tableau['Référence'].apply(self.extraire_infos_referenceFormationIRIS)
                # A cause de certains éléments None ou NaN, Pandas type la colonne en float. Je la retype en Int64 qui permet de stocker des NaN avec des entiers, contrairement au type int standard.
                df_colonnes_sup['Année'] = df_colonnes_sup['Année'].astype('Int64')
                #print(df_colonnes_sup)

        # Ajouter les colonnes supplémentaires au DataFrame principal ssi le DataFrame df_colonnes_sup exite
        if df_colonnes_sup is not None:
            self._df_tableau = pd.concat([self._df_tableau, df_colonnes_sup], axis=1)
    
    def extraire_infos_numSessionIRIS(self, reference:str):
        """
        Fonction pour extraire les colonnes à partir de la colonne 'Référence'. Je dois faire une fonction interne car j'emploie Split qui ne s'applique que sur des string. Je dois donc faire appel à cette fonction ligne par ligne et donc créer une fonction que j'appelle par DataFrame[colonne].apply().

        Exemples de cas à traiter  :
        #S-04934-FI1516-1512-GI_VBE_GBO
        #S-05251-FI1516-1510-AMS-LCH-CLE
        #S-05246-F1516-1510-OPE-HGR-NNO
        #S-04178-FC15-604-SES-CCO

        ANCIENNE METHODE DE TRAITEMENT
        self.__df_tableau['Numéro IRIS'] = self.__df_tableau['N° Session'].astype(str).str[2:7]
        self.__df_tableau['Type formation'] = self.__df_tableau['N° Session'].astype(str).str[8:10]
        self.__df_tableau['Trigramme AF'] = self.__df_tableau['N° Session'].astype(str).str[-3:] #tout sauf 3 derniers caract
        self.__df_tableau['Trigramme RP'] = self.__df_tableau['N° Session'].astype(str).str[-7:-4] #De -7 à -4
        self.__df_tableau['Trigramme formation'] = self.__df_tableau['N° Session'].astype(str).str[-11:-8]        
        """
        # Vérifier que la référence est une chaîne de caractères
        if not isinstance(reference, str):
            #print("Problème : la référence n'est pas une instance : ")
            #print(reference)
            return pd.Series({
                'Trigramme formation': None,
                'Code IRIS': None,
                'Type de formation': None,
                'Année': None,
                'Trigramme RP': None,
                'Trigramme AF': None,
                '3ème élément de la référence': None
            })

        blocs = reference.split('-')

        # Sécurité : vérifier qu'on n'a pas plus de 8 blocs
        if len(blocs) > 8:
            print("Problème : il y a plus de 8 blocs : ")
            print(reference)
            return pd.Series({
                'Trigramme formation': None,
                'Code IRIS': None,
                'Type de formation': None,
                'Année': None,
                'Trigramme RP': None,
                'Trigramme AF': None,
                '3ème élément de la référence': None
            })

        # Traitement commun aux cas 6, 7 et 8 blocs
        # le blocs[0] c'est "S" ça sert à rien
        code_IRIS = blocs[1] if len(blocs) >= 2 else None
        #print(reference)
        if len(blocs) > 2:
            type_formation = blocs[2][:2]
            annee_match = re.search(r'\d+', blocs[2][2:])
            annee = int(annee_match.group()) if annee_match else None
        else :
            type_formation = None
            annee = None

        if len(blocs) > 5:
            trigramme_AF = blocs[-1].rstrip('_')
            trigramme_RP = blocs[-2].rstrip('_')
            trigramme = blocs[-3].rstrip('_')
        else:
            trigramme_RP = None
            trigramme_AF = None
            trigramme = None

        # 3e élément uniquement si on a 7 ou 8 blocs
        troisieme_bloc = '-'.join(blocs[3:-3]) if len(blocs) in [7, 8] else None
        
        return pd.Series({
            'Trigramme formation': trigramme,
            'Code IRIS': code_IRIS,
            'Type de formation': type_formation,
            'Année': annee,
            'Trigramme RP': trigramme_RP,
            'Trigramme AF': trigramme_AF,
            '3ème élément de la référence': troisieme_bloc
        })

    def extraire_infos_referenceFormationIRIS(self, reference):
        """
        Fonction pour extraire les colonnes à partir de la colonne 'Référence'. Je dois faire une fonction interne car j'emploie Split qui ne s'applique que sur des string. Je dois donc faire appel à cette fonction ligne par ligne et donc créer une fonction que j'appelle par DataFrame[colonne].apply().
        """
        
        # Vérifier que la référence est une chaîne de caractères
        if not isinstance(reference, str):
            #print("Problème : la référence n'est pas une instance : ")
            #print(reference)
            return pd.Series({
                'Trigramme formation': None,
                'Type de formation': None,
                'Année': None,
                'Unité de formation': None,
                '3ème élément de la référence': None
            })

        blocs = reference.split('-')

        # Sécurité : vérifier qu'on a au moins 4 blocs
        if len(blocs) < 3:
            print("Problème : il y a moins de 3 blocs : ")
            print(reference)
            return pd.Series({
                'Trigramme formation': None,
                'Type de formation': None,
                'Année': None,
                'Unité de formation': None,
                '3ème élément de la référence': None
            })

        # Traitement commun aux cas 3 et 4+ blocs
        type_formation = blocs[0][:2]
        annee_match = re.search(r'\d+', blocs[0][2:])
        annee = int(annee_match.group()) if annee_match else None

        trigramme = blocs[1].rstrip('_')
        unite_formation = blocs[-1].rstrip('_')

        # 3e élément uniquement si on a 4 blocs ou plus
        troisieme_bloc = '-'.join(blocs[2:-1]) if len(blocs) > 3 else None
        
        return pd.Series({
            'Trigramme formation': trigramme,
            'Type de formation': type_formation,
            'Année': annee,
            'Unité de formation': unite_formation,
            '3ème élément de la référence': troisieme_bloc
        })
   
   
    # === Affichage ===
    def __str__(self):
        if self._chemins_fichiersInput:
            # Pour aff_repertoire et chemin 
            aff_cheminsFichiers = ""
            for cfichier in self._chemins_fichiersInput:
                aff_cheminsFichiers = aff_cheminsFichiers + "\n    " + cfichier
        else:
            aff_cheminsFichiers = "Aucun fichier spécifié"

        # Pour aff_df
        from io import StringIO
        buffer = StringIO()
        if self._df_tableau is not None:
            print(self._df_tableau, file=buffer)
            aff_df = buffer.getvalue()
        else:
            aff_df = "Non défini"

        return (
            f"TravauxFichiersIRIS\n"
            f"  Chemins fichiers à exploiter : {aff_cheminsFichiers}\n"
            f"  df_tableau :\n{aff_df}"
            )
        
class BilanFormation:
    "C'est la classe qui contient tous les éléments de ma formation pour mon bilan"

    def __init__(self, codeFormation:str, annee:int):
        
        self._codeFormation = codeFormation
        self._annee = annee

        self._sessions_nom_typeExport = sessions._nom_typeExport
        self._sessions_codeExport = sessions._codeExport
        self._sessions_repertoire = sessions._output.repertoire

        self._chemin_specsPedagogiques = None

        self._chemin_fdc = None

        self._repertoire_fdc_defaut = rep_fdc_defaut
        self._repertoire_specsPedagogiques_defaut = rep_specsPedagogiques_defaut
        
        #####
        # Exploitation de l'extract IRIS sessions
        #####

        # J'ouvre un export session de IRIS et load tous ses tableaux structurés dans des DataFrame (inclus dans un FichierExcel)
        chemin_fichier_session = filedialog.askopenfilename(title="Sélectionner l'export " + self._sessions_nom_typeExport + " (" + self._sessions_codeExport + ") Excel à employer", filetype=[("Fichiers Excel", "*.xlsx")], initialdir=self._sessions_repertoire)
        if not chemin_fichier_session:
            log_erreur("click sur cancel du filedialog → Pas de chemin de fichier session")
        self._fe_session = FichierExcel.depuis_fichier(chemin_fichier_session)
        #print(self._fe_session)
        self._df_sessions = self._fe_session._tableaux[self._sessions_nom_typeExport]._df  # Création d'un alias

        # Pour initialiser les valeurs communes, déjà on filtre le Dataframe principal avec le code formation
        #df_filtre = self.__df_sessionsIRIS[(self.__df_sessionsIRIS['Trigramme formation'] == self.__codeFormation) & (self.__df_sessionsIRIS['Année début ses.'] == self.__annee)]
        df_sessions_filtre = self._df_sessions[(self._df_sessions['Trigramme formation'] == self._codeFormation) & (self._df_sessions['Année début ses.'] == self._annee)]
        
        
        # Puis on récupère la dernière ligne pour avoir les valeurs les plus à jour (tri par index)
        [self._titreFormation, self._dureeHeures, self._dureeJours, self._minIRIS, self._maxIRIS, self._depassementAutoriseIRIS] = df_sessions_filtre[["Session", "Durée planif. (H.)", "Durée planif. (J.)", "Min.", "Max.", "Dépass. autorisé"]].iloc[-1]

        # Pour obtenir la liste des RP et de leurs lieux
        self._lieuxFormation = ""
        for rp in df_sessions_filtre["Nom responsable pédag."].unique():
            #print("\n\nRP = " + rp)
            [prenomRP, nomRP, lieuRP] = df_sessions_filtre[self._df_sessions["Nom responsable pédag."] == rp][["Prénom responsable pédag.", "Nom responsable pédag.", "Lieu principal"]].iloc[-1]
            self._lieuxFormation += prenomRP + " " + nomRP + " (" + lieuRP + "), "
        self._lieuxFormation = self._lieuxFormation[:-2]
        #print(self._lieuxFormation)




        #####
        # Exploitation des specs pédagogiques
        #####
        self._chemin_specsPedagogiques = filedialog.askopenfilename(title="Sélectionner les dernières specs pédagogiques", filetype=[("Fichiers PDF", "*.pdf"), ("Documents Word", "*.docx")], initialdir=optimiseCheminRepertoire(self._repertoire_specsPedagogiques_defaut.replace("XXX", self._codeFormation)))
        if not chemin_fichier_session:
            log_erreur("click sur cancel du filedialog → Pas de chemin des specs pédagogiques")
        self._dateSpecs = time.localtime(os.path.getmtime(self._chemin_specsPedagogiques))
        self._sDateSpecs = f"{self._dateSpecs[2]:02}/{self._dateSpecs[1]:02}/{self._dateSpecs[0]:04}"
        #print(self._sDateSpecs)



        #####
        # Exploitation de la fiche de coûts
        #####
        self._chemin_fdc = filedialog.askopenfilename(title="Sélectionner la dernière fiche de coûts", filetype=[("Fichiers Excel", "*.xlsx")], initialdir=optimiseCheminRepertoire(self._repertoire_fdc_defaut.replace("XXX", self._codeFormation)))
        if not chemin_fichier_session:
            log_erreur("click sur cancel du filedialog → Pas de chemin de fiche de coûts")
        self._dateFdC = time.localtime(os.path.getmtime(self._chemin_fdc))
        self._sDateFdC = f"{self._dateFdC[2]:02}/{self._dateFdC[1]:02}/{self._dateFdC[0]:04}"
        
        self._df_fdc_infos, self._df_fdc_couts, self._prixVenteRetenuParParticipant, self._dateCreationFormation, self._dureeJours_fdc, self._osThematique, self._nbCible_fcd = lire_fdc(self._chemin_fdc)
        




        #####
        # Exploitation de l'extract IRIS formation TODO
        #####
        #Date création formation
        #Type de reconnaissance (Autre, Certification, Diplôme)
        #Formation habilitante (bool)
        #Reconnu au RNCP
        #Eligible CPF
        #Reconnue au RS


    ### --------------------------------------------------------------------
    #  Méthodes de la classe
    ### --------------------------------------------------------------------
    def mergeBilan(self, chemin_word_bilan_input, chemin_word_bilan_output):
        """
        A partir d'un chemin de fichier word avec des champs de fusion, on crée le bilan de formation final en incluant les données à l'intérieur.
        Le fichier output est défini par l'utilisateur

        :param s_word_bilan_input: Chemin du fichier Word contenant les champs de fusion et a employer
        :type s_word_bilan_input: string
        :param s_word_bilan_output: Chemin du fichier Word apres fusion des donnees
        :type s_word_bilan_output: string
        :return: pas de donnee en retour
        :rtype: none

        :Example:

        >>> self.mergeBilan("C:\\Users\\wordIn.docx", "C:\\Users\\wordOut.docx")


        .. seealso:: Rien du tout.
        .. warning:: Rien du tout.
        .. note:: Rien du tout.
        .. todo:: Rien du tout.
        """

        #print(f"BilanFormation lancé avec : trigramme={self._codeFormation}, année={self._annee}")
        document = MailMerge(chemin_word_bilan_input)
        #print(document.get_merge_fields())

        document.merge(
            annee='{:%Y}'.format(date.today()),
            codeFormation=self._codeFormation,
            titreFormation=self._titreFormation,
            
            lienGED = r'file:///\\\\instnt\\PARTAGE\\FORMATIONS_C\\ACI\\',
            osThematique = self._osThematique,
            lieuxFormation = self._lieuxFormation,
            dureeJours = f"{self._dureeJours:.1f}",
            dureeHeures = f"{self._dureeHeures:.2f}",
            dateCreationFormation = str(self._dateCreationFormation), #Actuellement pris depuis fdc, réaffecter à partir extract IRIS formations

            dateSpecs = self._sDateSpecs,

            dateFdC = self._sDateFdC,
            minFdC = f"{self._df_fdc_couts.loc['Valeur fixée', 'Min participants T3']:d} p.",
            cibleFdC = f"{self._nbCible_fcd:d} p.",
            minIRIS = f"{self._minIRIS:d} p.",
            cibleIRIS = f"{self._maxIRIS:d} p.",
            maxIRIS = f"{self._depassementAutoriseIRIS:d} p.",

            PT1_1 = f"{self._df_fdc_couts.loc['T1', 'Montant cible par participant']:.0f} €/p.",
            PT1_2 = f"{self._df_fdc_couts.loc['T1', 'Montant cible par participant et par jour']:.0f} €/j/p.",
            PT1_3 = f"{self._df_fdc_couts.loc['T1', 'Min participants T1']:d} p.",

            PT3_1 = f"{self._df_fdc_couts.loc['1.1xT3', 'Montant cible par participant']:.0f} €/p.",
            PT3_2 = f"{self._df_fdc_couts.loc['1.1xT3', 'Montant cible par participant et par jour']:.0f} €/j/p.",
            PT3_3 = f"{self._df_fdc_couts.loc['1.1xT3', 'Min participants T1']:d} p.",
            PT3_4 = f"{self._df_fdc_couts.loc['1.1xT3', 'Min participants T3']:d} p.",

            PTR_1 = f"{self._df_fdc_couts.loc['Valeur fixée', 'Montant cible par participant']:.0f} €/p.",
            PTR_2 = f"{self._df_fdc_couts.loc['Valeur fixée', 'Montant cible par participant et par jour']:.0f} €/j/p.",
            PTR_3 = f"{self._df_fdc_couts.loc['Valeur fixée', 'Min participants T1']:d} p.",
            PTR_4 = f"{self._df_fdc_couts.loc['Valeur fixée', 'Min participants T3']:d} p."
            
            )
        
        document.write(chemin_word_bilan_output)

class Traiter_evalStat:
    # Obtenir :
    #   - Taux de recommandation en 2025 :	100%
    #   - Satisfaction sur l’année 2025 :	4,75 sur 5
    #   - Taux de retour (facultatif)

    # Enlever ligne total avant traitement
    # Remise ligne total après fin traitement

    # Aller chercher CSV (sélectionner le fichier à partir du chemin standard)
    # chemin standard vers CSV = "\\Instnt\partage\FORMATIONS_C\"&codeFormation&"\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\CSV\"

    # Importer le CSV dans le modèle Excel (onglet "CSV")
    # Ecrire nb d'inscrits (vient de sessions) + code formation et code IRIS + changer chemin CSV → Voir si je ne peux pas urtiliser des plages nommées
    # coller infos dans tableau travaillé
    #   - depuis colonne H jusqu'à la fin (AG) on copie et on va transposer
    #   - pour toutes les 2 colonnes ça fait critère puis commentaire
    #   - sauf pour "Recommanderiez-vous cette formation ?" (AD) + "Commentaires, remarques, suggestions " (AH)
    #   - transformer valeurs de Recommanderiez-vous cette formation ? et Avez-vous d'autres besoins de formation ? en oui = 5 et non = 0
    # Actualiser les TCD
    # Sauvegarder au bon endroit
    # TODO : pour l'instant à cause de pandas, je casse les segments

    # Possibilité d'importer plusieurs CSV pour une même année
    # Attention : impact sur nombre total de stagiaire (sommer dans sessions)
    
    def __init__(self):
        
        self._chemin_csv_stagiaires:str|None = None  # Fichier csv EvalStat stagiaire individuel
        self._chemin_excel_stagiaires_output:str|None = None  # Fichier xlsx EvalStat stagiaire individuel qu'on va créer à partir du CSV
        self._fe_stagiaires:FichierExcel|None = None  # Objet contenant les données EvalStat stagiaire individuel

        self._chemin_modeleExcel_stagiaires:str|None = None  # Modèle Excel dans lequel importer le CSV
        
        self._chemin_excel_sessions:str|None = None  # Fichier Excel (extract IRIS traité) dans lequel on a les informations des sessions (permet de compélter les CSV)
        self._fe_sessions:FichierExcel|None = None  # Objet contenant le modèle Excel

        self._trigrammeFormation:str|None = None
        self._codeIRIS:int|None = None

        self._chemins_csv_traites:List[str] = []
        self._chemins_csv_exclus:List[str] = []

        # Valeurs actuellement écrites en dur
        self._chemin_excel_evaluations_defaut = r'\\instnt\partage\FORMATIONS_C\###\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\Evaluation-Stagiaires-Global-###.xlsx'
        self._chemin_modeleExcel_stagiaires = r"C:\Users\vt238770\Documents\_CEA\Prog\Modèles\Evaluation-Stagiaires-Modèle.xlsx"
        self._chemin_excel_sessions = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts complets\R04110_Sessions-COMPLET-2025.08.24.xlsx" 
        #instance._chemin_excel_stagiaires_output = r"C:\Users\vt238770\Documents\_CEA\Prog\Modèles\Evaluation-Stagiaires-testOut.xlsx"

        # === Colonnes du CSV selon traitement à avoir ===
        # Colonnes descriptives à recopier
        self._colonnes_csv_fixes = [
            "Chemin fichier CSV", "Prénom", "Nom", "Entreprise", "Code session"]

        # Colonnes avec note/commentaire en binôme
        self._colonnes_csv_avec_commentaires = [
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
        self._colonnes_csv_commentaires_seuls = [
            "Comment avez-vous connu cette formation ?",
            "Commentaires, remarques, suggestions"]

        # Colonnes note seule (il se trouve que je vais aussi devoir convertir le booléen)
        self._colonnes_csv_bool = [
            "Recommanderiez-vous cette formation ?"]
        
        
        # === Colonnes de l'extract IRIS Sessions à récupérer ===
        self._colonnes_sessions = [
            "N° Session",
            "Formation",
            "Trigramme formation",
            "Code IRIS",
            "Date début ses.",
            "Année début ses.",
            "Type de formation",
            "Trigramme RP",
            "Trigramme AF",
            "Nb. Présents"]
        
    @classmethod
    def depuis_chemin_csv_stagiaires(cls, chemin_csv_stagiaires:str, fe_sessions:FichierExcel=None, ouvrirDossier:Boolean=False, remplace_df:Boolean=False) :
       
       # On créée l'instance
        instance = Traiter_evalStat()
        instance._chemin_csv_stagiaires = chemin_csv_stagiaires
        if fe_sessions is not None:
            instance._fe_sessions = fe_sessions
            # Quand je ferai la jointure plus tard sur "Code IRIS", il faudra que ce soit avec des strings
            instance._fe_sessions._tableaux["Sessions"]._df["Code IRIS"] = instance._fe_sessions._tableaux["Sessions"]._df["Code IRIS"].astype(str)

        # On récupère le trigramme de la formation depuis le chemin du CSV
        instance._trigrammeFormation = instance.recupere_trig_formation_depuis_chemin(chemin_csv_stagiaires)
        
        # Pour le nom de l'Excel output : on reprend le nom du csv et on remplace par xlsx
        instance._chemin_excel_stagiaires_output = os.path.join(os.path.dirname(instance._chemin_csv_stagiaires), os.path.basename(instance._chemin_csv_stagiaires).replace(".csv", ".xlsx"))

        # On génère le fichier Excel du CSV à partir du modèle
        instance._construit_FichierExcel_depuis_CSV(fe_sessions = fe_sessions, remplace_df=remplace_df)

        # Ouverture du dossier à la fin
        if ouvrirDossier:
            ouvrir_dossier(os.path.dirname(instance._chemin_excel_stagiaires_output))

        # On ajoute le chemin au tuple des éléments traités
        instance._chemins_csv_traites.append(chemin_csv_stagiaires)

        return instance        

    @classmethod
    def depuis_tuple_csv_stagiaires(cls, tuple_csv_stagiaires:Tuple(str), chemin_excel_evaluations_defaut:str=None, chemin_modeleExcel_stagiaires:str=None, chemin_excel_sessions:str=None, fe_sessions:FichierExcel=None, ouvrirDossier:Boolean=False) -> Traiter_evalStat:
        

        
        # On crée l'instance et on complète les infos avec les valeurs facultatives
        instance = Traiter_evalStat()
        instance._chemin_excel_evaluations_defaut = chemin_excel_evaluations_defaut
        instance._chemin_modeleExcel_stagiaires = chemin_modeleExcel_stagiaires
        instance._chemin_excel_sessions = chemin_excel_sessions
        instance._fe_sessions = fe_sessions
        
        df_formation_csv = None
        df_formation_stagiaires = None
        #dico_sessionsDejaTraitees = None
        timer = Timer()

        # On récupère les données de Sessions (on en aura besoin plus tard)
        if instance._fe_sessions is None:
            timer.debut("Lecture fichier session")
            instance._fe_sessions = FichierExcel.depuis_fichier(instance._chemin_excel_sessions)
            timer.fin()
        

        # On convertit le tuple de strings en dictionnaire avec les trigrammes formation en clef
        dico_chemins_csv_session = defaultdict(list)  #Dictionnaire spécial : lorsqu’on accède à une clé qui n’existe pas encore, il va automatiquement créer une nouvelle entrée avec une valeur par défaut, ici une liste vide (list())
        for chemin in tuple_csv_stagiaires:
            trigramme = instance.recupere_trig_formation_depuis_chemin(chemin)
            dico_chemins_csv_session[trigramme].append(chemin)
        dico_chemins_csv_session = dict(dico_chemins_csv_session)  # Optionnel : conversion en dict normal


        # Pour chaque trigramme on va traiter chaque session et soit créer soit append le fichier excel global de la formation
        for trigramme, chemins_csv_session in dico_chemins_csv_session.items(): #chemins est la liste des chemins des évaluations pour chaque sessions de ce trigramme formation
            print(f"\n\n{Style.BRIGHT}{Fore.RED}Gestion des formations {trigramme}")

            # On crée le chemin vers les évaluations de la formation (le fichier qui va concaténer toutes les évaluation d'une formation)
            chemin_excel_evaluations_formation = instance._chemin_excel_evaluations_defaut.replace("###", trigramme)

            # On vérifie que le répertoire dédié existe sinon on le créée : \\instnt\partage\FORMATIONS_C\###\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations
            os.makedirs(os.path.dirname(chemin_excel_evaluations_formation), exist_ok=True)
            
            # On regarde si Evaluation-Stagiaires-Global-XXX.xlsx existe
            if os.path.isfile(chemin_excel_evaluations_formation):
                # Alors on l'ouvre
                fe_evaluations_formation = FichierExcel.depuis_fichier(chemin_fichier=chemin_excel_evaluations_formation)
                fe_evaluations_formation._tableaux["CSV_stagiaires"].charge_df()
                df_formation_stagiaires = fe_evaluations_formation._tableaux["Stagiaires"]._df  # Alias

                # On récupère la liste des sessions déjà intégrées (liste des codes et des chemins) → Ce sera pour écrire dans le tkinter
                #dico_sessionsDejaTraitees = dict(
                #df_formation_stagiaires[df_formation_stagiaires["Trigramme formation"] == trigramme]     # 1. filtre sur le trigramme
                #.drop_duplicates(subset=["Code IRIS", "Chemin fichier CSV"])  # 2. élimine les doublons
                #[["Code IRIS", "Chemin fichier CSV"]]            # 3. sélection des colonnes
                #.values                               # 4. valeurs du DF
                #    )        

                # Il ne faudra pas supprimer les anciennes données de fe_evaluations_formation
                supprimeDonneesEtRemplace = False  
            else :
                # On créée le fichier : on ouvre le modèle (on sauvera avec le bon nom à la fin)
                fe_evaluations_formation = FichierExcel.depuis_modele(chemin_modele=instance._chemin_modeleExcel_stagiaires, chemin_fichier_sauv=chemin_excel_evaluations_formation)

                # Il faudra supprimer les anciennes données de fe_evaluations_formation
                supprimeDonneesEtRemplace = True

                vlog.ajouter_message("Création EvalStat Global formation", fe_evaluations_formation.chemin_fichier, style=["vert"])

            
            # Pour chaque chemin de session, on crée le fe_stagiaire dédié de la session et on ajoute les lignes de son dataframe au dataframe de fe_evaluations_formation
            for chemin_csv_session in chemins_csv_session:
                print(f"\n{Style.BRIGHT}{Fore.YELLOW}Gestion de la session {chemin_csv_session}")
                traiterCSV = True  # Par défaut, on traite le CSV

                # On vérifie que chemin_csv_session n'est pas déjà dans le fichier session pour savoir si on l'exclue du traitement
                if df_formation_stagiaires is not None:
                    if chemin_csv_session in df_formation_stagiaires["Chemin fichier CSV"].drop_duplicates().tolist():  
                        traiterCSV = False
                        instance._chemins_csv_exclus.append(chemin_csv_session)
                        print(f"Exclusion car csv déjà dans le fichier global : {chemin_csv_session}")
                        #global message_sortie
                        #message_sortie += f"\nExclusion car csv déjà dans le fichier global : {chemin_csv_session}"
                        vlog.ajouter_message("Exclusion car csv déjà dans le fichier global", chemin_csv_session, style=["orange"])

                # Traitement du CSV
                if traiterCSV:  
                    timer.debut("Traiter_evalStat.depuis_chemin_csv_stagiaires")
                    traite_csv_session = Traiter_evalStat.depuis_chemin_csv_stagiaires(chemin_csv_session, fe_sessions=instance._fe_sessions, ouvrirDossier=ouvrirDossier, remplace_df=True)

                    timer.debut("Copie des Dataframe csv et stagiaires")
                    # #Si df_formation_csv est vide, il faut l'initialiser avec le premier df sinon on concatène
                    if df_formation_csv is None:
                        df_formation_csv = traite_csv_session._fe_stagiaires._tableaux["CSV_stagiaires"]._df.copy()
                        df_formation_stagiaires = traite_csv_session._fe_stagiaires._tableaux["Stagiaires"]._df.copy()
                    else:
                        df_formation_csv = pd.concat([df_formation_csv, traite_csv_session._fe_stagiaires._tableaux["CSV_stagiaires"]._df], ignore_index=True)
                        df_formation_stagiaires = pd.concat([df_formation_stagiaires, traite_csv_session._fe_stagiaires._tableaux["Stagiaires"]._df], ignore_index=True)
                    
                    # On ajoute le chemin au tuple des éléments traités
                    instance._chemins_csv_traites.append(chemin_csv_session)
                    vlog.ajouter_message("Fichiers traités", chemin_csv_session, style=["vert"])

            # On concatène, on sauve et on ferme le fe de tous les CSV de la formation
            if instance._chemins_csv_traites :
                timer.debut("Ecriture dans les tableaux Excel + ")
                fe_evaluations_formation._tableaux["CSV_stagiaires"].ecrit_dataFrame_dans_tableauStructure(df_formation_csv, supprimeDonneesEtRemplace = supprimeDonneesEtRemplace)
                fe_evaluations_formation._tableaux["Stagiaires"].ecrit_dataFrame_dans_tableauStructure(df_formation_stagiaires, supprimeDonneesEtRemplace = supprimeDonneesEtRemplace)

                timer.debut("Sauvegarde et femerture")
                fe_evaluations_formation.save()
                fe_evaluations_formation.close()
                timer.fin()
            
                # Actualisation des TCD
                fe_evaluations_formation.actualiser_TCD()

        return instance

    # ===  (getter / setter) ===
    @property
    def chemins_csv_traites(self):
        return self._chemins_csv_traites


    ######
    #  === Méthodes internes ===
    ######

    def _construit_FichierExcel_depuis_CSV(self, fe_sessions:FichierExcel=None, remplace_df:Boolean=False):
        ######
        # === Import et traitement du CSV d'evalStat ===
        ######
        # On récupère l'encodage et on importe le CSV dans un DataFrame
        timer = Timer()
        codage_csv = trouve_encodage_csv(self._chemin_csv_stagiaires)
        print(f"\nConstruction de l'Excel pour {self._chemin_csv_stagiaires} ; codage : {codage_csv}")
        timer.debut("Import du CSV et traitement du DataFrame")
        df_csv_stagiaires = pd.read_csv(self._chemin_csv_stagiaires, sep=';', encoding=codage_csv)  # Ouverture du CSV et mise dans un DataFrame
        
        # Prise en compte qu'on a plusieurs formats de CSV : on doit traiter des colonnes en + ou - en conséquences
        if "Date de fin" in df_csv_stagiaires.columns:
            # Cas 1 (nouveau format de csv) : supprimer "Date de fin" → Test : r"P:\FORMATIONS_C\54C\P07-bilan-sessions-et-bilan-formation\rapports-sessions-evaluations\2023-06-S14317 UEM\S-14317-FC23-54C-VTE-LRA-Stagiaires.csv"
            df_csv_stagiaires = df_csv_stagiaires.drop(columns=["Date de fin"])
        else:
            # Cas 2 (ancien format de csv) : supprimer la 2e et 3e colonne (indices 1 et 2) → Test : r"P:\FORMATIONS_C\22B\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\S-17606 - 22B - 06-2025\S-17606-FC25-22B-VTE-CAR-Stagiaires.csv"
            df_csv_stagiaires = df_csv_stagiaires.drop(df_csv_stagiaires.columns[[1, 2]], axis=1)
        
        # On rajoute le chemin du CSV en première colonne
        df_csv_stagiaires.insert(0, "Chemin fichier CSV", self._chemin_csv_stagiaires)

        # Définition self._codeIRIS. Sinon Non existant, on récupère le numéro IRIS depuis le CSV (c'est la plus sur)
        if self._codeIRIS is None:
            match = re.search(r"\b\d{5}\b", self._chemin_csv_stagiaires)
            if match:
                self._codeIRIS = match.group(0)
            else:
                self.demander_code_iris()

        # Met à jour ou crée la colonne "Code session" avec self._codeIRIS
        df_csv_stagiaires["Code session"] = self._codeIRIS

        # Mise au format jj/mm/aaaa de la colonne "Date" (si elle existe)
        if "Date" in df_csv_stagiaires.columns:
            try:
                df_csv_stagiaires["Date"] = pd.to_datetime(df_csv_stagiaires["Date"], dayfirst=True, errors="coerce").dt.strftime("%d/%m/%Y")  # dayfirst=True indique que le premier nombre correspond au jour (format jj/mm/aaaa)
            except Exception as e:
                print(f"Erreur de conversion de la colonne Date : {e}")

        # A cause des espaces à la con qui trainent dans les noms des colonnes des CSV, je vais reload le dataframe depuis l'excel que je viens de créer car les colonnes du modèle sont bien nommées
        # Ainsi on sauve ici plutôt qu'à la fin et on reload le DataFrame
        self._fe_stagiaires = FichierExcel.depuis_modele(chemin_modele=self._chemin_modeleExcel_stagiaires, chemin_fichier_sauv=self._chemin_excel_stagiaires_output)
        if remplace_df:
            self._fe_stagiaires._tableaux["CSV_stagiaires"].remplace_df(df_csv_stagiaires)
        self._fe_stagiaires._tableaux["CSV_stagiaires"].ecrit_dataFrame_dans_tableauStructure(df_csv_stagiaires, supprimeDonneesEtRemplace=True)
        self._fe_stagiaires.save(self._chemin_excel_stagiaires_output)
        self._fe_stagiaires.close()
        self._fe_stagiaires = FichierExcel.depuis_fichier(self._chemin_excel_stagiaires_output)
        self._fe_stagiaires._tableaux["CSV_stagiaires"].charge_df()
        df_csv_stagiaires = self._fe_stagiaires._tableaux["CSV_stagiaires"]._df
        
        ######
        # === On crée la seconde partie du DataFrame qui sera dans l'onglet "Stagiaire" ===
        # On va découper le dataframe du CSV selon les différents critères et mettre dans un dataframe qu'on pourra exploiter par un TCD
        ######
        # Nouveau DataFrame à remplir
        df_long = []

        # Parcours des lignes
        for _, row in df_csv_stagiaires.iterrows():
            base = {col: row[col] for col in self._colonnes_csv_fixes}  # Création des colonnes qui seront répétées à chaque fois
            base["NOM Prénom"] = f"{str(row['Nom']).upper()} {row['Prénom']}".strip()  # Création du champ "NOM Prénom"

            # Cas 1 : colonnes avec note + commentaire associé
            for critere in self._colonnes_csv_avec_commentaires:
                if critere in row:
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

            # Cas 2 : colonnes texte seules
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


        ######
        # === On fait le left join entre df_stagiaires et les données qui proviennent de l'extract IRIS Sessions ===
        ######
        # On récupère les données de Sessions
        if fe_sessions is None:
            timer.debut("Lecture fichier session")
            self._fe_sessions = FichierExcel.depuis_fichier(self._chemin_excel_sessions)
            self._fe_sessions._tableaux["Sessions"]._df["Code IRIS"] = self._fe_sessions._tableaux["Sessions"]._df["Code IRIS"].astype(str)
        else:
            self._fe_sessions = fe_sessions


        # Pour faire le merge, il faut que les colonnes soient de même type (là "Code session" est de type int64 et "Code IRIS" est de type object (souvent des chaînes de caractères)).
        # Comme je ne peux être sûr que tous les "Code IRIS" issu des CSV soient bien convertibles en int (c’est-à-dire pas de chaînes vides, NaN, ou autres caractères non numériques), alors je passe par des strings
        df_stagiaires["Code session"] = df_stagiaires["Code session"].astype(str)
        #self._fe_sessions._tableaux["Sessions"]._df["Code IRIS"] = self._fe_sessions._tableaux["Sessions"]._df["Code IRIS"].astype(str)
        
        # On récupère uniquement les colonnes souhaitées dans une vue pour faciliter le codage (c'est un alias)
        df_sessions_filtre = self._fe_sessions._tableaux["Sessions"]._df[self._colonnes_sessions]
        
        # On fait la jointure entre df_stagiaires et df_sessions_filtre
        timer.debut("Création du DataFrame Stagiaires (jointure)")
        df_stagiaires = df_stagiaires.merge(
            df_sessions_filtre,
            left_on="Code session",
            right_on="Code IRIS",
            how="left"
        )

        # On réorganise les colonnes : d'abord celles de df_sessions puis celles de df_stagiaires
        colonnes_resultat = (
            df_sessions_filtre.columns.tolist() +  # colonnes de _df_sessions
            [col for col in df_stagiaires.columns if col not in df_sessions_filtre.columns]  # le reste (i.e. celles de df_stagiaires)
        )
        df_stagiaires = df_stagiaires[colonnes_resultat]
        #print(_df_stagiaires)
        
        # On vire "Code session" qui est redondante avec "Code IRIS"
        df_stagiaires.drop(columns=["Code session"], inplace=True)
        
        # On renomme les colonnes
        #df_sessions_filtre.rename(columns={"Date début ses.": "Date"}, inplace=True)
        #df_sessions_filtre.rename(columns={"Année début ses.": "Année"}, inplace=True)

        # On remplace le DataFrame existant par le nouveau
        if remplace_df:
            self._fe_stagiaires._tableaux["Stagiaires"].remplace_df(df_stagiaires)
        
        # On écrit et on sauve
        timer.debut("On écrit le DataFrame, on met à jour les TCD et on sauve")
        self._fe_stagiaires._tableaux["Stagiaires"].ecrit_dataFrame_dans_tableauStructure(df_stagiaires, supprimeDonneesEtRemplace=True)
        self._fe_stagiaires.save()


        # Màj des TCD
        self._fe_stagiaires.actualiser_TCD()
        timer.fin()

    def demander_code_iris(self):
        print("TODO : demander_code_iris appelé → Essayer de généraliser avec demander_code")
        def verifier_entree(*args):
            val = entry_code.get()
            bouton_valider.config(state="normal" if val.isdigit() and len(val) == 5 else "disabled")
            bouton_renommer.config(state="normal" if val.isdigit() and len(val) == 5 else "disabled")

        def valider():
            nonlocal code_iris
            self._codeIRIS = int(entry_code.get())
            fenetre.destroy()

        def valider_et_renommer():
            nonlocal code_iris
            self._codeIRIS = int(entry_code.get())
            nouveau_nom = f"S-{code_iris}-Stagiaires.csv"
            nouveau_chemin = os.path.join(
                os.path.dirname(self._chemin_csv_stagiaires),
                nouveau_nom
            )

            try:
                os.rename(self._chemin_csv_stagiaires, nouveau_chemin)
                self._chemin_csv_stagiaires = nouveau_chemin
            except Exception as e:
                tk.messagebox.showerror("Erreur", f"Impossible de renommer le fichier :\n{e}")
                log_erreur(f"impossible de renommer le fichier :\n{e}")
                return  # Ne pas fermer la fenêtre si erreur

            fenetre.destroy()

        def annuler():
            fenetre.destroy()
            log_erreur("code IRIS non renseignée pour " + self._chemin_csv_stagiaires + " → exit()")
            exit()

        code_iris = None

        fenetre = tk.Tk()
        fenetre.title("Code IRIS à renseigner manuellement")
        fenetre.resizable(False, False)
        fenetre.geometry("500x180")
        fenetre.eval('tk::PlaceWindow . center')

        # Fermer avec Échap
        fenetre.bind("<Escape>", lambda e: annuler())
        # Entrée = bouton Valider
        fenetre.bind("<Return>", lambda e: bouton_valider.invoke())

        label_warning = tk.Label(
            fenetre,
            text="⚠ Code IRIS non trouvé automatiquement, veuillez le spécifier manuellement",
            font=("Segoe UI", 10, "bold"),
            fg="orange"
        )
        label_warning.pack(pady=(10, 5))

        label_chemin = tk.Label(
            fenetre,
            text=f"Chemin du fichier source :\n  *{self._chemin_csv_stagiaires}*",
            font=("Segoe UI", 9),
            justify="left",
            wraplength=480
        )
        label_chemin.pack(pady=(0, 10))

        frame_saisie = tk.Frame(fenetre)
        frame_saisie.pack()

        entry_code = tk.Entry(frame_saisie, width=10, justify="center", font=("Segoe UI", 12))
        entry_code.pack()
        entry_code.focus()

        entry_code_var = tk.StringVar()
        entry_code["textvariable"] = entry_code_var
        entry_code_var.trace_add("write", verifier_entree)

        frame_boutons = tk.Frame(fenetre)
        frame_boutons.pack(pady=10)

        bouton_valider = tk.Button(frame_boutons, text="Valider", state="disabled", width=20, command=valider)
        bouton_valider.grid(row=0, column=0, padx=5)

        bouton_renommer = tk.Button(frame_boutons, text="Valider et remplacer le nom du CSV", state="disabled", width=30, command=valider_et_renommer)
        bouton_renommer.grid(row=0, column=1, padx=5)

        bouton_annuler = tk.Button(frame_boutons, text="Annuler", width=10, command=annuler)
        bouton_annuler.grid(row=0, column=2, padx=5)

        fenetre.mainloop()

        return code_iris
    
    def demander_code(self, typeCode) -> int|str:
        # Initialisation en fonction du type de code
        match typeCode:
            case "Code IRIS":
                print("Cas code IRIS")
                nbCaracteres = 5
                chemin_a_tester = self._chemin_csv_stagiaires
            case "Trigramme formation":
                print("Cas trigramme formation")
                nbCaracteres = 3
                chemin_a_tester = self._chemin_csv_stagiaires
            case _:
                log_erreur(f"cas non valide : soit 'IRIS' soit 'Trigramme formation', demandé : {typeCode} → exit()")
                exit()

            
        def verifier_entree(*args):
            val = entry_code.get()
            bouton_valider.config(state="normal" if val.isdigit() and len(val) == nbCaracteres else "disabled")
            if typeCode=="Code IRIS":
                bouton_renommer.config(state="normal" if val.isdigit() and len(val) == nbCaracteres else "disabled")

        def valider():
            nonlocal code # Adaptation : changer en "code"
            code = entry_code.get()            
            fenetre.destroy()

        # Cas IRIS uniquement : ne pas toucher / ne pas afficher si cas Trigramme formation
        def valider_et_renommer():
            nonlocal code
            code = entry_code.get()
            nouveau_nom = f"S-{code}-Stagiaires.csv"
            nouveau_chemin = os.path.join(
                os.path.dirname(self._chemin_csv_stagiaires),
                nouveau_nom
            )

            try:
                os.rename(self._chemin_csv_stagiaires, nouveau_chemin)
                self._chemin_csv_stagiaires = nouveau_chemin
            except Exception as e:
                tk.messagebox.showerror("Erreur", f"Impossible de renommer le fichier :\n{e}")
                log_erreur("Erreur", f"Impossible de renommer le fichier :\n{e}")
                return  # Ne pas fermer la fenêtre si erreur

            fenetre.destroy()

        def annuler():
            fenetre.destroy()
            log_erreur(f"{typeCode} non renseignée pour {chemin_a_tester} → exit()")
            exit()

        code = None

        fenetre = tk.Tk()
        fenetre.title(f"{typeCode} à renseigner manuellement")
        fenetre.resizable(False, False)
        fenetre.geometry("500x180")
        fenetre.eval('tk::PlaceWindow . center')

        # Fermer avec Échap
        fenetre.bind("<Escape>", lambda e: annuler())
        # Entrée = bouton Valider
        fenetre.bind("<Return>", lambda e: bouton_valider.invoke())

        label_warning = tk.Label(
            fenetre,
            text=f"⚠ {typeCode} non trouvé automatiquement.\nVeuillez le spécifier manuellement",
            font=("Segoe UI", 10, "bold"),
            fg="orange"
        )
        label_warning.pack(pady=(10, 5))

        label_chemin = tk.Label(
            fenetre,
            text=f"Chemin du fichier source :\n  *{chemin_a_tester}*",
            font=("Segoe UI", 9),
            justify="left",
            wraplength=480
        )
        label_chemin.pack(pady=(0, 10))

        frame_saisie = tk.Frame(fenetre)
        frame_saisie.pack()

        entry_code = tk.Entry(frame_saisie, width=10, justify="center", font=("Segoe UI", 12))
        entry_code.pack()
        entry_code.focus()

        entry_code_var = tk.StringVar()
        entry_code["textvariable"] = entry_code_var
        entry_code_var.trace_add("write", verifier_entree)

        frame_boutons = tk.Frame(fenetre)
        frame_boutons.pack(pady=10)

        bouton_valider = tk.Button(frame_boutons, text="Valider", state="disabled", width=20, command=valider)
        bouton_valider.grid(row=0, column=0, padx=5)

        if typeCode == "Code IRIS":
            bouton_renommer = tk.Button(frame_boutons, text="Valider et remplacer le nom du CSV", state="disabled", width=30, command=valider_et_renommer)
            bouton_renommer.grid(row=0, column=1, padx=5)

        bouton_annuler = tk.Button(frame_boutons, text="Annuler", width=10, command=annuler)
        bouton_annuler.grid(row=0, column=2, padx=5)

        fenetre.mainloop()

        code = int(code) if code.isdigit() else str(code)
        return code

    def recupere_trig_formation_depuis_chemin(self, chemin:str|None) -> str:
        """
        Extrait un trigramme (3 lettres/chiffres) depuis un chemin, ou le demande à l'utilisateur si introuvable.
        Gère les slashs / et \\ de manière robuste.
        """
        # On récupère le trigramme de la formation depuis le chemin du CSV ou alors on demande à l'utilisateur via tkinter
        trigramme = None

        if chemin:
            # Normalise les slashs pour s'assurer que le chemin est cohérent
            chemin_normalise = chemin.replace("\\", "/")

            # Découpe le chemin en parties
            parties = chemin_normalise.split("/")

            # Recherche un segment de 3 caractères alphanumériques
            for part in parties:
                if re.fullmatch(r"[a-zA-Z0-9]{3}", part):
                    trigramme = part
                    break

            # Si rien trouvé, on demande à l'utilisateur
            if not trigramme:
                trigramme = self.demander_code("Trigramme formation")

        else:
            trigramme = self.demander_code("Trigramme formation")

        #print(instance._trigrammeFormation)

        return trigramme

class Traiter_REE:
    def __init__(self):
        # Fichier renseigné par la ressource extérieure
        self._chemin_ficheAdministrative:str = None
        self._word_ficheAdministrative:FichierWord = None  

        # Fichier Excel à remplir pour Laetitia Da Mota (RH INSTN qui s'occupe de rentrer les REE dans IRIS)
        self._chemin_modele_excel_ficheIntervenant:str = r"\\harmonie\INSTN\UEM\_Documents_communs\Formations\Formateurs\P09-Pr01-Qualifier les ressources enseignantes\P09_Pr01_Ta.E_Grille des critères de qualification des compétences_V1.xlsx"
        self._excel_ficheIntervenant:FichierExcel = None  

        self._mail_rh:Mail = None  # Mail à envoyer à la RH INSTN qui s'occupe de rentrer les REE dans IRIS (Laetitita Da Mota)

        self._repertoire_sauvegarde_fichiersREE:str = r"\\harmonie\INSTN\UEM\_Documents_communs\Formations\Formateurs\1.Intervenants - Documents administratifs" # Lieu où sauvegarder les fichiers de l'intervenant

        # Association colonnes excel avec command control Word
        # La préparation de ce ditionnaire peut être faite avec : fe = FichierExcel.depuis_fichier(chemin_fichier=r"T:\_Documents_communs\Formations\Formateurs\P09-Pr01-Qualifier les ressources enseignantes\P09_Pr01_Ta.E_Grille des critères de qualification des compétences_V1.xlsx", charger_df=True) ; fe._tableaux["QualificationsREE"].generer_dictionnaire_depuis_excel()
        self._dict_colExcel_cc:Dict[str, str] = {
            "NOM": "Nom",
            "Pr\u00e9nom": "Prenoms",
            "Dipl\u00f4me ou formation/exp\u00e9rience professionnelle": "Diplome",
            "Dur\u00e9e exp\u00e9rience professionnelle": "DureeExperiencePro",
            "Niveau d'expertise permettant une reconnaissance": "NiveauExpertise",
            "Domaine / Sp\u00e9cialit\u00e9 \nde l'expertise": "DomaineExpertise",
            "ATTRIBUTION Niveau comp\u00e9tence": None,
            "Combien de jours anim\u00e9s, en moyenne par an": "formation_nbJoursAnimes",
            "Combien de jours de formations suivies en p\u00e9dagogie (=animation)": "formation_nbJoursFormationPedagogie",
            "Profils d'apprenants form\u00e9s": "formation_profilApprenants",
            "Taux consolid\u00e9 de la satisfaction des apprenants relativement \u00e0 l'enseignant-formateur consid\u00e9r\u00e9": None,
            "Estimation par le RP de la capacit\u00e9 de l'enseignant-formateur \u00e0 animer \n(fond de salle)": None,
            "Outils num\u00e9riques utilis\u00e9s durant les animations r\u00e9alis\u00e9es\n(serious game, blended-learning\u2026)": "formation_outilsNumeriques",
            "Combien de jours pass\u00e9s en conception de s\u00e9quence de formation, en moyenne par an": "IngPedago_nbJoursConception",
            "Combien de jours de formations suivies en ing\u00e9nierie p\u00e9dagogique (=conception de s\u00e9quences de formation)": "IngPedago_nbJoursFormationIngPedago",
            "Estimation par le RP de la conception de la s\u00e9quence en fonction des objectifs p\u00e9dagogiques fournis par le RP\n(fond de salle, analyse des supports fournis)": None,
            "Estimation par le RP de la pertinence de l'\u00e9valuation des acquis r\u00e9alis\u00e9e par l'enseignant-formateur sur sa s\u00e9quence\n(analyse de la progression des apprenants : tests avant/apr\u00e8s)": None,
            "Estimation par le RP de l'utilisation des m\u00e9thodes actives\n(\u00e9tudes de cas, r\u00e9solution de probl\u00e8mes, classes invers\u00e9es, travaux de groupes\u2026)": None,
            "Combien d'ann\u00e9es d'exp\u00e9rience en conception de dispositifs de formations\n(=cr\u00e9ation et coordination)": "IngFormation_nbJoursConception",
            "Combien de jours de formations suivies en ing\u00e9nierie de formation\n(=conception de dispositifs de formation)": "IngFormation_nbJoursFormationIngFormation",
            "Estimation par le chef de projet ou le CUE de la complexit\u00e9 des pr\u00e9c\u00e9dents dispositifs de formation con\u00e7us": None,
            "Profil des apprenants des dispositifs de formations prc\u00e9demment con\u00e7us": "IngFormation_profilApprenants",
            "Estimation par le chef de projet ou le CUE de l'\u00e9valuation des acquis r\u00e9alis\u00e9 dans le dispositif de formation\n(mesure de la progression des apprenants=estimation de la qualit\u00e9 du dispositif de formation)": None,
            "Combien d'ann\u00e9es d'exp\u00e9rience en tant que tuteur acad\u00e9mique": "IngFormation_nbAnneesTuteur",
            "Combien de r\u00e9f\u00e9rentiels d'activit\u00e9, de comp\u00e9tence et d'\u00e9valuation r\u00e9alis\u00e9s": "IngFormation_nbAnneesTuteur",
            "Combien de jours de formations suivies en ing\u00e9nierie de comp\u00e9tences": "IngCompetences_nbReferentiels",
            "Estimation par la cellule p\u00e9dagogique de DPF de la complexit\u00e9 des pr\u00e9c\u00e9dentes r\u00e9alisations de l'ing\u00e9nieur/consultant en ing\u00e9nierie de comp\u00e9tences \n(complexit\u00e9 du m\u00e9tier et de son environnement : risques, r\u00e9glementation...)": "IngCompetences_nbJoursFormationIngCompetences",
            "ATTRIBUTION Niveau comp\u00e9tence ": None,
            "Evaluation CECRL ou \u00e9quivalence TOEIC, TOEFL": "ResultatLangue2",
            "ATTRIBUTION Niveau comp\u00e9tence  ": None,
            "Curriculum vitae": None,
        }



        # Fichier Excel qui liste les AI des intervenants
        self._chemin_excel_AI:str = r"\\harmonie\INSTN\UEM\_Documents_communs\Formations\Formateurs\2.AI\Liste AI formateurs.xlsx"
        self._excel_AI:FichierExcel = None
        self._onglet_excel_AI:str = "Intervenants" #Ligne ET = 3, 1èreligne = 4
        self._tableauStructure_AI:str = "ListeIntervenants"
        nbLignes_avantET_AI:int = 2 # Si tableau structuré, normalement on n'en a pas besoin

        # Fichier Excel qui liste des intervenants qui centralise les coordonnées 
        self._chemin_excel_listeIntervenants:str = r"\\harmonie\INSTN\UEM\_Echanges\VTE\Planning UEM.xlsm"
        self._excel_listeIntervenants:FichierExcel = None
        self._onglet_excel_listeIntervenants:str = "Liste intervenants" #Ligne ET = 4, 1èreligne = 5
        self._tableauStructure_listeIntervenant:str = "ListeIntervenants"
        nbLignes_avantET_listeIntervenants:int = 3  # Si tableau structuré, normalement on n'en a pas besoin
        


        # === Début code ===

        # On ouvre le word et on charge tous les command control (filedialog depuis "Download"). On le ferme
        self._word_ficheAdministrative = FichierWord.depuisFichier()
        #print(self._word_ficheAdministrative)

        # On crée le répertoire dans le répertoire des REE s'il n'existe pas (ou assimilé) (NOM Prénom (Société - AAAA))
        self._repertoire_sauvegarde_fichiersREE += f"\\{self._word_ficheAdministrative.cc['Nom'].upper()} {self._word_ficheAdministrative.cc['Prenoms'].title()} ({self._word_ficheAdministrative.cc['RaisonSociale'] if self._word_ficheAdministrative.cc['RaisonSociale'] != 'Raison sociale employeur principal' else 'CEA'} - {datetime.now().year})"
        #print(self._repertoire_sauvegarde_fichiersREE)
        #os.makedirs(self._repertoire_sauvegarde_fichiersREE, exist_ok=True)

        # On sélectionne tous les fichiers de la REE et on les déplace dans le répertoire idoine
        #self.deplacer_fichiers(self._repertoire_sauvegarde_fichiersREE)

        # On ouvre le fichier Excel à remplir pour Laetitia Da Mota (c'est un modèle, on le collera avec le bon nom dans le répertoire idoine)
        fichier_sauvegarde_fichiersREE = os.path.join(self._repertoire_sauvegarde_fichiersREE, os.path.basename(self._chemin_modele_excel_ficheIntervenant))
        self._excel_ficheIntervenant = FichierExcel.depuis_modele(self._chemin_modele_excel_ficheIntervenant, fichier_sauvegarde_fichiersREE, charger_df=True)
        #print(self._excel_ficheIntervenant._tableaux["QualificationsREE"]._df)

        # On écrit le dataframe du tableau QualificationsREE avec les données de l'intervenant provenant du word
        df_REE = self._excel_ficheIntervenant._tableaux["QualificationsREE"]._df  # Alias
        nouvelle_ligne = {}
        for col_df, cc_key in self._dict_colExcel_cc.items():
            if cc_key is None:
                # Pas de clé correspondante => valeur vide dans la DataFrame
                nouvelle_ligne[col_df] = None
            else:
                # Récupérer la valeur dans le dictionnaire Word, ou None si la clé absente
                valeur = self._word_ficheAdministrative._cc.get(cc_key, None)
                nouvelle_ligne[col_df] = convertir_si_possible(valeur)
                #print(valeur, type(convertir_si_possible(valeur)))

        # Ajouter la nouvelle ligne au DataFrame
        # TODO : non pas sûr
        df_REE = pd.concat([df_REE, pd.DataFrame([nouvelle_ligne])], ignore_index=True)

        # On pré-rempli le fichier Excel fiche intervenant grâce aux CC et au dictionnaire
        self._excel_ficheIntervenant._tableaux["QualificationsREE"].ecrit_dataFrame_dans_tableauStructure(df_REE, remplace_df_par_nouveau=True)

        # On sauve la fiche intervenant
        #self._excel_ficheIntervenant.save()
        #self._excel_ficheIntervenant.close()

        # On ouvre l'Excel et le Word pour comparaison et adaptations manuelles

        # Dès que l'Excel est fermé, on prépare le mail pour Laetitia

        # On met à jour le fichier Excel Liste AI formateurs.xlsx : onglet intervenant, on cherche et remplace la date de validité de l'attestation employeur sinon nouvelle ligne (recopier formule + format)
        # On met à jour le fichier Excel  avec la liste des intervenants :  on cherche et remplace les données mail, tel, Ville, la date de validité de l'attestation employeur... sinon nouvelle ligne (recopier formule + format)

        
    def deplacer_fichiers(self, destination: str = None) -> None:
        """
        Ouvre un dialogue pour sélectionner des fichiers, puis les déplace vers un dossier choisi.

        Args:
            destination (str, optional): Chemin du dossier de destination.
                                        Si None, un dialogue s'ouvrira pour le choisir.
        """

        # Fenêtre Tkinter cachée
        root = tk.Tk()
        root.withdraw()

        # Sélection des fichiers à déplacer
        fichiers = filedialog.askopenfilenames(title="Sélectionner les fichiers à déplacer")
        if not fichiers:
            print("Aucun fichier sélectionné.")
            return

        # Sélection du dossier de destination
        if destination is None:
            destination = filedialog.askdirectory(title="Choisir le dossier de destination")
            if not destination:
                print("Aucun dossier de destination sélectionné.")
                return

        # Déplacement de chaque fichier
        for fichier in fichiers:
            nom_fichier = os.path.basename(fichier)
            chemin_destination = os.path.join(destination, nom_fichier)

            try:
                shutil.move(fichier, chemin_destination)
                print(f"✅ Déplacé : {nom_fichier}")
            except Exception as e:
                print(f"❌ Erreur avec {nom_fichier} : {e}")




### --------------------------------------------------------------------
#  Fonctions globales
### --------------------------------------------------------------------

def initialisationListeFichiersExportsIRIS():
    """
    Initialise les tuples des fichiers Excel (extracttions d'IRIS) à concaténer / traiter

    :Example:

    >>> initialisationListeFichiersExportsIRIS()

    .. seealso:: Rien du tout.
    .. warning:: Rien du tout.
    .. note:: Rien du tout.
    .. todo:: Rien du tout.
    """
    tSessions = (
        'R04110_Sessions-2011 à 2014 FINAL.xlsx',
        'R04110_Sessions-2015 FINAL.xlsx',
        'R04110_Sessions-2016 FINAL.xlsx',
        'R04110_Sessions-2017 FINAL.xlsx',
        'R04110_Sessions-2018 FINAL.xlsx',
        'R04110_Sessions-2019 FINAL.xlsx',
        'R04110_Sessions-2020 FINAL.xlsx',
        'R04110_Sessions-2021 FINAL.xlsx',
        'R04110_Sessions-2022 FINAL.xlsx',
        'R04110_Sessions-2023 FINAL.xlsx',
        'R04110_Sessions-2024 FINAL.xlsx',
        'R04110_Sessions-2025 au 2025.08.07.xlsx')

    tFormations = (
        "R0304_Ref_Formation-Listedesformations-2025.06.06.xlsx", )

    tVentes = (
        'R04301_Sessions-Ventes-FC2020 FINAL.xlsx',
        'R04301_Sessions-Ventes-FC2021 FINAL.xlsx',
        'R04301_Sessions-Ventes-FC2022 FINAL.xlsx',
        'R04301_Sessions-Ventes-FC2023 FINAL.xlsx',
        'R04301_Sessions-Ventes-FC2024 FINAL.xlsx',
        'R04301_Sessions-Ventes-filtre sur FC-2025 au 2025-02-05 LG.xlsx',
        'R04301_Sessions-Ventes-filtre sur FC-2025 au 2025-03-03 LG.xlsx',
        'R04301_Sessions-Ventes-filtre sur FC-2025 au 2025-04-01 LG.xlsx',
        'R04301_Sessions-Ventes-filtre sur FC-2025 au 2025-05-12 LG.xlsx',
        'R04301_Sessions-Ventes-filtre sur FC-2025 au 2025-06-02 LG.xlsx',
        'R04301_Sessions-Ventes-filtre sur FC-2025 au 2025-07-01 LG.xlsx',
        'R04301_Sessions-Ventes-filtre sur FC-2025 au 2025-08-01 LG.xlsx')
        
    tInscriptions = (
        'R04500_Sessions-Inscriptions-FC2020 FINAL.xlsx',
        'R04500_Sessions-Inscriptions-FC2021 FINAL.xlsx',
        'R04500_Sessions-Inscriptions-FC2022 FINAL.xlsx',
        'R04500_Sessions-Inscriptions-FC2023 FINAL.xlsx',
        'R04500_Sessions-Inscriptions-FC2024 FINAL.xlsx',
        'R04500_Sessions-Inscriptions-filtre sur FC-2025 au 2025-02-05.xlsx',
        'R04500_Sessions-Inscriptions-filtre sur FC-2025 au 2025-03-03.xlsx',
        'R04500_Sessions-Inscriptions-filtre sur FC-2025 au 2025-04-01.xlsx',
        'R04500_Sessions-Inscriptions-filtre sur FC-2025 au 2025-05-12.xlsx',
        'R04500_Sessions-Inscriptions-filtre sur FC-2025 au 2025-06-02.xlsx',
        'R04500_Sessions-Inscriptions-filtre sur FC-2025 au 2025-07-01.xlsx',
        'R04500_Sessions-Inscriptions-filtre sur FC-2025 au 2025-08-01.xlsx')

    return {
        "Sessions" : tSessions,
        "Formations" : tFormations,
        "Ventes" : tVentes,
        "Inscriptions" : tInscriptions
    }

def mettreAJourTousLesExportsIRIS_auto(tuple_types:Tuple(str)):
    """
    Permet de créer un seul fichier Excel à partir de plusieurs exports d'IRIS.
    Les fichiers à traités sont initialisés par la fonction initialisationListeFichiersExportsIRIS() qui permet à l'utilisateur de tout lister à la main (ça peut être plus pratique dans certains cas afin d'éviter de passer par une sélection manuelle)
    
    Les fichiers output sont des modèles avec les mêmes colonnes que les extracts d'IRIS mais avec de meilleures formes (format, couleurs...) + des colonnes adjointes à la fin pour extraire et séparer les infos du n° de session ou de la référence de la formation (ex. : trigramme formation, trigramme RP, trigramme AF...)

    Les tuples des fichiers Excel à traiter sont enregistrés dans des instances de TravauxFichiersIRIS et on emploie les méthodes de cette classe

    :param tuple_types: tuple de strings avec les noms des extracts 
    :type donnees: Tuple[str, ...]

    :Example:
    
    >>> mettreAJourTousLesExportsIRIS_auto(("Formations", ))
    >>> mettreAJourTousLesExportsIRIS_auto(("Formations", "Sessions", "Ventes", "Inscriptions"))


    .. seealso:: Rien du tout.
    .. warning:: Si une seule valeur pour tuple_types, bien mettre sous cette forme : ("Formations",) car sans la virgule Python interprête juste comme un string
    .. note:: Rien du tout.
    .. todo:: Rien du tout.
    """
    # On initialise tous les chemins des fichiers à concaténer
    dict_DE_IRIS["Fichiers"] = initialisationListeFichiersExportsIRIS()

    # On traite à la suite
    for clef in dict_DE_IRIS["CodesExports"].keys():
        if clef in tuple_types:
            print(Style.BRIGHT + Fore.YELLOW + "\nTraitement des exports " + clef)
            fichiers_formates = ["\n\t" + f for f in dict_DE_IRIS["Fichiers"][clef]]
            print("Liste des fichiers traités : " + ", ".join(fichiers_formates))
            chemins_fichiersInput = tuple(os.path.join(dict_DE_IRIS["PropExportIRIS"][clef]._input.repertoire, nom) for nom in dict_DE_IRIS["Fichiers"][clef])
            TravauxFichiersIRIS.avecEcritureOutputDefaut(dict_DE_IRIS["PropExportIRIS"][clef], chemins_fichiersInput=chemins_fichiersInput)

def mettreAJourTousLesExportsIRIS_fileDialog(tuple_types:Tuple[str, ...]):
    """
    Permet de créer un seul fichier Excel à partir de plusieurs exports d'IRIS.
    Les fichiers à traités sont sélectionnés à la suite par l'utilisateur à travers un filedialog dans l'ordre du doctionnaire d'entrée, puis tous les fichiers sont traités successivement après.
    
    Les fichiers output sont des modèles avec les mêmes colonnes que les extracts d'IRIS mais avec de meilleures formes (format, couleurs...) + des colonnes adjointes à la fin pour extraire et séparer les infos du n° de session ou de la référence de la formation (ex. : trigramme formation, trigramme RP, trigramme AF...)

    Les tuples des fichiers Excel à traiter sont enregistrés dans des instances de ExtractIRIS et on emploie les méthodes de cette classe

    :param tuple_types: tuple de strings avec les noms des extracts 
    :type donnees: Tuple[str, ...]

    :Example:
    >>> mettreAJourTousLesExportsIRIS_fileDialog(("Formations", ))
    >>> mettreAJourTousLesExportsIRIS_fileDialog(("Formations", "Sessions", "Ventes", "Inscriptions"))


    .. seealso:: Rien du tout.
    .. warning:: Si une seule valeur pour tuple_types, bien mettre sous cette forme : ("Formations",) car sans la virgule Python interprête juste comme un string
    .. note:: Rien du tout.
    .. todo:: Rien du tout.
    """
    # On fait choisir les fichiers à l'utilisateur
    dict_traitements = {}
    for clef, codeIRIS in dict_DE_IRIS["CodesExports"].items():
        if clef in tuple_types:
            traitement = TravauxFichiersIRIS(dict_DE_IRIS["PropExportIRIS"][clef])
            dict_traitements[clef] = traitement

    # On traite à la suite
    for clef in dict_traitements.keys():
        print(Style.BRIGHT + Fore.YELLOW + "\nTraitement des exports " + clef)
        fichiers_formates = ["\n\t" + f for f in dict_traitements[clef]._chemins_fichiersInput]
        print("Liste des fichiers traités : " + ", ".join(fichiers_formates))
        #print("Liste des fichiers traités : " + ", ".join(dictEI[clef].nomsFichiers_exportIRIS))
        #dict_traitements[clef].lire_extractIRIS()
        #dict_traitements[clef].ecrit_dataFrame_dans_tableauStructure(df=instance._df_chemins, supprimeDonneesEtRemplace=True)
        TravauxFichiersIRIS.avecEcritureOutputDefaut(dict_DE_IRIS["PropExportIRIS"][clef], chemins_fichiersInput=dict_traitements[clef]._chemins_fichiersInput)

def lire_fdc(chemin_fdc):

    """
    A partir d'un chemin de fichier Excel (qui se doit d'etre une fiche de couts, on retourne plusieurs dataframes.

    :param s_fdc: Chemin du fichier Excel a ouvrir (se doit d'etre une fiche de coûts)
    :type s_fdc: string
    :return: Un DataFrame de l'extract IRIS avec ajouts de colonnes (on a extrait les informations de la colonne 'N° Session' par decoupage)
    :rtype: DataFrame

    :Example:

    >>> DataFrame.df_sessionsIRIS = lire_fdc("C:\\Users\\fichier.xlsx")


    .. seealso:: Rien du tout.
    .. warning:: Rien du tout.
    .. note:: Rien du tout.
    .. todo:: Rien du tout.
    """

    nomOnglet_fdc = "Fiche de coûts"

    # Premier tableau : B5:C21 (informations génériques)
    df_fdc_infos = pd.read_excel(chemin_fdc, sheet_name=nomOnglet_fdc, usecols=[1, 2], names=["Critere", "Valeur"], header=3, nrows=17)
    #print(df_fdc_infos)
    dateCreationFormation = df_fdc_infos.iloc[4, 1] #C9
    if isinstance(dateCreationFormation, int):
        dateCreationFormation = dateCreationFormation
    elif hasattr(dateCreationFormation, 'year'):
        dateCreationFormation = dateCreationFormation.year
    else:
        print(f"Valeur inattendue pour une année : {dateCreationFormation} (type {type(dateCreationFormation)})")
        dateCreationFormation = 1900

    dureeJours_fdc = df_fdc_infos.iloc[6, 1] #C11
    osThematique = df_fdc_infos.iloc[8, 1] #C13
    nbCible_fcd = df_fdc_infos.iloc[12, 1] #C17

    # Prix de vente défini par le RP (cellule J22)
    prixVenteRetenuParParticipant = pd.read_excel(chemin_fdc, sheet_name=nomOnglet_fdc, usecols=[9], names=["Valeur"], skiprows=20, nrows=1).iloc[0,0]
    
    # Deuxième tableau : tableau des coûts (tout compris) et des prix par personne (T1, T2 et T3) : ref K35:N35
    df_fdc_couts = pd.read_excel(chemin_fdc, sheet_name=nomOnglet_fdc, usecols=[10, 11, 12, 13, 14, 15], names=["T1", "T3", "T2", "0.9xT3", "1.1xT3", "Valeur fixée"], header=33, nrows=2).transpose() #Grosse astuce : je mets 2 lignes de plus pour affecter les noms plus facilement et je les recalculerai après
    df_fdc_couts.rename(columns={0: "Couts fixes et variables nb cible"}, inplace=True) # Pour changer le nom de la colonne après transposition. Coûts prix fixes et prix
    df_fdc_couts.head() # Requis pour MàJ le nom de la colonne après transposition

    df_fdc_couts.loc["0.9xT3", "Couts fixes et variables nb cible"] = df_fdc_couts.loc["T3", "Couts fixes et variables nb cible"] * 0.9
    df_fdc_couts.loc["1.1xT3", "Couts fixes et variables nb cible"] = df_fdc_couts.loc["T3", "Couts fixes et variables nb cible"] * 1.1
    df_fdc_couts.loc["Valeur fixée", "Couts fixes et variables nb cible"] = prixVenteRetenuParParticipant / 1.1 * nbCible_fcd # Astuce : comme c'est une valeur que je n'ai pas, je fais le calcul inverse que pour avoir le montant par session avec aleas


    # On ajoute les colonnes montants cibles avec calculs
    l1 = []
    l2 = []
    l3 = []
    for index, row in df_fdc_couts.iterrows():
        l1.append(row["Couts fixes et variables nb cible"] * 1.1)
        l2.append(row["Couts fixes et variables nb cible"] * 1.1 / nbCible_fcd)
        l3.append(row["Couts fixes et variables nb cible"] * 1.1 / nbCible_fcd / dureeJours_fdc)
    df_fdc_couts["Montant cible par session avec aléas"] = l1
    df_fdc_couts["Montant cible par participant"] = l2
    df_fdc_couts["Montant cible par participant et par jour"] = l3


    # On ajoute les colonnes pour calcul nb participants
    l1 = []
    l2 = []
    l3 = []
    for index, row in df_fdc_couts.iterrows():
        if row["Montant cible par participant"] != 0 :
            l1.append(math.ceil(df_fdc_couts.loc["T1", "Montant cible par session avec aléas"] / row["Montant cible par participant"]))
            l2.append(math.ceil(df_fdc_couts.loc["T2", "Montant cible par session avec aléas"] / row["Montant cible par participant"]))
            l3.append(math.ceil(df_fdc_couts.loc["T3", "Montant cible par session avec aléas"] / row["Montant cible par participant"]))
        else :
            l1.append(0)
    df_fdc_couts["Min participants T1"] = l1
    df_fdc_couts["Min participants T2"] = l2
    df_fdc_couts["Min participants T3"] = l3 

    print(df_fdc_couts)

    return df_fdc_infos, df_fdc_couts, prixVenteRetenuParParticipant, dateCreationFormation, dureeJours_fdc, osThematique, nbCible_fcd

def optimiseCheminRepertoire(path_in):
    """
        Tout le monde n'emploie pas les noms de la GED miroir comme ils devraient.
        Ainsi pour faciliter l'utilisateur, je teste l'existance du repertoire qui devrait fonctionner.
        Si ce n'est pas le cas, je remonte d'un cran d'an l'arborescence.
        Je m'arrête si la longueur du chemin complet est nulle (chemin completement bidon). 
 
        :param path_in: Chemin du repertoire a tester
        :type path_in: string
        :return: un chemin optimal (i.e. avec la plus longue arborescence) qui est fonctionnel
        :rtype: string
 
        :Example:
 
        >>> string chemin = optimiseCheminRepertoire("C:\\Users\\fichier.xlsx")

 
        .. seealso:: Rien du tout.
        .. warning:: Rien du tout.
        .. note:: Rien du tout.
        .. todo:: Rien du tout.
    """
    

    path_out = path_in

    # On boucle de manière incrémentale vers la racine du répertoire donné en paramètre d'entrée jusqu'à ce qu'un répertoire soit ok
    while not(os.path.exists(path_out)):
        path_out = path_out[:-len(path_out.split("\\")[-1])-1]
        if len(path_out)==0 :
            path_out = "."
        #print(path_out)

    return(path_out)

def log_erreur(message: str, continuer:Boolean = False):
    # Récupérer le frame d’appel (un cran au-dessus dans la stack)
    stack = inspect.stack()
    frame = stack[1].frame

    # Nom de la fonction ou méthode appelante
    nom_fonction = stack[1].function

    # Essayer d’obtenir la classe via le premier argument (souvent self)
    cls_name = None
    if 'self' in frame.f_locals:
        cls_name = type(frame.f_locals['self']).__name__

    # Construction de l'en-tête d'erreur
    prefix = f"{Fore.YELLOW}❌ "
    location = f"{cls_name + '.' if cls_name else ''}{nom_fonction}()"
    

    # Affichage de l'erreur
    
    if continuer:
        suffix = " → On ignore la règle et on continue."
        print(f"{prefix}Erreur dans {location} : {message}{suffix}{Style.RESET_ALL}")
    else:
        print(f"{prefix}Erreur dans {location} : {message}{Style.RESET_ALL}")
        print("=== exit() ===")
        exit()

def trouve_encodage_csv(chemin_fichier: str) -> str:
    """
    Détecte l'encodage d'un fichier CSV.

    Args:
        chemin_fichier (str): Chemin vers le fichier CSV.

    Returns:
        str: L'encodage détecté (ex: 'utf-8', 'cp1252', etc.).
    """
    import chardet

    try:
        with open(chemin_fichier, 'rb') as f:
            result = chardet.detect(f.read())
        return result['encoding']
    except Exception as e:
        print(f"Erreur lors de la détection de l'encodage : {e}")
        return 'utf-8'  # Valeur par défaut en cas d'erreur

def fenetreBilan():
    def valider_champs(*args):
        trig = entry_trigramme.get().strip()
        annee = entry_annee.get().strip()
        bouton_generer.config(
            state="normal" if len(trig) == 3 and annee.isdigit() and len(annee) == 4 else "disabled"
        )

    def generer_bilan(event=None):
        trigrammeFormation = entry_trigramme.get().strip()
        anneeBilan_str = entry_annee.get().strip()
        try:
            anneeBilan = int(anneeBilan_str)
        except ValueError:
            messagebox.showerror("Erreur", "L'année doit être un entier à 4 chiffres.")
            return
        #BilanFormation(trig, annee)
        print(f"BilanFormation lancé avec : trigramme={trigrammeFormation}, année={anneeBilan}")
        bf = BilanFormation(trigrammeFormation, anneeBilan)
        bf.mergeBilan(chemin_word_bilan_input, chemin_word_bilan_output)
        fenetre.destroy()

    def annuler():
        sys.exit()

    # Création de la fenêtre
    fenetre = tk.Tk()
    fenetre.title("Création d'un bilan de formation")
    fenetre.geometry("350x180")
    fenetre.resizable(False, False)

    # Label + champ pour trigramme
    ttk.Label(fenetre, text="Trigramme formation (3 lettres) :").pack(pady=(10, 0))
    entry_trigramme = ttk.Entry(fenetre)
    entry_trigramme.pack(pady=5)

    # Label + champ pour année
    annee_defaut = str(datetime.now().year - 1)
    ttk.Label(fenetre, text="Année du bilan (4 chiffres) :").pack()
    entry_annee = ttk.Entry(fenetre)
    entry_annee.insert(0, annee_defaut)
    entry_annee.pack(pady=5)

    # Boutons
    frame_boutons = ttk.Frame(fenetre)
    frame_boutons.pack(pady=10)

    bouton_generer = ttk.Button(frame_boutons, text="Générer bilan", state="disabled", command=generer_bilan)
    bouton_generer.grid(row=0, column=0, padx=5)

    bouton_annuler = ttk.Button(frame_boutons, text="Annuler", command=annuler)
    bouton_annuler.grid(row=0, column=1, padx=5)

    # Validation en temps réel
    entry_trigramme.bind("<KeyRelease>", valider_champs)
    entry_annee.bind("<KeyRelease>", valider_champs)

    # Entrée = clic sur bouton générer
    fenetre.bind("<Return>", generer_bilan)

    # Échap = fermeture de la fenêtre
    fenetre.bind("<Escape>", lambda e: fenetre.destroy())

    # Lancer la fenêtre
    fenetre.mainloop()

def ouvrir_dossier(path) -> str:
    if platform.system() == "Windows":
        os.startfile(os.path.realpath(path))
    elif platform.system() == "Darwin":  # macOS
        subprocess.run(["open", path])
    else:  # Linux
        subprocess.run(["xdg-open", path])

class UNIVERSAL_NAME_INFO(ctypes.Structure):
    _fields_ = [("lpUniversalName", wintypes.LPWSTR)]

def chemin_vers_unc(path):
    path = os.path.normpath(path)

    if not os.path.isabs(path):
        path = os.path.abspath(path)

    if not path[1:3] == ':\\':
        return path

    buf = ctypes.create_string_buffer(1024)  # buffer brut pour la structure
    size = ctypes.c_ulong(ctypes.sizeof(buf))

    # Appel à WNetGetUniversalNameW
    result = ctypes.windll.mpr.WNetGetUniversalNameW(
        path,
        0x00000001,  # UNIVERSAL_NAME_INFO_LEVEL
        buf,
        ctypes.byref(size)
    )

    if result == 0:
        # Cast du buffer en pointeur vers UNIVERSAL_NAME_INFO
        uni_name_info = ctypes.cast(buf, ctypes.POINTER(UNIVERSAL_NAME_INFO)).contents
        return uni_name_info.lpUniversalName
    else:
        print("Marche pas, code erreur :", result)
        return path

def convertir_si_possible(valeur):
    if isinstance(valeur, str):
        valeur = valeur.strip().replace(',', '.')
        try:
            return int(valeur) if valeur.isdigit() else float(valeur)
        except ValueError:
            return valeur
    return valeur

### --------------------------------------------------------------------
#  Initialisations variables globales communes
### --------------------------------------------------------------------

# Pour couleur barres de progression
colorama.init(autoreset=True)

### --------------------------------------------------------------------
#  Code pour mises à jour des Extracts IRIS
### --------------------------------------------------------------------

# ==== Initialisation variables utilisateur ====
# Initialisation des chemins des répertoires
rep_extractIRIS_VTE = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits"
rep_extractIRIS_INSTNT = r"\\instnt\HOME\REFERENC\IRIS - rapports de synthese"
#rep_extractIRIS_INSTNT = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\GED"

test_fichierExcel_tabeauStructure = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts complets\R0304_Formations-Extract COMPLET-2025.08.19.xlsx"
test_fichierExcel_tableauNormal = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts originaux\R0304_Ref_Formation-Listedesformations-2025.06.06.xlsx"


# ==== Initialisation exports IRIS ====
dictCodesIRIS = {
        "Sessions" : "R04110",
        "Formations" : "R0304",
        "Ventes" : "R04301",
        "Inscriptions" : "R04500",
    }

sessions = PropExportIRIS(
    nom_typeExport = "Sessions",
    codeExport = "R04110",

    repertoire_input = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts originaux",
    nom_onglet_input = "Data",
    nbLignes_avantET_input = 1,

    repertoire_modele = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Modèles",
    nom_fichier_modele = "R04110_Sessions-Modèle.xlsx",

    repertoire_output = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts complets",
    nom_fichier_output = "R04110_Sessions-COMPLET.xlsx"
    )

formations = PropExportIRIS(
    nom_typeExport = "Formations",
    codeExport = "R0304",

    repertoire_input = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts originaux",
    nom_onglet_input = "Data",
    nbLignes_avantET_input = 0,

    repertoire_modele = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Modèles",
    nom_fichier_modele = "R0304_Formations-Modèle.xlsx",

    repertoire_output = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts complets",
    nom_fichier_output = "R0304_Formations-COMPLET.xlsx"
    )

ventes = PropExportIRIS(
    nom_typeExport = "Ventes",
    codeExport = "R04301",

    repertoire_input = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts originaux", #Car il y a des petits bugs sur certains CSV
    nom_onglet_input = "Data",
    nbLignes_avantET_input = 1,

    repertoire_modele = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Modèles",
    nom_fichier_modele = "R04301_Ventes-Modèle.xlsx",

    repertoire_output = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts complets",
    nom_fichier_output = "R04301_Ventes-COMPLET.xlsx"
    )

inscriptions = PropExportIRIS(
    nom_typeExport = "Inscriptions",
    codeExport = "R04500",

    repertoire_input = r"\\instnt\HOME\REFERENC\IRIS - rapports de synthese",
    nom_onglet_input = "Data",
    nbLignes_avantET_input = 1,

    repertoire_modele = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Modèles",
    nom_fichier_modele = "R04500_Inscriptions-Modèle.xlsx",

    repertoire_output = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts complets",
    nom_fichier_output = "R04500_Inscriptions-COMPLET.xlsx"
    )

dict_DE_IRIS = {}
dict_DE_IRIS["CodesExports"] = {
        "Sessions" : "R04110",
        "Formations" : "R0304",
        "Ventes" : "R04301",
        "Inscriptions" : "R04500",
    }
dict_DE_IRIS["PropExportIRIS"] = {
        "Sessions" : sessions, 
        "Formations" : formations,
        "Ventes" : ventes,
        "Inscriptions" : inscriptions}

# ==== Tests Fichier Excel ====
#fe = FichierExcel()
#fe = FichierExcel.depuis_repertoire(rep_extractIRIS_VTE)
#fe = FichierExcel.depuis_fichier(t_fichierExcel_tabeauStructure)
#fe.charger_tousTableauxStructures()

#fe = FichierExcel.depuis_fichier(r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts originaux\R04110_Sessions-2015 FINAL.xlsx")
#fe.charger_tableau("Data", nbLignes_avantET=1)
#print(fe._tableaux["Data"])
#print(fe)

# === Test ouverture fichier IRIS ===
#t_input = FichierExcel.depuis_repertoire(repertoire=r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts originaux")
#t_modele = FichierExcel.depuis_fichier(chemin_fichier=r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Modèles\R04110_Sessions-Modèle.xlsx")
#t_modele = FichierExcel.depuis_fichier(chemin_fichier=r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Modèles\R04110_Sessions-Modèle.xlsx", nom_onglet="Sessions")
#t_output = FichierExcel.depuis_fichier(chemin_fichier=r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts complets\R0304_Formations-Extract COMPLET", avec_ouverture_wb=False)

# === Test PropExportIRIS ===
#print(inscriptions)

#=== Tests TravauxFichiersIRIS ===
#env = TravauxFichiersIRIS.avecLecture(sessions, (r'C:/Users/vt238770/Documents/_CEA/_Formations/Extracts IRIS - Faits/Extracts originaux/R04110_Sessions-2021 FINAL.xlsx', r'C:/Users/vt238770/Documents/_CEA/_Formations/Extracts IRIS - Faits/Extracts originaux/R04110_Sessions-2022 FINAL.xlsx'))
#env = TravauxFichiersIRIS.avecEcritureOutputDefaut(sessions, (r'.\R04110_Sessions-2021 FINAL.xlsx', r'.\R04110_Sessions-2022 FINAL.xlsx'))
env = TravauxFichiersIRIS.avecEcritureOutputDefaut(sessions, (r'C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts originaux\R04110_Sessions-2021 FINAL.xlsx', r'C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts originaux\R04110_Sessions-2022 FINAL.xlsx'))
#env = TravauxFichiersIRIS.avecEcritureOutputDefaut(sessions, r'C:/Users/vt238770/Documents/_CEA/_Formations/Extracts IRIS - Faits/Extracts originaux/R04110_Sessions-2021 FINAL.xlsx')
#env = TravauxFichiersIRIS.avecEcritureOutputDefaut(formations, )
print(env)


# === mettreAJourTousLesExportsIRIS_auto ===
#mettreAJourTousLesExportsIRIS_auto(("Sessions", "Formations", "Ventes", "Inscriptions"))
#mettreAJourTousLesExportsIRIS_auto(("Sessions", "Formations"))
#mettreAJourTousLesExportsIRIS_fileDialog(("Formations", ))

# === Fonctions pour lancer MàJ Extracts (fonctionnelles) ===
#mettreAJourTousLesExportsIRIS_auto(("Sessions", "Formations", "Ventes", "Inscriptions"))
#mettreAJourTousLesExportsIRIS_auto(("Inscriptions",))
#mettreAJourTousLesExportsIRIS_fileDialog(("Formations", ))





### --------------------------------------------------------------------
#   Code pour création bilans pédagogiques
### --------------------------------------------------------------------

# ==== Initialisation variables utilisateur ====
# Initialisation des chemins des répertoires
rep_gedMiroir = r"\\instnt\partage\FORMATIONS_C"
rep_fdc_defaut = r"\\instnt\partage\FORMATIONS_C\XXX\P05-P06-dossier-conception-referentiel\fiche-de-cout-et-code-de-formation"
rep_specsPedagogiques_defaut = r"\\instnt\partage\FORMATIONS_C\XXX\P05-P06-dossier-conception-referentiel\specifications-pedagogiques-et-referentiel"


# Modèle du bilan à remplir
chemin_word_bilan_input = r'C:\Users\vt238770\Documents\_CEA\Prog\Modèles\Bilan formation.docx'

# Bilan en sortie après remplissage
chemin_word_bilan_output = r'C:\Users\vt238770\Documents\_CEA\Prog\Modèles\Bilan formation - output.docx'

# Fiche de coûts
#chemin_fdc = r'\\instnt\partage\FORMATIONS_C\948\P05-P06-dossier-conception-referentiel\fiche-de-cout-et-code-de-formation\Fiche de coûts - 948 - Elaboration de scénarios de DEM - 2025.01.24.xlsx'

# Specs pédagogiques
#chemin_specsPedagogiques = r'\\instnt\partage\FORMATIONS_C\948\P05-P06-dossier-conception-referentiel\specifications-pedagogiques-et-referentiel\P06-Pr01-F01_Specifications-pedagogiques - 948 - 2025.04.pdf'

# Pour chrono des fonctions
timer = Timer()


# Pour message de sortie applis externes
vlog = Vlog()


# === Lancer génération du bilan
#lire_fdc(r"\\instnt\PARTAGE\FORMATIONS_C\948\P05-P06-dossier-conception-referentiel\fiche-de-cout-et-code-de-formation\Fiche de coûts - 948 - Elaboration de scénarios de DEM - 2025.01.24.xlsx")
#fenetreBilan()

### --------------------------------------------------------------------
#   Code pour création EvalStat
### --------------------------------------------------------------------

# === Fichier CSV individuel
#es = Traiter_evalStat()
#es = Traiter_evalStat.depuis_chemin_csv_stagiaires(r"\\instnt\partage\FORMATIONS_C\22B\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\S-17606 - 22B - 06-2025\S-17606-FC25-22B-VTE-CAR-Stagiaires.csv")
#es = Traiter_evalStat.depuis_chemin_csv_stagiaires(r"P:\FORMATIONS_C\778\P07-bilan-sessions-et-bilan-formation\2024-Bilans 778\Evaluations 778(2024.11)\S-15715-FC24-778-VMO-VCA-Stagiaires.csv")

tuple_csv_stagiaires = (
    r"\\INSTNT\partage\FORMATIONS_C\948\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2021-11-S-11369 UEM\EVALSTAT\S-11369-FC21-948-VLE-SNA-Stagiaires.csv",
    r"\\INSTNT\partage\FORMATIONS_C\948\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2022-11-S-12995 UEM\S-12995-FC22-948-VTE-SNA-Stagiaires.csv",
    r"\\INSTNT\partage\FORMATIONS_C\948\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2024-01-S-15697 UEM\S-15697-FC24-948-VTE-VCA-Stagiaires.csv"
    )

#es = Traiter_evalStat.depuis_tuple_csv_stagiaires(tuple_csv_stagiaires=tuple_csv_stagiaires)



### --------------------------------------------------------------------
#   Code pour Fichier EE
### --------------------------------------------------------------------
#fw = FichierWord.depuisFichier(r"C:\Users\vt238770\Documents\_CEA\EXCEL - Tests ou Backup\Lien formulaire Word vers Excel\Fiche admin - ContentControl.docx")
#print(fw)
#mail= Mail()

#mail._creer_mail(
#    destinataires="vacataires.instn@cea.fr",
#    sujet="Documents pour mise à jour IRIS", # Pimper avec le nom de l'intervenant
#    corps_html="<p>Bonjour Laëtitia,</p><p>Je t’ai mis en PJ les documents pour intégrer/mettre à jour la fiche IRIS de ###.</p><p>Je te remercie, passe une excellente journée,</p>", # Pimper avec le nom de l'intervenant
#    pieces_jointes=r"C:\Users\vt238770\Documents\_CEA\EXCEL - Tests ou Backup\Lien formulaire Word vers Excel\Fiche admin - ContentControl.docx", # if None, sélectionner avec fileDialog
#    envoyer_directement=False  # envoie directement sans afficher
#)

#t_ree = Traiter_REE()





#Il me faudrait une classe Traiter_REE :
#   - variables de la classe : _word_ficheAdministrative:FichierWord= None, _excel_ficheIntervenant:FichierExcel = None, _mail_traitement _repertoire_sauvegarde + _mail_gestionnaire_REE +








# Todo : 
# Dans le modèle Word : gérer le lien vers la GED 
# Exploiter EvalStat
# optimiser copie format avec xlwings
# Mettre au propre
# Avoir un répertoire dédié où je vais chercher le modèle du bilan
# dans tkinter, mettre des points d'étapes dans une barre de status
# ? Exploiter export formation plutôt que export sessions pour les valeurs par défaut nmin/max...
# Faire un module / exe dédié traitements exports IRIS
# Faire un module / exe dédié traitements CSV EvalStat


# Todo Word
# Il y a des trous dans la raquette dans le word de sortie (checkboxes)
# coller des images depuis Excel
# 






# Todo : faire un truc pour la lecture sessions qui peut changer (dernière version de GLOBAL)
#def etat_boutons_traitement(state):
#    # Désactive ou active les 3 boutons
#    for btn in (btn_traiter_eval, bilan_button, iris_button):
#        btn.config(state=state)

#def chargements_initiaux():
#    fe_sessions = FichierExcel.depuis_fichier(config["Extractions d'IRIS"]["Fichier sessions R04110"])



# Au lancement, désactiver les 3 boutons de traitement
#update_status("Lecture fichier session")
#etat_boutons_traitement('disabled')

#fe_sessions = None
# Lancer le chargement en thread séparé
#threading.Thread(target=chargements_initiaux, daemon=True).start()

#print(fe_sessions)

# Quand c'est fini, on réactive les boutons, mais dans le thread Tkinter !
#etat_boutons_traitement('normal')
#update_status("Prêt.")



#fichiers = tuple(map(chemin_vers_unc, filedialog.askopenfilenames(filetypes=[("CSV files", "*.csv")])))
#print(fichiers)


#Tu peux maintenant :
#Afficher des messages dans la barre de statut avec update_status("ton message"),
#Suivre les étapes de traitement en direct pendant les clics,
#Ajouter des appels à update_status(...) dans ton futur code métier.