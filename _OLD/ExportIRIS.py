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

import warnings
from xmlrpc.client import Boolean
warnings.filterwarnings("ignore", message="Slicer List extension is not supported and will be removed")

from fileinput import filename
from typing import Tuple, List

import pandas as pd
import pandas as DataFrame

from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils.cell import range_boundaries, get_column_letter

import inspect
import os
import sys
import platform
import subprocess
import re
import time
import math
import win32com.client
import copy

import json

from datetime import date, datetime
from io import StringIO
from dataclasses import dataclass
from collections import defaultdict

from tqdm import tqdm
from colorama import Fore, Style, init

from mailmerge import MailMerge

import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, ttk, messagebox
import threading

### --------------------------------------------------------------------
#  Définitions classes utilisateur
### --------------------------------------------------------------------

@dataclass
class infosExportsIRIS:
    repertoire: str|None
    nom_fichier: str|None
    chemin_fichier: str|None
    nom_onglet: str|None
    nbLignes_avantET: int|None

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

    def ecrit_DataFrame_dans_tableauStructure(self, df:DataFrame, supprimeDonneesEtRemplace:Boolean=False, remplace_df_par_nouveau:Boolean=False):
        """
        On peuple le dataFrame dans un fichier output spécifié par un chemin, un nom de fichier, un nom d'onglet et un nom de tableau.
        On peut lui dire s'il faut remplacer les données ou écrire à la suite.

        :Example:
        >>> self.ecrit_DataFrame_dans_tableauStructure()

        .. seealso:: Rien du tout.
        .. warning:: Rien du tout.
        .. note:: Rien du tout.
        .. todo:: Rien du tout.        
        """
     
        """
        Ecrit un DataFrame dans un tableau structure d'une feuille de calcul  

        :param df: DataFrame à integrer dans le tableau structure
        :type df: DataFrame
        :param wb: classeur a lire 
        :type df: openpyxl.workbook
        :param nom_ws: nom de la feuille dans laquelle est le tableau structure
        :type nom_ws: string
        :param nom_table: nom du tableau structure. Si non renseigné, alors ce sera le même nom que l'onglet
        :type nom_table: string
        :param supprimeDonneesEtRemplace: Pour savoir si l'on ajoute les données du DataFrame à l'existant (False) ou si l'on supprime les données existantes et qu'on les remplace avec celles du DataFrame
        :type supprimeDonneesEtRemplace: Boolean
        :return: rien (on écrit/sauve un fichier excel)
        :rtype: None

        :Example:

        >>> writeDataFrameInStructuredRef(df_output, wb, nom_ws, nom_table = "Sessions", supprimeDonneesEtRemplace = False)


        .. seealso:: Rien du tout.
        .. warning:: Rien du tout.
        .. note:: Pour rajouter des lignes, on le fait à la suite de la feuille (worksheet) . Il en résulte qu'on gruge un peu : 1) on vire le tableau structuré, 2) on colle toutes les nouvelles lignes, 3) on supprime la ligne 1 du tableau, 4) on redéfinit les dimensions du tableau structuré (car l'ajout de nouvbelles lignes ne l'étend pas automatiquement)
        .. todo:: Rien du tout.
        """


        if remplace_df_par_nouveau:
            self.remplace_df(df)
        
        # On écrit toutes les autres lignes une par une (on garde les lignes initiales pour garder le format qu'on copiera)
        # Méthode 1 qui marche
        with tqdm(total=self._nb_lignes_tableau, unit=' ligne', desc=Fore.CYAN + "Écriture des lignes dans l'output" + Style.RESET_ALL) as pbar:
            for i, il in enumerate(df.itertuples(), 1):
                pbar.set_postfix(progress=f"{i}/{len(df)}")
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
        self._tableaux = {}  # Dict[str, TableauExcel]

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
        instance._chemin_fichier = chemin_fichier

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
        instance._chemin_fichier = chemin_fichier_sauv

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
        self._chemin_fichier = nouveau_chemin
        self._repertoire = os.path.dirname(nouveau_chemin)
        if os.path.isfile(nouveau_chemin):
            self._wb = load_workbook(nouveau_chemin)  # Recharge automatiquement le fichier

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
        fe_modele._tableaux[instance._nom_typeExport].ecrit_DataFrame_dans_tableauStructure(instance._df_tableau, supprimeDonneesEtRemplace=True)
        
        # On écrit les références des fichiers copiés dans le tableau structuré "Imports"
        instance._df_chemins = pd.DataFrame(instance._chemins_fichiersInput, columns=['Chemin fichier'])
        fe_modele._tableaux["Imports"].ecrit_DataFrame_dans_tableauStructure(df=instance._df_chemins, supprimeDonneesEtRemplace=True)
        
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

class TraiterEvalStat:
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
        instance = TraiterEvalStat()
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
    def depuis_tuple_csv_stagiaires(cls, tuple_csv_stagiaires:Tuple(str), chemin_excel_evaluations_defaut:str=None, chemin_modeleExcel_stagiaires:str=None, chemin_excel_sessions:str=None, ouvrirDossier:Boolean=False) -> TraiterEvalStat:
        

        
        # On crée l'instance et on complète les infos avec les valeurs facultatives
        instance = TraiterEvalStat()
        instance._chemin_excel_evaluations_defaut = chemin_excel_evaluations_defaut
        instance._chemin_modeleExcel_stagiaires = chemin_modeleExcel_stagiaires
        instance._chemin_excel_sessions = chemin_excel_sessions
        
        df_formation_csv = None
        df_formation_stagiaires = None
        #dico_sessionsDejaTraitees = None
        timer = Timer()

        # On récupère les données de Sessions (on en aura besoin plus tard)
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

                # Traitement du CSV
                if traiterCSV:  
                    timer.debut("TraiterEvalStat.depuis_chemin_csv_stagiaires")
                    traite_csv_session = TraiterEvalStat.depuis_chemin_csv_stagiaires(chemin_csv_session, fe_sessions=instance._fe_sessions, ouvrirDossier=ouvrirDossier, remplace_df=True)

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

            # On concatène, on sauve et on ferme le fe de tous les CSV de la formation
            timer.debut("Ecriture dans les tableaux Excel + ")
            fe_evaluations_formation._tableaux["CSV_stagiaires"].ecrit_DataFrame_dans_tableauStructure(df_formation_csv, supprimeDonneesEtRemplace = supprimeDonneesEtRemplace)
            fe_evaluations_formation._tableaux["Stagiaires"].ecrit_DataFrame_dans_tableauStructure(df_formation_stagiaires, supprimeDonneesEtRemplace = supprimeDonneesEtRemplace)

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
        self._fe_stagiaires._tableaux["CSV_stagiaires"].ecrit_DataFrame_dans_tableauStructure(df_csv_stagiaires, supprimeDonneesEtRemplace=True)
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
        self._fe_stagiaires._tableaux["Stagiaires"].ecrit_DataFrame_dans_tableauStructure(df_stagiaires, supprimeDonneesEtRemplace=True)
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
        #dict_traitements[clef].ecrit_DataFrame_dans_tableauStructure(df=instance._df_chemins, supprimeDonneesEtRemplace=True)
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

def lire_satisfaction_stagiaires():
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
    # dans l'onglet accueil : coller dans un tableau à part tous les commentaires en préfixant par ["Critère"] (ça alimentera le bilande foramtion à un moment) puis trier par critère
    # Actualiser les TCD
    # Sauvegarder au bon endroit

    # Possibilité d'importer plusieurs CSV pour une même année
    # Attention : impact sur nombre total de stagiaire (sommer dans sessions)
    
    # Ouvrir le dossier de sauvegarde à la fin si l'utilisateur veut voir ce que ça donne
    print()

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

def ouvrir_dossier(path):
    if platform.system() == "Windows":
        os.startfile(os.path.realpath(path))
    elif platform.system() == "Darwin":  # macOS
        subprocess.run(["open", path])
    else:  # Linux
        subprocess.run(["xdg-open", path])


### --------------------------------------------------------------------
#  Initialisations variables globales communes
### --------------------------------------------------------------------

# Pour couleur barres de progression
init(autoreset=True)

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
#env = TravauxFichiersIRIS.avecEcritureOutputDefaut(sessions, (r'C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts originaux\R04110_Sessions-2021 FINAL.xlsx', r'C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts originaux\R04110_Sessions-2022 FINAL.xlsx'))
#env = TravauxFichiersIRIS.avecEcritureOutputDefaut(sessions, r'C:/Users/vt238770/Documents/_CEA/_Formations/Extracts IRIS - Faits/Extracts originaux/R04110_Sessions-2021 FINAL.xlsx')
#env = TravauxFichiersIRIS.avecEcritureOutputDefaut(formations, )
#print(env)


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


# === Lancer génération du bilan
#lire_fdc(r"\\instnt\PARTAGE\FORMATIONS_C\948\P05-P06-dossier-conception-referentiel\fiche-de-cout-et-code-de-formation\Fiche de coûts - 948 - Elaboration de scénarios de DEM - 2025.01.24.xlsx")
#fenetreBilan()

### --------------------------------------------------------------------
#   Code pour création EvalStat
### --------------------------------------------------------------------

# === Fichier CSV individuel
#es = TraiterEvalStat()
#es = TraiterEvalStat.depuis_chemin_csv_stagiaires(r"\\instnt\partage\FORMATIONS_C\22B\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\S-17606 - 22B - 06-2025\S-17606-FC25-22B-VTE-CAR-Stagiaires.csv")
#es = TraiterEvalStat.depuis_chemin_csv_stagiaires(r"P:\FORMATIONS_C\778\P07-bilan-sessions-et-bilan-formation\2024-Bilans 778\Evaluations 778(2024.11)\S-15715-FC24-778-VMO-VCA-Stagiaires.csv")

tuple_csv_stagiaires = (
    r"\\INSTNT\partage\FORMATIONS_C\948\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2021-11-S-11369 UEM\EVALSTAT\S-11369-FC21-948-VLE-SNA-Stagiaires.csv",
    r"\\INSTNT\partage\FORMATIONS_C\948\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2022-11-S-12995 UEM\S-12995-FC22-948-VTE-SNA-Stagiaires.csv",
    r"\\INSTNT\partage\FORMATIONS_C\948\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2024-01-S-15697 UEM\S-15697-FC24-948-VTE-VCA-Stagiaires.csv"
    )

#es = TraiterEvalStat.depuis_tuple_csv_stagiaires(tuple_csv_stagiaires=tuple_csv_stagiaires)


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


