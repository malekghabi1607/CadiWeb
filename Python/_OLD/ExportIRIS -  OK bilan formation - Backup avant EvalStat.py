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
from fileinput import filename
from typing import Tuple

import pandas as pd
import pandas as DataFrame

from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils.cell import range_boundaries
from copy import copy

import inspect
import os
import sys
import re
import time
import math

from datetime import date
from datetime import datetime
from io import StringIO
from dataclasses import dataclass

from tqdm import tqdm
from colorama import Fore, Style, init

from mailmerge import MailMerge

import tkinter as tk
from tkinter import ttk, messagebox
from tkinter import filedialog

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

    def ecrit_DataFrame_dans_tableauStructure(self, df:DataFrame, supprimeDonneesEtRemplace=False):
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
        # Si nom_Table n'est pas défini en argument, c'est que par défaut c'est le même que nom_ws 
        #if nom_table == "": nom_table = nom_ws

        # On ouvre la feuille et le tableau structuré
        #ws = wb[nom_ws]
        #print(self._ws)
        #table = ws._tables[nom_table] #Tableau structuré nommé

        
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
                    #target.font = copy(source.font)
                    #target.border = copy(source.border)
                    target.fill = copy(source.fill)
                    #target.alignment = copy(source.alignment)
                    #target.protection = copy(source.protection)
                    target.number_format = copy(source.number_format)

                pbar.update(1)

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
    def depuis_fichier(cls, chemin_fichier:str=None, avec_ouverture:bool=True, charger_tableau:bool=True, nom_onglet:str=None, nom_tableau:str=None, nbLignes_avantET:int=0, charger_df:bool=True):
        """
        Par défaut ou ouvre le workbook, mais on peut sépcifier que non avec avec_ouverture=False
        Si charger_tableau = True
            Si pas de nom_onglet : on charge tous les tableaux structurés
            Si nom_onglet : on charge un seul tableau (structuré ou non) qui est dans cet onglet (il faut alors un nom_onglet)
        """
        #Reprise de comportement (si on ne demande pas l'ouverture (chargement wb), alors pas de raison de charger le tableau. Idemn avec charger_tableau et charger_df)
        if not avec_ouverture: charger_tableau=False
        if not charger_tableau: charger_df=False

        if not chemin_fichier:
            chemin_fichier = cls.choisirFichiers_filedialog()
        instance = cls.depuis_repertoire(os.path.dirname(chemin_fichier))
        instance._chemin_fichier = chemin_fichier
        
            

        if avec_ouverture:
            instance.charger_wb()

        if charger_tableau:
            # Si nom_onglet : on charge un tableau (structuré ou non) dans nom_onglet ; sinon on charge tous les tableaux structurés
            if nom_onglet :
                instance.charger_tableau(nom_onglet, nom_tableau, nbLignes_avantET, charger_df)
            else:
                # TODO Besoin de close _wb ?
                instance.charger_TousTableauxStructures(charger_df=charger_df)

        return instance


    # === Méthodes utilitaires ===
    def charger_wb(self):
        self._wb = load_workbook(self._chemin_fichier)
    
    def charger_tableau(self, nom_onglet:str, nom_tableau:str=None, nbLignes_avantET:int=0, charger_df:bool=True):
        self.ajouter_tableau(TableauExcel.depuis_FichierExcel_et_nomOnglet(self, nom_onglet=nom_onglet, nom_tableau=nom_tableau, nbLignes_avantET=nbLignes_avantET, charger_df=charger_df))
    
    def charger_TousTableauxStructures(self, charger_df:bool=True):
        for nom_onglet in self._wb.sheetnames:
            for nom_tableau in self._wb[nom_onglet].tables:
                self.ajouter_tableau(TableauExcel.depuis_FichierExcel_et_nomOnglet(self, nom_onglet=nom_onglet, nom_tableau=nom_tableau, charger_df=charger_df))

    def ajouter_tableau(self, tableau:TableauExcel):
        self._tableaux[tableau._nom_tableau] = tableau
    
    #def get_worksheet(self, nom_feuille):
    #    """Retourne une feuille du classeur."""
    #    return self._wb[nom_feuille]

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
        nom_onglet = instance._output.nom_onglet

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
            trigramme_RP = blocs[-1].rstrip('_')
            trigramme_AF = blocs[-2].rstrip('_')
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

def log_erreur(message: str):
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
    prefix = f"{Fore.RED}❌ "
    location = f"{cls_name + '.' if cls_name else ''}{nom_fonction}()"

    # Affichage de l'erreur
    print(f"{prefix}Erreur dans {location} : {message}{Style.RESET_ALL}\n=== exit() ===")
    exit()


###

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

    repertoire_input = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts originaux",
    nom_onglet_input = "Data",
    nbLignes_avantET_input = 0,

    repertoire_modele = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Modèles",
    nom_fichier_modele = "R04301_Ventes-Modèle.xlsx",

    repertoire_output = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts complets",
    nom_fichier_output = "R04301_Ventes-COMPLET.xlsx"
    )

inscriptions = PropExportIRIS(
    nom_typeExport = "Inscriptions",
    codeExport = "R04500",

    repertoire_input = r"\\instnt\partage\FORMATIONS_C\Extracts originaux",
    nom_onglet_input = "Data",
    nbLignes_avantET_input = 0,

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
#fe.charger_TousTableauxStructures()

#fe = FichierExcel.depuis_fichier(r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts originaux\R04110_Sessions-2015 FINAL.xlsx")
#fe.charger_tableau("Data", nbLignes_avantET=1)
#print(fe._tableaux["Data"])
#print(fe)

# === Test ouverture fichier IRIS ===
#t_input = FichierExcel.depuis_repertoire(repertoire=r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts originaux")
#t_modele = FichierExcel.depuis_fichier(chemin_fichier=r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Modèles\R04110_Sessions-Modèle.xlsx")
#t_modele = FichierExcel.depuis_fichier(chemin_fichier=r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Modèles\R04110_Sessions-Modèle.xlsx", nom_onglet="Sessions")
#t_output = FichierExcel.depuis_fichier(chemin_fichier=r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts complets\R0304_Formations-Extract COMPLET", avec_ouverture=False)

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
#mettreAJourTousLesExportsIRIS_fileDialog(("Formations", ))





### --------------------------------------------------------------------
#   Code pour création bilans pédagogiques
### --------------------------------------------------------------------

# ==== Initialisation variables utilisateur ====
# Initialisation des chemins des répertoires
rep_gedMiroir = r"\\instnt\partage\FORMATIONS_C"
rep_fdc_defaut = r"\\instnt\partage\FORMATIONS_C\XXX\P05-P06-dossier-conception-referentiel\fiche-de-cout-et-code-de-formation"
rep_specsPedagogiques_defaut = r"\\instnt\partage\FORMATIONS_C\948\P05-P06-dossier-conception-referentiel\specifications-pedagogiques-et-referentiel"


# Modèle du bilan à remplir
chemin_word_bilan_input = r'C:\Users\vt238770\Documents\_CEA\Prog\Modèles\Bilan formation.docx'

# Bilan en sortie après remplissage
chemin_word_bilan_output = r'C:\Users\vt238770\Documents\_CEA\Prog\Modèles\Bilan formation - output.docx'

# Fiche de coûts
#chemin_fdc = r'\\instnt\partage\FORMATIONS_C\948\P05-P06-dossier-conception-referentiel\fiche-de-cout-et-code-de-formation\Fiche de coûts - 948 - Elaboration de scénarios de DEM - 2025.01.24.xlsx'

# Specs pédagogiques
#chemin_specsPedagogiques = r'\\instnt\partage\FORMATIONS_C\948\P05-P06-dossier-conception-referentiel\specifications-pedagogiques-et-referentiel\P06-Pr01-F01_Specifications-pedagogiques - 948 - 2025.04.pdf'




#trigrammeFormation = "778"
#anneeBilan = 2023

#bf = BilanFormation(trigrammeFormation, anneeBilan, chemin_specsPedagogiques, chemin_fdc)
#bf.mergeBilan(chemin_word_bilan_input, chemin_word_bilan_output)

#fenetreBilan()
#lire_fdc(r"\\instnt\PARTAGE\FORMATIONS_C\948\P05-P06-dossier-conception-referentiel\fiche-de-cout-et-code-de-formation\Fiche de coûts - 948 - Elaboration de scénarios de DEM - 2025.01.24.xlsx")

# Todo : 
# Dans le modèle Word : gérer le lien vers la GED 
# Faire backup de cet état
# Exploiter EvalStat
# optimiser copie format avec xlwings
# Mettre au propre
# Avoir un répertoire dédié où je vais chercher le modèle du bilan
# A la fin de la réalisation, ouvrir le bon répertoire output
# dans tkinter, mettre des points d'étapes dans une barre de status
# ? Exploiter export formation plutôt que export sessions pour les valeurs par défaut nmin/max...
# Faire un module / exe dédié traitements exports IRIS
# Faire un module / exe dédié traitements CSV EvalStat


# Todo Word
# Il y a des trous dans la raquette dans le word de sortie (checkboxes)
# coller des images depuis Excel
# 