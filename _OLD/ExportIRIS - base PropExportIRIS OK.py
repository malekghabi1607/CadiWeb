import pandas as pd
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.utils.cell import range_boundaries

import os
from tqdm import tqdm
from colorama import Fore, Style, init

### --------------------------------------------------------------------
#  Définitions classes utilisateur
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
    def __init__(self, nom_onglet:str=None, nom_tableau:str=None, nbLignes_avantET:int=None):
        """
        Créer des instances de TableauExportIRIS sans charger immédiatement un fichier Excel
        """
        self._nom_onglet = nom_onglet
        self._ws = None
        
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
        # S'il n'y a pas de référence tableau, alors notre instance est un répertoire export IRIS par défaut
        sout =f"TableauExcel\n"
        f"  Onglet : {self._nom_onglet}\n"
        f"  Tableau : {self._nom_tableau or 'non structuré'}\n"

        if self._ref_tableau is None:
            sout = sout + f"  Lignes avant en-tête : {self._nbLignes_avantET}"
        else :
            sout = sout + f"  Réf : {self._ref_tableau or 'non définie'}\n" f"  Lignes avant en-tête : {self._nbLignes_avantET}\n" f"  Dimensions : lignes {self._min_row}-{self._max_row}, colonnes {self._min_col}-{self._max_col}"
        
        return sout

class TableauExportIRIS(TableauExcel):
    """
    Classe héritée de TableauExcel, ajoutant la gestion des fichiers Excel
    pour les cas spécifiques aux exports IRIS.
    """
    # === Constructeurs ===
    def __init__(self, chemin_fichier:str=None, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._chemin_fichier = chemin_fichier
        self._repertoire = os.path.dirname(chemin_fichier) if chemin_fichier else None
    
    @classmethod
    def depuis_tableau_structure(cls, chemin_fichier:str, nom_onglet:str, nom_tableau:str=None):
        """Si nom_tableau est omis, il est supposé égal à nom_onglet (convention de nommage des fichiers Excel)"""
        instance = cls(
            chemin_fichier = chemin_fichier,
            nom_onglet = nom_onglet,
            nom_tableau = nom_tableau or nom_onglet
        )
        instance.charger_workbook(load_workbook(chemin_fichier))
        return instance   

    @classmethod
    def depuis_tableau_normal(cls, chemin_fichier:str, nom_onglet:str, nbLignes_avantET:int=0):
        """Si nom_tableau est omis, il est supposé égal à nom_onglet (convention de nommage des fichiers Excel)"""
        instance = cls(
            chemin_fichier = chemin_fichier,
            nom_onglet = nom_onglet,
            nbLignes_avantET = nbLignes_avantET
        )
        instance.charger_workbook(load_workbook(chemin_fichier))
        return instance

    @classmethod
    def depuis_output(cls, repertoire:str, nom_fichier:str, nom_onglet:str, nom_tableau:str=None, nbLignes_avantET:int=None):
        """Si nom_tableau est omis, il est supposé égal à nom_onglet (convention de nommage des fichiers Excel)"""
        instance = cls(
            chemin_fichier = os.path.join(repertoire, nom_fichier),
            nom_onglet = nom_onglet
        )

        if nbLignes_avantET: # Cas tableau normal
            instance._nbLignes_avantET = nbLignes_avantET
        else: # Cas tableau structuré
            instance._nom_tableau = nom_onglet
        
        return instance

    @classmethod
    def depuis_repertoire(cls, repertoire:str, nom_onglet:str, nbLignes_avantET:int=0):
        """Si nom_tableau est omis, il est supposé égal à nom_onglet (convention de nommage des fichiers Excel)"""
        instance = cls(
            nom_onglet = nom_onglet,
            nbLignes_avantET = nbLignes_avantET
        )
        instance._repertoire = repertoire
        return instance

    # === Propriétés ===
    @property
    def chemin_fichier(self):
        return self._chemin_fichier

    @property
    def nom_fichier(self):
        return os.path.basename(self._chemin_fichier) if self._chemin_fichier else None

    @property
    def repertoire(self):
        return self._repertoire

    # === Affichage ===

    def __str__(self):
        base = super().__str__()
        return (
            f"TableauExportIRIS\n"
            f"  Nom : {self.nom_fichier or 'Aucun fichier'}\n"
            f"  Répertoire : {self._repertoire or 'Non défini'}\n"
            +base.replace('TableauExcel\n', '')
        )

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
        self.tableaux = {}  # Dict[str, TableauExcel]

    @classmethod
    def depuis_repertoire(cls, repertoire:str):
        instance = cls()
        instance._repertoire = repertoire
        return instance

    @classmethod
    def depuis_fichier(cls, chemin_fichier:str):
        instance = cls.depuis_repertoire(os.path.dirname(chemin_fichier))
        instance._chemin_fichier = chemin_fichier
        instance._wb = load_workbook(chemin_fichier)
        return instance

    # === Méthodes utilitaires ===
    def ajouter_tableau(self, tableau:TableauExcel):
        self._tableaux[tableau.nom_onglet] = tableau
    
    #def get_worksheet(self, nom_feuille):
    #    """Retourne une feuille du classeur."""
    #    return self._wb[nom_feuille]

    def close(self):
        """Ferme le fichier (facultatif avec openpyxl mais pratique)."""
        self._wb.close()

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
        return (
        f"Propriétés exports IRIS par défaut\n"
        f"  Fichier : {self.nom_fichier}\n"
        f"  Répertoire : {self._repertoire}\n"
        f"  Tableaux : {self.tableaux}\n"
        )

class PropExportIRIS:
    """C'est une fabrique. Contient toutes les propriétés des fichiers Excel issus des Exports IRIS (structure de configuration d'un export IRIS)"""
    # === Constructeurs ===
    def __init__(self, nom_typeExport:str, codeExport:str, input:TableauExportIRIS=None, modele:TableauExportIRIS=None, output:TableauExportIRIS=None):
        # Type d'export
        self._nom_typeExport = nom_typeExport
        self._codeExport = codeExport

        # Informations input données
        self._input = input

        # Informations sur le modèle Excel à employer pour remplir l'output
        self._modele = modele

        # Informations output
        self._output = output

    @classmethod
    def informationsParDefaut(cls, nom_typeExport, codeExport, repertoire_input, nom_onglet_input, nbLignes_avantET_input, repertoire_modele, nom_fichier_modele, repertoire_output, nom_fichier_output):
        instance = cls(nom_typeExport, codeExport)

        instance._input = TableauExportIRIS.depuis_repertoire(
            repertoire = repertoire_input,
            nom_onglet = nom_onglet_input,
            nbLignes_avantET = nbLignes_avantET_input)

        instance._modele = TableauExportIRIS.depuis_tableau_structure(
            chemin_fichier = os.path.join(repertoire_modele, nom_fichier_modele),
            nom_onglet = nom_typeExport)

        instance._output = TableauExportIRIS.depuis_output(
            repertoire = repertoire_output,
            nom_fichier = nom_fichier_output,
            nom_onglet = nom_typeExport)
        
        return instance

   # === Affichage ===
    def __str__(self):
        return (
                f"PropExportIRIS\n"
                f"  Type export : {self._nom_typeExport} ({self._codeExport})\n"
                f"  Input : {str(self._input).replace('  ', '    ') or 'non défini'}\n"
                f"  Modèle : {str(self._modele).replace('  ', '    ') or 'non défini'}\n"
                f"  Output : {str(self._output).replace('  ', '    ') or 'non défini'}"
            )




### --------------------------------------------------------------------
#  Code
### --------------------------------------------------------------------



### --------------------------------------------------------------------
#   Initialisation variables globales de l'utilisateur
### --------------------------------------------------------------------

# Initialisation des chemins des répertoires
rep_gedMiroir = r"\\instnt\partage\FORMATIONS_C"
rep_fdc_defaut = r"\P05-P06-dossier-conception-referentiel\fiche-de-cout-et-code-de-formation"
rep_specsPedagogiques_defaut = r"\P05-P06-dossier-conception-referentiel\specifications-pedagogiques-et-referentiel"

rep_extractIRIS_VTE = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits"
rep_extractIRIS_INSTNT = r"\\instnt\HOME\REFERENC\IRIS - rapports de synthese"
#rep_extractIRIS_INSTNT = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\GED"

# ==== Initialisation exports IRIS ====
dictCodesIRIS = {
        "Sessions" : "R04110",
        "Formations" : "R0304",
        "Ventes" : "R04301",
        "Inscriptions" : "R04500",
    }

#test = TableauExportIRIS.depuis_tableau_structure(r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\R0304_Formations-Modèle.xlsx", "Formations")
#test = TableauExportIRIS.depuis_tableau_normal(r"\\instnt\HOME\REFERENC\IRIS - rapports de synthese\R04301_Sessions-Ventes-filtre sur FC-2025 au 2025-03-03 LG.xlsx", "Data", 1)
#test = TableauExportIRIS.depuis_repertoire(r"\\instnt\HOME\REFERENC\IRIS - rapports de synthese", "Data", 1)
#test = TableauExportIRIS.depuis_repertoire_et_fichier(r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts complets", "R04500_Inscriptions-COMPLET.xlsx", nom_onglet="Inscriptions")

sessions = PropExportIRIS.informationsParDefaut(
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

print(sessions)
