
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries

import os
from tqdm import tqdm
from colorama import Fore, Style, init

### --------------------------------------------------------------------
#  Définitions classes utilisateur
### --------------------------------------------------------------------
class FichierExcel:
    """
    Classe pour manipuler un fichier Excel avec openpyxl.
    """

    # === Constructeurs ===
    def __init__(self, repertoire):
        self._repertoire = repertoire
        self._chemin_fichier = None
        self._wb = None

    @classmethod
    def depuis_fichier(cls, chemin_fichier):
        instance = cls(os.path.dirname(chemin_fichier))
        instance._chemin_fichier = chemin_fichier
        instance._wb = load_workbook(chemin_fichier)
        return instance

    # === Propriétés ===
    @property
    def chemin_fichier(self):
        return self._chemin_fichier

    @chemin_fichier.setter
    def chemin_fichier(self, nouveau_chemin):
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

    # === Méthodes utilitaires ===
    #def get_worksheet(self, nom_feuille):
    #    """Retourne une feuille du classeur."""
    #    return self._wb[nom_feuille]

    def close(self):
        """Ferme le fichier (facultatif avec openpyxl mais pratique)."""
        self._wb.close()

    # === Affichage ===
    def __str__(self):
        # S'il n'y a pas de fichier défini, alors notre instance est un répertoire export IRIS par défaut
        if self._chemin_fichier is None:
            return (
            f"Propriétés exports IRIS par défaut\n"
            f"  Répertoire : {self._repertoire}"
        )
        else :
            return (
                f"FichierExcel\n"
                f"  Nom : {self.nom_fichier or 'Pas de fichier défini (répertoire export IRIS par défaut)'}\n"
                f"  Répertoire : {self._repertoire}"
            )

class TableauExcel(FichierExcel):
    """
    Représente un tableau contenu dans un fichier Excel, pouvant être un tableau structuré (Excel Table)
    ou un tableau "normal" (données tabulaires sans table Excel explicite).

    Cette classe hérite de `FichierExcel` et ajoute des fonctionnalités spécifiques à la manipulation
    de tableaux présents dans un onglet d’un fichier Excel.

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
    def __init__(self, repertoire):
        super().__init__(repertoire)
        self._nom_onglet = None
        self._ws = None
        
        # Informations si tableau structuré (peuvent être être None si on a un tableau normal)
        self._nom_tableau = None
        self._table = None  # Ne contient que les méta-données du tableau, pas les données elles-mêmes

        # Informations si tableau normal (peuvent être être None si on a un tableau structuré)
        self._nbLignes_avantET = None

        # Dimensions et références tableau
        self._ref_tableau = None  # Pour mémoriser la référence (ex: A1:D12)
        self._min_row = None
        self._max_row = None
        self._min_col = None
        self._max_col = None

    @classmethod
    def depuis_tableau_structure(cls, chemin_fichier, nom_onglet, nom_tableau = ""):
        """Si nom_tableau n'est pas donné, alors nom_tableau = nom_onglet (convention de nommage des fichiers Excel)"""
        instance = cls.depuis_fichier(chemin_fichier)
        instance._nom_onglet = nom_onglet
        
        # Normalement : nom du Tableau = nom de l'onglet (Comportement de règle fichiers Excel) ; mais au cas où on peut spécifier manuellement
        if nom_tableau == "":
            instance._nom_tableau = nom_onglet
        else:
            instance._nom_tableau = nom_tableau 

        instance._ws = instance._wb[nom_onglet]
        instance._table = instance._ws.tables[instance._nom_tableau]

        instance._initialiser_dimensions_structured_table()
        return instance

    @classmethod
    def depuis_tableau_normal(cls, chemin_fichier, nom_onglet, nbLignes_avantET):
        instance = cls.depuis_fichier(chemin_fichier)
        instance._nom_onglet = nom_onglet
        instance._nbLignes_avantET = nbLignes_avantET

        instance._ws = instance._wb[nom_onglet]
        
        instance._initialiser_dimensions_tableau_normal()
        return instance

    @classmethod
    def depuis_repertoire(cls, repertoire, nom_onglet, nbLignes_avantET):
        """
        Cas particulier : c'est pour renseigner les valeurs par défaut des exports IRIS.
        Il n'y a pas de tableau à proprement parler.
        On rentre :
            un répertoire (répertoire par défaut où sont stockés ces exports IRIS)
            un nom d'onglet (où est le tableau utile pour nous)
            un nombre de lignes avant l'en-tête du tableau
        """
        instance = cls(repertoire)
        instance._nom_onglet = nom_onglet
        instance._nbLignes_avantET = nbLignes_avantET

        return instance

    @classmethod
    def depuis_output(cls, repertoire, nom_fichier, nom_onglet, nom_tableau = "", nbLignes_avantET = None):
        """Si nom_tableau n'est pas donné, alors nom_tableau = nom_onglet (convention de nommage des fichiers Excel)"""
        instance = cls(repertoire)
        instance._chemin_fichier = os.path.join(repertoire, nom_fichier)
        instance._nom_onglet = nom_onglet
        
        # Normalement : nom du Tableau = nom de l'onglet (Comportement de règle fichiers Excel) ; mais au cas où on peut spécifier manuellement
        if nom_tableau == "":
            instance._nom_tableau = nom_onglet
        else:
            instance._nom_tableau = nom_tableau 

        """
        # a initialiser plus tard
        instance._ws = instance._wb[nom_onglet]
        instance._table = instance._ws.tables[instance._nom_tableau]

        instance._initialiser_dimensions_structured_table()
        """
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
        base = super().__str__()
        # S'il n'y a pas de fichier défini, alors notre instance est un répertoire export IRIS par défaut
        if self._chemin_fichier is None:
            return (
                f"{base}\n"
                f"  Onglet : {self._nom_onglet}\n"
                f"  Tableau : {self._nom_tableau or 'non structuré'}\n"
                f"  Lignes avant en-tête : {self._nbLignes_avantET}"
            )
        else :
            return (
                f"{base}\n"
                f"  Onglet : {self._nom_onglet}\n"
                f"  Tableau : {self._nom_tableau or 'non structuré'}\n"
                f"  Réf : {self._ref_tableau or 'non définie'}\n"
                f"  Lignes avant en-tête : {self._nbLignes_avantET}\n"
                f"  Dimensions : lignes {self._min_row}-{self._max_row}, colonnes {self._min_col}-{self._max_col}"
            )

class PropExportIRIS:
    """Contient toutes les propriétés des fichiers Excel issus des Exports IRIS """
    # === Constructeurs ===
    def __init__(self, nom_typeExport, codeExport, input=None, modele=None, output=None):
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

        instance._input = TableauExcel.depuis_repertoire(
            repertoire = repertoire_input,
            nom_onglet = nom_onglet_input,
            nbLignes_avantET = nbLignes_avantET_input)

        instance._modele = TableauExcel.depuis_tableau_structure(
            chemin_fichier = os.path.join(repertoire_modele, nom_fichier_modele),
            nom_onglet = nom_typeExport)

        instance._output = TableauExcel.depuis_output(
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

class ExportIRIS(PropExportIRIS):
    "C'est la classe qui contient tous les éléments d'un type d'export IRIS"   
    # === Constructeurs ===
    def __init__(self, nom_typeExport, codeExport):
        super().__init__(nom_typeExport, codeExport)
        self._inputDonnees_chemins = None
        self._df_tableau = None
        
    
    @classmethod
    def depuisChemins(cls, propExportIRIS, chemins_fichiersInput):
        instance = cls(chemin_fichier)
        instance._nom_onglet = nom_onglet
        # _inputDonnees_chemins doit être un tuple de strings. Si c'est un string c'est qu'un seul fichier a été donné. Alors on convertit en tuple
        if isinstance(chemins_fichiersInput, str):
            instance._chemins_fichiersInput = (chemins_fichiersInput,)
        else:
            instance._chemins_fichiersInput = chemins_fichiersInput
        
        return instance
    
    def peuple_df(self):
        """
        Peuple le DataFrame self.__df_tableau du ou des tableaux issus de l'export IRIS.
        On selectionne la bonne methode en fonction du type d'export

        :Example:
        >>> self.peuple_df()


        .. seealso:: Rien du tout.
        .. warning:: Rien du tout.
        .. note:: Rien du tout.
        .. todo:: Rien du tout.
        """

        # Initialisation : on crée un DataFrame vide pour recevoir (peut-être) des infos que l'on traitera et qui nécessitera d'adjoindre des colonnes à self.__df_tableau
        df_colonnes_sup = None
        df_list = [] # Liste des DataFrame qui contiendra chaque fichier Excel séparément

        # On parcourt le tuple des fichiers à lire
        taille_totale = sum(os.path.getsize(cheminFichier) for cheminFichier in self.__inputDonnees_chemins) # Calcul taille totale pour barre de progression

        with tqdm(total=taille_totale, unit='o', unit_scale=True, desc=Fore.CYAN+"Lecture des fichiers Excel" + Style.RESET_ALL) as pbar:
            for i, cheminFichier in enumerate(self.__inputDonnees_chemins, 1):
                #On crée les indicateurs pour tqdm
                taille = os.path.getsize(cheminFichier)

                # Affichage dynamique dans la barre
                pbar.set_postfix(file=os.path.dirname(cheminFichier), progress=f"{i}/{len(self.__inputDonnees_chemins)}")
                
                # On lit l'excel ; skiprows = 1 car l'extract IRIS a une première ligne que l'on doit sauter
                df = pd.read_excel(cheminFichier, skiprows=self.__nbLignes_avantET_exportIRIS)
                #print("\nTraitement "+ifichier)
                #print("\tdf")
                #print(df)
                
                # On ajoute le DataFrame à notre liste de DataFrame
                df_list.append(df)
                #print("\tdf_list")
                #print(df)

                # Mise à jour de la barre avec la taille du fichier
                pbar.update(taille)

        # Concaténation finale (note : toute la fin de la méthode se fait quasi-instantanément)
        self.__df_tableau = pd.concat(df_list, ignore_index=True) 

        # Selon le type d'export à traiter, on va faire des traitements spécifiques (extraction d'info des colonnes référence formation ou n° Iris)
        match self.__codeExport:
            # Cas Sessions ou Inscriptions ou Ventes (sensiblement comme 'Inscription R04500' mais groupé par Client (pas de détail de chaque stagiaire))
            case "R04110" | "R04500" | "R04301":
                # On extrait / retravaille les informations de la colonne 'N° Session'
                df_colonnes_sup = self.__df_tableau['N° Session'].apply(self.extraire_infos_numSessionIRIS)
                # A cause de certains éléments None ou NaN, Pandas type la colonne en float. Je la retype en Int64 qui permet de stocker des NaN avec des entiers, contrairement au type int standard.
                df_colonnes_sup['Année'] = df_colonnes_sup['Année'].astype('Int64')
                #print(df_colonnes_sup)

            # Cas Formations
            case "R0304":
                # On extrait / retravaille les informations de la colonne 'Référence'
                df_colonnes_sup = self.__df_tableau['Référence'].apply(self.extraire_infos_referenceFormationIRIS)
                # A cause de certains éléments None ou NaN, Pandas type la colonne en float. Je la retype en Int64 qui permet de stocker des NaN avec des entiers, contrairement au type int standard.
                df_colonnes_sup['Année'] = df_colonnes_sup['Année'].astype('Int64')
                #print(df_colonnes_sup)

        # Ajouter les colonnes supplémentaires au DataFrame principal ssi le DataFrame df_colonnes_sup exite
        if df_colonnes_sup is not None:
            self.__df_tableau = pd.concat([self.__df_tableau, df_colonnes_sup], axis=1)


### --------------------------------------------------------------------
#  Code
### --------------------------------------------------------------------

#fichier1 = FichierExcel(r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\R04110_Sessions-UEM-2025.01.28.xlsx")
#print(fichier1.nomFichier)

#test = TableauExcel.depuis_tableau_structure(r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\R0304_Formations-Modèle.xlsx", "Formations")
#test = TableauExcel.depuis_tableau_normal(r"\\instnt\HOME\REFERENC\IRIS - rapports de synthese\R04301_Sessions-Ventes-filtre sur FC-2025 au 2025-03-03 LG.xlsx", "Data", 1)
#test = TableauExcel.depuis_repertoire(r"\\instnt\HOME\REFERENC\IRIS - rapports de synthese", "Data", 1)
#print(test)


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

formations = PropExportIRIS.informationsParDefaut(
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

ventes = PropExportIRIS.informationsParDefaut(
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

inscriptions = PropExportIRIS.informationsParDefaut(
    nom_typeExport = "Inscriptions",
    codeExport = "R04500",

    repertoire_input = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts originaux",
    nom_onglet_input = "Data",
    nbLignes_avantET_input = 0,

    repertoire_modele = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Modèles",
    nom_fichier_modele = "R04500_Inscriptions-Modèle.xlsx",

    repertoire_output = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts complets",
    nom_fichier_output = "R04500_Inscriptions-COMPLET.xlsx"
    )

print(inscriptions)