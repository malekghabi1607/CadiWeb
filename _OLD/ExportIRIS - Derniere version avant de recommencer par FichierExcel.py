import pandas as pd
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.utils.cell import range_boundaries

import os
import re
from tqdm import tqdm
from colorama import Fore, Style, init

from tkinter import filedialog

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

    def importe_df(self, df, supprimeDonneesEtRemplace):
        print()

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
        if self._ref_tableau is None:
            return(
                f"TableauExcel\n"
                f"  Onglet : {self._nom_onglet}\n"
                f"  Tableau : {self._nom_tableau or 'non structuré'}\n"
                f"  Lignes avant en-tête : {self._nbLignes_avantET}"
            )
        else:
            return(
                f"TableauExcel\n"
                f"  Onglet : {self._nom_onglet}\n"
                f"  Tableau : {self._nom_tableau or 'non structuré'}\n"
                f"  Réf : {self._ref_tableau or 'non définie'}\n"
                f"  Lignes avant en-tête : {self._nbLignes_avantET}\n"
                f"  Dimensions : lignes {self._min_row}-{self._max_row}, colonnes {self._min_col}-{self._max_col}"
            )

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
        self._tableaux = {}  # Dict[str, TableauExcel]

    @classmethod
    def depuis_repertoire(cls, repertoire:str):
        instance = cls()
        instance._repertoire = repertoire
        return instance

    @classmethod
    def depuis_fichier(cls, chemin_fichier:str):
        instance = cls.depuis_repertoire(os.path.dirname(chemin_fichier))
        instance._chemin_fichier = chemin_fichier
        return instance

    @classmethod
    def depuis_fichier_avecOuvertureWB(cls, chemin_fichier:str):
        instance = cls.depuis_fichier(chemin_fichier)
        instance._wb = load_workbook(chemin_fichier)
        return instance

    # === Méthodes utilitaires ===
    def ajouter_tableau(self, tableau:TableauExcel):
        self._tableaux[tableau.nom_tableau] = tableau
    
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
        # Pour aff_df
        aff_tableau = ""
        if self._tableaux is not None:
            for clef, tableau in self._tableaux.items():
                aff_tableau = aff_tableau + "\n    " + tableau.nom_tableau
        else:
            aff_tableau = "Non défini"

        return (
        f"Propriétés fichier Excel\n"
        f"  Fichier : {self.nom_fichier}\n"
        f"  Répertoire : {self._repertoire}\n"
        f"  Liste des tableaux du fichier : {aff_tableau}\n"
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

    # === Accesseurs ===


    # === Affichage ===
    def __str__(self):
        return (
            f"PropExportIRIS\n"
            f"  Type export : {self._nom_typeExport} ({self._codeExport})\n"
            f"  Input : {str(self._input).replace('  ', '    ') or 'non défini'}\n"
            f"  Modèle : {str(self._modele).replace('  ', '    ') or 'non défini'}\n"
            f"  Output : {str(self._output).replace('  ', '    ') or 'non défini'}"
            )

class TravauxFichiersIRIS:
    "C'est la classe qui contient l'environnement pour bosser sur des fichiers Exports IRIS"
    def __init__(self, prop:PropExportIRIS, chemins_fichiersInput:str|tuple[str, ...]=None):
        self._prop = prop
        self._df_tableau = None

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
    #def avecEcriture(cls, propExportIRIS:PropExportIRIS, chemins_fichiersInput:str|tuple[str, ...], fichier_output = "", repertoire_output = "", nom_ws_output = "", nom_tableau_output = "", supprimeDonneesEtRemplace = False, nom_ws_imports = ""):
    def avecEcritureOutputDefaut(cls, propExportIRIS:PropExportIRIS, chemins_fichiersInput:str|tuple[str, ...]=None):
        
        instance = cls.avecLecture(propExportIRIS, chemins_fichiersInput)
        fe_output = FichierExcel.depuis_fichier(instance._prop._output._chemin_fichier)
        fe_output.ajouter_tableau(instance._prop._output)
        print(fe_output)
        
        # On lit le/les extract IRIS et on stocke dans self.__df_tableau
        #instance.ecrit_DataFrame_dans_tableauStructure(fichier_output = fichier_output, supprimeDonneesEtRemplace = supprimeDonneesEtRemplace, repertoire_output = repertoire_output, nom_ws_output = nom_ws_output, nom_tableau_output = nom_tableau_output, nom_ws_imports = nom_ws_imports)
        #instance.ecrit_DataFrame_dans_tableauStructure(fichier_output = fichier_output, supprimeDonneesEtRemplace = supprimeDonneesEtRemplace, repertoire_output = repertoire_output, nom_ws_output = nom_ws_output, nom_tableau_output = nom_tableau_output, nom_ws_imports = nom_ws_imports)

        return instance
 

    # === Méthodes ===
    def choisirFichiers_filedialog(self):
        # Lister/sélectionner les documents à concaténer
        cheminsExcel = filedialog.askopenfilenames(title="Sélectionner les fichiers " + self._prop._nom_typeExport + " (" + self._prop._codeExport + ") Excel à concaténer", filetype=[("Fichiers Excel", "*.xlsx")], initialdir=self._prop._input._repertoire)
        
        # Gestion du cas où il y a non-sélection de fichiers
        if not cheminsExcel:
            exit()
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
                
                df = pd.read_excel(chemin, skiprows=self._prop._input._nbLignes_avantET)
                df_list.append(df)  # On ajoute le DataFrame à notre liste de DataFrame
                
                # Mise à jour de la barre avec la taille du fichier
                pbar.update(taille)

        # Concaténation finale (note : toute la fin de la méthode se fait quasi-instantanément)
        self._df_tableau = pd.concat(df_list, ignore_index=True) 

        # Selon le type d'export à traiter, on va faire des traitements spécifiques (extraction d'info des colonnes référence formation ou n° Iris)
        match self._prop._codeExport:
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
    
    def extraire_infos_numSessionIRIS(self, reference):
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
        'R04110_Sessions-Extract année 2011 2012 2013 2014 complète-Fait le 2023.05.04.xlsx',
        'R04110_Sessions-Extract année 2015 complète-Fait le 2023.05.04.xlsx',
        'R04110_Sessions-Extract année 2016 complète-Fait le 2023.05.04.xlsx',
        'R04110_Sessions-Extract année 2017 complète-Fait le 2023.05.04.xlsx',
        'R04110_Sessions-Extract année 2018 complète-Fait le 2023.05.04.xlsx',
        'R04110_Sessions-Extract année 2019 complète-Fait le 2023.05.04.xlsx',
        'R04110_Sessions-Extract année 2020 complète-Fait le 2023.05.04.xlsx',
        'R04110_Sessions-Extract année 2021 complète-Fait le 2023.05.04.xlsx',
        'R04110_Sessions-Extract année 2022 complète-Fait le 2023.05.02.xlsx',
        'R04110_Sessions-Extract année 2023 complète-Fait le 2025.08.07.xlsx',
        'R04110_Sessions-Extract année 2024 complète-Fait le 2025.08.07.xlsx',
        'R04110_Sessions-Extract année 2025 en cours-Fait le 2025.08.07.xlsx')

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

def mettreAJourTousLesExportsIRIS_auto(dictCodesIRIS):
    """
    Permet de créer un seul fichier Excel à partir de plusieurs exports d'IRIS.
    Les fichiers à traités sont initialisés par la fonction initialisationListeFichiersExportsIRIS() qui permet à l'utilisateur de tout lister à la main (ça peut être plus pratique dans certains cas afin d'éviter de passer par une sélection manuelle)
    
    Les fichiers output sont des modèles avec les mêmes colonnes que les extracts d'IRIS mais avec de meilleures formes (format, couleurs...) + des colonnes adjointes à la fin pour extraire et séparer les infos du n° de session ou de la référence de la formation (ex. : trigramme formation, trigramme RP, trigramme AF...)

    Les tuples des fichiers Excel à traiter sont enregistrés dans des instances de ExtractIRIS et on emploie les méthodes de cette classe

    :param dictCodesIRIS: dictionnaire de la forme {"Sessions" : "R04110", ...}
    :type donnees: Dictionnaire

    :Example:
    >>> mettreAJourTousLesExportsIRIS_auto({"Sessions" : "R04110"})


    .. seealso:: Rien du tout.
    .. warning:: Rien du tout.
    .. note:: Rien du tout.
    .. todo:: Rien du tout.
    """
    # On initialise tous les chemins des fichiers à concaténer
    dictChemins = initialisationListeFichiersExportsIRIS()

    # On traite à la suite
    for clef, codeIRIS in dictCodesIRIS.items():
        print(Style.BRIGHT + Fore.YELLOW + "\nTraitement des exports " + clef)
        fichiers_formates = ["\n\t" + f for f in dictChemins[clef]]
        print("Liste des fichiers traités : " + ", ".join(fichiers_formates))
        EI = TravauxFichiersIRIS.avecEcriture(codeExport=codeIRIS, nomsFichiers_exportIRIS=dictChemins[clef], supprimeDonneesEtRemplace=True, nom_ws_imports = "Imports")





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




# Pour couleur barres de progression
init(autoreset=True)


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

    repertoire_input = r"\\instnt\partage\FORMATIONS_C\Extracts originaux",
    nom_onglet_input = "Data",
    nbLignes_avantET_input = 0,

    repertoire_modele = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Modèles",
    nom_fichier_modele = "R04500_Inscriptions-Modèle.xlsx",

    repertoire_output = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts complets",
    nom_fichier_output = "R04500_Inscriptions-COMPLET.xlsx"
    )

dictCodesIRIS = {
        "Sessions" : "R04110",
        "Formations" : "R0304",
        "Ventes" : "R04301",
        "Inscriptions" : "R04500",
    }
#mettreAJourTousLesExportsIRIS_auto(dictCodesIRIS)


#env = TravauxFichiersIRIS.avecLecture(sessions, r'C:/Users/vt238770/Documents/_CEA/_Formations/Extracts IRIS - Faits/Extracts originaux/R04110_Sessions-2021 FINAL.xlsx')

#env = TravauxFichiersIRIS.avecLecture(sessions, (r'C:/Users/vt238770/Documents/_CEA/_Formations/Extracts IRIS - Faits/Extracts originaux/R04110_Sessions-2021 FINAL.xlsx', r'C:/Users/vt238770/Documents/_CEA/_Formations/Extracts IRIS - Faits/Extracts originaux/R04110_Sessions-2022 FINAL.xlsx', r'C:/Users/vt238770/Documents/_CEA/_Formations/Extracts IRIS - Faits/Extracts originaux/R04110_Sessions-2023 FINAL.xlsx'))
#env = TravauxFichiersIRIS.avecEcritureOutputDefaut(sessions, (r'C:/Users/vt238770/Documents/_CEA/_Formations/Extracts IRIS - Faits/Extracts originaux/R04110_Sessions-2021 FINAL.xlsx', r'C:/Users/vt238770/Documents/_CEA/_Formations/Extracts IRIS - Faits/Extracts originaux/R04110_Sessions-2022 FINAL.xlsx', r'C:/Users/vt238770/Documents/_CEA/_Formations/Extracts IRIS - Faits/Extracts originaux/R04110_Sessions-2023 FINAL.xlsx'))
env = TravauxFichiersIRIS.avecEcritureOutputDefaut(sessions, r'C:/Users/vt238770/Documents/_CEA/_Formations/Extracts IRIS - Faits/Extracts originaux/R04110_Sessions-2021 FINAL.xlsx')

#print(env)
