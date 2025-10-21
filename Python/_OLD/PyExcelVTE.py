# Exploitation Word
#https://pbpython.com/python-word-template.html

# Optimisation vitesse
#https://progresser-en-maths.com/la-maniere-la-plus-efficace-diterer-sur-des-lignes-dans-un-dataframe-pandas/
#https://pandas.pydata.org/docs/user_guide/enhancingperf.html
#https://moncoachdata.com/blog/7-techniques-doptimisation-de-la-memoire-avec-pandas/

#Tableaux structurés
#https://openpyxl.readthedocs.io/en/3.1/api/openpyxl.worksheet.table.html#module-openpyxl.worksheet.table
#https://openpyxl-readthedocs-io.translate.goog/en/3.1/worksheet_tables.html?_x_tr_sl=en&_x_tr_tl=fr&_x_tr_hl=fr&_x_tr_pto=sc




import pandas as pd
from openpyxl import load_workbook
#from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils.cell import get_column_letter
from openpyxl.utils.cell import range_boundaries
from openpyxl.utils.dataframe import dataframe_to_rows

import xlwings as xw # Avec xlwings, Excel gère automatiquement l’expansion du tableau, la mise en forme, et même les formules mais c'est plus long que openpyxl

from mailmerge import MailMerge
from datetime import date

import os
import time
import math
import re
from copy import copy
from tqdm import tqdm
from colorama import Fore, Style, init

from tkinter import filedialog

### --------------------------------------------------------------------
#  Définitions classes utilisateur
### --------------------------------------------------------------------
class fichierExcel:
    "C'est la classe qui contient les princnipales données et fonctions agir sur le fichier"

    # ====  Constructeurs
    def __init__(self, cheminFichier):
        self.__cheminFichier = "" # Chemin vers le fichier (chemin + fichier)
        self.__wb = load_workbook(cheminFichier)
        self.__ws = None

    @classmethod
    def ouvreWS(cls, cheminFichier, nom_ws):
        FE = cls(cheminFichier)
        cls.__ws = FE.__wb[nom_ws]
        return FE

    # ====  Méthodes
    def repertoire(self):
        """Donne le répertoire du fichier uniquement"""
        return os.path.dirname(self.__cheminFichier)
    
    def nomFichier(self):
        """Donne le nom du fichier uniquement"""
        return os.path.basename(self.__cheminFichier)

    # ====  Accesseurs / mutateurs
    @property
    def cheminFichier(self):
        """Accesseur pour l'attribut cheminFichier"""
        return self.__cheminFichier

    @cheminFichier.setter
    def cheminFichier(self, cheminFichier):
        """Mutateur pour l'attribut cheminFichier"""
        self.__cheminFichier = cheminFichier

class ExportIRIS_INFOS_SEULEMENT:
    """C'est la classe qui contient tous les éléments d'un type d'export IRIS"

            # J'ai besoin du codeExport pour connaitre le type et sélectionner mes actions

        # === ENTREE
        # j'ai un ou plusieurs fichiers à lire : (cheminFichier ou cheminsFichiers)
        #   pour mettre dans df_tableau
        #   pour mettre dans df_tableau et concaténer plus tard
        # (si plusieurs fichiers : faire une liste ?)

        # Pour ça j'ai besoin de connaître (si tableau qcq)
        #   le nom de l'onglet
        #   le nombre de lignes à sauter avant ET

        # Pour ça j'ai besoin de connaître (si tableau structuré)
        #   le nom de l'onglet
        #   le nom du tableau structuré

        # === SORTIE
        # j'ai un fichier modèle dans lequel écrire mais ne pas écraser
        # j'ai un fichier output (qui peut être tout nouveau ou déjà créé)


        # j'ai 
        """




class ExportIRIS:
    "C'est la classe qui contient tous les éléments d'un type d'export IRIS"
    
    def __init__(self, codeExport, nomsFichiers_exportIRIS = "", cheminRepertoire_exportIRIS = "", cheminRepertoire_exportsGED = "", format_nomExportsGED = "", nomOngletTableau = "", nbLignes_avantET_exportIRIS = -1, nbLignes_avantET_exportsGED = -1):






        # Initialisation des données obligatoirement spécifiées
        # nomsFichiers_exportIRIS doit être un tuple de strings. Si c'est un string c'est qu'un seul fichier a été donné. Alors on convertit en tuple
        if isinstance(nomsFichiers_exportIRIS, str):
            self.__nomsFichiers_exportIRIS = (nomsFichiers_exportIRIS,)
        else:
            self.__nomsFichiers_exportIRIS = nomsFichiers_exportIRIS
        
        self.__codeExport = codeExport
        self.__df_tableau = None # Lui je l'initialise car à un moment je dois faire un test s'il est rempli ou non

        # Initialisation des données facultatives
        # 1er cas par défaut (le plus générique): on initialise selon le type de rapport
        match self.__codeExport:
            # Cas Sessions
            case "R04110":
                # Informations de référence pour la GED
                self.__cheminRepertoire_exportsGED = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits" #Pas dans la GED
                self.__format_nomExportsGED="R04110_Sessions-Extract année*.xlsx"
                self.__nomOnglet_tableauGED = "Data"
                self.__nbLignes_avantET_exportsGED = 1

                # Informations du fichier à exploiter ou le plus complet ou qui concatene tous les fichiers IRIS
                self.__cheminRepertoire_exportIRIS = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits"
                self.__nomBase_exportReferenceIRIS = "R04110_Sessions-Extract COMPLET"
                self.__nomOnglet_tableauDonnesIRIS = "Sessions"
                self.__nbLignes_avantET_exportIRIS = 1

            # Cas Inscriptions
            case "R04500":
                # Informations de référence pour la GED
                self.__cheminRepertoire_exportsGED = r"\\instnt\HOME\REFERENC\IRIS - rapports de synthese"
                self.__format_nomExportsGED="R04500_Sessions-Inscriptions-*.xlsx"
                self.__nomOnglet_tableauGED = "Data"
                self.__nbLignes_avantET_exportsGED = 0

                # Informations du fichier à exploiter ou le plus complet ou qui concatene tous les fichiers IRIS
                self.__cheminRepertoire_exportIRIS = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits"
                self.__nomBase_exportReferenceIRIS = "R04500_Sessions-Inscriptions-COMPLET"
                self.__nomOnglet_tableauDonnesIRIS = "Inscriptions"
                self.__nbLignes_avantET_exportIRIS = 1
            
            # Cas Formations
            case "R0304":
                # Informations de référence pour la GED
                self.__cheminRepertoire_exportsGED = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits" #Pas dans la GED
                self.__format_nomExportsGED = "R0304_Ref_Formation-Listedesformations-*.xlsx"
                self.__nomOnglet_tableauGED = "Data"
                self.__nbLignes_avantET_exportsGED = 0

                # Informations du fichier à exploiter ou le plus complet ou qui concatene tous les fichiers IRIS
                self.__cheminRepertoire_exportIRIS = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits"
                self.__nomBase_exportReferenceIRIS = "R0304_Formations-Extract COMPLET"
                self.__nomOnglet_tableauDonnesIRIS = "Formations"
                self.__nbLignes_avantET_exportIRIS = 0
            
            # Cas Ventes : sensiblement comme 'Inscription R04500' mais groupé par Client (pas de détail de chaque stagiaire)     
            case "R04301":
                # Informations de référence pour la GED
                #self.__cheminRepertoire_exportsGED = r"\\instnt\HOME\REFERENC\IRIS - rapports de synthese"
                self.__cheminRepertoire_exportsGED = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits" #Tant qu'il y a des bugs dans la GED sur les fichiers R04301_Sessions-Ventes-FC2021 FINAL et R04301_Sessions-Ventes-FC2022 FINAL
                self.__format_nomExportsGED="R04301_Sessions-Ventes-*.xlsx"
                self.__nomOnglet_tableauGED = "Data"
                self.__nbLignes_avantET_exportsGED = 0

                # Informations du fichier à exploiter ou le plus complet ou qui concatene tous les fichiers IRIS
                self.__cheminRepertoire_exportIRIS = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits"
                self.__nomBase_exportReferenceIRIS = "R04301_Sessions-Ventes-Extract COMPLET"
                self.__nomOnglet_tableauDonnesIRIS = "Ventes"
                self.__nbLignes_avantET_exportIRIS = 1

            # Code session différent des autres cas
            case _:
                print("Code session non valide : aucune valeur par défaut renseignée")
                #exit()

        """
        def __init__(self, codeExport, cheminRepertoire_exportsGED, format_nomExportsGED, nomOngletTableau, nbLignes_avantET, cheminRepertoire_exportIRIS, nomFichier_exportIRIS):
            # Informations des fichiers sur la GED INSTN
            self.__codeExport = codeExport
            self.__cheminRepertoire_exportsGED = cheminRepertoire_exportsGED
            self.__format_nomExportsGED = format_nomExportsGED
            self.__nomOngletTableau = nomOngletTableau
            self.__nbLignes_avantET = nbLignes_avantET

            # Informations du fichier à exploiter ou le plus complet ou qui concatene tous les fichiers IRIS
            self.__cheminRepertoire_exportIRIS = cheminRepertoire_exportIRIS
            self.__nomFichier_exportIRIS = nomFichier_exportIRIS
            self.__cheminFichier_exportIRIS = cheminRepertoire_exportIRIS + "\\" + nomFichier_exportIRIS


        # Ancienne gestion par défaut
        if cheminRepertoire_exportIRIS == "": cheminRepertoire_exportIRIS = rep_extractIRIS_VTE
        if cheminRepertoire_exportsGED == "": cheminRepertoire_exportsGED = rep_extractIRIS_INSTNT
        if format_nomExportsGED == "": format_nomExportsGED = ""
        if nomOngletTableau == "": nomOngletTableau = nomGenerique_ongletExportIRISGED
        if nbLignes_avantET == -1: nbLignes_avantET = nbLignes_avantET_ExportsIRISGED


        # Informations du fichier à exploiter ou le plus complet ou qui concatene tous les fichiers IRIS
        self.__cheminRepertoire_exportIRIS = cheminRepertoire_exportIRIS
        self.__cheminFichier_exportIRIS = self.__cheminRepertoire_exportIRIS + "\\" + self.__nomFichier_exportIRIS

        # Informations des fichiers sur la GED INSTN
        self.__cheminRepertoire_exportsGED = cheminRepertoire_exportsGED
        self.__format_nomExportsGED = format_nomExportsGED
        self.__nomOngletTableau = nomOngletTableau
        self.__nbLignes_avantET = nbLignes_avantET
        """

        # 2eme cas par défaut : si spécifié explicitement dans la déclaration, alors on écrase les valeur 1er défaut par ce qui est spécifié
        if cheminRepertoire_exportIRIS != "": self.__cheminRepertoire_exportIRIS = cheminRepertoire_exportIRIS
        if cheminRepertoire_exportsGED != "": self.__cheminRepertoire_exportsGED = cheminRepertoire_exportsGED
        if format_nomExportsGED != "": self.__format_nomExportsGED = ""
        if nomOngletTableau != "": self.__nomOngletTableau = nomOngletTableau
        if nbLignes_avantET_exportIRIS != -1: self.__nbLignes_avantET_exportIRIS = nbLignes_avantET_exportIRIS
        if nbLignes_avantET_exportsGED != -1: self.__nbLignes_avantET_exportsGED = nbLignes_avantET_exportsGED
        
        # Informations du fichier à exploiter ou le plus complet ou qui concatene tous les fichiers IRIS #ne peut pas s'appliquer si l'on a un tuple comme fichiers Excel à traiter
        # self.__cheminFichier_exportIRIS = self.__cheminRepertoire_exportIRIS + "\\" + self.__nomFichier_exportIRIS

    @classmethod
    def avecLecture(cls, IHMSelectionFichiers = False, **kwargs):
        
        EI = cls(**kwargs)

        # Si demandé par l'utilisateur, on ouvre un filedialog
        if IHMSelectionFichiers:
            EI.choisirFichiers_filedialog()
        
        # On lit le/les extract IRIS et on stocke dans self.__df_tableau
        EI.lire_extractIRIS()
        return EI
        
    @classmethod
    def avecEcriture(cls, fichier_output = "", repertoire_output = "", nom_ws_output = "", nom_tableau_output = "", supprimeDonneesEtRemplace = False, nom_ws_imports = "", **kwargs):
        
        EI = cls.avecLecture(**kwargs)
        
        # On lit le/les extract IRIS et on stocke dans self.__df_tableau
        EI.ecrit_DataFrame_dans_tableauStructure(fichier_output = fichier_output, supprimeDonneesEtRemplace = supprimeDonneesEtRemplace, repertoire_output = repertoire_output, nom_ws_output = nom_ws_output, nom_tableau_output = nom_tableau_output, nom_ws_imports = nom_ws_imports)

        return EI



    ### --------------------------------------------------------------------
    #  Accesseurs / mutateurs (TODO)
    ### --------------------------------------------------------------------

    @property
    def nomsFichiers_exportIRIS(self):
        """Accesseur pour l'attribut nomsFichiers_exportIRIS"""
        return self.__nomsFichiers_exportIRIS

    @property
    def df_tableau(self):
        """Accesseur pour l'attribut df_tableau"""
        return self.__df_tableau


    #@cheminFichier_exportIRIS.setter
    #def cheminFichier_exportIRIS(self, cheminFichier_exportIRIS):
    #    """Mutateur pour l'attribut cheminFichier_exportIRIS"""
    #    self.__cheminFichier_exportIRIS = cheminFichier_exportIRIS




    ### --------------------------------------------------------------------
    #  Méthodes de la classe
    ### --------------------------------------------------------------------
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
        taille_totale = sum(os.path.getsize(os.path.join(self.__cheminRepertoire_exportIRIS, fichier)) for fichier in self.__nomsFichiers_exportIRIS) # Calcul taille totale pour barre de progression

        #with tqdm(total=taille_totale, unit='o', unit_scale=True, desc=f"{cyan}Lecture des fichiers Excel{resetCouleur}") as pbar:
        with tqdm(total=taille_totale, unit='o', unit_scale=True, desc=Fore.CYAN+"Lecture des fichiers Excel" + Style.RESET_ALL) as pbar:
            for i, ifichier in enumerate(self.__nomsFichiers_exportIRIS, 1):
                # On crée le chemin complet du fichier
                chemin_complet = os.path.join(self.__cheminRepertoire_exportIRIS, ifichier)

                #On crée les indicateurs pour tqdm
                taille = os.path.getsize(chemin_complet)

                # Affichage dynamique dans la barre
                pbar.set_postfix(file=ifichier, progress=f"{i}/{len(self.__nomsFichiers_exportIRIS)}")
                
                # On lit l'excel ; skiprows = 1 car l'extract IRIS a une première ligne que l'on doit sauter
                df = pd.read_excel(chemin_complet, skiprows=self.__nbLignes_avantET_exportIRIS)
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

    def extraire_infos_numSessionIRIS_BAK(self, reference):
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

        # Sécurité : vérifier qu'on a au moins 6 blocs
        if len(blocs) < 6:
            print("Problème : il y a moins de 6 blocs : ")
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
        type_formation = blocs[2][:2]
        annee_match = re.search(r'\d+', blocs[2][2:])
        annee = int(annee_match.group()) if annee_match else None

        trigramme_RP = blocs[-1].rstrip('_')
        trigramme_AF = blocs[-2].rstrip('_')
        trigramme = blocs[-3].rstrip('_') if len(blocs) > 5 else None

        # 3e élément uniquement si on a 7 blocs
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

    def ecrit_DataFrame_dans_tableauStructure(self, fichier_output = "", supprimeDonneesEtRemplace = False, repertoire_output = "", nom_ws_output = "", nom_tableau_output = "", nom_ws_imports = ""):
        """
        On écrit le dataFrame dans un fichier output spécifié par un chemin, un nom de fichier, un nom d'onglet et un nom de tableau.
        On peut lui dire s'il faut remplacer les données ou écrire à la suite.

        :Example:
        >>> self.ecrit_DataFrame_dans_tableauStructure()

        .. seealso:: Rien du tout.
        .. warning:: Rien du tout.
        .. note:: Rien du tout.
        .. todo:: Rien du tout.        
        """
        
        # On écrit un DataFrame dans un fichier
        # Le fichier output peut-être quelconque → si pas de réf, alors on prend celui donné en argument
        # On peut soit vouloir concaténer le DataFrame à l'Excel, soit le remplacer
        # 
        
        # Cas par défaut
        if fichier_output == "": fichier_output = self.__nomBase_exportReferenceIRIS + ".xlsx" # Si pas de fichier_output, alors on veut coller dans l'export de référence IRIS
        if repertoire_output == "": repertoire_output = self.__cheminRepertoire_exportIRIS # Si pas de repertoire_output, alors on veut coller dans l'export de référence IRIS
        if nom_ws_output == "": nom_ws_output = self.__nomOnglet_tableauDonnesIRIS # Si pas de nom_ws, alors on veut coller dans l'export de référence IRIS
        if nom_tableau_output == "": nom_tableau_output = self.__nomOnglet_tableauDonnesIRIS # Si pas de nom_tableau, alors on veut coller dans l'export de référence IRIS (par défaut, nom du tableau et nom de l'onbglet sont les mêmes)


        #On vérifie que le DataFrame avec les données existe
        if self.__df_tableau is None:
            print("Le dataFrame n'a pas encore été créé")
            exit()
        
        # Définition du nom du classeur de base (i.e. le modèle) qui va servir à la sortie par la suite
        s_output = os.path.join(repertoire_output, fichier_output)

        # On ouvre le classeur qui va recevoir la concaténation
        wb_output = load_workbook(filename = s_output)

        # On écrit dans le tableau structuré 
        writeDataFrameInStructuredRef_openpyxl(self.__df_tableau, wb_output, nom_ws_output, supprimeDonneesEtRemplace = supprimeDonneesEtRemplace)

        # On écrit les références des fichiers copiés dans le tableau structuré "Imports"
        if nom_ws_imports != "":
            # Si un seul fichier est sélectionné, 'chemins' est une string, donc on convertit en tuple
            #if isinstance(tupleCheminsExcel, str):
            #    tupleCheminsExcel = (tupleCheminsExcel,)

            df_chemins = pd.DataFrame(self.__nomsFichiers_exportIRIS, columns=['Chemin fichier'])
            writeDataFrameInStructuredRef_openpyxl(df_chemins, wb_output, nom_ws_imports, supprimeDonneesEtRemplace = supprimeDonneesEtRemplace)
            #ws = wb_output[nom_ws_imports]
            #table = ws.tables[nom_ws_imports]
            #table.range.end('down').offset(1, 0).value = ['val1']

        #On enregistre et on ferme
        if fichier_output == self.__nomBase_exportReferenceIRIS + ".xlsx": #Alors c'est l'export de référence qu'on concatène et on veut le dater
            wb_output.save(filename=os.path.join(repertoire_output, self.__nomBase_exportReferenceIRIS + "-" + date.today().strftime("%Y.%m.%d") + ".xlsx")) #ou f"{datetime.now():%Y.%m.%d}")
        else: # Alors on sauvegarde sur le fichier output tel que désigné)
            wb_output.save(filename=s_output)

        #On fermele workbook    
        wb_output.close()

    def ecrit_DataFrame_dans_extractReference(self, supprimeDonneesEtRemplace = False, nom_ws_imports = "", tupleCheminsExcel = ""):
        """
        On écrit le dataFrame dans l'extract IRIS de référence

        :Example:
        >>> self.ecritRemplace_DataFrame_dans_extractReference()

        .. seealso:: Rien du tout.
        .. warning:: Rien du tout.
        .. note:: Rien du tout.
        .. todo:: Rien du tout.        
        """
        
        #On vérifie que le DataFrame avec les données existe
        if self.__df_tableau is None:
            print("Le dataFrame n'a pas encore été créé")
            exit()
        
        # Définition du nom du classeur de base qui va servir à la sortie par la suite
        s_output = self.__cheminRepertoire_exportIRIS + "\\" + self.__nomBase_exportReferenceIRIS + ".xlsx"

        # On ouvre le classeur qui va recevoir la concaténation
        wb_output = load_workbook(filename = s_output)

        # On écrit dans le tableau structuré 
        writeDataFrameInStructuredRef_openpyxl(self.__df_tableau, wb_output, self.__nomOnglet_tableauDonnesIRIS, supprimeDonneesEtRemplace)

        # On écrit les références des fichiers copiés dans le tableau structuré "Imports"
        if nom_ws_imports != "":
            # Si un seul fichier est sélectionné, 'chemins' est une string, donc on convertit en tuple
            if isinstance(tupleCheminsExcel, str):
                tupleCheminsExcel = (tupleCheminsExcel,)

            df_chemins = pd.DataFrame(tupleCheminsExcel, columns=['Chemin fichier'])
            writeDataFrameInStructuredRef_openpyxl(df_chemins, wb_output, nom_ws_imports)
            #ws = wb_output[nom_ws_imports]
            #table = ws.tables[nom_ws_imports]
            #table.range.end('down').offset(1, 0).value = ['val1']

        #On enregistre et on ferme
        wb_output.save(filename=self.__cheminRepertoire_exportIRIS + "\\" + self.__nomBase_exportReferenceIRIS + "-" + date.today().strftime("%Y.%m.%d") + ".xlsx") #ou f"{datetime.now():%Y.%m.%d}")
        wb_output.close()

    def choisirFichiers_filedialog(self):
        # Lister/sélectionner les documents à concaténer
        cheminsExcel = filedialog.askopenfilename(title="Sélectionner les fichiers " + self.__codeExport + " Excel à concaténer", filetype=[("Fichiers Excel", "*.xlsx")], initialdir=optimiseCheminRepertoire(self.__cheminRepertoire_exportsGED), multiple=True)
        
        # Gestion du cas où il y a non-sélection de fichiers
        if not cheminsExcel:
            exit()
        
        # Convertir le tuple original en liste pour manipulation
        cheminsExcel = list(cheminsExcel)

        # Extraire le répertoire commun (s’il y en a au moins un)
        self.__cheminRepertoire_exportIRIS = os.path.dirname(cheminsExcel[0]) if cheminsExcel else ""

        # Extraire uniquement les noms de fichiers et mise en tuple
        self.__nomsFichiers_exportIRIS = tuple(os.path.basename(f) for f in cheminsExcel)

class Timer:
    def __init__(self, description = ""):
        self.__debut = 0
        self.__fin = 0
        self.__duree = 0
        self.__description = description

    def debut(self, description = ""):
        self.__debut = time.time()
        self.__description = description

    def fin(self):
        self.__fin = time.time()
        self.__duree = self.__fin - self.__debut

        # Affichage lisible du temps
        minutes, secondes = divmod(int(self.__duree), 60)
        
        if(self.__description == ""):
            print(f"✅ Traitement terminé en {minutes} min {secondes} s.")
        else:
            print(f"✅ {self.__description} terminé en {minutes} min {secondes} s.")
 
class BilanFormation:
    "C'est la classe qui contient tous les éléments de ma formation pour mon bilan"
 
    def __init__(self, codeFormation, annee, s_extractIRIS_sessions, s_extractIRIS_formations, s_fdc, s_specsPedagogiques):
        self.__codeFormation = codeFormation
        self.__annee = annee
        self.__s_extractIRIS_sessions = s_extractIRIS_sessions
        self.__s_extractIRIS_formations = s_extractIRIS_formations
        self.__s_fdc = s_fdc
        self.__s_specsPedagogiques = s_specsPedagogiques

        self.__lieuxFormation = ""


        #####
        # Exploitation de l'extract IRIS sessions
        #####
        self.__df_sessionsIRIS = lire_extractIRIS_sessionsR04110(self.__s_extractIRIS_sessions)

        # Pour initialiser les valeurs communes, déjà on filtre le Dataframe principal avec le code formation
        #df_sorted = self.__df_sessionsIRIS[self.__df_sessionsIRIS['Trigramme formation'] == self.__codeFormation]
        df_sorted = self.__df_sessionsIRIS[(self.__df_sessionsIRIS['Trigramme formation'] == self.__codeFormation) & (self.__df_sessionsIRIS['Année début ses.'] == self.__annee)]

        # Puis on récupère la dernière ligne pour avoir les valeurs les plus à jour (tri par index)
        [self.__titreFormation, self.__dureeHeures, self.__dureeJours, self.__minIRIS, self.__maxIRIS, self.__depassementAutoriseIRIS] = df_sorted[["Session", "Durée planif. (H.)", "Durée planif. (J.)", "Min.", "Max.", "Dépass. autorisé"]].iloc[-1]


        # Pour obtenir la liste des RP et de leurs lieux
        for rp in df_sorted["Nom responsable pédag."].unique():
            #print("\n\nRP = " + rp)
            [prenomRP, nomRP, lieuRP] = df_sorted[self.__df_sessionsIRIS["Nom responsable pédag."] == rp][["Prénom responsable pédag.", "Nom responsable pédag.", "Lieu principal"]].iloc[-1]
            self.__lieuxFormation += prenomRP + " " + nomRP + " (" + lieuRP + "), "
        self.__lieuxFormation = self.__lieuxFormation[:-2]
        #print(self.lieuxFormation)




        #####
        # Exploitation des specs pédagogiques
        #####

        self.__dateSpecs = time.localtime(os.path.getmtime(self.__s_specsPedagogiques))
        self.__sDateSpecs = f"{self.__dateSpecs[2]:02}/{self.__dateSpecs[1]:02}/{self.__dateSpecs[0]:04}"

        

        #####
        # Exploitation de la fiche de coûts
        #####
        self.__dateFdC = time.localtime(os.path.getmtime(self.__s_fdc))
        self.__sDateFdC = f"{self.__dateFdC[2]:02}/{self.__dateFdC[1]:02}/{self.__dateFdC[0]:04}"
        
        self.__df_fdc_infos, self.__df_fdc_couts, self.__prixVenteRetenuParParticipant, self.__dateCreationFormation, self.__dureeJours_fdc, self.__osThematique, self.__nbCible_fcd = lire_fdc(self.__s_fdc)
        




        #####
        # Exploitation de l'extract IRIS formation TODO
        #####
        #Date création formation
        #Type de reconnaissance (Autre, Certification, Diplôme)
        #Formation habilitante (bool)
        #Reconnu au RNCP
        #Eligible CPF
        #Reconnue au RS


        
    def __init__(self, codeFormation, annee, s_extractIRIS_sessions, s_extractIRIS_formations, s_fdc, s_specsPedagogiques):
        print()



    ### --------------------------------------------------------------------
    #  Méthodes de la classe
    ### --------------------------------------------------------------------
    def mergeBilan(self, s_word_bilan_input, s_word_bilan_output):
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

        document = MailMerge(s_word_bilan_input)
        #print(document.get_merge_fields())

        document.merge(
            annee='{:%Y}'.format(date.today()),
            codeFormation=self.__codeFormation,
            titreFormation=self.__titreFormation,
            
            lienGED = r'file:///\\\\instnt\\PARTAGE\\FORMATIONS_C\\ACI\\',
            osThematique = self.__osThematique,
            lieuxFormation = self.__lieuxFormation,
            dureeJours = f"{self.__dureeJours:.1f}",
            dureeHeures = f"{self.__dureeHeures:.2f}",
            dateCreationFormation = str(self.__dateCreationFormation), #Réaffecter à partir extract IRIS formations

            dateSpecs = self.__sDateSpecs,

            dateFdC = self.__sDateFdC,
            minFdC = f"{self.__df_fdc_couts.loc['Valeur fixée', 'Min participants T3']:d} p.",
            cibleFdC = f"{self.__nbCible_fcd:d} p.",
            minIRIS = f"{self.__minIRIS:d} p.",
            cibleIRIS = f"{self.__maxIRIS:d} p.",
            maxIRIS = f"{self.__depassementAutoriseIRIS:d} p.",

            PT1_1 = f"{self.__df_fdc_couts.loc['T1', 'Montant cible par participant']:.0f} €/p.",
            PT1_2 = f"{self.__df_fdc_couts.loc['T1', 'Montant cible par participant et par jour']:.0f} €/j/p.",
            PT1_3 = f"{self.__df_fdc_couts.loc['T1', 'Min participants T1']:d} p.",

            PT3_1 = f"{self.__df_fdc_couts.loc['T3', 'Montant cible par participant']:.0f} €/p.",
            PT3_2 = f"{self.__df_fdc_couts.loc['T3', 'Montant cible par participant et par jour']:.0f} €/j/p.",
            PT3_3 = f"{self.__df_fdc_couts.loc['T3', 'Min participants T1']:d} p.",
            PT3_4 = f"{self.__df_fdc_couts.loc['T3', 'Min participants T3']:d} p.",

            PTR_1 = f"{self.__df_fdc_couts.loc['Valeur fixée', 'Montant cible par participant']:.0f} €/p.",
            PTR_2 = f"{self.__df_fdc_couts.loc['Valeur fixée', 'Montant cible par participant et par jour']:.0f} €/j/p.",
            PTR_3 = f"{self.__df_fdc_couts.loc['Valeur fixée', 'Min participants T1']:d} p.",
            PTR_4 = f"{self.__df_fdc_couts.loc['Valeur fixée', 'Min participants T3']:d} p."
            
            )
        
        document.write(s_word_bilan_output)

   



### --------------------------------------------------------------------
#  Définitions fonctions utilisateur
### --------------------------------------------------------------------


def lire_extractIRIS_sessionsR04110(s_extractIRIS_sessions):
    """
        A partir d'un chemin de fichier Excel (qui se doit d'etre un extract IRIS de session R04110, on retourne un dataframe.
        On traite les colonnes de sorte a extraire les informations de la colonne 'N° Session'
 
 
        :param s_extractIRIS_sessions: Chemin du fichier Excel a ouvrir (se doit d'etre un extract IRIS R04110_Sessions)
        :type s_extractIRIS_sessions: string
        :return: Un DataFrame de l'extract IRIS avec ajouts de colonnes (on a extrait les informations de la colonne 'N° Session' par decoupage)
        :rtype: DataFrame
 
        :Example:
 
        >>> DataFrame.df_sessionsIRIS = lireExtractIRIS_sessions("C:\\Users\\fichier.xlsx")

 
        .. seealso:: Rien du tout.
        .. warning:: Rien du tout.
        .. note:: Rien du tout.
        .. todo:: Rien du tout.
    """
    
    # On lit l'excel ; skiprows = 1 car l'extract IRIS a une première ligne que l'on doit sauter
    df_sessionsIRIS = pd.read_excel(s_extractIRIS_sessions, skiprows=1) 

    # On découpe la colonne N° Session pour avoir le détail d'éléments supplémentaires (trigrammes)
    #df[['S', 'Numéro IRIS', 'Type formation et année', 'Trigramme formation', 'Trigramme RP', 'Trigramme AF']] = df['N° Session'].str.split('-', expand=True)
    df_sessionsIRIS['Numéro IRIS'] = df_sessionsIRIS['N° Session'].astype(str).str[2:7]
    df_sessionsIRIS['Type formation'] = df_sessionsIRIS['N° Session'].astype(str).str[8:10]
    df_sessionsIRIS['Trigramme AF'] = df_sessionsIRIS['N° Session'].astype(str).str[-3:] #tout sauf 3 derniers caract
    df_sessionsIRIS['Trigramme RP'] = df_sessionsIRIS['N° Session'].astype(str).str[-7:-4] #De -7 à -4
    df_sessionsIRIS['Trigramme formation'] = df_sessionsIRIS['N° Session'].astype(str).str[-11:-8]

    return(df_sessionsIRIS)

def lire_fdc(s_fdc):

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
    df_fdc_infos = pd.read_excel(s_fdc, sheet_name=nomOnglet_fdc, usecols=[1, 2], names=["Critere", "Valeur"], header=3, nrows=17)
    #print(df_fdc_infos)
    dateCreationFormation = df_fdc_infos.iloc[4, 1] #C9
    dureeJours_fdc = df_fdc_infos.iloc[6, 1] #C11
    osThematique = df_fdc_infos.iloc[8, 1] #C13
    nbCible_fcd = df_fdc_infos.iloc[15, 1] #C17

    # Prix de vente défini par le RP (cellule J22)
    prixVenteRetenuParParticipant = pd.read_excel(s_fdc, sheet_name=nomOnglet_fdc, usecols=[9], names=["Valeur"], skiprows=20, nrows=1).iloc[0,0]
    
    # Deuxième tableau : tableau des coûts (tout compris) et des prix par personne (T1, T2 et T3)
    df_fdc_couts = pd.read_excel(s_fdc, sheet_name=nomOnglet_fdc, usecols=[10, 11, 12, 13, 14, 15], names=["T1", "T3", "T2", "0.9xT3", "1.1xT3", "Valeur fixée"], header=33, nrows=2).transpose() #Grosse astuce : je mets 2 lignes de plus pour affecter les noms plus facilement et je les recalculerai après
    df_fdc_couts.rename(columns={0: "Couts PF et PV nb cible"}, inplace=True) # Pour changer le nom de la colonne après transposition
    df_fdc_couts.head() # Requis pour MàJ le nom de la colonne après transposition

    df_fdc_couts.loc["0.9xT3", "Couts PF et PV nb cible"] = df_fdc_couts.loc["T3", "Couts PF et PV nb cible"] * 0.9
    df_fdc_couts.loc["1.1xT3", "Couts PF et PV nb cible"] = df_fdc_couts.loc["T3", "Couts PF et PV nb cible"] * 1.1
    df_fdc_couts.loc["Valeur fixée", "Couts PF et PV nb cible"] = prixVenteRetenuParParticipant / 1.1 * nbCible_fcd # Asrtuce : comme c'est une valeur que je n'ai pas, je fais le calcul inverse que pour avoir le montant par session avec aleas


    # On ajoute les colonnes montants cibles avec calculs
    l1 = []
    l2 = []
    l3 = []
    for index, row in df_fdc_couts.iterrows():
        l1.append(row["Couts PF et PV nb cible"] * 1.1)
        l2.append(row["Couts PF et PV nb cible"] * 1.1 / nbCible_fcd)
        l3.append(row["Couts PF et PV nb cible"] * 1.1 / nbCible_fcd / dureeJours_fdc)
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

def writeDataFrameInStructuredRef_openpyxl(df, wb, nom_ws, nom_table = "", supprimeDonneesEtRemplace = False):
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
    if nom_table == "": nom_table = nom_ws

    # On ouvre la feuille et le tableau structuré
    ws = wb[nom_ws]
    #print(ws)
    table = ws._tables[nom_table] #Tableau structuré nommé

    # Infos de longueurs de mon tableau
    min_col, min_row, max_col, max_row = range_boundaries(table.ref)
    total_rows = max_row - min_row + 1
    data_rows = total_rows - table.headerRowCount
    #print(table.ref, total_rows, data_rows, min_col, min_row, max_col, max_row)
    
    # On écrit toutes les autres lignes une par une (on garde les lignes initiales pour garder le format qu'on copiera)
    # Méthode 1 qui marche
    #for il in tqdm(df.itertuples(), desc="Ecriture output Excel"):
    total_lignes = len(df)
    with tqdm(total=total_lignes, unit=' ligne', desc=Fore.CYAN + "Écriture des lignes dans l'output" + Style.RESET_ALL) as pbar:
        for i, il in enumerate(df.itertuples(), 1):
            pbar.set_postfix(progress=f"{i}/{len(df)}")
            row = [val if pd.notna(val) else None for val in il[1:]] #Je dois rajouter cette ligne car il faut tester si je n'ai pas de valeurs <NA> qu'il faut retravailler sinon ça plante
            ws.append(row)
            pbar.update(1)
    
    # Méthode 2 - Bug
    #for r in dataframe_to_rows(df, index=False, header=False):
    #    ws.append(r)
    
    # On redéfinit le dimensionnement du tableau (/!\ +1 ligne pour récupérer le format de la dernière ligne)
    table.ref = f"{ws.cell(row=min_row, column=min_col).coordinate}:{ws.cell(row=max_row + len(df), column=max_col).coordinate}" #On saut le nb de ligens avant l'en-tête, puis l'en-tête, puis on va à la première ligne de données nbLignes_avantET1+1+1)
    #table.ref = "A1:{}{}".format(lettreFinTableau, len(df)+1) #+1 car on a la ligne d'en-tête #Méthodo initiale qui requiert de connaître la lettre de fin du tableau

    # On copie le format sur toutes les nouvelles lignes du tableau
    copieFormatTableauStructure_openpyxl(ws, table, indexLigneSourceFormat = min_row + 1, indexLigneDebutCopie = max_row + 1, indexLigneFinCopie = max_row + len(df) +1)

    # Si désiré par l'utilisateur, alors on supprime les anciennes lignes de la feuille Excel (ça garde la dimension initiale du tableau structuré)
    if supprimeDonneesEtRemplace:
        # Suppression des anciennes lignes
        ws.delete_rows(idx=min_row + 1, amount=data_rows)

        # On redéfinit les dimensions du tableau structuré
        table.ref = f"{ws.cell(row=min_row, column=min_col).coordinate}:{ws.cell(row=min_row + len(df), column=max_col).coordinate}" #On saut le nb de ligens avant l'en-tête, puis l'en-tête, puis on va à la première ligne de données nbLignes_avantET1+1+1)


def writeDataFrameInStructuredRef_openpyxl_xlwings(df, wb, chemin_wb, nom_ws, nom_table = "", supprimeDonneesEtRemplace = False):
    """
    Ecrit un DataFrame dans un tableau structuré existant d'une feuille de calcul  
    Utilise openpyxl pour l’écriture des données, puis xlwings pour copier rapidement le format.

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
    if nom_table == "": nom_table = nom_ws

    # On ouvre la feuille et le tableau structuré
    ws = wb[nom_ws]
    table = ws._tables[nom_table] #Tableau structuré nommé

    # Infos de dimension du tableau
    min_col, min_row, max_col, max_row = range_boundaries(table.ref)
    total_rows = max_row - min_row + 1
    data_rows = total_rows - table.headerRowCount
    #print(table.ref, total_rows, data_rows, min_col, min_row, max_col, max_row)
    
    # On écrit toutes les autres lignes une par une (on garde les lignes initiales pour garder le format qu'on copiera)
    # Méthode 1 qui marche
    #for il in tqdm(df.itertuples(), desc="Ecriture output Excel"):
    total_lignes = len(df)
    with tqdm(total=total_lignes, unit=' ligne', desc=Fore.CYAN + "Écriture des lignes dans l'output" + Style.RESET_ALL) as pbar:
        for i, il in enumerate(df.itertuples(), 1):
            pbar.set_postfix(progress=f"{i}/{len(df)}")
            row = [val if pd.notna(val) else None for val in il[1:]] #Je dois rajouter cette ligne car il faut tester si je n'ai pas de valeurs <NA> qu'il faut retravailler sinon ça plante
            ws.append(row)
            pbar.update(1)
    
    # Méthode 2 - Bug
    #for r in dataframe_to_rows(df, index=False, header=False):
    #    ws.append(r)
    
    # Redimensionnement du tableau 
    nouvelle_max_row = max_row + len(df)
    table.ref = f"{ws.cell(row=min_row, column=min_col).coordinate}:{ws.cell(row=nouvelle_max_row, column=max_col).coordinate}"

    # Nécessaire avant d'utiliser xlwings
    wb.save(chemin_wb)  
    wb.close()

    # ==== Copie du format avec xlwings ====
    # Définir la plage source (= ligne de format) et plage cible (= nouvelles lignes)
    col_lettre_debut = get_column_letter(min_col)
    col_lettre_fin = get_column_letter(max_col)

    range_modele = f"{col_lettre_debut}{min_row+1}:{col_lettre_fin}{min_row+1}"  # première ligne de données
    range_cible = f"{col_lettre_debut}{max_row+1}:{col_lettre_fin}{nouvelle_max_row}"  # nouvelles lignes

    # On copie le format sur toutes les nouvelles lignes du tableau
    copierFormat_xlwings(chemin_wb, nom_ws, range_modele, range_cible)
    
    # Réouvrir pour finaliser les suppressions éventuelles
    wb = load_workbook(filename=chemin_wb)
    ws = wb[nom_ws]
    table = ws._tables[nom_table]

    # Si désiré par l'utilisateur, alors on supprime les anciennes lignes de la feuille Excel (ça garde la dimension initiale du tableau structuré)
    if supprimeDonneesEtRemplace:
        # Suppression des anciennes lignes
        ws.delete_rows(idx=min_row + 1, amount=data_rows)

        # On redéfinit les dimensions du tableau structuré
        table.ref = f"{ws.cell(row=min_row, column=min_col).coordinate}:{ws.cell(row=min_row + len(df), column=max_col).coordinate}" #On saut le nb de ligens avant l'en-tête, puis l'en-tête, puis on va à la première ligne de données nbLignes_avantET1+1+1)


def copieFormatTableauStructure_openpyxl(ws, table, indexLigneSourceFormat = 1, indexLigneDebutCopie = -1, indexLigneFinCopie = -1) :
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
    min_col, min_row, max_col, max_row = range_boundaries(table.ref)
    
    # Initialisattion paramètres non renseignés
    if indexLigneDebutCopie == -1: indexLigneDebutCopie = indexLigneSourceFormat + 1
    if indexLigneFinCopie == -1: indexLigneFinCopie = max_row


    # On récupère les formats de chaque cellule de la première ligne du tableau structuré 
    formats=[]
    for icol in range(min_col, max_col+1) :
        formats.append(ws.cell(indexLigneSourceFormat, icol))
        #print(formats[icol-1].number_format)
    #print(formats)

    # On copie colle les formats avec ces cellules
    with tqdm(total=max_col, unit=' colonnes', desc=Fore.CYAN + "Copie des formats" + Style.RESET_ALL, ncols=150) as pbar:
        for icol in range(min_col, max_col+1) :
            pbar.set_postfix(progress=f"{icol}/{max_col}")
            source = formats[icol-1] #Je prends un index de liste et pas un numéro de colonne, donc -1
            #print(source.number_format, source.number_format == "General")

            # Optimisation : on ne fait les copies que si le format est différent de General
            #if source.number_format != "General":
            for il in range(indexLigneDebutCopie, indexLigneFinCopie) : #+1 pour le row car en-tête
                #print(il, icol, ws.cell(row=il, column=icol).value, source.number_format)
                target = ws.cell(row=il, column=icol)
                #target.font = copy(source.font)
                #target.border = copy(source.border)
                target.fill = copy(source.fill)
                #target.alignment = copy(source.alignment)
                #target.protection = copy(source.protection)
                target.number_format = copy(source.number_format)

            pbar.update(1)

#concateneExcels(rep_extractIRIS_VTE, "_A")
#concateneExcels(rep_defaut = ".", nomBaseFichier = "Excel-output", nbLignes_avantET = 1)

def concatene_ongletsExcels_openpyxl(nom_ws, nom_ws_imports = "", rep_defaut = ".", rep_output = "", nomBaseFichier = "Excel-output", nbLignes_avantET = 0):
    """
    Concatene des fichiers Excel avec une même stucture 

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

    if rep_output == "": rep_output=rep_defaut


    # Lister/sélectionner les documents à concaténer
    listeCheminsExcel = filedialog.askopenfilename(title="Sélectionner les fichiers Excel à concaténer", filetype=[("fichiers excel","*.xlsx")], initialdir=rep_defaut, multiple=True)

    # On met les fichiers input dans un DataFrame
    df_input = pd.concat((pd.read_excel(iFichier, skiprows=nbLignes_avantET) for iFichier in listeCheminsExcel), ignore_index=True)

    # Définition du nom du classeur de base qui va servir à la sortie par la suite
    s_output = rep_output + "\\" + nomBaseFichier + ".xlsx" #ou f"{datetime.now():%Y.%m.%d}"
    #print(s_output)

    # On copie le classeur
    #shutil.copy(s_classeurIni, s_classeurDestination)

    # On ouvre le classeur qui va recevoir la concaténation
    wb_output = load_workbook(filename = s_output)
    
    # On met les fichiers input dans un DataFrame
    df_input = pd.concat((pd.read_excel(rep_defaut + "\\" + iFichier, skiprows=nbLignes_avantET) for iFichier in listeCheminsExcel), ignore_index=True)

    # On écrit dans le tableau structuré 
    writeDataFrameInStructuredRef_openpyxl(df_input, wb_output, nom_ws)

    # On écrit les références des fichiers copiés dans le tableau structuré "Imports"
    if nom_ws_imports != "":
        ws = wb_output.sheets[nom_ws_imports]
        table = ws.tables[nom_ws_imports]
        table.range.end('down').offset(1, 0).value = ['val1']

    #On enregistre et on ferme
    wb_output.save(filename=rep_output + "\\" + nomBaseFichier + "-" + date.today().strftime("%Y.%m.%d") + ".xlsx") #ou f"{datetime.now():%Y.%m.%d}")
    wb_output.close()

def concatene_ongletsExcels_xlwings(nom_ws, nom_ws_imports = "", rep_defaut = ".", rep_output = "", nomBaseFichier = "Excel-output", nbLignes_avantET = 0):
    """
    Concatene des fichiers Excel avec une même stucture.
    Si rep_output == "", alors rep_output=rep_defaut
    nom_ws_imports est facultatif, c'est juste si l'on souhaite sauvegarder la liste des fichiers concatenes

    :param nom_ws: nom du worksheet
    :type nom_ws: str
    :param nom_ws_imports: nom du worksheet ou on enregistre les fichiers excel qui vont etre concatene (defaut = "")
    :type nom_ws_imports: str
    :param rep_defaut: repertoire par defaut ou on va aller chercher les fichiers excel a concatener (defaut = ".")
    :type rep_defaut: str
    :param rep_output: repertoire ou on va enregistrer le fichier excel contenant la concatenation des fichiers (defaut = "", si rep_output == "", alors rep_output=rep_defaut)
    :type rep_output: str
    :param nomBaseFichier: nom du fichier output (defaut = "Excel-output")
    :type nomBaseFichier: str
    :param nbLignes_avantET: nombre de ligne avant l'en-tete des tableaux que l'on va concatener
    :type nbLignes_avantET: int




    :Example:
    >>> concatene_ongletsExcels("Inscriptions", nom_ws_imports = "Imports", rep_defaut = rep_extractIRIS_INSTNT, rep_output = rep_extractIRIS_VTE, nomBaseFichier = "R04500_Sessions-Inscriptions-COMPLET", nbLignes_avantET = 1)


    .. seealso:: Rien du tout.
    .. warning:: Rien du tout.
    .. note:: si rep_output == "", alors rep_output=rep_defaut
    .. todo:: C'est tres long, il faudra que je teste les vitesses entre xlwings et openpyxl.
    """

    if rep_output == "": rep_output=rep_defaut

    #Lister/sélectionner les documents à concaténer
    listeCheminsExcel = filedialog.askopenfilename(title="Sélectionner les fichiers Excel à concaténer", filetype=[("fichiers excel","*.xlsx")], initialdir=rep_defaut, multiple=True)

    # On met les fichiers input dans un DataFrame
    #for iFichier in listeCheminsExcel :
    #    df_input = pd.read_excel(rep_defaut + "\\" + iFichier, skiprows=nbLignes_avantET)
    df_input = pd.concat((pd.read_excel(iFichier, skiprows=nbLignes_avantET) for iFichier in listeCheminsExcel), ignore_index=True)
    
    # On met le dataframe en liste pour l'envoyer à ajouter_lignes_tableau_xlwings
    donnees = df_input.values.tolist()
    
    # Forme à avoir
    #donnees = [
    #    ['Jean', 'Durand', 'jean@example.com', 42],
    #    ['Claire', 'Martin', 'claire@example.com', 35]
    #]
    #donnees = [[2], [4]]
    #donnees = df.values.tolist()


    # Définition du nom du classeur de base qui va servir à la sortie par la suite
    s_output = rep_output + "\\" + nomBaseFichier + ".xlsx" #ou f"{datetime.now():%Y.%m.%d}"
    #print(s_output)

    # On ouvre l'app pour xlwings
    app = xw.App(visible=False)  # Excel s'ouvre en arrière-plan
        
    # On copie le classeur
    #shutil.copy(s_classeurIni, s_classeurDestination)

    # On ouvre le classeur qui va recevoir la concaténation
    wb_output = app.books.open(s_output)

    # On écrit dans le tableau structuré 
    #writeDataFrameInStructuredRef_openpyxl(df_input, wb_output, nom_ws) #Marche
    ajouter_lignes_tableau_xlwings(donnees, wb_output, nom_ws) #Tres long (~20 minutes pour mon test avec toutes les sessions)


    # On écrit les références des fichiers copiés dans le tableau structuré "Imports" ssi il y a un nom dans nom_ws_imports
    if nom_ws_imports != "":
        ajouter_lignes_tableau_xlwings(tuple_vers_liste_de_listes(listeCheminsExcel), wb_output, nom_ws_imports)

    #On enregistre et on ferme
    #wb_output.save(filename=rep_defaut + "\\" + nomBaseFichier + "-" + date.today().strftime("%Y.%m.%d") + ".xlsx") #ou f"{datetime.now():%Y.%m.%d}")
    wb_output.save(rep_output + "\\" + nomBaseFichier + "-" + date.today().strftime("%Y.%m.%d") + ".xlsx") #ou f"{datetime.now():%Y.%m.%d}")
    wb_output.close()
    app.quit()

def ajouter_lignes_tableau_xlwings(donnees, wb, nom_feuille, nom_tableau = ""):
    """
    Insere une ou plusieurs lignes dans un tableau structure Excel en minimisant les appels COM.
    Méthode longue → Préférer openpyxl si possible

    :param donnees: Liste contenant les donnees a inserer dans le tableau Excel
    :type donnees: List
    :param wb: Classeur dans lequel on souhaite inscrire nos donnees
    :type wb: xlwing book
    :param nom_feuille: nom de la feuille contenant le tableau
    :type nom_feuille: str
    :param nom_tableau: nom du tableau structuré (ListObject) (defaut = "")
    :type nom_tableau: str

    :Example:
    >>> ajouter_lignes_tableau_xlwings(donnees, wb_output, nom_ws)


    .. seealso:: Rien du tout.
    .. warning:: Rien du tout.
    .. note:: Rien du tout.
    .. todo:: C'est tres long (~20 minutes pour mon test avec toutes les sessions), il faudra que je teste les vitesses entre xlwings et openpyxl.
    """




    #On gère le cas par défaut où on ne donne pas de nom_tableau car c'est le même que le nom de la feuille
    if nom_tableau == "": nom_tableau = nom_feuille


    ws = wb.sheets[nom_feuille]
    table = ws.api.ListObjects(nom_tableau)
    
    data_body_range = table.DataBodyRange

    nb_lignes_nouvelles = len(donnees)
    nb_colonnes = len(donnees[0])

    # Détermine l'endroit où écrire : soit première ligne du tableau, soit après la dernière ligne existante
    if data_body_range is None or data_body_range.Value is None:
        start_cell = ws.range((table.HeaderRowRange.Row + 1, table.HeaderRowRange.Column))
    else :
        nb_lignes_existantes = data_body_range.Rows.Count
        next_row = data_body_range.Row + nb_lignes_existantes
        start_cell = ws.range((next_row, data_body_range.Column))

    # Écriture en bloc pour performance    
    ws.range(start_cell.address).resize(nb_lignes_nouvelles, nb_colonnes).value = donnees

def copierFormat_xlwings(chemin_fichier, nom_ws, range_modele, range_cible):
    """
    Copie rapidement le format d'une plage source vers une plage cible dans un fichier Excel 
    en utilisant `xlwings` et l'API COM (équivalent au collage spécial > formats dans Excel).

    :param chemin_fichier: Chemin complet du fichier Excel à modifier.
    :type chemin_fichier: str
    :param nom_ws: Nom de la feuille contenant les plages.
    :type nom_ws: str
    :param range_modele: Adresse de la plage source contenant les formats à copier (ex: "A2:G2").
    :type range_modele: str
    :param range_cible: Adresse de la plage cible à laquelle appliquer les formats (ex: "A3:G100").
    :type range_cible: str

    :return: Aucun. Le fichier Excel est modifié et enregistré.
    :rtype: None

    :example:
    >>> copier_format_rapide(
            nom_fichier="mon_fichier.xlsx",
            nom_feuille="Données",
            range_modele="A2:G2",
            range_cible="A3:G100"
        )

    .. note::
        Cette fonction nécessite Microsoft Excel installé sur votre machine (Windows uniquement).

    .. warning::
        Seuls les formats (style, bordures, police, etc.) sont copiés. Les valeurs ne sont pas modifiées.
    """
    
    app = xw.App(visible=False)
    wb = app.books.open(chemin_fichier)
    ws = wb.sheets[nom_ws]

    # Copie de la première ligne (format uniquement)
    ws.range(range_modele).copy()

    # Collage spécial des formats uniquement
    tqdm.write("⏳ Application des formats avec Excel (xlwings)...")
    ws.range(range_cible).api.PasteSpecial(Paste=-4122)  # -4122 = xlPasteFormats
    tqdm.write("✅ Formats collés avec succès.")

    wb.save()
    wb.close()
    app.quit()

def tuple_vers_liste_de_listes(t):
    """
    Permet de transformer un tuple (2, 4) en liste de listes [[2], [4]]
    C'est la forme qu'il nous faut pour ajouter efficacement des lignes dans un tableau avec xlwings

    donnees = [
        ['Jean', 'Durand', 'jean@example.com', 42],
        ['Claire', 'Martin', 'claire@example.com', 35]
    ]
    donnees = [[2], [4]]


    :param t: le tuple
    :type donnees: Tuple
    :type nom_tableau: str
    :return: Liste de listes
    :rtype: List

    :Example:
    >>> tuple_vers_liste_de_listes((2, 4))


    .. seealso:: Rien du tout.
    .. warning:: Rien du tout.
    .. note:: Rien du tout.
    .. todo:: Rien du tout.
    """
    return [[elem] for elem in t]

def open_file():
   filepath = filedialog.askopenfilename(title="Ouvrir un fichier texte", filetypes=(("fichiers texte","*.txt"), ("tous les fichiers","*.*")))
   file = open(filepath,'r')
   print(file.read())
   file.close()

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
        EI = ExportIRIS.avecEcriture(codeExport=codeIRIS, nomsFichiers_exportIRIS=dictChemins[clef], supprimeDonneesEtRemplace=True, nom_ws_imports = "Imports")

def mettreAJourTousLesExportsIRIS_fileDialog(dictCodesIRIS):
    """
    Permet de créer un seul fichier Excel à partir de plusieurs exports d'IRIS.
    Les fichiers à traités sont sélectionnés à la suite par l'utilisateur à travers un filedialog dans l'ordre du doctionnaire d'entrée, puis tous les fichiers sont traités successivement après.
    
    Les fichiers output sont des modèles avec les mêmes colonnes que les extracts d'IRIS mais avec de meilleures formes (format, couleurs...) + des colonnes adjointes à la fin pour extraire et séparer les infos du n° de session ou de la référence de la formation (ex. : trigramme formation, trigramme RP, trigramme AF...)

    Les tuples des fichiers Excel à traiter sont enregistrés dans des instances de ExtractIRIS et on emploie les méthodes de cette classe

    :param dictCodesIRIS: dictionnaire de la forme {"Sessions" : "R04110", ...}
    :type donnees: Dictionnaire

    :Example:
    >>> mettreAJourTousLesExportsIRIS_fileDialog({"Sessions" : "R04110"})


    .. seealso:: Rien du tout.
    .. warning:: Rien du tout.
    .. note:: Rien du tout.
    .. todo:: Rien du tout.
    """
    # On fait choisir à l'utilisateur
    dictEI = {}
    for clef, codeIRIS in dictCodesIRIS.items():
        EI = ExportIRIS(codeIRIS)
        EI.choisirFichiers_filedialog()
        dictEI[clef] = EI

    # On traite à la suite
    for clef in dictEI.keys():
        print(Style.BRIGHT + Fore.YELLOW + "\nTraitement des exports " + clef)
        fichiers_formates = ["\n\t" + f for f in dictEI[clef].nomsFichiers_exportIRIS]
        print("Liste des fichiers traités : " + ", ".join(fichiers_formates))
        #print("Liste des fichiers traités : " + ", ".join(dictEI[clef].nomsFichiers_exportIRIS))
        dictEI[clef].lire_extractIRIS()
        dictEI[clef].ecrit_DataFrame_dans_tableauStructure(supprimeDonneesEtRemplace=True, nom_ws_imports = "Imports", repertoire_output=rep_extractIRIS_VTE)
    
    """
    # Extract IRIS sessions (R04110)
    EI_sessions = ExportIRIS.avecEcriture("R04110", IHMSelectionFichiers=True, supprimeDonneesEtRemplace=True, nom_ws_imports = "Imports")
    #print(EI_sessions.nomsFichiers_exportIRIS)

    # Extract IRIS formations (R0304)
    EI_formations = ExportIRIS.avecEcriture("R0304", r'R0304_Ref_Formation-Listedesformations-2025.06.06.xlsx', supprimeDonneesEtRemplace=True, nom_ws_imports = "Imports")
    #print(EI_formations.nomsFichiers_exportIRIS)

    # Extract IRIS sessions Ventes (R04301)
    EI_ventes = ExportIRIS.avecEcriture("R04301", IHMSelectionFichiers=True, supprimeDonneesEtRemplace=True, nom_ws_imports = "Imports", repertoire_output=rep_extractIRIS_VTE)
    #print(EI_ventes.nomsFichiers_exportIRIS)


    # Extract IRIS sessions inscriptions (R04500) #C'est le plus long, on finit par lui
    EI_inscriptions = ExportIRIS.avecEcriture("R04500", IHMSelectionFichiers=True, supprimeDonneesEtRemplace=True, nom_ws_imports = "Imports", repertoire_output=rep_extractIRIS_VTE)
    print(EI_inscriptions.nomsFichiers_exportIRIS)
    #('R04500_Sessions-Inscriptions-FC2020 FINAL.xlsx', 'R04500_Sessions-Inscriptions-FC2021 FINAL.xlsx', 'R04500_Sessions-Inscriptions-FC2022 FINAL.xlsx', 'R04500_Sessions-Inscriptions-FC2023 FINAL.xlsx', 'R04500_Sessions-Inscriptions-FC2024 FINAL.xlsx', 'R04500_Sessions-Inscriptions-filtre sur FC-2025 au 2025-02-05.xlsx', 'R04500_Sessions-Inscriptions-filtre sur FC-2025 au 2025-03-03.xlsx', 'R04500_Sessions-Inscriptions-filtre sur FC-2025 au 2025-04-01.xlsx', 'R04500_Sessions-Inscriptions-filtre sur FC-2025 au 2025-05-12.xlsx', 'R04500_Sessions-Inscriptions-filtre sur FC-2025 au 2025-06-02.xlsx', 'R04500_Sessions-Inscriptions-filtre sur FC-2025 au 2025-07-01.xlsx', 'R04500_Sessions-Inscriptions-filtre sur FC-2025 au 2025-08-01.xlsx') 
    """





### --------------------------------------------------------------------
#  Initialisation variables globales de l'utilisateur
### --------------------------------------------------------------------


# Initialisation des chemins des répertoires
rep_gedMiroir = r"\\instnt\partage\FORMATIONS_C"

rep_extractIRIS_VTE = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits"
#rep_extractIRIS_INSTNT = r"\\instnt\HOME\REFERENC\IRIS - rapports de synthese"
rep_extractIRIS_INSTNT = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\GED"

rep_fdc_defaut = r"\P05-P06-dossier-conception-referentiel\fiche-de-cout-et-code-de-formation"
rep_specsPedagogiques_defaut = r"\P05-P06-dossier-conception-referentiel\specifications-pedagogiques-et-referentiel"

# Pour chrono des fonctions
timer = Timer()

# Pour couleur barres de progression
init(autoreset=True)
cyan = "\033[96m"
resetCouleur = "\033[0m"



# Modèle du bilan à remplir
s_word_bilan_input = r'C:\Users\vt238770\Documents\_CEA\EXCEL - Tests ou Backup\Python-TraitementWord\Bilan formation.docx'

# Bilan en sortie après remplissage
s_word_bilan_output = r'C:\Users\vt238770\Documents\_CEA\EXCEL - Tests ou Backup\Python-TraitementWord\Bilan formation - output.docx'

# Fiche de coûts
s_fdc = r'\\instnt\partage\FORMATIONS_C\948\P05-P06-dossier-conception-referentiel\fiche-de-cout-et-code-de-formation\Fiche de coûts - 948 - Elaboration de scénarios de DEM - 2025.01.24.xlsx'

# Specs pédagogiques
s_specsPedagogiques = r'\\instnt\partage\FORMATIONS_C\948\P05-P06-dossier-conception-referentiel\specifications-pedagogiques-et-referentiel\P06-Pr01-F01_Specifications-pedagogiques - 948 - 2025.04.pdf'


### --------------------------------------------------------------------
#  Début code
### --------------------------------------------------------------------
dictCodesIRIS = {
        "Sessions" : "R04110",
        "Formations" : "R0304",
        "Ventes" : "R04301",
        "Inscriptions" : "R04500",
    }
mettreAJourTousLesExportsIRIS_auto(dictCodesIRIS)

EI_sessions = ExportIRIS.avecLecture(codeExport = dictCodesIRIS["Sessions"], cheminRepertoire_exportIRIS = rep_extractIRIS_VTE, nomsFichiers_exportIRIS ="R04110_Sessions-Extract COMPLET-2025.08.19.xlsx")
EI_formations = ExportIRIS.avecLecture(codeExport = dictCodesIRIS["Formations"], cheminRepertoire_exportIRIS = rep_extractIRIS_VTE, nomsFichiers_exportIRIS ="R0304_Formations-Extract COMPLET-2025.08.19.xlsx")
EI_ventes = ExportIRIS.avecLecture(codeExport = dictCodesIRIS["Ventes"], cheminRepertoire_exportIRIS = rep_extractIRIS_VTE, nomsFichiers_exportIRIS ="R04301_Sessions-Ventes-Extract COMPLET-2025.08.19.xlsx")
EI_inscriptions = ExportIRIS.avecLecture(codeExport = dictCodesIRIS["Inscriptions"], cheminRepertoire_exportIRIS = rep_extractIRIS_VTE, nomsFichiers_exportIRIS ="R04500_Sessions-Inscriptions-COMPLET-2025.08.19.xlsx")

#PARTIE POUR TRAITER BILAN FORMATION

# TODO : faire choisir à l'utilisateur
trigrammeFormation = "948"
anneeBilan = 2023

rep_formation = rep_gedMiroir + "\\" + trigrammeFormation

#On optimise les chemins des répertoires avec ce qu'il y a dans la GED miroir
rep_fdc = optimiseCheminRepertoire(rep_formation + rep_fdc_defaut)
filepath = filedialog.askopenfilename(title="Ouvrir la fiche de coûts", filetype=[("fichiers excel","*.xlsx")], initialdir=rep_fdc, multiple=False)

rep_specsPedagogiques = optimiseCheminRepertoire(rep_formation + rep_specsPedagogiques_defaut)
filepath = filedialog.askopenfilename(title="Ouvrir les spécifications pédagogiques", filetype=[("fichiers pdf","*.pdf"), ("fichiers word","*.docx .doc")], initialdir=rep_specsPedagogiques, multiple=False)

b1 = BilanFormation(trigrammeFormation, anneeBilan, EI_sessions.che, EI_formations.cheminFichier_exportIRIS, s_fdc, s_specsPedagogiques)
b1.mergeBilan(s_word_bilan_input, s_word_bilan_output)
