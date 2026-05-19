from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass
from datetime import date
from functools import cached_property
from pathlib import Path
from tkinter.ttk import Style
from typing import Iterable, List, Optional, Tuple, Any, Type

#import pandas as pd
from pandas import *
import tqdm
from colorama import Fore

from vte.core import config, config_extractsIRIS
from vte.utils.office import FichierExcel
from vte.utils.utils import *
from vte.utils.utils import periode as utils_periode
from vte.utils.utils_instn import demander_code

# ======================================================================================
# CLASSE IRIS
# Objet fichier IRIS + logique directement liée au fichier
# ======================================================================================

# ======================================================================================
# STRUCTURES DE DONNÉES
# ======================================================================================

@dataclass
class InfosExportsIRIS:
    repertoire: Optional[Path]
    chemin_fichier: Optional[Path]
    nom_onglet: Optional[str]
    nbLignes_avantET: Optional[int]
    ordre_colonne: Optional[List[str]]

class ConfigExportIRIS:
    """
    Configuration d'un type d'export IRIS.
    Contient les propriétés input, modèle, output.
    """
    def __init__(self,
                 nom_typeExport: str,
                 codeExport: str,

                 repertoire_input: Path,
                 nom_onglet_input: str = "Data",
                 nbLignes_avantET_input: int = 0,

                 repertoire_modele: Optional[Path] = config.REPERTOIRES_MODELES, 
                 nom_fichier_modele: Optional[str | Path] = None,
                 ordre_colonnes_modele: Optional[List[str]] = None,

                 repertoire_output: Optional[Path] = config.REPERTOIRE_EXCEL_IRIS_OUTPUT,
                 nom_fichier_output: Optional[str | Path] = None):

        # Type d'export
        self._nom_typeExport = nom_typeExport
        self._codeExport = codeExport

        if not nom_fichier_modele:
            nom_fichier_modele = Path(f"{codeExport}_{nom_typeExport}-Modèle.xlsx")
        if not nom_fichier_output:
            nom_fichier_output = Path(
                f"{codeExport}_{nom_typeExport}-COMPLET-{date.today():%Y.%m.%d}.xlsx"
            )

        # Informations input données (i.e. extracts natifs d'IRIS)
        self._input:InfosExportsIRIS = InfosExportsIRIS(
            repertoire=repertoire_input, # TODO est-ce qu'avec Path je suis obligé d'avoir les 2 ?
            chemin_fichier=None,
            nom_onglet=nom_onglet_input,
            nbLignes_avantET=nbLignes_avantET_input,
            ordre_colonne = None
        )

        # Informations sur le modèle Excel à employer pour remplir l'output
        self._modele:InfosExportsIRIS = InfosExportsIRIS(
            repertoire=repertoire_modele,
            chemin_fichier=(repertoire_modele / nom_fichier_modele) if repertoire_modele else None,
            nom_onglet=nom_typeExport,
            nbLignes_avantET=None,
            ordre_colonne = ordre_colonnes_modele
        )

        # Informations output
        self._output:InfosExportsIRIS = InfosExportsIRIS(
            repertoire=repertoire_output,
            chemin_fichier=(repertoire_output / nom_fichier_output) if repertoire_output else None,
            nom_onglet=nom_typeExport,
            nbLignes_avantET=None,
            ordre_colonne = None
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

# TODO : _typeExport est fragile : mettre enum
"""
class TypeExport(Enum):
    SESSIONS = "Sessions"
    VENTES = "Ventes"
"""
# TODO : ensuite, remplacer avec get_iris(TypeExport.SESSIONS)
# ======================================================================================
# CLASSE IRIS (objet fichier)
# ======================================================================================

class IRIS:
    """
    Classe représentant un fichier IRIS (natif ou traité).
    """

    # Variables de classe
    _SESSIONS = ConfigExportIRIS(**config.IRIS_SESSIONS_PARAMS)
    _FORMATIONS = ConfigExportIRIS(**config.IRIS_FORMATIONS_PARAMS)
    _VENTES = ConfigExportIRIS(**config.IRIS_VENTES_PARAMS)
    _INSCRIPTIONS = ConfigExportIRIS(**config.IRIS_INSCRIPTIONS_PARAMS)

    """
    Dictionnaire (clefs = [ ; ]) :
    - "ConfigExportIRIS" : propriétés des Exports IRIS (chemins, modèles, nb lignes avant tableau...). Ces données sont renseignées pour :
            - les inputs (natifs IRIS) ;
            - les modèles à employer (qui sont peuplés par les inputs) ;
            - les output (fichiers traités / concaténés).
    - "chemins_fichiersInput" : liste des chemins des fichiers natifs IRIS à concaténer pour obtenir les fichiers output
    - "colonnes_modele" : je ne suis plus sûr : soit liste soit ordre soit nouveau nom pour le modèle versus l'input
    """
    DICT_EXPORTS_IRIS = {

        "ConfigExportIRIS": {
            "Sessions": _SESSIONS,
            "Formations": _FORMATIONS,
            "Ventes": _VENTES,
            "Inscriptions": _INSCRIPTIONS,
        },

        "chemins_fichiersInput": {
            "Sessions": config_extractsIRIS._tSessions,
            "Formations": config_extractsIRIS._tFormations,
            "Ventes": config_extractsIRIS._tVentes,
            "Inscriptions": config_extractsIRIS._tInscriptions,
        },

        "colonnes_modele": {
            "Sessions": None,
            "Formations": None,
            "Ventes": None,
            "Inscriptions": config_extractsIRIS._colonnes_modele_inscriptions,
        },
    }






    # ====================
    # === Constructeur ===
    # ====================
    def __init__(self, typeExport: str):
        """
        Initialise un fichier IRIS traité delin d'un type d'export
        
        :param typeExport: spécifie le type d'export. Doit être dans cette liste : ["Sessions", "Formations", "Ventes", "Insciptions"]
        :type typeExport: str
        """
        self._typeExport = typeExport  # Nom du type d'export : ["Sessions", "Formations", "Ventes", "Insciptions"]
        self._fe: Optional[FichierExcel] = None  # Fichier Excel de l'export IRIS




        # TODO quand j'ai des match ; passer par mapping:
        """
        CONVERSIONS = {
            "Sessions": self._convert_sessions,
            "Ventes": self._convert_ventes
        }
        """

    # ================================================
    # === Méthodes statiques de traitement d'infos ===
    # ================================================   
    @staticmethod
    def extraire_infos_numSessionIRIS(numSession: str) -> pd.Series:
        """
        Fonction pour extraire les colonnes à partir de la colonne 'Référence'. Je dois faire une fonction interne car j'emploie Split qui ne s'applique que sur des string. Je dois donc faire appel à cette fonction ligne par ligne et donc créer une fonction que j'appelle par DataFrame[colonne].apply().

        Exemples de cas à traiter  :
        S-04178-FC15-604-SES-CCO

        S-04934-FI1516-1512-GI_VBE_GBO
        S-05251-FI1516-1510-AMS-LCH-CLE
        S-05246-F1516-1510-OPE-HGR-NNO
        S-01131-FI12-1111-NPC-CSI-CSI
        
        S-00754-FC12-TDA BCDE-JVI-MLR

        S-05672FC17-894-MCG-CCO
        S08125-FC18-ACI-OCR-MBO

        S-00674-FC12-ACT-2-1-JV-MLR
        S-03291-FA1415-1410-STN-SCO-MNC

        S-01273-FC12-PBO-LGE  : RP et AF peuvent être extrait
        S-00653-FC12-T30-2-1-JV-MLR  : trigramme = T30
        S-00674-FC12-ACT-2-1-JV-MLR : trigramme = ACT
        S-00826-FC12047-ALA-CBR : trigramme = 047
        S-00915-FC12-SCA-1-1-JV-MLR : trigramme = SCA
        S-01444-FC13-J32 : trigramme = J32
        S-01731-FC13-ACI.SPE.SLC : on a un point en séparateur à la fin entre trigramme, rp et af
        S-02614-FC14.ACI.SPE.SLC  : on a des points en séparateurs à la fin
        S-02922-P57-HBR-MME : trigramme = P57
        S-06172-FC16.470-OCR-SDA : point en séparateur, trigramme = 470

        ANCIENNE METHODE DE TRAITEMENT
        self.__df_tableau['Numéro IRIS'] = self.__df_tableau['N° Session'].astype(str).str[2:7]
        self.__df_tableau['Type formation'] = self.__df_tableau['N° Session'].astype(str).str[8:10]
        self.__df_tableau['Trigramme AF'] = self.__df_tableau['N° Session'].astype(str).str[-3:] #tout sauf 3 derniers caract
        self.__df_tableau['Trigramme RP'] = self.__df_tableau['N° Session'].astype(str).str[-7:-4] #De -7 à -4
        self.__df_tableau['Trigramme formation'] = self.__df_tableau['N° Session'].astype(str).str[-11:-8]        
        """
        pattern_debut = re.compile(
            r"""
            ^S-?                               # S initial optionnel
            (?P<code_iris>\d{4,6})             # Code IRIS 4 à 6 chiffres
            -?                                 # Tiret optionnel
            (?P<type>[A-Z]{1,2}F?){0,1}        # Type formation optionnel
            (?P<annee>\d{2,4})?                # Année optionnelle
            [-._]?                              # Tiret/point/underscore optionnel
            (?P<reste>.*)                       # Tout le reste
            """,
            re.VERBOSE | re.IGNORECASE
        )


        colonnes = [
            'Trigramme formation',
            'Code IRIS',
            'Type de formation',
            'Année',
            'Trigramme RP',
            'Trigramme AF',
            '3ème élément de la référence'
        ]

        # Si ce n'est pas une chaîne, retourner None pour tout
        if not isinstance(numSession, str):
            return pd.Series([None]*7, index=colonnes)

        numSession = numSession.strip().replace(" ", "")
        match = pattern_debut.match(numSession)
        if not match:
            return pd.Series([None]*7, index=colonnes)

        code_iris = match.group("code_iris")
        type_formation = match.group("type")
        annee = match.group("annee")
        reste = match.group("reste")

        code_iris = int(code_iris) if code_iris else None
        annee = int(annee) if annee else None

        # Normaliser tous les séparateurs en "-"
        reste = re.sub(r"[-._]", "-", reste)
        blocs = [b for b in reste.split("-") if b]

        # Initialisation
        trigramme_formation = None
        trigramme_rp = None
        trigramme_af = None
        bloc_central = None

        nb_blocs = len(blocs)

        if nb_blocs >= 3:
            # Bloc avant les 2 derniers = trigramme formation
            trigramme_formation = blocs[-3]
            trigramme_rp = blocs[-2]
            trigramme_af = blocs[-1]
            # Tout ce qui précède trigramme formation = 3ème élément
            bloc_central = "-".join(blocs[:-3]) if nb_blocs > 3 else None

        elif nb_blocs == 2:
            # Pas de trigramme formation, juste RP et AF
            trigramme_formation = None
            trigramme_rp = blocs[0]
            trigramme_af = blocs[1]
            bloc_central = None

        elif nb_blocs == 1:
            # Seulement RP, pas de trigramme formation ni AF
            trigramme_formation = None
            trigramme_rp = blocs[0]
            trigramme_af = None
            bloc_central = None

        return pd.Series({
            'Trigramme formation': trigramme_formation or None,
            'Code IRIS': code_iris,
            'Type de formation': type_formation or None,
            'Année': annee,
            'Trigramme RP': trigramme_rp or None,
            'Trigramme AF': trigramme_af or None,
            '3ème élément de la référence': bloc_central or None
        })


    @staticmethod
    def extraire_infos_numSessionIRIS_BAK(numSession:str) -> pd.Series:
        """
        Fonction pour extraire les colonnes à partir de la colonne 'Référence'. Je dois faire une fonction interne car j'emploie Split qui ne s'applique que sur des string. Je dois donc faire appel à cette fonction ligne par ligne et donc créer une fonction que j'appelle par DataFrame[colonne].apply().

        Exemples de cas à traiter  :
        S-04178-FC15-604-SES-CCO

        S-04934-FI1516-1512-GI_VBE_GBO
        S-05251-FI1516-1510-AMS-LCH-CLE
        S-05246-F1516-1510-OPE-HGR-NNO
        S-01131-FI12-1111-NPC-CSI-CSI
        
        S-00754-FC12-TDA BCDE-JVI-MLR

        S-05672FC17-894-MCG-CCO
        S08125-FC18-ACI-OCR-MBO

        S-00674-FC12-ACT-2-1-JV-MLR
        S-03291-FA1415-1410-STN-SCO-MNC

        S-01273-FC12-PBO-LGE  : RP et AF peuvent être extrait
        S-00653-FC12-T30-2-1-JV-MLR  : trigramme = T30
        S-00674-FC12-ACT-2-1-JV-MLR : trigramme = ACT
        S-00826-FC12047-ALA-CBR : trigramme = 047
        S-00915-FC12-SCA-1-1-JV-MLR : trigramme = SCA
        S-01444-FC13-J32 : trigramme = J32
        S-01731-FC13-ACI.SPE.SLC : on a un point en séparateur à la fin entre trigramme, rp et af
        S-02614-FC14.ACI.SPE.SLC  : on a des points en séparateurs à la fin
        S-02922-P57-HBR-MME : trigramme = P57
        S-06172-FC16.470-OCR-SDA : point en séparateur, trigramme = 470

        ANCIENNE METHODE DE TRAITEMENT
        self.__df_tableau['Numéro IRIS'] = self.__df_tableau['N° Session'].astype(str).str[2:7]
        self.__df_tableau['Type formation'] = self.__df_tableau['N° Session'].astype(str).str[8:10]
        self.__df_tableau['Trigramme AF'] = self.__df_tableau['N° Session'].astype(str).str[-3:] #tout sauf 3 derniers caract
        self.__df_tableau['Trigramme RP'] = self.__df_tableau['N° Session'].astype(str).str[-7:-4] #De -7 à -4
        self.__df_tableau['Trigramme formation'] = self.__df_tableau['N° Session'].astype(str).str[-11:-8]        
        """
        # Vérifier que la référence est une chaîne de caractères
        if not isinstance(numSession, str):
            #print("Problème : la référence n'est pas une instance : ")
            #print(numSession)
            return pd.Series({
                'Trigramme formation': None,
                'Code IRIS': None,
                'Type de formation': None,
                'Année': None,
                'Trigramme RP': None,
                'Trigramme AF': None,
                '3ème élément de la référence': None
            })

        blocs = numSession.split('-')

        # Sécurité : vérifier qu'on n'a pas plus de 8 blocs
        if len(blocs) > 8:
            print("Problème : il y a plus de 8 blocs : ")
            print(numSession)
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
        #print(numSession)
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

    @staticmethod
    def extraire_infos_referenceFormationIRIS(referenceFormation:str) -> pd.Series:
        """
        Fonction pour extraire les colonnes à partir de la colonne 'Référence'. Je dois faire une fonction interne car j'emploie Split qui ne s'applique que sur des string. Je dois donc faire appel à cette fonction ligne par ligne et donc créer une fonction que j'appelle par DataFrame[colonne].apply().
        """
        
        # Vérifier que la référence est une chaîne de caractères
        if not isinstance(referenceFormation, str):
            #print("Problème : la référence n'est pas une instance : ")
            #print(referenceFormation)
            return pd.Series({
                'Trigramme formation': None,
                'Type de formation': None,
                'Année': None,
                'Unité de formation': None,
                '3ème élément de la référence': None
            })

        blocs = referenceFormation.split('-')

        # Sécurité : vérifier qu'on a au moins 4 blocs
        if len(blocs) < 3:
            print("Problème : il y a moins de 3 blocs : ")
            print(referenceFormation)
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

    @staticmethod
    def verifier_code_IRIS(valeur: Any, type_sortie: Type = str) -> Tuple[bool, Any]:
        """
        Vérifie si une valeur correspond à un entier à 5 chiffres (code IRIS).

        Args:
            valeur (Any):
                La valeur à tester. Peut être de n'importe quel type (int, float, str, etc.).
            type_sortie (Type, optionnel):
                Le type dans lequel renvoyer la valeur si elle est valide.
                Par défaut : str.
                Autres valeurs possibles : int, float, etc.

        Returns:
            Tuple[bool, Any]:
                - Le premier élément est un booléen indiquant si la valeur est un entier à 5 chiffres.
                - Le second élément est la valeur convertie dans le type demandé (ou None si invalide).

        Exemple:
            >>> verifier_code_IRIS(12345)
            (True, '12345')

            >>> verifier_code_IRIS("01234")
            (True, '01234')

            >>> verifier_code_IRIS("9999")
            (False, None)

            >>> verifier_code_IRIS("12345.0")
            (True, '12345')

            >>> verifier_code_IRIS("abcde")
            (False, None)

            >>> verifier_code_IRIS("67890", int)
            (True, 67890)

        Remarques:
            - Les zéros initiaux sont conservés si le type de sortie est `str`.
            - Les valeurs numériques flottantes représentant un entier à 5 chiffres (ex: "12345.0") sont acceptées.
            - Si la valeur ne correspond pas à 5 chiffres, la fonction renvoie (False, None).
        """
        # Conversion en chaîne pour analyse initiale
        if isinstance(valeur, str):
            str_val = valeur.strip()
        else:
            try:
                # On convertit float -> int -> str pour éviter les ".0"
                str_val = str(int(float(valeur)))
            except (ValueError, TypeError):
                return False, None

        # Vérifie qu'on a bien 5 chiffres
        if str_val.isdigit() and len(str_val) == 5:
            try:
                valeur_convertie = type_sortie(str_val)
            except Exception:
                return False, None
            return True, valeur_convertie

        return False, None

    @staticmethod
    def extraire_code_IRIS_depuis_chemin(chemin:Path|str) -> int:
        """
        Récupère le numéro IRIS (5 chiffres) depuis un chemin (a priori chemin CSV) si pas possible on demande le code à l'utilisateur
        """
        match = re.search(r"\b\d{5}\b", str(chemin))
        if match:
            codeIRIS = int(match.group(0))
        else:
            codeIRIS = demander_code(typeCode="Code IRIS", chemin=chemin)

        return codeIRIS

    # ==========================
    # === IHM ===
    # ==========================
    @staticmethod
    def demander_liste_codes_IRIS(message="Pour exclure des sessions : entrez un ou plusieurs code IRIS (numéro à 5 chiffres) séparés par des espaces ou des virgules (ou rien pour passer) : ", type_sortie:int|str=int) -> list[int|str]:
        """
        Demande en console à l'utilisateur une liste de codes IRIS séparés par des virgules

        :param message: Message à afficher à l'utilisateur pour demander cette liste. Defaut = "Pour exclure des sessions : entrez un ou plusieurs code IRIS (numéro à 5 chiffres) séparés par des espaces ou des virgules (ou rien pour passer) : ".
        :type message: str, optional
        :param type_sortie: type des éléments de la liste de sortie (int ou str), défaut = int
        :type type_sortie: int | str, optional
        :return: Liste des codes IRIS sélectionnés par l'utilisateur
        :rtype: list[int|str]
        """
        while True:
            entree = input(message).strip()
            if not entree:
                # Pas de saisie => retourner liste vide
                return []

            # On met mes codes IRIS dans une liste de str en vrifiant que toutes les valeurs sont des entiers
            try:
                # On teste le typage en int
                l_entiers = [int(v.strip()) for v in entree.replace(',', ' ').split()]

                if type_sortie == int:
                    return l_entiers
                else:
                    # On reconvertit en str avant sortie méthode
                    l_str = [str(v) for v in l_entiers]
                    return l_str

            except ValueError:
                print("Erreur : veuillez entrer uniquement des nombres entiers, séparés par des espaces ou des virgules.")
                IRIS.demander_liste_codes_IRIS()
     

    # =========================
    # ===  GETTER / SETTER  === 
    # =========================
    @property
    def cei(self) -> ConfigExportIRIS:
        """
        Config export IRIS de l'instance
        (contient toutes les propriétés de référence de ces exports : chemins input IRIS natifs, output VTE, type d'export...)
        
        :return: Description
        :rtype: ConfigExportIRIS
        """
        return IRIS.DICT_EXPORTS_IRIS["ConfigExportIRIS"][self._typeExport]

    @property
    def typeExport(self) -> str:
        """
        Type de l'export IRIS
        """
        return self._typeExport
    
    @property
    def fe(self) -> FichierExcel:
        """
        retourne le fichier Excel de l'objet
        
        :return: fichier Excel de l'objet
        :rtype: FichierExcel
        """
        return self._fe

    @fe.setter
    def fe(self, valeur:FichierExcel):
        """
        Modifie le fichier Excel de l'objet.
        
        :param valeur: nouveau fichier Excel
        :type valeur: FichierExcel
        """
        self._fe = valeur

    @property
    def df(self) -> DataFrame:
        """
        DataFrame de _fe._tableaux[self._typeExport]._df (ex. : _fe._tableaux["Sessions"]._df)
        
        :return: DataFrame de _fe._tableaux[self._typeExport]._df (ex. : _fe._tableaux["Sessions"]._df)
        :rtype: DataFrame
        """
        #return self._fe._tableaux[self._typeExport]._df
        return self._fe.get_df_tableau(self._typeExport)

    @df.setter
    def df(self, valeur:DataFrame):
        """
        Remplace le DataFrame de self._fe._tableaux[self._typeExport]._df par un nouveau DataFrame
        
        :param valeur: DataFrame devant écraser l'ancien dans self._fe._tableaux[self._typeExport]._df
        :type valeur: DataFrame
        """
        #self._fe._tableaux[self._typeExport]._df = valeur
        self._fe.set_df_tableau(self._typeExport, df=valeur)

    @property
    def tableau_donnees_iris(self) -> FichierExcel._TableauExcel:
        #return self._fe._tableaux[self.typeExport]
        return self._fe.get_tableau(self.typeExport)
    
    @property
    def tableau_fichiers_importes(self) -> FichierExcel._TableauExcel:
        #return self._fe._tableaux["Imports"]
        return self._fe.get_tableau("Imports")

    @property
    def chemin(self) -> Path:
        """
        Renvoie le chemin du fichier Excel de l'instance
        
        :return: le chemin du fichier Excel de l'instance
        :rtype: Path
        """
        return self._fe.chemin_fichier

    @chemin.setter
    def chemin(self, valeur:Path):
        """
        Change la valeur du chemin du fichier Excel
        
        :param valeur: chemin à remplacer
        :type valeur: Path
        """
        self._fe.chemin_fichier = valeur


# ======================================================================================
# CLASSE IRIS (objet fichier) pour les extracts IRIS natifs (héritage de IRIS)
# ======================================================================================
class IRIS_natif(IRIS):
    """
    Classe qui gère les extracts IRIS natifs
    """
    
    # =====================
    # === CONSTRUCTEURS ===
    # =====================
    def __init__(self, typeExport:str):
        # On initialise la classe mère
        super().__init__(typeExport=typeExport)

    @classmethod
    def avec_traitement(cls, typeExport:str, chemins_fichiersInput:Optional[str|Path|Iterable[str|Path]] = None, chemin_fichier_sauv:Optional[Path]=None) -> IRIS_natif:
        r"""
        Crée un fichier Excel unique à partir de plusieurs exports IRIS natifs.

        Si pas de chemins_input, alors on ouvre un filedialog.

        Sauvegarde soit à un endroit en argument soit à sa place par défaut (donné dans config)
        Les fichiers générés reprennent les colonnes d'origine mais avec de meilleures formes (format, couleurs...) + des colonnes adjointes à la fin pour extraire et séparer les infos du n° de session ou de la référence de la formation (ex. : trigramme formation, trigramme RP, trigramme AF...)

        :param typeExport: spécifie le type d'export. Doit être dans cette liste : ["Sessions", "Formations", "Ventes", "Insciptions"]
        :type typeExport: str
        :param chemins_fichiersInput: Chemin(s) du ou des fichiers à traiter. Si None, on ouvre un filedialog
        :type chemins_fichiersInput: str|Path|Iterable[str|Path]
        :param chemin_fichier_sauv: Chemin de savegarde si l'utilisateur ne veut pas employer celui qui est dans la config
        :type chemin_fichier_sauv: Path

        :return: un objet IRIS_natif
        :rtype: IRIS_natif

        :example:

        >>> avec_traitement(chemins_fichiersInput=(r'.\R04110_Sessions-2021 FINAL.xlsx', r'.\R04110_Sessions-2022 FINAL.xlsx'))
        >>> avec_traitement(chemins_fichiersInput=(r'.\R04110_Sessions-2021 FINAL.xlsx', r'.\R04110_Sessions-2022 FINAL.xlsx'), chemin_fichier_sauv=Path(r"C:/fichier_out.xlsx"))


        .. seealso:: Rien du tout.
        .. warning:: On ne doit traiter que des extracts IRIS du même type (sessions, formation...)
        .. note:: Remplace complètement l'ancien fichier (pas de mise à jour incrémentale)
        .. todo:: Rien du tout.
        """
        instance = cls(typeExport)
        instance._creer_export_IRIS(
            chemins_fichiersInput=chemins_fichiersInput,
            chemin_fichier_sauv=chemin_fichier_sauv
            )
        return instance
    

    # =========================
    # === Méthodes internes ===
    # =========================
    def _charger_iris_natifs(self, chemins_fichiersInput:Optional[Path|Iterable[Path]]=None) -> tuple[DataFrame, tuple[Path]]:
        r"""
        Lit le/les extract(s) IRIS et on génère dans un seul dataframe
        On concatène si besoin
        Si pas de chemins en input, alors on ouvre un filedialog.

        :param chemins_fichiersInput: Chemin(s) du ou des fichiers IRIS natifs (.xlsx) à charger dans le dataframe. Si None, on ouvre un filedialog
        :type chemins_fichiersInput: Optional[Path|Iterable[Path]]

        :return: 2 variables:
           - un DataFrame contenant les données de chemins_fichiersInput concaténées
           - un tuple des chemins concaténés
        :rtype: tuple[DataFrame, tuple[Path]]

        :example:
        >>> self._charger_iris_natifs(Path(r"C:/fichier1.xlsx"))
        >>> self._charger_iris_natifs((Path(r"C:/fichier1.xlsx"), Path(r"C:/fichier2.xlsx"))


        .. seealso:: Rien du tout.
        .. warning:: Rien du tout.
        .. note:: Rien du tout.
        .. todo:: Rien du tout.
        """

        # ===
        # === GESTION DES FICHIERS D'ENTREE ===
        # ===
        # Si aucun fichier input n'est donné, alors on ouvre un filedialog
        if chemins_fichiersInput is None:
            chemins_fichiersInput = self.choisir_fichiers(self._typeExport)

        # On convertit en Tuple[Path]
        chemins_fichiersInput = convertir_tuple_path(chemins_fichiersInput)

        # On enlève les raccourcis réseau système (dépendant de l'utilisateur) pour mettre le chemin réel (\\harmonie\instn\uem...)
        chemins_fichiersInput = tuple(chemin_vers_unc(chemin) for chemin in chemins_fichiersInput)

        # Affichage de la liste des fichiers à traiter
        print(Style.BRIGHT + Fore.YELLOW + "\nTraitement des exports " + self.typeExport)
        chemins_fichiersInput_str = [f"\n\t{f}" for f in chemins_fichiersInput]
        print("Liste des fichiers à concaténer : " + ", ".join(chemins_fichiersInput_str))


        # ===
        # === CHARGEMENT DES DATAFRAMES & CONCATENATION ===
        # ===
        # On parcourt le tuple des fichiers à lire
        df_list = [] # Liste des DataFrame qui contiendra chaque fichier Excel séparément
        taille_totale = sum(fichier.stat().st_size for fichier in chemins_fichiersInput) # Calcul taille totale pour barre de progression

        print()  # Pour avoir une ligne à écraser avec le tqdm
        with tqdm(total=taille_totale, unit='o', unit_scale=True, desc=Fore.CYAN+"Lecture des fichiers Excel" + Style.RESET_ALL) as pbar:
            #for i, ifichier in enumerate((os.path.basename(chemin) for chemin in chemins_fichiersInput), 1):
            for i, chemin in enumerate(chemins_fichiersInput, 1):
                # Données pour tqdm
                fichier = chemin.name
                taille = chemin.stat().st_size  # taille en octets 
                pbar.set_postfix(file=fichier, progress=f"{i}/{len(chemins_fichiersInput)}")  # Affichage dynamique dans la barre
                
                df = pd.read_excel(chemin, skiprows=self.cei(self.typeExport)._input.nbLignes_avantET)
                df_list.append(df)  # On ajoute le DataFrame à notre liste de DataFrame
                
                # Mise à jour de la barre avec la taille du fichier
                pbar.update(taille)

        # Concaténation finale (note : toute la fin de la méthode se fait quasi-instantanément)
        df_concat_iris = pd.concat(df_list, ignore_index=True) 


        # ===
        # === EXTRACTION INFOS DEPUIS N°IRIS ET REFERENCE FORMATION & AJOUT COLONNES ===
        # ===
        # Extraction des infos depuis "référence formation" ou "n° Iris" (dépend du type d'export)
        df_colonnes_sup = None  # Si on extrait des données depuis "référence formation" ou "n° Iris", alors on devra ajouter des colonnes au df initial
        match self.typeExport:
            # Cas Sessions ou Inscriptions ou Ventes (sensiblement comme 'Inscription R04500' mais groupé par Client (pas de détail de chaque stagiaire))
            case "Sessions" | "Inscriptions" | "Ventes":
                # On extrait / retravaille les informations de la colonne 'N° Session'
                df_colonnes_sup = df_concat_iris['N° Session'].apply(self.extraire_infos_numSessionIRIS)
                # A cause de certains éléments None ou NaN, Pandas type la colonne en float. Je la retype en Int64 qui permet de stocker des NaN avec des entiers, contrairement au type int standard.
                df_colonnes_sup['Année'] = df_colonnes_sup['Année'].astype('Int64')
                #print(df_colonnes_sup)

            # Cas Formations
            case "Formations":
                # On extrait / retravaille les informations de la colonne 'Référence'
                df_colonnes_sup = df_concat_iris['Référence'].apply(self.extraire_infos_referenceFormationIRIS)
                # A cause de certains éléments None ou NaN, Pandas type la colonne en float. Je la retype en Int64 qui permet de stocker des NaN avec des entiers, contrairement au type int standard.
                df_colonnes_sup['Année'] = df_colonnes_sup['Année'].astype('Int64')
                #print(df_colonnes_sup)

        # Ajouter les colonnes supplémentaires au DataFrame principal ssi le DataFrame df_colonnes_sup exite
        if df_colonnes_sup is not None:
            df_concat_iris = pd.concat([df_concat_iris, df_colonnes_sup], axis=1)

        # Modification structure du DataFrame (emplacement des colonnes) (dépend du type d'export)
        if IRIS.DICT_EXPORTS_IRIS["colonnes_modele"][self.typeExport] is not None:
            df_concat_iris = df_concat_iris[IRIS.DICT_EXPORTS_IRIS["colonnes_modele"][self.typeExport]]
    


        return df_concat_iris, chemins_fichiersInput

    def _ecrire_et_sauver_df_dans_excel(self, df:DataFrame, chemins_natifs_iris:tuple[Path], chemin_fichier_sauv:Optional[Path]=None) -> None:
        r"""
        Écrit et sauve un df issu d'extracts IRIS natifs dans un fichier Excel issu d'un modèle.
        Si chemin_fichier_sauv n'est pas donné, alors on emploie la valeur par défaut qui est dans cei.
        Nota : on sauvegarde également dans ce fichier la liste des chemins ayant servi à faire ce fichier (onglet "Imports")

        Example:
        >>> self._ecrire_et_sauver_df_dans_excel(df, chemins_natifs_iris)
        >>> self._ecrire_et_sauver_df_dans_excel(df, chemins_natifs_iris, Path(r"C:/fichier_out.xlsx"))

        Args:
            df (DataFrame): le DataFrame des extracts IRIS natifs concaténés
            chemins_natifs_iris (tuple[Path]): tuple des chemins ayant servi à faire df
            chemin_fichier_sauv (Optional[Path]): chemin de sauvegarde du fichier. Si chemin_fichier_sauv n'est pas donné, alors on emploie la valeur par défaut qui est dans cei.
        """
        # Si aucun fichier de sortie n'est donnée, alors on prend le chemin par défaut
        if chemin_fichier_sauv is None:
            chemin_fichier_sauv = self.cei(self.typeExport)._output.chemin_fichier
 
        # On ouvre le modèle et tous ses tableaux structurés
        self._fe = FichierExcel.depuis_modele(
                                        chemin_modele=self.cei(self.typeExport)._modele.chemin_fichier, 
                                        chemin_fichier_sauv=chemin_fichier_sauv
                                        )

        # On copie le DataFrame df avec les nouvelles données dans le modèle (on écrase les anciennes données car depuis modèle, donc tableau vide)
        self.tableau_donnees_iris.ecrit_dataFrame_dans_tableauStructure(df, supprimeDonneesEtRemplace=True, remplace_df_par_nouveau=True)

        # On écrit les références des fichiers copiés dans le tableau structuré "Imports". Nota : openpyxl ne prend pas en charge les Path donc on passe avec des str
        df_chemins = DataFrame([str(chemin) for chemin in chemins_natifs_iris], columns=['Chemin fichier'])
        self.tableau_fichiers_importes.ecrit_dataFrame_dans_tableauStructure(df=df_chemins, supprimeDonneesEtRemplace=True, remplace_df_par_nouveau=True)

        #On enregistre et on ferme (par précaution car copieformat xlwings sauvegarde)
        self._fe.save()    
        self._fe.close()

    def _creer_export_IRIS(self, chemins_fichiersInput:Optional[str|Path|Iterable[str|Path]] = None, chemin_fichier_sauv:Optional[Path]=None) -> None:
        r"""
        Crée un fichier Excel unique à partir de plusieurs exports IRIS natifs.

        Si pas de chemins_input, alors on ouvre un filedialog.

        Sauvegarde soit à un endroit en argument soit à sa place par défaut (donné dans config)
        Les fichiers générés reprennent les colonnes d'origine mais avec de meilleures formes (format, couleurs...) + des colonnes adjointes à la fin pour extraire et séparer les infos du n° de session ou de la référence de la formation (ex. : trigramme formation, trigramme RP, trigramme AF...)

        :param chemins_fichiersInput: Chemin(s) du ou des fichiers à traiter. Si None, on ouvre un filedialog
        :type chemins_fichiersInput: str|Path|Iterable[str|Path]
        :param chemin_fichier_sauv: Chemin de savegarde si l'utilisateur ne veut pas employer celui qui est dans la config
        :type chemin_fichier_sauv: Path

        :return: Ne retourne rien
        :rtype: None

        :example:

        >>> creer_export_IRIS(chemins_fichiersInput=(r'.\R04110_Sessions-2021 FINAL.xlsx', r'.\R04110_Sessions-2022 FINAL.xlsx'))
        >>> creer_export_IRIS(chemins_fichiersInput=(r'.\R04110_Sessions-2021 FINAL.xlsx', r'.\R04110_Sessions-2022 FINAL.xlsx'), chemin_fichier_sauv=Path(r"C:/fichier_out.xlsx"))


        .. seealso:: Rien du tout.
        .. warning:: On ne doit traiter que des extracts IRIS du même type (sessions, formation...)
        .. note:: Remplace complètement l'ancien fichier (pas de mise à jour incrémentale)
        .. todo:: Rien du tout.
        """
        
        # On lit le/les extract(s) IRIS et on le/les charge dans un seul dataframe
        df_concat_iris, chemins_natifs_iris = self._charger_iris_natifs(chemins_fichiersInput=chemins_fichiersInput)

        # On écrit et on sauve dans un nouvel Excel issu d'un modèle
        self._ecrire_et_sauver_df_dans_excel(
            df=df_concat_iris, 
            chemins_natifs_iris=chemins_natifs_iris, 
            chemin_fichier_sauv=chemin_fichier_sauv
            ) 



    # =========================
    # === Méthodes externes ===
    # =========================
    @staticmethod
    def cei(typeExport:str) -> ConfigExportIRIS:
        """
        Renvoie les propriétés 

        Args:
            typeExport (str): _description_

        Returns:
            ConfigExportIRIS: _description_
        """
        return IRIS.DICT_EXPORTS_IRIS["ConfigExportIRIS"][typeExport]



    # ===========
    # === IHM ===
    # ===========
    @staticmethod
    def choisir_fichiers(typeExport:str) -> Path|List[Path]:
        """
        Ouvre un filedialog pour demander à l'utilisateur de sélectionner un ou plusieurs extract IRIS natifs.
        On pointe au mieux sur le répertoire des extracts IRIS pour la boîte de dialogue.

        Args:
            typeExport (str): Type de l'export à sélectionner, doit être parmis ["Sessions", "Formations", "Ventes", "Inscriptions"]

        Returns:
            Path|List[Path]: Le ou les chemins des fichiers IRIS natifs pointés par l'utilisateur
        """
        return choisir_fichier(
            titre=f"Sélectionner un ou plusieurs fichiers Extract IRIS {str.lower(IRIS_natif.cei(typeExport)._nom_typeExport)} ({IRIS_natif.cei(typeExport)._codeExport})",
            types_fichiers=[("Fichiers Excel", "*.xlsx")],
            dossier_initial=IRIS_natif.cei(typeExport)._input.repertoire,
            texte_bouton_choisir=f"Choisir extract IRIS {str.lower(IRIS_natif.cei(typeExport)._nom_typeExport)} {IRIS_natif.cei(typeExport)._codeExport} à nouveau",
            multi_fichiers=True
        )





# ======================================================================================
# CLASSE IRIS_TRAITE (classe abstraite générique, hérite de IRIS)
# ======================================================================================
class IRIS_traite(IRIS, ABC):

    _registry = {}

    def __new__(cls, *args, **kwargs):
        """
        Surcharge de la méthode spéciale __new__ afin de transformer la classe abstraite
        IRIS_traite en "fabrique" (factory) de ses sous-classes concrètes.

        🎯 Objectif
        ----------
        Permettre une instanciation centralisée de la forme :

            >>> iris = IRIS_traite(typeExport="Sessions")

        sans que l'utilisateur ait besoin de connaître explicitement la sous-classe
        correspondante (ex. IRIS_sessions, IRIS_ventes, etc.).

        Autrement dit, cette méthode permet de retourner dynamiquement une instance
        de la bonne sous-classe en fonction de la valeur de `typeExport`.

        ⚙️ Fonctionnement
        ----------------
        - Si la classe appelée est directement IRIS_traite :
            1. On récupère l'argument `typeExport` depuis kwargs
            2. On recherche dans le registre (_registry) la sous-classe associée
            3. On instancie dynamiquement cette sous-classe
            4. On retourne cette instance

        - Si la classe appelée est déjà une sous-classe (ex. IRIS_sessions) :
            → Comportement normal (instanciation classique)

        📦 Registre (_registry)
        ----------------------
        Le dictionnaire `_registry` contient les correspondances :

            {
                "Sessions": IRIS_sessions,
                "Ventes": IRIS_ventes,
                ...
            }

        Ce registre est automatiquement rempli via la méthode __init_subclass__.

        ⚠️ Points importants
        -------------------
        - Cette méthode est appelée AVANT __init__
        - Elle contrôle le type réel de l'objet instancié
        - Elle permet de contourner l'impossibilité d'instancier une classe abstraite directement

        🚫 Cas d'erreur
        --------------
        - Si `typeExport` n'est pas fourni → ValueError
        - Si `typeExport` n'est pas reconnu → ValueError

        🔁 Exemple de flux complet
        -------------------------
            >>> iris = IRIS_traite(typeExport="Sessions")

            → appel de __new__(IRIS_traite, ...)
            → récupération de "Sessions"
            → lookup dans _registry
            → trouve IRIS_sessions
            → retourne instance de IRIS_sessions

            → ensuite __init__ de IRIS_sessions est appelé

        Returns
        -------
        instance
            Instance de la sous-classe appropriée (ex. IRIS_sessions, IRIS_ventes, etc.)
        """
        
        # Si on instancie directement IRIS_traite
        if cls is IRIS_traite:
            typeExport = kwargs.pop("typeExport", None)
            chemin = kwargs.get("chemin")

            if typeExport is None:
                raise ValueError("typeExport est requis")
        
            # On charge les modules des sous-classes (astuce pour initialiser IRIS_traite._registry)
            if not cls._registry:
                _load_subclasses()

            # Sélection de la sous-classe
            try:
                subclass = cls._registry[typeExport]
            except KeyError:
                raise ValueError(f"typeExport inconnu : {typeExport}")

            # On instancie la bonne sous-classe
            instance = super().__new__(subclass)
            return instance

        # Cas normal (appel depuis une sous-classe)
        return super().__new__(cls)

    def __init_subclass__(cls, **kwargs):
        """
        Méthode spéciale appelée automatiquement par Python à chaque définition
        d'une sous-classe de IRIS_traite.

        🎯 Objectif
        ----------
        Enregistrer automatiquement chaque sous-classe concrète dans un registre central
        (_registry), afin de permettre à __new__ de retrouver dynamiquement la bonne
        classe à instancier.

        Cela permet d'éviter :
        - un mapping manuel (typeExport → classe)
        - la duplication de logique dans plusieurs fichiers

        ⚙️ Fonctionnement
        ----------------
        Lorsqu'une sous-classe est définie :

            class IRIS_sessions(IRIS_traite):
                _typeExport = "Sessions"

        Python exécute automatiquement :

            IRIS_traite.__init_subclass__(IRIS_sessions)

        Cette méthode :
            1. Vérifie que la sous-classe possède un attribut `_typeExport`
            2. Ajoute une entrée dans le registre :

                _registry["Sessions"] = IRIS_sessions

        📦 Registre (_registry)
        ----------------------
        Dictionnaire de correspondance entre :
            - une clé métier (typeExport)
            - une classe concrète

        Exemple après chargement du module :

            {
                "Sessions": IRIS_sessions,
                "Ventes": IRIS_ventes
            }

        ⏱️ Moment d'exécution
        --------------------
        Très important :
        - Cette méthode est appelée **au moment de la définition de la classe**
        (donc à l'import du module), et NON à l'instanciation.
        - Le registre est donc prêt avant tout appel à IRIS_traite(...)

        ⚠️ Points importants
        -------------------
        - Fonctionne uniquement si les sous-classes sont importées (donc exécutées)
        - Si une sous-classe n'est jamais importée → elle ne sera pas enregistrée

        🔁 Exemple de flux complet
        -------------------------
            Chargement du module :

            class IRIS_sessions(IRIS_traite):
                _typeExport = "Sessions"

            → appel automatique de __init_subclass__
            → enregistrement dans _registry

            Plus tard :

            >>> IRIS_traite(typeExport="Sessions")

            → __new__ utilise _registry pour trouver IRIS_sessions

        Parameters
        ----------
        cls : type
            La sous-classe en cours de définition (ex. IRIS_sessions)

        Returns
        -------
        None
        """
        super().__init_subclass__(**kwargs)

        if hasattr(cls, "_typeExport"):
            IRIS_traite._registry[cls._typeExport] = cls

    def _load_subclasses():
        """
        Charge explicitement les sous-classes concrètes de IRIS_traite afin de déclencher
        leur enregistrement automatique dans le registre interne (_registry).

        🎯 Objectif
        ----------
        Permettre l'utilisation de l'API suivante :

            >>> iris = IRIS_traite(typeExport="Sessions")

        sans avoir besoin d'importer explicitement les classes concrètes
        (IRIS_sessions, IRIS_ventes, etc.) dans chaque module appelant.

        ⚙️ Fonctionnement
        ----------------
        Cette fonction réalise des imports "tardifs" (lazy imports) des modules contenant
        les sous-classes concrètes :

            from .iris_sessions import IRIS_sessions
            from .iris_ventes import IRIS_ventes

        Lors de ces imports, Python exécute la définition des classes, ce qui déclenche
        automatiquement la méthode spéciale __init_subclass__ de IRIS_traite.

        Cette méthode ajoute alors chaque sous-classe dans le registre interne :

            IRIS_traite._registry = {
                "Sessions": IRIS_sessions,
                "Ventes": IRIS_ventes,
                ...
            }

        Ainsi, après appel de _load_subclasses, la méthode __new__ de IRIS_traite
        est capable de retrouver dynamiquement la bonne classe à instancier.

        ⏱️ Quand cette fonction est-elle appelée ?
        -----------------------------------------
        Typiquement, elle est appelée dans __new__ de IRIS_traite, uniquement si le
        registre est vide :

            if not cls._registry:
                _load_subclasses()

        Cela garantit que :
        - les sous-classes sont chargées uniquement si nécessaire
        - on évite les imports circulaires au chargement du module

        🔄 Pourquoi utiliser des imports tardifs ?
        -----------------------------------------
        Pour éviter les imports circulaires :

            iris.py → iris_sessions.py → iris.py

        En déplaçant les imports à l'intérieur d'une fonction, on diffère leur exécution
        jusqu'au moment où ils sont réellement nécessaires.

        ⚠️ Points importants
        -------------------
        - Cette fonction ne retourne rien : son effet est de bord (remplissage du registre)
        - Elle doit être appelée avant toute tentative d'instanciation dynamique
        - Si une sous-classe n'est pas importée ici (ou ailleurs), elle ne sera pas disponible

        🔁 Exemple de flux complet
        -------------------------
            >>> iris = IRIS_traite(typeExport="Sessions")

            → __new__ détecte que _registry est vide
            → appel de _load_subclasses()
            → import de IRIS_sessions
            → __init_subclass__ enregistre "Sessions"
            → __new__ récupère IRIS_sessions
            → instance créée correctement

        Returns
        -------
        None
        """
        
        # Ca semblait marcher avant mais VSCode avait du mal avec l'import relatif même s'il marchait
        #from .iris_sessions import IRIS_sessions
        #from .iris_ventes import IRIS_ventes

        # Chemins complet des modules (non testé)
        from vte.domain.iris import IRIS_sessions
        from vte.domain.iris import IRIS_ventes


    def __init__(self, chemin:Optional[Path]=None, fe:Optional[FichierExcel]=None, IRIS_plus_recent:Optional[bool]=True, **kwargs):
        """
        Classe abstraite pour charger un fichier IRIS déjà traité.

        Par défaut on prend automatiquement le fichier IRIS le plus récent (IRIS_plus_recent = True).
        Si chemin est donné, alors on pointe vers celui-ci.

        Si fe est donné, alors on ne fait rien (c'est qu'il est déjà chargé)

        ⚠️ Tous les kwargs supplémentaires sont ignorés volontairement
        pour permettre l'utilisation de la factory via typeExport.

        :param typeExport: spécifie le type d'export. Doit être dans cette liste : ["Sessions", "Formations", "Ventes", "Insciptions"]
        :type typeExport: str
        :param chemin: Chemin de l'Excel IRIS à ouvrir
        :type chemin: Optional[Path]
        :param fe: Fichier Excel 
        :type fe: Optional[FichierExcel]
        :param IRIS_plus_recent: Défaut = True. Si True, on prend automatiquement le fichier IRIS le plus récent.
        :type IRIS_plus_recent: bool
        """
        # On initialise la classe mère
        #super().__init__(typeExport = self.TYPE_EXPORT)
        super().__init__(typeExport = self._typeExport)

        # On affecte les arguments aux variables d'instance
        self._fe = fe

        # Si chemin non donné et IRIS_plus_recent = True, alors on prend le fichier IRIS le plus récent.
        # Si l'utilisateur force un chemin, alors on prendra cette valeur (chemin) 
        if (chemin is None) and IRIS_plus_recent:
            chemin = self._chemin_IRIS_traite_plus_recent()

        # On ouvre le fichier IRIS s'il n'existe pas encore
        # Si chemin = None, alors _charger_excel le gèrera pour faire pointer l'utilisateur
        self._charger_excel(chemin=chemin)


    # =========================
    # === METHODES INTERNES ===
    # =========================
    def _charger_excel(self, chemin:Optional[Path]=None) -> None:
        """
            Charge dans _fe l’extract IRIS traité (excel VTE qui concatène plusieurs natifs) s'il n'existe pas déjà.
            Soit on fournit un FichierExcel, soit un chemin vers ce fichier .xlsx

            Si self._fe existe (not None), alors on ne fait rien (l'Excel est déjà chargé).
            Si chemin est vide, alors on demande à l'utilisateur de pointer un fichier.
            On crée un FichierExcel depuis le chemin et on le met dans _fe.
    
            :param chemin: Chemin de l'extract IRIS que l'on souhaite traiter
            :type chemin: Path
    
            :Example:
    
            >>> charger_excel(chemin=chemin_input)
            >>> charger_excel()

    
            .. seealso:: Rien du tout.
            .. warning:: Rien du tout.
            .. note:: Rien du tout.
            .. todo:: Rien du tout.
        """
        if self._fe is None:
            timer.debut(f"Lecture fichier IRIS {self._typeExport.lower()}")
                
            # S'il n'y a pas de chemin, alors l'utilisateur le pointe
            if chemin is None:
                chemin = self._choisir_fichier()

            # On ouvre le fichier IRIS s'il n'existe pas encore
            self._fe = FichierExcel.depuis_fichier(chemin_fichier=chemin)

            # Conversions des colonnes si besoin (l'import du df convertit en entiers ou en d'autres types certaines colonnes alors que ça ne devrait pas (ex. : trigrammes, code IRIS))
            self._convertit_types_colonnes_df()

            # On trie # TODO : pour l'instant ça marche avec tous mes types (sessions et ventes ; j'ai pas testé avec les autres)
            self.df = self.df.sort_values(by=self._nom_colonne_debut_session)  # Trie par "Date début ses."

            timer.fin()

    def _chemin_IRIS_traite_plus_recent(self) -> Path:
        """
        Récupère le chemin du fichier le plus récent dans le répertoire spécifié.
        Le nom du fichier doit correspondre au format :
        f"{self.cei._codeExport}_{self._typeExport}-COMPLET-{date.today():%Y.%m.%d}.xlsx"

        Returns:
            Path: Chemin du fichier le plus récent, ou None si aucun fichier correspondant n'est trouvé.
        """
        # Fonction pour extraire la date du nom de fichier
        def extraire_date(fichier):
            match = re.search(r"-COMPLET-(\d{4}\.\d{2}\.\d{2})$", fichier.stem)
            if not match:
                raise ValueError(
                    f"Format de date incorrect dans le nom du fichier : {fichier.name}"
                )
            return datetime.strptime(match.group(1), "%Y.%m.%d")
        
      
        # Récupérer le répertoire depuis self.cei._output.chemin_fichier
        repertoire = self.cei._output.chemin_fichier.parent
        #print(repertoire)

        # Générer le pattern de nom de fichier à rechercher
        pattern = f"{self.cei._codeExport}_{self._typeExport}-COMPLET-*.xlsx"

        # Lister tous les fichiers dans le répertoire qui correspondent au pattern
        fichiers = list(repertoire.glob(pattern))
        #print(fichiers)

        if not fichiers:
            return None

        # Trier les fichiers par date extraite du nom (du plus récent au plus ancien)
        fichiers_tries = sorted(fichiers, key=extraire_date, reverse=True)

        fichier_iris_plus_recent = fichiers_tries[0]
        #print(fichier_iris_plus_recent)

        # Retourner le chemin du fichier le plus récent
        return fichier_iris_plus_recent
    
    @abstractmethod
    def _convertit_types_colonnes_df(self) -> None:
        pass


    # ====================
    # === METHODES GET ===
    # ====================
    def get_champ_depuis_codes_IRIS(self, champ:str, codes_IRIS:int|Iterable[int], valeurUnique:bool=False) -> str|tuple[str]|int|tuple[int]|date|tuple[date]:
        """
        Méthode générique quel que soit le champ.

        Récupère la ou les valeurs d'un champ d'un IRIS traité (sessions...) à partir d'un ou plusieurs codes IRIS (on ne conserve pas les doublons).

        Si valeurUnique est True on vérifie qu'on n'a qu'un seul retour sinon on lance une erreur (et type retour = str sinon type retour = tuple[str]).

        Si le champ est de type date (datetime64 pandas), il est converti en `datetime.date`

        :param champ: Nom de la colonne de l'export IRIS dont il faut récupérer la valeur ("Trigramme formation", "Session" (intitulé foramtion), "Date début ses.", "Année début ses.", "Lieu principal"...)
        :type champ: str
        :param codes_IRIS: Codes IRIS de la formation dont on souhaite récupérer les valeurs du champ.
        :type codes_IRIS: int|Iterable[int]
        :param valeurUnique: Si valeurUnique est True on vérifie qu'on n'a qu'un seul retour sinon on lance une erreur. Défaut = False
        :type valeurUnique: bool, optional
        :return: La ou les valeurs d'un champ d'un IRIS traité (sessions...) correspondant à codes_IRIS. Si valeurUnique=True, type retour = str|int|date ; sinon un tuple.
        :rtype: str|tuple[str]|int|tuple[int]|date|tuple[date]
        """      
        # On évalue notre sortie
        df = self.df_filtre_codes_IRIS(codes_IRIS=codes_IRIS)
        
        if pd.api.types.is_datetime64_any_dtype(df[champ]):  # Si champ de type date  # Si Excel mal typé, alors on pourra faire df[champ] = pd.to_datetime(df[champ], errors="coerce")
            resultat = tuple(df[champ].dropna().dt.date.unique())
        else:  # Sinon
            resultat = tuple(df[champ].dropna().unique())

        # On gère le return en fonction du nombre de résultat et de valeurUnique
        if len(resultat) == 1:  # Si resultat unique
            if valeurUnique:
                return resultat[0]  # str|int|date
            else:
                return resultat  # tuple[str]|tuple[int]|tuple[date]
        else:  # Si plusieurs résultats
            if valeurUnique:
                return vlog.log_erreur(f"{champ} unique attendu, or resultat = {resultat}")  # Erreur car valeur unique attendue
            else:
                return resultat  # tuple[str]|tuple[int]|tuple[date]

    def get_periode_depuis_codes_IRIS(self, codes_IRIS: int | Iterable[int]) -> tuple[int, str]:
        """
        Récupère l'année et la période ("1er semestre", "2nd semestre", "Année")
        correspondant à un ou plusieurs codes IRIS, à partir des dates de début de session.

        Les dates sont récupérées via get_champ_depuis_codesIRIS, qui renvoie des objets `datetime.date`.

        Hypothèses :
        - Toutes les sessions doivent appartenir à la même année (sinon on lève une erreur).
        - La période est déterminée en fonction de la date minimale et maximale.

        :param codes_IRIS: Code(s) IRIS des sessions à analyser
        :type codes_IRIS: int | Iterable[int]

        :return: Tuple (année, période)
                - année : int
                - période : str ("1er semestre", "2nd semestre", "Année")
        :rtype: tuple[int, str]
        """

        # On récupère les dates de début (déjà en datetime.date)
        dates_debuts = self.get_champ_depuis_codes_IRIS(
            champ=self.nom_colonne_debut_session,
            codes_IRIS=codes_IRIS
        )

        date_min = min(dates_debuts)
        date_max = max(dates_debuts)

        # Vérification même année
        meme_annee = all(date_debut.year == date_min.year for date_debut in dates_debuts)

        if meme_annee:
            annee = date_min.year
        else:
            vlog.log_erreur(f"Tous les codes IRIS sont sensés être sur la même année, or date min = {date_min} et date max = {date_max}")

        # Bornes en datetime.date (pas pandas)
        debut_annee = date(annee, 1, 1)
        fin_semestre_1 = date(annee, 6, 30)
        debut_semestre_2 = date(annee, 7, 1)
        fin_annee = date(annee, 12, 31)

        # Détermination de la période
        if date_min >= debut_annee and date_max <= fin_semestre_1:
            periode = "1er semestre"
        elif date_min >= debut_semestre_2 and date_max <= fin_annee:
            periode = "2nd semestre"
        elif date_min >= debut_annee and date_max <= fin_annee:
            periode = "Année"
        else:
            vlog.log_erreur(
                f"Période incohérente : date min = {date_min}, date max = {date_max}"
            )

        return annee, periode


    # ============================
    # === METHODES LIEES AU DF ===
    # ============================
    def affiche_df_colonnes_principales(self, df:DataFrame) -> None:
        """
        print le DataFrame avec filtre des colonnes pour affichage : ['Code IRIS', 'Trigramme RP', 'Trigramme AF', 'Date début ses.', 'Date fin ses.', 'Durée réal. (J.)', 'Nb. Présents', 'Statut Session', 'N° Session']

        :param df: Dataframe dont on souhaite n'afficher que les colonnes principales
        :type df: DataFrame
        """
        print(tabulate(
            df[self._colonnes_principales], 
            headers='keys', 
            tablefmt='pretty', 
            showindex=False
        ))

    def df_filtre_codes_IRIS(self, codes_IRIS:int|Iterable[int])  -> DataFrame:
        """
        Renvoie le dataframe de l'extract IRIS filtré VTE selon Code IRIS in codes_IRIS.
        
        :param codes_IRIS: Codes IRIS sur lesquels on veut filtrer IRIS sessions
        :type codes_IRIS: int|Iterable[int]
        :return: extract IRIS VTE filtré sur codes_IRIS
        :rtype: DataFrame
        """     
        return self.df[self.df['Code IRIS'].isin(convertir_collection(codes_IRIS))]

    def df_filtre_periode(self, trigramme_formation:Optional[str] = None, annee:Optional[int] = None, periode:Optional[str] = None) -> DataFrame:
        """
        Renvoie le dataframe de l'extract IRIS filtré VTE selon plusieurs critères.
         
         Les critères dépendent du typé d'export ; pour sessions :
            - Statut Session != "Annulée" (toujours) ;
            - Nb. Nommés != 0 (toujours ; ssi IRIS sessions) ;
            - l'année de la session (si donné en argument) ;
            - la période de la session (si donné en argument) ;
            - le trigramme de la formation en cours (si donné en argument).

        Par défaut la période est toute l'année, sinon il faut préciser "1er semestre" ou "2nd semestre".

        Trié de sorte que la session la plus récente est en 1ère ligne.
        
        :param trigramme_formation: trigramme filtré. Si non renseigné : pas de filtre sur ce critère.
        :type trigramme_formation: Optional[str]
        :param annee: année du filtre. Si non renseignée, alors année en cours.
        :type annee: Optional[int]
        :param periode: période du filtre ("1er semestre", "2nd semestre"). Si non renseigné, alors période="Année".
        :type periode: Optional[str]
        :return: extract IRIS VTE filtré
        :rtype: DataFrame
        """
        #TODO à faire gaffe à la fin :
        # if annee is None : annee = self._annee

        # On ouvre et lit l'export sessions IRIS si et seulement si il n'est pas déjà ouvert et lu avant
        #self.charger_excel_IRIS()

        # Application du pré-filtre immuable : statut (session != annulé) et (Nb. nommés != 0)
        df = self.df_prefiltre
        #print(self.df)


        # Si trigramme en argument, alors on filtre sur le trigramme
        if trigramme_formation is not None:
            df = df[
                (df['Trigramme formation'] == trigramme_formation)
            ]
            #print(self.df)

        # Si données temporelles (année ou période ou les deux)
        #if periode is not None:
        #    # Si période en argument, alors on filtre sur la période (plus restrictif que l'année)
        #    date_debut, date_fin = debut_fin_periode(annee=annee, periode=periode)
        #    df = df[
        #        (df[self.nom_colonne_debut_session] >= date_debut) &
        #        (df[self.nom_colonne_debut_session] <= date_fin)
        #    ]   
        #elif annee is not None:
        #    df = df[
        #        (df[self.nom_colonne_debut_session] == annee)
        #    ]
        
        # Si période en argument, alors on filtre sur la période (plus restrictif que l'année)
        date_debut, date_fin = debut_fin_periode(annee=annee, periode=periode)
        df = df[
            (df[self.nom_colonne_debut_session] >= date_debut) &
            (df[self.nom_colonne_debut_session] <= date_fin)
        ]   
        #print(self.df)


        # On trie
        df = df.sort_values(self.nom_colonne_debut_session, ascending=False)

        return df




    # =========================
    # === GETTERS / SETTERS === 
    # =========================
    @property
    @abstractmethod
    def nom_colonne_debut_session(self) -> str:
        pass
    
    @property
    @abstractmethod
    def nom_colonne_statut_session(self) -> str:
        pass

    @cached_property
    @abstractmethod
    def df_prefiltre(self) -> DataFrame:
        """
        Fait un premier pré-filtre sur le DataFrame de l'IRIS :
            - sur IRIS sessions, on fait :
                - Statut Session != "Annulée",
                - Nb. Nommés != 0 ;
            - sur IRIS ventes, on fait :
                - Statut session != "Annulée".
        """
        pass

    #    @property
    #    @abstractmethod
    #    def TYPE_EXPORT(self) -> str:
    #        pass





    # ===========
    # === IHM ===
    # ===========

    def _choisir_fichier(self) -> Path | None:
        """
        Ouvre un filedialog pour demander à l'utilisateur de sélectionner un extract IRIS VTE (excel VTE qui concatène plusieurs natifs).
        On pointe au mieux sur le répertoire des extracts IRIS pour la boîte de dialogue.
    
        :return: Le chemin du fichier IRIS pointé par l'utilisateur
        :rtype: Path

        :example:

        >>> choisir_fichier_IRIS()


        .. seealso:: Rien du tout.
        .. warning:: Rien du tout
        .. note:: Rien du tout
        .. todo:: Rien du tout.
        """

        return choisir_fichier(
                    titre=f"Sélectionner l'extract IRIS {str.lower(self._typeExport)} {self.cei._codeExport} à employer.",
                    types_fichiers=[("Fichiers Excel", "*.xlsx")],
                    dossier_initial=self.cei._output.repertoire, # Pour aller vers mes fichiers concaténés, sinon pour les originaux il faut pointer vers input
                    texte_bouton_choisir=f"Choisir extract IRIS {str.lower(self._typeExport)} {self.cei._codeExport} à nouveau"
                    )




# ==========================================================================================================
# SOUS-CLASSE CONCRETE IRIS_sessions (objet fichier) pour les IRIS sessions traités (spécifie IRIS_traite, hérite de IRIS)
# ==========================================================================================================
class IRIS_sessions(IRIS_traite):
    """
    Classe qui permet d'ouvrir et traiter un Extract IRIS sessions déjà traité
    """

    # ===========================
    # === VARIABLES DE CLASSE ===
    # ===========================
    #TYPE_EXPORT = "Sessions"
    _typeExport:str = "Sessions"
    _nom_colonne_debut_session:str = "Date début ses."
    _nom_colonne_statut_session:str = "Statut Session"

    _colonnes_principales:list[str] = ['Code IRIS', 'Trigramme RP', 'Trigramme AF', 'Date début ses.', 'Date fin ses.', 'Durée réal. (J.)', 'Nb. Nommés', 'Statut Session', 'N° Session']

    # =========================
    # === METHODES INTERNES ===
    # =========================
    def _convertit_types_colonnes_df(self) -> None:
        """
            L'import du df convertit en entiers ou en d'autres types certaines colonnes alors que ça ne devrait pas (ex. : trigrammes, code IRIS sur l'extract sessions).
            
            Il en résulte que des filtres ne fonctionnement pas sans conversion.
            On convertit donc des colonnes à cette fin.
            
            Ex. sur l'extract Session de IRIS :
               - trigrammes -> str;
               - code IRIS -> str (? Pourquoi str et pas int).
    
            :Example:

            >>> convertit_types_colonnes_df()

    
            .. seealso:: Rien du tout.
            .. warning:: Rien du tout.
            .. note:: Rien du tout.
            .. todo:: Rien du tout.
        """
        # Les conversion dépendent du type d'extract à cause des noms des colonnes
        self.df["Trigramme formation"] = self.df["Trigramme formation"].astype(str)  # Retype "Trigramme formation"
        #self.df["Code IRIS"] = self.df["Code IRIS"].astype(str)  # Retype "Code IRIS"



    # ============================
    # === METHODES LIEES AU DF ===
    # ============================




    # ===========
    # === IHM ===
    # ===========
    def demande_sessions_a_retenir(self, trigramme_formation:str = None, annee:Optional[int] = None, periode:Optional[str] = None) -> list[int]:
        """
        Demande à l'utilisateur les sessions qu'il souhaite retenir de la période choisie :
           - on applique un filtre sur une période + trigramme au dataframe de l'extract IRIS
           - on affiche le résultat
           - l'utilisateur définit les codes IRIS qu'il veut retenir
        
        On retourne :
           - la liste des codes IRIS retenus
        
        :param trigramme_formation: trigramme de la formation à filtrer
        :type trigramme_formation: str
        :param annee: année à filtrer
        :type annee: Optional[int]
        :param periode: période à filtrer
        :type periode: Optional[str]
        :return: la liste des codes IRIS retenus
        :rtype: list[int]
        """

        # Renseigne les valeurs par défaut de période et année si None
        annee, periode = utils_periode(annee, periode)

        # On filtre sur le trigramme et la période demandée
        df_filtre = self.df_filtre_periode(trigramme_formation=trigramme_formation, annee=annee, periode=periode).copy()

        # Adaptation format date
        #df_filtre['Date début ses.'] = pd.to_datetime(df_filtre['Date début ses.']).dt.strftime("%d/%m/%Y")
        #df_filtre['Date fin ses.'] = pd.to_datetime(df_filtre['Date fin ses.']).dt.strftime("%d/%m/%Y")

        # On affiche à l'utilisateur les sessions et dates et statuts 
        vlog.print("Info", f"\nListe des sessions {trigramme_formation} dans {self.chemin.name} - {periode} {annee}", style=["jaune"])
        self.affiche_df_colonnes_principales(df_filtre)
        
        # On demande à l'utilisateur les sessions qu'il veut exclure
        sessionsRetenues = self.demander_liste_codes_IRIS(message="Pour retenir des sessions : entrez un ou plusieurs code IRIS (numéro à 5 chiffres) séparés par des espaces ou des virgules (ou rien pour passer) :")

        if sessionsRetenues:  # si la liste n'est pas vide
            # On met à jour _df_sessions_filtre en enlevant les sessions exclues
            df_filtre = df_filtre[df_filtre['Code IRIS'].isin(sessionsRetenues)]
        else:
            # la liste est vide, on ne filtre rien, on garde tout
            pass

        # On affiche à l'utilisateur les sessions finalement retenues
        vlog.print("Info", "\nSessions retenues pour le bilan :", style=["jaune"])
        self.affiche_df_colonnes_principales(df_filtre)

        return sessionsRetenues

    def demande_sessions_a_exclure(self, trigramme_formation:str = None, annee:Optional[int] = None, periode:Optional[str] = None) -> List[int]:  #Tuple[DataFrame, List[int]]:
        """
        Retourne une liste de codes IRIS retenus après exclusions de certains éléments par l'utilisateur
        
        Demande à l'utilisateur les sessions qu'il souhaite exclure de la période choisie :
           - on applique un filtre sur une période + trigramme au dataframe de l'extract IRIS
           - on affiche le résultat
           - l'utilisateur définit les codes IRIS qu'il veut exclure
        
        On retourne :
           - un dataframe df_sessions_filtre à jour
           - la liste des codes IRIS retenus (plus maintenant : pour l'avoir on peut faire df_filtre["Code IRIS"].tolist())
           - #la liste des codes IRIS exclus
        
        :param trigramme_formation: trigramme de la formation à filtrer
        :type trigramme_formation: str
        :param annee: année à filtrer
        :type annee: Optional[int]
        :param periode: période à filtrer
        :type periode: Optional[str]
        :return: la liste des codes IRIS retenus
        :rtype: Tuple[DataFrame, List[int]]
        """

        # Liste des sessions qui seront exclues par l'utilisateur
        codes_sessions_exclues_par_utilisateur = []

        # Renseigne les valeurs par défaut de période et année si None
        annee, periode = utils_periode(annee, periode)

        # On filtre sur le trigramme et la période demandée
        df_filtre = self.df_filtre_periode(trigramme_formation=trigramme_formation, annee=annee, periode=periode)

        # On affiche à l'utilisateur les sessions et dates et statuts 
        vlog.print("Info", f"\nListe des sessions {trigramme_formation} dans {self.chemin.name} - {periode} {annee}", style=["jaune"])
        self.affiche_df_colonnes_principales(df_filtre)

        
        # On demande à l'utilisateur les sessions qu'il veut exclure
        exclusionSessions = self.demander_liste_codes_IRIS()
        if exclusionSessions:  # si la liste n'est pas vide
            # On trace l'exclusion des sessions
            codes_sessions_exclues_par_utilisateur = df_filtre.loc[df_filtre["Code IRIS"].isin(exclusionSessions), "N° Session"].unique().tolist()
            #for session_exclue in exclusionSessions:
            #    codes_sessions_exclues_par_utilisateur.append(df_filtre.loc[df_filtre["Code IRIS"] == session_exclue, "N° Session"].iloc[0])

            # On met à jour _df_sessions_filtre en enlevant les sessions exclues
            df_filtre = df_filtre[~df_filtre['Code IRIS'].isin(exclusionSessions)]
        else:
            # la liste est vide, on ne filtre rien, on garde tout
            pass

        # On affiche à l'utilisateur les sessions finalement retenues
        vlog.print("Info", "\nSessions retenues pour le bilan :", style=["jaune"])
        self.affiche_df_colonnes_principales(df_filtre)


        # pour avoir la liste des codes IRIS retenus : df_filtre["Code IRIS"].tolist()
        codes_IRIS_retenus = df_filtre["Code IRIS"].tolist()

        # On retourne les codes_IRIS des sessions retenues + exploitationBilan 
        #return df_filtre, codes_sessions_exclues_par_utilisateur
        return codes_IRIS_retenus




    # =========================
    # === GETTERS / SETTERS === 
    # =========================
    @property
    def nom_colonne_debut_session(self) -> str:
        return self._nom_colonne_debut_session
    
    @property
    def nom_colonne_statut_session(self) -> str:
        return self._nom_colonne_statut_session

    @cached_property
    def df_prefiltre(self) -> DataFrame:
        """
        Fait un premier pré-filtre sur le DataFrame de l'IRIS :
            - sur IRIS sessions, on fait :
                - Statut Session != "Annulée",
                - Nb. Nommés != 0.
        """
        return self.df[
            (self.df[self.nom_colonne_statut_session] != "Annulée") &
            (self.df[self.nom_colonne_debut_session] != 0)
        ]




# ==========================================================================================================
# SOUS CLASSE CONCRETE IRIS_ventes (objet fichier) pour les IRIS sessions traités (spécifie IRIS_traite, hérite de IRIS)
# ==========================================================================================================
class IRIS_ventes(IRIS_traite):
    """
    Classe qui permet d'ouvrir et traiter un Extract IRIS sessions déjà traité
    """
    

    # ===========================
    # === VARIABLES DE CLASSE ===
    # ===========================
    #TYPE_EXPORT = "Ventes"
    _typeExport:str = "Ventes"
    _nom_colonne_debut_session:str = "Date de début"
    _nom_colonne_statut_session:str = "Statut session"

    _colonnes_principales:list[str] = ['Code IRIS', 'RP', 'Date de début', 'Lieux Principal', "Statut session", 'Type', 'N° Session', 'Intitulé Client', 'Nb Inscriptions', 'Total HT', 'Type tarif']

    # TODO à virer après les phases de test
    #chemin_IRIS_ventes:Path = Path(chemin_vers_unc(r"R:\_Echanges\VTE\Prog\IRIS\Extracts complets\R04301_Ventes-COMPLET-2026.02.12.xlsx"))
    
    # =========================
    # === METHODES INTERNES ===
    # =========================
    def _convertit_types_colonnes_df(self) -> None:
        """
            L'import du df convertit en entiers ou en d'autres types certaines colonnes alors que ça ne devrait pas (ex. : trigrammes, code IRIS sur l'extract sessions).
            
            Il en résulte que des filtres ne fonctionnement pas sans conversion.
            On convertit donc des colonnes à cette fin.
            
            Ex. sur l'extract Ventes de IRIS :
               - Date de début -> datetime.
    
            :Example:

            >>> convertit_types_colonnes_df()

    
            .. seealso:: Rien du tout.
            .. warning:: Rien du tout.
            .. note:: Rien du tout.
            .. todo:: Rien du tout.
        """
        # Les conversion dépendent du type d'extract à cause des noms des colonnes
        # On convertit la colonne "Date de début" en datetime
        self.df['Date de début'] = pd.to_datetime(self.df['Date de début'])


    # =========================
    # === METHODES EXTERNES ===
    # =========================
    # Forcer écriture unité
    def prix_formation_annee_str(self, trigramme_formation:str, annee:int) -> str:
        """
        Renvoie le ou les prix d'une formation pour une année donnée en argument.
        Ce ou ces prix sont donnés au format str pour être inclus dans un bilan (ex. formation)

        :param trigramme_formation: Trigramme de la formation
        :type trigramme_formation: str
        :param annee: Année pour laquelle on souhaite avoir le ou les prix
        :type annee: int
        :return: Prix de la formation au format str prête à être intégrée à un bilan de formation
        :rtype: str
        """

        # Création DataFrame avec les colonnes créées pour calculer le prix de la formation de l'année donnée en argument
        df_travail = self.df_prix_formation_annee(trigramme_formation=trigramme_formation, annee=annee)

        if len(df_travail["Prix HT EE"].unique()) == 1:  # Cas avec une seule valeur de prix
            prix_annee = f"{int(df_travail.iloc[0]["Prix HT EE"])} {df_travail.iloc[0]["Unité prix"]}"  # Je convertis mon prix en int
        else :  # Cas avec plusieurs valeurs de prix
            # On affiche le tableau avec les différents prix
            print(df_travail[['Date de début', 'N° Session','Intitulé Client', 'CEA', 'Nb Inscriptions', 'Total HT', 'Type tarif', 'Prix HT EE', 'Unité prix']])

            # Regrouper les lignes par "Prix HT EE" et "Unité prix"
            df_travail_groupe = df_travail.groupby(['Prix HT EE', 'Unité prix'])

            # Initialiser une liste pour stocker les lignes de texte
            lignes_texte = []

            # Parcourir chaque groupe
            for (prix, unite), groupe in df_travail_groupe:
                # Créer une liste des éléments pour chaque ligne du groupe
                elements = []
                for _, row in groupe.iterrows():
                    element = f"{row['N° Session']}; {row['Date de début']}; {row['Intitulé Client']}; {row['Nb Inscriptions']}; {row['Total HT']}; {row['Type tarif']}"
                    elements.append(element)

                # Créer la ligne de texte pour le groupe
                ligne_texte = f"• {prix} {unite} :\n\t- " + "\n\t- ".join(elements)
                lignes_texte.append(ligne_texte)

            # Joindre toutes les lignes de texte avec des sauts de ligne
            prix_annee = "\n".join(lignes_texte)

        return prix_annee



    # ====================
    # === METHODES GET ===
    # ====================


    def get_periode_depuis_codes_IRIS_AVIRER(self, codes_IRIS: int | Iterable[int]) -> tuple[int, str]:
        """
        Récupère l'année et la période ("1er semestre", "2nd semestre", "Année")
        correspondant à un ou plusieurs codes IRIS, à partir des dates de début de session.

        Les dates sont récupérées via get_champ_depuis_codesIRIS, qui renvoie des objets `datetime.date`.

        Hypothèses :
        - Toutes les sessions doivent appartenir à la même année (sinon on lève une erreur).
        - La période est déterminée en fonction de la date minimale et maximale.

        :param codes_IRIS: Code(s) IRIS des sessions à analyser
        :type codes_IRIS: int | Iterable[int]

        :return: Tuple (année, période)
                - année : int
                - période : str ("1er semestre", "2nd semestre", "Année")
        :rtype: tuple[int, str]
        """

        # On récupère les dates de début (déjà en datetime.date)
        dates_debuts = self.get_champ_depuis_codes_IRIS(
            champ="Date début ses.",
            codes_IRIS=codes_IRIS
        )

        date_min = min(dates_debuts)
        date_max = max(dates_debuts)

        # Vérification même année
        meme_annee = all(date_debut.year == date_min.year for date_debut in dates_debuts)

        if meme_annee:
            annee = date_min.year
        else:
            vlog.log_erreur(f"Tous les codes IRIS sont censés être sur la même année, or date min = {date_min} et date max = {date_max}")

        # Bornes en datetime.date (pas pandas)
        debut_annee = date(annee, 1, 1)
        fin_semestre_1 = date(annee, 6, 30)
        debut_semestre_2 = date(annee, 7, 1)
        fin_annee = date(annee, 12, 31)

        # Détermination de la période
        if date_min >= debut_annee and date_max <= fin_semestre_1:
            periode = "1er semestre"
        elif date_min >= debut_semestre_2 and date_max <= fin_annee:
            periode = "2nd semestre"
        elif date_min >= debut_annee and date_max <= fin_annee:
            periode = "Année"
        else:
            vlog.log_erreur(
                f"Période incohérente : date min = {date_min}, date max = {date_max}"
            )

        return annee, periode


    # ============================
    # === METHODES LIEES AU DF ===
    # ============================
    def df_prix_formation_annee(self, trigramme_formation:str, annee:int) -> DataFrame:
        """
        Génère le DataFrame qui permettra d'évaluer le prix de la formation d'une année.
        Crée entre autres les colonnes :
        - CEA ;
        - Prix HT EE ;
        - Unité prix.

        :param trigramme_formation: Trigramme de la formation
        :type trigramme_formation: str
        :param annee: Année pour laquelle on souhaite créer le dataframe (le prix par extention)
        :type annee: int
        :return: le dataframe avec les colonnes rajoutées qui permettront de calculer le prix de cette année
        :rtype: DataFrame
        """

        # Filtrer le DataFrame sur le "Trigramme formation" = "TEL" et sur les années n et n-1
        #df_filtre = self.df[(self.df['Trigramme formation'] == self._trigramme_formation) & (iris_ventes.df['Date de début'].dt.year.isin([self._annee-1, self._annee]))]
        df_filtre = self.df_filtre_periode(trigramme_formation=trigramme_formation, annee=annee)

        # Créer un DataFrame df_travail avec les colonnes spécifiées
        df_travail = df_filtre[self._colonnes_principales]

        # Ajouter une colonne "CEA" (booléen) pour tester si les 3 premières lettres de "Intitulé Client" = "CEA"
        df_travail['CEA'] = df_travail['Intitulé Client'].str[:3] == 'CEA'

        # Ajouter une colonne "Prix HT EE" (si tarif CEA, alors on divise par 0.9 ; si on est sur un forfait/pers., alors on divise par nombre d'inscriptions)
        df_travail['Prix HT EE'] = df_travail.apply(
            lambda row: (row['Total HT'] / 0.9 if row['CEA'] else row['Total HT']) / row['Nb Inscriptions'] if row['Type tarif'] == 'Forfait/Pers.' else (row['Total HT'] / 0.9 if row['CEA'] else row['Total HT'])
            , axis=1)

        # Arrondir le résultat à l'entier le plus proche
        df_travail['Prix HT EE'] = df_travail['Prix HT EE'].round()

        # Ajouter une colonne "Unité prix" en fonction du "Type de tarif"
        df_travail['Unité prix'] = df_travail['Type tarif'].apply(lambda x: '€ HT (forfait)' if x == 'Forfait' else '€ HT/pers.')


        # Vérification que tous les tarifs sont bien les mêmes
        # J'exclue les lignes si "Prix HT EE" = NaN
        df_travail = df_travail.dropna(subset=['Prix HT EE'])

        # J'exclue les lignes si "Prix HT EE" = 0
        df_travail = df_travail[df_travail['Prix HT EE'] != 0]

        return df_travail

    # =========================
    # === GETTERS / SETTERS === 
    # =========================
    @property
    def nom_colonne_debut_session(self) -> str:
        return self._nom_colonne_debut_session
    
    @property
    def nom_colonne_statut_session(self) -> str:
        return self._nom_colonne_statut_session

    @cached_property
    def df_prefiltre(self) -> DataFrame:
        """
        Fait un premier pré-filtre sur le DataFrame de l'IRIS :
            - sur IRIS sessions, on fait :
                - Statut Session != "Annulée".
        """
        return self.df[
            (self.df[self.nom_colonne_statut_session] != "Annulée")
        ]
















# ==========================
# === Méthodes statiques pour ouvir un extract existant IRIS Session (excel VTE qui concatène plusieurs natifs)  ===
# Anciennes méthodes avant que je ne mette _fe en variable d'instance.
# TODO : à transférer à priori s'il reste des anciens appels
# TODO à vérifier : est-ce que ces fonctions n'étaient utiles que pour IRIS_traite (avant que je ne la fasse) → Soit supprimer, soit adapter pour IRIS non traités
# ==========================
r"""
@staticmethod
def choisir_fichier_IRIS_sessions() -> Path | None:

    return choisir_fichier(titre=f"Sélectionner l'extract IRIS session {IRIS.cei("Sessions")._codeExport} à employer.",
                    types_fichiers=[("Fichiers Excel", "*.xlsx")],
                    dossier_initial=IRIS.cei("Sessions")._output.repertoire, # Pour aller vers mes fichiers concaténés, sinon pour les originaux il faut pointer vers input
                    texte_bouton_choisir=f"Choisir extract IRIS session {IRIS.cei("Sessions")._codeExport} à nouveau"
                    )

@staticmethod
def charger_excel_IRIS_sessions(fe_IRIS_sessions:Optional[FichierExcel]=None, chemin_IRIS_sessions:Optional[Path]=None) -> FichierExcel:   
    # Vérifie existance de fe_IRIS sinon on le charge
    if fe_IRIS_sessions is None:
        timer.debut("Lecture fichier IRIS sessions")

        # S'il n'y a pas de chemin, alors l'utilisateur le pointe
        if chemin_IRIS_sessions is None:
            chemin_IRIS_sessions = IRIS.choisir_fichier_IRIS_sessions()

        # On charge le fichier IRIS session
        fe_IRIS_sessions = FichierExcel.depuis_fichier(chemin_fichier=chemin_IRIS_sessions)
        
        # Alias
        df_IRIS_sessions = fe_IRIS_sessions._tableaux["Sessions"]._df
        
        # On convertit trigramme formation et code_IRIS
        df_IRIS_sessions = IRIS.convertit_types_colonnes_df_sessions(fe_IRIS_sessions)

        # On trie
        df_IRIS_sessions = df_IRIS_sessions.sort_values(by="Date début ses.")  # Trie par "Date début ses."
        
        
        timer.fin()

    return fe_IRIS_sessions

@staticmethod
def convertit_types_colonnes_df_sessions(fe_IRIS_sessions:FichierExcel) -> DataFrame:

    df_IRIS_sessions = fe_IRIS_sessions._tableaux["Sessions"]._df  # Alias

    # On convertit
    df_IRIS_sessions["Trigramme formation"] = df_IRIS_sessions["Trigramme formation"].astype(str)  # Retype "Trigramme formation"
    df_IRIS_sessions["Code IRIS"] = df_IRIS_sessions["Code IRIS"].astype(str)  # Retype "Code IRIS"

    return df_IRIS_sessions

"""

# ============================
# === Affichage - OBSOLETE === 
# ============================
"""
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
"""
