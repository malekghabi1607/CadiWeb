from __future__ import annotations

from vte.config import UNITE
from .office import *

import math
from babel.dates import format_date

from dataclasses import dataclass
from collections import defaultdict
from tabulate import tabulate

from mailmerge import MailMerge
from tkinter import ttk, messagebox

from pprint import pprint

from typing import TYPE_CHECKING
if TYPE_CHECKING:
    import config as ConfigType  # pour que Pylance ait une base d’autocomplétion

config, config_extractsIRIS, user_config = charger_config()

config:ConfigType  # type hint explicite

### --------------------------------------------------------------------
#  Définitions classes et fonctions spécifiques INSTN
### --------------------------------------------------------------------


@dataclass
class InfosExportsIRIS:
    repertoire:Optional[Path]
    chemin_fichier:Optional[Path]
    nom_onglet:Optional[str]
    nbLignes_avantET:Optional[int]

class PropExportIRIS:
    """C'est une fabrique. Contient toutes les propriétés des fichiers Excel issus des Exports IRIS (structure de configuration d'un export IRIS)"""
    # === Constructeurs ===
    def __init__(self, 
                nom_typeExport:str, 
                codeExport:str, 
                
                repertoire_input:Path, 
                nom_onglet_input:str = "Data", 
                nbLignes_avantET_input:int = 0, 
                
                repertoire_modele:Optional[Path] = config.REPERTOIRES_MODELES, 
                nom_fichier_modele:Optional[str|Path] = None, 
                
                repertoire_output:Optional[Path] = config.REPERTOIRE_EXCEL_IRIS_OUTPUT, 
                nom_fichier_output:Optional[str|Path] = None
            ):
        # Type d'export
        self._nom_typeExport:str = nom_typeExport
        self._codeExport:str = codeExport

        if not nom_fichier_modele:
            nom_fichier_modele = Path(f"{codeExport}_{nom_typeExport}-Modèle.xlsx")
        if not nom_fichier_output:
            nom_fichier_output = Path(f"{codeExport}_{nom_typeExport}-COMPLET-{date.today():%Y.%m.%d}.xlsx")

        # Informations input données (i.e. extracts natifs d'IRIS)
        self._input:InfosExportsIRIS = InfosExportsIRIS(
            repertoire=repertoire_input,
            chemin_fichier=None,
            nom_onglet=nom_onglet_input,
            nbLignes_avantET=nbLignes_avantET_input
            )

        # Informations sur le modèle Excel à employer pour remplir l'output
        self._modele:InfosExportsIRIS = InfosExportsIRIS(
            repertoire=repertoire_modele,
            chemin_fichier=repertoire_modele / nom_fichier_modele,
            nom_onglet=nom_typeExport,
            nbLignes_avantET=None
            )

        # Informations output
        self._output:InfosExportsIRIS = InfosExportsIRIS(
            repertoire=repertoire_output,
            chemin_fichier=repertoire_output / nom_fichier_output,
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



class IRIS:
    "C'est la classe qui contient l'environnement pour bosser sur des fichiers Exports IRIS"

    # === VARIABLES DE CLASSE ===
    # --- Paramètres d'environnement - exports IRIS

    _dict_DE_IRIS = {}
    _dict_DE_IRIS["CodesExports"] = {
        "Sessions" : "R04110",
        "Formations" : "R0304",
        "Ventes" : "R04301",
        "Inscriptions" : "R04500",
        }
    _dict_DE_IRIS["PropExportIRIS"] = {
        "Sessions" : config.IRIS_SESSIONS, 
        "Formations" : config.IRIS_FORMATIONS,
        "Ventes" : config.IRIS_VENTES,
        "Inscriptions" : config.IRIS_INSCRIPTIONS
        }
    _dict_DE_IRIS["Fichiers"] = {
        "Sessions" : config_extractsIRIS._tSessions,
        "Formations" : config_extractsIRIS._tFormations,
        "Ventes" : config_extractsIRIS._tVentes,
        "Inscriptions" : config_extractsIRIS._tInscriptions
        }


    # Association colonnes excel avec command control Word

    # --- Paramètres utilisateur

    # --- Autres variables de la classe
    # Type d'export
    _nom_typeExport:str
    _codeExport:str

    # Informations génériques sur les exports, modèles et output (dépend du type d'export)
    _input:InfosExportsIRIS  # Informations input données
    _modele:InfosExportsIRIS  # Informations sur le modèle Excel à employer pour remplir l'output
    _output:InfosExportsIRIS  # Informations output

    _df_tableau:Optional[pd.DataFrame] = None  # dataframe du fichier IRIS
    _df_chemins:Optional[pd.DataFrame] = None  # dateframe des chemins CSV à traiter

    # === CONSTRUCTEUR ===
    def __init__(self, prop:PropExportIRIS, chemins_fichiersInput:Optional[Path|tuple[Path, ...]]=None):
        # Type d'export
        self._nom_typeExport = prop._nom_typeExport
        self._codeExport = prop._codeExport

        # Informations génériques sur les exports, modèles et output (dépend du type d'export)
        self._input = prop._input  # Informations input données
        self._modele = prop._modele  # Informations sur le modèle Excel à employer pour remplir l'output
        self._output = prop._output  # Informations output

        self._chemins_fichiersInput:Tuple[Path, ...]
        # S'il n'y a pas de chemin_fichiersInput de donné, c'est qu'il faut les sélectionner manuellement
        if not chemins_fichiersInput:
            self._choisirFichiers_filedialog()
        else:
            self._chemins_fichiersInput = convertir_tuple_path(chemins_fichiersInput)


    @classmethod
    def avecLecture(cls, propExportIRIS:PropExportIRIS, chemins_fichiersInput:Optional[Path|tuple[Path, ...]]=None) -> IRIS:
        """
        On lit le/les extract(s) IRIS et on le/les stocke dans un seul dataframe self._df_tableau
        """
        instance = cls(prop=propExportIRIS, chemins_fichiersInput=chemins_fichiersInput)
        instance._lire_extractIRIS()
        return instance


    @classmethod
    def avecEcritureOutputDefaut(cls, propExportIRIS:PropExportIRIS, chemins_fichiersInput:Optional[Path|tuple[Path, ...]]=None) -> IRIS:
        
        instance = cls.avecLecture(propExportIRIS, chemins_fichiersInput)

        # On récupère le chemin du modèle à ouvrir
        chemin_fichier = instance._modele.chemin_fichier
        #nom_onglet = instance._output.nom_onglet

        # Création du chemin pour l'output
        chemin_fichier_output = instance._output.chemin_fichier
        


        # On ouvre le modèle et tous ses tableaux structurés
        #fe_modele = FichierExcel.depuis_fichier(chemin_fichier=chemin_fichier)
        fe_modele = FichierExcel.depuis_modele(chemin_modele=chemin_fichier, chemin_fichier_sauv=chemin_fichier_output)

        # On copie le DataFrame instance._df_tableau (le dataframe créé par cls.avecLecture) avec les nouvelles données dans le modèle
        fe_modele._tableaux[instance._nom_typeExport].ecrit_dataFrame_dans_tableauStructure(instance._df_tableau, supprimeDonneesEtRemplace=True)
        
        # On écrit les références des fichiers copiés dans le tableau structuré "Imports"
        instance._df_chemins = pd.DataFrame(instance._chemins_fichiersInput, columns=['Chemin fichier'])
        fe_modele._tableaux["Imports"].ecrit_dataFrame_dans_tableauStructure(df=instance._df_chemins, supprimeDonneesEtRemplace=True)
        
        #On enregistre et on ferme (par précaution car copieformat xlwings sauvegarde)
        fe_modele.save()    
        fe_modele.close()

        return instance
 

    # === Méthodes statiques
    @staticmethod
    def mettreAJourTousLesExportsIRIS_auto(tuple_types:Tuple(str)) -> None:
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

        # On traite à la suite
        for clef in IRIS._dict_DE_IRIS["CodesExports"].keys():
            if clef in tuple_types:
                print(Style.BRIGHT + Fore.YELLOW + "\nTraitement des exports " + clef)
                fichiers_formates = ["\n\t" + f for f in IRIS._dict_DE_IRIS["Fichiers"][clef]]
                print("Liste des fichiers traités : " + ", ".join(fichiers_formates))
                chemins_fichiersInput = tuple((IRIS._dict_DE_IRIS["PropExportIRIS"][clef]._input.repertoire / nom) for nom in IRIS._dict_DE_IRIS["Fichiers"][clef])
                IRIS.avecEcritureOutputDefaut(IRIS._dict_DE_IRIS["PropExportIRIS"][clef], chemins_fichiersInput=chemins_fichiersInput)

    @staticmethod
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
        for clef, codeIRIS in IRIS._dict_DE_IRIS["CodesExports"].items():
            if clef in tuple_types:
                traitement = IRIS(IRIS._dict_DE_IRIS["PropExportIRIS"][clef])
                dict_traitements[clef] = traitement

        # On traite à la suite
        for clef in dict_traitements.keys():
            print(Style.BRIGHT + Fore.YELLOW + "\nTraitement des exports " + clef)
            fichiers_formates = ["\n\t" + f for f in dict_traitements[clef]._chemins_fichiersInput]
            print("Liste des fichiers traités : " + ", ".join(fichiers_formates))
            #print("Liste des fichiers traités : " + ", ".join(dictEI[clef].nomsFichiers_exportIRIS))
            #dict_traitements[clef].lire_extractIRIS()
            #dict_traitements[clef].ecrit_dataFrame_dans_tableauStructure(df=instance._df_chemins, supprimeDonneesEtRemplace=True)
            IRIS.avecEcritureOutputDefaut(IRIS._dict_DE_IRIS["PropExportIRIS"][clef], chemins_fichiersInput=dict_traitements[clef]._chemins_fichiersInput)



    # === Méthodes internes ===
    def _choisirFichiers_filedialog(self):
        # Lister/sélectionner les documents à concaténer
        cheminsExcel_str = filedialog.askopenfilenames(title="Sélectionner les fichiers " + self._nom_typeExport + " (" + self._codeExport + ") Excel à concaténer", filetype=[("Fichiers Excel", "*.xlsx")], initialdir=self._input.repertoire)
        # Gestion du cas où il y a non-sélection de fichiers
        if not cheminsExcel_str:
            vlog("click sur cancel du filedialog → Pas de chemins de fichier")
        
        self._chemins_fichiersInput = tuple(Path(p) for p in cheminsExcel_str)
        

    def _lire_extractIRIS(self):
        """
        Crée le DataFrame pour l'export IRIS. On le stocke dans self._df_tableau
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
                fichier = chemin.name
                taille = chemin.stat().st_size  # taille en octets 
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
                df_colonnes_sup = self._df_tableau['N° Session'].apply(self._extraire_infos_numSessionIRIS)
                # A cause de certains éléments None ou NaN, Pandas type la colonne en float. Je la retype en Int64 qui permet de stocker des NaN avec des entiers, contrairement au type int standard.
                df_colonnes_sup['Année'] = df_colonnes_sup['Année'].astype('Int64')
                #print(df_colonnes_sup)

            # Cas Formations
            case "R0304":
                # On extrait / retravaille les informations de la colonne 'Référence'
                df_colonnes_sup = self._df_tableau['Référence'].apply(self._extraire_infos_referenceFormationIRIS)
                # A cause de certains éléments None ou NaN, Pandas type la colonne en float. Je la retype en Int64 qui permet de stocker des NaN avec des entiers, contrairement au type int standard.
                df_colonnes_sup['Année'] = df_colonnes_sup['Année'].astype('Int64')
                #print(df_colonnes_sup)

        # Ajouter les colonnes supplémentaires au DataFrame principal ssi le DataFrame df_colonnes_sup exite
        if df_colonnes_sup is not None:
            self._df_tableau = pd.concat([self._df_tableau, df_colonnes_sup], axis=1)
    
    def _extraire_infos_numSessionIRIS(self, reference:str) -> pd.Series:
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

    def _extraire_infos_referenceFormationIRIS(self, reference:str) -> pd.Series:
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

class ContexteFormation:
    """
    Gère le fichier d'évaluation global d'une formation donnée.
    Utilisable comme contexte : ouvre au début, sauvegarde et ferme à la fin.
    """
    def __init__(self, trigramme_formation: str):
        self.trigramme_formation = trigramme_formation
        self.fe_evaluations_formation: Optional[FichierExcel] = None
        self.df_evaluations_formation: Optional[pd.DataFrame] = None
        self.supprimeDonneesEtRemplace_evaluations_formation:Optional[bool] = None

    def __enter__(self):
        print(f"🔹 Ouverture du fichier global pour {self.trigramme_formation}")

        # On ouvre ou on créée (si inexistant) le fichier Excel qui concatène toutes les sessions d'une formation
        #instance._fe_evaluations_formation, instance._df_evaluations_formation,
        self.supprimeDonneesEtRemplace_evaluations_formation = self._ouvrir_ou_creer_evaluationsFormation()

        self.fe_evaluations_formation = FichierExcel(depuis_formation=self.trigramme_formation)
        self.df_evaluations_formation = self.fe_evaluations_formation.tableau_principal
        EvalStat._contexte_formation = self  # ← toutes les instances peuvent y accéder
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.fe_evaluations_formation:
            print(f"💾 Sauvegarde et fermeture du fichier global pour {self.trigramme_formation}")
            self.fe_evaluations_formation.sauvegarder()
            self.fe_evaluations_formation.fermer()
        EvalStat._contexte_formation = None  # Nettoyage


    @classmethod
    def _ouvrir_ou_creer_evaluationsFormation(cls, trigramme_formation:str) -> bool: #-> Tuple[FichierExcel, pd.DataFrame, bool] :
        """
        Ouvre ou crée le fichier Excel d'évaluations globales pour une formation donnée.

        Cette méthode construit le chemin vers le fichier d'évaluations correspondant au 
        trigramme de la formation. Si ce fichier existe, il est ouvert et les données 
        des stagiaires sont chargées dans un DataFrame. Sinon, un nouveau fichier est 
        créé à partir d'un modèle, et les données seront à initialiser.

        Args:
            trigramme (str): Code trigramme de la formation.

        Returns:
                - Un booléen indiquant si les anciennes données doivent être supprimées et remplacées 
                (`True` si nouveau fichier créé, `False` sinon).
        """ 
        # On définit le chemin vers les évaluations de la formation (le fichier qui va concaténer toutes les évaluation d'une formation)
        chemin_excel_evaluations_formation = config.format_path(config.CHEMIN_EXCEL_EVALUATIONS_FORMATION, trigramme_formation=trigramme_formation)

        # On vérifie que le répertoire dédié existe sinon on le créée : \\instnt\partage\FORMATIONS_C\###\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\AAAA
        chemin_excel_evaluations_formation.parent.mkdir(parents=True, exist_ok=True)
        
        # On regarde si Evaluation-Stagiaires-Global-XXX.xlsx existe
        if chemin_excel_evaluations_formation.is_file():
            # Alors on l'ouvre

            cls._fe_evaluations_formation = FichierExcel.depuis_fichier(chemin_fichier=chemin_excel_evaluations_formation)
            #self._fe_evaluations_formation._tableaux["CSV_stagiaires"].charge_df()
            #self._fe_evaluations_formation._tableaux["Stagiaires"]._df['Trigramme formation'] = self._fe_evaluations_formation._tableaux["Stagiaires"]._df['Trigramme formation'].astype(str)

            
            # On récupère la liste des sessions déjà intégrées (liste des codes et des chemins) → Ce sera pour écrire dans le tkinter
            #dico_sessionsDejaTraitees = dict(
            #df_formation_stagiaires[df_formation_stagiaires["Trigramme formation"] == trigramme]     # 1. filtre sur le trigramme
            #.drop_duplicates(subset=["Code IRIS", "Chemin fichier CSV"])  # 2. élimine les doublons
            #[["Code IRIS", "Chemin fichier CSV"]]            # 3. sélection des colonnes
            #.values                               # 4. valeurs du DF
            #    )        

            # Il ne faudra pas supprimer les anciennes données de fe_evaluations_formation
            supprimeDonneesEtRemplace = False

            vlog.ajouter_message("Ouverture EvalStat Global formation", cls._fe_evaluations_formation.chemin_fichier, style=["vert"])
        else :
            # On créée le fichier excel à partir du modèle
            cls._fe_evaluations_formation = FichierExcel.depuis_modele(
                chemin_modele = cls._CHEMIN_MODELE_EXCEL_EVALUATIONS_STAGIAIRES, 
                chemin_fichier_sauv = chemin_excel_evaluations_formation
                )
            
            # Il faudra supprimer les anciennes données de fe_evaluations_formation
            supprimeDonneesEtRemplace = True

            vlog.ajouter_message("Création EvalStat Global formation", self._fe_evaluations_formation.chemin_fichier, style=["vert"])

        # On crée un alias pour le dataframe des évaluations de la formation
        self._df_evaluations_formation = self._fe_evaluations_formation._tableaux["Stagiaires"]._df  # Alias

        #return self._fe_evaluations_formation, self._df_evaluations_formation, supprimeDonneesEtRemplace
        return supprimeDonneesEtRemplace



class EvalStat:


    # === VARIABLES DE CLASSE COMMUNES A TOUTES LES INSTANCES ===

    # Extract IRIS Sessions (R04110)
    _fe_IRIS_sessions:Optional[FichierExcel] = None  # Fichier Excel qui contient les extracts IRIS Sessions (ou a minima celles de la période en cours)

    # Fichier Excel évaluation de la formation (peut être commun à plusieurs instances si même trigramme formation) → Sera mis à jour
    _trigramme_formation:Optional[str] = None
    _fe_evaluations_formation:Optional[FichierExcel] = None # Fichier Excel qui contient tous les CSV d'évaluation d'une formation
    _df_evaluations_formation:Optional[pd.DataFrame] = None # DataFrame de self._fe_evaluations_formation (Alias)
       

    #_statuts_csv: dict[Path, str] = {}  # ex: {Path("...csv"): "traite" | "exclu" | "probleme"}


    # === Colonnes du CSV selon traitement à avoir ===
    # Colonnes descriptives à recopier
    _colonnes_csv_fixes = [
        "Chemin fichier CSV", "Prénom", "Nom", "Entreprise", "Code session"]

    # Colonnes avec note/commentaire en binôme
    _colonnes_csv_avec_commentaires = [
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
    _colonnes_csv_commentaires_seuls = [
        "Comment avez-vous connu cette formation ?",
        "Commentaires, remarques, suggestions"]

    # Colonnes note seule (il se trouve que je vais aussi devoir convertir le booléen)
    _colonnes_csv_bool = [
        "Recommanderiez-vous cette formation ?"]
    
    # === Colonnes de l'extract IRIS Sessions à récupérer ===
    _colonnes_sessions = [
        "N° Session",
        "Formation",
        "Trigramme formation",
        "Code IRIS",
        "Date début ses.",
        "Année début ses.",
        "Type de formation",
        "Trigramme RP",
        "Trigramme AF",
        "Nb. Présents"]

    




    # === CONSTRUCTEUR ===
    def __init__(self) -> None:

        # Variables pour un EvalStat individuel        
        #self._trigramme_formation:Optional[str] = None
        self._codeIRIS:Optional[str] = None

        self._chemin_csv_evaluations_stagiaires:Optional[Path] = None  # Fichier csv EvalStat stagiaire individuel
        #self._chemin_excel_evaluations_stagiaires:Optional[Path] = None  # Fichier xlsx EvalStat stagiaire individuel qu'on va créer à partir du CSV
        self._fe_evaluations_stagiaires:Optional[FichierExcel] = None  # Objet contenant les données EvalStat stagiaire individuel

        # === Résultat du traitement ===
        self._statut_csv: str | None = None  # ex: "traite", "exclu", "probleme", "CSV vide"

    # === CONSTRUCTEURS ALTERNATIFS ===    
    @classmethod
    def ouvrir_ou_creer_evaluationsFormation(cls, trigramme_formation:str) -> EvalStat:
        """
        A partir d'un trigramme de foramtion, on ouvre et on charge le fichier excel qui concatène tous les CSV d'une formation 
        """
        # On crée l'instance et on complète les infos avec les valeurs facultatives
        instance = EvalStat()

        instance._trigramme_formation = trigramme_formation

        #instance._ouvrir_fe_evaluations_formation()
        instance._ouvrir_ou_creer_evaluationsFormation()

        return instance

    @classmethod
    def depuis_chemin_csv_evaluations_stagiaires(cls, chemin_csv_stagiaires:Path|str, ouvrirDossier:bool=False, remplace_df:bool=False) -> EvalStat:
        # On convertit le Path si nécessaire
        if isinstance(chemin_csv_stagiaires, str):
            chemin_csv_stagiaires = Path(chemin_csv_stagiaires)

        # On créée l'instance
        instance = EvalStat()

        # On récupère IRIS sessions, seulement si nécessaire
        instance._charge_iris_sessions()


        timer.debut(f"\n{Style.BRIGHT}{Fore.YELLOW}Gestion de la session {chemin_csv_stagiaires.name}")

        # Par défaut, on traite le CSV
        traiterCSV = True

        # On vérifie que chemin_csv_session n'est pas déjà dans le fichier session pour savoir si on l'exclue du traitement
        if instance._df_evaluations_formation is not None:
            if str(chemin_csv_session) in instance._df_evaluations_formation["Chemin fichier CSV"].drop_duplicates().tolist():  
                traiterCSV = False
                instance._chemins_csv_exclus.append(chemin_csv_session)
                print(f"Exclusion car csv déjà dans le fichier global : {chemin_csv_session}")
                vlog.ajouter_message("Info", f"Exclusion car csv déjà dans le fichier global {chemin_csv_session}", style=["orange"])

        # si le chemin est avec un raccourci réseau alors on récupère le chemin en entier
        instance._chemin_csv_evaluations_stagiaires = chemin_vers_unc(chemin_csv_stagiaires)

        # On récupère le trigramme de la formation depuis le chemin du CSV
        instance._trigramme_formation = recupere_trig_formation_depuis_chemin(instance._chemin_csv_evaluations_stagiaires)

        # On génère le fichier Excel du CSV à partir du modèle
        instance._construit_FichierExcel_depuis_CSV(remplace_df=remplace_df)

        # Ouverture du dossier à la fin
        if ouvrirDossier:
            ouvrir_dossier(instance._chemin_excel_evaluations_stagiaires.parent)

        timer.fin()
        return instance        

    @classmethod
    def depuis_tuple_csv_stagiaires(cls, tuple_csv_stagiaires:Tuple[Path], ouvrirDossier:bool=False) -> EvalStat:
        """
        A partir d'un tuple de chemins de CSV stagiaire (il peut il y avoir plusieurs trigrammes de formations différents)
        Permet de générer :
           - le fichier excel stagiaires de chaque session (via le CSV)
           - le fichier excel stagiaires de chaque formation (celui qui concatène tous les CSV d'une session) [Il est créé ou on l'append avec les nouvelles valeurs]
        
        On exclue du traitement les chemin_csv_session qui sont déjà dans le FE formation (on considère que le CSV a déjà été traité)

        """

        
        # On crée l'instance et on complète les infos avec les valeurs facultatives
        instance = EvalStat()
        
        df_formation_csv = None
        df_formation_stagiaires = None

        # On convertit le tuple de strings en dictionnaire avec les trigrammes formation en clef
        dico_chemins_csv_session = defaultdict(list)  #Dictionnaire spécial : lorsqu’on accède à une clé qui n’existe pas encore, il va automatiquement créer une nouvelle entrée avec une valeur par défaut, ici une liste vide (list())
        for chemin in tuple_csv_stagiaires:
            trigramme_formation = recupere_trig_formation_depuis_chemin(chemin)
            dico_chemins_csv_session[trigramme_formation].append(chemin)
        dico_chemins_csv_session = dict(dico_chemins_csv_session)  # Optionnel : conversion en dict normal


        # Pour chaque trigramme on va traiter chaque session et soit créer soit append le fichier excel global de la formation
        #for chemins_csv_session in tuple_csv_stagiaires: #chemins est la liste des chemins des évaluations pour chaque sessions de ce trigramme formation~
        for trigramme_formation in dico_chemins_csv_session.keys() :  # chemins_csv_session est la liste des chemins des évaluations pour chaque sessions de ce trigramme formation~
            print(f"\n\n{Style.BRIGHT}{Fore.RED}Gestion des formations {trigramme_formation}")

            instance._trigramme_formation = trigramme_formation

            # Pour chaque chemin de session, on crée le fe_stagiaire dédié de la session et on ajoute les lignes de son dataframe au dataframe de fe_evaluations_formation
            for chemin_csv_session in dico_chemins_csv_session[trigramme_formation]:
                # TODO faire traitement avec fonction



                # Traitement du CSV
                if traiterCSV:
                    try:  
                        #timer.debut("Traiter_evalStat.depuis_chemin_csv_evaluations_stagiaires")
                        traite_csv_session = EvalStat.depuis_chemin_csv_evaluations_stagiaires(chemin_csv_session, fe_IRIS_sessions=instance._fe_IRIS_sessions, ouvrirDossier=ouvrirDossier, remplace_df=True)
                    except Exception as e:
                        instance._chemins_csv_probleme.append(chemin_csv_session)
                        print(f"❌ Erreur de traitement sur le CSV {chemin_csv_session} (CSV exclu) :", e)
                    else:
                        #timer.debut("Copie des Dataframe csv et stagiaires")

                        # On vérifie l'existance de traite_csv_session._fe_evaluations_stagiaires (i.e. le csv avait au moins 1 ligne)
                        if traite_csv_session._fe_evaluations_stagiaires:
                            #Si df_formation_csv est vide, il faut l'initialiser avec le premier df sinon on concatène
                            if df_formation_csv is None:
                                df_formation_csv = traite_csv_session._fe_evaluations_stagiaires._tableaux["CSV_stagiaires"]._df.copy()
                                df_formation_stagiaires = traite_csv_session._fe_evaluations_stagiaires._tableaux["Stagiaires"]._df.copy()
                            else:
                                df_formation_csv = pd.concat([df_formation_csv, traite_csv_session._fe_evaluations_stagiaires._tableaux["CSV_stagiaires"]._df], ignore_index=True)
                                df_formation_stagiaires = pd.concat([df_formation_stagiaires, traite_csv_session._fe_evaluations_stagiaires._tableaux["Stagiaires"]._df], ignore_index=True)
                            
                            # On ajoute le chemin au tuple des éléments traités
                            instance._chemins_csv_traites.append(chemin_csv_session)  # Déjà fait dans depuis_chemin_csv_evaluations_stagiaires
                            #vlog.ajouter_message("Fichiers traités", chemin_csv_session, style=["vert"])
                pass

            # On concatène, on sauve et on ferme le fe de tous les CSV de la formation
            if instance._chemins_csv_traites :
                print(f"\n\n{Style.BRIGHT}{Fore.RED}Écriture du fichier global des évaluations de la formation {trigramme_formation}")
                timer.debut("Écriture, sauvegarde et fermeture")
                instance._fe_evaluations_formation._tableaux["CSV_stagiaires"].ecrit_dataFrame_dans_tableauStructure(df_formation_csv, supprimeDonneesEtRemplace = supprimeDonneesEtRemplace_evaluations_formation)
                instance._fe_evaluations_formation._tableaux["Stagiaires"].ecrit_dataFrame_dans_tableauStructure(df_formation_stagiaires, supprimeDonneesEtRemplace = supprimeDonneesEtRemplace_evaluations_formation)
                instance._fe_evaluations_formation._tableaux["CSV_stagiaires"].charge_df()
                instance._fe_evaluations_formation._tableaux["Stagiaires"].charge_df()

                instance._df_evaluations_formation = instance._fe_evaluations_formation._tableaux["Stagiaires"]._df  # Alias

                instance._fe_evaluations_formation.save()
                instance._fe_evaluations_formation.close()
                timer.fin()
            
                # Actualisation des TCD
                instance._fe_evaluations_formation.actualiser_TCD()

        return instance

    # === METHODES INTERNES DE CLASSE === 
    @classmethod
    def _charge_iris_sessions(cls) -> FichierExcel:
        """Retourne l’extract IRIS, en le chargeant si nécessaire."""

        if cls._fe_IRIS_sessions is None:
            timer.debut("Lecture fichier session")
            chemin_iris_sessions = choisir_fichier_iris_sessions()
            cls._fe_IRIS_sessions = FichierExcel.depuis_fichier(chemin_fichier=chemin_iris_sessions)
            timer.fin()
            
            # Quand je ferai la jointure plus tard sur "Code IRIS", il faudra que ce soit avec des strings
            cls._fe_IRIS_sessions._tableaux["Sessions"]._df["Code IRIS"] = cls._fe_IRIS_sessions._tableaux["Sessions"]._df["Code IRIS"].astype(str)
        return cls._fe_IRIS_sessions





    # === METHODES INTERNES ===
    def _construit_FichierExcel_depuis_CSV(self, remplace_df:bool=False):
        """
        Import et traitement du CSV d'evalStat
        """

        # On récupère l'encodage et on importe le CSV dans un DataFrame
        codage_csv = trouve_encodage_csv(self._chemin_csv_evaluations_stagiaires)
        df_csv_stagiaires = pd.read_csv(self._chemin_csv_evaluations_stagiaires, sep=';', encoding=codage_csv)  # Ouverture du CSV et mise dans un DataFrame
        
        if (not df_csv_stagiaires.empty):
            ###
            # === Traitement première partie du dataframe du CSV ===
            ###

            # Prise en compte qu'on a plusieurs formats de CSV : on doit traiter des colonnes en + ou - en conséquences
            if "Date de fin" in df_csv_stagiaires.columns:
                # Cas 1 (nouveau format de csv) : supprimer "Date de fin" → Test : r"P:\FORMATIONS_C\54C\P07-bilan-sessions-et-bilan-formation\rapports-sessions-evaluations\2023-06-S14317 UEM\S-14317-FC23-54C-VTE-LRA-Stagiaires.csv"
                df_csv_stagiaires = df_csv_stagiaires.drop(columns=["Date de fin"])
            else:
                # Cas 2 (ancien format de csv) : supprimer la 2e et 3e colonne (indices 1 et 2) → Test : r"P:\FORMATIONS_C\22B\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\S-17606 - 22B - 06-2025\S-17606-FC25-22B-VTE-CAR-Stagiaires.csv"
                df_csv_stagiaires = df_csv_stagiaires.drop(df_csv_stagiaires.columns[[1, 2]], axis=1)
            
            # On rajoute le chemin du CSV en première colonne
            df_csv_stagiaires.insert(0, "Chemin fichier CSV", str(self._chemin_csv_evaluations_stagiaires))

            # Définition self._codeIRIS. Sinon Non existant, on récupère le numéro IRIS depuis le CSV (c'est la plus sur)
            if self._codeIRIS is None:
                match = re.search(r"\b\d{5}\b", str(self._chemin_csv_evaluations_stagiaires))
                if match:
                    self._codeIRIS = match.group(0)
                else:
                    self._demander_code("Code IRIS")

            # Met à jour ou crée la colonne "Code session" avec self._codeIRIS
            df_csv_stagiaires["Code session"] = self._codeIRIS

            # Mise au format jj/mm/aaaa de la colonne "Date" (si elle existe)
            if "Date" in df_csv_stagiaires.columns:
                try:
                    df_csv_stagiaires["Date"] = pd.to_datetime(df_csv_stagiaires["Date"], dayfirst=True, errors="coerce").dt.strftime("%d/%m/%Y")  # dayfirst=True indique que le premier nombre correspond au jour (format jj/mm/aaaa)
                except Exception as e:
                    print(f"Erreur de conversion de la colonne Date : {e}")

            ###
            # === Save / reload Excel ===
            #
            # A cause des espaces à la con qui trainent dans les noms des colonnes des CSV, je vais reload le dataframe depuis l'excel que je viens de créer car les colonnes du modèle sont bien nommées
            # Ainsi on sauve ici plutôt qu'à la fin et on reload le DataFrame
            ### 
            
            # On colle le dataframe dans le modèle et on sauve
            self._fe_evaluations_stagiaires = FichierExcel.depuis_modele(chemin_modele=self._CHEMIN_MODELE_EXCEL_EVALUATIONS_STAGIAIRES, chemin_fichier_sauv=self._chemin_excel_evaluations_stagiaires)
            if remplace_df:
                self._fe_evaluations_stagiaires._tableaux["CSV_stagiaires"].remplace_df(df_csv_stagiaires)
            self._fe_evaluations_stagiaires._tableaux["CSV_stagiaires"].ecrit_dataFrame_dans_tableauStructure(df_csv_stagiaires, supprimeDonneesEtRemplace=True)
            self._fe_evaluations_stagiaires.save(self._chemin_excel_evaluations_stagiaires)
            self._fe_evaluations_stagiaires.close()

            # On recharge le modèle
            self._fe_evaluations_stagiaires = FichierExcel.depuis_fichier(self._chemin_excel_evaluations_stagiaires)
            #self._fe_evaluations_stagiaires._tableaux["CSV_stagiaires"].charge_df()
            df_csv_stagiaires = self._fe_evaluations_stagiaires._tableaux["CSV_stagiaires"]._df
            
            ######
            # === On crée la seconde partie du DataFrame qui sera dans l'onglet "Stagiaire" ===
            #
            #  On va découper le dataframe du CSV selon les différents critères et mettre dans un dataframe qu'on pourra exploiter par un TCD
            ######
            # Nouveau DataFrame à remplir
            df_long = []

            # Parcours des lignes
            for _, row in df_csv_stagiaires.iterrows():
                #print("Ligne en cours : ")
                #print(row)
                base = {col: row[col] for col in self._colonnes_csv_fixes}  # Création des colonnes qui seront répétées à chaque fois
                base["NOM Prénom"] = f"{str(row['Nom']).upper()} {row['Prénom']}".strip()  # Création du champ "NOM Prénom"

                # Cas 1 : colonnes avec note + commentaire associé
                for critere in self._colonnes_csv_avec_commentaires:
                    if critere in row:
                        #print(f"'{critere}'")
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

            # Pour faire le merge, il faut que les colonnes soient de même type (là "Code session" est de type int64 et "Code IRIS" est de type object (souvent des chaînes de caractères)).
            # Comme je ne peux être sûr que tous les "Code IRIS" issu des CSV soient bien convertibles en int (c’est-à-dire pas de chaînes vides, NaN, ou autres caractères non numériques), alors je passe par des strings
            df_stagiaires["Code session"] = df_stagiaires["Code session"].astype(str)
            #self._fe_IRIS_sessions._tableaux["Sessions"]._df["Code IRIS"] = self._fe_IRIS_sessions._tableaux["Sessions"]._df["Code IRIS"].astype(str)
            
            # On récupère uniquement les colonnes souhaitées dans une vue pour faciliter le codage (c'est un alias)
            df_sessions_filtre = self._fe_IRIS_sessions._tableaux["Sessions"]._df[self._colonnes_sessions]
            
            # On fait la jointure entre df_stagiaires et df_sessions_filtre
            #timer.debut("Création du DataFrame Stagiaires (jointure)")
            df_stagiaires = df_stagiaires.merge(
                df_sessions_filtre,
                left_on="Code session",
                right_on="Code IRIS",
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
                self._fe_evaluations_stagiaires._tableaux["Stagiaires"].remplace_df(df_stagiaires)
            
            # On écrit et on sauve
            #timer.debut("On écrit le DataFrame, on met à jour les TCD et on sauve")
            self._fe_evaluations_stagiaires._tableaux["Stagiaires"].ecrit_dataFrame_dans_tableauStructure(df_stagiaires, supprimeDonneesEtRemplace=True)
            self._fe_evaluations_stagiaires.save()


            # Màj des TCD
            self._fe_evaluations_stagiaires.actualiser_TCD()
            #timer.fin()

            self._definit_statut_csv("traité")

        else :
            self._definit_statut_csv("CSV vide")

    def _definit_statut_csv(self, statut:str) -> None:
        self._statut_csv = statut
        self._statuts_csv[self._statut_csv] = self._chemin_csv_evaluations_stagiaires

        if statut == "traité":
            vlog.ajouter_message("OK", f"Le CSV {self._chemin_csv_evaluations_stagiaires} a été traité.")
        elif statut == "CSV vide":
            vlog.print("Warning", f"Le CSV {self._chemin_csv_evaluations_stagiaires} est vide, mais on continue.")


    # === POPUP ===
    def _filedialog_csv(self, code_IRIS: int, trigramme_formation: Optional[str] = None) -> str | None:
        """
        Ouvre un filedialog pour demander à l'utilisateur de sélectionner un CSV.
        On pointe au mieux sur le répertoire des CSV de cette formation pour la boîte de dialogue.
        """
        
        if (trigramme_formation is None) and (self._chemin_excel_evaluations_formation is not None):
            trigramme_formation = recupere_trig_formation_depuis_chemin(self._chemin_excel_evaluations_formation)

        if trigramme_formation:
            chemin_repertoire_csv = optimiseCheminRepertoire(self._chemin_excel_evaluations_formation.parent)
        else:
            chemin_repertoire_csv = Path.cwd()

        return choisir_fichier(titre=f"Sélectionner le fichier CSV de la session {code_IRIS}",
                        types_fichiers=[("Fichiers CSV", "*.csv")],
                        dossier_initial=chemin_repertoire_csv,
                        obligatoire=False,
                        texte_bouton_choisir="Choisir CSV à nouveau",
                        texte_bouton_aucun="Pas de CSV pour cette session"
                        )

    @property
    def _chemin_excel_evaluations_stagiaires(self) -> Path:
        """
        Retourne le chemin Excel correspondant au CSV courant.
        Pour le nom de l'Excel output : on reprend le nom du csv et on remplace par xlsx.
        """
        return self._chemin_csv_evaluations_stagiaires.with_suffix(".xlsx")


class BilanSession:
    """
    C'est la classe qui contient tous les éléments de ma formation pour mon bilan de session V3
    """
    # === VARIABLES PARTAGÉES ENTRE TOUTES LES INSTANCES
    _chemin_excel_IRIS_sessions:Optional[Path] = None
    # Récupération du dernier export (extract IRIS le plus récent dans le répertoire)
    #_chemin_excel_IRIS_sessions:Optional[str] = obtenir_fichier_plus_recent_repertoire(
    #            config.REPERTOIRE_EXCEL_IRIS_SESSIONS,
    #            r"^R04110_Sessions.*"
    #        )  # Chemin vers le fichier Excel qui contient les extracts IRIS Sessions (ou a minima celles de la période en cours)
    _fe_IRIS_sessions:Optional[FichierExcel] = None  # Fichier Excel qui contient les extracts IRIS Sessions (ou a minima celles de la période en cours)
    _df_sessions:Optional[pd.DataFrame] = None  # DataFrame de _fe_IRIS_sessions (self._fe_IRIS_sessions._tableaux["Sessions"]._df)


    _CRITERES_A_ENLEVER:list[str] = [  # Critères à ne pas retenir pour le calcul des moyennes < 3
        "Comment avez-vous connu cette formation ?", "Avez-vous d'autres besoins de formation ?", "Commentaires, remarques, suggestions", "Recommanderiez-vous cette formation ?"]


    # === CONSTRUCTEURS ===
    def __init__(self) -> None:
        # Variables d’instance → propres à chaque bilan
        self._chemin_word_bilan_session_output: Optional[Path] = None  # Bilan de session

        self._codeFormation: Optional[str] = None
        self._annee: Optional[int] = None
        self._periode: Optional[str] = None

        self._df_sessions_filtre:pd.DataFrame = None  # dataframe de _fe_IRIS_sessions filtré avec le bon trigramme formation, la bonne année et la bonne période

        self._stats_stagiaires: Optional[dict] = None  # Dictionnaire des stats des CSV
        self._es: Optional[EvalStat] = None  # EvalStat global des évaluations stagiaires de la formation
        
        self._codes_IRIS_communs: list[str] = []  # Codes IRIS en commun entre le fichier Excel des sessions et le fichier Excel global des évaluations stagiaires de la formation
        self._codes_IRIS_absents_fin: list[str] = []  # Disparités restantes après traitement des CSV manquants entre le fichier Excel des sessions et le fichier Excel global des évaluations stagiaires de la formation

        self._exploitationBilan:dict[list] = {
            "Exploités pour les stats générales" : [],  # Exploités pour stats initiales → Dans _demande_sessions_a_exclure
            "Exploités pour les évaluations (CSV présents)" : [],  # Exploités pour les stats stagiaires → Dans _maj_evalstat_formation
            "Exclus des évaluations (CSV manquants)" : [],   # Exclus des évaluations car CSV stagiaires manquants → Dans _maj_evalstat_formation
            "Exclus des évaluations (problème traitement CSV)" : [],   # Exclus des évaluations car problème au traitement des CSV → Dans _maj_evalstat_formation
            "Exclus entièrement du bilan" : [],  # Exclus entièrement du bilan car sessions non réalisées ou mauvais RP (exclus par l'utilisateur) → Dans _demande_sessions_a_exclure
        }

        # Liste des champs de fusion du Word (pour la fonction .mergefields, il faut des str)
        self._titreFormation:str = ""
        #self._codeFormation:str = ""  # Déjà déclaré pour fonctionnement de la classe
        self._periodeSessionsEvaluees:str = ""
        self._nbSessionsEvaluees:str = ""
        self._numerosSessions:str = ""
        self._nbApprenants:str = ""
        self._rp:str = ""
        self._af:str = ""

        self._commentairesBilan:str = ""
        self._satisfactionGlobale_moy:str = ""
        self._satisfactionGlobale_com:str = ""
        self._recommandation_moy:str = ""
        self._commentairesRemarquesSuggestions_com:str = ""
        self._evalInf3_val:str = ""
        self._evalInf3_com:str = ""
        self._tauxRetours_val:str = ""

    @classmethod   
    def bilanUnique_parCodeIRIS(cls, code_IRIS:int) -> None:
        """
        Permet de générer un bilan de session par code IRIS
        Ex : BilanSession.bilanUnique_parPeriode(16411)
        """

        instance = cls()

        # Vérifier que code_iris est bien un entier à 5 chiffres, on le convertit en str
        est_code_IRIS_valide, instance._code_IRIS = verifier_code_iris(code_IRIS)

        if not est_code_IRIS_valide:
            vlog.log_erreur(f"Code IRIS en entrée non valide : {instance._code_IRIS}")
    
        # Ouverture / création du dataframe de l'extract IRIS sessions de la formation
        instance._creer_df_extractIRIS_sessions_codeIRIS()

        # Initialisation données
        instance._codeFormation = instance._df_sessions_filtre["Trigramme formation"].iloc[0]
        instance._annee = instance._df_sessions_filtre["Année début ses."].iloc[0]
        moisSession = mois_fr_depuis_date(instance._df_sessions_filtre["Date début ses."].iloc[0])
        numSession = instance._df_sessions_filtre["N° Session"].iloc[0]
        instance._periode = f"Session {numSession} uniquement ({moisSession} {instance._annee})"
        instance._periodeSessionsEvaluees = f"{numSession}"



        # On met à jour l'Excel evalstat de la formation si des sessions demandées par l'utilisateur ne s'y trouvent pas
        instance._maj_evalstat_formation()

        # On calcule les stats
        instance._calculer_stats_criteres()
        #pprint(instance._stats_stagiaires)

        # On construit le bilan de session (bilan de session V3)
        instance._bilanSessionV3()

        # On ouvre le word
        FichierWord.depuisFichier(chemin_fichier=instance._chemin_word_bilan_session_output, charger_contentControl=False, afficherWord=True)

        # On envoie un mail au chef d'unité pour la signature du pdf
        instance._envoyer_mail_chef_unite()

    @classmethod
    def plusieursBilans_parCodeIRIS(cls, liste_codes_IRIS:list[int]) -> None:
        """
        Permet de lancer une série de bilans de sessions à partir d'une liste d'entiers (codes IRIS) ex. [61235, 54231]
        """
        print(liste_codes_IRIS)
        for codeIRIS in liste_codes_IRIS:
            print(codeIRIS)
            cls.bilanUnique_parCodeIRIS(codeIRIS)

    @classmethod
    def plusieursBilans_parCodeIRIS_fichierConfig(cls) -> None:
        cls.plusieursBilans_parCodeIRIS(user_config.liste_codes_IRIS)


    @classmethod   
    def bilanUnique_parPeriode(cls, codeFormation:str, annee:int, periode:str) -> None:
        """
        Permet de générer un bilan de session selon une période qui est l'un de ces éléments : ["1er semestre", "2nd semestre", "Année"]
        Ex : BilanSession.bilanUnique_parPeriode("948", 2024, "Année")
        """

        instance = cls()
    
        instance._codeFormation = codeFormation
        instance._annee = annee
        instance._periode = periode
        instance._periodeSessionsEvaluees = f"{instance._periode} {instance._annee}"


        # Ouverture / création du dataframe de l'extract IRIS sessions de la formation
        instance._creer_df_extractIRIS_sessions_periode()

        # On met à jour df_sessions_filtre selon les sessions que souhaite garder / exclure l'utilisateur
        instance._demande_sessions_a_exclure()
        #print("\nÉtat de Excel sessions filtré sur période et trigramme :")
        #pprint(instance._df_sessions_filtre)

        # S'il n'y a plus de session à lire dans _df_sessions_filtre, alors il n'y a plus de raison de faire le bilan
        if len(instance._df_sessions_filtre) != 0 :
            # On met à jour l'Excel evalstat de la formation si des sessions demandées par l'utilisateur ne s'y trouvent pas
            instance._maj_evalstat_formation()

            # On calcule les stats
            instance._calculer_stats_criteres()
            #pprint(instance._stats_stagiaires)

            # On construit le bilan de session (bilan de session V3)
            instance._bilanSessionV3()

            # On ouvre le word
            FichierWord.depuisFichier(chemin_fichier=instance._chemin_word_bilan_session_output, charger_contentControl=False, afficherWord=True)

            # On envoie un mail au chef d'unité pour la signature du pdf
            instance._envoyer_mail_chef_unite()


        else:
            vlog.print("Info", f"⚠️ Toutes les sessions sont exclues : il n'y a plus de raison de faire le bilan de session.")

    @classmethod
    def plusieursBilans_parPeriode(cls, liste_periodes:list[Tuple[str, int, str]]) -> None:
        """
        Permet de lancer une série de bilans de sessions à partir d'une liste de tuples ex. [("948", 2022, "Année"), ("948", 2023, "1er semestre"), ("948", 2023, "2nd semestre")]
        """
        for codeFormation, annee, periode in liste_periodes:
            cls.bilanUnique_parPeriode(codeFormation, annee, periode)

    # === MÉTHODES ===
    @classmethod
    def _charger_fe_IRIS_sessions(cls) -> None:
        """
        Charge le fichier Excel IRIS Sessions si ce n'est pas déjà fait.
        Cette méthode met à jour _fe_IRIS_sessions et _df_sessions.
        """
        if cls._fe_IRIS_sessions is None:

            # Lecture de l'Excel
            cls._fe_IRIS_sessions = FichierExcel.depuis_fichier(repertoire_recherche_ini=config.IRIS_SESSIONS._output.repertoire)

            # Copie du tableau structuré "Sessions"
            cls._df_sessions = cls._fe_IRIS_sessions._tableaux["Sessions"]._df.copy()

            # Trie par "Date début ses."
            cls._df_sessions = cls._df_sessions.sort_values(by="Date début ses.")

            # Retype "Trigramme formation" et "Code IRIS"
            cls._df_sessions["Trigramme formation"] = cls._df_sessions["Trigramme formation"].astype(str)
            cls._df_sessions["Code IRIS"] = cls._df_sessions["Code IRIS"].astype(str)

    @staticmethod
    def _demander_entiers(message="Pour exclure des sessions : entrez un ou plusieurs code IRIS (numéro à 5 chiffres) séparés par des espaces (ou rien pour passer) : ") -> list[str]:
        while True:
            entree = input(message).strip()
            if not entree:
                # Pas de saisie => retourner liste vide
                return []

            # On met mes codes IRIS dans une liste de str
            #valeurs = [v.strip() for v in entree.replace(',', ' ').split()]

            # Vérifier que toutes les valeurs sont des entiers
            #try:
            #    entiers = [int(v) for v in valeurs]
            #    return entiers
            #except ValueError:
            #    print("Erreur : veuillez entrer uniquement des nombres entiers, séparés par des espaces ou des virgules.")

            # On met mes codes IRIS dans une liste de str en vrifiant que toutes les valeurs sont des entiers
            try:
                # On teste le typage en int
                l_entiers = [int(v.strip()) for v in entree.replace(',', ' ').split()]
                # On reconvertit en str avant sortie méthode
                l_str = [str(v) for v in l_entiers]
                return l_str

            except ValueError:
                print("Erreur : veuillez entrer uniquement des nombres entiers, séparés par des espaces ou des virgules.")
                BilanSession._demander_entiers()

    def _creer_df_extractIRIS_sessions_codeIRIS(self) -> None:
        """
        Méthode pour créer df_sessions_filtre
        C'est le dataframe issu de l'extract IRIS session. Il est filtré sur :
           - le trigramme de la formation en cours ;
           - l'année de la session ;
           - Statut Session != "Annulée" ;
           - la période souhaitée pour le bilan (1er semestre, 2nd semestre ou toute l'année)

        On se base sur l'export session de IRIS le plus récent
        """
        # On ouvre et lit l'export sessions IRIS si et seulement si il n'est pas déjà ouvert et lu avant
        self.__class__._charger_fe_IRIS_sessions()

        # Application du filtre sur la session
        self._df_sessions_filtre = self.__class__._df_sessions[self._df_sessions['Code IRIS'] == self._code_IRIS]
        #print(self._df_sessions_filtre)

        if len(self._df_sessions_filtre) < 1:
            print(self._df_sessions_filtre)
            vlog.log_erreur(f"Le fichier Excel Session ne contient pas ce code IRIS : {self._df_sessions_filtre}")

    def _creer_df_extractIRIS_sessions_periode(self) -> None:
        """
        Méthode pour créer df_sessions_filtre
        C'est le dataframe issu de l'extract IRIS session. Il est filtré sur :
           - le trigramme de la formation en cours ;
           - l'année de la session ;
           - Statut Session != "Annulée" ;
           - la période souhaitée pour le bilan (1er semestre, 2nd semestre ou toute l'année)

        On se base sur l'export session de IRIS le plus récent
        """
        # On ouvre et lit l'export sessions IRIS si et seulement si il n'est pas déjà ouvert et lu avant
        self.__class__._charger_fe_IRIS_sessions()


        ####
        # Filtration du dataframe
        ####

        # Application du pré-filtre avec les 3 critères trigramme, statut session et période
        self._df_sessions_filtre = self.__class__._df_sessions[
            (self._df_sessions['Trigramme formation'] == str(self._codeFormation)) &
            (self._df_sessions['Année début ses.'] == self._annee) &
            (self._df_sessions['Statut Session'] != "Annulée") &
            (self._df_sessions['Nb. Présents'] != 0)
        ]
        #print(self._df_sessions_filtre)


        # Définition date de début et de fin de la période choisie par l'utilisateur
        if self._periode == "1er semestre":
            date_debut = pd.Timestamp(f'{self._annee}-01-01')
            date_fin = pd.Timestamp(f'{self._annee}-06-30')
        elif self._periode == "2nd semestre":
            date_debut = pd.Timestamp(f'{self._annee}-07-01')
            date_fin = pd.Timestamp(f'{self._annee}-12-31')
        elif self._periode == "Année":
            date_debut = pd.Timestamp(f'{self._annee}-01-01')
            date_fin = pd.Timestamp(f'{self._annee}-12-31')
        else:
            # Cas par défaut : on ne filtre pas sur la date
            date_debut = None
            date_fin = None


        # Application second filtre sur la période
        if date_debut is not None and date_fin is not None:
            self._df_sessions_filtre = self._df_sessions_filtre[
                (self._df_sessions_filtre['Date début ses.'] >= date_debut) &
                (self._df_sessions_filtre['Date début ses.'] <= date_fin)
            ]   

    def _demande_sessions_a_exclure(self) -> None:
        """
        Demande à l'utilisateur les sessions qu'il souhaite exclure de la période choisie
        On retourne un dataframe df_sessions_filtre à jour
        """
        # On affiche à l'utilisateur les sessions et dates et statuts 
        vlog.print("Info", f"\nListe des sessions {self._codeFormation} dans {self._fe_IRIS_sessions._chemin_fichier.name} - {self._periode} {self._annee}", style=["jaune"])
        print(tabulate(
            self._df_sessions_filtre[['Code IRIS', 'Trigramme RP', 'Trigramme AF', 'Date début ses.', 'Date fin ses.', 'Durée réal. (J.)', 'Nb. Présents', 'Statut Session', 'N° Session']], 
            headers='keys', 
            tablefmt='pretty', 
            showindex=False
        ))
        
        # On demande à l'utilisateur les sessions qu'il veut exclure
        exclusionSessions = BilanSession._demander_entiers()
        if exclusionSessions:  # si la liste n'est pas vide
            # On trace l'exclusion des sessions
            for session_exclue in exclusionSessions:
                self._exploitationBilan["Exclus entièrement du bilan"].append(self._df_sessions_filtre.loc[self._df_sessions_filtre["Code IRIS"] == session_exclue, "N° Session"].iloc[0])

            # On met à jour _df_sessions_filtre en enlevant les sessions exclues
            self._df_sessions_filtre = self._df_sessions_filtre[~self._df_sessions_filtre['Code IRIS'].isin(exclusionSessions)]
        else:
            # la liste est vide, on ne filtre rien, on garde tout
            pass

        # On affiche à l'utilisateur les sessions finalement retenues
        vlog.print("Info", "\nSessions retenues pour le bilan :", style=["jaune"])
        print(tabulate(
            self._df_sessions_filtre[['Code IRIS', 'Trigramme RP', 'Trigramme AF', 'Date début ses.', 'Date fin ses.', 'Durée réal. (J.)', 'Nb. Présents', 'Statut Session', 'N° Session']], 
            headers='keys', 
            tablefmt='pretty', 
            showindex=False
        ))

        # On met à jour _df_sessions_filtre en enlevant les sessions exclues
        self._exploitationBilan["Exploités pour les stats générales"] = self._df_sessions_filtre["N° Session"].tolist()

    def _maj_evalstat_formation(self) -> None:
        """
        Met à jour l'Excel evalstat de la formation si des sessions demandées par l'utilisateur ne s'y trouvent pas
        (on regarde les CSV qui ne sont pas dans le fichier Excel global à partir de la liste df_sessions_filtre['Code IRIS'])

        """
        # On ouvre le fichier Excel global des évaluations de la formation
        timer.debut("Lecture du fichier Excel global des évaluations des stagiaires")
        self._es = EvalStat.ouvrir_ou_creer_evaluationsFormation(self._codeFormation)
        self._df_stagiaires = self._es._fe_evaluations_formation._tableaux["Stagiaires"]._df  # Création d'un alias pour faciliter le code
        
        # Retype "Trigramme formation" et "Code IRIS"
        self._df_stagiaires["Trigramme formation"] = self._df_stagiaires["Trigramme formation"].astype(str)
        self._df_stagiaires["Code IRIS"] = (
            pd.to_numeric(self._df_stagiaires["Code IRIS"], errors="coerce")  # "13414" ou 13414.0 → 13414
            .astype("Int64")                               # reste un entier (NaN compatible)
            .astype(str)                                   # enfin en texte propre
        )
        

        #print(f"\nÉtat de Excel évaluations filtré sur période et trigramme ({self._es._chemin_excel_evaluations_formation}) :")
        #pprint(self._df_stagiaires)

        vlog.ajouter_message("OK", f"✅ Lecture du fichier Excel global des évaluations des stagiaires {self._es._fe_evaluations_formation.chemin_fichier}")
        timer.fin()



        # --- Gestion des CSV manquants
        # On isole depuis ce fichier les CSV manquants (df_sessions_filtre = sessions demandées par l'utilisateur ; es._df_evaluations_formation = existant dans l'excel global)
        self._code_session_absents = list(set(self._df_sessions_filtre['Code IRIS']) - set(self._df_stagiaires['Code IRIS']))
        #print(set(self._df_sessions_filtre['Code IRIS']))
        #print(set(self._df_stagiaires['Code IRIS']))
        if self._code_session_absents:
            vlog.print("Info", f"🔎 Des codes session sont absents de l'Excel global des évaluations des stagiaires : {self._code_session_absents}")

            # Pour les CSV manquants, on demande à l'utilisateur de sélectionner les CSV à la main
            chemins_csv_a_traiter = {}
            chemins_csv_manquants = {}
            for code_IRIS in self._code_session_absents:
                # L'utilisateur sélectionne le CSV de code_IRIS
                chemin_csv_supp = self._es._filedialog_csv(code_IRIS=code_IRIS, trigramme_formation=self._codeFormation)
                
                if chemin_csv_supp is not None: # CSV sélectionné 
                    chemins_csv_a_traiter[code_IRIS] = chemin_vers_unc(chemin_csv_supp)
                    vlog.print("OK", f"✅ CSV {code_IRIS} : {chemin_csv_supp}")
                else: #CSV manquant
                    # On trace l'exclusion du CSV
                    self._exploitationBilan["Exclus des évaluations (CSV manquants)"].append(self._df_sessions_filtre.loc[self._df_sessions_filtre["Code IRIS"] == code_IRIS, "N° Session"].iloc[0])

                    chemins_csv_manquants[code_IRIS] = f"❌ Pas de CSV disponible pour cette session {code_IRIS}, session exclue."
                    vlog.print("CSV non disponible", f"❌ CSV {code_IRIS} : pas de CSV disponible pour cette session, session exclue.")
            #print(chemins_csv_manquants)


            # On traite les CSV sélectionnés par l'utilisateur
            tuple_csv_stagiaires = tuple(val for val in chemins_csv_a_traiter.values())
            if tuple_csv_stagiaires:
                vlog.print("Info", f"⏳ Traitement des CSV non déjà présents dans {os.path.basename(self._es._fe_evaluations_formation.chemin_fichier)} :" + "".join(f"\n• {chemin}" for chemin in tuple_csv_stagiaires))
                self._es = EvalStat.depuis_tuple_csv_stagiaires(tuple_csv_stagiaires=tuple_csv_stagiaires, fe_IRIS_sessions=self._fe_IRIS_sessions)
                self._df_stagiaires = self._es._fe_evaluations_formation._tableaux["Stagiaires"]._df.copy()  # MàJ de l'alias pour faciliter le code
                self._df_stagiaires["Trigramme formation"] = self._df_stagiaires["Trigramme formation"].astype(str)
                self._df_stagiaires["Code IRIS"] = (
                    pd.to_numeric(self._df_stagiaires["Code IRIS"], errors="coerce")  # "13414" ou 13414.0 → 13414
                    .astype("Int64")                               # reste un entier (NaN compatible)
                    .astype(str)                                   # enfin en texte propre
                )

            #print(self._df_stagiaires)

            # Après traitement des CSV on regarde dans self._df_stagiaires les N° de session
            #self._exploitationBilan["Exploités pour les évaluations (CSV présents)"] = list(set(self._df_sessions_filtre['Code IRIS']) & set(self._df_stagiaires['Code IRIS']))
            #codesIRIS_absents_apres_traitement = list(set(self._code_session_absents) - set(self._df_stagiaires['Code IRIS']) - set(self._exploitationBilan["Exclus des évaluations (CSV manquants)"]))
            #print("codesIRIS_absents_apres_traitement : ", codesIRIS_absents_apres_traitement)
            #print(set(self._code_session_absents))
            #print(set(self._df_stagiaires['Code IRIS']))

            #self._exploitationBilan["Exclus des évaluations (problème traitement CSV)"] = self._df_sessions_filtre[self._df_sessions_filtre['Code IRIS'].isin(codesIRIS_absents_apres_traitement)]['N° Session'].tolist()


            # Affichage selon retours traitement EvalStat
            if (self._es._chemins_csv_traites or self._es._chemins_csv_probleme or self._es._chemins_csv_exclus):
                vlog.print("Info", "\nBilan traitement des CSV :")
            if self._es._chemins_csv_traites:
                vlog.print("OK", f"   ✅ Chemins traités : " + "".join(f"\n      • {i_csv}" for i_csv in self._es._chemins_csv_traites))
            if self._es._chemins_csv_probleme:
                vlog.print("Problème traitement CSV", f"   ❌ Chemins ayant eu des problèmes (exclus) : " + "".join(f"\n      • {i_csv}" for i_csv in self._es._chemins_csv_probleme))
            if self._es._chemins_csv_exclus:
                vlog.print("Problème traitement CSV", f"   ❌ Chemins exclus lors du traitement : " + "".join(f"\n      • {i_csv}" for i_csv in self._es._chemins_csv_exclus))
                
        else:
            vlog.print("OK", f"✅ Tous les CSV sont bien déjà importés dans {self._es._fe_evaluations_formation.chemin_fichier.name}")



        self._codes_IRIS_communs = list(set(self._df_sessions_filtre['Code IRIS']) & set(self._df_stagiaires['Code IRIS']))
        #codes_IRIS_absents = list(set(self._df_sessions_filtre['Code IRIS']) - set(self._df_stagiaires['Code IRIS']))
        #vlog.print("Info", f"\nIn fine, voici la liste des éléments qui seront :")
        #vlog.print("Info", f"\t• inclus dans le bilan : {self._codes_IRIS_communs}")
        #vlog.print("Info", f"\t• exclus du bilan : {codes_IRIS_absents}")


 
        
        # On remet à jour les alias avec les nouvelles données (après traitement CSV)
        self._df_stagiaires_final = self._df_stagiaires[self._df_stagiaires['Code IRIS'].isin(self._codes_IRIS_communs)]
        #vlog.print("Info", self._df_stagiaires_final)
        self._df_stagiaires_final_1ligne_session = self._df_stagiaires_final.drop_duplicates(subset=['Code IRIS'])  # Ne garde qu'une ligne par Code IRIS (la première rencontrée)
        #vlog.print("Info", self._df_stagiaires_final_1ligne_session)

        # Dernière vérif qu'on a bien tout importé les CSV dans l'excel global
        self._codes_IRIS_absents_fin = list(set(self._df_sessions_filtre['Code IRIS']) - set(self._df_stagiaires_final['Code IRIS']))
        #if not self._codes_IRIS_absents_fin:
            #print(self._code_session_absents_fin)
            #vlog.print("Erreur", f"Erreur il reste encore des disparités avec des CSV non importés qui sont sensés être dans le bilan après traitement : {self._codes_IRIS_absents_fin}")
            #exit()

        # On définit la liste finale des N° session traités
        codesIRIS_avec_CSV = list(set(self._df_sessions_filtre['Code IRIS']) & set(self._df_stagiaires_final['Code IRIS']))
        self._exploitationBilan["Exploités pour les évaluations (CSV présents)"] = self._df_sessions_filtre[self._df_sessions_filtre['Code IRIS'].isin(codesIRIS_avec_CSV)]['N° Session'].tolist()
        self._exploitationBilan["Exclus des évaluations (problème traitement CSV)"] = list(
            set(self._exploitationBilan["Exploités pour les stats générales"])
            - set(self._exploitationBilan["Exploités pour les évaluations (CSV présents)"])
            - set(self._exploitationBilan["Exclus des évaluations (CSV manquants)"])
        )
        #self._exploitationBilan["Exploités pour les évaluations (CSV présents)"] = list(
        #    set(self._exploitationBilan["Exploités pour les stats générales"])
        #    - set(self._exploitationBilan["Exclus des évaluations (CSV manquants)"])
        #    - set(self._exploitationBilan["Exclus des évaluations (problème traitement CSV)"])
        #    )

        #pprint(self._exploitationBilan)
        
        self._commentairesBilan += "\nListe des sessions :"
        for critere, lsessions in self._exploitationBilan.items():
            if lsessions:
                self._commentairesBilan += f"\n   • {critere} :" + "".join(f"\n       - {isession}" for isession in lsessions)
        vlog.print("Info", f"\n{self._commentairesBilan}")        

    def _calculer_stats_criteres(self) -> dict:
        """
        Retourne un dictionnaire de la forme :
        {
            "Nom du critère": {
                "Nombre": ...,
                "Moyenne": ...,
                "Commentaires": ...
            },
            ...
        }
        """      
        self._stats_stagiaires = {}

        # S'il n'y a pas de CSV disponibles pour les stats, alors ce n'est pas la peine de faire les stats
        if len(self._exploitationBilan["Exploités pour les évaluations (CSV présents)"]) != 0 :
            liste_criteres = self._df_stagiaires_final['Critère'].dropna().unique()

            #for critere in liste_criteres:
            #    if critere in self._criteres_a_enlever:
            #        continue

            for critere in liste_criteres:
                df_filtre = self._df_stagiaires_final[self._df_stagiaires_final['Critère'] == critere]
                nb = len(df_filtre)
                moyenne = df_filtre['Note'].mean() if nb > 0 else None

                commentaires_concat = "\n".join(
                    "• " + c.strip()
                    for c in df_filtre['Commentaires'].dropna().astype(str)
                    if c.strip() != ""
                )

                self._stats_stagiaires[critere] = {
                    "Nombre": nb,
                    "Moyenne": moyenne,
                    "Commentaires": commentaires_concat
                }
        else:
            vlog.print("Info", f"⚠️ Aucun CSV disponible pour ce bilan : les statistiques des stagiaires ne seront pas évaluées.")

        return self._stats_stagiaires

    def _bilanSessionV3(self) -> None:
        # On construit champs de fusion
        self.__construit_champsFusionV3()

        # On merge les champs de fusion
        self.__mergeBilanV3()

    def __construit_champsFusionV3(self) -> None:
        """
        Calcule puis définit les str des champs de fusion
        """
        #####
        # On évalue les valeurs requises pour la fin de la méthode
        #####
        #self._periodeSessionsEvaluees = f"{self._periode} {self._annee}"  #Déjà évalué avant : dépend de si on fait une session unique ou une période
        self._chemin_word_bilan_session_output = config.format_path(config.CHEMIN_WORD_BILAN_SESSION_OUTPUT, trigramme_formation=self._codeFormation, annee=self._annee, periode=self._periodeSessionsEvaluees, unite=config.UNITE)
        
        #liste_numerosSession_statsgenerales_seulement = list(
        #    set(self._exploitationBilan["Exploités pour les stats générales"])
        #    - set(self._exploitationBilan["Exploités pour les évaluations (CSV présents)"])
        #)


        #####
        # Affectation des valeurs pour les champs de fusion (ce sont des str)
        #####
        self._titreFormation = self._df_sessions_filtre["Session"].iloc[-1]  # On est sensé travailler sur un même trigramme de formation, donc quelle que soit la ligne on a le bon nom de formation
        # self._codeFormation = codeFormation  # (donné en argument)
        # self._periodeSessionsEvaluees = f"{self._periode} {self._annee}"  # Evalué plus haut
        self._nbSessionsEvaluees = f"{len(self._df_sessions_filtre)}"  # Valeur toutes les données
        #self._numerosSessions = "\n".join(liste_numerosSession_statsgenerales_seulement)  # Valeur toutes les données
        self._numerosSessions = "\n".join(self._df_sessions_filtre["N° Session"].dropna().astype(str).unique())
        self._nbApprenants = f"{self._df_sessions_filtre['Nb. Nommés'].sum()}"  # Valeur toutes les données
        #self._rp = ", ".join(self._df_sessions_filtre["Trigramme RP"].dropna().astype(str).unique())  # Valeur toutes les données
        self._rp = ", ".join(self._df_sessions_filtre["Nom responsable pédag."].dropna().astype(str).unique() + " " + self._df_sessions_filtre["Prénom responsable pédag."].dropna().astype(str).unique())  # Valeur toutes les données
        #self._af = ", ".join(self._df_sessions_filtre["Trigramme AF"].dropna().astype(str).unique())  # Valeur toutes les données
        self._af = ", ".join(self._df_sessions_filtre["Créée par"].dropna().astype(str).unique())  # Valeur toutes les données


        # On n'affecte les champs suivants que si des CSV sont disponibles pour les stats
        if len(self._exploitationBilan["Exploités pour les évaluations (CSV présents)"]) != 0 :

            # On évalue les données requises pour les stats
            nb_stagiaires_retours = self._df_stagiaires_final['NOM Prénom'].nunique()
            nb_apprenants = self._df_stagiaires_final_1ligne_session['Nb présents'].sum()

            stats_sous_3 = { # Dictionnaire pour les critères dont la moyenne est inférieure à 3 et non exclus (critères dans la liste self._CRITERES_A_ENLEVER)
                critere: valeurs
                for critere, valeurs in self._stats_stagiaires.items()
                if (
                    critere not in self._CRITERES_A_ENLEVER
                    and valeurs["Moyenne"] is not None
                    and valeurs["Moyenne"] < 3
                )
            }

            # On gère les données entre parenthèses s'il y a des sessions exclues d'une manière ou d'une autre
            #if len(self._exploitationBilan["Exploités pour les stats générales"]) != len(self._exploitationBilan["Exploités pour les évaluations (CSV présents)"]) :
            #    self._nbSessionsEvaluees += f" ({len(self._codes_IRIS_communs)})"  # Valeur si on ne prend que les données CSV
            #    self._numerosSessions += "\n".join("(" + self._df_stagiaires_final_1ligne_session["N° Session"].dropna().astype(str).unique() + ")")  # Valeur si on ne prend que les données CSV
            #    self._nbApprenants += f" ({nb_apprenants:.0f})"  # Valeur si on ne prend que les données CSV
            #    #self._rp += " (" + ", ".join(self._df_stagiaires_final_1ligne_session["Trigramme RP"].dropna().astype(str).unique()) + ")"  # Valeur si on ne prend que les données CSV
            #    #self._af += " (" + ", ".join(self._df_stagiaires_final_1ligne_session["Trigramme AF"].dropna().astype(str).unique()) + ")"  # Valeur si on ne prend que les données CSV
            #
            #    self._commentairesBilan = "Les valeurs entre parenthèses dans les statistiques générales sont les données des sessions pour lesquelles nous avons des CSV exploitables.\n" + self._commentairesBilan
            
            # Si des champs ne sont pas dans le CSV, alors on garde "" qui est déjà définit dans le constructeur
            try:
                self._satisfactionGlobale_moy = f'{self._stats_stagiaires["Satisfaction globale"]["Moyenne"]:.1f}'
            except:
                pass
            try:
                self._satisfactionGlobale_com = self._stats_stagiaires["Satisfaction globale"]["Commentaires"].replace("_x000D_", "\n")
            except:
                pass
            try:
                self._recommandation_moy = f'{self._stats_stagiaires["Recommanderiez-vous cette formation ?"]["Moyenne"]/5*100:.0f}%'  # (on divise par 5 car on a un booléen stcké sous forme de note sur 5 : 0 = False, 5 = True)
            except:
                pass
            try:
                self._commentairesRemarquesSuggestions_com = self._stats_stagiaires["Commentaires, remarques, suggestions"]["Commentaires"].replace("_x000D_", "\n")
            except:
                pass
            try:
                self._evalInf3_val = f"{len(stats_sous_3)}"
            except:
                pass
            try:
                self._evalInf3_com = "\n".join(f"• {clef} ({valeurs['Moyenne']:.1f}) :{valeurs['Commentaires'].replace('•', '\n   -').replace('\n\n', '\n')}"
                    for clef, valeurs in stats_sous_3.items()
                )

                #self._evalInf3_com = "\n".join(
                #    valeurs["Commentaires"]
                #    for valeurs in stats_sous_3.values()
                #    if valeurs["Commentaires"]
                #).replace("_x000D_", "\n")
            except:
                pass
            try:
                self._tauxRetours_val = f"{(nb_stagiaires_retours/nb_apprenants)*100:.0f}%"
            except:
                pass

    def __mergeBilanV3(self) -> None:
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
        document = MailMerge(config.CHEMIN_MODELE_WORD_BILAN_SESSION)
        #print(document.get_merge_fields())

        document.merge(
            titreFormation = self._titreFormation,
            codeFormation = self._codeFormation,
            periodeSessionsEvaluees = self._periodeSessionsEvaluees,
            nbSessionsEvaluees = self._nbSessionsEvaluees,
            numerosSessions = self._numerosSessions,
            nbApprenants = self._nbApprenants,
            rp = self._rp,
            af = self._af,
            commentairesBilan = self._commentairesBilan,
            satisfactionGlobale_moy = self._satisfactionGlobale_moy,
            satisfactionGlobale_com = self._satisfactionGlobale_com,
            recommandation_moy = self._recommandation_moy,
            commentairesRemarquesSuggestions_com = self._commentairesRemarquesSuggestions_com,
            evalInf3_val = self._evalInf3_val,
            evalInf3_com = self._evalInf3_com,
            tauxRetours_val = self._tauxRetours_val
            )
        
        # On crée le répertoire pour les bilans de session de cette année s'il n'existe pas
        self._chemin_word_bilan_session_output.parent.mkdir(parents=True, exist_ok=True)

        # On écrit le fichier
        document.write(self._chemin_word_bilan_session_output)

    def _envoyer_mail_chef_unite(self, pj:Optional[list[str]] = None):
        """
        Envoie un mail au chef d'unité avec en lien le PDF à signer
        """       
        


        chemin_pdf_bilan_output = self._chemin_word_bilan_session_output.with_suffix(".pdf")
        #self._CORPS_MAIL_CHEF_UNITE.replace()
        corps_html = remplacer_champs(config.CORPS_MAIL_CHEF_UNITE, [
            ["lien_pdf_bilan", chemin_pdf_bilan_output],
            ["formation", f"{self._titreFormation} ({self._codeFormation})"],
            ["periode", minuscule_premiere_lettre(self._periode)],
        ])

        Mail.creer_mail(
            destinataires=config.ADRESSE_MAIL_CHEF_UNITE,
            sujet=f"Signature bilan de session {self._titreFormation} ({self._codeFormation}) : {chemin_pdf_bilan_output.name}",
            corps_html=corps_html,
            pieces_jointes=pj,
            envoyer_mail=False  # envoie directement sans afficher
        ) 

class BilanFormation:
    """
    C'est la classe qui contient tous les éléments de ma formation pour mon bilan

    # TODO j'en suis là
    # Todo Word
    # Dans le modèle Word : gérer le lien vers la GED 
    # Exploiter EvalStat
    # Il y a des trous dans la raquette dans le word de sortie (checkboxes)
    # coller des images depuis Excel
    # ? Exploiter export formation plutôt que export sessions pour les valeurs par défaut nmin/max...
    """
    def __init__(self, codeFormation:str, annee:int):
        
        self._codeFormation:str = codeFormation
        self._annee:int = annee

        self._sessions_nom_typeExport:str = config.IRIS_SESSIONS._nom_typeExport  #sessions._nom_typeExport #  Provient de la valeur globale sessions
        self._sessions_codeExport:str = config.IRIS_SESSIONS._codeExport #sessions._codeExport #  Provient de la valeur globale sessions
        self._sessions_repertoire:Path = config.IRIS_SESSIONS._output.repertoire #sessions._output.repertoire #  Provient de la valeur globale sessions

        self._chemin_modele_word_bilan_formation = config.CHEMIN_MODELE_WORD_BILAN_FORMATION
        self._chemin_word_bilan_formation_output = config.format_path(config.CHEMIN_WORD_BILAN_FORMATION_OUTPUT, trigramme_formation=codeFormation, annee=annee)


        self._chemin_specsPedagogiques = None

        self._chemin_fdc = None

        self._repertoire_fdc_defaut:Path = config.format_path(config.REPERTOIRE_FDC, trigramme_formation=codeFormation)
        self._repertoire_specsPedagogiques_defaut:Path = config.format_path(config.REPERTOIRE_SPECS, trigramme_formation=codeFormation)
        
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
        self._dateSpecs = datetime.fromtimestamp(self._chemin_specsPedagogiques.stat().st_mtime)
        self._sDateSpecs = self._dateSpecs.strftime("%d/%m/%Y")
        #print(self._sDateSpecs)



        #####
        # Exploitation de la fiche de coûts
        #####
        self._chemin_fdc = filedialog.askopenfilename(title="Sélectionner la dernière fiche de coûts", filetype=[("Fichiers Excel", "*.xlsx")], initialdir=optimiseCheminRepertoire(self._repertoire_fdc_defaut.replace("XXX", self._codeFormation)))
        if not chemin_fichier_session:
            log_erreur("click sur cancel du filedialog → Pas de chemin de fiche de coûts")
        self._dateFdC = datetime.fromtimestamp(self._chemin_fdc.stat().st_mtime)
        self._sDateFdC = self._dateFdC.strftime("%d/%m/%Y")
        
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
    def mergeBilan(self):
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
        document = MailMerge(self._chemin_modele_word_bilan_formation)
        #print(document.get_merge_fields())

        document.merge(
            annee='{:%Y}'.format(date.today()),
            codeFormation=self._codeFormation,
            titreFormation=self._titreFormation,
            
            lienGED = config.format_path(config.REPERTOIRE_FORMATION, trigramme_formation=self._codeFormation),
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
        
        document.write(self._chemin_word_bilan_formation_output)



@dataclass
class TypeIntervenant:
    nom:str
    docs:list[REE.DocREE]

class REE:
    
    @dataclass
    class DocREE:
        frequence_maj:Optional[list[str]] = None
        nom_fichier:Optional[str] = None
        chemin_fichier:Optional[str] = None
        intervenants:Optional[list[str]] = None

    """
    def enregistrer_docRee_dico(nom: str, frequence_maj:Optional[list[str]], nom_fichier:Optional[str] = None) -> DocREE:
        doc = DocREE(nom, frequence_maj, nom_fichier)
        REE._docsREE[nom] = doc
        return doc


    def enregistrer_intervenant_dico(cls, nom: str, docs: list[REE.DocREE]) -> TypeIntervenant:
        ti = cls.TypeIntervenant(nom, docs)
        cls._typesIntervenants[nom] = ti
        return ti"""
    # === VARIABLES DE CLASSE ===
    # --- Paramètres d'environnement
    _repertoire_documents_ree:str = r"\\harmonie\instn\uem\_Documents_communs\Formations\Formateurs\0.Docs à envoyer"  # Répertoire de la GED où sont 
    _chemin_mailtype_informationsAdministratives = r"\\harmonie\instn\uem\_Documents_communs\Formations\Formateurs\Mails types\Demande des informations administratives.msg"  # Message type à envoyer aux intervenants
    _repertoire_sauvegarde_fichiersREE:str = r"\\harmonie\INSTN\UEM\_Documents_communs\Formations\Formateurs\1.Intervenants - Documents administratifs" # Lieu où sauvegarder les fichiers de l'intervenant
    _chemin_modele_excel_ficheIntervenant:str = r"\\harmonie\INSTN\UEM\_Documents_communs\Formations\Formateurs\P09-Pr01-Qualifier les ressources enseignantes\P09_Pr01_Ta.E_Grille des critères de qualification des compétences_V1.xlsx"  # Fichier Excel à remplir pour Laetitia Da Mota (RH INSTN qui s'occupe de rentrer les REE dans IRIS)
    _adresse_mail_gestionnaire_ree_INSTN:str = "vacataires.instn@cea.fr"
    _corps_html_mail_gestionnaire_ree_INSTN:str = "<p>Bonjour Laëtitia,</p><p>Je t’ai mis en PJ les documents pour intégrer/mettre à jour la fiche IRIS de ###.</p><p>Je te remercie, passe une excellente journée,</p>"
    
    # Association colonnes excel avec command control Word
    # La préparation de ce ditionnaire peut être faite avec : ree._generer_dictionnaire_depuis_excel()
    _dict_colExcel_cc:Dict[str, str] = {
        "NOM": "Nom",
        "Pr\u00e9nom": "Prenoms",
        "Dipl\u00f4me ou formation/exp\u00e9rience professionnelle": "Diplome",
        "Dur\u00e9e exp\u00e9rience professionnelle": "DureeExperiencePro",
        "Niveau d'expertise permettant une reconnaissance": "NiveauExpertise",
        "Domaine / Sp\u00e9cialit\u00e9 de l'expertise": "DomaineExpertise",
        "ATTRIBUTION Niveau comp\u00e9tences techniques": None,
        "Combien de jours anim\u00e9s, en moyenne par an": "formation_nbJoursAnimes",
        "Combien de jours de formations suivies en p\u00e9dagogie (=animation)": "formation_nbJoursFormationPedagogie",
        "Profils d'apprenants form\u00e9s": "formation_profilApprenants",
        "Taux consolid\u00e9 de la satisfaction des apprenants relativement \u00e0 l'enseignant-formateur consid\u00e9r\u00e9": None,
        "Estimation par le RP de la capacit\u00e9 de l'enseignant-formateur \u00e0 animer (fond de salle)": None,
        "Outils num\u00e9riques utilis\u00e9s durant les animations r\u00e9alis\u00e9es (serious game, blended-learning\u2026)": "formation_outilsNumeriques",
        "Combien de jours pass\u00e9s en conception de s\u00e9quence de formation, en moyenne par an": "IngPedago_nbJoursConception",
        "Combien de jours de formations suivies en ing\u00e9nierie p\u00e9dagogique (=conception de s\u00e9quences de formation)": "IngPedago_nbJoursFormationIngPedago",
        "Estimation par le RP de la conception de la s\u00e9quence en fonction des objectifs p\u00e9dagogiques fournis par le RP (fond de salle, analyse des supports fournis)": None,
        "Estimation par le RP de la pertinence de l'\u00e9valuation des acquis r\u00e9alis\u00e9e par l'enseignant-formateur sur sa s\u00e9quence (analyse de la progression des apprenants : tests avant/apr\u00e8s)": None,
        "Estimation par le RP de l'utilisation des m\u00e9thodes actives (\u00e9tudes de cas, r\u00e9solution de probl\u00e8mes, classes invers\u00e9es, travaux de groupes\u2026)": None,
        "Combien d'ann\u00e9es d'exp\u00e9rience en conception de dispositifs de formations (=cr\u00e9ation et coordination)": "IngFormation_nbJoursConception",
        "Combien de jours de formations suivies en ing\u00e9nierie de formation (=conception de dispositifs de formation)": "IngFormation_nbJoursFormationIngFormation",
        "Estimation par le chef de projet ou le CUE de la complexit\u00e9 des pr\u00e9c\u00e9dents dispositifs de formation con\u00e7us": None,
        "Profil des apprenants des dispositifs de formations prc\u00e9demment con\u00e7us": "IngFormation_profilApprenants",
        "Estimation par le chef de projet ou le CUE de l'\u00e9valuation des acquis r\u00e9alis\u00e9 dans le dispositif de formation (mesure de la progression des apprenants=estimation de la qualit\u00e9 du dispositif de formation)": None,        
        "Combien d'ann\u00e9es d'exp\u00e9rience en tant que tuteur acad\u00e9mique": "IngFormation_nbAnneesTuteur",
        "Combien de r\u00e9f\u00e9rentiels d'activit\u00e9, de comp\u00e9tence et d'\u00e9valuation r\u00e9alis\u00e9s": "IngCompetences_nbReferentiels",
        "Combien de jours de formations suivies en ing\u00e9nierie de comp\u00e9tences": "IngCompetences_nbJoursFormationIngCompetences",
        "Estimation par la cellule p\u00e9dagogique de DPF de la complexit\u00e9 des pr\u00e9c\u00e9dentes r\u00e9alisations de l'ing\u00e9nieur/consultant en ing\u00e9nierie de comp\u00e9tences (complexit\u00e9 du m\u00e9tier et de son environnement : risques, r\u00e9glementation...)": None,
        "ATTRIBUTION Niveau comp\u00e9tences p\u00e9dagogiques": None,
        "Evaluation CECRL ou \u00e9quivalence TOEIC, TOEFL": "ResultatLangue2",
        "ATTRIBUTION Niveau comp\u00e9tences linguistiques": None,
        "Curriculum vitae": None
    }

    # --- Paramètres utilisateur
    # Documents à envoyer / demander
    _docsREE:dict[DocREE] ={
        "Fiche administrative" : DocREE(
            nom_fichier=r"Fiche administrative vacataire INSTN.docx", 
            frequence_maj=["Initialisation", "Mise à jour"],
            intervenants=["CEA", "vacataire", "contrat spécifique de collaboration", "auto-entrepreneur"]
        ),
        "CV" : DocREE(
            nom_fichier=r"CV-Type.docx", 
            frequence_maj=["Initialisation", "Mise à jour"],
            intervenants=["CEA", "vacataire", "contrat spécifique de collaboration", "auto-entrepreneur"]
        ),
        "RIB" : DocREE( 
            frequence_maj=["Initialisation", "Mise à jour"],
            intervenants=["vacataire", "auto-entrepreneur"]
        ),
        "Attestation employeur" : DocREE(
            nom_fichier=r"Attestation employeur.docx", 
            frequence_maj=["Initialisation", "Tous les ans"],
            intervenants=["vacataire", "contrat spécifique de collaboration"]
        ),
        "Devis" : DocREE(
            frequence_maj=["Initialisation", "Tous les ans"],
            intervenants=["auto-entrepreneur"]
        ),
        "Bilan pédagogique et financier année n-1" : DocREE(
            frequence_maj=["Initialisation", "Tous les ans"],
            intervenants=["auto-entrepreneur"]
        ),
        "Guide pour l'intervenant" : DocREE(
            nom_fichier=r"Guide pour l'intervenant.pdf",
            intervenants=["CEA", "vacataire", "contrat spécifique de collaboration", "auto-entrepreneur"]
        ),

    }

    _correspondance_frequence_texte:dict[str] = {
        "Initialisation" : "initialisation du dossier",
        "Mise à jour" : "s'il y a une mise à jour",
        "Tous les ans" : "<strong><u>chaque année civile</u></strong>"
    }
 


    # --- Autres variables de la classe
    #_docsREE:dict[DocREE] = {}  # Dictionnaire des documents (fiche admin, CV...)
    #_typesIntervenants:dict[TypeIntervenant] = {}  # Dictionnaire des types d'intervenant (CEA, vacataire...)


    # === CONSTRUCTEUR ===
    def __init__(self):
        """
        for doc in self._docsREE.values():
            if doc.nom_fichier is not None:
                # accès direct à la variable de classe REE._repertoire_documents_ree
                doc.chemin_fichier = os.path.join(self._repertoire_documents_ree, doc.nom_fichier)"""
        pass



    # === ENVOI MAIL REE ===
    def envoyerMail_REE(self, 
        statut:str, 
        destinataire:Optional[Union[str, list[str], pd.Series]] = None, 
        copie:Optional[Union[str, list[str], pd.Series]] = None
    ) -> None:
        # TODO : comment récupérer le statu depuis Excel ?
        # TODO : comment faire une boucle auto sur les personnes à qui envoyer ?
        """
        Prépare et envoie automatiquement un e-mail d'informations administratives
        via Outlook à partir d'un modèle pré-défini (.msg), selon le statut de l'intervenant.

        Cette méthode :
        - récupère la liste des documents requis pour un type d'intervenant (ex. Vacataire, CEA, etc.)
        - construit dynamiquement le corps du message (HTML) listant les documents à fournir
        - ajoute automatiquement les pièces jointes correspondantes
        - ouvre un e-mail Outlook basé sur le modèle type, avec remplacement des balises
            ###statut### et ###listeInfos### dans le corps du message.

        Args:
            statut (str):
                Statut de l'intervenant (ex. `"Vacataire"`, `"CEA"`, `"Auto-entrepreneur"`, ...).
                Sert à filtrer les documents associés à ce type d'intervenant.
            destinataire (Optional[Union[str, list[str], pd.Series]], optional):
                Adresse(s) e-mail des destinataires principaux.
                Peut être une chaîne unique, une liste de chaînes ou une série Pandas.
                Exemple : `"nom.prenom@domaine.com"` ou `["a@x.com", "b@y.com"]`.
            copie (Optional[Union[str, list[str], pd.Series]], optional):
                Adresse(s) e-mail des destinataires en copie (CC).

        Returns:
            None

        Raises:
            FileNotFoundError:
                Si le modèle de mail (_chemin_mailtype_informationsAdministratives) est introuvable.
            Exception:
                Toute erreur lors de la création ou de l'ouverture du mail Outlook.

        Example:
            >>> ree = REE()
            >>> ree.envoyerMail_REE(
            ...     statut="Vacataire",
            ...     destinataire="vacataire@exemple.com",
            ...     copie=["admin@instn.fr", "drh@instn.fr"]
            ... )
            # Ouvre un mail Outlook basé sur le modèle "Demande des informations administratives"
            # avec la liste des documents à fournir par le vacataire et les fichiers joints.
        """


        # Faire choix auto pour existant ou nouveau

        #Je crée le texte pour listeInfo

        
        listeInfos:str = ""
        listePJ:list[str] = []
        #self._correspondance_frequence_texte
        for nom_doc, doc in self._docsREE.items():
            if statut in doc.intervenants:
                # Gestion de la liste des infos à afficher dans le mail
                if doc.frequence_maj is not None:
                    # On mappe chaque élément de doc.frequence_maj via le dictionnaire
                    frequences = [
                        self._correspondance_frequence_texte.get(freq, freq)
                        for freq in doc.frequence_maj or []
                    ]

                    listeInfos += f"<li><strong>{nom_doc}</strong> <em>({', '.join(frequences)})</em></li>\n"  

                # Gestion des PJ à mettre dans le mail
                if doc.nom_fichier is not None:
                    listePJ.append(os.path.join(self._repertoire_documents_ree, doc.nom_fichier))            
        
        if listeInfos != "":
            listeInfos = "<ul> "+listeInfos+" </ul>\n"

        #Remplacer les textes avec statut et listeInfos

        Mail.depuis_modele(
            chemin_modele=self._chemin_mailtype_informationsAdministratives,
            destinataires=destinataire,
            copies=copie,
            pieces_jointes=listePJ,
            remplaceBalises=[["###statut###", statut], ["###listeInfos###", listeInfos]]
        )



    # === RECEPTION / TRAITEMENT DOC REE
    def traiter_docs_REE(self) -> None:

        # TODO : mettre à jour le fichier Excel des coordonnées des intervenants
        # TODO : mettre à jour le fichier Excel des AI

        # On ouvre le word et on charge tous les command control (filedialog depuis "Download"). On le ferme
        self._word_ficheAdministrative = FichierWord.depuisFichier()
        #print(self._word_ficheAdministrative)

        # TODO : Peut-être afficher NOM et prénoms pour que l'utilisateur redéfinisse quel nom et quel prénom écrire (peut-être enlever les prénos en sus)

        # On crée le répertoire dans le répertoire des REE s'il n'existe pas (ou assimilé) (NOM Prénom (Société - AAAA))
        self._creer_repertoire_REE()
        
        # L'utilisateur sélectionne tous les fichiers de la REE et on les déplace dans le répertoire idoine
        fichiers_sortie = self._deplacer_fichiers(self._repertoire_sauvegarde_fichiersREE)

        # On emplit le fichier Excel à transférer à Laetitia Da Mota à partir d'un modèle
        self._remplit_excel_avecInfos_word()


        # On ouvre l'Excel et le Word pour comparaison et adaptations manuelles
        chemin_word = os.path.join(self._repertoire_sauvegarde_fichiersREE, os.path.basename(self._word_ficheAdministrative._chemin_fichier))
        chemin_excel = os.path.join(self._repertoire_sauvegarde_fichiersREE, os.path.basename(self._chemin_modele_excel_ficheIntervenant))
        fichiers_sortie.append(chemin_excel)
        ouvrir_word_excel_cote_a_cote(chemin_word, chemin_excel, split_ecranPrincipal=True)  # on peut rajouter split_ecranPrincipal=True

        # On attend que l'utilisateur ait adapté/validé le fichier Excel pour avancer
        input("🕒 Attente pour adaptations de l'Excel.\nAppuyez sur une touche après adaptation/sauvegarde de l'Excel REE pour continuer")

        # Dès que l'Excel est fermé, on prépare le mail pour Laetitia
        self._envoyer_mail_gestionnaire_ree_instn(pj=fichiers_sortie)  # On peut aussi mettre delai=timedelta(days=30)

        # On met à jour le fichier Excel Liste AI formateurs.xlsx : onglet intervenant, on cherche et remplace la date de validité de l'attestation employeur sinon nouvelle ligne (recopier formule + format)
        # On met à jour le fichier Excel  avec la liste des intervenants :  on cherche et remplace les données mail, tel, Ville, la date de validité de l'attestation employeur... sinon nouvelle ligne (recopier formule + format)        



    # === Méthodes internes à la classe
    def _creer_repertoire_REE(self, test:bool=False) -> None:
        """
        Crée le répertoire du REE sur le réseau local
        """
        # TODO : comment faire si pas de content control ?
        if all(k in self._word_ficheAdministrative.cc for k in ["Nom", "Prenoms", "RaisonSociale"]):
            self._repertoire_sauvegarde_fichiersREE += f"\\{self._word_ficheAdministrative.cc['Nom'].upper()} {self._word_ficheAdministrative.cc['Prenoms'].title()} ({self._word_ficheAdministrative.cc['RaisonSociale'] if self._word_ficheAdministrative.cc['RaisonSociale'] != 'Raison sociale employeur principal' else 'CEA'} - {datetime.now().year})"
        else :
            print("⚠️ Certaines clés sont manquantes dans cc :", [k for k in ["Nom", "Prenoms", "RaisonSociale"] if k not in self._word_ficheAdministrative.cc])
        #print(self._repertoire_sauvegarde_fichiersREE)
        if not test:
            os.makedirs(self._repertoire_sauvegarde_fichiersREE, exist_ok=True)
            # TODO : self._repertoire_sauvegarde_fichiersREE.mkdir(parents=True, exist_ok=True)
        else :
            print(self._repertoire_sauvegarde_fichiersREE)

    def _deplacer_fichiers(self, destination: str = None) -> list[str]:
        """
        Ouvre un dialogue pour sélectionner des fichiers, puis les déplace vers un dossier choisi.

        Args:
            destination (str, optional): Chemin du dossier de destination.
                                        Si None, un dialogue s'ouvrira pour le choisir.
        """

        fichiers_sortie:list[str] = []

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
                fichiers_sortie.append(chemin_destination)
                print(f"✅ Déplacé : {nom_fichier}")
            except Exception as e:
                print(f"❌ Erreur avec {nom_fichier} : {e}")

        
        #print(fichiers_sortie)
        return fichiers_sortie

    def _remplit_excel_avecInfos_word(self) -> None:
        """
        Remplit le fichier Excel à transférer à Laetitia Da Mota à partir d'un modèle

        Doit avoir lu un word avec les content control en amont
        """
        # On ouvre le fichier Excel à remplir pour Laetitia Da Mota (c'est un modèle, on l'enregistrera avec le bon nom dans le répertoire idoine)
        excel_ficheIntervenant = FichierExcel.depuis_modele(
            chemin_modele = self._chemin_modele_excel_ficheIntervenant,
            chemin_fichier_sauv = os.path.join(self._repertoire_sauvegarde_fichiersREE, os.path.basename(self._chemin_modele_excel_ficheIntervenant)),
            charger_df = True
        )
        
        #df_REE = excel_ficheIntervenant._tableaux["QualificationsREE"]._df  # Alias
        #print(df_REE)


        # On pré-rempli le fichier Excel fiche intervenant grâce aux contecnt control du word et au dictionnaire
        nouvelle_ligne = {}
        for col_df, cc_key in self._dict_colExcel_cc.items():
            if cc_key is None:
                # Pas de clé correspondante => valeur vide dans la DataFrame
                nouvelle_ligne[col_df] = None
            else:
                # Récupérer la valeur dans le dictionnaire Word, ou None si la clé absente
                valeur = self._word_ficheAdministrative._cc.get(cc_key, None)
                nouvelle_ligne[col_df] = convertir_si_possible(valeur)
                #print(col_df, cc_key, valeur, type(convertir_si_possible(valeur)))

        #print(pd.DataFrame([nouvelle_ligne]))

        # On écrit le dataframe dans le tableau structuré
        excel_ficheIntervenant._tableaux["QualificationsREE"].ecrit_dataFrame_dans_tableauStructure(pd.DataFrame([nouvelle_ligne]), supprimeDonneesEtRemplace=True, remplace_df_par_nouveau=True, copie_formules=True)


        # On sauve la fiche intervenant
        excel_ficheIntervenant.save()
        excel_ficheIntervenant.close()

    def _envoyer_mail_gestionnaire_ree_instn(self, pj:Optional[list[str]] = None, delai: Optional[timedelta] = None):
        """
        Envoie un mail au gestionnaire des REE de l'INSTN (Laëtitia Da Mota)

        Pour les PJ, si elles ne sont pas données en argument, alors on récupère automatiquement tous les documents qui ont été mis dans le répertoire de l'intervenant il y a moins de 'delai'
        """

        if (pj is None) and (delai is not None):
            pj = lister_fichiers_repertoire(
                self._repertoire_sauvegarde_fichiersREE,
                delai=delai,
                inclure_sous_dossiers=True,        # Inclut les sous-dossiers
            )
        

        Mail.creer_mail(
            destinataires=self._adresse_mail_gestionnaire_ree_INSTN,
            sujet="Documents pour mise à jour IRIS", # Pimper avec le nom de l'intervenant
            corps_html=self._corps_html_mail_gestionnaire_ree_INSTN.replace("###", self._word_ficheAdministrative.cc['Prenoms'].title()+" "+self._word_ficheAdministrative.cc['Nom'].upper()),
            pieces_jointes=pj,
            envoyer_mail=False  # envoie directement sans afficher
        )  

    #  Pour aider à générer le dictionnaire des noms de colonne du fichier Excel REE
    def _generer_dictionnaire_depuis_excel(
        self,
        chemin_fichier: str=None,
        nomOngletQualifications: str=None,
        nomOngletAssociationCC: str=None,
        nomColonneCC: str=None,
    ) -> None:
        """
        Génère et affiche un dictionnaire Python liant les colonnes Excel de l'onglet 'nomOngletQualifications'
        aux content controls Word listés dans l'onglet 'nomOngletAssociationCC'.
        
        Si un content control est vide, sa valeur sera 'None'.

        Args:
            chemin_fichier (str): Chemin du fichier Excel.
            nomOngletQualifications (str): Nom de l’onglet contenant les colonnes de qualifications.
            nomOngletAssociationCC (str): Nom de l’onglet contenant les correspondances CC.
            nomColonneCC (str): Nom de la colonne dans l’onglet d’association qui contient les noms des content controls.
        """
        import json

        if chemin_fichier is None :
            chemin_fichier = self._chemin_modele_excel_ficheIntervenant
        if nomOngletQualifications is None :
            nomOngletQualifications = "QualificationsREE"
        if nomOngletAssociationCC is None :
            nomOngletAssociationCC = "AssociationCC"
        if nomColonneCC is None :
            nomColonneCC = "Nom CC"

        # Charger le fichier Excel via ta classe personnalisée
        fe = FichierExcel.depuis_fichier(chemin_fichier=chemin_fichier, charger_df=True)

        # Récupérer le DataFrame des colonnes de qualifications
        df_qualif = fe._tableaux[nomOngletQualifications].df
        noms_colonnes = list(df_qualif.columns)

        # Récupérer le DataFrame contenant les associations
        df_assoc = fe._tableaux[nomOngletAssociationCC].df

        if nomColonneCC not in df_assoc.columns:
            raise ValueError(f"La colonne '{nomColonneCC}' n'existe pas dans l'onglet '{nomOngletAssociationCC}'.")

        # Récupérer la colonne des content controls, en forçant la taille à celle des colonnes
        liste_cc = df_assoc[nomColonneCC].tolist()
        # Compléter si jamais il y a moins de lignes que de colonnes dans le premier onglet
        while len(liste_cc) < len(noms_colonnes):
            liste_cc.append(None)
            raise ValueError(f"Le nombre de lignes de CC '{len(liste_cc)}' est plus petit que le nombre de colonnes du tableau principal '{len(noms_colonnes)}'.")

        # Générer le dictionnaire
        print("mon_dictionnaire = {")
        for i, (nom_col, cc) in enumerate(zip(noms_colonnes, liste_cc)):
            virgule = "," if i < len(noms_colonnes) - 1 else ""
            # Si le CC est vide ou NaN → None
            if cc is None or (isinstance(cc, float) and pd.isna(cc)) or str(cc).strip() == "":
                cc_str = "None"
            else:
                cc_str = json.dumps(str(cc))
            print(f"    {json.dumps(nom_col)}: {cc_str}{virgule}")
        print("}")

class Traiter_contactsApprentis:
    """
    Classe permettant de contacter les apprentis et tuteurs lors d'un suivi d'apprentissage :
       - mail de premier contact (à partir d'un modèle .msg) ;
       - préparer les RDV Outlook pour les entretiens (à partir de modèles .oft) ;
       - faire les mails de relance pour des documents ou Studea.
    
    Pour avoir accès aux mails et aux noms, j'ouvre le fichier Excel Etudiants (ADIN ou LP3D)
    Dans ce fichier, on note aussi la réception des docs ou des signatures Studea pour filtrer les mails à envoyer

    TODO :
       - Pouvoir faire un RDV Skpe (lié à classe RDV_outlook)
       - Ouvrir un mail vide, on le remplit, on sauve, et ça envoie à tous les apprentis et tuteurs

    Validation :
       - mail de premier contact (à partir d'un modèle .msg) → Oui mais TODO signature à enlever
       - préparer les RDV Outlook pour les entretiens (à partir de modèles .oft) ;
       - faire les mails de relance pour des documents ou Studea.

    """
    
    @dataclass
    class PropEntretien:
        sujet: str
        chemin_modele: str
        duree: timedelta
        date_debut:datetime
        categorie: str = "FI"
        
        def __post_init__(self):
            # Normaliser en datetime
            if isinstance(self.date_debut, date) and not isinstance(self.date_debut, datetime):
                self.date_debut = datetime.combine(self.date_debut, time(9, 0))

    @dataclass
    class PropFichierARenvoyer:
        sujet: str
        chemin_fichier: str
        periode:str
        deadline_retour:Optional[datetime]=None

    # Fichier Excel avec les informations des étudiants
    _fe_etudiants:FichierExcel
    _df_etudiants:DataFrame

    # Année en cours
    _annee_scolaire:str

    # Infos mails
    _prefixe_sujet:str
    _chemin_modele_mail_priseContact:Optional[str] = None
    _mail_responsables_univ:Optional[str] = None
    
    _entretiens:list[PropEntretien] = []
    _relances:list[str] = []
    _fichiersARenvoyer:list[PropFichierARenvoyer] = []

    _envoyer_mail:bool = False

    # === Initialisations statiques ===
    # Colonnes du fichier etudiant à récupérer
    _colonnes_fe_etudiants = []


    def __init__(self, chemin_fichier_etudiants:str, nom_onglet:str) -> None:

        # Fichier Excel avec les informations des étudiants
        self._chemin_fichier_etudiants = chemin_fichier_etudiants

        # Par défaut, il faut charger l'Excel des coordonnées des étudiants
        self._fe_etudiants = FichierExcel.depuis_fichier(
            chemin_fichier = chemin_fichier_etudiants,
            nom_onglet = nom_onglet
        )

    @classmethod
    def UGA(cls) -> Traiter_contactsApprentis:
        """
        """

        # === Initialisations statiques ===
        annee_scolaire = "2025-2026"
        envoyer_mail = False

        prefixe_sujet = "Master IN - Suivi d'alternance"
        mail_responsables_univ = "master-in-responsables@univ-grenoble-alpes.fr"

        chemin_modele_mail_priseContact = r"C:\Users\vt238770\Documents\_CEA\Modèles adaptés\Mails\Master IN\Master IN - Tutorat en entreprise - Prise de contact.msg"

        chemin_fichier_etudiants =  fr"\\instnt\partage\FORMATIONS_I\GDRA - Master IN (parcours ADIN-GDRA-SN)\{annee_scolaire}\1-dossier etudiants\Master ADIN - {annee_scolaire.replace('-', '_')}.xlsx"
        
        nom_onglet = "Etudiants"
        
        colonnes_fe_etudiants = [
            "Cursus", 
            "Nom", "Prénom", "Mail apprenti", "Téléphone apprenti", 
            "Entreprise ", "Lieu entreprise", "Nom TE", "Prénom TE", "Mail TE", "Téléphone TE", "Ma fonction de suivi de l'alternant", 
            "Engagement des parties", "Entretien de prise de fonction", "1ère visite en entreprise", "2ème visite en entreprise", "Fiche évaluation 1", "Fiche évaluation 2"]

        # Entretiens
        prise_de_fonction = cls.PropEntretien(
            sujet = "Entretien de prise de fonction",
            chemin_modele = r"C:\Users\vt238770\Documents\_CEA\Modèles adaptés\Mails\Master IN\Master IN - Suivi d'alternance - Entretien.oft",
            duree = timedelta(hours=0, minutes=45),
            date_debut = RDV_Outlook.get_lundi_depuis_num_semaine(numero_semaine=41) # (Autour du 6 octobre : dernière semaine de la première période en entreprise → A faire avant mi-novembre)
        )
        
        premiere_visite = cls.PropEntretien(
            sujet = "1ère visite en entreprise",
            chemin_modele = r"C:\Users\vt238770\Documents\_CEA\Modèles adaptés\Mails\Master IN\Master IN - Suivi d'alternance - Entretien.oft",
            duree = timedelta(hours=1),
            date_debut = RDV_Outlook.get_lundi_depuis_num_semaine(numero_semaine=51) # (Autour du 15 décembre : dernière semaine avant vacances Noël et reprise école → A faire avant mi-janvier)
        )

        deuxieme_visite = cls.PropEntretien(
            sujet = "2ème visite en entreprise",
            chemin_modele = r"C:\Users\vt238770\Documents\_CEA\Modèles adaptés\Mails\Master IN\Master IN - Suivi d'alternance - Entretien.oft",
            duree = timedelta(hours=1),
            date_debut = RDV_Outlook.get_lundi_depuis_num_semaine(numero_semaine=17) # (Autour du 20 avril : dernière semaine reprise école → A faire avant fin mai)
        )
        
        # Fichiers à renvoyer
        ficheEvaluation1 = cls.PropFichierARenvoyer(
            sujet = "Fiche évaluation 1",
            chemin_fichier = r"\\instnt\partage\FORMATIONS_I\GDRA - Master IN (parcours ADIN-GDRA-SN)\2024-2025\9-stages\Fiche_Evaluation_alternance_M2_Ingénierie_Nucléaire_janvier.docx",
            periode = "mi-année",
            deadline_retour = RDV_Outlook.get_lundi_depuis_num_semaine(numero_semaine=6) # (Autour du 5 février)
        )
        
        ficheEvaluation2 = cls.PropFichierARenvoyer(
            sujet = "Fiche évaluation 2",
            chemin_fichier = r"\\instnt\partage\FORMATIONS_I\GDRA - Master IN (parcours ADIN-GDRA-SN)\2024-2025\9-stages\Fiche_Evaluation_alternance_M2_Ingénierie_Nucléaire_aout.docx",
            periode = "fin d'année",
            deadline_retour = RDV_Outlook.get_lundi_depuis_num_semaine(numero_semaine=35) # (Autour du 25 août)
        )


        # Initialisation de l'instance
        instance = cls(chemin_fichier_etudiants, nom_onglet)

        # Constantes du contexte UGA
        instance._annee_scolaire = annee_scolaire
        instance._prefixe_sujet = prefixe_sujet
        instance._chemin_modele_mail_priseContact = chemin_modele_mail_priseContact
        instance._envoyer_mail = envoyer_mail
        instance._mail_responsables_univ = mail_responsables_univ
        instance._colonnes_fe_etudiants = colonnes_fe_etudiants
        
        # On réduit le DataFrame aux informations qui nous sont utiles
        instance._df_etudiants = instance._fe_etudiants._tableaux[nom_onglet]._df[instance._colonnes_fe_etudiants] # On ne garde que les colonnes qui nous intéressent mais attention ça reste une vue dont les modifications affectent le dataframe initial
        instance._df_etudiants = instance._df_etudiants[instance._df_etudiants["Ma fonction de suivi de l'alternant"] == "Tuteur"]

        # RDV entretiens
        instance._entretiens.append(prise_de_fonction)
        instance._entretiens.append(premiere_visite)
        instance._entretiens.append(deuxieme_visite)

        # Mail + RDV de rappel fiche évaluation TODO : faire liste
        instance._fichiersARenvoyer.append(ficheEvaluation1)
        instance._fichiersARenvoyer.append(ficheEvaluation2)

        # Relances → Doit avoir la même structure que les colonnes Excel qui trace les retours tuteurs et apprentis
        instance._relances.append("Engagement des parties")
        for entretien in instance._entretiens:
            instance._relances.append(entretien.sujet)
        for fichierARenvoyer in instance._fichiersARenvoyer:
            instance._relances.append(fichierARenvoyer.sujet)


        # Tests :
        #instance._df_etudiants = instance._df_etudiants.head(1)
        
        return instance

    @classmethod
    def L3D(cls) -> Traiter_contactsApprentis:
        """
        """

        # === Initialisations statiques ===
        annee_scolaire = "2025-2026"
        envoyer_mail = False

        prefixe_sujet = "LP3D - Suivi d'alternance"
        mail_responsables_univ = "isabelle.techer@unimes.fr"

        chemin_modele_mail_priseContact = r"C:\Users\vt238770\Documents\_CEA\Modèles adaptés\Mails\LP3D\LP3D - Tutorat en entreprise - Prise de contact.msg"

        chemin_fichier_etudiants =  fr"\\instnt\partage\FORMATIONS_I\LP3D+-démantelement désamiantage dépollution\{annee_scolaire}\1-dossier etudiants\LP3D - {annee_scolaire.replace('-', '_')}.xlsx"
        
        nom_onglet = "Etudiants"
        
        colonnes_fe_etudiants = [
            "Cursus", 
            "Nom", "Prénom", "Mail apprenti", "Téléphone apprenti", 
            "Entreprise ", "Lieu entreprise", "Nom TE", "Prénom TE", "Mail TE", "Téléphone TE", "Ma fonction de suivi de l'alternant", 
            "Première prise de contact", "Documents CFA à viser", "Entretien d'installation", "Entretien d’installation + fin période 1 en entreprise", "2ème entretien : fin période 2 en entreprise", "3ème entretien : milieu période 3 en entreprise"]

        # Entretiens
        premiere_visite = cls.PropEntretien(
            sujet = "Entretien d’installation + fin période 1 en entreprise",
            chemin_modele = r"C:\Users\vt238770\Documents\_CEA\Modèles adaptés\Mails\LP3D\LP3D - Suivi d'alternance - Entretien.oft",
            duree = timedelta(hours=0, minutes=45),
            date_debut = RDV_Outlook.get_lundi_depuis_num_semaine(numero_semaine=51) # (Autour du 15 décembre : dernière semaine avant vacances Noël et reprise école → A faire avant début janvier)
        )
        
        deuxieme_visite = cls.PropEntretien(
            sujet = "2ème entretien : fin période 2 en entreprise",
            chemin_modele = r"C:\Users\vt238770\Documents\_CEA\Modèles adaptés\Mails\LP3D\LP3D - Suivi d'alternance - Entretien.oft",
            duree = timedelta(hours=1),
            date_debut = RDV_Outlook.get_lundi_depuis_num_semaine(numero_semaine=15) # (Autour du 6 avril : dernière semaine avant reprise école → A faire avant 17 avril)
        )

        troisieme_visite = cls.PropEntretien(
            sujet = "3ème entretien : milieu période 3 en entreprise",
            chemin_modele = r"C:\Users\vt238770\Documents\_CEA\Modèles adaptés\Mails\LP3D\LP3D - Suivi d'alternance - Entretien.oft",
            duree = timedelta(hours=1),
            date_debut = RDV_Outlook.get_lundi_depuis_num_semaine(numero_semaine=28) # (Autour du 6 juillet : avant vacances de chacun → A faire avant fin août)
        )
        

        # Initialisation de l'instance
        instance = cls(chemin_fichier_etudiants, nom_onglet)

        # Constantes du contexte UGA
        instance._annee_scolaire = annee_scolaire
        instance._prefixe_sujet = prefixe_sujet
        instance._chemin_modele_mail_priseContact = chemin_modele_mail_priseContact
        instance._envoyer_mail = envoyer_mail
        instance._mail_responsables_univ = mail_responsables_univ
        instance._colonnes_fe_etudiants = colonnes_fe_etudiants
        
        # On réduit le DataFrame aux informations qui nous sont utiles
        instance._df_etudiants = instance._fe_etudiants._tableaux[nom_onglet]._df[instance._colonnes_fe_etudiants] # On ne garde que les colonnes qui nous intéressent mais attention ça reste une vue dont les modifications affectent le dataframe initial
        instance._df_etudiants = instance._df_etudiants[instance._df_etudiants["Ma fonction de suivi de l'alternant"] == "Tuteur"]

        # RDV entretiens
        instance._entretiens.append(premiere_visite)
        instance._entretiens.append(deuxieme_visite)
        instance._entretiens.append(troisieme_visite)

        # Relances → Doit avoir la même structure que les colonnes Excel qui trace les retours tuteurs et apprentis
        instance._relances.append("Documents CFA à viser")
        instance._relances.append("Entretien d'installation")
        for entretien in instance._entretiens:
            instance._relances.append(entretien.sujet)


        # Tests :
        #instance._df_etudiants = instance._df_etudiants.head(1)
        
        return instance

    def creer_mails_contactInitial(self, chemin_modele_mail_priseContact:Optional(str)=None) -> None:
        """
        Prépare un mail de contact initial des alternants.
            - Soit à partir de self._chemin_modele_mail_priseContact
            - Soit à partir d'un modèle .msg donné en argument
        """

        if chemin_modele_mail_priseContact is None:
            chemin_modele_mail_priseContact = self._chemin_modele_mail_priseContact

        # Prise de contact
        # On ouvre un modèle de mail existant et on prépare un mail pour chaque apprenti
        for _, ligne in self._df_etudiants.iterrows():
            destinataire = ligne["Mail apprenti"]
            copie = ligne["Mail TE"]

            # Vérification basique (on peut aussi tester si c'est une adresse mail valide)
            if pd.notna(destinataire):
                Mail.depuis_modele(
                    chemin_modele=chemin_modele_mail_priseContact,
                    destinataires=destinataire,
                    copies=copie
                )

    def creer_mails_ficheEvaluation(self, date_deadline_retourFiche:date=None) -> None:
        """
        Prépare un mail pour envoyer les fiches d'évaluation aux tuteurs entreprises.
            - mail avec pj
            - rdv à deadline retour
        """
        periode = self.periode_scolaire_UGA()
        if periode == "mi-année" :
            prop = self._fichiersARenvoyer[0]
        else:
            prop = self._fichiersARenvoyer[1]

        
        # Si non existant, on demande la deadline à l'utilisateur :
        if date_deadline_retourFiche is None:
            date_deadline_retourFiche = prop.deadline_retour

        # On prépare le mail et le RDV outlook pour chaque tuteur entreprise
        for _, ligne in self._df_etudiants.iterrows():
            self._creer_mail_ficheEvaluation(ligne, date_deadline_retourFiche)

    def creer_rdv(self, prop:PropEntretien) -> None:
        # On ouvre un modèle de mail existant et on prépare un mail pour chaque apprenti
        for _, ligne in self._df_etudiants.iterrows(): # Si je,veux boucler que le 1er élément : self._df_etudiants.head(1).iterrows()
            # On prépare l'adaptation du RDV
            participants_obligatoires = ";".join([
                    ligne["Mail apprenti"],
                    ligne["Mail TE"]
                ])

            # On créée le RDV à partir du modèle
            RDV_Outlook.depuis_modele_oft(
                chemin_modele=prop.chemin_modele,
                sujet = f"{self._prefixe_sujet} - {prop.sujet} - {ligne['Prénom']} {ligne['Nom']}",
                lieu = f'{ligne["Entreprise "]} {ligne["Lieu entreprise"]}',
                date_debut = prop.date_debut,
                duree = prop.duree,
                categorie = prop.categorie, # Si je veux en mettre plusieurs, je peux mettre "Urgent, Projet A"
                participants_obligatoires=participants_obligatoires
                )

    @staticmethod
    def creer_html_mail_relance_BAK(choix:int, ligne:pd.Series, lettreInterlocuteur:str) -> str|None:
        corps_html = """
        <html>
        <body>
            <p>Bonjour,</p>
            <p>Dans le cadre du suivi de l'alternance il vous manque certaines actions.</p>
            <p>S'il vous plaît, est-ce que vous pouvez <span style="color: red; font-weight: bold;">au plus tôt</span> :</p>

            <ul>
        """

        len_corps_html_ini = len(corps_html)
        
        if choix in (1, 2, 3, 4, 5, 6) :
            if lettreInterlocuteur not in str(ligne.get("Engagement des parties") or "").strip():
                #print("Relance engagement des parties")
                corps_html += """<li>Signer dans Studea la section <strong>"Engagement des parties"</strong></li>\n"""
        if choix in (2, 3, 4, 5, 6) :
            if lettreInterlocuteur not in str(ligne.get("Entretien initial") or "").strip():
                #print("Relance entretien prise de fonction")
                corps_html += """<li>Compléter et signer dans Studea <strong>l'entretien de prise de fonction</strong> dans la partie "Visites en entreprise"</li>\n"""
        if choix in (3, 4, 5, 6) :
            if lettreInterlocuteur not in str(ligne.get("1ère visite") or "").strip():
                #print("Relance 1ère visite")
                corps_html += """<li>Compléter et signer dans Studea <strong>le relevé de conclusion de la première visite en entreprise</strong> dans la partie "Visites en entreprise"</li>\n"""
        if choix in (4, 5, 6) :
            if lettreInterlocuteur not in str(ligne.get("2ème visite") or "").strip():
                #print("Relance 2ème visite")
                corps_html += """<li>Compléter et signer dans Studea <strong>le relevé de conclusion de la deuxième visite en entreprise</strong> dans la partie "Visites en entreprise"</li>\n"""
        if choix == 5:
            if lettreInterlocuteur not in str(ligne.get("Fiche évaluation 1") or "").strip():
                #print("Relance fiche d'évaluation 1")
                corps_html += """<li>Compléter et nous renvoyer <span style="color: red; font-weight: bold;">la fiche d'évaluation du 1er semestre nécessaire pour la soutenance de début janvier</span></li>\n"""
        if choix == 6:
            if lettreInterlocuteur not in str(ligne.get("Fiche évaluation 2") or "").strip():
                #print("Relance fiche d'évaluation 2")
                corps_html += """<li>Compléter et nous renvoyer <span style="color: red; font-weight: bold;">la fiche d'évaluation du 2ème semestre nécessaire pour la soutenance finale</span></li>\n"""

        if len(corps_html) > len_corps_html_ini:
            corps_html += """
                </ul>

                <p>Je vous remercie, passez une excellente fin de journée,</p>
            </body>
            </html>"""
            return corps_html
        else:
            return None      

    def creer_html_mail_ficheEvaluation(self, ligne:pd.Series, date_deadline_retourFiche:date) -> str:
        
        periode = self.periode_scolaire_UGA(date_deadline_retourFiche)
        date_str = format_date(date_deadline_retourFiche, "d MMMM", locale="fr")

        corps_html = """
        <html>
        <body>
            <p>Bonjour,</p>
            <p>Dans le cadre de la notation de """ + f'{ligne["Prénom"]} {ligne["Nom"]}' + f" lors de sa prochaine soutenance de {periode}, " + """est-ce que vous pouvez compléter et signer le document en PJ s'il vous plait ? Ce document restera strictement confidentiel à l'équipe pédagogique et ne sera pas divulgué à votre apprenti.</p>
            <p>Il sera à retourner à moi-même en mettant en copie """ + f"{self._mail_responsables_univ}" + f" pour le <span style='color: red; font-weight: bold;'>{date_str} au plus tard</span> " + """.</p>
            <p>Je vous envoie un avis de rdv pour faire office de pense-bête. Il sera mis au """ + f"{date_str}" + """ avec un rappel une semaine avant (mais en mode « libre » histoire de ne pas bloquer le créneau sur votre agenda donc vous pouvez l’accepter dans risque).</p>
            """

        if periode == "fin d'année" :
            corps_html += """<p>Si vous êtes en vacances durant cette période, veillez me l’envoyer avant de profiter de votre repos mérité !</p>"""

        corps_html += """
                <p>Je vous remercie, passez une excellente fin de journée,</p>
            </body>
            </html>"""
       
        return corps_html

    def _creer_mail_ficheEvaluation(self, ligne:pd.Series, date_deadline_retourFiche:date):
        """
        Pour un alternant (fonction appelée depuis une boucle) : 
           - envoie un mail avec la fiche d'évaluation en pj
           - envoie un rdv outlook avec rappel 1 semaine avant la date butoir
        """
        corps_html = self.creer_html_mail_ficheEvaluation(ligne, date_deadline_retourFiche)
        pieces_jointes = self.defini_pj_ficheEvaluation_UGA()
        sujet = self._prefixe_sujet + " - Fiche d'évaluation à compléter et retourner"
        
        Mail.creer_mail(
            destinataires = ligne["Mail TE"],
            sujet = sujet,
            corps_html = corps_html,
            pieces_jointes = pieces_jointes,
            envoyer_mail = self._envoyer_mail
        )

        date_debut = datetime.combine(date_deadline_retourFiche, time(hour=8, minute=0))
        date_fin = date_debut + timedelta(minutes=30)
        RDV_Outlook.depuis_html(
            sujet = sujet,
            date_debut = date_debut,
            date_fin = date_fin,
            categorie = "FI",
            html = corps_html,
            participants_obligatoires = ligne["Mail TE"],
            pieces_jointes = pieces_jointes,
            disponibilite = "Libre",
            importance = 2,
            rappel_active = True,
            rappel_minutes = RDV_Outlook.calculer_rappel_minutes(jours=7),
            envoyer = self._envoyer_mail
        )

    def creer_mails_relances(self, relance: str):
        """
        Envoie les mails de relance pour un intitulé donné (élément de self._relances).
        """
        try:
            ind = self._relances.index(relance) + 1  # convertit en indice 1-based
        except ValueError:
            raise ValueError(f"Relance inconnue: {relance!r}")

        # Pièces jointes seulement pour les fiches d’évaluation
        if "évaluation" in self._relances[ind-1]:
            pieces_jointes = self.defini_pj_ficheEvaluation_UGA()
        else:
            pieces_jointes = None

        for _, ligne in self._df_etudiants.iterrows():
            # Apprenti
            corps_html = self.creer_html_mail_relance(ind, ligne, "A")
            if corps_html:
                Mail.creer_mail(
                    destinataires=ligne["Mail apprenti"],
                    sujet=self._prefixe_sujet + "  - Relance actions suivis de l'alternance",
                    corps_html=corps_html,
                    pieces_jointes=pieces_jointes,
                    envoyer_mail=self._envoyer_mail,
                )

            # Tuteur entreprise
            corps_html = self.creer_html_mail_relance(ind, ligne, "T")
            if corps_html:
                Mail.creer_mail(
                    destinataires=ligne["Mail TE"],
                    sujet=self._prefixe_sujet + " - Relance actions suivis de l'alternance",
                    corps_html=corps_html,
                    pieces_jointes=pieces_jointes,
                    envoyer_mail=self._envoyer_mail,
                )


    def creer_html_mail_relance(self, ind:int, ligne:pd.Series, lettreInterlocuteur:str) -> str|None:
        """
        On choisit par définition que tout rappel d'un certain indice du tableau des rappel implique un rappel des entretiens/trucs précédents
        ind est la valeur sélectionnée par l'utilisateur et commence à 1. Donc sa correspondance dans le tableau self._relances c'est ind-1
        """
        corps_html = """
        <html>
        <body>
            <p>Bonjour,</p>
            <p>Dans le cadre du suivi de l'alternance il vous manque certaines actions.</p>
            <p>S'il vous plaît, est-ce que vous pouvez <span style="color: red; font-weight: bold;">au plus tôt</span> :</p>

            <ul>
        """

        len_corps_html_ini = len(corps_html)
        #for i, typeRelance in enumerate(self._relances):
        for i in range(ind):
            if lettreInterlocuteur not in str(ligne.get(self._relances[i]) or "").strip():
                if "évaluation" in self._relances[i]: # Cas fiches d'évaluation
                    corps_html += f"<li>Compléter et nous renvoyer <span style=\"color: red; font-weight: bold;\">la fiche d'évaluation nécessaire pour la soutenance de {self.periode_scolaire_UGA()}</span></li>\n"
                else: # Cas studea
                    corps_html += fr"<li>Compléter et signer dans Studea la ligne <strong>{self._relances[i]}</strong></li>"

        
        if len(corps_html) > len_corps_html_ini:
            corps_html += """
                </ul>

                <p>Je vous remercie, passez une excellente fin de journée,</p>
            </body>
            </html>"""
            return corps_html
        else:
            return None      

    def defini_pj_ficheEvaluation_UGA(self) -> str:
        """
        Sélectionne et renvoie le chemin de la bonne fiche d'évaluation en fonction de la période (mi-année ou fin d'année) → Ancienne méthode
        """
        periode = self.periode_scolaire_UGA()
        fichier = next(
            (f for f in self._fichiersARenvoyer if f.periode == periode),
            None  # valeur par défaut si rien trouvé
        )
        
        """
        if (periode == "mi-année"):
            #print("On est entre septembre et février (inclus)")
            pieces_jointes = self._ficheEvaluation1
        else:
            #print("On est entre mars et août")
            pieces_jointes = self._ficheEvaluation2
        """
        
        return fichier.chemin_fichier

    def rdv_prise_de_fonction_BAK(self, chemin_modele_rdv:str) -> None:

        # Durée type : 1h
        duree = timedelta(hours=0, minutes=45) #timedelta(hours=1, minutes=30)

        # On ouvre un modèle de mail existant et on prépare un mail pour chaque apprenti
        for _, ligne in self._df_etudiants.iterrows():

            # On prépare l'adaptation du RDV
            sujet = f"Master IN - Suivi d'alternance - Entretien de prise de fonction - {ligne['Prénom']} {ligne['Nom']}"
            #date_debut = 
            #date_fin = 
            lieu = f'{ligne["Entreprise "]} {ligne["Lieu entreprise"]}'
            participants_obligatoires = ";".join([
                    ligne["Mail apprenti"],
                    ligne["Mail TE"]
                ])

            # On créée le RDV à partir du modèle
            RDV_Outlook.depuis_modele_oft(
                chemin_modele=chemin_modele_rdv,
                sujet = sujet,
                lieu = lieu,
                duree = duree,
                categorie = "FI", # Si je veux en mettre plusieurs, je peux mettre "Urgent, Projet A"
                participants_obligatoires=participants_obligatoires
                )

    @staticmethod
    def periode_scolaire_UGA(date_input:date=None) -> str:
        """
        La fonction emploie soit une date rentrée en argument, soit la date du jour.
        
        Renvoie "mi-année" si date est entre septembre et février.
        Renvoie "fin d'année" sinon
        """
        # Période à définir pour la fiche d'évaluation à employer
        if date_input is None:
            mois = date.today().month  # 1=janvier, ..., 12=décembre
        else:
            mois = date_input.month

        if ((9 <= mois) or (mois <= 2)):  # septembre (9) → février (2)
            #print("On est entre septembre et février (inclus)")
            periode = "mi-année"
        else:
            #print("On est entre mars et août")
            periode = "fin d'année"        
        return periode


    @staticmethod
    def _diagnostiquer_msg(chemin_modele: str):
        print("== Diagnostique du fichier .msg ==\n")
        print(f"Chemin du fichier : {chemin_modele}\n")

        # 1. Analyse via extract_msg
        print("--> Analyse avec extract_msg")
        try:
            msg = extract_msg.Message(chemin_modele)
            print("Sujet         :", msg.subject)
            print("Expéditeur    :", msg.sender)
            print("Date          :", msg.date)
            print("Contient HTML :", bool(msg.htmlBody))
            print("Contient Texte:", bool(msg.body))
            if msg.htmlBody:
                print(" - Taille HTMLBody :", len(msg.htmlBody))
                print(" - HTMLBody contient <ul> :", "<ul>" in msg.htmlBody.lower())
                print(" - HTMLBody contient puce (•) :", "•" in msg.htmlBody)
        except Exception as e:
            print("Erreur extract_msg :", e)

        print("\n--> Analyse avec Outlook COM (OpenSharedItem)")
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            session = outlook.Session
            modele = session.OpenSharedItem(os.path.abspath(chemin_modele))

            print("Sujet         :", modele.Subject)
            print("Location      :", modele.Location)
            print("Body (200c)   :", modele.Body[:200] if hasattr(modele, "Body") else "[Aucun]")
            
            # Test HTMLBody
            try:
                html = modele.HTMLBody
                print("Contient HTMLBody :", True)
                print(" - Taille HTMLBody :", len(html))
                print(" - HTMLBody contient <ul> :", "<ul>" in html.lower())
                print(" - HTMLBody contient puce (•) :", "•" in html)
            except AttributeError:
                print("Contient HTMLBody :", False)
            
            # Test RTFBody si dispo
            if hasattr(modele, "RTFBody"):
                print("Contient RTFBody :", True)
            else:
                print("Contient RTFBody :", False)

        except Exception as e:
            print("Erreur Outlook COM :", e)




### --------------------------------------------------------------------
#  Fonctions globales INSTN
### --------------------------------------------------------------------

def initialiser_PropExportIRIS_de_config() -> None:
    """
    Initialise les dataclass PropExportIRIS de config.py
    Je suis obligé de fonctionner comme ça car sinon :
       - config.py importe instn.py
       - instn.py importe config.py
    → Référence circulaire
    """
    config.IRIS_SESSIONS = PropExportIRIS(**config.IRIS_SESSIONS_PARAMS)
    config.IRIS_FORMATIONS = PropExportIRIS(**config.IRIS_FORMATIONS_PARAMS)
    config.IRIS_VENTES = PropExportIRIS(**config.IRIS_VENTES_PARAMS)
    config.IRIS_INSCRIPTIONS = PropExportIRIS(**config.IRIS_INSCRIPTIONS_PARAMS)

def verifier_code_iris(valeur: Any, type_sortie: Type = str) -> Tuple[bool, Any]:
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
        >>> verifier_code_iris(12345)
        (True, '12345')

        >>> verifier_code_iris("01234")
        (True, '01234')

        >>> verifier_code_iris("9999")
        (False, None)

        >>> verifier_code_iris("12345.0")
        (True, '12345')

        >>> verifier_code_iris("abcde")
        (False, None)

        >>> verifier_code_iris("67890", int)
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


def fenetreBilanFormation():
    def valider_champs(*args):
        trig = entry_trigramme.get().strip()
        annee = entry_annee.get().strip()
        bouton_generer.config(
            state="normal" if len(trig) == 3 and annee.isdigit() and len(annee) == 4 else "disabled"
        )

    def generer_bilan_formation(event=None):
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
        bf.mergeBilan()
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

    bouton_generer = ttk.Button(frame_boutons, text="Générer bilan", state="disabled", command=generer_bilan_formation)
    bouton_generer.grid(row=0, column=0, padx=5)

    bouton_annuler = ttk.Button(frame_boutons, text="Annuler", command=annuler)
    bouton_annuler.grid(row=0, column=1, padx=5)

    # Validation en temps réel
    entry_trigramme.bind("<KeyRelease>", valider_champs)
    entry_annee.bind("<KeyRelease>", valider_champs)

    # Entrée = clic sur bouton générer
    fenetre.bind("<Return>", generer_bilan_formation)

    # Échap = fermeture de la fenêtre
    fenetre.bind("<Escape>", lambda e: fenetre.destroy())

    # Lancer la fenêtre
    fenetre.mainloop()

def recupere_trig_formation_depuis_chemin(chemin:Optional[Path] = None) -> str:
    """
    Extrait un trigramme (3 lettres/chiffres) depuis un chemin, ou le demande à l'utilisateur si introuvable.
    Gère les slashs / et \\ de manière robuste, et accepte un `Path` ou une `str`.

    Paramètres
    ----------
    chemin : Path | str | None
        Chemin d'où extraire le trigramme. Peut être un objet `pathlib.Path`, une chaîne, ou `None`.

    Retour
    ------
    str
        Le trigramme de la formation (3 caractères alphanumériques).
        Si aucun trigramme n'est trouvé dans le chemin, la méthode le demande à l'utilisateur.

    Exemples
    --------
    >>> self._recupere_trig_formation_depuis_chemin("C:/Formations/L3D/stagiaires.csv")
    'L3D'

    >>> self._recupere_trig_formation_depuis_chemin(Path("D:/data/UGA/stagiaires.csv"))
    'UGA'

    >>> self._recupere_trig_formation_depuis_chemin(None)
    [ouvre une boîte de dialogue pour demander le trigramme]
    """

    trigramme_formation = None

    if chemin:
        # Découpe le chemin en segments
        parties = chemin.parts

        # Recherche un segment de 3 caractères alphanumériques
        for part in parties:
            if re.fullmatch(r"[A-Z0-9]{3}", part):
                trigramme_formation = part
                break

        # Si rien trouvé, on demande à l'utilisateur
        if not trigramme_formation:
            trigramme_formation = demander_code(typeCode="Trigramme formation", chemin=chemin)

    else:
        # Aucun chemin fourni → demande directe à l'utilisateur
        trigramme_formation = demander_code(typeCode="Trigramme formation", chemin=chemin)

    return trigramme_formation

def demander_code(typeCode:str, chemin:Path) -> int|str:
    # Initialisation en fonction du type de code
    
    match typeCode:
        case "Code IRIS":
            print("Cas code IRIS")
            nbCaracteres = 5
        case "Trigramme formation":
            print("Cas trigramme formation")
            nbCaracteres = 3
        case _:
            log_erreur(f"cas non valide : soit 'IRIS' soit 'Trigramme formation', demandé : {typeCode} → exit()")
            exit()

        
    def verifier_entree(*args):
        val = entry_code.get()
        bouton_valider.config(state="normal" if val.isdigit() and len(val) == nbCaracteres else "disabled")

    def valider():
        nonlocal code # Adaptation : changer en "code"
        code = entry_code.get()            
        fenetre.destroy()


    def annuler():
        fenetre.destroy()
        log_erreur(f"{typeCode} non renseigné → exit()")
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
        text=f"Chemin du fichier source :\n  *{chemin}*",
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

    bouton_annuler = tk.Button(frame_boutons, text="Annuler", width=10, command=annuler)
    bouton_annuler.grid(row=0, column=2, padx=5)

    fenetre.mainloop()

    code = int(code) if code.isdigit() else str(code)
    return code

def choisir_fichier_iris_sessions() -> Path | None:
    """
    Ouvre un filedialog pour demander à l'utilisateur de sélectionner un extract IRIS sessions.
    On pointe au mieux sur le répertoire des extracts IRIS pour la boîte de dialogue.
    """

    return choisir_fichier(titre=f"Sélectionner l'extract IRIS session {config.IRIS_SESSIONS._codeExport} à employer.",
                    types_fichiers=[("Fichiers Excel", "*.xlsx")],
                    dossier_initial=config.IRIS_SESSIONS._output.repertoire, # Pour aller vers mes ficheirs concaténés, sinon pour les originaux il faut pointer vers input
                    texte_bouton_choisir=f"Choisir extract IRIS session {config.IRIS_SESSIONS._codeExport} à nouveau"
                    )




def recupere_trig_formation_depuis_chemin_avec_renommage(chemin:Optional[Path] = None, pointeur_chemin:Optional[list[Path]] = None, renommage:Optional[Callable[[Path], None]] = None) -> str:
    """
    Extrait un trigramme (3 lettres/chiffres) depuis un chemin, ou le demande à l'utilisateur si introuvable.
    Gère les slashs / et \\ de manière robuste, et accepte un `Path` ou une `str`.

    J'emploie une astuce :
        Dans demander_code() il peut il y avoir un changement denom du fichier chemin
        Comme il n'existe pas de pointeurs dans Python, j'emploie un objet mutable (i.e. qui peut être modifié en dehors comme des pointeurs : ici une liste)
        Donc s'il n'y a pas de chemin entrée, il faut que ce soit son pointeur

    Paramètres
    ----------
    chemin : Path | str | None
        Chemin d'où extraire le trigramme. Peut être un objet `pathlib.Path`, une chaîne, ou `None`.

    Retour
    ------
    str
        Le trigramme de la formation (3 caractères alphanumériques).
        Si aucun trigramme n'est trouvé dans le chemin, la méthode le demande à l'utilisateur.

    Exemples
    --------
    >>> self._recupere_trig_formation_depuis_chemin("C:/Formations/L3D/stagiaires.csv")
    'L3D'

    >>> self._recupere_trig_formation_depuis_chemin(Path("D:/data/UGA/stagiaires.csv"))
    'UGA'

    >>> self._recupere_trig_formation_depuis_chemin(None)
    [ouvre une boîte de dialogue pour demander le trigramme]
    """

    trigramme_formation = None

    # On vérifie qu'il y a soit chemin, soit pointeur_chemin sinon on coupe
    if (not chemin) and (not pointeur_chemin):
        raise TypeError("Il faut que soit chemin, soit pointeur_chemin soit renseigné")

    chemin_a_tester = None
    if chemin:
        chemin_a_tester = chemin
    elif pointeur_chemin:
        chemin_a_tester = pointeur_chemin[0]

    if chemin_a_tester:
        # Découpe le chemin en segments
        parties = chemin.parts

        # Recherche un segment de 3 caractères alphanumériques
        for part in parties:
            if re.fullmatch(r"[A-Z0-9]{3}", part):
                trigramme_formation = part
                break

        # Si rien trouvé, on demande à l'utilisateur
        if not trigramme_formation:
            #trigramme_formation = demander_code("Trigramme formation", renommage=renommage)
            trigramme_formation = demander_code("Trigramme formation", renommage=renommage)

    else:
        # Aucun chemin fourni → demande directe à l'utilisateur
        trigramme_formation = demander_code("Trigramme formation", renommage=renommage)

    return trigramme_formation

def demander_code_avec_renommage(typeCode:str, chemin:Path, renommage:Optional[Callable[[Path], None]] = None) -> int|str:
    # Initialisation en fonction du type de code
    
    match typeCode:
        case "Code IRIS":
            print("Cas code IRIS")
            nbCaracteres = 5
        case "Trigramme formation":
            print("Cas trigramme formation")
            nbCaracteres = 3
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
        nouveau_chemin = chemin.parent / nouveau_nom

        try:
            os.rename(chemin, nouveau_chemin)
            renommage(chemin)  # On appelle le callback
        except Exception as e:
            tk.messagebox.showerror("Erreur", f"Impossible de renommer le fichier :\n{e}")
            log_erreur("Erreur", f"Impossible de renommer le fichier :\n{e}")
            return  # Ne pas fermer la fenêtre si erreur

        fenetre.destroy()

    def annuler():
        fenetre.destroy()
        log_erreur(f"{typeCode} non renseigné pour {chemin} → exit()")
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
        text=f"Chemin du fichier source :\n  *{chemin}*",
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

### --------------------------------------------------------------------
#  Initialisations variables globales communes
### --------------------------------------------------------------------

# Requis pour avoir des PropExportIRIS dans config.py (dinon références circulaires à l'import)
initialiser_PropExportIRIS_de_config()

# Pour couleur barres de progression
colorama.init(autoreset=True)

# Pour chrono des fonctions
timer = Timer()

# Pour message de sortie applis externes
vlog = Vlog()


