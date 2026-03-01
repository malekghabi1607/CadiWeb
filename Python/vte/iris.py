from __future__ import annotations
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from tkinter.ttk import Style
from typing import List, Optional, Tuple

from colorama import Fore
import pandas as pd
import tqdm

from vte import config, config_extractsIRIS
from vte.office import FichierExcel
from vte.utils import *
from vte.utils_instn import demander_code


# ======================================================================================
# CLASSE IRIS
# ======================================================================================
@dataclass
class InfosExportsIRIS:
    repertoire:Optional[Path]
    chemin_fichier:Optional[Path]
    nom_onglet:Optional[str]
    nbLignes_avantET:Optional[int]

    ordre_colonne:Optional[List[str]]
    #TODO : à moyen terme je pourrai avoir une double liste avec le nom des colonnes initiales + ordre ; voire mettre 2 lignes d'en-têtes dans le modèle (nom extract IRIS & nom plus adapté)

class ConfigExportIRIS:
    """
    Configuration d'un type d'export IRIS.
    Contient les propriétés input, modèle, output.
    """
    # === Constructeurs ===
    def __init__(self, 
                nom_typeExport:str, 
                codeExport:str, 
                
                repertoire_input:Path, 
                nom_onglet_input:str = "Data", 
                nbLignes_avantET_input:int = 0, 
                
                repertoire_modele:Optional[Path] = config.REPERTOIRES_MODELES, 
                nom_fichier_modele:Optional[str|Path] = None, 
                ordre_colonnes_modele:Optional[List[str]] = None,
                
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

class IRIS:
    """
    C'est la classe qui contient l'environnement pour bosser sur des fichiers Exports IRIS
    """
    
    # TODO Mettre à jour extracts IRIS complets
    """
    # Lire les extracts
    # Faire un ls du répertoire cible des extracts
    # Ouvrir extract vomplet et regarder les fichiers inclus
    # Comparer fichiers inclus et ls
    # Supprimer des lignes (si besoin)
    # Ajouter fichiers nécessaires 
    """


    # ===========================
    # === Variables de classe ===
    # ===========================
    _SESSIONS:ConfigExportIRIS = ConfigExportIRIS(**config.IRIS_SESSIONS_PARAMS)
    _FORMATIONS = ConfigExportIRIS(**config.IRIS_FORMATIONS_PARAMS)
    _VENTES = ConfigExportIRIS(**config.IRIS_VENTES_PARAMS)
    _INSCRIPTIONS = ConfigExportIRIS(**config.IRIS_INSCRIPTIONS_PARAMS)


    _dict_exports_IRIS = {}
    _dict_exports_IRIS["ConfigExportIRIS"] = {
        "Sessions" : _SESSIONS, 
        "Formations" : _FORMATIONS,
        "Ventes" : _VENTES,
        "Inscriptions" : _INSCRIPTIONS
        }
    _dict_exports_IRIS["chemins_fichiersInput"] = {
        "Sessions" : config_extractsIRIS._tSessions,
        "Formations" : config_extractsIRIS._tFormations,
        "Ventes" : config_extractsIRIS._tVentes,
        "Inscriptions" : config_extractsIRIS._tInscriptions
        }
    _dict_exports_IRIS["colonnes_modele"] = {
        "Sessions" : None,
        "Formations" : None,
        "Ventes" : None,
        "Inscriptions" : config_extractsIRIS._colonnes_modele_inscriptions
        }


    # ====================
    # === Constructeur ===
    # ====================
    def __init__(self, typeExport:str):
        """
        Initialise un fichier IRIS traité delin d'un type d'export
        
        :param typeExport: spécifie le type d'export. Doit être dans cette liste : ["Sessions", "Formations", "Ventes", "Insciptions"]
        :type typeExport: str
        """
        self._typeExport:str = typeExport  # Nom du type d'export : ["Sessions", "Formations", "Ventes", "Insciptions"]
        
        # Fichier Excel de l'export IRIS
        self._fe:Optional[FichierExcel] = None

    # =================================
    # === Constructeurs alternatifs ===
    # =================================
    @classmethod
    def concatener_exportsIRIS_typeUnique(cls, typeExport:str, chemins_fichiersInput:Optional[str|Path|Tuple[str|Path]] = None, chemin_fichier_sauv:Optional[Path]=None) -> None:
        r"""
        Crée un fichier Excel unique à partir de plusieurs exports IRIS.

        Sauvegarde soit à un endroit en argument soit à sa place par défaut (donné dans config)
        Les fichiers générés reprennent les colonnes d'origine mais avec de meilleures formes (format, couleurs...) + des colonnes adjointes à la fin pour extraire et séparer les infos du n° de session ou de la référence de la formation (ex. : trigramme formation, trigramme RP, trigramme AF...)

        :param typeExport: Type des exports IRIS à traiter, un de ceux-ci ["Formations", "Sessions", "Ventes", "Inscriptions"].
        :type typeExport: str
        :param chemins_fichiersInput: Chemin(s) du ou des fichiers à traiter. Si None, on ouvre un filedialog
        :type chemins_fichiersInput: str|Path|Tuple[str|Path]
        :param chemin_fichier_sauv: Chemin de savegarde si l'utilisateur ne veut pas employer celui qui est dans la config
        :type chemin_fichier_sauv: Path

        :return: Ne retourne rien
        :rtype: None

        :example:

        >>> concatener_exportsIRIS("Sessions", (r'.\R04110_Sessions-2021 FINAL.xlsx', r'.\R04110_Sessions-2022 FINAL.xlsx'))
        >>> concatener_exportsIRIS("Sessions", r'C:/Users/vt238770/Documents/_CEA/_Formations/Extracts IRIS - Faits/Extracts originaux/R04110_Sessions-2021 FINAL.xlsx')


        .. seealso:: Rien du tout.
        .. warning:: On ne doit traiter que des extracts IRIS du même type (sessions, foramtion...)
        .. note:: Remplace complètement l'ancien fichier (pas de mise à jour incrémentale)
        .. todo:: Rien du tout.
        """

        # On crée l'instance
        instance = cls(typeExport=typeExport)

        # Si aucun fichier input n'est donné, alors on ouvre un filedialog
        if chemins_fichiersInput is None:
            chemins_fichiersInput = IRIS.choisir_fichier(typeExport)
        
        # Si aucun fichier de sortie n'est donnée, alors on prend le chemin par défaut
        if chemin_fichier_sauv is None:
            chemin_fichier_sauv = instance.cei._output.chemin_fichier

        # On convertit en Path

        # On convertit en Tuple[Path]
        chemins_fichiersInput = convertir_tuple_path(chemins_fichiersInput)
        chemins_fichiersInput = tuple(chemin_vers_unc(chemin) for chemin in chemins_fichiersInput)

        # Affichage de la liste des fichiers à traiter
        print(Style.BRIGHT + Fore.YELLOW + "\nTraitement des exports " + typeExport)
        chemins_fichiersInput_str = [f"\n\t{f}" for f in chemins_fichiersInput]
        print("Liste des fichiers à concaténer : " + ", ".join(chemins_fichiersInput_str))
        
        # On ouvre le modèle et tous ses tableaux structurés
        fe_modele = FichierExcel.depuis_modele(
                                        chemin_modele=instance.cei._modele.chemin_fichier, 
                                        chemin_fichier_sauv=chemin_fichier_sauv
                                        )

        # On lit le/les extract(s) IRIS et on le/les stocke dans un seul dataframe self._df_tableau
        df_concat_iris = instance._charger_df_extractsIRIS_originaux(chemins_fichiersInput=chemins_fichiersInput)

        # On copie le DataFrame df_concat_iris (le dataframe créé par cls._charger_df_extractsIRIS_originaux) avec les nouvelles données dans le modèle
        fe_modele._tableaux[instance._typeExport].ecrit_dataFrame_dans_tableauStructure(df_concat_iris, supprimeDonneesEtRemplace=True)
        
        # On écrit les références des fichiers copiés dans le tableau structuré "Imports". Nota : openpyxl ne prend pas en charge les Path
        df_chemins = pd.DataFrame([str(chemin) for chemin in chemins_fichiersInput], columns=['Chemin fichier'])
        fe_modele._tableaux["Imports"].ecrit_dataFrame_dans_tableauStructure(df=df_chemins, supprimeDonneesEtRemplace=True)
        
        #On enregistre et on ferme (par précaution car copieformat xlwings sauvegarde)
        fe_modele.save()    
        fe_modele.close()

    @classmethod
    def concatener_exportsIRIS_plusieursTypes(cls, typesExports:str|Tuple[str], depuis_config:bool=False) -> None:
        """
        Génère un ou plusieurs fichier IRIS en concaténant plusieurs extracts originaux.

        Les fichiers à traiter son soit spécifiés dans config_extractsIRIS.py (depuis_config = True) soit on ouvre un filedialog (depuis_config = False = valeur par défaut).
        Gère le traitement de plusieurs types.

        :param typeExport: Type des exports IRIS à traiter, un ou plusieurs de ceux-ci ["Formations", "Sessions", "Ventes", "Inscriptions"].
        :type typeExport: str
        :param depuis_config: Si True, on exploite les fichiers spécifiés dans config_extractsIRIS.py. Sinon (défaut) on ouvre un filedialog.
        :type depuis_config: bool
        
        :return: Ne retourne rien
        :rtype: None

        :example:

        >>> mettreAJourTousLesExportsIRIS_auto(("Sessions", "Formations", "Ventes", "Inscriptions"), depuis_config=True)
        >>> mettreAJourTousLesExportsIRIS_auto("Sessions")


        .. seealso:: concatener_exportsIRIS_typeUnique()
        .. warning:: Rien du tout.
        .. note:: Remplace complètement l'ancien fichier (pas de mise à jour incrémentale)
        .. todo:: Rien du tout.      
        """
        def f_dict_chemins(depuis_config:bool,typeExport:str) -> Tuple[Path]:
            """
            Fonction à employer en fonction du choix utilisateur depuis_config/filedialog.
            
            Si depuis_config = True  : on charge les chemins depuis config_extractsIRIS.
            Si depuis config = False : on ouvre un filedialog pour choix utilisateur.

            :return: Un tuple de chemins d'extracts IRIS originaux (Path)
            :rtype: Tuple[Path]
            """
            if depuis_config:
                # Cas depuis_config
                return tuple((cls.cei(typeExport)._input.repertoire / nom) for nom in cls._dict_exports_IRIS["chemins_fichiersInput"][typeExport])
            else :
                # Cas avec ouverture filedialog
                return choisir_fichier(
                    titre=f"Sélectionner un ou plusieurs fichiers Extract IRIS {cls.cei(typeExport)._nom_typeExport} ({cls.cei(typeExport)._codeExport})",
                    types_fichiers=[("Fichiers Excel", "*.xlsx")],
                    dossier_initial=cls.cei(typeExport)._input.repertoire,
                    multi_fichiers=True
                )

                

        # On convertit en tuple si besoin
        typesExports = convertir_tuple_str(typesExports)

        # On définit le dictionnaire des chemins à employer pour chaque type d'export
        dict_chemins_fichiersInpout = {
            typeExport: f_dict_chemins(depuis_config, typeExport) 
            for typeExport in typesExports
            }

        # On traite les fichiers
        for typeExport in typesExports:
            cls.concatener_exportsIRIS_typeUnique(typeExport, dict_chemins_fichiersInpout[typeExport])

    # TODO : AVIRER
    @classmethod
    def concatener_exportsIRIS_plusieursTypes_depuis_config_extractsIRIS(cls, typesExports:str|Tuple[str]) -> None:
        """
        Génère un fichier IRIS en concaténant plusieurs extracts originaux spécifiés dans config_extractsIRIS.py
        Gère le traitement de plusieurs types.

        :param typeExport: Type des exports IRIS à traiter, un ou plusieurs de ceux-ci ["Formations", "Sessions", "Ventes", "Inscriptions"].
        :type typeExport: str
        
        :return: Ne retourne rien
        :rtype: None

        :example:

        >>> mettreAJourTousLesExportsIRIS_auto(("Sessions", "Formations", "Ventes", "Inscriptions"))
        >>> mettreAJourTousLesExportsIRIS_auto("Sessions")


        .. seealso:: concatener_exportsIRIS_typeUnique()
        .. warning:: Rien du tout.
        .. note:: Remplace complètement l'ancien fichier (pas de mise à jour incrémentale)
        .. todo:: Rien du tout.      
        """
        # On convertit en tuple si besoin
        typesExports = convertir_tuple_str(typesExports)

        # On fait choisir les fichiers à l'utilisateur
        dict_chemins_fichiersInpout = {}
        for typeExport in typesExports:
            dict_chemins_fichiersInpout[typeExport] = tuple((cls.cei(typeExport)._input.repertoire / nom) for nom in cls._dict_exports_IRIS["chemins_fichiersInput"][typeExport])

        # On traite les fichiers
        for typeExport in typesExports:
            IRIS.concatener_exportsIRIS_typeUnique(typeExport, dict_chemins_fichiersInpout[typeExport])

    # TODO : AVIRER
    @classmethod
    def concatener_exportsIRIS_plusieursTypes_depuis_filedialog(cls, typesExports:str|Tuple[str]) -> None:
        # On convertit en tuple si besoin
        typesExports = convertir_tuple_str(typesExports)
        
        # On fait choisir les fichiers à l'utilisateur
        dict_chemins_fichiersInpout = {}
        for typeExport in typesExports:
            cei = cls.cei(typeExport)
            dict_chemins_fichiersInpout[typeExport] = choisir_fichier(
                titre=f"Sélectionner un ou plusieurs fichiers Extract IRIS {cei._nom_typeExport} ({cei._codeExport})",
                types_fichiers=[("Fichiers Excel", "*.xlsx")],
                dossier_initial=cei._input.repertoire,
                multi_fichiers=True
            )

        # On traite les fichiers
        for typeExport in typesExports:
            IRIS.concatener_exportsIRIS_typeUnique(typeExport, dict_chemins_fichiersInpout[typeExport])










    # =========================
    # === Méthodes internes ===
    # =========================
    def _charger_df_extractsIRIS_originaux(self, chemins_fichiersInput:Path|tuple[Path, ...]) -> pd.DataFrame:
        """
        Lit le/les extract(s) IRIS et on le/les stocke dans un seul dataframe self._df_tableau
        (i.e. on concatène si besoin)

        :param chemins_fichiersInput: Chemin(s) du ou des fichiers IRIS (.xlsx) à charger dans le dataframe
        :type chemins_fichiersInput: Path|tuple[Path, ...]

        :return: Un DataFrame contenant les données de chemins_fichiersInput concaténées
        :rtype: pd.DataFrame

        :example:
        >>> self._charger_df_extractsIRIS_originaux(Path(r"C:/fichier1.xlsx"))
        >>> self._charger_df_extractsIRIS_originaux((Path(r"C:/fichier1.xlsx"), Path(r"C:/fichier2.xlsx"))


        .. seealso:: Rien du tout.
        .. warning:: Rien du tout.
        .. note:: Rien du tout.
        .. todo:: Rien du tout.
        """
        # Initialisation : on crée un DataFrame vide pour recevoir (peut-être) des infos que l'on traitera et qui nécessitera d'adjoindre des colonnes à self.__df_tableau
        df_colonnes_sup = None

        if isinstance(chemins_fichiersInput, Path):
            chemins_fichiersInput = (chemins_fichiersInput,)

        # On parcourt le tuple des fichiers à lire
        df_list = [] # Liste des DataFrame qui contiendra chaque fichier Excel séparément
        taille_totale = sum(fichier.stat().st_size for fichier in chemins_fichiersInput) # Calcul taille totale pour barre de progression

        with tqdm(total=taille_totale, unit='o', unit_scale=True, desc=Fore.CYAN+"Lecture des fichiers Excel" + Style.RESET_ALL) as pbar:
            #for i, ifichier in enumerate((os.path.basename(chemin) for chemin in chemins_fichiersInput), 1):
            for i, chemin in enumerate(chemins_fichiersInput, 1):
                # Données pour tqdm
                fichier = chemin.name
                taille = chemin.stat().st_size  # taille en octets 
                pbar.set_postfix(file=fichier, progress=f"{i}/{len(chemins_fichiersInput)}")  # Affichage dynamique dans la barre
                
                df = pd.read_excel(chemin, skiprows=self.cei._input.nbLignes_avantET)
                df_list.append(df)  # On ajoute le DataFrame à notre liste de DataFrame
                
                # Mise à jour de la barre avec la taille du fichier
                pbar.update(taille)

        # Concaténation finale (note : toute la fin de la méthode se fait quasi-instantanément)
        df_concat_iris = pd.concat(df_list, ignore_index=True) 



        # Extraction des infos depuis "référence formation" ou "n° Iris" (dépend du type d'export)
        match self.cei._nom_typeExport:
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
        if self._dict_exports_IRIS["colonnes_modele"][self.cei._nom_typeExport] is not None:
            df_concat_iris = df_concat_iris[self._dict_exports_IRIS["colonnes_modele"][self.cei._nom_typeExport]]
    
        return df_concat_iris


    # =================================
    # === Méthodes pour les filtres ===
    # =================================
    def df_filtre_periode(self, trigramme_formation:Optional[str] = None, annee:Optional[int] = None, periode:Optional[str] = None) -> pd.DataFrame:
        """
        Renvoie le dataframe de l'extract IRIS filtré VTE selon plusieurs critères.
         
         Les critères dépendent du typé d'export ; pour sessions :
            - Statut Session != "Annulée" (toujours) ;
            - Nb. Nommés != 0 (toujours) ;
            - l'année de la session (si donné en argument) ;
            - la période de la session (si donné en argument) ;
            - le trigramme de la formation en cours (si donné en argument).

        Par défaut la période est toute l'année, sinon il faut préciser "1er semestre" ou "2nd semestre".

        #On ouvre et lit l'export sessions IRIS si et seulement si il n'est pas déjà ouvert et lu avant
        
        :param trigramme_formation: trigramme filtré. Si non renseigné : pas de filtre sur ce critère.
        :type trigramme_formation: Optional[str]
        :param annee: année du filtre. Si non renseignée, alors année en cours.
        :type annee: Optional[int]
        :param periode: période du filtre ("1er semestre", "2nd semestre"). Si non renseigné, alors période="Année".
        :type periode: Optional[str]
        :return: extract IRIS VTE filtré
        :rtype: DataFrame
        """
        match self._typeExport:
            case "Sessions":
                return self.df_filtre_periode_sessions(self, annee=annee, periode=periode, trigramme_formation=trigramme_formation)

    def df_filtre_periode_sessions(self, trigramme_formation:Optional[str] = None, annee:Optional[int] = None, periode:Optional[str] = None) -> pd.DataFrame:
        """
        Renvoie le dataframe de l'extract IRIS filtré VTE selon plusieurs critères :
            - Statut Session != "Annulée" (toujours) ;
            - Nb. Nommés != 0 (toujours) ;
            - l'année de la session (si donné en argument) ;
            - la période de la session (si donné en argument) ;
            - le trigramme de la formation en cours (si donné en argument).

        Par défaut la période est toute l'année, sinon il faut préciser "1er semestre" ou "2nd semestre".

        #On ouvre et lit l'export sessions IRIS si et seulement si il n'est pas déjà ouvert et lu avant
        
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
        df_sessions_filtre = self.df[
            (self.df['Statut Session'] != "Annulée") &
            (self.df['Nb. Nommés'] != 0)
        ]
        #print(self.df_sessions_filtre)


        # Si trigramme en argument, alors on filtre sur le trigramme
        if trigramme_formation is not None:
            df_sessions_filtre = df_sessions_filtre[
                (df_sessions_filtre['Trigramme formation'] == trigramme_formation)
            ]
            #print(self.df_sessions_filtre)

        # Si données temporelles (année ou période ou les deux)
        if periode is not None:
            # Si période en argument, alors on filtre sur la période (plus restrictif que l'année)
            date_debut, date_fin = debut_fin_periode(annee=annee, periode=periode)
            df_sessions_filtre = df_sessions_filtre[
                (df_sessions_filtre['Date début ses.'] >= date_debut) &
                (df_sessions_filtre['Date début ses.'] <= date_fin)
            ]   
        elif annee is not None:
            df_sessions_filtre = df_sessions_filtre[
                (df_sessions_filtre['Année début ses.'] == annee)
            ]
        #print(self.df_sessions_filtre)


        # On trie
        df_sessions_filtre.sort_values("Date début ses.", ascending=False)

        return df_sessions_filtre



    # ==========================
    # === Méthodes pour ouvir un extract existant IRIS Session (excel VTE qui concatène plusieurs natifs)  ===
    # Avant j'avais des méthodes externes (cf. ci-dessous) mais j'ai voulu intégrer _fe en variable d'instance
    # Il en résulte que je peux en faire des fonctions de classe non statiques
    # Les anciennes mléthodes sont ci-après
    # ==========================
    def convertit_types_colonnes_df(self):
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
        match self._typeExport:
            case "Sessions":
                self.df["Trigramme formation"] = self.df["Trigramme formation"].astype(str)  # Retype "Trigramme formation"
                self.df["Code IRIS"] = self.df["Code IRIS"].astype(str)  # Retype "Code IRIS"

            case "Ventes":
                # On convertit la colonne "Date de début" en datetime
                self.df['Date de début'] = pd.to_datetime(self.df['Date de début'])


    # ==========================
    # === Méthodes statiques pour ouvir un extract existant IRIS Session (excel VTE qui concatène plusieurs natifs)  ===
    # Anciennes méthodes avant que je ne mette _fe en variable d'instance.
    # TODO : à transférer à priori s'il reste des anciens appels
    # TODO à vérifier : est-ce que ces fonctions n'étaient utiles que pour IRIS_traite (avant que je ne la fasse) → Soit supprimer, soit adapter pour IRIS non traités
    # ==========================
    @staticmethod
    def choisir_fichier_IRIS_sessions() -> Path | None:
        """
        Ouvre un filedialog pour demander à l'utilisateur de sélectionner un extract IRIS sessions.
        On pointe au mieux sur le répertoire des extracts IRIS pour la boîte de dialogue.
        """

        return choisir_fichier(titre=f"Sélectionner l'extract IRIS session {IRIS.cei("Sessions")._codeExport} à employer.",
                        types_fichiers=[("Fichiers Excel", "*.xlsx")],
                        dossier_initial=IRIS.cei("Sessions")._output.repertoire, # Pour aller vers mes fichiers concaténés, sinon pour les originaux il faut pointer vers input
                        texte_bouton_choisir=f"Choisir extract IRIS session {IRIS.cei("Sessions")._codeExport} à nouveau"
                        )

    @staticmethod
    def charger_excel_IRIS_sessions(fe_IRIS_sessions:Optional[FichierExcel]=None, chemin_IRIS_sessions:Optional[Path]=None) -> FichierExcel:
        """
            Retourne l’extract IRIS sessions (excel VTE qui concatène plusieurs natifs) s'il n'existe pas déjà.
            Soit on fournit un FichierExcel, soit un chemin vers ce fichier .xlsx

            Si fe_IRIS_sessions existe (not None), alors on ne fait rien.
            Si chemin_IRIS_sessions est vide, alors on demande à l'utilisateur de pointer un fichier.
            On crée un FichierExcel depuis le chemin.
    
            :param fe_IRIS_sessions: FichierExcel de l'extract IRIS sessions qu'on souhaite traiter
            :type fe_IRIS_sessions: FichierExcel
            :param chemin_IRIS_sessions: Chemin de l'extract IRIS sessions qu'on souhaite traiter
            :type chemin_IRIS_sessions: Path
            :return: un FichierExcel de l'extract IRIS Sessions
            :rtype: FichierExcel
    
            :Example:
    
            >>> charger_excel_IRIS_sessions(chemin_IRIS_sessions=chemin_input)
            >>> charger_excel_IRIS_sessions()

    
            .. seealso:: Rien du tout.
            .. warning:: Rien du tout.
            .. note:: Rien du tout.
            .. todo:: Rien du tout.
        """
        
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
    def convertit_types_colonnes_df_sessions(fe_IRIS_sessions:FichierExcel) -> pd.DataFrame:
        """
            Convertit les types suivants de l'extract Session de IRIS :
               - trigrammes ;
               - code IRIS.

            On le fait car l'import du df convertit en entiers ou en d'autres types certaines colonnes alors que ça ne devrait pas (trigrammes, code IRIS).
            Il en résulte que des filtres ne fonctionnement pas sans conversion.
    
            :param fe_IRIS_sessions: FichierExcel de l'extract IRIS sessions qu'on souhaite traiter
            :type fe_IRIS_sessions: FichierExcel
            :return: un DataFrame convertit sur les colonnes
            :rtype: DataFrame
    
            :Example:

            >>> convertit_types_colonnes_df_sessions(fe_IRIS_sessions = fe)

    
            .. seealso:: Rien du tout.
            .. warning:: Rien du tout.
            .. note:: Rien du tout.
            .. todo:: Rien du tout.
        """

        df_IRIS_sessions = fe_IRIS_sessions._tableaux["Sessions"]._df  # Alias

        # On convertit
        df_IRIS_sessions["Trigramme formation"] = df_IRIS_sessions["Trigramme formation"].astype(str)  # Retype "Trigramme formation"
        df_IRIS_sessions["Code IRIS"] = df_IRIS_sessions["Code IRIS"].astype(str)  # Retype "Code IRIS"

        return df_IRIS_sessions





    # ==========================
    # === Méthodes statiques de traitement d'infos ===
    # ==========================
    @staticmethod
    def cei(typeExport:str) -> ConfigExportIRIS:
        return IRIS._dict_exports_IRIS["ConfigExportIRIS"][typeExport]
    
    @staticmethod
    def extraire_infos_numSessionIRIS(numSession:str) -> pd.Series:
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
    def extraire_code_IRIS_depuis_chemin(chemin:Path|str) -> str:
        """
        Récupère le numéro IRIS (5 chiffres) depuis un chemin (a priori chemin CSV) si pas possible on demande le code à l'utilisateur
        """
        match = re.search(r"\b\d{5}\b", str(chemin))
        if match:
            codeIRIS = match.group(0)
        else:
            codeIRIS = str(demander_code(typeCode="Code IRIS", chemin=chemin))

        return codeIRIS



    # ==========================
    # === IHM ===
    # ==========================
    @staticmethod
    def demander_liste_codes_IRIS(message="Pour exclure des sessions : entrez un ou plusieurs code IRIS (numéro à 5 chiffres) séparés par des espaces ou des virgules (ou rien pour passer) : ") -> list[str]:
        while True:
            entree = input(message).strip()
            if not entree:
                # Pas de saisie => retourner liste vide
                return []

            # On met mes codes IRIS dans une liste de str en vrifiant que toutes les valeurs sont des entiers
            try:
                # On teste le typage en int
                l_entiers = [int(v.strip()) for v in entree.replace(',', ' ').split()]
                # On reconvertit en str avant sortie méthode
                l_str = [str(v) for v in l_entiers]
                return l_str

            except ValueError:
                print("Erreur : veuillez entrer uniquement des nombres entiers, séparés par des espaces ou des virgules.")
                IRIS.demander_liste_codes_IRIS()
        
    @staticmethod
    def choisir_fichier(typeExport:str) -> Path | List[Path]:
        return choisir_fichier(
            titre=f"Sélectionner un ou plusieurs fichiers Extract IRIS {IRIS.cei(typeExport)._nom_typeExport} ({IRIS.cei(typeExport)._codeExport})",
            types_fichiers=[("Fichiers Excel", "*.xlsx")],
            dossier_initial=IRIS.cei(typeExport)._input.repertoire,
            multi_fichiers=True
        )


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
        return self._dict_exports_IRIS["ConfigExportIRIS"][self._typeExport]

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
    def df(self) -> pd.DataFrame:
        """
        DataFrame de _fe._tableaux[self._typeExport]._df (ex. : _fe._tableaux["Sessions"]._df)
        
        :return: DataFrame de _fe._tableaux[self._typeExport]._df (ex. : _fe._tableaux["Sessions"]._df)
        :rtype: DataFrame
        """
        return self._fe._tableaux[self._typeExport]._df

    @df.setter
    def df(self, valeur:pd.DataFrame):
        """
        Remplace le DataFrame de self._fe._tableaux[self._typeExport]._df par un nouveau DataFrame
        
        :param valeur: DataFrame devant écraser l'ancien dans self._fe._tableaux[self._typeExport]._df
        :type valeur: pd.DataFrame
        """
        self._fe._tableaux[self._typeExport]._df = valeur

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

class IRIS_traite(IRIS):
    """
    Classe qui permet d'ouvrir et traiter un Extract IRIS déjà traité
    """
    def __init__(self, typeExport:str, chemin:Optional[Path]=None, fe:Optional[FichierExcel]=None):
        # On initialise la classe mère
        super().__init__(typeExport=typeExport)

        # Type d'export traité
        self.typeExport = typeExport

        # TODO : pas sûr d'en avoir besoin : redondant avec fe Chemin du fichier Excel IRIS
        """
        self._chemin:Optional[Path]=None
        if chemin:
            self._chemin = chemin"""

        # Fichier Excel IRIS déjà traité
        if fe:
            self._fe = fe

        # On ouvre le fichier IRIS s'il n'existe pas encore
        self.charger_excel(fe=self._fe, chemin=chemin)

    def charger_excel(self, fe:Optional[FichierExcel]=None, chemin:Optional[Path]=None):
        """
            Charge dans _fe l’extract IRIS traité (excel VTE qui concatène plusieurs natifs) s'il n'existe pas déjà.
            Soit on fournit un FichierExcel, soit un chemin vers ce fichier .xlsx

            Si self._fe existe (not None), alors on ne fait rien.
            Si chemin est vide, alors on demande à l'utilisateur de pointer un fichier.
            On crée un FichierExcel depuis le chemin et on le met dans _fe.
    
            :param fe: FichierExcel de l'extract IRIS que l'on souhaite traiter
            :type fe: FichierExcel
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
        if fe is None:
            timer.debut(f"Lecture fichier IRIS {self._typeExport.lower()}")

            # On remplace la valeur par défaut si renseigné par l'utilisateur
            if chemin is not None:
                self.chemin = chemin
                
            # S'il n'y a pas de chemin, alors l'utilisateur le pointe
            if self.chemin is None:
                self.chemin = self.choisir_fichier()

            # On ouvre le fichier IRIS s'il n'existe pas encore
            self._fe = FichierExcel.depuis_fichier(chemin_fichier=self.chemin)

            # Conversions des colonnes si besoin (l'import du df convertit en entiers ou en d'autres types certaines colonnes alors que ça ne devrait pas (ex. : trigrammes, code IRIS))
            self.convertit_types_colonnes_df()

            # On trie # TODO : pour l'instant ça marche avec tous mes types (sessions et ventes ; j'ai pas testé avec les autres)
            self.df = self.df.sort_values(by="Date début ses.")  # Trie par "Date début ses."

            timer.fin()

    def choisir_fichier(self) -> Path | None:
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

        return choisir_fichier(titre=f"Sélectionner l'extract IRIS {str.lower(self._typeExport)} {self.cei(self._typeExport)._codeExport} à employer.",
                    types_fichiers=[("Fichiers Excel", "*.xlsx")],
                    dossier_initial=self.cei(self._typeExport)._output.repertoire, # Pour aller vers mes fichiers concaténés, sinon pour les originaux il faut pointer vers input
                    texte_bouton_choisir=f"Choisir extract IRIS {str.lower(self._typeExport)} {self.cei(self._typeExport)._codeExport} à nouveau"
                    )

    def demande_sessions_a_exclure(self, trigramme_formation:str = None, annee:Optional[int] = None, periode:Optional[str] = None) -> Tuple[pd.DataFrame, List[int]]:
        """
        Demande à l'utilisateur les sessions qu'il souhaite exclure de la période choisie :
           - on applique un filtre sur une période + trigramme au dataframe de l'extract IRIS
           - on affiche le résultat
           - l'utilisateur définit les codes IRIS qu'il veut exclure
        
        On retourne :
           - un dataframe df_sessions_filtre à jour
           - #la liste des codes IRIS retenus (plus maintenant : pour l'avoir on peut faire df_filtre["Code IRIS"].tolist())
           - la liste des codes IRIS exclus
        
        :param trigramme_formation: trigramme de la formation à filtrer
        :type trigramme_formation: str
        :param annee: année à filtrer
        :type annee: Optional[int]
        :param periode: période à filtrer
        :type periode: Optional[str]
        :return: un dataframe df_sessions_filtre à jour, la liste des codes IRIS exclus
        :rtype: Tuple[pd.Dataframe, List[int]]
        """

        # Liste des sessions qui seront exclues pâr l'utilisateur
        codes_sessions_exclues_par_utilisateur = []

        # Renseigne les valeurs par défaut de période et année si None
        annee, periode = periode(annee, periode)

        # On filtre sur le trigramme et la période demandée
        df_filtre = self.df_filtre_periode(trigramme_formation=trigramme_formation, annee=annee, periode=periode)

        # On affiche à l'utilisateur les sessions et dates et statuts 
        vlog.print("Info", f"\nListe des sessions {trigramme_formation} dans {self.chemin.name} - {periode} {annee}", style=["jaune"])

        # Adaptation format date
        df_filtre['Date début ses.'] = pd.to_datetime(df_filtre['Date début ses.']).dt.strftime("%d/%m/%Y")
        df_filtre['Date fin ses.'] = pd.to_datetime(df_filtre['Date fin ses.']).dt.strftime("%d/%m/%Y")

        print(tabulate(
            df_filtre[['Code IRIS', 'Trigramme RP', 'Trigramme AF', 'Date début ses.', 'Date fin ses.', 'Durée réal. (J.)', 'Nb. Nommés', 'Statut Session', 'N° Session']], 
            headers='keys', 
            tablefmt='pretty', 
            showindex=False
        ))
        
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
        print(tabulate(
            df_filtre[['Code IRIS', 'Trigramme RP', 'Trigramme AF', 'Date début ses.', 'Date fin ses.', 'Durée réal. (J.)', 'Nb. Nommés', 'Statut Session', 'N° Session']], 
            headers='keys', 
            tablefmt='pretty', 
            showindex=False
        ))


        # pour avoir la liste des codes IRIS retenus : df_filtre["Code IRIS"].tolist()

        # On retourne les codes_IRIS des sessions retenues + exploitationBilan 
        return df_filtre, codes_sessions_exclues_par_utilisateur

