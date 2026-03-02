from __future__ import annotations

from pathlib import Path
from typing import Optional, Tuple

from vte.domain.fichiers.iris import IRIS
from vte.utils.utils import convertir_tuple_path, convertir_tuple_str


class IRISServices:
    """
    Services métier autour des exports IRIS.
    """

    @staticmethod
    def concatener_type_unique(
        typeExport: str,
        chemins_fichiersInput: Optional[str | Path | Tuple[str | Path]] = None,
        chemin_fichier_sauv: Optional[Path] = None
    ) -> None:

        instance = IRIS(typeExport)

        if chemins_fichiersInput is None:
            chemins_fichiersInput = IRIS.choisir_fichier(typeExport)

        if chemin_fichier_sauv is None:
            chemin_fichier_sauv = instance.cei._output.chemin_fichier

        chemins_fichiersInput = convertir_tuple_path(chemins_fichiersInput)

        fe_modele = FichierExcel.depuis_modele(
            chemin_modele=instance.cei._modele.chemin_fichier,
            chemin_fichier_sauv=chemin_fichier_sauv
        )

        df_concat = instance._charger_df_extractsIRIS_originaux(chemins_fichiersInput)
        fe_modele._tableaux[typeExport].ecrit_dataFrame_dans_tableauStructure(
            df_concat, supprimeDonneesEtRemplace=True
        )

        fe_modele.save()
        fe_modele.close()

    @staticmethod
    def concatener_plusieursTypes(cls, typesExports:str|Tuple[str], depuis_config:bool=False) -> None:
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
    @staticmethod
    def concatener_plusieursTypes_depuis_config_extractsIRIS(cls, typesExports:str|Tuple[str]) -> None:
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
    @staticmethod
    def concatener_plusieursTypes_depuis_filedialog(cls, typesExports:str|Tuple[str]) -> None:
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
    # Sera à adapter
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

