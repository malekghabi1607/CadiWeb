

# ======================================================================================
# CLASSE FDC
# ======================================================================================
class FdC:
    """
    Classe qui gère tous les éléments relatifs aux fiches de coûts INSTN
    """

    
    # === VARIABLES DE CLASSE COMMUNES À TOUTES LES INSTANCES ===




    # ==================================================================================
    # CONSTRUCTEUR
    # ==================================================================================
    def __init__(self, chemin_fdc: Optional[Path] = None, trigramme_formation: Optional[str] = None) -> None:
        """
        Initialise une instance FdC.

        Args:
            codeIRIS (str | None): Code IRIS de la formation.
            chemin_fdc (Path | None): Chemin vers le fichier Excel de la fiche de coûts.
        """
        # Variables propres à la FdC
        self._trigramme_formation: Optional[str] = trigramme_formation
        self._chemin_fdc: Optional[Path] = chemin_fdc  # Chemin de la fiche de coûts

        self._fe_fdc: Optional[FichierExcel] = None  # Objet Excel contenant la fiche de coûts
        self._tableau_fdc:Optional[FichierExcel._TableauExcel] = None


        # Si aucun fichier input n'est donné, alors on ouvre un filedialog
        if self._chemin_fdc is None:
            self._chemin_fdc = FdC.choisir_fdc(self._trigramme_formation)
            #TODO : else : si j'ai le chemin, je peux récupérer le trigramme

        # On ouvre l'Excel
        self._fe_fdc, self._tableau_fdc = FdC.charger_excel_IRIS_sessions(fe_fdc=self._fe_fdc, chemin_fdc=self._chemin_fdc)


        """ # TODO - Date de modif de la FdC
        # Pour connaître la date de la fiche de coût (basé sur date de modif)
        self._dateFdC = datetime.fromtimestamp(self._chemin_fdc.stat().st_mtime)
        self._sDateFdC = self._dateFdC.strftime("%d/%m/%Y")
        """
        

    @property
    def nomFormation(self) -> str:
        """
        Renvoie le nom de la formation (C5)
        
        :return: le dataframe de l'Exctract IRIS des sessions
        :rtype: str
        """
        return self._tableau_fdc["C5"]

    @property
    def nb_participants_prevus(self) -> int:
        return self._tableau_fdc["C17"]

    @property
    def date_creationFormation(self) -> int:
        """
        Retourne l'année de conception de la formation (C9).
        TODO : on retourne une année alors qu'on pourrait retourner un DateTime
        
        :return: l'année de conception de la formation
        :rtype: int
        """
        dateCreationFormation = self._tableau_fdc["C9"]

        if isinstance(dateCreationFormation, int):
            dateCreationFormation = dateCreationFormation
        elif hasattr(dateCreationFormation, 'year'):
            dateCreationFormation = dateCreationFormation.year
        else:
            print(f"Valeur inattendue pour une année : {dateCreationFormation} (type {type(dateCreationFormation)})")
            dateCreationFormation = 1900
        return dateCreationFormation


        """
        Renvoie le nom de la formation (C5)
        
        :return: le dataframe de l'Exctract IRIS des sessions
        :rtype: str
        """
        return self.fdc_ref("C5")

    @property
    def min_participants_cea(self) -> int:
        return self._tableau_fdc["M22"]
    
    @property
    def min_participants_ee(self) -> int:
        return self._tableau_fdc["N23"]
    
    @property
    def min_participants(self) -> str:
        return f"{self.min_participants_ee} pers. (EE) / {self.min_participants_cea} pers. (CEA)"
    
    @property
    def prevus_participants(self) -> int:
        return self._tableau_fdc["C17"]

    def max_participants(self, depassementAutorise:int) -> int:
        return self.prevus_participants + depassementAutorise


    @staticmethod
    def choisir_fdc(trigramme_formation:str = "") -> Path | None:
        """
        Ouvre un filedialog pour demander à l'utilisateur de sélectionner une fiche de coûts.
        On pointe au mieux sur le répertoire des FdC pour la boîte de dialogue.
        """

        return choisir_fichier(titre=f"Sélectionner la fiche de coûts à employer.",
                        types_fichiers=[("Fichiers Excel", "*.xlsx")],
                        dossier_initial=config.format_path(config.REPERTOIRE_FDC, trigramme_formation=trigramme_formation),
                        texte_bouton_choisir=f"Choisir FdC à nouveau"
                        )


    @staticmethod
    def charger_excel_IRIS_sessions(fe_fdc:Optional[FichierExcel]=None, chemin_fdc:Optional[Path]=None) -> Tuple[FichierExcel, FichierExcel._TableauExcel]:
        """
            Retourne l’extract IRIS sessions s'il n'existe pas déjà.
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
        if fe_fdc is None:
            timer.debut("Lecture fiche de coûts")

            if chemin_fdc is None:
                chemin_fdc = FdC.choisir_fdc()

            fe_fdc = FichierExcel.depuis_fichier(chemin_fichier=chemin_fdc, nom_onglet="Fiche de coûts")
            tableau_fdc = fe_fdc._tableaux["Fiche de coûts"]
            
            #fe_fdc._tableaux["Fiche de coûts"]._df = IRIS.convertit_types_colonnes_df_sessions(fe_IRIS_sessions)

            timer.fin()

        return fe_fdc, tableau_fdc

    def lire_fdc_bak(self):

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



        """
        # On charge le fichier fdc si non déjà ouvert
        self._fe_IRIS_sessions = IRIS.charger_excel_IRIS_sessions(fe_IRIS_sessions=self._fe_IRIS_sessions, chemin_IRIS_sessions=chemin_IRIS_sessions)



        self._chemin_fdc = filedialog.askopenfilename(title="Sélectionner la dernière fiche de coûts", filetype=[("Fichiers Excel", "*.xlsx")], initialdir=optimiseCheminRepertoire(self._repertoire_fdc_defaut.replace("XXX", self._codeFormation)))
        if not chemin_fichier_session:
            log_erreur("click sur cancel du filedialog → Pas de chemin de fiche de coûts")
        self._dateFdC = datetime.fromtimestamp(self._chemin_fdc.stat().st_mtime)
        self._sDateFdC = self._dateFdC.strftime("%d/%m/%Y")
        
        self._df_fdc_infos, self._df_fdc_couts, self._prixVenteRetenuParParticipant, self._dateCreationFormation, self._dureeJours_fdc, self._osThematique, self._nbCible_fcd = lire_fdc(self._chemin_fdc)
        """

        nomOnglet_fdc = "Fiche de coûts"

        # Premier tableau : B5:C21 (informations génériques)
        df_fdc_infos = pd.read_excel(self._chemin_fdc, sheet_name=nomOnglet_fdc, usecols=[1, 2], names=["Critere", "Valeur"], header=3, nrows=17)
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
        prixVenteRetenuParParticipant = pd.read_excel(self._chemin_fdc, sheet_name=nomOnglet_fdc, usecols=[9], names=["Valeur"], skiprows=20, nrows=1).iloc[0,0]
        
        # Deuxième tableau : tableau des coûts (tout compris) et des prix par personne (T1, T2 et T3) : ref K35:N35
        df_fdc_couts = pd.read_excel(self._chemin_fdc, sheet_name=nomOnglet_fdc, usecols=[10, 11, 12, 13, 14, 15], names=["T1", "T3", "T2", "0.9xT3", "1.1xT3", "Valeur fixée"], header=33, nrows=2).transpose() #Grosse astuce : je mets 2 lignes de plus pour affecter les noms plus facilement et je les recalculerai après
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


