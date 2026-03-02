
# ======================================================================================
# CLASSE BilanSession
# ======================================================================================

class BilanSession:
    """
    C'est la classe qui contient tous les éléments de ma formation pour mon bilan de session V3
    """
    
    # TODO : Une fois fini BilanFormation, il faudra que j'adapte cette classe pour employer la nouvelle classe IRIS et ses méthodes associées

    # === VARIABLES PARTAGÉES ENTRE TOUTES LES INSTANCES
    _chemin_excel_IRIS_sessions:Optional[Path] = None
    # Récupération du dernier export (extract IRIS le plus récent dans le répertoire)
    #_chemin_excel_IRIS_sessions:Optional[str] = obtenir_fichier_plus_recent_repertoire(
    #            config.REPERTOIRE_EXCEL_IRIS_SESSIONS,
    #            r"^R04110_Sessions.*"
    #        )  # Chemin vers le fichier Excel qui contient les extracts IRIS Sessions (ou a minima celles de la période en cours)
    _fe_IRIS_sessions:Optional[FichierExcel] = None  # Fichier Excel qui contient les extracts IRIS Sessions (ou a minima celles de la période en cours)
    _df_IRIS_sessions:Optional[pd.DataFrame] = None  # DataFrame de _fe_IRIS_sessions (self._fe_IRIS_sessions._tableaux["Sessions"]._df)


    _CRITERES_A_ENLEVER:list[str] = [  # Critères à ne pas retenir pour le calcul des moyennes < 3
        "Comment avez-vous connu cette formation ?", "Avez-vous d'autres besoins de formation ?", "Commentaires, remarques, suggestions", "Recommanderiez-vous cette formation ?"]


    # === CONSTRUCTEURS ===
    def __init__(self) -> None:
        # Variables d’instance → propres à chaque bilan
        self._chemin_word_bilan_session_output: Optional[Path] = None  # Bilan de session

        self._codeFormation: Optional[str] = None  # Trigramme formation
        self._annee: Optional[int] = None
        self._periode: Optional[str] = None

        self._df_sessions_filtre:pd.DataFrame = None  # dataframe de _fe_IRIS_sessions filtré avec le bon trigramme formation, la bonne année et la bonne période

        self._stats_stagiaires: Optional[dict] = None  # Dictionnaire des stats des CSV
        self._es: Optional[EvalStat] = None  # EvalStat global des évaluations stagiaires de la formation

        self._codes_IRIS:Optional[List[str]] = []
        
        #self._codes_IRIS_communs: list[str] = []  # Codes IRIS en commun entre le fichier Excel des sessions et le fichier Excel global des évaluations stagiaires de la formation
        #self._codes_IRIS_absents_fin: list[str] = []  # Disparités restantes après traitement des CSV manquants entre le fichier Excel des sessions et le fichier Excel global des évaluations stagiaires de la formation

        self._exploitationBilan:dict[list] = {
            "Exploités pour les stats générales" : [],  # Exploités pour stats initiales → Dans _demande_sessions_a_exclure
            "Exploités pour les évaluations (CSV présents)" : [],  # Exploités pour les stats stagiaires → Dans _maj_evalstat_formation
            "Exclus des évaluations (CSV manquants)" : [],   # Exclus des évaluations car CSV stagiaires manquants → Dans _maj_evalstat_formation
            "Exclus des évaluations (problème traitement CSV)" : [],   # Exclus des évaluations car problème au traitement des CSV → Dans _maj_evalstat_formation
            "Exclus des évaluations (CSV vide / aucun retour)" : [],   # Exclus des évaluations car le CSV est vide (i.e. aucun retour d'utilisateur)
            "Exclus entièrement du bilan (exclus par utilisateur)" : [],  # Exclus entièrement du bilan car sessions non réalisées ou mauvais RP (exclus par l'utilisateur) → Dans _demande_sessions_a_exclure
            "Exclus entièrement du bilan (non présent dans IRIS / mauvais code)" : [],  # Exclus entièrement du bilan car sessions non réalisées ou mauvais RP (exclus par l'utilisateur) → Dans _maj_evalstat_formation
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
        est_code_IRIS_valide, instance._code_IRIS = IRIS.verifier_code_IRIS(code_IRIS)

        if not est_code_IRIS_valide:
            vlog.log_erreur(f"Code IRIS en entrée non valide : {instance._code_IRIS} non traité", continuer=True)
            return
    
        # Ouverture / création du dataframe de l'extract IRIS sessions filtré selon le code IRIS en cours
        instance._charger_df_sessions_filtre_selon_codeIRIS()

        # Initialisation données
        instance._codeFormation = instance._df_sessions_filtre["Trigramme formation"].iloc[0]
        instance._annee = instance._df_sessions_filtre["Année début ses."].iloc[0]
        moisSession = mois_fr_depuis_date(instance._df_sessions_filtre["Date début ses."].iloc[0])
        numSession = instance._df_sessions_filtre["N° Session"].iloc[0]
        instance._periode = f"Session {numSession} uniquement ({moisSession} {instance._annee})"
        #instance._periodeSessionsEvaluees = f"{numSession}"
        instance._periodeSessionsEvaluees = instance._periode


        # Chemin du bilan word
        instance._chemin_word_bilan_session_output = config.format_path(config.CHEMIN_WORD_BILAN_SESSION_OUTPUT, trigramme_formation=instance._codeFormation, annee=instance._annee, periode=f"{numSession}", unite=config.UNITE)

        # On teste la pré-existance du bilan Word
        continuer = tester_existance_fichier(instance._chemin_word_bilan_session_output)

        if continuer:
            # On met à jour l'Excel evalstat de la formation si la session demandée par l'utilisateur ne s'y trouve pas
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
            print("❌  Bilan de session déjà existant → Arrêt du traitement du bilan par l'utilisateur")

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


        # Chemin du bilan word
        instance._chemin_word_bilan_session_output = config.format_path(config.CHEMIN_WORD_BILAN_SESSION_OUTPUT, trigramme_formation=instance._codeFormation, annee=instance._annee, periode=instance._periodeSessionsEvaluees, unite=config.UNITE)

        # On teste la pré-existance du bilan Word
        continuer = tester_existance_fichier(instance._chemin_word_bilan_session_output)

        if continuer:
            # Ouverture / création du dataframe de l'extract IRIS sessions filtré selon la période demandée
            instance._charger_df_sessions_filtre_selon_periode()

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
        
        else:
            print("❌  Bilan de session déjà existant → Arrêt du traitement du bilan par l'utilisateur")


    @classmethod
    def plusieursBilans_parPeriode(cls, liste_periodes:list[Tuple[str, int, str]]) -> None:
        """
        Permet de lancer une série de bilans de sessions à partir d'une liste de tuples ex. [("948", 2022, "Année"), ("948", 2023, "1er semestre"), ("948", 2023, "2nd semestre")]
        """
        for codeFormation, annee, periode in liste_periodes:
            cls.bilanUnique_parPeriode(codeFormation, annee, periode)

    # === MÉTHODES INTERNES ===
    @classmethod
    def _charger_IRIS_sessions(cls) -> None:
        """
        Charge le fichier Excel IRIS Sessions si ce n'est pas déjà fait.
        Cette méthode met à jour _fe_IRIS_sessions et _df_sessions.
        """

        # TODO il faudra que je vire ces anciennes méthodes statiques par celles de la classe IRIS : charger_excel_IRIS_VTE
        # Vérifie existance de fe_IRIS sinon on le charge
        cls._fe_IRIS_sessions = IRIS.charger_excel_IRIS_sessions(fe_IRIS_sessions=cls._fe_IRIS_sessions)

        # On convertit les colonnes
        cls._df_IRIS_sessions = IRIS.convertit_types_colonnes_df_sessions(fe_IRIS_sessions=cls._fe_IRIS_sessions)


    @classmethod
    def _charger_IRIS_sessions_OLD(cls) -> None:
        """
        Charge le fichier Excel IRIS Sessions si ce n'est pas déjà fait.
        Cette méthode met à jour _fe_IRIS_sessions et _df_sessions.
        """

        # Vérifie existance de fe_IRIS sinon on le charge
        cls._fe_IRIS_sessions = IRIS.charger_excel_IRIS_sessions(fe_IRIS_sessions=cls._fe_IRIS_sessions)

        # Traitement du DataFrame
        cls._df_IRIS_sessions = cls._fe_IRIS_sessions._tableaux["Sessions"]._df  # Alias
        cls._df_IRIS_sessions = cls._df_IRIS_sessions.sort_values(by="Date début ses.")  # Trie par "Date début ses."
        cls._df_IRIS_sessions["Trigramme formation"] = cls._df_IRIS_sessions["Trigramme formation"].astype(str)  # Retype "Trigramme formation"
        cls._df_IRIS_sessions["Code IRIS"] = cls._df_IRIS_sessions["Code IRIS"].astype(str)  # Retype "Code IRIS"

    def _charger_df_sessions_filtre_selon_codeIRIS(self) -> None:
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
        self._charger_IRIS_sessions()

        # Application du filtre sur la session
        self._df_sessions_filtre = self._df_IRIS_sessions[self._df_IRIS_sessions['Code IRIS'] == self._code_IRIS]
        #print(self._df_sessions_filtre)

        if len(self._df_sessions_filtre) < 1:
            print(self._df_sessions_filtre)
            vlog.log_erreur(f"Le fichier Excel Session ne contient pas ce code IRIS : {self._df_sessions_filtre}")

    def _charger_df_sessions_filtre_selon_periode(self) -> None:
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
        self._charger_IRIS_sessions()


        ####
        # Filtre du dataframe
        ####

        # Application du pré-filtre avec les 3 critères trigramme, statut session et période
        self._df_sessions_filtre = self._df_IRIS_sessions[
            (self._df_IRIS_sessions['Trigramme formation'] == str(self._codeFormation)) &
            (self._df_IRIS_sessions['Année début ses.'] == self._annee) &
            (self._df_IRIS_sessions['Statut Session'] != "Annulée") &
            (self._df_IRIS_sessions['Nb. Présents'] != 0)
        ]
        #print(self._df_IRIS_sessions_filtre)


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

        # Adaptation format date
        self._df_sessions_filtre['Date début ses.'] = pd.to_datetime(self._df_sessions_filtre['Date début ses.']).dt.strftime("%d/%m/%Y")
        self._df_sessions_filtre['Date fin ses.'] = pd.to_datetime(self._df_sessions_filtre['Date fin ses.']).dt.strftime("%d/%m/%Y")

        print(tabulate(
            self._df_sessions_filtre[['Code IRIS', 'Trigramme RP', 'Trigramme AF', 'Date début ses.', 'Date fin ses.', 'Durée réal. (J.)', 'Nb. Présents', 'Statut Session', 'N° Session']], 
            headers='keys', 
            tablefmt='pretty', 
            showindex=False
        ))
        
        # On demande à l'utilisateur les sessions qu'il veut exclure
        exclusionSessions = IRIS.demander_liste_codes_iris()
        if exclusionSessions:  # si la liste n'est pas vide
            # On trace l'exclusion des sessions
            for session_exclue in exclusionSessions:
                self._exploitationBilan["Exclus entièrement du bilan (exclus par utilisateur)"].append(self._df_sessions_filtre.loc[self._df_sessions_filtre["Code IRIS"] == session_exclue, "N° Session"].iloc[0])

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

        # On met à jour _df_sessions_filtre en enlevant les sessions exclues → Plus besoin : on a ça dans _maj_evalstat_formation
        #self._exploitationBilan["Exploités pour les stats générales"] = self._df_sessions_filtre["N° Session"].tolist()

        # On renseigne _codes_IRIS
        self._codes_IRIS = self._df_sessions_filtre["Code IRIS"].tolist()

    def _maj_evalstat_formation(self) -> None:
        """
        Crée les EvalStat de la/les sessions demandées et met à jour le fichier EvalStat de la formation
        Ne s'applique que si des sessions demandées par l'utilisateur ne s'y trouvent pas
        (on regarde les CSV qui ne sont pas dans le fichier Excel global à partir de la liste df_sessions_filtre['Code IRIS'])

        """

        fe_evaluations_formation, statuts_csv = EvalStat.depuis_liste_codes_IRIS(self._codes_IRIS, fe_IRIS_sessions=self._fe_IRIS_sessions)

        # A ce stade, toutes les valeurs de _codes_IRIS sont sensées être a minima présents dans IRIS avec données pour stats générales
        for code_IRIS in self._codes_IRIS:
            self._exploitationBilan["Exploités pour les stats générales"].append(code_IRIS)
        
        # Cas s'il y a au moins un CSV de lu
        if statuts_csv is not None:
            # === Gestion retour du traitement des CSV
            for code_IRIS, donnees in statuts_csv.items():
                if donnees["statut"] in ("Traité", "Exclu - CSV déjà dans fichier global"):
                    self._exploitationBilan["Exploités pour les évaluations (CSV présents)"].append(code_IRIS)


                if donnees["statut"] == "Exclu - Problème lecture CSV":
                    self._exploitationBilan["Exclus des évaluations (problème traitement CSV)"].append(code_IRIS)

                if donnees["statut"] == "Exclu - Fichier non existant":
                    self._exploitationBilan["Exclus des évaluations (CSV manquants)"].append(code_IRIS)

                if donnees["statut"] == "Exclu - CSV vide / Aucun retour":
                    self._exploitationBilan["Exclus des évaluations (CSV vide / aucun retour)"].append(code_IRIS)
                

                if donnees["statut"] == "Exclu - Code IRIS pas dans Extract IRIS sessions":
                    self._exploitationBilan["Exclus entièrement du bilan (non présent dans IRIS / mauvais code)"].append(code_IRIS)
                    # Alors ne pas exploiter pour les stats générales
                    self._exploitationBilan["Exploités pour les stats générales"].remove(code_IRIS)

            # === Préparation des dataframes des évaluation de la formation pour calcul des stats ===
            self._df_stagiaires = fe_evaluations_formation._tableaux["Stagiaires"]._df  # Création d'un alias pour faciliter le code
            
            # df_stagiaires filtré sur les codes IRIS exploités pour le bilan (infos générales)
            self._df_stagiaires_final = self._df_stagiaires[self._df_stagiaires['Code IRIS'].isin(self._exploitationBilan["Exploités pour les stats générales"])]
            #vlog.print("Info", self._df_stagiaires_final)

            # df_stagiaires avec 1 ligne pour chaque session différente (pour avoir les infos globales issues de IRIS comme le nb d'apprenants)
            self._df_stagiaires_final_1ligne_session = self._df_stagiaires_final.drop_duplicates(subset=['Code IRIS'])  # Ne garde qu'une ligne par Code IRIS (la première rencontrée)
            #vlog.print("Info", self._df_stagiaires_final_1ligne_session)

        # Cas s'il n'y a aucun CSV
        else:
            # Si aucun CSV, alors il faut le traiter manuellement 
            for code_IRIS in self._codes_IRIS:
                self._exploitationBilan["Exclus des évaluations (CSV manquants)"].append(code_IRIS)
            
            self._df_stagiaires = None
            self._df_stagiaires_final = None
            self._df_stagiaires_final_1ligne_session = None
            



        # === Afficher les statuts des CSV employés pour le bilan ===
        #pprint(self._exploitationBilan)
        self._commentairesBilan += "\nListe des sessions :"
        for critere, lsessions in self._exploitationBilan.items():
            if lsessions:
                self._commentairesBilan += f"\n   • {critere} :" + "".join(f"\n       - {isession}" for isession in lsessions)
        vlog.print("Info", f"\n{self._commentairesBilan}")      

        



    def _maj_evalstat_formation_BAK(self) -> None:
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
            vlog.print("Info", f"⚠️  Aucun CSV disponible pour ce bilan : les statistiques des stagiaires ne seront pas évaluées.")
            self._commentairesBilan += f"\n⚠️  Aucun CSV disponible pour ce bilan : les statistiques des stagiaires ne seront pas évaluées.\n"



        return self._stats_stagiaires

    def _envoyer_mail_chef_unite(self, pj:Optional[list[str]] = None):
        """
        Envoie un mail au chef d'unité avec en lien le PDF à signer
        """       
        


        chemin_pdf_bilan_output = self._chemin_word_bilan_session_output.with_suffix(".pdf")
        #self._CORPS_MAIL_CHEF_UNITE.replace()
        corps_html = remplacer_champs(config.CORPS_MAIL_CHEF_UNITE, [
            ["lien_pdf_bilan", chemin_pdf_bilan_output],
            ["formation", f"{self._titreFormation} ({self._codeFormation})"],
            ["periode", minuscule_premiere_lettre(self._periodeSessionsEvaluees)],
        ])

        Mail.creer_mail(
            destinataires=config.ADRESSE_MAIL_CHEF_UNITE,
            sujet=f"Signature bilan de session {self._titreFormation} ({self._codeFormation}) : {chemin_pdf_bilan_output.name}",
            corps_html=corps_html,
            pieces_jointes=pj,
            envoyer_mail=False  # envoie directement sans afficher
        ) 

    # === MÉTHODES POUR LA V3 DU BILAN DE SESSION ===
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
        
        # TODO à mettre ailleurs au début
        #self._chemin_word_bilan_session_output = config.format_path(config.CHEMIN_WORD_BILAN_SESSION_OUTPUT, trigramme_formation=self._codeFormation, annee=self._annee, periode=self._periodeSessionsEvaluees, unite=config.UNITE)
        
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
                self._satisfactionGlobale_moy = "Pas de donnée"
            try:
                self._satisfactionGlobale_com = self._stats_stagiaires["Satisfaction globale"]["Commentaires"].replace("_x000D_", "\n")
            except:
                pass
            try:
                self._recommandation_moy = f'{self._stats_stagiaires["Recommanderiez-vous cette formation ?"]["Moyenne"]/5*100:.0f}%'  # (on divise par 5 car on a un booléen stcké sous forme de note sur 5 : 0 = False, 5 = True)
            except:
                self._recommandation_moy = "Pas de donnée"
            try:
                self._commentairesRemarquesSuggestions_com = self._stats_stagiaires["Commentaires, remarques, suggestions"]["Commentaires"].replace("_x000D_", "\n")
            except:
                pass
            try:
                self._evalInf3_val = f"{len(stats_sous_3)}"
            except:
                self._evalInf3_val = "Pas de donnée"
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
                self._tauxRetours_val = "Pas de donnée"

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

    # === PROPRIÉTÉS ===
    @property
    def _code_IRIS(self) -> str:
        """
        Un code IRIS (en considérant qu'on est sur un bilan contenant un code IRIS unique)
        """
        return self._codes_IRIS[0]
    
    @_code_IRIS.setter
    def _code_IRIS(self, valeur:int|str):
        """
        Renseigne un code IRIS (en considérant qu'on est sur un bilan contenant un code IRIS unique)
        """
        if not isinstance(valeur, str|int):
            raise TypeError("La valeur doit être une chaîne ou un entier.")
        
        if isinstance(valeur, int):
            valeur = str(valeur)

        if not self._codes_IRIS: # Cas d'une liste vide
            self._codes_IRIS.append(valeur)
        else:
            self._codes_IRIS[0] = valeur

