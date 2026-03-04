
# ==========================================================================================
# CLASSE EVALSTAT_SERVICES
#
# Pour lancer des fonctions et méthodes faisant appel à EvalStat et classes fonctionnelles mères (ex. : Formation)
# ==========================================================================================
class EvalStat_services:
    """
    Services métier autour des exports IRIS.
    """
    @classmethod
    def depuis_tuple_csv_stagiaires(cls, tuple_csv_stagiaires:Tuple[Path|str], ouvrirDossier:bool=False) -> None:
        """
        A partir d'un tuple de chemins de CSV stagiaire (il peut il y avoir plusieurs trigrammes de formations différents)
        Permet de générer :
           - le fichier excel stagiaires de chaque session (via le CSV)
           - le fichier excel stagiaires de chaque formation (celui qui concatène tous les CSV d'une session) [Il est créé ou on l'append avec les nouvelles valeurs]
        
        On exclue du traitement les chemin_csv_session qui sont déjà dans le FE formation (on considère que le CSV a déjà été traité)
        """

        # On convertit le tuple de strings en dictionnaire avec les trigrammes formation en clef
        dico_chemins_csv_session = defaultdict(list)  #Dictionnaire spécial : lorsqu’on accède à une clé qui n’existe pas encore, il va automatiquement créer une nouvelle entrée avec une valeur par défaut, ici une liste vide (list())
        for chemin in tuple_csv_stagiaires:
            chemin = Path(chemin)
            trigramme_formation = recupere_trig_formation_depuis_chemin(chemin)
            dico_chemins_csv_session[trigramme_formation].append(chemin)
        dico_chemins_csv_session = dict(dico_chemins_csv_session)  # Optionnel : conversion en dict normal
        
        EvalStat.depuis_dico_csv_stagiaires(dico_chemins_csv_session=dico_chemins_csv_session, ouvrirDossier=ouvrirDossier)

    @classmethod
    def depuis_dico_csv_stagiaires(cls, dico_chemins_csv_session: dict[str, list[Path]], ouvrirDossier:bool=False) -> Tuple[Optional[FichierExcel], Optional[dict[str, dict[str, str]]]]:
        """
        Traite un ensemble de fichiers CSV groupés par formation.

        Args :
            dico_chemins_csv_session (dict[str, list[Path]]):
                Dictionnaire {trigramme_formation: [liste_de_csv]}.

        Return :
            Une liste de dictionnaires {
                es._codeIRIS,
                    {
                    "fichier": chemin_csv.name,
                    "statut": es._statut_csv
                    }      
                }
            Je pourrai accéder à la valeur par nom_dico[codeIRIS]["statut"]  
        Exemple :
            EvalStat.depuis_tuple_csv_stagiaires({
                "ABC": [Path("R04110.csv"), Path("R04112.csv")],
                "DEF": [Path("R04201.csv")]
            })
        """

        # Cas avec dictionnaire non-vide
        if dico_chemins_csv_session :
            statuts_csv:dict[str, dict[str, str]] = {}
            for trigramme_formation, chemins_csv in dico_chemins_csv_session.items():
                print(f"\n\n{Style.BRIGHT}{Fore.RED}Gestion des formations {trigramme_formation}")
                with Contexte_formation(trigramme_formation) as ctx:
                    for chemin_csv in chemins_csv:
                        #cls.depuis_chemin_csv_evaluations_stagiaires(chemin_csv)
                        es = EvalStat.depuis_chemin_csv_evaluations_stagiaires_avec_contexte(chemin_csv_stagiaires=Path(chemin_csv),
                                                                                ouvrirDossier=ouvrirDossier
                                                                                )
                        statuts_csv[es._codeIRIS] = {
                            "fichier": chemin_csv.name,
                            "statut": es._statut_csv
                            }
                        print(f"Fin traitement : {es._codeIRIS}\t{chemin_csv.name}\t{es._statut_csv}")
            return ctx._fe_evaluations_formation, statuts_csv
        
        # Cas vec dictionnaire vide
        else:
            return None, None


        

    @classmethod
    def depuis_liste_codes_IRIS(cls, liste_codes_IRIS:list[int]|int, fe_IRIS_sessions:FichierExcel=None, ouvrirDossier:bool=False) -> Tuple[Optional[FichierExcel], Optional[dict[str, dict[str, str]]]]:
        """
        A partir d'une liste de codes IRIS (il peut il y avoir plusieurs trigrammes de formations différents)
        Permet de générer :
           - le fichier excel stagiaires de chaque session (via le CSV)
           - le fichier excel stagiaires de chaque formation (celui qui concatène tous les CSV d'une session) [Il est créé ou on l'append avec les nouvelles valeurs]
        
        On exclue du traitement les chemin_csv_session qui sont déjà dans le FE formation (on considère que le CSV a déjà été traité)

        
        Return :
            Une liste de dictionnaires {
                es._codeIRIS,
                    {
                    "fichier": chemin_csv.name,
                    "statut": es._statut_csv
                    }      
                }
            Je pourrai accéder à la valeur par nom_dico[codeIRIS]["statut"]  
        """
        if fe_IRIS_sessions is not None :
            cls._fe_IRIS_sessions = fe_IRIS_sessions
            # Conversion du code IRIS en string pour jointure future
            cls._df_IRIS_sessions = cls._fe_IRIS_sessions._tableaux["Sessions"]._df
            cls._df_IRIS_sessions["Code IRIS"] = cls._df_IRIS_sessions["Code IRIS"].astype(str)
        
        # Dictionnaire de retour
        statuts_csv:Optional[dict[str, dict[str, str]]] = {}

        # Si en entrée on a un entier, alors on convertit en liste
        if isinstance(liste_codes_IRIS, int):
            liste_codes_IRIS = [liste_codes_IRIS]

        # On lit le fichier extract IRIS
        cls._charger_IRIS_sessions()


        


        # On crée le dictionnaire des csv stagiaires
        dico_chemins_csv_session = defaultdict(list)  #Dictionnaire spécial : lorsqu’on accède à une clé qui n’existe pas encore, il va automatiquement créer une nouvelle entrée avec une valeur par défaut, ici une liste vide (list())
        for code_IRIS in liste_codes_IRIS:
            #vlog.print("Info", f"Recherche infos code IRIS {code_IRIS}")

            # On récupère le trigramme formation depuis le code IRIS (extract IRIS sessions)
            #res = cls._df_IRIS_sessions.loc[cls._df_IRIS_sessions["Code IRIS"] == str(code_IRIS), "Trigramme formation"]
            #trigramme_formation = res.iloc[0] if not res.empty else demander_code(typeCode="Trigramme formation", info=code_IRIS)
            trigramme_formation = rechercheX_dataframe(
                dataframe=cls._df_IRIS_sessions, 
                colonne_recherche="Code IRIS", 
                valeur_recherche=str(code_IRIS), 
                colonne_souhaitee="Trigramme formation", 
                fallback=lambda: demander_code(typeCode="Trigramme formation", info=code_IRIS)
            )

            # On récupère le CSV depuis le fichier Excel global de la formation
            with Contexte_formation(trigramme_formation):
                #res = cls._df_evaluations_formation.loc[cls._df_evaluations_formation["Code IRIS"] == str(code_IRIS), "Chemin fichier CSV"]
                #if not res.empty:
                #    chemin_csv = res.iloc[0]
                #else:
                #    # Il faut que l'utilisateur pointe le CSV de la session correspondante
                #    chemin_csv = cls._filedialog_csv(code_IRIS=code_IRIS, trigramme_formation=trigramme_formation)
                chemin_csv = rechercheX_dataframe(
                    cls._df_evaluations_formation, 
                    "Code IRIS", 
                    str(code_IRIS), 
                    "Chemin fichier CSV",
                    fallback=lambda: cls._filedialog_csv(code_IRIS=code_IRIS, trigramme_formation=trigramme_formation)
                    )

            if chemin_csv is not None:
                chemin_csv = Path(chemin_csv)
                dico_chemins_csv_session[trigramme_formation].append(chemin_csv)
            else:
                statuts_csv[trigramme_formation] = {
                    "fichier": "Fichier non existant",
                    "statut": "Exclu - Fichier non existant"
                    }
        dico_chemins_csv_session = dict(dico_chemins_csv_session)  # Optionnel : conversion en dict normal

        # On lance le traitement des EvalStat 
        fe_evaluations_formation, statuts_csv = EvalStat.depuis_dico_csv_stagiaires(dico_chemins_csv_session=dico_chemins_csv_session, ouvrirDossier=ouvrirDossier)

        return fe_evaluations_formation, statuts_csv