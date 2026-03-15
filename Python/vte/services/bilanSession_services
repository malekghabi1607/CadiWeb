#self._codes_IRIS:Optional[List[str]] = []


# TODO : définir à partir de code_IRIS sans avoir à rentrer formation, année... car tout ça on peut avoir de session IRIS et code IRIS
def bilanUnique_parCodeIRIS(cls, code_IRIS:int):
        # On charge IRIS sessions
        instance._iris_sessions = get_iris(typeExport="Sessions", chemin=chemin_IRIS_sessions)
        
        
        # Initialisation données
        #instance._codeFormation = instance._df_sessions_filtre["Trigramme formation"].iloc[0]
        instance._annee = instance._df_sessions_filtre["Année début ses."].iloc[0]
        moisSession = mois_fr_depuis_date(instance._df_sessions_filtre["Date début ses."].iloc[0])
        numSession = instance._df_sessions_filtre["N° Session"].iloc[0]
        instance._periode = f"Session {numSession} uniquement ({moisSession} {instance._annee})"
        #instance._periodeSessionsEvaluees = f"{numSession}"
        instance._periodeSessionsEvaluees = instance._periode

# TODO : définir à partir de période sans avoir besoin de formation
def depuis_periode(cls, codeFormation:str, annee:int, periode:str) -> None:
     pass

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
def plusieursBilans_parPeriode(cls, liste_periodes:list[Tuple[str, int, str]]) -> None:
    """
    Permet de lancer une série de bilans de sessions à partir d'une liste de tuples ex. [("948", 2022, "Année"), ("948", 2023, "1er semestre"), ("948", 2023, "2nd semestre")]
    """
    for codeFormation, annee, periode in liste_periodes:
        cls.bilanUnique_parPeriode(codeFormation, annee, periode)

