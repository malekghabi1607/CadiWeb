

### --------------------------------------------------------------------
#  Fonctions globales INSTN
### --------------------------------------------------------------------

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

def demander_code(typeCode:str, info:Optional[Path] = None) -> int|str:
    """
    Permet de demander un code à l'utilisateur : soit code IRIS, soit un trigramme formation

    :Exemples:
       demander_code(typeCode="Code IRIS")
       demander_code(typeCode="Trigramme formation")
    """
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
        match typeCode:
            case "Code IRIS":
                bouton_valider.config(state="normal" if val.isdigit() and len(val) == nbCaracteres else "disabled")
            case "Trigramme formation":
                bouton_valider.config(state="normal" if len(val) == nbCaracteres else "disabled")
        

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

    if info:
        label_chemin = tk.Label(
            fenetre,
            text=f"Donnée :\n  * {info} *",
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




# Versions avec "pointeur"
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

# Requis pour avoir des ConfExportIRIS dans config.py (dinon références circulaires à l'import)
#initialiser_fichierConfig_ExportIRIS_standards()

# Pour couleur barres de progression
#colorama.init(autoreset=True)
# forcer la conversion ANSI dans toutes les consoles
colorama.init(autoreset=True, convert=True, strip=False)

# Pour chrono des fonctions
timer = Timer()

# Pour message de sortie applis externes
vlog = Vlog()