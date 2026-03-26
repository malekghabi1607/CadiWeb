from __future__ import annotations

from vte.core import config
from vte.utils.utils import *

from pathlib import Path
import re
from typing import Optional




### --------------------------------------------------------------------
#  Fonctions utilitaires globales INSTN
### --------------------------------------------------------------------

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
            vlog.log_erreur(f"cas non valide : soit 'IRIS' soit 'Trigramme formation', demandé : {typeCode} → exit()")
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
        vlog.log_erreur(f"{typeCode} non renseigné → exit()")
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

def construire_chemin_config(
    chemin_a_completer: Path,
    *,
    optimiser_chemin: bool = True,
    convertir_unc: bool = True,
    #fallback_path: Optional[Path] = config.REPERTOIRE_FORMATION.parent,
    **kwargs: Any
) -> Path:
    """
    Construit un chemin à partir d'un template de configuration (config.format_path) :
        - les paramètres à compléter du chemin peuvent être rentrés de manière générique ;
        - le chemin est transformé en unc (s'il y a un raccourci lecteur réseau sur le poste de l'utilisateur on transforme en chemin réseau complet) ;
        - l'utilisateur peut optimiser le chemin (chemin le plus long entre l'attendu et ce qui existe).


    Fonction générique utilisée pour tous les chemins métiers (CSV, Word, Excel, etc.) qui sont dans la conf.

    :param chemin_a_completer: Template de chemin provenant de la config avec éléments à compléter (ex. config.REPERTOIRE_FORMATION)
    :type chemin_a_completer: Path

    :param optimiser_chemin: Si True, applique optimiseCheminRepertoire pour trouver le chemin le plus long entre l'attendu et ce qui existe. Défaut = True
    :type optimiser_chemin: bool

    :param convertir_unc: Si True, convertit le chemin en UNC. Défaut = True
    :type convertir_unc: bool

    :param kwargs: Paramètres à injecter dans config.format_path (trigramme_formation, annee, etc.)
    :type kwargs: dict

    :return: Chemin construit et prêt à l'emploi
    :rtype: Path
    """
    """
    :param fallback_path: Chemin utilisé si les paramètres nécessaires ne sont pas fournis
    :type fallback_path: Optional[Path]

    if kwargs:
        chemin = config.format_path(chemin_a_completer, **kwargs)
    else:
        if fallback_path is None:
            raise ValueError("Aucun paramètre fourni et aucun fallback_path défini")
        chemin = fallback_path
    """
    # --- Filtrer les kwargs None ---
    kwargs = {k: v for k, v in kwargs.items() if v is not None}
    
    # --- Construction du chemin ---
    chemin = config.format_path(chemin_a_completer, **kwargs)

    # --- Conversion UNC ---
    if convertir_unc:
        chemin = chemin_vers_unc(chemin)

    # --- Optimisation chemin ssi le chemin est un répertoire ---
    if chemin.is_dir() and optimiser_chemin:
        chemin = optimiseCheminRepertoire(chemin)

    return chemin


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
            vlog.log_erreur(f"cas non valide : soit 'IRIS' soit 'Trigramme formation', demandé : {typeCode} → exit()")
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
            vlog.log_erreur("Erreur", f"Impossible de renommer le fichier :\n{e}")
            return  # Ne pas fermer la fenêtre si erreur

        fenetre.destroy()

    def annuler():
        fenetre.destroy()
        vlog.log_erreur(f"{typeCode} non renseigné pour {chemin} → exit()")
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





