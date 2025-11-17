from __future__ import annotations

from typing import Dict, List, Tuple, Optional, Union, Any, Type, Callable
from types import ModuleType

import pandas as pd

from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries, get_column_letter

import xlwings as xw

import inspect
import os
import sys
import platform
import subprocess
import re
import time as time_module
import locale
import copy
import ctypes
from ctypes import wintypes
from pathlib import Path

import time as time_module
from datetime import date, datetime, time, timedelta

from tqdm import tqdm
import colorama
from colorama import Fore, Style

import tkinter as tk
from tkinter import filedialog

from bs4 import BeautifulSoup

import pygetwindow as gw

import importlib.util

import hashlib

### --------------------------------------------------------------------
#  Tests (log et timer)
### --------------------------------------------------------------------

class Timer:
    stack = []        # Pile des timers imbriqués
    indent = "   "
    last_from_timer = False   # True si la dernière ligne imprimée est un ⏳ du timer

    def __init__(self, description=""):
        self.description = description
        self.start = None
        self.depth = 0
        self.line_length = 0
        self.active = False    # Timer encore "ouvert"

    def debut(self, description=""):
        self.description = description
        self.start = time.time()
        self.depth = len(Timer.stack)
        Timer.stack.append(self)
        self.active = True

        indent = Timer.indent * self.depth
        msg = f"{indent}⏳ {self.description}..."

        # Si la dernière ligne venait d’un timer → on remplace
        if Timer.last_from_timer:
            sys.stdout.write("\r" + " " * self.line_length + "\r")
        else:
            # Sinon, on passe à une nouvelle ligne
            sys.stdout.write("\n")

        sys.stdout.write(msg)
        sys.stdout.flush()

        self.line_length = len(msg)
        Timer.last_from_timer = True

    def fin(self):
        if not self.active:
            return  # Déjà terminé

        end = time.time()
        minutes, seconds = divmod(int(end - self.start), 60)
        indent = Timer.indent * self.depth

        final = f"{indent}✅ {self.description} terminé en {minutes} min {seconds} s."

        # Supprime ce timer de la pile
        Timer.stack.remove(self)
        self.active = False

        # Si la dernière ligne venait du timer → on remplace le ⏳
        if Timer.last_from_timer:
            sys.stdout.write("\r" + " " * self.line_length + "\r")
        else:
            sys.stdout.write("\n")

        sys.stdout.write(final)
        sys.stdout.flush()

        # Maintenant la dernière ligne appartient au Timer
        Timer.last_from_timer = True


def timer_safe_print(*args, **kwargs):
    """
    Permet d'imprimer du texte sans casser les timers.
    À utiliser si tu veux un print "compatible Timer".
    """
    Timer.last_from_timer = False
    print(*args, **kwargs)

class Timer_V1:
    stack = []   # pile des timers ouverts
    indent = "   "

    def __init__(self, description=""):
        self.description = description
        self.start = None
        self.depth = 0
        self.line_length = 0

    def debut(self, description=""):
        self.description = description
        self.start = time_module.time()
        self.depth = len(Timer.stack)
        Timer.stack.append(self)

        indent = Timer.indent * self.depth
        msg = f"{indent}⏳ {self.description}..."

        # Timer courant → affiché sur la dernière ligne → remplaçable
        sys.stdout.write(msg + "\n")
        sys.stdout.flush()

        self.line_length = len(msg)

    def fin(self):
        if self not in Timer.stack:
            return

        end = time_module.time()
        minutes, seconds = divmod(int(end - self.start), 60)
        indent = Timer.indent * self.depth

        # Ligne finale à afficher
        final = f"{indent}✅ {self.description} terminé en {minutes} min {seconds} s."

        # Retire cet élément de la pile
        Timer.stack.remove(self)

        # Efface la dernière ligne (le ⏳)
        # → on remplace uniquement le timer courant
        sys.stdout.write("\r" + " " * self.line_length + "\r")
        sys.stdout.flush()

        # On réécrit correctement la ligne finale
        print(final)
        sys.stdout.flush()

class Timer_V0:
    """
    Classe Timer simple pour mesurer et afficher la durée de traitements dans un script.

    Fonctionnalités :
    -----------------
    - Affiche un message au début d'un traitement : "⏳ Traitement de <description>..."
    - Remplace ce message à la fin (automatique ou manuelle) par : "✅ <description> terminé en X min Y s."
    - Gère automatiquement la fin du chrono précédent à chaque nouvel appel de `debut(...)`.

    Utilisation :
    ------------
    >>> timer = Timer()
    >>> timer.debut("Chargement des données")
    >>> # ... traitement ...
    >>> timer.debut("Traitement des résultats")
    >>> # ... autre traitement ...
    >>> timer.fin()  # Optionnel si on veut terminer explicitement le dernier chrono

    Méthodes :
    ----------
    - debut(description: str): démarre un nouveau chronomètre et affiche un message.
                                Termine automatiquement le précédent s'il est en cours.
    - fin(): termine le chronomètre en cours et affiche la durée du traitement.

    Remarques :
    -----------
    - Aucun module externe requis (comme tqdm).
    - L'affichage est propre dans la console grâce à l'effacement dynamique de la ligne.
    - Conçu pour les scripts où l'on veut chronométrer plusieurs étapes sans se répéter.
    """
    def __init__(self, description=""):
        self.__debut = None
        self.__fin = None
        self.__duree = 0
        self.__description = description
        self.__timer_en_cours = False
        self.__last_message = ""

    def debut(self, description=""):
        # Si un timer est déjà en cours, on le termine proprement
        if self.__timer_en_cours:
            self.fin()

        self.__description = description
        self.__debut = time_module.time()
        self.__timer_en_cours = True

        self.__last_message = f"⏳ {self.__description}..."
        print(self.__last_message, end='', flush=True)

    def fin(self):
        if not self.__timer_en_cours:
            return  # Rien à terminer

        self.__fin = time_module.time()
        self.__duree = self.__fin - self.__debut
        minutes, secondes = divmod(int(self.__duree), 60)

        message_final = f"✅ {self.__description} terminé en {minutes} min {secondes} s."

        # Nettoyer la ligne précédente et afficher le nouveau message
        clean_line = '\r' + ' ' * len(self.__last_message) + '\r'
        print(clean_line + message_final)

        self.__timer_en_cours = False

class Vlog:
    """
    Classe de journalisation visuelle (Visual Logger).

    Permet de centraliser les messages d'information, d'erreur, d'exclusion, etc.
    Affichage possible en console, en popup Tkinter, et export vers un fichier.

    Exemple :
        vlog = Vlog()
        vlog.ajouter_message("Infos", "Traitement terminé", style=["vert"])
        vlog.log_erreur("Impossible de lire le fichier")
        print(vlog)
        vlog.afficher_popup("Résultat du traitement")
    """

    def __init__(self):
        self._dict_messages: Dict[str, List[Tuple[str, List[str]]]] = {}

        # Définition des styles disponibles pour le widget Tkinter
        self._styles: Dict[str, Dict] = {
            "normal": {"font": ("TkDefaultFont", 10)},
            "gras": {"font": ("TkDefaultFont", 10, "bold")},
            "italique": {"font": ("TkDefaultFont", 10, "italic")},
            "souligne": {"underline": True},
            "rouge": {"foreground": "#cc0000"},
            "rouge clair": {"foreground": "#ff6666"},
            "bleu": {"foreground": "#0000cc"},
            "bleu clair": {"foreground": "#66b3ff"},
            "vert": {"foreground": "#009933"},
            "vert clair": {"foreground": "#66ff99"},
            "jaune": {"foreground": "#e6b800"},
            "orange": {"foreground": "#ff9933"},
        }

    def ajouter_message(self, categorie: str, texte: str, style: Union[str, List[str]] = "normal") -> None:
        """
        Ajoute un message dans une catégorie avec un style donné.

        :param categorie: Nom de la catégorie (ex. : "Erreurs", "Exclusions")
        :param texte: Contenu du message
        :param style: Style (ou liste de styles) appliqué au texte

        :Example:

        >>> # Cas avec un seul style
        >>> vlog.ajouter_message("CSV traités", "fichier_X.csv", style=["vert"])

        >>> # Cas avec styles combinés
        >>> vlog.ajouter_message("Traitement", "Le fichier a été ajouté avec succès", style=["gras", "vert"])


        """
        if isinstance(style, str):
            style = [style]
        if categorie not in self._dict_messages:
            self._dict_messages[categorie] = []
        self._dict_messages[categorie].append((texte, style))

    def print(self, categorie: str, texte: str, style: Union[str, List[str]] = "normal") -> None:
        self.ajouter_message(categorie=categorie, texte=texte, style=style)
        print(texte)

    def reinitialiser_messages(self) -> None:
        """
        Vide complètement les messages enregistrés.
        """
        self._dict_messages.clear()

    def log_erreur(self, message: str, continuer: bool = False) -> None:
        """
        Affiche une erreur enrichie en console et l'enregistre dans les messages.

        :param message: Le message d'erreur
        :param continuer: Si False, appelle `exit()` après affichage

        :Example:

        >>> vlog.log_erreur("Une erreur est survenue.")
        """
        #import colorama
        #from colorama import Fore, Style
        #colorama.init()
        colorama.init(autoreset=True)

        stack = inspect.stack()
        frame = stack[1].frame
        nom_fonction = stack[1].function
        cls_name = None
        if 'self' in frame.f_locals:
            cls_name = type(frame.f_locals['self']).__name__
        location = f"{cls_name + '.' if cls_name else ''}{nom_fonction}()"

        prefix = f"{Fore.YELLOW}❌ "
        suffix = " → On ignore la règle et on continue." if continuer else ""
        message_console = f"{prefix}Erreur dans {location} : {message}{suffix}{Style.RESET_ALL}"
        print(message_console)

        # Ajout dans le logger visuel
        self.ajouter_message("Erreurs", f"{location} : {message}", style=["gras", "rouge clair"])

        if not continuer:
            print("=== exit() ===")
            exit()

    def copier_dans_presse_papiers(self) -> None:
        """
        Copie le contenu du log (version texte brut) dans le presse-papiers.
        """
        r = tk.Tk()
        r.withdraw()
        r.clipboard_clear()
        r.clipboard_append(str(self))
        r.update()
        r.destroy()

    def sauvegarder_vers_fichier(self, chemin: Optional[str] = None) -> None:
        """
        Sauvegarde le contenu du log dans un fichier texte.

        :param chemin: Chemin du fichier. Si None, un nom par défaut est généré.
        """
        if chemin is None:
            timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            chemin = f"log_{timestamp}.txt"
        with open(chemin, "w", encoding="utf-8") as f:
            f.write(str(self))

    def afficher_popup(self, titre: str = "Information", reinitialiser_messages=True) -> None:
        """
        Affiche une popup enrichie avec les messages du log.

        :param titre: Titre de la fenêtre popup

        :Example:

        >>> vlog.afficher_popup("Rapport d'import")
        """
        popup = tk.Toplevel()
        popup.title(titre)
        popup.resizable(True, True)

        # Calcul de la largeur optimale selon le message le plus long
        largeur_min = 60
        largeur_max = 400
        largeur_calculee = largeur_min

        for categorie, messages in self.dict_messages.items():
            for texte, _ in messages:
                ligne = f"{categorie} :\n\t• {texte}"
                largeur_calculee = max(largeur_calculee, len(ligne))

        # Ajuster pour ne pas dépasser une largeur raisonnable
        print(largeur_calculee)
        largeur_calculee = min(largeur_calculee, largeur_max)
        print(largeur_calculee)

        text_widget = tk.Text(popup, wrap="word", height=25, width=largeur_calculee)
        text_widget.pack(expand=True, fill="both", padx=10, pady=10)

        # Configurer tous les styles
        for style_nom, style_conf in self._styles.items():
            text_widget.tag_configure(style_nom, **style_conf)

        # Insérer les messages formatés
        for categorie, messages in self._dict_messages.items():
            text_widget.insert("end", f"\n{categorie} :\n", ("gras", "souligne"))
            for texte, styles in messages:
                if not styles:
                    styles = ["normal"]
                text_widget.insert("end", "\t• ", tuple(styles))
                text_widget.insert("end", texte + "\n", tuple(styles))

        text_widget.config(state="disabled")

        bouton_frame = tk.Frame(popup)
        bouton_frame.pack(pady=(0, 10))

        tk.Button(bouton_frame, text="Copier dans presse-papiers", command=self.copier_dans_presse_papiers).pack(side="left", padx=5)
        tk.Button(bouton_frame, text="Fermer", command=popup.destroy).pack(side="right", padx=5)

        # Centrer la popup
        popup.update_idletasks()
        w = popup.winfo_width()
        h = popup.winfo_height()
        x = (popup.winfo_screenwidth() // 2) - (w // 2)
        y = (popup.winfo_screenheight() // 2) - (h // 2)
        popup.geometry(f"+{x}+{y}")

        # Si demandé, on réinitialise les messages
        if reinitialiser_messages:
            self.reinitialiser_messages()

    @property
    def dict_messages(self) -> Dict[str, List[Tuple[str, List[str]]]]:
        """
        Retourne le dictionnaire complet des messages.
        """
        return self._dict_messages

    def __str__(self) -> str:
        """
        Affichage texte brut pour la console, structuré par catégorie.
        """
        sortie = ""
        for categorie, messages in self._dict_messages.items():
            sortie += f"{categorie} :\n"
            for texte, _ in messages:
                sortie += f"\t• {texte}\n"
        return sortie

def log_erreur(message: str, continuer:bool = False) -> None:
    """
    Permet de print un log avec la location du message (fonction d'appel) avec option de continuer (warning) ou d'arrêter le code (erreur)
    """
    # Récupérer le frame d'appel (un cran au-dessus dans la stack)
    stack = inspect.stack()
    frame = stack[1].frame

    # Nom de la fonction ou méthode appelante
    nom_fonction = stack[1].function

    # Essayer d'obtenir la classe via le premier argument (souvent self)
    cls_name = None
    if 'self' in frame.f_locals:
        cls_name = type(frame.f_locals['self']).__name__

    # Construction de l'en-tête d'erreur
    prefix = f"{Fore.YELLOW}❌ "
    location = f"{cls_name + '.' if cls_name else ''}{nom_fonction}()"
    

    # Affichage de l'erreur
    
    if continuer:
        suffix = " → On ignore la règle et on continue."
        print(f"{prefix}Erreur dans {location} : {message}{suffix}{Style.RESET_ALL}")
    else:
        print(f"{prefix}Erreur dans {location} : {message}{Style.RESET_ALL}")
        print("=== exit() ===")
        exit()



### --------------------------------------------------------------------
#  Dossiers
### --------------------------------------------------------------------

def optimiseCheminRepertoire(path_in) -> str:
    """
        Tout le monde n'emploie pas les noms de la GED miroir comme ils devraient.
        Ainsi pour faciliter l'utilisateur, je teste l'existance du repertoire qui devrait fonctionner.
        Si ce n'est pas le cas, je remonte d'un cran d'an l'arborescence.
        Je m'arrête si la longueur du chemin complet est nulle (chemin completement bidon). 
 
        :param path_in: Chemin du repertoire a tester
        :type path_in: string
        :return: un chemin optimal (i.e. avec la plus longue arborescence) qui est fonctionnel
        :rtype: string
 
        :Example:
 
        >>> string chemin = optimiseCheminRepertoire("C:\\Users\\fichier.xlsx")

 
        .. seealso:: Rien du tout.
        .. warning:: Rien du tout.
        .. note:: Rien du tout.
        .. todo:: Rien du tout.
    """
    

    path_out = path_in

    # On boucle de manière incrémentale vers la racine du répertoire donné en paramètre d'entrée jusqu'à ce qu'un répertoire soit ok
    while not(os.path.exists(path_out)):
        path_out = path_out[:-len(path_out.split("\\")[-1])-1]
        if len(path_out)==0 :
            path_out = "."
        #print(path_out)

    return(path_out)

def ouvrir_dossier(path) -> str:
    if platform.system() == "Windows":
        os.startfile(os.path.realpath(path))
    elif platform.system() == "Darwin":  # macOS
        subprocess.run(["open", path])
    else:  # Linux
        subprocess.run(["xdg-open", path])

def lister_fichiers_repertoire(
    dossier: str,
    delai: timedelta | None = None,
    inclure_sous_dossiers: bool = False,
    extensions: list[str] | None = None,
    trier_par_date: bool = False
) -> list[str]:
    """
    Liste les fichiers d'un répertoire selon certains critères de date, d'extension,
    et optionnellement les trie par date de modification.

    Args:
        dossier (str): Chemin du répertoire à analyser.
        delai (timedelta | None, optional): Durée limite. 
            Si None, tous les fichiers sont listés.  
            Exemple : `timedelta(days=30)` pour les fichiers modifiés depuis 30 jours.
        inclure_sous_dossiers (bool, optional): 
            Si True, parcourt aussi les sous-dossiers (par défaut False).
        extensions (list[str] | None, optional): 
            Liste d'extensions à filtrer (ex: ['.txt', '.py']). 
            Si None, aucun filtrage d'extension n'est appliqué.
        trier_par_date (bool, optional): 
            Si True, trie les fichiers du plus récent au plus ancien (par défaut False).

    Returns:
        list[str]: Liste des chemins complets des fichiers correspondant aux critères.

    Exemple:
        >>> from datetime import timedelta
        >>> fichiers = lister_fichiers_repertoire(
        ...     dossier="/chemin/vers/dossier",
        ...     delai=timedelta(days=30),
        ...     inclure_sous_dossiers=True,
        ...     extensions=[".py", ".txt"],
        ...     trier_par_date=True
        ... )
        >>> for f in fichiers:
        ...     print(f)
        /chemin/vers/dossier/script.py
        /chemin/vers/dossier/notes.txt
    """
    fichiers = []
    maintenant = time_module.time()
    limite_secondes = delai.total_seconds() if delai is not None else None

    def fichier_valide(chemin: str) -> bool:
        # Vérifie l'extension
        if extensions is not None:
            if not any(chemin.lower().endswith(ext.lower()) for ext in extensions):
                return False
        # Vérifie la date
        if limite_secondes is not None:
            if (maintenant - os.path.getmtime(chemin)) > limite_secondes:
                return False
        return True

    # Parcours du répertoire
    if inclure_sous_dossiers:
        for racine, _, fichiers_local in os.walk(dossier):
            for f in fichiers_local:
                chemin = os.path.join(racine, f)
                if os.path.isfile(chemin) and fichier_valide(chemin):
                    fichiers.append(chemin)
    else:
        for f in os.listdir(dossier):
            chemin = os.path.join(dossier, f)
            if os.path.isfile(chemin) and fichier_valide(chemin):
                fichiers.append(chemin)

    # Tri par date de modification (du plus récent au plus ancien)
    if trier_par_date:
        fichiers.sort(key=lambda f: os.path.getmtime(f), reverse=True)

    return fichiers

def obtenir_fichier_plus_recent_repertoire(repertoire: str, motif: str):
    """
    Retourne le fichier le plus récent d'un répertoire correspondant à une expression régulière.

    Args:
        repertoire (str): Chemin du répertoire à parcourir.
        motif (str): Expression régulière pour filtrer les fichiers.

    Example:
        dernier_log = obtenir_fichier_plus_recent("/var/logs", r"^journal_.*")

    Returns:
        str | None: Chemin complet du fichier le plus récent correspondant, ou None si aucun ne correspond.
    """
    chemin_repertoire = Path(repertoire)
    if not chemin_repertoire.is_dir():
        raise NotADirectoryError(f"{repertoire} n'est pas un répertoire valide.")

    expression = re.compile(motif)

    # Liste des fichiers qui correspondent au motif
    fichiers_correspondants = [
        fichier for fichier in chemin_repertoire.iterdir()
        if fichier.is_file() and expression.search(fichier.name)
    ]

    if not fichiers_correspondants:
        return None

    # Tri par date de création (ou de modification selon le système)
    fichier_plus_recent = max(fichiers_correspondants, key=lambda f: f.stat().st_ctime)
    
    print(f"✅ Récupération de l'extract IRIS Sessions le plus récent : {os.path.basename(fichier_plus_recent)}")

    return str(fichier_plus_recent)


### --------------------------------------------------------------------
#  Réseau
### --------------------------------------------------------------------
def chemin_vers_unc(path:Path|str, retour_type:Type[Path]|Type[str]=Path) -> Path|str:
    """
    Convertit un chemin local (y compris via un lecteur réseau mappé, ex: 'Z:\\...') 
    en chemin réseau complet de type UNC (ex: '\\\\serveur\\partage\\...') sous Windows.

    Si le chemin n'est pas un lecteur réseau ou si la conversion échoue,
    la fonction renvoie le chemin normalisé d'origine.

    Paramètres
    ----------
    path : str | Path
        Chemin à convertir. Peut être une chaîne (`str`) ou un objet `pathlib.Path`.
    retour_type : type, optionnel
        Type de retour souhaité : `Path` (par défaut) ou `str`.

    Retour
    ------
    Path | str
        Le chemin UNC complet sous le type spécifié (`Path` ou `str`).
        Si la conversion échoue, retourne le chemin d'origine converti au bon type.

    Notes
    -----
    - Fonctionne uniquement sous **Windows**.
    - Nécessite l'accès à `mpr.dll` (via ctypes) pour appeler `WNetGetUniversalNameW`.

    Exemples
    --------
    >>> from pathlib import Path
    >>> chemin_local = Path("Z:/projets/formation/stagiaires.csv")

    # Retourne un Path (par défaut)
    >>> chemin_vers_unc(chemin_local)
    WindowsPath('\\\\serveur\\projets\\formation\\stagiaires.csv')

    # Retourne une chaîne de caractères
    >>> chemin_vers_unc(chemin_local, retour_type=str)
    '\\\\serveur\\projets\\formation\\stagiaires.csv'

    # Fonctionne aussi si l'entrée est déjà une str
    >>> chemin_vers_unc("Z:/projets/formation/stagiaires.csv")
    WindowsPath('\\\\serveur\\projets\\formation\\stagiaires.csv')
    """

    # Vérification et conversion du type d'entrée
    if isinstance(path, Path):
        path = str(path)
    elif not isinstance(path, str):
        raise TypeError("L'argument 'path' doit être une chaîne ou un objet Path.")

    # Normalisation du chemin
    path = os.path.normpath(path)

    # Conversion en chemin absolu si nécessaire
    if not os.path.isabs(path):
        path = os.path.abspath(path)

    # Si le chemin ne correspond pas à un lecteur local (ex: 'C:\\'), on renvoie tel quel
    if not path[1:3] == ':\\':
        return retour_type(path)

    # Définition de la structure UNIVERSAL_NAME_INFO
    class UNIVERSAL_NAME_INFO(ctypes.Structure):
        _fields_ = [("lpUniversalName", wintypes.LPWSTR)]

    # Création du buffer pour recevoir le résultat
    buf = ctypes.create_string_buffer(1024)
    size = ctypes.c_ulong(ctypes.sizeof(buf))

    # Appel à l'API Windows pour obtenir le chemin UNC
    result = ctypes.windll.mpr.WNetGetUniversalNameW(
        path,
        0x00000001,  # UNIVERSAL_NAME_INFO_LEVEL
        buf,
        ctypes.byref(size)
    )

    # Si succès
    if result == 0:
        uni_name_info = ctypes.cast(buf, ctypes.POINTER(UNIVERSAL_NAME_INFO)).contents
        unc_path = uni_name_info.lpUniversalName
        return retour_type(unc_path)
    else:
        print(f"Conversion UNC échouée (code erreur : {result}). Chemin renvoyé tel quel.")
        return retour_type(path)

### --------------------------------------------------------------------
#  Divers
### --------------------------------------------------------------------

def trouve_encodage_csv(chemin_fichier: str|Path) -> str:
    """
    Détecte l'encodage d'un fichier CSV.

    Args:
        chemin_fichier (str): Chemin vers le fichier CSV.

    Returns:
        str: L'encodage détecté (ex: 'utf-8', 'cp1252', etc.).
    """
    import chardet

    try:
        with open(chemin_fichier, 'rb') as f:
            result = chardet.detect(f.read())
        return result['encoding']
    except Exception as e:
        print(f"Erreur lors de la détection de l'encodage : {e}")
        return 'utf-8'  # Valeur par défaut en cas d'erreur

def nettoyer_nom_colonne(nom:str) -> str:
    """
    Nettoie un nom de colonne Excel :
    - Remplace les retours à la ligne (\n, \r) par des espaces
    - Supprime les espaces en début et fin de chaîne
    - Conserve les espaces successifs à l'intérieur

    Gestion spécifique des colonnes 'Commentaires' suivies d'espaces (cas des CSV où il y a des colonnes qui feintent le nom des colonnes des tableaux structurés "Commentaires", "Commentaires ", "Commenatires  "...)
    - Si nom commence par 'Commentaires' suivi d'au moins un espace : nom distinct préservé et rendu unique si nécessaire.
    - Sinon : nettoyage standard (retours ligne, espaces multiples, strip).
    """
    
    if not re.match(r'^Commentaires\s+$', nom):  # Match exact "Commentaires" + espaces
        # Supprime les retours à la ligne, espaces au début/fin et caractères spéciaux invisibles
        nom = nom.replace('\n', ' ').replace('\r', ' ')
        nom = re.sub(r'\s+', ' ', nom)  # remplace plusieurs espaces par un seul
        #nom = nom.strip()  # Enlève les espaces doublés

    return nom

def remplacer_champs(
    str_in: str,
    liste_remplacements: list[list[str]] | list[str],
    format_champ: str = "{%s}"
) -> str:
    """
    Remplace dans une chaîne de texte des champs encadrés par des accolades (ou autre format) par leurs valeurs associées.

    Chaque champ à remplacer doit être écrit dans le texte sous la forme définie par `format_champ`.
    Exemple : "Bonjour {nom}, votre formation {formation} est prévue."

    Args:
        str_in (str): 
            Le texte d'entrée contenant des champs à remplacer.
        liste_remplacements (list[list[str]] | list[str]): 
            - Soit une liste de paires [nom_champ, valeur_remplacement].
              Exemple : [["nom", "Dupont"], ["formation", "Python avancé"]]
            - Soit une seule paire [nom_champ, valeur_remplacement] pour un seul remplacement.
              Exemple : ["nom", "Dupont"]
        format_champ (str, optionnel): 
            Format des champs à rechercher. 
            Doit contenir "%s" à l'emplacement du nom du champ.
            Par défaut "{%s}" correspond à "{champ}".
            Exemples valides :
                - "{%s}" → {champ}
                - "<%s>" → <champ>

    Returns:
        str: 
            Le texte de sortie après remplacement de tous les champs.

    Exemples:
        >>> texte = "Bonjour {nom}, votre formation {formation} est prévue le {date}."
        >>> remplacements = [["nom", "Dupont"], ["formation", "Python"], ["date", "15/03/2025"]]
        >>> resultat = remplacer_champs(texte, remplacements)
        >>> print(resultat)
        Bonjour Dupont, votre formation Python est prévue le 15/03/2025.

        >>> texte2 = "Bonjour <nom>, bienvenue dans la formation <formation>."
        >>> remplacements2 = [["nom", "Alice"], ["formation", "Pandas"]]
        >>> resultat2 = remplacer_champs(texte2, remplacements2, format_champ="<%s>")
        >>> print(resultat2)
        Bonjour Alice, bienvenue dans la formation Pandas.

        >>> # Cas d'un seul remplacement
        >>> texte3 = "Bonjour {nom} !"
        >>> resultat3 = remplacer_champs(texte3, ["nom", "Bob"])
        >>> print(resultat3)
        Bonjour Bob !
    """
    # Normalisation : si un seul remplacement est fourni sous forme [clé, valeur]
    if isinstance(liste_remplacements[0], str):
        liste_remplacements = [liste_remplacements]

    str_out = str_in
    for champ, valeur in liste_remplacements:
        champ_formate = format_champ % champ
        str_out = str_out.replace(champ_formate, str(valeur))
    return str_out

def hash_df(df: pd.DataFrame) -> str:
    """
    Retourne une signature unique du DataFrame.
    S’utilise pour détecter les modifications.
    """
    if df is None:
        return "NONE"

    # On convertit en octets de manière stable
    data = pd.util.hash_pandas_object(df, index=True).values

    # On hash le résultat
    return hashlib.md5(data).hexdigest()

### --------------------------------------------------------------------
#  Chargement config
### --------------------------------------------------------------------

def _get_app_dir() -> str:
    """
    Retourne le dossier racine de l'application.
    Gère l'exécution normale et l'exécutable PyInstaller.
    """
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(sys.argv[0]))

def _charger_module_depuis_chemin(path: str, nom_module: str) -> ModuleType:
    """Charge dynamiquement un module Python depuis un fichier."""
    spec = importlib.util.spec_from_file_location(nom_module, path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)  # type: ignore
    return module

def charger_config() -> Tuple[ModuleType, Optional[ModuleType]]:
    """
    Charge les modules de configuration :
    - config.py (global, obligatoire)
    - user_config.py (optionnel)

    Retourne
    -------
    tuple(ModuleType, Optional[ModuleType]) :
        (config, user_config)
    """
    app_dir = _get_app_dir()

    print(f"\n{Style.BRIGHT}{Fore.YELLOW}Chargement des fichiers de configuration")

    # --- config.py (global) ---
    config_path_local = os.path.join(app_dir, "config.py")
    if os.path.exists(config_path_local):
        config = _charger_module_depuis_chemin(config_path_local, "config")
        print(f"✅  Configuration globale config.py chargée depuis {config_path_local}")
    else:
        try:
            import vte.config as config
            print("⚙️  Configuration globale config.py importée depuis vte.config")
        except ModuleNotFoundError:
            raise FileNotFoundError(
                "Impossible de trouver config.py ni dans le dossier local ni dans vte/"
            )

    # --- config_extractsIRIS.py (global optionnel) ---
    config_path_local = os.path.join(app_dir, "config_extractsIRIS.py")
    if os.path.exists(config_path_local):
        config_extractsIRIS = _charger_module_depuis_chemin(config_path_local, "config_extractsIRIS")
        print(f"✅  Configuration globale config_extractsIRIS.py chargée depuis {config_path_local}")
    else:
        try:
            import vte.config_extractsIRIS as config_extractsIRIS
            print("⚙️  Configuration globale config_extractsIRIS.py importée depuis vte.config")
        except ModuleNotFoundError:
            config_extractsIRIS = None
            print("ℹ️ Impossible de trouver config_extractsIRIS.py ni dans le dossier local ni dans vte/ — ce n'est pas bloquant.")


    # --- user_config.py (optionnel) ---
    user_config_path = os.path.join(app_dir, "user_config.py")
    if os.path.exists(user_config_path):
        user_config = _charger_module_depuis_chemin(user_config_path, "user_config")
        print(f"✅  Configuration utilisateur chargée depuis {user_config_path}")
    else:
        user_config = None
        print("ℹ️  Aucun fichier user_config.py trouvé — Ce n'est pas bloquant.")

    print("\n")

    return config, config_extractsIRIS, user_config

def charger_config_user() -> ModuleType:
    """
    Recharge uniquement le module user_config.py.
    """
    app_dir = _get_app_dir()
    user_config_path = os.path.join(app_dir, "user_config.py")

    if not os.path.exists(user_config_path):
        raise FileNotFoundError(
            f"Le fichier de configuration utilisateur est introuvable : {user_config_path}"
        )

    user_config = _charger_module_depuis_chemin(user_config_path, "user_config")
    print(f"✅ Configuration utilisateur rechargée depuis {user_config_path}")
    return user_config

def chargement_config_demander_verif_utilisateur(
    nom_variable: str,
    fonction_execution: Callable[[Any], None],
) -> bool:
    """
    Demande à l'utilisateur de vérifier les données de `user_config.py`
    avant d'exécuter une fonction avec la variable correspondante.

    Paramètres
    ----------
    nom_variable : str
        Nom de la variable définie dans `user_config.py` (ex. : 'liste_codes_IRIS')
    fonction_execution : Callable[[Any], None]
        Fonction à exécuter si l'utilisateur valide (reçoit la variable en argument)

    Retour
    ------
    bool
        True si l'action a été validée et exécutée, False si l'utilisateur a annulé.
    """

    while True:
        # === Étape 1 : Charger la configuration utilisateur ===
        user_conf = charger_config_user()

        if not hasattr(user_conf, nom_variable):
            print(f"⚠️ Le fichier user_config.py ne contient pas la variable '{nom_variable}'.")
            return False

        valeur = getattr(user_conf, nom_variable)

        # === Étape 2 : Afficher le contenu actuel ===
        print("\n📂 Voici la donnée actuellement chargée depuis user_config.py :")
        print(f"\n{nom_variable} = {valeur}\n")

        # === Étape 3 : Demander à l'utilisateur quoi faire ===
        saisie = input(
            "Si vous voulez :\n"
            "  • valider, appuyer sur Entrée\n"
            "  • adapter le fichier puis saisir 1 pour recharger\n"
            "  • annuler la procédure, saisir 0\n"
            "Votre choix : "
        ).strip()

        # === Étape 4 : Gérer le choix ===
        if saisie == "":
            print("\n→ Exécution de l'action...\n")
            fonction_execution(valeur)
            return True

        elif saisie == "1":
            print("\n↻ Rechargement du fichier user_config.py...\n")
            continue  # reboucle après modification du fichier

        elif saisie == "0":
            print("\n❌ Procédure annulée par l'utilisateur.\n")
            return False

        else:
            print("⚠️ Saisie invalide, veuillez appuyer sur Entrée, saisir 1 ou 0.\n")

def charger_config_user_BAK(nom_fichier: str = "user_config.py") -> Optional[ModuleType]:
    """
    Charge dynamiquement un fichier de configuration utilisateur Python.

    Cette fonction recherche le fichier dans :
    1. Le dossier courant (utile pour les tests dans VSCode)
    2. Le dossier contenant l'exécutable (utile après compilation avec PyInstaller)

    Paramètres
    ----------
    nom_fichier : str
        Nom du fichier de configuration utilisateur (par défaut : "user_config.py")

    Retour
    ------
    types.ModuleType | None
        Le module importé s'il existe, sinon None.

    Exemple
    -------
    >>> conf = charger_config_user()
    >>> if conf:
    ...     print(conf.liste_codes_IRIS)
    ...     print(conf.liste_periodes)
    """
    # 1️⃣ - Chemin du fichier selon le mode d'exécution
    if getattr(sys, 'frozen', False):  # Cas EXE (PyInstaller)
        base_dir = os.path.dirname(sys.executable)
    else:  # Cas développement (VSCode, script Python classique)
        base_dir = os.path.dirname(os.path.abspath(__file__))

    chemin_conf = os.path.join(base_dir, nom_fichier)

    # 2️⃣ - Vérification existence
    if not os.path.exists(chemin_conf):
        print(f"⚠️ Fichier de configuration non trouvé : {chemin_conf}. → Exit()")
        #return None
        exit()

    # 3️⃣ - Chargement dynamique du module
    spec = importlib.util.spec_from_file_location("user_config", chemin_conf)
    user_config = importlib.util.module_from_spec(spec)
    assert spec.loader is not None  # pour typer proprement
    spec.loader.exec_module(user_config)

    print(f"✅ Configuration utilisateur chargée depuis : {chemin_conf}")
    return user_config

 

### --------------------------------------------------------------------
#  Conversions
### --------------------------------------------------------------------

def convertir_si_possible(valeur) -> str | int | float:
    """
    Tente de convertir une chaîne en int ou float si c'est pertinent.
    Sinon, renvoie la valeur telle quelle.
    - '42'       -> 42 (int)
    - '3,14'     -> 3.14 (float)
    - 'Bonjour'  -> 'Bonjour'
    - '3,14 cm'  -> '3,14 cm' (reste string)
    """
    if isinstance(valeur, str):
        val_strip = valeur.strip()

        # Cas entier pur
        if val_strip.isdigit():
            return int(val_strip)

        # Cas nombre flottant avec virgule ou point uniquement
        # On autorise une seule virgule ou un seul point, et rien d'autre
        val_clean = val_strip.replace(',', '.')
        try:
            # Vérifie que la chaîne représente uniquement un nombre (optionnellement avec signe)
            if all(c in "0123456789+-. " for c in val_strip) and val_strip.replace(',', '.').count('.') <= 1:
                return float(val_clean)
        except ValueError:
            pass

        # Sinon, renvoyer la chaîne d'origine (non modifiée)
        return valeur.strip()

    return valeur

def convertir_en_liste(val: Optional[Union[str, List[str], tuple, pd.Series]]) -> List[str]:
    """
    Convertit une valeur en liste Python de chaînes de caractères.

    Supporte :
        - str → [str]
        - list → list inchangée
        - tuple → list
        - pandas.Series → list
        - None ou NaN → []

    Args:
        val: valeur à convertir (str, list, tuple, Series, None ou NaN)

    Returns:
        List[str]: liste prête à être utilisée
    """
    # Cas None ou NaN (ex: valeur vide venant d'Excel)
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return []

    elif isinstance(val, str):
        return [val.strip()] if val.strip() else []

    elif isinstance(val, pd.Series):
        return val.dropna().astype(str).tolist()

    elif isinstance(val, (list, tuple)):
        # Convertit tuple en list et nettoie les éléments vides ou NaN
        return [str(v).strip() for v in val if pd.notna(v) and str(v).strip()]

    else:
        # Cas inattendu
        raise TypeError(f"Type non supporté pour conversion en liste : {type(val)}")

def tuple_vers_liste_de_listes(t):
    """
    Permet de transformer un tuple (2, 4) en liste de listes [[2], [4]]
    C'est la forme qu'il nous faut pour ajouter efficacement des lignes dans un tableau avec xlwings

    donnees = [
        ['Jean', 'Durand', 'jean@example.com', 42],
        ['Claire', 'Martin', 'claire@example.com', 35]
    ]
    donnees = [[2], [4]]


    :param t: le tuple
    :type donnees: Tuple
    :type nom_tableau: str
    :return: Liste de listes
    :rtype: List

    :Example:
    >>> tuple_vers_liste_de_listes((2, 4))


    .. seealso:: Rien du tout.
    .. warning:: Rien du tout.
    .. note:: Rien du tout.
    .. todo:: Rien du tout.
    """
    return [[elem] for elem in t]

def html_vers_texte(html: str) -> str:
    """
    Convertit du HTML en texte brut simple (compatible Outlook Body).

    - Supprime les balises et styles.
    - Conserve les paragraphes comme des sauts de ligne.
    - Ajoute des retours à la ligne pour les <br>, <p>, <li>.

    Exemple
    -------
    >>> html_vers_texte("<p>Bonjour <b>tout le monde</b></p><p>Deuxième paragraphe</p>")
    'Bonjour tout le monde\n\nDeuxième paragraphe'
    """
    soup = BeautifulSoup(html, "html.parser")

    # Gérer les retours à la ligne
    for br in soup.find_all("br"):
        br.replace_with("\n")
    for p in soup.find_all("p"):
        p.insert_before("\n")
        p.insert_after("\n")
    for li in soup.find_all("li"):
        li.insert_before("- ")

    # Texte brut
    texte = soup.get_text()
    
    # Nettoyage : strip + normalisation des sauts de ligne
    lignes = [l.strip() for l in texte.splitlines()]
    texte_final = "\n".join([l for l in lignes if l])  # supprime les lignes vides en trop

    return texte_final

def convertir_tuple_path(input:Path|Tuple[Path]) -> Tuple[Path]:
        if isinstance(input, Path):
            return (input,)
        else:
            return input

### --------------------------------------------------------------------
#  Dates
### --------------------------------------------------------------------
def mois_fr_depuis_date(date_val: Union[datetime, str, int, float, pd.Timestamp], 
                        use_locale: bool = False) -> str:
    """
    Renvoie le nom du mois en français à partir d'une valeur de date (Excel, datetime ou texte).

    Args:
        date_val (Union[datetime, str, int, float, pd.Timestamp]): 
            Valeur représentant une date (ex. : datetime, timestamp, texte ISO ou valeur Excel).
            Si Excel stocke une date comme nombre (ex. 45678), la fonction la convertira automatiquement.
        use_locale (bool, optional): 
            Si True, utilise la locale système ("fr_FR") pour récupérer le mois.
            Si False (par défaut), utilise une liste interne de mois en français.

    Returns:
        str: Le nom du mois en français, par exemple "avril".

    Raises:
        ValueError: Si la valeur ne peut pas être convertie en date valide.

    Exemple:
    --------
    >>> mois_fr_depuis_date("2024-04-15")
    'avril'

    >>> mois_fr_depuis_date(datetime(2025, 10, 27))
    'octobre'

    >>> mois_fr_depuis_date(45678)  # Valeur Excel correspondant à une date
    'janvier'

    >>> mois_fr_depuis_date("2025/12/01", use_locale=True)
    'décembre'
    """

    # Conversion en datetime si nécessaire
    if not isinstance(date_val, datetime):
        date_val = pd.to_datetime(date_val, errors="coerce")

    if pd.isna(date_val):
        raise ValueError(f"Impossible de convertir la valeur '{date_val}' en date valide.")

    # Option 1 : via locale système
    if use_locale:
        try:
            locale.setlocale(locale.LC_TIME, "fr_FR.UTF-8")
        except locale.Error:
            try:
                locale.setlocale(locale.LC_TIME, "fr_FR")
            except locale.Error:
                pass  # Fallback si non disponible
        return date_val.strftime("%B").capitalize()

    # Option 2 : liste interne (plus fiable sur Windows)
    noms_mois = [
        "janvier", "février", "mars", "avril", "mai", "juin",
        "juillet", "août", "septembre", "octobre", "novembre", "décembre"
    ]

    return noms_mois[date_val.month - 1]


### --------------------------------------------------------------------
#  Chaînes de caractères
### --------------------------------------------------------------------
def minuscule_premiere_lettre(s: str) -> str:
    """Met en minuscule uniquement la première lettre d'une chaîne."""
    return s[:1].lower() + s[1:] if s else s

### --------------------------------------------------------------------
#  Fenêtres
### --------------------------------------------------------------------
def choisirFichiers_filedialog_excel(initialdir: Optional[Path] = None) -> Path:
    """
    Lister/sélectionner les documents à concaténer.
    Le filedialog s'affiche au premier plan.
    """
    # Crée une fenêtre racine temporaire
    root = tk.Tk()
    root.withdraw()            # cache la fenêtre principale
    root.attributes('-topmost', True)  # met le filedialog au premier plan

    # Affiche le filedialog
    chemin_fichier = filedialog.askopenfilename(
        title="Sélectionner le fichier à charger",
        filetypes=[("Fichiers Excel", "*.xlsx")],
        initialdir=initialdir or os.getcwd()
    )

    # Détruire la fenêtre racine pour libérer les ressources
    root.destroy()

    chemin_fichier = Path(chemin_fichier) if chemin_fichier else None

    if not chemin_fichier:
        print("Click sur cancel du filedialog → Pas de chemins de fichier")

    return chemin_fichier


def choisir_fichier(
    titre: str = "Sélectionner un fichier",
    types_fichiers: List[Tuple[str, str]] = [("Tous les fichiers", "*.*")],
    dossier_initial: Optional[Path] = None,
    obligatoire: bool = True,
    multi_fichiers: bool = False,
    texte_bouton_choisir: str = "Choisir à nouveau",
    texte_bouton_aucun: str = "Pas de fichier à sélectionner",
    texte_bouton_quitter: str = "Quitter l'application",
) -> Optional[Union[Path, List[Path]]]:
    """
    Ouvre une boîte de dialogue pour sélectionner un ou plusieurs fichiers, avec gestion
    élégante du cas où aucun fichier n'est sélectionné.

    Parameters
    ----------
    titre : str, optional
        Titre de la fenêtre de sélection de fichier (par défaut "Sélectionner un fichier").
    types_fichiers : list[tuple[str, str]], optional
        Liste des types de fichiers acceptés (ex: [("Fichiers Excel", "*.xlsx")]).
    dossier_initial : Path, optional
        Dossier dans lequel ouvrir la boîte de dialogue (par défaut : répertoire courant).
    obligatoire : bool, optional
        Si True (défaut), force l'utilisateur à choisir un fichier ou quitter.
        Si False, l'utilisateur peut ignorer la sélection.
    multi_fichiers : bool, optional
        Si True, permet la sélection multiple (retourne une liste de chemins).
        Si False (défaut), ne permet qu'un seul fichier.
    texte_bouton_choisir : str, optional
        Libellé du bouton pour relancer la sélection.
    texte_bouton_aucun : str, optional
        Libellé du bouton "aucun fichier", affiché seulement si `obligatoire=False`.
    texte_bouton_quitter : str, optional
        Libellé du bouton pour quitter le programme.

    Returns
    -------
    Path | list[Path] | None
        - Chemin unique (Path) si `multi_fichiers=False`.
        - Liste de chemins (List[Path]) si `multi_fichiers=True`.
        - None si aucun fichier n'est sélectionné et `obligatoire=False`.

    Raises
    ------
    SystemExit
        Si l'utilisateur clique sur “Quitter l'application”.

    Examples
    --------
    >>> # Sélection d'un seul fichier Excel
    >>> fichier = choisir_fichier(
    ...     titre="Choisir un fichier Excel",
    ...     types_fichiers=[("Fichiers Excel", "*.xlsx")],
    ... )
    >>> print(fichier)
    Path('data/suivi_formations.xlsx')

    >>> # Sélection de plusieurs fichiers CSV (facultatif)
    >>> fichiers = choisir_fichier(
    ...     titre="Sélectionner des fichiers CSV",
    ...     types_fichiers=[("Fichiers CSV", "*.csv")],
    ...     multi_fichiers=True,
    ...     obligatoire=False
    ... )
    >>> print(fichiers)
    [Path('data/iris1.csv'), Path('data/iris2.csv')]
    """

    # === Fonctions internes ===
    def centrer_fenetre(fenetre: tk.Toplevel, largeur: int = 400, hauteur: int = 250):
        """Centre la fenêtre sur l'écran."""
        fenetre.update_idletasks()
        x = (fenetre.winfo_screenwidth() // 2) - (largeur // 2)
        y = (fenetre.winfo_screenheight() // 2) - (hauteur // 2)
        fenetre.geometry(f"{largeur}x{hauteur}+{x}+{y}")

    def popup_aucun_fichier() -> str:
        """Boîte de dialogue si aucun fichier n'a été sélectionné."""
        choix = {}

        def choisir_nouveau():
            choix["reponse"] = "choisir"
            fenetre.quit()
            fenetre.destroy()

        def aucun():
            choix["reponse"] = "aucun"
            fenetre.quit()
            fenetre.destroy()

        def quitter():
            choix["reponse"] = "quitter"
            fenetre.quit()
            fenetre.destroy()

        fenetre = tk.Toplevel()
        fenetre.title("Aucun fichier sélectionné")
        fenetre.resizable(False, False)
        fenetre.attributes('-topmost', True)
        fenetre.grab_set()
        fenetre.focus_force()

        # === Calcul dynamique de la largeur des boutons ===
        # On mesure la largeur réelle (en pixels) du texte le plus long
        police_bouton = tk.Font(family="Segoe UI", size=10, weight="bold")
        textes_boutons = [texte_bouton_choisir, texte_bouton_quitter]
        if not obligatoire:
            textes_boutons.append(texte_bouton_aucun)

        largeur_max_px = max(police_bouton.measure(t) for t in textes_boutons)
        marge_px = 40  # marges internes du bouton
        largeur_bouton_px = largeur_max_px + marge_px
        largeur_bouton_car = max(20, largeur_bouton_px // 8)  # conversion approximative pour paramètre "width" de tk.Button

        # === Texte explicatif ===
        label = tk.Label(
            fenetre,
            text="Aucun fichier n'a été sélectionné.\nQue souhaitez-vous faire ?",
            pady=20,
            font=("Segoe UI", 11),
        )
        label.pack()

        # === Cadre pour les boutons ===
        cadre_boutons = tk.Frame(fenetre)
        cadre_boutons.pack(pady=10)

        # Bouton "Choisir à nouveau" (vert)
        bouton_choisir = tk.Button(
            cadre_boutons,
            text=texte_bouton_choisir,
            width=largeur_bouton_car,
            bg="#4CAF50", fg="white",
            font=("Segoe UI", 10, "bold"),
            command=choisir_nouveau
        )
        bouton_choisir.grid(row=0, column=0, padx=5, pady=5)

        # Bouton "Aucun fichier" (optionnel)
        if not obligatoire:
            bouton_aucun = tk.Button(
                cadre_boutons,
                text=texte_bouton_aucun,
                width=largeur_bouton_car,
                bg="#DDDDDD",
                font=("Segoe UI", 10),
                command=aucun
            )
            bouton_aucun.grid(row=1, column=0, padx=5, pady=5)

        # Bouton "Quitter" (rouge)
        bouton_quitter = tk.Button(
            cadre_boutons,
            text=texte_bouton_quitter,
            width=largeur_bouton_car,
            bg="#E74C3C", fg="white",
            font=("Segoe UI", 10, "bold"),
            command=quitter
        )
        bouton_quitter.grid(row=2 if not obligatoire else 1, column=0, padx=5, pady=5)

        # === Adapter automatiquement la largeur de la fenêtre ===
        fenetre.update_idletasks()
        largeur_fenetre = max(420, largeur_bouton_px + 120)
        hauteur_fenetre = 240
        centrer_fenetre(fenetre, largeur_fenetre, hauteur_fenetre)

        fenetre.wait_window()
        return choix.get("reponse")

    # === Fenêtre racine invisible ===
    root = tk.Tk()
    root.withdraw()

    if dossier_initial is not None:
        dossier_initial = Path(dossier_initial)
    else:
        dossier_initial = Path.cwd()

    # === Boucle principale ===
    while True:
        root.attributes('-topmost', True)
        root.update()

        if multi_fichiers:
            chemins = filedialog.askopenfilenames(
                parent=root,
                title=titre,
                filetypes=types_fichiers,
                initialdir=dossier_initial
            )
        else:
            chemins = filedialog.askopenfilename(
                parent=root,
                title=titre,
                filetypes=types_fichiers,
                initialdir=dossier_initial
            )

        root.attributes('-topmost', False)

        # --- Si un ou plusieurs fichiers sont choisis
        if chemins:
            root.destroy()
            if multi_fichiers:
                return [Path(c) for c in chemins]
            else:
                return Path(chemins)

        # --- Aucun fichier sélectionné → boîte modale
        reponse = popup_aucun_fichier()

        if reponse == "choisir":
            continue
        elif reponse == "aucun":
            print(f"⚠️  Pas de fichier sélectionné.")
            root.destroy()
            return None
        elif reponse == "quitter":
            print("❌ Application quittée par l'utilisateur.")
            root.destroy()
            sys.exit()
        else:
            continue
            print("⚠️ Réponse inattendue. Fermeture.")
            root.destroy()
            sys.exit()



def arranger_fenetres(word_app, excel_app):
    """
    Place Word à gauche et Excel à droite sur l'écran principal,
    rend les fenêtres visibles et met Word au premier plan.
    """
    # Taille de l'écran principal
    screen = gw.getWindowsWithTitle("Program Manager")[0]
    screen_width, screen_height = screen.width, screen.height
    w_half = screen_width // 2

    # --- Word à gauche ---
    word_app.Visible = True
    word_window = word_app.ActiveWindow
    word_window.WindowState = 0        # wdWindowStateNormal
    word_window.Left = 0
    word_window.Top = 0
    word_window.Width = w_half
    word_window.Height = screen_height
    word_app.Activate()                 # Word au premier plan

    # --- Excel à droite ---
    excel_app.Visible = True
    excel_window = excel_app.ActiveWindow
    excel_window.WindowState = -4143   # xlNormal
    excel_window.Left = w_half
    excel_window.Top = 0
    excel_window.Width = w_half
    excel_window.Height = screen_height

    # Note : Excel est visible et à côté, Word reste au premier plan













### --------------------------------------------------------------------
#  TOUT CE QUI EST EN DESSOUS N'A PLUS ÉTÉ TESTÉ DEPUIS LONGTEMPS : CE SONT DES BASES DE REFLEXION
### --------------------------------------------------------------------

### --------------------------------------------------------------------
#  openpyxl
### --------------------------------------------------------------------

def writeDataFrameInStructuredRef_openpyxl(df, wb, nom_ws, nom_table = "", supprimeDonneesEtRemplace = False) -> None:
    """
    Ecrit un DataFrame dans un tableau structure d'une feuille de calcul  

    :param df: DataFrame à integrer dans le tableau structure
    :type df: DataFrame
    :param wb: classeur a lire 
    :type df: openpyxl.workbook
    :param nom_ws: nom de la feuille dans laquelle est le tableau structure
    :type nom_ws: string
    :param nom_table: nom du tableau structure. Si non renseigné, alors ce sera le même nom que l'onglet
    :type nom_table: string
    :param supprimeDonneesEtRemplace: Pour savoir si l'on ajoute les données du DataFrame à l'existant (False) ou si l'on supprime les données existantes et qu'on les remplace avec celles du DataFrame
    :type supprimeDonneesEtRemplace: Boolean
    :return: rien (on écrit/sauve un fichier excel)
    :rtype: None

    :Example:

    >>> writeDataFrameInStructuredRef(df_output, wb, nom_ws, nom_table = "Sessions", supprimeDonneesEtRemplace = False)


    .. seealso:: Rien du tout.
    .. warning:: Rien du tout.
    .. note:: Pour rajouter des lignes, on le fait à la suite de la feuille (worksheet) . Il en résulte qu'on gruge un peu : 1) on vire le tableau structuré, 2) on colle toutes les nouvelles lignes, 3) on supprime la ligne 1 du tableau, 4) on redéfinit les dimensions du tableau structuré (car l'ajout de nouvbelles lignes ne l'étend pas automatiquement)
    .. todo:: Rien du tout.
    """
    # Si nom_Table n'est pas défini en argument, c'est que par défaut c'est le même que nom_ws 
    if nom_table == "": nom_table = nom_ws

    # On ouvre la feuille et le tableau structuré
    ws = wb[nom_ws]
    #print(ws)
    table = ws._tables[nom_table] #Tableau structuré nommé

    # Infos de longueurs de mon tableau
    min_col, min_row, max_col, max_row = range_boundaries(table.ref)
    total_rows = max_row - min_row + 1
    data_rows = total_rows - table.headerRowCount
    #print(table.ref, total_rows, data_rows, min_col, min_row, max_col, max_row)
    
    # On écrit toutes les autres lignes une par une (on garde les lignes initiales pour garder le format qu'on copiera)
    # Méthode 1 qui marche
    #for il in tqdm(df.itertuples(), desc="Ecriture output Excel"):
    total_lignes = len(df)
    with tqdm(total=total_lignes, unit=' ligne', desc=Fore.CYAN + f"Écriture des lignes dans l'output {nom_table}" + Style.RESET_ALL) as pbar:
        for i, il in enumerate(df.itertuples(), 1):
            pbar.set_postfix(progress=f"{i}/{len(df)}")
            row = [val if pd.notna(val) else None for val in il[1:]] #Je dois rajouter cette ligne car il faut tester si je n'ai pas de valeurs <NA> qu'il faut retravailler sinon ça plante
            ws.append(row)
            pbar.update(1)
    
    # Méthode 2 - Bug
    #for r in dataframe_to_rows(df, index=False, header=False):
    #    ws.append(r)
    
    # On redéfinit le dimensionnement du tableau (/!\ +1 ligne pour récupérer le format de la dernière ligne)
    table.ref = f"{ws.cell(row=min_row, column=min_col).coordinate}:{ws.cell(row=max_row + len(df), column=max_col).coordinate}" #On saut le nb de ligens avant l'en-tête, puis l'en-tête, puis on va à la première ligne de données nbLignes_avantET1+1+1)
    #table.ref = "A1:{}{}".format(lettreFinTableau, len(df)+1) #+1 car on a la ligne d'en-tête #Méthodo initiale qui requiert de connaître la lettre de fin du tableau

    # On copie le format sur toutes les nouvelles lignes du tableau
    copieFormatTableauStructure_openpyxl(ws, table, indexLigneSourceFormat = min_row + 1, indexLigneDebutCopie = max_row + 1, indexLigneFinCopie = max_row + len(df) +1)

    # Si désiré par l'utilisateur, alors on supprime les anciennes lignes de la feuille Excel (ça garde la dimension initiale du tableau structuré)
    if supprimeDonneesEtRemplace:
        # Suppression des anciennes lignes
        ws.delete_rows(idx=min_row + 1, amount=data_rows)

        # On redéfinit les dimensions du tableau structuré
        table.ref = f"{ws.cell(row=min_row, column=min_col).coordinate}:{ws.cell(row=min_row + len(df), column=max_col).coordinate}" #On saut le nb de ligens avant l'en-tête, puis l'en-tête, puis on va à la première ligne de données nbLignes_avantET1+1+1)

def copieFormatTableauStructure_openpyxl(ws, table, indexLigneSourceFormat = 1, indexLigneDebutCopie = -1, indexLigneFinCopie = -1) -> None:
    """
    Recopie le format d'une ligne d'un tableau structure a une plage du tableau structure

    :param ws: feuille dans laquelle est le tableau structure
    :type ws: openpyxl.worksheet
    :param table: tableau structure
    :type table: openpyxl.Table
    :param indexLigneSourceFormat: indice de la ligne qui est a recopier (indice absolu dans la feuille excel, i.e. pas #ligne dans le tableau structure). Defaut = 1.
    :type indexLigneSourceFormat: int
    :param indexLigneDebutCopie: indice de la 1ere ligne ou il faut copier le format (indice absolu dans la feuille excel, i.e. pas #ligne dans le tableau structure). Defaut = indexLigneSourceFormat + 1
    :type indexLigneDebutCopie: int
    :param indexLigneFintCopie: indice de la derniere ligne ou il faut copier le format (indice absolu dans la feuille excel, i.e. pas #ligne dans le tableau structure). Defaut = derniere ligne tableau structure
    :type indexLigneFinCopie: int
    :return: rien (on a copie les format dans le worksheet)
    :rtype: None

    :Example:

    >>> copieFormatTableauStructure(ws, table, indexLigneSourceFormat = min_row + 1, indexLigneDebutCopie = max_row + 1, indexLigneFinCopie = max_row + len(df) +1)


    .. seealso:: Rien du tout.
    .. warning:: Rien du tout.
    .. note:: Pour rajouter des lignes, on le fait à la suite de la feuille (worksheet) . Il en résulte qu'on gruge un peu : 1) on vire le tableau structuré, 2) on colle toutes les nouvelles lignes, 3) on supprime la ligne 1 du tableau, 4) on redéfinit les dimensions du tableau structuré (car l'ajout de nouvbelles lignes ne l'étend pas automatiquement)
    .. todo:: Rien du tout.
    """

    # On met à jour les styles des cellules (openpywl ne sait pas insérer de lignes en conservant les formats ; par ailleurs on ne sait pas appliquer ça ligne par ligne ou colonne par colonne : on va donc le faire cellule par cellule)
    # On n'agrandit pas automatiquement le tableau structuré si l'on rajoute une cellule
    # Quand on agrandit un tableau en redéfinissant le ref, on ne colle pas le format



    # Infos de longueurs de mon tableau
    min_col, min_row, max_col, max_row = range_boundaries(table.ref)
    
    # Initialisattion paramètres non renseignés
    if indexLigneDebutCopie == -1: indexLigneDebutCopie = indexLigneSourceFormat + 1
    if indexLigneFinCopie == -1: indexLigneFinCopie = max_row


    # On récupère les formats de chaque cellule de la première ligne du tableau structuré 
    formats=[]
    for icol in range(min_col, max_col+1) :
        formats.append(ws.cell(indexLigneSourceFormat, icol))
        #print(formats[icol-1].number_format)
    #print(formats)

    # On copie colle les formats avec ces cellules
    with tqdm(total=max_col, unit=' colonnes', desc=Fore.CYAN + "Copie des formats" + Style.RESET_ALL, ncols=150) as pbar:
        for icol in range(min_col, max_col+1) :
            pbar.set_postfix(progress=f"{icol}/{max_col}")
            source = formats[icol-1] #Je prends un index de liste et pas un numéro de colonne, donc -1
            #print(source.number_format, source.number_format == "General")

            # Optimisation : on ne fait les copies que si le format est différent de General
            #if source.number_format != "General":
            for il in range(indexLigneDebutCopie, indexLigneFinCopie) : #+1 pour le row car en-tête
                #print(il, icol, ws.cell(row=il, column=icol).value, source.number_format)
                target = ws.cell(row=il, column=icol)
                #target.font = copy(source.font)
                #target.border = copy(source.border)
                target.fill = copy(source.fill)
                #target.alignment = copy(source.alignment)
                #target.protection = copy(source.protection)
                target.number_format = copy(source.number_format)

            pbar.update(1)

def concatene_ongletsExcels_openpyxl(nom_ws, nom_ws_imports = "", rep_defaut = ".", rep_output = "", nomBaseFichier = "Excel-output", nbLignes_avantET = 0) -> None:
    """
    Concatene des fichiers Excel avec une même stucture 

    :param path_in: Chemin du repertoire a tester
    :type path_in: string
    :return: un chemin optimal (i.e. avec la plus longue arborescence) qui est fonctionnel
    :rtype: string

    :Example:

    >>> string chemin = optimiseCheminRepertoire("C:\\Users\\fichier.xlsx")


    .. seealso:: Rien du tout.
    .. warning:: Rien du tout.
    .. note:: Rien du tout.
    .. todo:: Rien du tout.
    """

    if rep_output == "": rep_output=rep_defaut


    # Lister/sélectionner les documents à concaténer
    listeCheminsExcel = filedialog.askopenfilename(title="Sélectionner les fichiers Excel à concaténer", filetype=[("fichiers excel","*.xlsx")], initialdir=rep_defaut, multiple=True)

    # On met les fichiers input dans un DataFrame
    df_input = pd.concat((pd.read_excel(iFichier, skiprows=nbLignes_avantET) for iFichier in listeCheminsExcel), ignore_index=True)

    # Définition du nom du classeur de base qui va servir à la sortie par la suite
    s_output = rep_output + "\\" + nomBaseFichier + ".xlsx" #ou f"{datetime.now():%Y.%m.%d}"
    #print(s_output)

    # On copie le classeur
    #shutil.copy(s_classeurIni, s_classeurDestination)

    # On ouvre le classeur qui va recevoir la concaténation
    wb_output = load_workbook(filename = s_output)
    
    # On met les fichiers input dans un DataFrame
    df_input = pd.concat((pd.read_excel(rep_defaut + "\\" + iFichier, skiprows=nbLignes_avantET) for iFichier in listeCheminsExcel), ignore_index=True)

    # On écrit dans le tableau structuré 
    writeDataFrameInStructuredRef_openpyxl(df_input, wb_output, nom_ws)

    # On écrit les références des fichiers copiés dans le tableau structuré "Imports"
    if nom_ws_imports != "":
        ws = wb_output.sheets[nom_ws_imports]
        table = ws.tables[nom_ws_imports]
        table.range.end('down').offset(1, 0).value = ['val1']

    #On enregistre et on ferme
    wb_output.save(filename=rep_output + "\\" + nomBaseFichier + "-" + date.today().strftime("%Y.%m.%d") + ".xlsx") #ou f"{datetime.now():%Y.%m.%d}")
    wb_output.close()


### --------------------------------------------------------------------
#  xlwings
### --------------------------------------------------------------------

def writeDataFrameInStructuredRef_openpyxl_xlwings(df, wb, chemin_wb, nom_ws, nom_table = "", supprimeDonneesEtRemplace = False) -> None:
    """
    Ecrit un DataFrame dans un tableau structuré existant d'une feuille de calcul  
    Utilise openpyxl pour l'écriture des données, puis xlwings pour copier rapidement le format.

    :param df: DataFrame à integrer dans le tableau structure
    :type df: DataFrame
    :param wb: classeur a lire 
    :type df: openpyxl.workbook
    :param nom_ws: nom de la feuille dans laquelle est le tableau structure
    :type nom_ws: string
    :param nom_table: nom du tableau structure. Si non renseigné, alors ce sera le même nom que l'onglet
    :type nom_table: string
    :param supprimeDonneesEtRemplace: Pour savoir si l'on ajoute les données du DataFrame à l'existant (False) ou si l'on supprime les données existantes et qu'on les remplace avec celles du DataFrame
    :type supprimeDonneesEtRemplace: Boolean
    :return: rien (on écrit/sauve un fichier excel)
    :rtype: None

    :Example:

    >>> writeDataFrameInStructuredRef(df_output, wb, nom_ws, nom_table = "Sessions", supprimeDonneesEtRemplace = False)


    .. seealso:: Rien du tout.
    .. warning:: Rien du tout.
    .. note:: Pour rajouter des lignes, on le fait à la suite de la feuille (worksheet) . Il en résulte qu'on gruge un peu : 1) on vire le tableau structuré, 2) on colle toutes les nouvelles lignes, 3) on supprime la ligne 1 du tableau, 4) on redéfinit les dimensions du tableau structuré (car l'ajout de nouvbelles lignes ne l'étend pas automatiquement)
    .. todo:: Rien du tout.
    """
    # Si nom_Table n'est pas défini en argument, c'est que par défaut c'est le même que nom_ws 
    if nom_table == "": nom_table = nom_ws

    # On ouvre la feuille et le tableau structuré
    ws = wb[nom_ws]
    table = ws._tables[nom_table] #Tableau structuré nommé

    # Infos de dimension du tableau
    min_col, min_row, max_col, max_row = range_boundaries(table.ref)
    total_rows = max_row - min_row + 1
    data_rows = total_rows - table.headerRowCount
    #print(table.ref, total_rows, data_rows, min_col, min_row, max_col, max_row)
    
    # On écrit toutes les autres lignes une par une (on garde les lignes initiales pour garder le format qu'on copiera)
    # Méthode 1 qui marche
    #for il in tqdm(df.itertuples(), desc="Ecriture output Excel"):
    total_lignes = len(df)
    with tqdm(total=total_lignes, unit=' ligne', desc=Fore.CYAN + f"Écriture des lignes dans l'output {nom_table}" + Style.RESET_ALL) as pbar:
        for i, il in enumerate(df.itertuples(), 1):
            pbar.set_postfix(progress=f"{i}/{len(df)}")
            row = [val if pd.notna(val) else None for val in il[1:]] #Je dois rajouter cette ligne car il faut tester si je n'ai pas de valeurs <NA> qu'il faut retravailler sinon ça plante
            ws.append(row)
            pbar.update(1)
    
    # Méthode 2 - Bug
    #for r in dataframe_to_rows(df, index=False, header=False):
    #    ws.append(r)
    
    # Redimensionnement du tableau 
    nouvelle_max_row = max_row + len(df)
    table.ref = f"{ws.cell(row=min_row, column=min_col).coordinate}:{ws.cell(row=nouvelle_max_row, column=max_col).coordinate}"

    # Nécessaire avant d'utiliser xlwings
    wb.save(chemin_wb)  
    wb.close()

    # ==== Copie du format avec xlwings ====
    # Définir la plage source (= ligne de format) et plage cible (= nouvelles lignes)
    col_lettre_debut = get_column_letter(min_col)
    col_lettre_fin = get_column_letter(max_col)

    range_modele = f"{col_lettre_debut}{min_row+1}:{col_lettre_fin}{min_row+1}"  # première ligne de données
    range_cible = f"{col_lettre_debut}{max_row+1}:{col_lettre_fin}{nouvelle_max_row}"  # nouvelles lignes

    # On copie le format sur toutes les nouvelles lignes du tableau
    copierFormat_xlwings(chemin_wb, nom_ws, range_modele, range_cible)
    
    # Réouvrir pour finaliser les suppressions éventuelles
    wb = load_workbook(filename=chemin_wb)
    ws = wb[nom_ws]
    table = ws._tables[nom_table]

    # Si désiré par l'utilisateur, alors on supprime les anciennes lignes de la feuille Excel (ça garde la dimension initiale du tableau structuré)
    if supprimeDonneesEtRemplace:
        # Suppression des anciennes lignes
        ws.delete_rows(idx=min_row + 1, amount=data_rows)

        # On redéfinit les dimensions du tableau structuré
        table.ref = f"{ws.cell(row=min_row, column=min_col).coordinate}:{ws.cell(row=min_row + len(df), column=max_col).coordinate}" #On saut le nb de ligens avant l'en-tête, puis l'en-tête, puis on va à la première ligne de données nbLignes_avantET1+1+1)

def concatene_ongletsExcels_xlwings(nom_ws, nom_ws_imports = "", rep_defaut = ".", rep_output = "", nomBaseFichier = "Excel-output", nbLignes_avantET = 0) -> None:
    """
    Concatene des fichiers Excel avec une même stucture.
    Si rep_output == "", alors rep_output=rep_defaut
    nom_ws_imports est facultatif, c'est juste si l'on souhaite sauvegarder la liste des fichiers concatenes

    :param nom_ws: nom du worksheet
    :type nom_ws: str
    :param nom_ws_imports: nom du worksheet ou on enregistre les fichiers excel qui vont etre concatene (defaut = "")
    :type nom_ws_imports: str
    :param rep_defaut: repertoire par defaut ou on va aller chercher les fichiers excel a concatener (defaut = ".")
    :type rep_defaut: str
    :param rep_output: repertoire ou on va enregistrer le fichier excel contenant la concatenation des fichiers (defaut = "", si rep_output == "", alors rep_output=rep_defaut)
    :type rep_output: str
    :param nomBaseFichier: nom du fichier output (defaut = "Excel-output")
    :type nomBaseFichier: str
    :param nbLignes_avantET: nombre de ligne avant l'en-tete des tableaux que l'on va concatener
    :type nbLignes_avantET: int




    :Example:
    >>> concatene_ongletsExcels("Inscriptions", nom_ws_imports = "Imports", rep_defaut = rep_extractIRIS_INSTNT, rep_output = rep_extractIRIS_VTE, nomBaseFichier = "R04500_Sessions-Inscriptions-COMPLET", nbLignes_avantET = 1)


    .. seealso:: Rien du tout.
    .. warning:: Rien du tout.
    .. note:: si rep_output == "", alors rep_output=rep_defaut
    .. todo:: C'est tres long, il faudra que je teste les vitesses entre xlwings et openpyxl.
    """

    if rep_output == "": rep_output=rep_defaut

    #Lister/sélectionner les documents à concaténer
    listeCheminsExcel = filedialog.askopenfilename(title="Sélectionner les fichiers Excel à concaténer", filetype=[("fichiers excel","*.xlsx")], initialdir=rep_defaut, multiple=True)

    # On met les fichiers input dans un DataFrame
    #for iFichier in listeCheminsExcel :
    #    df_input = pd.read_excel(rep_defaut + "\\" + iFichier, skiprows=nbLignes_avantET)
    df_input = pd.concat((pd.read_excel(iFichier, skiprows=nbLignes_avantET) for iFichier in listeCheminsExcel), ignore_index=True)
    
    # On met le dataframe en liste pour l'envoyer à ajouter_lignes_tableau_xlwings
    donnees = df_input.values.tolist()
    
    # Forme à avoir
    #donnees = [
    #    ['Jean', 'Durand', 'jean@example.com', 42],
    #    ['Claire', 'Martin', 'claire@example.com', 35]
    #]
    #donnees = [[2], [4]]
    #donnees = df.values.tolist()


    # Définition du nom du classeur de base qui va servir à la sortie par la suite
    s_output = rep_output + "\\" + nomBaseFichier + ".xlsx" #ou f"{datetime.now():%Y.%m.%d}"
    #print(s_output)

    # On ouvre l'app pour xlwings
    app = xw.App(visible=False)  # Excel s'ouvre en arrière-plan
        
    # On copie le classeur
    #shutil.copy(s_classeurIni, s_classeurDestination)

    # On ouvre le classeur qui va recevoir la concaténation
    wb_output = app.books.open(s_output)

    # On écrit dans le tableau structuré 
    #writeDataFrameInStructuredRef_openpyxl(df_input, wb_output, nom_ws) #Marche
    ajouter_lignes_tableau_xlwings(donnees, wb_output, nom_ws) #Tres long (~20 minutes pour mon test avec toutes les sessions)


    # On écrit les références des fichiers copiés dans le tableau structuré "Imports" ssi il y a un nom dans nom_ws_imports
    if nom_ws_imports != "":
        ajouter_lignes_tableau_xlwings(tuple_vers_liste_de_listes(listeCheminsExcel), wb_output, nom_ws_imports)

    #On enregistre et on ferme
    #wb_output.save(filename=rep_defaut + "\\" + nomBaseFichier + "-" + date.today().strftime("%Y.%m.%d") + ".xlsx") #ou f"{datetime.now():%Y.%m.%d}")
    wb_output.save(rep_output + "\\" + nomBaseFichier + "-" + date.today().strftime("%Y.%m.%d") + ".xlsx") #ou f"{datetime.now():%Y.%m.%d}")
    wb_output.close()
    app.quit()

def ajouter_lignes_tableau_xlwings(donnees, wb, nom_feuille, nom_tableau = "") -> None:
    """
    Insere une ou plusieurs lignes dans un tableau structure Excel en minimisant les appels COM.
    Méthode longue → Préférer openpyxl si possible

    :param donnees: Liste contenant les donnees a inserer dans le tableau Excel
    :type donnees: List
    :param wb: Classeur dans lequel on souhaite inscrire nos donnees
    :type wb: xlwing book
    :param nom_feuille: nom de la feuille contenant le tableau
    :type nom_feuille: str
    :param nom_tableau: nom du tableau structuré (ListObject) (defaut = "")
    :type nom_tableau: str

    :Example:
    >>> ajouter_lignes_tableau_xlwings(donnees, wb_output, nom_ws)


    .. seealso:: Rien du tout.
    .. warning:: Rien du tout.
    .. note:: Rien du tout.
    .. todo:: C'est tres long (~20 minutes pour mon test avec toutes les sessions), il faudra que je teste les vitesses entre xlwings et openpyxl.
    """




    #On gère le cas par défaut où on ne donne pas de nom_tableau car c'est le même que le nom de la feuille
    if nom_tableau == "": nom_tableau = nom_feuille


    ws = wb.sheets[nom_feuille]
    table = ws.api.ListObjects(nom_tableau)
    
    data_body_range = table.DataBodyRange

    nb_lignes_nouvelles = len(donnees)
    nb_colonnes = len(donnees[0])

    # Détermine l'endroit où écrire : soit première ligne du tableau, soit après la dernière ligne existante
    if data_body_range is None or data_body_range.Value is None:
        start_cell = ws.range((table.HeaderRowRange.Row + 1, table.HeaderRowRange.Column))
    else :
        nb_lignes_existantes = data_body_range.Rows.Count
        next_row = data_body_range.Row + nb_lignes_existantes
        start_cell = ws.range((next_row, data_body_range.Column))

    # Écriture en bloc pour performance    
    ws.range(start_cell.address).resize(nb_lignes_nouvelles, nb_colonnes).value = donnees

def copierFormat_xlwings(chemin_fichier, nom_ws, range_modele, range_cible) -> None:
    """
    Copie rapidement le format d'une plage source vers une plage cible dans un fichier Excel 
    en utilisant `xlwings` et l'API COM (équivalent au collage spécial > formats dans Excel).

    :param chemin_fichier: Chemin complet du fichier Excel à modifier.
    :type chemin_fichier: str
    :param nom_ws: Nom de la feuille contenant les plages.
    :type nom_ws: str
    :param range_modele: Adresse de la plage source contenant les formats à copier (ex: "A2:G2").
    :type range_modele: str
    :param range_cible: Adresse de la plage cible à laquelle appliquer les formats (ex: "A3:G100").
    :type range_cible: str

    :return: Aucun. Le fichier Excel est modifié et enregistré.
    :rtype: None

    :example:
    >>> copier_format_rapide(
            nom_fichier="mon_fichier.xlsx",
            nom_feuille="Données",
            range_modele="A2:G2",
            range_cible="A3:G100"
        )

    .. note::
        Cette fonction nécessite Microsoft Excel installé sur votre machine (Windows uniquement).

    .. warning::
        Seuls les formats (style, bordures, police, etc.) sont copiés. Les valeurs ne sont pas modifiées.
    """
    
    app = xw.App(visible=False)
    wb = app.books.open(chemin_fichier)
    ws = wb.sheets[nom_ws]

    # Copie de la première ligne (format uniquement)
    ws.range(range_modele).copy()

    # Collage spécial des formats uniquement
    #tqdm.write("⏳ Application des formats avec Excel (xlwings)...")
    ws.range(range_cible).api.PasteSpecial(Paste=-4122)  # -4122 = xlPasteFormats
    #tqdm.write("✅ Formats collés avec succès.")

    wb.save()
    wb.close()
    app.quit()

