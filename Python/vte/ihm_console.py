import inspect
from typing import Any, Callable, List, Optional, Union
import re

class IHM_console:
    """
    Interface Homme-Machine en mode console, générique.

    Cette classe permet :
    - d'afficher des menus hiérarchiques (statique ou dynamique),
    - d'exécuter des actions associées,
    - de demander à l'utilisateur les arguments manquants automatiquement,
    - d'afficher des indications facultatives pour aider l'utilisateur.

    Exemple de menu :

    menus = {
        "Mails": {
            "Mail de premier contact": {
                "action": ca.creer_mails_contactInitial,
                "kwargs": {},
                "demander": []
            },
        },
        "RDV Outlook": {
            "Créer un RDV pour un entretien": {
                "sous-menu": [ca._entretiens, "sujet"],
                "action": ca.creer_rdv,
                "kwargs": lambda e: {"entretien": e},
                "demander": []
            }
        },
        "Relances": {
            "Choisir un type de relance": {
                "sous-menu": [ca._relances, None],
                "action": ca.creer_mails_relances,
                "kwargs": lambda r: {"relance": r},
                "demander": []
            }
        }
    }

    Exemple d'utilisation :

    >>> ihm = IHM_console(menus)
    >>> ihm.afficher_menu()
    """

    def __init__(self, menus: dict, contexte: Any = None):
        """
        Étape 1 : Initialisation

        Paramètres
        ----------
        menus : dict
            Arborescence des menus avec actions et sous-menus
        contexte : Any, optionnel
            Objet partagé passé aux callbacks si nécessaire (ex. instance de gestion)
        """
        self.menus = menus
        self.contexte = contexte

    # -------------------------------------------------------------------------
    # MÉTHODE : demander_saisie
    # -------------------------------------------------------------------------

    def demander_saisie(self, texte: str, type_attendu: type = str) -> Any:
        """
        Demande une saisie utilisateur en console, conversion robuste.

        - Gère types simples (int, float, str, bool).
        - Gère listes typées (list[int], list[str], List[int], etc.).
        - Séparateurs acceptés pour listes : virgules et/ou espaces.
        """

        # Normalisation si annotation donnée comme chaîne (ex: "list[int]" ou "int")
        if isinstance(type_attendu, str):
            s = type_attendu.lower()
            if "list" in s and "int" in s:
                type_attendu = list[int]
            elif "list" in s and "str" in s:
                type_attendu = list[str]
            elif s in ("int", "integer"):
                type_attendu = int
            elif s == "float":
                type_attendu = float
            elif s in ("bool", "boolean"):
                type_attendu = bool
            else:
                type_attendu = str

        # helper: obtenir callable pour un sous-type (int/str/float)
        def _get_callable(t):
            # si typing alias (ex list[int]) on prend le premier arg
            if hasattr(t, "__origin__") and getattr(t, "__args__", None):
                inner = t.__args__[0]
                # si inner est typing alias like 'int' as str, try to map
                if isinstance(inner, str):
                    m = {"int": int, "str": str, "float": float, "bool": bool}
                    return m.get(inner.lower(), str)
                return inner if callable(inner) else str
            # si t est type direct
            if isinstance(t, type) and callable(t):
                return t
            return str

        while True:
            try:
                # Affichage du libellé du type attendu (lisible)
                type_label = getattr(type_attendu, "__name__", str(type_attendu))
                saisie = input(f"{texte} ({type_label}) : ").strip()

                # Booléen
                if type_attendu == bool:
                    return saisie.lower() in ("o", "oui", "y", "yes", "true", "1")

                # Cas liste typée : detecte list[...] via __origin__
                origin = getattr(type_attendu, "__origin__", None)
                args = getattr(type_attendu, "__args__", None) or ()

                if origin is list or origin is list or (isinstance(type_attendu, type) and issubclass(type_attendu, list) if isinstance(type_attendu, type) else False):
                    # récupère type interne si possible, sinon str
                    inner_callable = _get_callable(type_attendu)
                    # split flexible : accepte virgules et/ou espaces
                    parts = [p for p in re.split(r"[,\s]+", saisie) if p != ""]
                    parsed = []
                    for p in parts:
                        try:
                            parsed.append(inner_callable(p))
                        except Exception:
                            # si conversion échoue, garde la string brute
                            parsed.append(p)
                    # DEBUG — tu peux commenter la ligne suivante si tu veux
                    # print(f"[DEBUG] parsed list for '{texte}': {parsed}")
                    return parsed

                # Cas liste non-typée (annotation exactly 'list')
                if type_attendu is list:
                    parts = [p for p in re.split(r"[,\s]+", saisie) if p != ""]
                    return parts

                # Types simples (int/float/str)
                if callable(type_attendu):
                    try:
                        return type_attendu(saisie)
                    except Exception:
                        # si conversion échoue, on alerte et reprompt
                        print(f"⚠️ Impossible de convertir '{saisie}' en {type_label}. Réessayez.")
                        continue

                # Par défaut : retourne chaîne
                return saisie

            except Exception as e:
                print(f"⚠️ Erreur de saisie ({e}). Réessayez.")

# -------------------------------------------------------------------------
    # MÉTHODE : executer_action
    # -------------------------------------------------------------------------
    def executer_action(self, action_def: dict, choix_sousmenu: Any = None) -> Any:
        """
        Étape 3 : Exécuter l'action d'un menu

        Paramètres
        ----------
        action_def : dict
            Définition d'une action, pouvant contenir :
            - "action" : callable
            - "kwargs" : dict ou callable pour générer kwargs dynamiquement
            - "demander" : liste des arguments à demander explicitement
            - "indications" : texte facultatif pour guider l'utilisateur
        choix_sousmenu : Any
            Objet choisi par l'utilisateur dans un sous-menu dynamique

        Retourne
        --------
        result : Any
            Résultat de la fonction exécutée
        """
        action = action_def.get("action")
        if not callable(action):
            raise ValueError("Action non exécutable.")

        # Étape 3a : afficher indications facultatives
        if "indications" in action_def and action_def["indications"]:
            print("\n💡 Indications :")
            print(action_def["indications"].strip())
            print()

        # Étape 3b : préparer les kwargs
        kwargs_def = action_def.get("kwargs", {})
        if callable(kwargs_def):
            if choix_sousmenu is None:
                raise ValueError("Un choix de sous-menu est requis pour générer les kwargs.")
            kwargs = kwargs_def(choix_sousmenu)
        else:
            kwargs = dict(kwargs_def)

        # Étape 3c : introspection pour récupérer les paramètres manquants
        sig = inspect.signature(action)
        demander = list(action_def.get("demander", []))

        for nom, param in sig.parameters.items():
            if nom == "cls":  # ignorer cls dans les classmethods
                continue
            if nom not in kwargs and param.default == inspect.Parameter.empty:
                # Type attendu (par annotation si dispo)
                type_attendu = (
                    param.annotation
                    if (hasattr(param, "annotation") and param.annotation != inspect._empty)
                    else str
                )
                # 🔥 Correction : on garde le type brut, pas son nom en str
                kwargs[nom] = self.demander_saisie(f"Entrez {nom}", type_attendu)

        # Étape 3d : demander explicitement les arguments dans "demander"
        for nom in demander:
            if nom not in kwargs:
                param = sig.parameters.get(nom)
                type_attendu = (
                    param.annotation
                    if (param and param.annotation != inspect._empty)
                    else str
                )
                kwargs[nom] = self.demander_saisie(f"Entrez {nom}", type_attendu)

        # Étape 3e : message avant exécution (placé APRÈS les saisies)
        print("\n\t→ Exécute l'action\n")

        # Étape 3f : exécution finale
        return action(**kwargs)

    # -------------------------------------------------------------------------
    # MÉTHODE : afficher_menu
    # -------------------------------------------------------------------------
    def afficher_menu(self, menu: Optional[dict] = None, titre: str = "Menu principal"):
        """
        Étape 4 : Afficher le menu et gérer la navigation

        Paramètres
        ----------
        menu : dict, optionnel
            Menu à afficher (par défaut menu racine)
        titre : str
            Titre du menu

        Notes
        -----
        - "0" permet de revenir ou quitter
        - Les sous-menus dynamiques sont de la forme :
            "sous-menu": [liste_objets, attribut_affichage]
            - attribut_affichage=None → str(obj)
            - attribut_affichage='sujet' → getattr(obj, 'sujet')
        """
        if menu is None:
            menu = self.menus

        while True:
            print(f"\n=== {titre} ===")
            options = list(menu.keys())
            for i, opt in enumerate(options, 1):
                print(f"{i}. {opt}")
            print("0. Quitter / Retour")

            try:
                choix = int(input("Votre choix : "))
            except ValueError:
                print("⚠️ Entrée invalide, merci de saisir un nombre.")
                continue

            if choix == 0:
                return

            if not (1 <= choix <= len(options)):
                print("⚠️ Choix invalide, réessayez.")
                continue

            cle = options[choix - 1]
            valeur = menu[cle]

            # Étape 4a : sous-menu dynamique
            if isinstance(valeur, dict) and "sous-menu" in valeur:
                liste_objets, attr = valeur["sous-menu"]
                if attr:
                    sous_menu_temp = {getattr(obj, attr): obj for obj in liste_objets}
                else:
                    sous_menu_temp = {str(obj): obj for obj in liste_objets}

                print(f"\n--- {cle} ---")
                sous_options = list(sous_menu_temp.keys())
                for i, opt in enumerate(sous_options, 1):
                    print(f"{i}. {opt}")
                print("0. Retour")

                try:
                    sous_choix = int(input("Votre choix : "))
                except ValueError:
                    print("⚠️ Entrée invalide, merci de saisir un nombre.")
                    continue

                if sous_choix == 0:
                    continue

                label = sous_options[sous_choix - 1]
                element_choisi = sous_menu_temp[label]

                # Étape 4b : exécuter action sur l'élément choisi
                self.executer_action(valeur, choix_sousmenu=element_choisi)

            # Étape 4c : action directe
            elif isinstance(valeur, dict) and "action" in valeur:
                self.executer_action(valeur)

            # Étape 4d : sous-menu statique
            elif isinstance(valeur, dict):
                self.afficher_menu(valeur, titre=cle)

            else:
                raise ValueError(f"⚠️ Entrée de menu non valide : {valeur}")
