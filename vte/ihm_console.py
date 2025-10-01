import inspect

class IHM_console:
    """
    Interface console générique pour naviguer dans des menus et exécuter des actions.
    Supporte :
        - Sous-menus statiques (dict imbriqués)
        - Sous-menus dynamiques à partir de listes d’objets
        - Arguments obligatoires et facultatifs à demander à l'utilisateur
        - Types d'arguments (str, int, bool, float, date, timedelta, etc.)

    Déclaration d'un menu dynamique exemple :

    menus = {
        "RDV Outlook": {
            "Créer un RDV pour un entretien": {
                "sous-menu": [ca._entretiens, "sujet"],  # [liste d'objets, attribut pour l'affichage]
                "action": ca.creer_rdv,
                "kwargs": lambda e: {"entretien": e},  # kwargs dynamiques basés sur l'objet
                "demander": []  # arguments supplémentaires à demander
            }
        },
        "Envoi fichiers": {
            "Envoyer les fichiers à renvoyer": {
                "sous-menu": [ca._fichiersARenvoyer, "sujet"],
                "action": ca.envoyer_fichier,
                "kwargs": lambda f: {"fichier": f},
                "demander": []
            }
        }
    }
    """

    def __init__(self, menus, contexte=None):
        """
        Initialise l'IHM console.

        :param menus: dict représentant la structure des menus
        :param contexte: objet contexte (ex: 'ca') à passer aux sous-menus dynamiques si besoin
        """
        self.menus = menus
        self.contexte = contexte

    def demander_saisie(self, message, type_attendu=str):
        """
        Demande une saisie à l'utilisateur et convertit au type attendu.
        Gère les booléens avec quelques valeurs acceptées.
        Si type_attendu est déjà une str (ex: annotation textuelle), on l'interprète comme 'str'.

        :param message: texte de la question
        :param type_attendu: type attendu de la réponse (str par défaut)
        :return: valeur convertie au type attendu
        """

        if isinstance(type_attendu, str):
            # parfois l’annotation arrive sous forme de string -> fallback sur str
            type_attendu = str

        while True:
            try:
                valeur = input(f"{message} ({type_attendu.__name__}) : ")
                if type_attendu == bool:
                    return valeur.lower() in ["true", "1", "oui", "o", "y", "yes"]
                return type_attendu(valeur)
            except ValueError:
                print(f"Erreur : veuillez entrer une valeur de type {type_attendu.__name__}.")

    def executer_action(self, action_def, choix_sousmenu=None):
        action = action_def["action"]

        kwargs_def = action_def.get("kwargs", {})
        if callable(kwargs_def):
            # Si kwargs est une fonction → on la résout avec le choix du sous-menu
            if choix_sousmenu is None:
                raise ValueError("Un choix de sous-menu est requis pour générer les kwargs.")
            kwargs = kwargs_def(choix_sousmenu)
        else:
            # Sinon → c’est déjà un dictionnaire fixe
            kwargs = dict(kwargs_def)

        # Gestion des paramètres à demander (idem ton code actuel)
        for nom in action_def.get("demander", []):
            sig = inspect.signature(action)
            param = sig.parameters.get(nom)
            type_attendu = param.annotation if param and param.annotation != inspect._empty else str
            kwargs[nom] = self.demander_saisie(nom, type_attendu)

        # Enfin on exécute
        return action(**kwargs)

    def afficher_menu(self, menu=None, titre="Menu principal"):
        """
        Affiche le menu et gère la navigation.

        :param menu: sous-menu à afficher (None = menu principal)
        :param titre: titre affiché pour ce menu
        """
        if menu is None:
            menu = self.menus

        while True:
            # Gestion des sous-menus dynamiques si menu callable
            if callable(menu):
                menu = menu(self.contexte)

            print(f"\n{titre}")
            for i, cle in enumerate(menu.keys(), 1):
                print(f"{i}. {cle}")
            print("0. Quitter / Retour")

            try:
                choix = int(input("Votre choix : "))
            except ValueError:
                print("Veuillez entrer un nombre valide.")
                continue

            if choix == 0:
                return

            if 1 <= choix <= len(menu):
                cle = list(menu.keys())[choix - 1]
                valeur = menu[cle]

                # ---- Gestion des sous-menus dynamiques ----
                if isinstance(valeur, dict) and "sous-menu" in valeur:
                    liste_objets = valeur["sous-menu"]
                    action = valeur.get("action")
                    kwargs_func = valeur.get("kwargs", {})
                    demander = valeur.get("demander", [])

                    # Détection de la syntaxe [liste_objets, label_attr]
                    label_attr = None
                    if isinstance(liste_objets, (list, tuple)) and len(liste_objets) == 2:
                        liste_objets, label_attr = liste_objets

                    # Création du sous-menu temporaire
                    sous_menu_temp = {}
                    for obj in liste_objets:
                        if label_attr:
                            label = getattr(obj, label_attr)
                        else:
                            label = str(obj)
                        sous_menu_temp[label] = {
                            "action": action,
                            "kwargs": kwargs_func,
                            "demander": demander
                        }

                    # Appel récursif sur le sous-menu
                    self.afficher_menu(sous_menu_temp, titre=cle)

                # ---- Action finale ----
                elif isinstance(valeur, dict) and "action" in valeur:
                    self.executer_action(valeur)

                # ---- Sous-menu statique classique ----
                elif isinstance(valeur, dict):
                    self.afficher_menu(valeur, titre=cle)

                else:
                    print(f"⚠️ Valeur non reconnue dans le menu pour {cle} : {valeur}")