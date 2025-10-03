import inspect

class IHM_console:
    """
    Interface Homme-Machine en mode console.

    Cette classe permet d’afficher des menus hiérarchiques en console,
    et d’exécuter des actions associées. Elle a été pensée pour rester générique :
    - L’affichage est ici textuel, mais la même logique peut être utilisée
      dans une interface graphique (chaque entrée devient un bouton).
    - Les menus peuvent être statiques (actions fixes) ou dynamiques
      (choix dans une liste d’objets, comme des entretiens ou des fichiers à envoyer).

    Exemple d’utilisation
    ---------------------
    >>> menus = {
    ...     "Mails": {
    ...         "Mail de premier contact": {
    ...             "action": ca.creer_mails_contactInitial,
    ...             "kwargs": {},
    ...             "demander": []
    ...         }
    ...     },
    ...     "RDV Outlook": {
    ...         "Créer un RDV pour un entretien": {
    ...             "sous-menu": [ca._entretiens, "sujet"],  # affiche l’attribut .sujet
    ...             "action": ca.creer_rdv,
    ...             "kwargs": lambda e: {"entretien": e},
    ...             "demander": []
    ...         }
    ...     },
    ...     "Relances": {
    ...         "Choisir un type de relance": {
    ...             "sous-menu": [ca._relances, None],  # liste de str
    ...             "action": ca.creer_mails_relances,
    ...             "kwargs": lambda r: {"relance": r},
    ...             "demander": []
    ...         }
    ...     }
    ... }
    >>> ihm = IHM_console(menus)
    >>> ihm.afficher_menu()
    """

    def __init__(self, menus, contexte=None):
        """
        Initialise l’IHM avec un dictionnaire de menus.

        Paramètres
        ----------
        menus : dict
            Dictionnaire décrivant l’arborescence des menus.
            Chaque entrée peut être :
              - un sous-menu (dict imbriqué)
              - une action directe (dict avec "action", "kwargs", "demander")
              - un menu dynamique avec "sous-menu"

        contexte : object, optionnel
            Un objet partagé qui peut être transmis aux callbacks
            si besoin (par ex. l’instance de gestion des apprentis).
        """
        self.menus = menus
        self.contexte = contexte

    # -------------------------------------------------------------------------
    # MÉTHODE : demander_saisie
    # -------------------------------------------------------------------------
    def demander_saisie(self, message, type_attendu=str):
        """
        Demande une saisie utilisateur en console, avec validation de type.

        Paramètres
        ----------
        message : str
            Texte affiché à l’utilisateur pour la saisie.
        type_attendu : type, optionnel
            Type attendu pour la conversion (par défaut : str).
            Peut être `int`, `float`, `bool`, etc.

        Retourne
        --------
        valeur : object
            La valeur saisie, convertie au bon type.

        Notes
        -----
        - Pour les booléens, les valeurs suivantes sont acceptées comme True :
          "true", "1", "oui", "o", "y", "yes".
        """
        while True:
            try:
                valeur = input(f"{message} ({getattr(type_attendu, '__name__', str(type_attendu))}) : ")
                if type_attendu == bool:
                    return valeur.lower() in ["true", "1", "oui", "o", "y", "yes"]
                return type_attendu(valeur)
            except ValueError:
                print(f"Erreur : veuillez entrer une valeur de type {type_attendu}.")

    # -------------------------------------------------------------------------
    # MÉTHODE : executer_action
    # -------------------------------------------------------------------------
    def executer_action(self, action_def, choix_sousmenu=None):
        """
        Exécute une action définie dans un menu.

        Paramètres
        ----------
        action_def : dict
            Dictionnaire décrivant l’action :
              - "action"   : fonction à appeler
              - "kwargs"   : dict d’arguments fixes ou fonction prenant le choix utilisateur
              - "demander" : liste de noms d’arguments à demander explicitement
        choix_sousmenu : object, optionnel
            Élément choisi par l’utilisateur dans un sous-menu (si applicable).
            Peut être un objet métier (ex. PropEntretien) ou une simple string.

        Retourne
        --------
        result : object
            Résultat de l’appel à la fonction associée.

        Notes
        -----
        - Si "kwargs" est un callable, on l’évalue avec choix_sousmenu.
        - Si des paramètres obligatoires ne sont pas fournis, ils sont demandés
          en console à l’utilisateur.
        """
        action = action_def["action"]
        kwargs_def = action_def.get("kwargs", {})
        demander = action_def.get("demander", [])

        # Cas 1 : kwargs est une fonction → on la calcule avec l’élément choisi
        if callable(kwargs_def):
            kwargs = kwargs_def(choix_sousmenu)
        else:
            kwargs = dict(kwargs_def)

        # Vérification de la signature de la fonction
        sig = inspect.signature(action)

        # Demande automatique des arguments obligatoires manquants
        for nom, param in sig.parameters.items():
            if nom in kwargs:
                continue
            if param.default == inspect.Parameter.empty:
                # Type attendu si indiqué dans l’annotation
                type_attendu = param.annotation if param.annotation != inspect._empty else str
                kwargs[nom] = self.demander_saisie(f"Entrez {nom}", type_attendu)

        # Demande explicite des arguments listés dans "demander"
        for nom in demander:
            if nom not in kwargs:
                param = sig.parameters.get(nom)
                type_attendu = str
                if param and param.annotation != inspect._empty:
                    type_attendu = param.annotation
                kwargs[nom] = self.demander_saisie(f"Entrez {nom}", type_attendu)

        return action(**kwargs)

    # -------------------------------------------------------------------------
    # MÉTHODE : afficher_menu
    # -------------------------------------------------------------------------
    def afficher_menu(self, menu=None, titre="Menu principal"):
        """
        Affiche un menu interactif en console et permet à l’utilisateur
        de naviguer dans les options.

        Paramètres
        ----------
        menu : dict, optionnel
            Le menu à afficher. Si None, affiche le menu racine.
        titre : str
            Titre affiché en en-tête du menu.

        Notes
        -----
        - L’utilisateur saisit un nombre correspondant à une entrée.
        - L’entrée "0" permet de revenir en arrière ou quitter.
        - Les sous-menus dynamiques peuvent être définis ainsi :
            "sous-menu": [liste, "attribut"]   → affichage basé sur getattr(obj, "attribut")
            "sous-menu": [liste, None]         → affichage basé sur str(obj)
        """
        if menu is None:
            menu = self.menus  # Menu racine par défaut

        while True:
            print(f"\n=== {titre} ===")

            # Affiche toutes les options disponibles
            options = list(menu.keys())
            for i, opt in enumerate(options, start=1):
                print(f"{i}. {opt}")
            print("0. Quitter / Retour")

            # Lecture du choix utilisateur
            try:
                choix = int(input("Votre choix : "))
            except ValueError:
                print("⚠️ Entrée invalide, merci de saisir un nombre.")
                continue

            if choix == 0:
                return  # Retour au menu précédent

            if not (1 <= choix <= len(options)):
                print("⚠️ Choix invalide, réessayez.")
                continue

            cle = options[choix - 1]
            valeur = menu[cle]

            # Cas 1 : sous-menu dynamique
            if isinstance(valeur, dict) and "sous-menu" in valeur:
                liste, attr = valeur["sous-menu"]

                # Préparation des options utilisateur
                if attr:  # on affiche un attribut spécifique
                    sous_menu_temp = {getattr(obj, attr): obj for obj in liste}
                else:  # on affiche directement la valeur
                    sous_menu_temp = {str(obj): obj for obj in liste}

                # Affichage du sous-menu
                print(f"\n--- {cle} ---")
                sous_options = list(sous_menu_temp.keys())
                for i, opt in enumerate(sous_options, start=1):
                    print(f"{i}. {opt}")
                print("0. Retour")

                try:
                    sous_choix = int(input("Votre choix : "))
                except ValueError:
                    print("⚠️ Entrée invalide, merci de saisir un nombre.")
                    continue

                if sous_choix == 0:
                    continue  # Retour au menu parent

                if not (1 <= sous_choix <= len(sous_options)):
                    print("⚠️ Choix invalide, réessayez.")
                    continue

                label = sous_options[sous_choix - 1]
                element_choisi = sous_menu_temp[label]

                # Exécute l’action avec l’élément choisi
                self.executer_action(valeur, choix_sousmenu=element_choisi)

            # Cas 2 : action directe
            elif isinstance(valeur, dict) and "action" in valeur:
                self.executer_action(valeur)

            # Cas 3 : sous-menu statique
            elif isinstance(valeur, dict):
                self.afficher_menu(valeur, titre=cle)

            else:
                raise ValueError(f"⚠️ Entrée de menu non valide : {valeur}")
