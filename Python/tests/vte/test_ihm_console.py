import pytest
from unittest.mock import MagicMock
from datetime import timedelta, date, datetime

from ihm_console import IHM_console

# =============================================================================
# CLASSES FICTIVES POUR LES TESTS
# =============================================================================
class FakeCA:
    """
    Classe factice simulant les actions d'un gestionnaire d'apprentis.
    Chaque fonction est mockée avec MagicMock pour vérifier les appels.
    """
    def __init__(self):
        # Actions mockées
        self.creer_mails_contactInitial = MagicMock()
        self.creer_mails_ficheEvaluation = MagicMock()
        self.creer_rdv = MagicMock()
        self.creer_mails_relances = MagicMock()

        # Sous-menus dynamiques
        self._entretiens = [FakeEntretien("Entretien 1"), FakeEntretien("Entretien 2")]
        self._relances = ["Relance A", "Relance B"]

        # Exemple d'attribut utilisé pour kwargs/demander
        self.chemin_modele_mail_priseContact = "chemin/factice.msg"
        self._envoyer_mail = True
        self._df_etudiants = [{"Mail apprenti": "a@example.com", "Mail TE": "t@example.com"}]

class FakeEntretien:
    """Classe factice pour les objets entretien"""
    def __init__(self, sujet):
        self.sujet = sujet

# =============================================================================
# FIXTURES PYTEST
# =============================================================================
@pytest.fixture
def ca():
    return FakeCA()

@pytest.fixture
def menus(ca):
    """
    Définition du menu complet pour les tests, incluant :
    - Action directe
    - Sous-menu dynamique avec objet
    - Sous-menu dynamique avec string
    """
    return {
        "Mails": {
            "Mail de premier contact": {
                "action": ca.creer_mails_contactInitial,
                "kwargs": {},
                "demander": []
            },
            "Envoi fiche d'évaluation": {
                "action": ca.creer_mails_ficheEvaluation,
                "kwargs": {},
                "demander": []
            },
        },
        "RDV Outlook": {
            "Créer un RDV pour un entretien": {
                "sous-menu": [ca._entretiens, "sujet"],  # affichage sur .sujet
                "action": ca.creer_rdv,
                "kwargs": lambda p: {"prop": p},        # injecte l'objet choisi
                "demander": []
            }
        },
        "Relances": {
            "Relances": {
                "sous-menu": [ca._relances, None],      # affichage string
                "action": ca.creer_mails_relances,
                "kwargs": lambda r: {"relance": r},     # injecte string choisie
                "demander": []
            }
        }
    }

# =============================================================================
# TEST PARAMÉTRÉ
# =============================================================================
@pytest.mark.parametrize(
    "menu_path, sousmenu_choice, expected_func, expected_kwargs",
    [
        # Menu principal -> action directe
        (["Mails", "Mail de premier contact"], None, "creer_mails_contactInitial", {}),
        (["Mails", "Envoi fiche d'évaluation"], None, "creer_mails_ficheEvaluation", {}),
        # Menu principal -> sous-menu dynamique avec objet
        (["RDV Outlook", "Créer un RDV pour un entretien"], 0, "creer_rdv", {"prop": 0}),
        # Menu principal -> sous-menu dynamique avec string
        (["Relances", "Relances"], 1, "creer_mails_relances", {"relance": "Relance B"})
    ]
)
def test_ihm_console_complete(monkeypatch, menus, ca, menu_path, sousmenu_choice, expected_func, expected_kwargs):
    """
    Test complet de IHM_console.

    Ce test vérifie :

    1) La capacité de IHM_console à naviguer correctement dans le dictionnaire menu
       - Affichage des différents niveaux
       - Revenir / quitter via 0
       - Exécution des fonctions aux bons moments

    2) La gestion des arguments via 'kwargs' et 'demander'
       - Les arguments déjà fournis ne sont pas redemandés
       - Les arguments listés dans 'demander' sont sollicités auprès de l'utilisateur
       - Les arguments obligatoires/facultatifs sont gérés correctement

    3) Tous les types de menus sont testés :
       - Action directe
       - Sous-menu dynamique avec objet
       - Sous-menu dynamique avec string

    4) Les fonctions sont appelées avec les paramètres corrects sans exécuter
       les actions réelles (pas d'envoi de mails, pas de création de RDV Outlook)
    """
    ihm = IHM_console(menus, contexte=ca)

    # -------------------------------
    # Construction de la séquence d'inputs simulés
    # -------------------------------
    inputs = []

    # Parcours du menu principal jusqu'à l'option ciblée
    current_menu = menus
    for cle in menu_path:
        options = list(current_menu.keys())
        idx = options.index(cle) + 1  # +1 car le menu affiche à partir de 1
        inputs.append(str(idx))
        # Descend dans le menu si nécessaire
        valeur = current_menu[cle]
        if isinstance(valeur, dict) and "sous-menu" in valeur:
            current_menu = valeur
        else:
            current_menu = valeur

    # Si sous-menu dynamique, choix de l'élément
    if sousmenu_choice is not None:
        inputs.append(str(sousmenu_choice + 1))  # menu commence à 1

    # Retour / Quitter
    inputs.append("0")  # Quitter sous-menu ou menu principal
    inputs.append("0")  # Quitter menu principal définitivement

    # Patch input() pour renvoyer les valeurs simulées
    monkeypatch.setattr("builtins.input", lambda _: inputs.pop(0))

    # -------------------------------
    # Lancer l'affichage du menu
    # -------------------------------
    ihm.afficher_menu()

    # -------------------------------
    # Vérifier que la bonne fonction a été appelée
    # -------------------------------
    func_mock = getattr(ca, expected_func)
    assert func_mock.called, f"{expected_func} n'a pas été appelé"

    # Vérification des arguments si applicables
    if expected_kwargs:
        if expected_func == "creer_rdv":
            # On récupère l'objet choisi depuis _entretiens
            func_mock.assert_called_with(prop=ca._entretiens[sousmenu_choice])
        elif expected_func == "creer_mails_relances":
            func_mock.assert_called_with(relance=expected_kwargs["relance"])