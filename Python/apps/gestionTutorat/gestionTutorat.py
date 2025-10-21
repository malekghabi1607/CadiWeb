from vte.instn import *
from vte.ihm_console import IHM_console
from vte.ihm_tkinter import IHMTkinter

#import Traiter_contactsApprentis
#import Mail

# =====================================================
# MAPPING CENTRALISÉ DES CONTEXTES DISPONIBLES
# =====================================================
CONTEXTES_MAPPING = {
    "UGA": Traiter_contactsApprentis.UGA,
    "L3D": Traiter_contactsApprentis.L3D
}
    
# mode = input("Choisir mode (console/tkinter) : ").strip().lower()
mode = "console"
#mode = "tkinter"


def main(retour_contexte_menus=False):
    """
    Fonction principale :
    1) Affiche un mini-menu IHM_console pour choisir le contexte.
    2) Lance le menu principal correspondant.
    
    Ex déclaration de menu:
    menus = {
        "Mails": {
            # Ici on force un chemin par défaut
            "Mail de premier contact (chemin forcé)": {
                "action": ca.creer_mails_contactInitial,
                "kwargs": {"chemin_modele_mail_priseContact": "C:/modele_contact.msg"},
                "demander": []  # rien à demander à l'utilisateur
            },

            # Ici on laisse le chemin par défaut de la classe,
            # mais on demande explicitement si on veut envoyer les mails
            "Mail de premier contact (avec confirmation)": {
                "action": ca.creer_mails_contactInitial,
                "kwargs": {},  # pas d'argument forcé
                "demander": ["envoyer_mail"]  # cet argument sera toujours demandé
            },
        }
    }
    """
    #contexte = Traiter_contactsApprentis.UGA()
    #contexte = Traiter_contactsApprentis.L3D()
    #gt = GestionTutoratUGA(ca, Mail)


    # --- Étape 1 : choix du contexte ---
    #menu_contexte = {
    #    "sous-menu": [list(CONTEXTES_MAPPING.keys()), None],
    #    "action": lambda choix: CONTEXTES_MAPPING[choix]()  # crée et renvoie le bon contexte
    #}

    #ihm_context = IHM_console(menu_contexte)
    #contexte = ihm_context.afficher_menu(arret_apres_action=True)
    #if contexte is None:
    #    print("⚠️ Aucun contexte sélectionné, arrêt du programme.")
    #    return  # ou sys.exit(), ou un comportement par défaut

    # Appel direct, menu interactif pour choisir le contexte
    contexte = IHM_console.depuis_sous_menu(
        liste=list(CONTEXTES_MAPPING.keys()),
        action=lambda choix: CONTEXTES_MAPPING[choix](),  # renvoie l’instance du contexte choisi
        titre="Choisir le contexte"
    )

# contexte contient maintenant Traiter_contactsApprentis.UGA() ou L3D()

    # --- Étape 2 : définition du menu principal ---

    # Définition des menus
    menus = {
        "Mails": {
            "Mail de premier contact": {
                "action": contexte.creer_mails_contactInitial,
                "kwargs": {},
                "demander": []
            },
            "Envoi fiche d'évaluation": {
                "action": contexte.creer_mails_ficheEvaluation,
                "kwargs": {},
                "demander": []
            },
        },

        "RDV Outlook": {
            "Créer un RDV pour un entretien": {
                "sous-menu": [contexte._entretiens, "sujet"],  # affiche l'attribut .sujet
                "action": contexte.creer_rdv,
                "kwargs": lambda p: {"prop": p},
                "demander": []
            }
        },

        "Relances": {
            "sous-menu": [contexte._relances, None],  # liste de str, donc on affiche directement str(obj)
            "action": contexte.creer_mails_relances,
            "kwargs": lambda r: {"relance": r},  # injecte la string choisie dans l’appel (i.e. le nom de la relance)
            "demander": []
        }
    }

    

    if retour_contexte_menus :
        return contexte, menus


    if mode == "console":
        ihm = IHM_console(menus, contexte=contexte)
        ihm.afficher_menu()
    else:
        ihm = IHMTkinter(menus)
        ihm.lancer()


if __name__ == "__main__":
    main()