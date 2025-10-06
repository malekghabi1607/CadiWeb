from vte.instn import *
from vte.ihm_console import IHM_console
from vte.ihm_tkinter import IHMTkinter

#import Traiter_contactsApprentis
#import Mail


def main():
    """
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
    contexte = Traiter_contactsApprentis.L3D()
    #gt = GestionTutoratUGA(ca, Mail)

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
    
    # mode = input("Choisir mode (console/tkinter) : ").strip().lower()
    mode = "console"
    #mode = "tkinter"
    

    if mode == "console":
        ihm = IHM_console(menus, contexte=contexte)
        ihm.afficher_menu()
    else:
        ihm = IHMTkinter(menus)
        ihm.lancer()



if __name__ == "__main__":
    main()