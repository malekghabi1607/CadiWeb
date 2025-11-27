from vte.instn import *
from vte.ihm_console import IHM_console

from colorama import init
init(autoreset=True)


MODE = "console"

MENUS = {
    "Préparer un mail pour un type d'intervenant": {
        "CEA": {
            "action": REE.envoyerMail_REE,
            "kwargs": {"statut": "CEA"},
            "demander": []
        },
        "Vacataire": {
            "action": REE.envoyerMail_REE,
            "kwargs": {"statut": "vacataire"},
            "demander": []
        },"Auto-entrepreneur": {
            "action": REE.envoyerMail_REE,
            "kwargs": {"statut": "auto-entrepreneur"},
            "demander": []
        },"contrat spécifique de collaboration": {
            "action": REE.envoyerMail_REE,
            "kwargs": {"statut": "contrat spécifique de collaboration"},
            "demander": []
        },
    },

    "Traiter un retour de document d'un REE": {
            "action": REE.traiter_docs_REE,
            "kwargs": {},
            "demander": [],
        },
}

def main():
    
    vlog.print("Info", "*************\nBienvenue dans le script pour générer les REE\n*************", style=["vert clair"])

    ihm = IHM_console(MENUS)
    ihm.afficher_menu()


def test():
    #ree = REE()

    # === Envoyer mail : ok avec bon texte et PJ ===
    # TODO : comment récupérer le statut depuis Excel ?
    # TODO : comment faire une boucle auto sur les personnes à qui envoyer ?
    REE.envoyerMail_REE(statut = "auto-entrepreneur")

    
    # Réception / traitement REE
    REE.traiter_docs_REE() # OK
    # TODO : mettre à jour le fichier Excel des coordonnées des intervenants
    # TODO : mettre à jour le fichier Excel des AI


    

if __name__ == "__main__":
    #test()
    main()




