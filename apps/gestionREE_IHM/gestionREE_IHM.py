from vte.instn import *
from vte.ihm_console import IHM_console
from vte.ihm_tkinter import IHMTkinter

def main():
    print("")

def test():
    ree = REE()

    # === Envoyer mail : ok avec bon texte et PJ ===
    # TODO : comment récupérer le statu depuis Excel ?
    # TODO : comment faire une boucle auto sur les personnes à qui envoyer ?
    #ree.envoyerMail_REE(statut = "auto-entrepreneur")

    
    # Réception / traitement REE
    ree.traiter_docs_REE()
    # TODO : mettre à jour le fichier Excel des coordonnées des intervenants
    # TODO : mettre à jour le fichier Excel des AI


    

if __name__ == "__main__":
    test()
    #main()
