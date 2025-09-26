#from __future__ import annotations

from vte.instn import *


def main():
    """
    Appli console permettant de contacter les apprentis et tuteurs lors d'un suivi d'apprentissage :
       - mail de premier contact (à partir d'un modèle .msg) ;
       - préparer les RDV Outlook pour les entretiens (à partir de modèles .oft) ;
       - faire les mails de relance pour des documents ou Studea.
    
    Pour avoir accès aux mails et aux noms, j'ouvre le fichier Excel Etudiants (ADIN ou LP3D)
    Dans ce fichier, on note aussi la réception des docs ou des signatures Studea pour filtrer les mails à envoyer
    
    TODO :
       - UGA : rajouter mail pour envoyer fiche évaluation + RDV Outlook (1 et 2)
       - LP3D : tout

    OK FONCTIONNEL VÉRIFIÉ 25/09/2025
    """
    ca = Traiter_contactsApprentis.UGA()

    # On demande quel traitement on veut faire :
    print(
        f"Quel type de traitement voulez-vous faire ?\n"
        "\t• 1 : Mails\n"
        "\t• 2 : RDV Outlook pour les entretiens\n"
        "\t• 3 : Relance pour des documents ou Studea\n"
        "\n0 : Sortir de l'application\n"
    )
    while True:
        try:
            choix = int(input("Entrez un nombre entier : "))
            if (choix in range(0, 4)) :
                break  # Sort de la boucle si conversion réussie et dans le bon intervalle
        except ValueError:
            print("Ce n'est pas un entier valide. Essayez encore.")


    # On quitte l'appli
    if choix == 0 :
        exit()

    ### --------------------------------------------------------------------
    #  1 : Mail de premier contact (mi-septembre)
    ### --------------------------------------------------------------------
    if choix == 1 :

        # On demande quel traitement on veut faire :
        print(
            f"Quel type de traitement voulez-vous faire ?\n"
            "\t• 1 : Mail de premier contact\n"
            "\t• 2 : Envoi fiche d'évaluation\n"
            "\t• 3 : Mail à tous les apprentis et tuteurs suivis par tutorat (en CC) [TODO]\n"
            "\t• 4 : Mail à tous les étudiants de la promo [TODO]\n"
            "\n0 : Sortir de l'application\n"
        )
        while True:
            try:
                choix2 = int(input("Entrez un nombre entier : "))
                if (choix2 in range(0, 5)) :
                    break  # Sort de la boucle si conversion réussie et dans le bon intervalle
            except ValueError:
                print("Ce n'est pas un entier valide. Essayez encore.")

        
        if (choix2 == 0):
            exit()
        
        elif (choix2 == 1):
            ca.creer_mails_contactInitial()

        elif (choix2 == 2):
            ca.creer_mails_ficheEvaluation()

        elif (choix2 == 3):
            print("Méthode non encore faite")
            
        elif (choix2 == 4):
            print("Méthode non encore faite")

    ### --------------------------------------------------------------------
    #  2 : Préparer les RDV Outlook pour les entretiens
    ### --------------------------------------------------------------------
    if choix == 2 :
        q2 = f"Pour quel entretien voulez-faire un RDV Outlook ?\n"
        for i, entretien in enumerate(ca._entretiens) :
            q2 += f"\t• {i+1} : {entretien.sujet}\n"
        q2 += "\n0 : Sortir de l'application\n"

        print(q2)

        while True:
            try:
                choix2 = int(input("Entrez un nombre entier : "))
                if (choix in range(0, i+1)) :
                    break  # Sort de la boucle si conversion réussie et dans le bon intervalle
            except ValueError:
                print("Ce n'est pas un entier valide. Essayez encore.")

        if choix2 == 0 :  # 0 : On quitte l'appli
            exit()
        else :  # On créée les RDV Outlook associés aux entretiens
            ca.creer_rdv(ca._entretiens[choix2-1])

    

    ### --------------------------------------------------------------------
    #  3 : Faire les mails de relance pour des documents ou Studea
    ### --------------------------------------------------------------------
    if choix == 3 :
        print(
            f"Jusqu'à quel niveau voulez-vous faire les relances ?\n"
            "\t• 1 : Engagement des parties (Studea)\n"
            "\t• 2 : Entretien prise de fonction (Studea)\n"
            "\t• 3 : 1ère visite (Studea)\n"
            "\t• 4 : 2ème visite (Studea)\n"
            "\t• 5 : Fiche d'évaluation 1\n"
            "\t• 6 : Fiche d'évaluation 2\n"
            "\n0 : Sortir de l'application\n"
        )
        while True:
            try:
                choix2 = int(input("Entrez un nombre entier : "))
                if (choix2 in range(0, 7)) :
                    break  # Sort de la boucle si conversion réussie et dans le bon intervalle
            except ValueError:
                print("Ce n'est pas un entier valide. Essayez encore.")

        if choix2 == 0 :  # 0 : On quitte l'appli
            exit()
        
        else :  # On créée les mails de rappel
            ca.creer_mails_relances(choix2)

if __name__ == "__main__":
    main()

    # 
    #ca = Traiter_contactsApprentis.UGA()
    #ligne = ca._df_etudiants.iloc[0]
    #ca.creer_mail_ficheEvaluation(ligne, date(2025, 12, 5))







# Prise de contact - (mi-septembre)
#instance.contactInitial(chemin_modele_mail_priseContact)

# RDV Outlook prise de fonction - (Autour du 6 octobre : dernière semaine de la première période en entreprise → A faire avant mi-novembre)
#instance.creer_rdv(prise_de_fonction)

# RDV Outlook prise de fonction - (Autour du 15 décembre : dernière semaine avant vacances Noël et reprise école → A faire avant mi-janvier)
#instance.creer_rdv(premiere_visite)

# RDV Outlook prise de fonction - (Autour du 20 avril : dernière semaine reprise école → A faire avant fin mai)
#instance.creer_rdv(deuxieme_visite)

# Relances
# Demander jusqu'à quel niveau il faut faire les relances
"""
print(
    f"Jusqu'à quel niveau voulez-vous faire les relances ?\n"
    "\t• Engagement des parties : 0\n"
    "\t• Entretien prise de fonction : 1\n"
    "\t• 1ère visite : 2\n"
    "\t• 2ème visite : 3\n"
    "\t• Fiche d'évaluation 1 : 4\n"
    "\t• Fiche d'évaluation 2 : 5"
)
while True:
    try:
        choix = int(input("Entrez un nombre entier : "))
        break  # Sort de la boucle si conversion réussie
    except ValueError:
        print("Ce n'est pas un entier valide. Essayez encore.")

#TODO : rajouter PJ pour les fiches d'évaluation
# pieces_jointes = r"\\instnt\partage\FORMATIONS_I\GDRA - Master IN (parcours ADIN-GDRA-SN)\2025-2026\9-stages\consignes_missions_entreprises_2526.pdf"
pieces_jointes = None

for _, ligne in instance._df_etudiants.iterrows():
    # Apprenti
    corps_html = cls.creer_html_mail_relance(choix, ligne, "A")
    if corps_html:
        Mail.creer_mail(
            destinataires = ligne["Mail apprenti"],
            sujet = "Master IN - Relance actions suivis de l'alternance",
            corps_html = corps_html,
            pieces_jointes = pieces_jointes,
            envoyer_mail = envoyer_mail
        )
    
    # Tuteur entreprise
    corps_html = cls.creer_html_mail_relance(choix, ligne, "T")
    if corps_html:
        Mail.creer_mail(
            destinataires = ligne["Mail apprenti"],
            sujet = "Master IN - Relance actions suivis de l'alternance",
            corps_html = corps_html,
            pieces_jointes = pieces_jointes,
            envoyer_mail = envoyer_mail
        )
"""











