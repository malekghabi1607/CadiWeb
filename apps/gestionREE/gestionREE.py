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


    # === Ouvrir un ficheier word et récupérer les content control
    ree._word_ficheAdministrative = FichierWord.depuisFichier()
    print(ree._word_ficheAdministrative)

    # === Créer le répertoire du REE sur le réseau
    ree._creer_repertoire_REE(test=True)

    # === Déplacer les fichiers du REE dans le répertoire
    #ree._deplacer_fichiers(ree._repertoire_sauvegarde_fichiersREE)


    # === Remplit le fichier Excel à transférer à Laetitia Da Mota à partir d'un modèle
    ree._remplit_excel_avecInfos_word()

    # === On ouvre l'Excel et le Word pour comparaison et adaptations manuelles
    chemin_word = os.path.join(ree._repertoire_sauvegarde_fichiersREE, os.path.basename(ree._word_ficheAdministrative._chemin_fichier))
    chemin_excel = os.path.join(ree._repertoire_sauvegarde_fichiersREE, os.path.basename(ree._chemin_modele_excel_ficheIntervenant))

    ouvrir_word_excel_cote_a_cote(chemin_word, chemin_excel, split_ecranPrincipal=True)
    
    #arranger_fenetres(word_app, excel_app)

    # On attend pour avancer
    input("🕒 Attente pour adaptations de l'Excel. Appuyez sur une touche pour continuer")

    # Dès que l'Excel est fermé, on prépare le mail pour Laetitia
    print("Mail Laetitia")

if __name__ == "__main__":
    test()
    #main()



### --------------------------------------------------------------------
#   Code pour Fichier EE
### --------------------------------------------------------------------
#fw = FichierWord.depuisFichier(r"C:\Users\vt238770\Documents\_CEA\EXCEL - Tests ou Backup\Lien formulaire Word vers Excel\Fiche admin - ContentControl.docx")
#print(fw)
#mail= Mail()

#mail._creer_mail(
#    destinataires="vacataires.instn@cea.fr",
#    sujet="Documents pour mise à jour IRIS", # Pimper avec le nom de l'intervenant
#    corps_html="<p>Bonjour Laëtitia,</p><p>Je t’ai mis en PJ les documents pour intégrer/mettre à jour la fiche IRIS de ###.</p><p>Je te remercie, passe une excellente journée,</p>", # Pimper avec le nom de l'intervenant
#    pieces_jointes=r"C:\Users\vt238770\Documents\_CEA\EXCEL - Tests ou Backup\Lien formulaire Word vers Excel\Fiche admin - ContentControl.docx", # if None, sélectionner avec fileDialog
#    envoyer_mail=False  # envoie directement sans afficher
#)

#t_ree = Traiter_REE()
#Il me faudrait une classe Traiter_REE :
#   - variables de la classe : _word_ficheAdministrative:FichierWord= None, _excel_ficheIntervenant:FichierExcel = None, _mail_traitement _repertoire_sauvegarde + _mail_gestionnaire_REE +




