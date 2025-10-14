from vte.instn import *
from vte.ihm_console import IHM_console
from vte.ihm_tkinter import IHMTkinter


def main():
    print("")

def test():
    ree = REE()
    ree.envoyerMail_REE(statut = "auto-entrepreneur")
    print("")


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




