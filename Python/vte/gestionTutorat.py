
# ======================================================================================
# CLASSES GestionTutorat
# ======================================================================================
class Traiter_contactsApprentis:
    """
    Classe permettant de contacter les apprentis et tuteurs lors d'un suivi d'apprentissage :
       - mail de premier contact (à partir d'un modèle .msg) ;
       - préparer les RDV Outlook pour les entretiens (à partir de modèles .oft) ;
       - faire les mails de relance pour des documents ou Studea.
    
    Pour avoir accès aux mails et aux noms, j'ouvre le fichier Excel Etudiants (ADIN ou LP3D)
    Dans ce fichier, on note aussi la réception des docs ou des signatures Studea pour filtrer les mails à envoyer

    TODO :
       - Pouvoir faire un RDV Skpe (lié à classe RDV_outlook)
       - Ouvrir un mail vide, on le remplit, on sauve, et ça envoie à tous les apprentis et tuteurs
       - Heures RDV avec rappel pour appréciations = 7h00-7h30 → Faire 8h-8h
       - Si pb pour aller rechercher la PJ : faire un file picker
       - Quid si mails tutoiement

    Validation :
       - mail de premier contact (à partir d'un modèle .msg) → Oui mais TODO signature à enlever
       - préparer les RDV Outlook pour les entretiens (à partir de modèles .oft) ;
       - faire les mails de relance pour des documents ou Studea.

    """
    
    @dataclass
    class PropEntretien:
        sujet: str
        chemin_modele: str
        duree: timedelta
        date_debut:datetime
        categorie: str = "FI"
        
        def __post_init__(self):
            # Normaliser en datetime
            if isinstance(self.date_debut, date) and not isinstance(self.date_debut, datetime):
                self.date_debut = datetime.combine(self.date_debut, time(9, 0))

    @dataclass
    class PropFichierARenvoyer:
        sujet: str
        chemin_fichier: str
        periode:str
        deadline_retour:Optional[datetime]=None

    def __init__(self) -> None:
        
        # Année en cours
        self._annee_scolaire:str

        # Infos mails
        self._prefixe_sujet:str
        self._mail_responsables_univ:Optional[str] = None
        self._chemin_modele_mail_priseContact:Optional[str] = None
        self._envoyer_mail:bool = False

        # Fichier Excel avec les informations des étudiants
        self._chemin_fichier_etudiants:str
        self._nom_onglet:str
        self._fe_etudiants:FichierExcel
        self._df_etudiants:DataFrame
        
        self._colonnes_fe_etudiants = [] # Colonnes à récupérer du fichier Excel étudiant, hors relances, envoi de docs et entretiens
        
        # Entretiens, relances, fichiers à envoyer
        self._relances:list[str] = []
        self._entretiens:list[Traiter_contactsApprentis.PropEntretien] = []
        self._fichiersARenvoyer:list[Traiter_contactsApprentis.PropFichierARenvoyer] = []
  
    @classmethod
    def depuis_nomUniversite(cls, nomUniversite:str) -> Traiter_contactsApprentis:

        # Pour charger les bonnes propriétés selon l'université
        match nomUniversite:
            case "UGA":
                cfg = importlib.import_module("user_config_UGA")
            case "LP3D":
                cfg = importlib.import_module("user_config_LP3D")
            case _:
                log_erreur("Code université faux")
        
        # Initialisation de l'instance
        instance = cls()  

        # On initialises les variables d'instance avec les données de la config UGA ou LP3D
        # Année en cours
        instance._annee_scolaire = cfg.annee_scolaire

        # Infos mails
        instance._prefixe_sujet = cfg.prefixe_sujet
        instance._mail_responsables_univ = cfg.mail_responsables_univ
        instance._chemin_modele_mail_priseContact = cfg.chemin_modele_mail_priseContact
        instance._envoyer_mail = cfg.envoyer_mail

        # Fichier Excel avec les informations des étudiants
        instance._chemin_fichier_etudiants = cfg.chemin_fichier_etudiants
        instance._nom_onglet = cfg.nom_onglet
        
        # RDV, fiches d'évaluation et relances
        instance._entretiens = [instance.PropEntretien(**iEntretien) for iEntretien in cfg.entretiens]
        instance._fichiersARenvoyer = [instance.PropFichierARenvoyer(**iFichiersARenvoyer) for iFichiersARenvoyer in cfg.fichiersARenvoyer]
        instance._relances = cfg.relances + [iEntretien.sujet for iEntretien in instance._entretiens] + [iFichiersARenvoyer.sujet for iFichiersARenvoyer in instance._fichiersARenvoyer] # On ajoute aux relances les RDV entretiens (on relancera pour les signatures Studea/Lea) + les fiches d'évaluation (à relancer également)
        
        # Colonnes à récupérer du fichier Excel étudiant avec hors relances et entretiens
        instance._colonnes_fe_etudiants = cfg.colonnes_fe_etudiants + [iRelance for iRelance in instance._relances] # On ajoute à la suite des colonnes Excel à récupérer les relances qui comprend les relances initiales + les entretiens
        # On ajoute à la suite des colonnes Excel à récupérer les relances qui comprend les relances initiales + les entretiens
        #colonnes_fe_etudiants.append(iRelance for iRelance in instance._relances)
        


        # On charge l'Excel qui contient les coordonnées des étudiants + traçage relances
        instance._fe_etudiants = FichierExcel.depuis_fichier(
            chemin_fichier = instance._chemin_fichier_etudiants,
            nom_onglet = instance._nom_onglet
        )
        # On réduit le DataFrame aux informations qui nous sont utiles
        instance._df_etudiants = instance._fe_etudiants._tableaux[instance._nom_onglet]._df[instance._colonnes_fe_etudiants] # On ne garde que les colonnes qui nous intéressent mais attention ça reste une vue dont les modifications affectent le dataframe initial
        instance._df_etudiants = instance._df_etudiants[instance._df_etudiants["Ma fonction de suivi de l'alternant"] == "Tuteur"]


        # Tests :
        #instance._df_etudiants = instance._df_etudiants.head(1)

        return instance




    def creer_mails_contactInitial(self, chemin_modele_mail_priseContact:Optional[str]=None) -> None:
        """
        Prépare un mail de contact initial des alternants.
            - Soit à partir de self._chemin_modele_mail_priseContact
            - Soit à partir d'un modèle .msg donné en argument
        """

        if chemin_modele_mail_priseContact is None:
            chemin_modele_mail_priseContact = self._chemin_modele_mail_priseContact

        # Prise de contact
        # On ouvre un modèle de mail existant et on prépare un mail pour chaque apprenti
        for _, ligne in self._df_etudiants.iterrows():
            destinataire = ligne["Mail apprenti"]
            copie = ligne["Mail TE"]

            # Vérification basique (on peut aussi tester si c'est une adresse mail valide)
            if pd.notna(destinataire):
                Mail.depuis_modele(
                    chemin_modele=chemin_modele_mail_priseContact,
                    destinataires=destinataire,
                    copies=copie
                )

    def creer_mails_ficheEvaluation(self, date_deadline_retourFiche:date=None) -> None:
        """
        Prépare un mail pour envoyer les fiches d'évaluation aux tuteurs entreprises.
            - mail avec pj
            - rdv à deadline retour
        """
        periode = self.periode_scolaire_UGA()
        if periode == "mi-année" :
            prop = self._fichiersARenvoyer[0]
        else:
            prop = self._fichiersARenvoyer[1]

        
        # Si non existant, on demande la deadline à l'utilisateur :
        if date_deadline_retourFiche is None:
            date_deadline_retourFiche = prop.deadline_retour

        # On prépare le mail et le RDV outlook pour chaque tuteur entreprise
        for _, ligne in self._df_etudiants.iterrows():
            self._creer_mail_ficheEvaluation(ligne, date_deadline_retourFiche)

    def creer_rdv(self, prop:PropEntretien) -> None:
        # On ouvre un modèle de mail existant et on prépare un mail pour chaque apprenti
        for _, ligne in self._df_etudiants.iterrows(): # Si je,veux boucler que le 1er élément : self._df_etudiants.head(1).iterrows()
            # On prépare l'adaptation du RDV
            participants_obligatoires = ";".join([
                    ligne["Mail apprenti"],
                    ligne["Mail TE"]
                ])

            # On créée le RDV à partir du modèle
            RDV_Outlook.depuis_modele_oft(
                chemin_modele=prop.chemin_modele,
                sujet = f"{self._prefixe_sujet} - {prop.sujet} - {ligne['Prénom']} {ligne['Nom']}",
                lieu = f'{ligne["Entreprise "]} {ligne["Lieu entreprise"]}',
                date_debut = prop.date_debut,
                duree = prop.duree,
                categorie = prop.categorie, # Si je veux en mettre plusieurs, je peux mettre "Urgent, Projet A"
                participants_obligatoires=participants_obligatoires
                )

    @staticmethod
    def creer_html_mail_relance_BAK(choix:int, ligne:pd.Series, lettreInterlocuteur:str) -> str|None:
        corps_html = """
        <html>
        <body>
            <p>Bonjour,</p>
            <p>Dans le cadre du suivi de l'alternance il vous manque certaines actions.</p>
            <p>S'il vous plaît, est-ce que vous pouvez <span style="color: red; font-weight: bold;">au plus tôt</span> :</p>

            <ul>
        """

        len_corps_html_ini = len(corps_html)
        
        if choix in (1, 2, 3, 4, 5, 6) :
            if lettreInterlocuteur not in str(ligne.get("Engagement des parties") or "").strip():
                #print("Relance engagement des parties")
                corps_html += """<li>Signer dans Studea la section <strong>"Engagement des parties"</strong></li>\n"""
        if choix in (2, 3, 4, 5, 6) :
            if lettreInterlocuteur not in str(ligne.get("Entretien initial") or "").strip():
                #print("Relance entretien prise de fonction")
                corps_html += """<li>Compléter et signer dans Studea <strong>l'entretien de prise de fonction</strong> dans la partie "Visites en entreprise"</li>\n"""
        if choix in (3, 4, 5, 6) :
            if lettreInterlocuteur not in str(ligne.get("1ère visite") or "").strip():
                #print("Relance 1ère visite")
                corps_html += """<li>Compléter et signer dans Studea <strong>le relevé de conclusion de la première visite en entreprise</strong> dans la partie "Visites en entreprise"</li>\n"""
        if choix in (4, 5, 6) :
            if lettreInterlocuteur not in str(ligne.get("2ème visite") or "").strip():
                #print("Relance 2ème visite")
                corps_html += """<li>Compléter et signer dans Studea <strong>le relevé de conclusion de la deuxième visite en entreprise</strong> dans la partie "Visites en entreprise"</li>\n"""
        if choix == 5:
            if lettreInterlocuteur not in str(ligne.get("Fiche évaluation 1") or "").strip():
                #print("Relance fiche d'évaluation 1")
                corps_html += """<li>Compléter et nous renvoyer <span style="color: red; font-weight: bold;">la fiche d'évaluation du 1er semestre nécessaire pour la soutenance de début janvier</span></li>\n"""
        if choix == 6:
            if lettreInterlocuteur not in str(ligne.get("Fiche évaluation 2") or "").strip():
                #print("Relance fiche d'évaluation 2")
                corps_html += """<li>Compléter et nous renvoyer <span style="color: red; font-weight: bold;">la fiche d'évaluation du 2ème semestre nécessaire pour la soutenance finale</span></li>\n"""

        if len(corps_html) > len_corps_html_ini:
            corps_html += """
                </ul>

                <p>Je vous remercie, passez une excellente fin de journée,</p>
            </body>
            </html>"""
            return corps_html
        else:
            return None      

    def creer_html_mail_ficheEvaluation(self, ligne:pd.Series, date_deadline_retourFiche:date) -> str:
        
        periode = self.periode_scolaire_UGA(date_deadline_retourFiche)
        date_str = format_date(date_deadline_retourFiche, "d MMMM", locale="fr")

        corps_html = """
        <html>
        <body>
            <p>Bonjour,</p>
            <p>Dans le cadre de la notation de """ + f'{ligne["Prénom"]} {ligne["Nom"]}' + f" lors de sa prochaine soutenance de {periode}, " + """est-ce que vous pouvez compléter et signer le document en PJ s'il vous plait ? Ce document restera strictement confidentiel à l'équipe pédagogique et ne sera pas divulgué à votre apprenti.</p>
            <p>Il sera à retourner à moi-même en mettant en copie """ + f"{self._mail_responsables_univ}" + f" pour le <span style='color: red; font-weight: bold;'>{date_str} au plus tard</span> " + """.</p>
            <p>Je vous envoie un avis de rdv pour faire office de pense-bête. Il sera mis au """ + f"{date_str}" + """ avec un rappel une semaine avant (mais en mode « libre » histoire de ne pas bloquer le créneau sur votre agenda donc vous pouvez l’accepter dans risque).</p>
            """

        if periode == "fin d'année" :
            corps_html += """<p>Si vous êtes en vacances durant cette période, veillez me l’envoyer avant de profiter de votre repos mérité !</p>"""

        corps_html += """
                <p>Je vous remercie, passez une excellente fin de journée,</p>
            </body>
            </html>"""
       
        return corps_html

    def _creer_mail_ficheEvaluation(self, ligne:pd.Series, date_deadline_retourFiche:date):
        """
        Pour un alternant (fonction appelée depuis une boucle) : 
           - envoie un mail avec la fiche d'évaluation en pj
           - envoie un rdv outlook avec rappel 1 semaine avant la date butoir
        """
        corps_html = self.creer_html_mail_ficheEvaluation(ligne, date_deadline_retourFiche)
        pieces_jointes = self.defini_pj_ficheEvaluation_UGA()
        sujet = self._prefixe_sujet + " - Fiche d'évaluation à compléter et retourner"
        
        Mail.creer_mail(
            destinataires = ligne["Mail TE"],
            sujet = sujet,
            corps_html = corps_html,
            pieces_jointes = pieces_jointes,
            envoyer_mail = self._envoyer_mail
        )

        date_debut = datetime.combine(date_deadline_retourFiche, time(hour=8, minute=0))
        date_fin = date_debut + timedelta(minutes=30)
        RDV_Outlook.depuis_html(
            sujet = sujet,
            date_debut = date_debut,
            date_fin = date_fin,
            categorie = "FI",
            html = corps_html,
            participants_obligatoires = ligne["Mail TE"],
            pieces_jointes = pieces_jointes,
            disponibilite = "Libre",
            importance = 2,
            rappel_active = True,
            rappel_minutes = RDV_Outlook.calculer_rappel_minutes(jours=7),
            envoyer = self._envoyer_mail
        )

    def creer_mails_relances(self, relance: str):
        """
        Envoie les mails de relance pour un intitulé donné (élément de self._relances).
        """
        try:
            ind = self._relances.index(relance) + 1  # convertit en indice 1-based
        except ValueError:
            raise ValueError(f"Relance inconnue: {relance!r}")

        # Pièces jointes seulement pour les fiches d’évaluation
        if "évaluation" in self._relances[ind-1]:
            pieces_jointes = self.defini_pj_ficheEvaluation_UGA()
        else:
            pieces_jointes = None

        for _, ligne in self._df_etudiants.iterrows():
            # Apprenti
            corps_html = self.creer_html_mail_relance(ind, ligne, "A")
            if corps_html:
                Mail.creer_mail(
                    destinataires=ligne["Mail apprenti"],
                    sujet=self._prefixe_sujet + "  - Relance actions suivis de l'alternance",
                    corps_html=corps_html,
                    pieces_jointes=pieces_jointes,
                    envoyer_mail=self._envoyer_mail,
                )

            # Tuteur entreprise
            corps_html = self.creer_html_mail_relance(ind, ligne, "T")
            if corps_html:
                Mail.creer_mail(
                    destinataires=ligne["Mail TE"],
                    sujet=self._prefixe_sujet + " - Relance actions suivis de l'alternance",
                    corps_html=corps_html,
                    pieces_jointes=pieces_jointes,
                    envoyer_mail=self._envoyer_mail,
                )


    def creer_html_mail_relance(self, ind:int, ligne:pd.Series, lettreInterlocuteur:str) -> str|None:
        """
        On choisit par définition que tout rappel d'un certain indice du tableau des rappel implique un rappel des entretiens/trucs précédents
        ind est la valeur sélectionnée par l'utilisateur et commence à 1. Donc sa correspondance dans le tableau self._relances c'est ind-1
        """
        corps_html = """
        <html>
        <body>
            <p>Bonjour,</p>
            <p>Dans le cadre du suivi de l'alternance il vous manque certaines actions.</p>
            <p>S'il vous plaît, est-ce que vous pouvez <span style="color: red; font-weight: bold;">au plus tôt</span> :</p>

            <ul>
        """

        len_corps_html_ini = len(corps_html)
        #for i, typeRelance in enumerate(self._relances):
        for i in range(ind):
            if lettreInterlocuteur not in str(ligne.get(self._relances[i]) or "").strip():
                if "évaluation" in self._relances[i]: # Cas fiches d'évaluation
                    corps_html += f"<li>Compléter et nous renvoyer <span style=\"color: red; font-weight: bold;\">la fiche d'évaluation nécessaire pour la soutenance de {self.periode_scolaire_UGA()}</span></li>\n"
                else: # Cas studea
                    corps_html += fr"<li>Compléter et signer dans Studea la ligne <strong>{self._relances[i]}</strong></li>"

        
        if len(corps_html) > len_corps_html_ini:
            corps_html += """
                </ul>

                <p>Je vous remercie, passez une excellente fin de journée,</p>
            </body>
            </html>"""
            return corps_html
        else:
            return None      

    def defini_pj_ficheEvaluation_UGA(self) -> str:
        """
        Sélectionne et renvoie le chemin de la bonne fiche d'évaluation en fonction de la période (mi-année ou fin d'année) → Ancienne méthode
        """
        periode = self.periode_scolaire_UGA()
        fichier = next(
            (f for f in self._fichiersARenvoyer if f.periode == periode),
            None  # valeur par défaut si rien trouvé
        )
        
        """
        if (periode == "mi-année"):
            #print("On est entre septembre et février (inclus)")
            pieces_jointes = self._ficheEvaluation1
        else:
            #print("On est entre mars et août")
            pieces_jointes = self._ficheEvaluation2
        """
        
        return fichier.chemin_fichier

    def rdv_prise_de_fonction_BAK(self, chemin_modele_rdv:str) -> None:

        # Durée type : 1h
        duree = timedelta(hours=0, minutes=45) #timedelta(hours=1, minutes=30)

        # On ouvre un modèle de mail existant et on prépare un mail pour chaque apprenti
        for _, ligne in self._df_etudiants.iterrows():

            # On prépare l'adaptation du RDV
            sujet = f"Master IN - Suivi d'alternance - Entretien de prise de fonction - {ligne['Prénom']} {ligne['Nom']}"
            #date_debut = 
            #date_fin = 
            lieu = f'{ligne["Entreprise "]} {ligne["Lieu entreprise"]}'
            participants_obligatoires = ";".join([
                    ligne["Mail apprenti"],
                    ligne["Mail TE"]
                ])

            # On créée le RDV à partir du modèle
            RDV_Outlook.depuis_modele_oft(
                chemin_modele=chemin_modele_rdv,
                sujet = sujet,
                lieu = lieu,
                duree = duree,
                categorie = "FI", # Si je veux en mettre plusieurs, je peux mettre "Urgent, Projet A"
                participants_obligatoires=participants_obligatoires
                )

    @staticmethod
    def periode_scolaire_UGA(date_input:date=None) -> str:
        """
        La fonction emploie soit une date rentrée en argument, soit la date du jour.
        
        Renvoie "mi-année" si date est entre septembre et février.
        Renvoie "fin d'année" sinon
        """
        # Période à définir pour la fiche d'évaluation à employer
        if date_input is None:
            mois = date.today().month  # 1=janvier, ..., 12=décembre
        else:
            mois = date_input.month

        if ((9 <= mois) or (mois <= 2)):  # septembre (9) → février (2)
            #print("On est entre septembre et février (inclus)")
            periode = "mi-année"
        else:
            #print("On est entre mars et août")
            periode = "fin d'année"        
        return periode


    @staticmethod
    def _diagnostiquer_msg(chemin_modele: str):
        print("== Diagnostique du fichier .msg ==\n")
        print(f"Chemin du fichier : {chemin_modele}\n")

        # 1. Analyse via extract_msg
        print("--> Analyse avec extract_msg")
        try:
            msg = extract_msg.Message(chemin_modele)
            print("Sujet         :", msg.subject)
            print("Expéditeur    :", msg.sender)
            print("Date          :", msg.date)
            print("Contient HTML :", bool(msg.htmlBody))
            print("Contient Texte:", bool(msg.body))
            if msg.htmlBody:
                print(" - Taille HTMLBody :", len(msg.htmlBody))
                print(" - HTMLBody contient <ul> :", "<ul>" in msg.htmlBody.lower())
                print(" - HTMLBody contient puce (•) :", "•" in msg.htmlBody)
        except Exception as e:
            print("Erreur extract_msg :", e)

        print("\n--> Analyse avec Outlook COM (OpenSharedItem)")
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            session = outlook.Session
            modele = session.OpenSharedItem(os.path.abspath(chemin_modele))

            print("Sujet         :", modele.Subject)
            print("Location      :", modele.Location)
            print("Body (200c)   :", modele.Body[:200] if hasattr(modele, "Body") else "[Aucun]")
            
            # Test HTMLBody
            try:
                html = modele.HTMLBody
                print("Contient HTMLBody :", True)
                print(" - Taille HTMLBody :", len(html))
                print(" - HTMLBody contient <ul> :", "<ul>" in html.lower())
                print(" - HTMLBody contient puce (•) :", "•" in html)
            except AttributeError:
                print("Contient HTMLBody :", False)
            
            # Test RTFBody si dispo
            if hasattr(modele, "RTFBody"):
                print("Contient RTFBody :", True)
            else:
                print("Contient RTFBody :", False)

        except Exception as e:
            print("Erreur Outlook COM :", e)



