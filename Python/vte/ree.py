
# ======================================================================================
# CLASSES LIÉES AUX REE
# ======================================================================================
@dataclass
class TypeIntervenant:
    nom:str
    docs:list[REE.DocREE]

class REE:
    
    @dataclass
    class DocREE:
        frequence_maj:Optional[list[str]] = None
        nom_fichier:Optional[str] = None
        chemin_fichier:Optional[str] = None
        intervenants:Optional[list[str]] = None

    """
    def enregistrer_docRee_dico(nom: str, frequence_maj:Optional[list[str]], nom_fichier:Optional[str] = None) -> DocREE:
        doc = DocREE(nom, frequence_maj, nom_fichier)
        REE._docsREE[nom] = doc
        return doc


    def enregistrer_intervenant_dico(cls, nom: str, docs: list[REE.DocREE]) -> TypeIntervenant:
        ti = cls.TypeIntervenant(nom, docs)
        cls._typesIntervenants[nom] = ti
        return ti"""
    # === VARIABLES DE CLASSE ===
    # --- Paramètres d'environnement
    #_repertoire_documents_ree:str = r"\\harmonie\instn\uem\_Documents_communs\Formations\Formateurs\0.Docs à envoyer"  # Répertoire de la GED où sont 
    #_chemin_mailtype_informationsAdministratives = r"\\harmonie\instn\uem\_Documents_communs\Formations\Formateurs\Mails types\Demande des informations administratives.msg"  # Message type à envoyer aux intervenants
    #_repertoire_sauvegarde_fichiersREE:str = r"\\harmonie\INSTN\UEM\_Documents_communs\Formations\Formateurs\1.Intervenants - Documents administratifs" # Lieu où sauvegarder les fichiers de l'intervenant
    #_chemin_modele_excel_ficheIntervenant:str = r"\\harmonie\INSTN\UEM\_Documents_communs\Formations\Formateurs\P09-Pr01-Qualifier les ressources enseignantes\P09_Pr01_Ta.E_Grille des critères de qualification des compétences_V1.xlsx"  # Fichier Excel à remplir pour Laetitia Da Mota (RH INSTN qui s'occupe de rentrer les REE dans IRIS)
    #_adresse_mail_gestionnaire_ree_INSTN:str = "vacataires.instn@cea.fr"
    #_corps_html_mail_gestionnaire_ree_INSTN:str = "<p>Bonjour Laëtitia,</p><p>Je t’ai mis en PJ les documents pour intégrer/mettre à jour la fiche IRIS de ###.</p><p>Je te remercie, passe une excellente journée,</p>"
    
    # Association colonnes excel avec command control Word
    # La préparation de ce ditionnaire peut être faite avec : ree._generer_dictionnaire_depuis_excel()
    _dict_colExcel_cc_V4:Dict[str, str] = {
        "NOM": "Nom",
        "Pr\u00e9nom": "Prenoms",
        "Dipl\u00f4me ou formation/exp\u00e9rience professionnelle": "Diplome",
        "Dur\u00e9e exp\u00e9rience professionnelle": "DureeExperiencePro",
        "Niveau d'expertise permettant une reconnaissance": "NiveauExpertise",
        "Domaine / Sp\u00e9cialit\u00e9 de l'expertise": "DomaineExpertise",
        "ATTRIBUTION Niveau comp\u00e9tences techniques": None,
        "Combien de jours anim\u00e9s, en moyenne par an": "formation_nbJoursAnimes",
        "Combien de jours de formations suivies en p\u00e9dagogie (=animation)": "formation_nbJoursFormationPedagogie",
        "Profils d'apprenants form\u00e9s": "formation_profilApprenants",
        "Taux consolid\u00e9 de la satisfaction des apprenants relativement \u00e0 l'enseignant-formateur consid\u00e9r\u00e9": None,
        "Estimation par le RP de la capacit\u00e9 de l'enseignant-formateur \u00e0 animer (fond de salle)": None,
        "Outils num\u00e9riques utilis\u00e9s durant les animations r\u00e9alis\u00e9es (serious game, blended-learning\u2026)": "formation_outilsNumeriques",
        "Combien de jours pass\u00e9s en conception de s\u00e9quence de formation, en moyenne par an": "IngPedago_nbJoursConception",
        "Combien de jours de formations suivies en ing\u00e9nierie p\u00e9dagogique (=conception de s\u00e9quences de formation)": "IngPedago_nbJoursFormationIngPedago",
        "Estimation par le RP de la conception de la s\u00e9quence en fonction des objectifs p\u00e9dagogiques fournis par le RP (fond de salle, analyse des supports fournis)": None,
        "Estimation par le RP de la pertinence de l'\u00e9valuation des acquis r\u00e9alis\u00e9e par l'enseignant-formateur sur sa s\u00e9quence (analyse de la progression des apprenants : tests avant/apr\u00e8s)": None,
        "Estimation par le RP de l'utilisation des m\u00e9thodes actives (\u00e9tudes de cas, r\u00e9solution de probl\u00e8mes, classes invers\u00e9es, travaux de groupes\u2026)": None,
        "Combien d'ann\u00e9es d'exp\u00e9rience en conception de dispositifs de formations (=cr\u00e9ation et coordination)": "IngFormation_nbJoursConception",
        "Combien de jours de formations suivies en ing\u00e9nierie de formation (=conception de dispositifs de formation)": "IngFormation_nbJoursFormationIngFormation",
        "Estimation par le chef de projet ou le CUE de la complexit\u00e9 des pr\u00e9c\u00e9dents dispositifs de formation con\u00e7us": None,
        "Profil des apprenants des dispositifs de formations prc\u00e9demment con\u00e7us": "IngFormation_profilApprenants",
        "Estimation par le chef de projet ou le CUE de l'\u00e9valuation des acquis r\u00e9alis\u00e9 dans le dispositif de formation (mesure de la progression des apprenants=estimation de la qualit\u00e9 du dispositif de formation)": None,        
        "Combien d'ann\u00e9es d'exp\u00e9rience en tant que tuteur acad\u00e9mique": "IngFormation_nbAnneesTuteur",
        "Combien de r\u00e9f\u00e9rentiels d'activit\u00e9, de comp\u00e9tence et d'\u00e9valuation r\u00e9alis\u00e9s": "IngCompetences_nbReferentiels",
        "Combien de jours de formations suivies en ing\u00e9nierie de comp\u00e9tences": "IngCompetences_nbJoursFormationIngCompetences",
        "Estimation par la cellule p\u00e9dagogique de DPF de la complexit\u00e9 des pr\u00e9c\u00e9dentes r\u00e9alisations de l'ing\u00e9nieur/consultant en ing\u00e9nierie de comp\u00e9tences (complexit\u00e9 du m\u00e9tier et de son environnement : risques, r\u00e9glementation...)": None,
        "ATTRIBUTION Niveau comp\u00e9tences p\u00e9dagogiques": None,
        "Evaluation CECRL ou \u00e9quivalence TOEIC, TOEFL": "ResultatLangue1",
        "ATTRIBUTION Niveau comp\u00e9tences linguistiques": None,
        "Curriculum vitae": None
    }

    _dict_colExcel_cc_versionVTE_refusee:Dict[str, str] = {
        "NOM": "Nom",
        "Pr\u00e9nom": "Prenoms",
        "Dipl\u00f4me ou formation/exp\u00e9rience professionnelle": "Diplome",
        "Dur\u00e9e exp\u00e9rience professionnelle": "DureeExperiencePro",
        "Niveau d'expertise permettant une reconnaissance": "NiveauExpertise",
        "Domaine / Sp\u00e9cialit\u00e9 de l'expertise": "DomaineExpertise",
        "ATTRIBUTION Niveau comp\u00e9tences techniques": None,
        "Combien de jours anim\u00e9s, en moyenne par an": "formation_nbJoursAnimes",
        "Combien de jours de formations suivies en p\u00e9dagogie (=animation)": "formation_nbJoursFormationPedagogie",
        "Profils d'apprenants form\u00e9s": "formation_profilApprenants",
        "Taux consolid\u00e9 de la satisfaction des apprenants relativement \u00e0 l'enseignant-formateur consid\u00e9r\u00e9": None,
        "Estimation par le RP de la capacit\u00e9 de l'enseignant-formateur \u00e0 animer (fond de salle)": None,
        "Outils num\u00e9riques utilis\u00e9s durant les animations r\u00e9alis\u00e9es (serious game, blended-learning\u2026)": "formation_outilsNumeriques",
        "Combien de jours pass\u00e9s en conception de s\u00e9quence de formation, en moyenne par an": "IngPedago_nbJoursConception",
        "Combien de jours de formations suivies en ing\u00e9nierie p\u00e9dagogique (=conception de s\u00e9quences de formation)": "IngPedago_nbJoursFormationIngPedago",
        "Estimation par le RP de la conception de la s\u00e9quence en fonction des objectifs p\u00e9dagogiques fournis par le RP (fond de salle, analyse des supports fournis)": None,
        "Estimation par le RP de la pertinence de l'\u00e9valuation des acquis r\u00e9alis\u00e9e par l'enseignant-formateur sur sa s\u00e9quence (analyse de la progression des apprenants : tests avant/apr\u00e8s)": None,
        "Estimation par le RP de l'utilisation des m\u00e9thodes actives (\u00e9tudes de cas, r\u00e9solution de probl\u00e8mes, classes invers\u00e9es, travaux de groupes\u2026)": None,
        "Combien d'ann\u00e9es d'exp\u00e9rience en conception de dispositifs de formations (=cr\u00e9ation et coordination)": "IngFormation_nbJoursConception",
        "Combien de jours de formations suivies en ing\u00e9nierie de formation (=conception de dispositifs de formation)": "IngFormation_nbJoursFormationIngFormation",
        "Estimation par le chef de projet ou le CUE de la complexit\u00e9 des pr\u00e9c\u00e9dents dispositifs de formation con\u00e7us": None,
        "Profil des apprenants des dispositifs de formations prc\u00e9demment con\u00e7us": "IngFormation_profilApprenants",
        "Estimation par le chef de projet ou le CUE de l'\u00e9valuation des acquis r\u00e9alis\u00e9 dans le dispositif de formation (mesure de la progression des apprenants=estimation de la qualit\u00e9 du dispositif de formation)": None,        
        "Combien d'ann\u00e9es d'exp\u00e9rience en tant que tuteur acad\u00e9mique": "IngFormation_nbAnneesTuteur",
        "Combien de r\u00e9f\u00e9rentiels d'activit\u00e9, de comp\u00e9tence et d'\u00e9valuation r\u00e9alis\u00e9s": "IngCompetences_nbReferentiels",
        "Combien de jours de formations suivies en ing\u00e9nierie de comp\u00e9tences": "IngCompetences_nbJoursFormationIngCompetences",
        "Estimation par la cellule p\u00e9dagogique de DPF de la complexit\u00e9 des pr\u00e9c\u00e9dentes r\u00e9alisations de l'ing\u00e9nieur/consultant en ing\u00e9nierie de comp\u00e9tences (complexit\u00e9 du m\u00e9tier et de son environnement : risques, r\u00e9glementation...)": None,
        "ATTRIBUTION Niveau comp\u00e9tences p\u00e9dagogiques": None,
        "Evaluation CECRL ou \u00e9quivalence TOEIC, TOEFL": "ResultatLangue2",
        "ATTRIBUTION Niveau comp\u00e9tences linguistiques": None,
        "Curriculum vitae": None
    }

    # --- Paramètres utilisateur
    # Documents à envoyer / demander
    _docsREE:dict[DocREE] ={
        "Fiche administrative" : DocREE(
            nom_fichier=r"P09-Pr01-F01 - Fiche administrative vacataire INSTN-V4.docx", 
            frequence_maj=["Initialisation", "Mise à jour"],
            intervenants=["CEA", "vacataire", "contrat spécifique de collaboration", "auto-entrepreneur"]
        ),
        "CV" : DocREE(
            nom_fichier=r"CV-Type.docx", 
            frequence_maj=["Initialisation", "Mise à jour"],
            intervenants=["CEA", "vacataire", "contrat spécifique de collaboration", "auto-entrepreneur"]
        ),
        "RIB" : DocREE( 
            frequence_maj=["Initialisation", "Mise à jour"],
            intervenants=["vacataire", "auto-entrepreneur"]
        ),
        "Attestation employeur" : DocREE(
            nom_fichier=r"Attestation employeur.docx", 
            frequence_maj=["Initialisation", "Tous les ans"],
            intervenants=["vacataire", "contrat spécifique de collaboration"]
        ),
        "Attestation sur l'honneur" : DocREE(
            nom_fichier=r"Attestation sur l'honneur profession libérale.docx", 
            frequence_maj=["Initialisation", "Tous les ans"],
            intervenants=["auto-entrepreneur"]
        ),
        "Devis" : DocREE(
            frequence_maj=["Initialisation", "Tous les ans"],
            intervenants=["auto-entrepreneur"]
        ),
        "Bilan pédagogique et financier année n-1" : DocREE(
            frequence_maj=["Initialisation", "Tous les ans"],
            intervenants=["auto-entrepreneur"]
        ),
        "Guide pour l'intervenant" : DocREE(
            nom_fichier=r"Guide pour l'intervenant - 2026.pdf",
            intervenants=["CEA", "vacataire", "contrat spécifique de collaboration", "auto-entrepreneur"]
        ),
        "Lettre REE" : DocREE(
            nom_fichier=r"Lettre REE - 2026.pdf",
            intervenants=["CEA", "vacataire", "contrat spécifique de collaboration", "auto-entrepreneur"]
        ),

    }

    _correspondance_frequence_texte:dict[str] = {
        "Initialisation" : "initialisation du dossier",
        "Mise à jour" : "s'il y a une mise à jour",
        "Tous les ans" : "<strong><u>chaque année civile</u></strong>"
    }
 


    # --- Autres variables de la classe
    #_docsREE:dict[DocREE] = {}  # Dictionnaire des documents (fiche admin, CV...)
    #_typesIntervenants:dict[TypeIntervenant] = {}  # Dictionnaire des types d'intervenant (CEA, vacataire...)


    # === CONSTRUCTEUR ===
    def __init__(self):
        """
        for doc in self._docsREE.values():
            if doc.nom_fichier is not None:
                # accès direct à la variable de classe REE._repertoire_documents_ree
                doc.chemin_fichier = os.path.join(self._repertoire_documents_ree, doc.nom_fichier)"""
        self._repertoire_sauvegarde_fichiersREE:Path = None  # Répertoire de sauvegarde qui sera construit à partir du répertoire défini dans config et du NOM prénoms de la REE

        # Variables créées / employées par traitement retour mail REE
        #self._word_ficheAdministrative:FichierWord

    # === ENVOI MAIL REE ===
    @classmethod
    def envoyerMail_REE(cls, 
        statut:str, 
        destinataire:Optional[Union[str, list[str], pd.Series]] = None, 
        copie:Optional[Union[str, list[str], pd.Series]] = None
    ) -> None:
        # TODO : comment récupérer le statut depuis Excel ?
        # TODO : comment faire une boucle auto sur les personnes à qui envoyer ?
        """
        Prépare et envoie automatiquement un e-mail d'informations administratives
        via Outlook à partir d'un modèle pré-défini (.msg), selon le statut de l'intervenant.

        Cette méthode :
        - récupère la liste des documents requis pour un type d'intervenant (ex. Vacataire, CEA, etc.)
        - construit dynamiquement le corps du message (HTML) listant les documents à fournir
        - ajoute automatiquement les pièces jointes correspondantes
        - ouvre un e-mail Outlook basé sur le modèle type, avec remplacement des balises
            ###statut### et ###listeInfos### dans le corps du message.

        Args:
            statut (str):
                Statut de l'intervenant (ex. `"Vacataire"`, `"CEA"`, `"Auto-entrepreneur"`, ...).
                Sert à filtrer les documents associés à ce type d'intervenant.
            destinataire (Optional[Union[str, list[str], pd.Series]], optional):
                Adresse(s) e-mail des destinataires principaux.
                Peut être une chaîne unique, une liste de chaînes ou une série Pandas.
                Exemple : `"nom.prenom@domaine.com"` ou `["a@x.com", "b@y.com"]`.
            copie (Optional[Union[str, list[str], pd.Series]], optional):
                Adresse(s) e-mail des destinataires en copie (CC).

        Returns:
            None

        Raises:
            FileNotFoundError:
                Si le modèle de mail (_chemin_mailtype_informationsAdministratives) est introuvable.
            Exception:
                Toute erreur lors de la création ou de l'ouverture du mail Outlook.

        Example:
            >>> ree = REE()
            >>> ree.envoyerMail_REE(
            ...     statut="Vacataire",
            ...     destinataire="vacataire@exemple.com",
            ...     copie=["admin@instn.fr", "drh@instn.fr"]
            ... )
            # Ouvre un mail Outlook basé sur le modèle "Demande des informations administratives"
            # avec la liste des documents à fournir par le vacataire et les fichiers joints.
        """


        instance = cls()

        # Faire choix auto pour existant ou nouveau

        # On crée le texte pour listeInfo + la liste des PJ
        listeInfos:str = ""
        listePJ:list[str] = []
        #self._correspondance_frequence_texte
        for nom_doc, doc in instance._docsREE.items():
            if statut in doc.intervenants:
                # Gestion de la liste des infos à afficher dans le mail
                if doc.frequence_maj is not None:
                    # On mappe chaque élément de doc.frequence_maj via le dictionnaire
                    frequences = [
                        instance._correspondance_frequence_texte.get(freq, freq)
                        for freq in doc.frequence_maj or []
                    ]

                    listeInfos += f"<li><strong>{nom_doc}</strong> <em>({', '.join(frequences)})</em></li>\n"  

                # Gestion des PJ à mettre dans le mail
                if doc.nom_fichier is not None:
                    #listePJ.append(os.path.join(self._repertoire_documents_ree, doc.nom_fichier))            
                    listePJ.append(config.REPERTOIRE_DOCUMENTS_REE / doc.nom_fichier)
        
        if listeInfos != "":
            listeInfos = "<ul> "+listeInfos+" </ul>\n"
        #print(listeInfos)



        # Mail en remplaçant les textes avec statut et listeInfos
        Mail.depuis_modele(
            chemin_modele=config.CHEMIN_MAIL_DEMANDE_INFOS_ADMIN_REE,
            destinataires=destinataire,
            copies=copie,
            pieces_jointes=listePJ,
            balises_a_remplacer=[["statut", statut], ["listeInfos", listeInfos]]
        )



    # === RECEPTION / TRAITEMENT DOC REE
    @classmethod
    def traiter_docs_REE(cls) -> None:

        # TODO : mettre à jour le fichier Excel des coordonnées des intervenants
        # TODO : mettre à jour le fichier Excel des AI
        # TODO : il y aura de la recherche approximative de noms à faire

        
        instance = cls()
        #instance._word_ficheAdministrative:FichierWord

        # On ouvre le word et on charge tous les command control (filedialog depuis "Download"). On le ferme
        instance._word_ficheAdministrative = FichierWord.depuisFichier()
        #print(self._word_ficheAdministrative)

        # TODO : Peut-être afficher NOM et prénoms pour que l'utilisateur redéfinisse quel nom et quel prénom écrire (peut-être enlever les prénos en sus)

        # On crée le répertoire dans le répertoire des REE s'il n'existe pas (ou assimilé) (NOM Prénom (Société - AAAA))
        instance._creer_repertoire_REE()
        
        # L'utilisateur sélectionne tous les fichiers de la REE et on les déplace dans le répertoire idoine
        fichiers_sortie = instance._deplacer_fichiers(instance._repertoire_sauvegarde_fichiersREE)
        if fichiers_sortie is None:
            vlog.log_erreur("Aucun fichier sélectionné pour le déplacement dans le répertoire de la REE, sortie de la procédure REE", continuer=True)
            return

        # On emplit le fichier Excel à transférer à Laetitia Da Mota à partir d'un modèle
        instance._remplit_excel_avecInfos_word()


        # On ouvre l'Excel et le Word pour comparaison et adaptations manuelles
        # chemin_word = os.path.join(self._repertoire_sauvegarde_fichiersREE, os.path.basename(self._word_ficheAdministrative._chemin_fichier))
        chemin_word = instance._repertoire_sauvegarde_fichiersREE / instance._word_ficheAdministrative._chemin_fichier.name
        # chemin_excel = os.path.join(self._repertoire_sauvegarde_fichiersREE, os.path.basename(self._chemin_modele_excel_ficheIntervenant))
        chemin_excel = instance._repertoire_sauvegarde_fichiersREE / config.CHEMIN_MODELE_EXCEL_FICHE_INTERVENANT.name
        
        fichiers_sortie.append(chemin_excel)
        ouvrir_word_excel_cote_a_cote(chemin_word, chemin_excel, split_ecranPrincipal=True)  # on peut rajouter split_ecranPrincipal=True

        # On attend que l'utilisateur ait adapté/validé le fichier Excel pour avancer
        input("🕒 Attente pour adaptations de l'Excel.\nAppuyez sur une touche après adaptation/sauvegarde de l'Excel REE pour continuer")

        # Dès que l'Excel est fermé, on prépare le mail pour Laetitia
        instance._envoyer_mail_gestionnaire_ree_instn(pj=fichiers_sortie)  # On peut aussi mettre delai=timedelta(days=30)

        # On met à jour le fichier Excel Liste AI formateurs.xlsx : onglet intervenant, on cherche et remplace la date de validité de l'attestation employeur sinon nouvelle ligne (recopier formule + format)
        # On met à jour le fichier Excel  avec la liste des intervenants :  on cherche et remplace les données mail, tel, Ville, la date de validité de l'attestation employeur... sinon nouvelle ligne (recopier formule + format)        



    # === Méthodes internes à la classe
    def _creer_repertoire_REE(self, test:bool=False) -> None:
        """
        Crée le répertoire du REE sur le réseau local
        """
        # TODO : comment faire si pas de content control ?
        if all(k in self._word_ficheAdministrative.cc for k in ["Nom", "Prenoms", "RaisonSociale"]):
            self._repertoire_sauvegarde_fichiersREE = config.REPERTOIRE_SAUVEGARDE_FICHIERS_REE / f"{self._word_ficheAdministrative.cc['Nom'].upper()} {self._word_ficheAdministrative.cc['Prenoms'].title()} ({self._word_ficheAdministrative.cc['RaisonSociale'] if self._word_ficheAdministrative.cc['RaisonSociale'] != 'Raison sociale employeur principal' else 'CEA'} - {datetime.now().year})"
        else :
            vlog.log_erreur("⚠️ Certaines clés sont manquantes dans cc :", [k for k in ["Nom", "Prenoms", "RaisonSociale"] if k not in self._word_ficheAdministrative.cc])
        #print(self._repertoire_sauvegarde_fichiersREE)
        if not test:
            # os.makedirs(self._repertoire_sauvegarde_fichiersREE, exist_ok=True)
            self._repertoire_sauvegarde_fichiersREE.mkdir(parents=True, exist_ok=True)
        else :
            print(self._repertoire_sauvegarde_fichiersREE)

    def _deplacer_fichiers(self, destination: Path = None) -> list[Path]:
        """
        Ouvre un dialogue pour sélectionner des fichiers, puis les déplace vers un dossier choisi.

        Args:
            destination (str, optional): Chemin du dossier de destination.
                                        Si None, un dialogue s'ouvrira pour le choisir.
        """

        fichiers_sortie:list[Path] = []

        # Fenêtre Tkinter cachée
        root = tk.Tk()
        root.withdraw()

        # Sélection des fichiers à déplacer
        fichiers = filedialog.askopenfilenames(title="Sélectionner les fichiers à déplacer")
        if not fichiers:
            print("Aucun fichier sélectionné.")
            return

        # Sélection du dossier de destination
        if destination is None:
            destination = filedialog.askdirectory(title="Choisir le dossier de destination")
            if not destination:
                print("Aucun dossier de destination sélectionné.")
                return

        # Déplacement de chaque fichier
        for fichier_str in fichiers:
            fichier = Path(fichier_str)
            nom_fichier:Path = fichier.name  # os.path.basename(fichier)
            chemin_destination:Path = destination / nom_fichier  # os.path.join(destination, nom_fichier)

            try:
                shutil.move(fichier, chemin_destination)
                fichiers_sortie.append(chemin_destination)
                print(f"✅ Déplacé : {nom_fichier}")
            except Exception as e:
                print(f"❌ Erreur avec {nom_fichier} : {e}")

        
        #print(fichiers_sortie)
        return fichiers_sortie

    def _remplit_excel_avecInfos_word(self) -> None:
        """
        Remplit le fichier Excel à transférer à Laetitia Da Mota à partir d'un modèle

        Doit avoir lu un word avec les content control en amont
        """
        # On ouvre le fichier Excel à remplir pour Laetitia Da Mota (c'est un modèle, on l'enregistrera avec le bon nom dans le répertoire idoine)
        excel_ficheIntervenant = FichierExcel.depuis_modele(
            chemin_modele = config.CHEMIN_MODELE_EXCEL_FICHE_INTERVENANT,
            chemin_fichier_sauv = self._repertoire_sauvegarde_fichiersREE / config.CHEMIN_MODELE_EXCEL_FICHE_INTERVENANT.name,  # os.path.join(self._repertoire_sauvegarde_fichiersREE, os.path.basename(self._chemin_modele_excel_ficheIntervenant)),
            charger_df = True
        )
        
        #df_REE = excel_ficheIntervenant._tableaux["QualificationsREE"]._df  # Alias
        #print(df_REE)


        # On pré-rempli le fichier Excel fiche intervenant grâce aux contecnt control du word et au dictionnaire
        nouvelle_ligne = {}
        for col_df, cc_key in self._dict_colExcel_cc_V4.items():
            if cc_key is None:
                # Pas de clé correspondante => valeur vide dans la DataFrame
                nouvelle_ligne[col_df] = None
            else:
                # Récupérer la valeur dans le dictionnaire Word, ou None si la clé absente
                valeur = self._word_ficheAdministrative._cc.get(cc_key, None)
                nouvelle_ligne[col_df] = convertir_si_possible(valeur)
                #print(col_df, cc_key, valeur, type(convertir_si_possible(valeur)))

        #print(pd.DataFrame([nouvelle_ligne]))

        # On écrit le dataframe dans le tableau structuré
        excel_ficheIntervenant._tableaux["QualificationsREE"].ecrit_dataFrame_dans_tableauStructure(pd.DataFrame([nouvelle_ligne]), supprimeDonneesEtRemplace=True, remplace_df_par_nouveau=True, copie_formules=True)


        # On sauve la fiche intervenant
        excel_ficheIntervenant.save()
        excel_ficheIntervenant.close()

    def _envoyer_mail_gestionnaire_ree_instn(self, pj:Optional[list[str]] = None, delai: Optional[timedelta] = None):
        """
        Envoie un mail au gestionnaire des REE de l'INSTN (Laëtitia Da Mota)

        Pour les PJ, si elles ne sont pas données en argument, alors on récupère automatiquement tous les documents qui ont été mis dans le répertoire de l'intervenant il y a moins de 'delai'
        """

        if (pj is None) and (delai is not None):
            pj = lister_fichiers_repertoire(
                self._repertoire_sauvegarde_fichiersREE,
                delai=delai,
                inclure_sous_dossiers=True,        # Inclut les sous-dossiers
            )
        

        Mail.creer_mail(
            destinataires=config.ADRESSE_MAIL_GESTIONNAIRE_REE_INSTN,  # self._adresse_mail_gestionnaire_ree_INSTN,
            sujet=f"Documents pour mise à jour IRIS de {self._word_ficheAdministrative.cc['Prenoms'].title()} {self._word_ficheAdministrative.cc['Nom'].upper()}", # Pimper avec le nom de l'intervenant
            corps_html=remplacer_champs(config.CORPS_MAIL_GESTIONNAIRE_REE_INSTN, [["Prenoms", self._word_ficheAdministrative.cc['Prenoms'].title()], ["NOM", self._word_ficheAdministrative.cc['Nom'].upper()]]),  # self._corps_html_mail_gestionnaire_ree_INSTN.replace("###", self._word_ficheAdministrative.cc['Prenoms'].title()+" "+self._word_ficheAdministrative.cc['Nom'].upper()),
            pieces_jointes=pj,
            envoyer_mail=False  # envoie directement sans afficher
        )  

    #  Pour aider à générer le dictionnaire des noms de colonne du fichier Excel REE
    def _generer_dictionnaire_depuis_excel(
        self,
        chemin_fichier:Path=None,
        nomOngletQualifications:str=None,
        nomOngletAssociationCC:str=None,
        nomColonneCC:str=None,
    ) -> None:
        """
        Génère et affiche un dictionnaire Python liant les colonnes Excel de l'onglet 'nomOngletQualifications'
        aux content controls Word listés dans l'onglet 'nomOngletAssociationCC'.
        
        Si un content control est vide, sa valeur sera 'None'.

        Args:
            chemin_fichier (str): Chemin du fichier Excel.
            nomOngletQualifications (str): Nom de l’onglet contenant les colonnes de qualifications.
            nomOngletAssociationCC (str): Nom de l’onglet contenant les correspondances CC.
            nomColonneCC (str): Nom de la colonne dans l’onglet d’association qui contient les noms des content controls.
        """
        import json

        if chemin_fichier is None :
            chemin_fichier = config.CHEMIN_MODELE_EXCEL_FICHE_INTERVENANT  # self._chemin_modele_excel_ficheIntervenant
        if nomOngletQualifications is None :
            nomOngletQualifications = "QualificationsREE"
        if nomOngletAssociationCC is None :
            nomOngletAssociationCC = "AssociationCC"
        if nomColonneCC is None :
            nomColonneCC = "Nom CC"

        # Charger le fichier Excel via ta classe personnalisée
        fe = FichierExcel.depuis_fichier(chemin_fichier=chemin_fichier, charger_df=True)

        # Récupérer le DataFrame des colonnes de qualifications
        df_qualif = fe._tableaux[nomOngletQualifications].df
        noms_colonnes = list(df_qualif.columns)

        # Récupérer le DataFrame contenant les associations
        df_assoc = fe._tableaux[nomOngletAssociationCC].df

        if nomColonneCC not in df_assoc.columns:
            raise ValueError(f"La colonne '{nomColonneCC}' n'existe pas dans l'onglet '{nomOngletAssociationCC}'.")

        # Récupérer la colonne des content controls, en forçant la taille à celle des colonnes
        liste_cc = df_assoc[nomColonneCC].tolist()
        # Compléter si jamais il y a moins de lignes que de colonnes dans le premier onglet
        while len(liste_cc) < len(noms_colonnes):
            liste_cc.append(None)
            raise ValueError(f"Le nombre de lignes de CC '{len(liste_cc)}' est plus petit que le nombre de colonnes du tableau principal '{len(noms_colonnes)}'.")

        # Générer le dictionnaire
        print("mon_dictionnaire = {")
        for i, (nom_col, cc) in enumerate(zip(noms_colonnes, liste_cc)):
            virgule = "," if i < len(noms_colonnes) - 1 else ""
            # Si le CC est vide ou NaN → None
            if cc is None or (isinstance(cc, float) and pd.isna(cc)) or str(cc).strip() == "":
                cc_str = "None"
            else:
                cc_str = json.dumps(str(cc))
            print(f"    {json.dumps(nom_col)}: {cc_str}{virgule}")
        print("}")

