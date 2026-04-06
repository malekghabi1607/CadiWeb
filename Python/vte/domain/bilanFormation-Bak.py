from __future__ import annotations
from functools import cached_property
from pathlib import Path
from typing import Optional, Protocol

#import pandas as pd
from pandas import DataFrame

from vte.core.iris_referentiel import *
from vte.core import config
from vte.domain.fdc import FdC
from vte.domain.iris import IRIS_sessions
from vte.utils.office import FichierWord
from vte.utils.utils import *
from vte.utils.utils_instn import construire_chemin_config

# TODO : Pour l'instant c'est une classe de traitement. Le jour où j'ai besoin d'ouvrir un EvalStat pour le lire uniquement, prendre modèle sur IRIS avec des classes de lecture et de traitement

# ======================================================================================
# PROTOCOLES
# (pour faire passer les informations des objets parents sans ref circulaires)
# ======================================================================================
class Formation_protocol(Protocol):
    """
    Protocol de Formation : permet de simuler une formation en évitant les références circulaires
    """
    @property
    def trigramme_formation(self) -> Optional[str]: ...

    @property
    def fdc(self) -> Optional[FdC]: ...
    
    @fdc.setter
    def fdc(self, valeur:FdC) -> None: ...

    #@property
    #def eval(self) -> EvalStat_formation|None: ...

# ======================================================================================
# CLASSE BilanFormation
# ======================================================================================
class BilanFormation:
    # === VARIABLES PARTAGÉES ENTRE TOUTES LES INSTANCES
    _iris_sessions:Optional[IRIS_sessions] = None  # Fichier Excel qui contient les extracts IRIS Sessions (ou a minima celles de la période en cours) → Plusieurs bilan peuvent être fait à partir de cet extract, c'est donc une variable de classe
    _iris_ventes:Optional[IRIS_sessions] = None  # Fichier Excel qui contient les extracts IRIS Ventes (ou a minima celles de la période en cours) → Plusieurs bilan peuvent être fait à partir de cet extract, c'est donc une variable de classe

    #_df_sessions_filtre:Optional[pd.DataFrame] = None

    def __init__(self, formation:Formation_protocol, annee:int):
        # Données de l'utilisateur
        self._formation:Formation_protocol = formation
        self._annee:int = annee


        #self._chemin_word_bilan_output: Optional[Path] = None  # Chemin de sauvegarde du bilan de formation



    # =========================
    # === METHODES EXTERNES === 
    # =========================
    @staticmethod
    def construire_chemin_word_bilan_formation_output(trigramme_formation:Optional[str]=None, annee:Optional[int]=None) -> Path:
        """
        Construit le chemin de sortie du bilan de formation (à partir des données de la config CHEMIN_WORD_BILAN_FORMATION_OUTPUT) :
            - le chemin est transformé en unc (s'il y a un raccourci lecteur réseau sur le poste de l'utilisateur on transforme en chemin réseau complet) ;
            - l'utilisateur peut optimiser le chemin (chemin le plus long entre l'attendu et ce qui existe).

        :param trigramme_formation: Trigramme de la formation. Défaut = None
        :type trigramme_formation: Optional[str], optional

        :param annee: Année du bilan de formation. Défaut = None.
        :type annee: Optional[int], optional

        :return: Le chemin de sortie du bilan de formation.
        :rtype: Path
        """
        return construire_chemin_config(
            chemin_a_completer = config.CHEMIN_WORD_BILAN_FORMATION_OUTPUT,
            trigramme_formation = trigramme_formation, 
            annee = annee
        )


    # =========================
    # === GETTERS / SETTERS === 
    # =========================
    @property
    def annee(self) -> Optional[int]:
        return self._annee
    
    @cached_property
    def chemin_word_bilan_formation_output(self) -> Path:
        """
        Renvoie le chemin de sortie du bilan de formation.
        Cette donnée est stockée dans le fichier de config (valeur par défaut).
        
        :return: Chemin de sortie du bilan de formation.
        :rtype: Path
        """
        #return config.format_path(config.CHEMIN_WORD_BILAN_FORMATION_OUTPUT, trigramme_formation=self.trigramme_formation, annee=self._annee)
        return BilanFormation.construire_chemin_word_bilan_formation_output(trigramme_formation=self.trigramme_formation, annee=self._annee)

    #TODO mettre les cached_property (on peut prendre exeemple sur bilan session)
    @property
    def df_sessions_filtre(self) -> DataFrame:
        """
        Renvoie le dataframe de l'extract IRIS filtré VTE selon plusieurs critères :
            - Statut Session != "Annulée" (toujours) ;
            - Nb. Nommés != 0 (toujours) ;
            - l'année de la session (année n du bilan de formation) ;
            - le trigramme de la formation traité par le bilan (self.trigramme_formation).
        
        :return: extract IRIS VTE filtré
        :rtype: DataFrame
        """
        return self._iris_sessions.df_filtre_periode(trigramme_formation=self.trigramme_formation, annee=self._annee)

    @property
    def df_sessions_filtre_nm1(self) -> DataFrame:
        """
        Renvoie le dataframe de l'extract IRIS filtré VTE selon plusieurs critères :
            - Statut Session != "Annulée" (toujours) ;
            - Nb. Nommés != 0 (toujours) ;
            - l'année de la session (année n-1 du bilan de formation) ;
            - le trigramme de la formation traité par le bilan (self.trigramme_formation).
        
        :return: extract IRIS VTE filtré
        :rtype: DataFrame
        """
        return self._iris_sessions.df_filtre_periode(trigramme_formation=self.trigramme_formation, annee=self._annee-1)

    @property
    def trigramme_formation(self) -> str:
        return self._formation.trigramme_formation

    @property
    def fdc(self) -> str:
        return self._formation.fdc

    @fdc.setter
    def fdc(self, valeur:FdC) -> None:
        self.fdc = valeur

class BilanFormation_V3(BilanFormation):
    """
    C'est la classe qui contient tous les éléments de ma formation pour mon bilan

    TODO j'en suis là
    ? Exploiter export formation plutôt que export sessions pour les valeurs par défaut nmin/max...
    """
    def __init__(self, formation:Formation_protocol, annee:int) -> None:
        super().__init__(formation=formation, annee=annee)

        self._recapDonnees:DataFrame = DataFrame(columns=['Source', 'Titre formation', 'Min participants', 'Cible participants', 'Max participants', 'Dépassement autorisé'])

        # Liste des champs de fusion du word (nom variable = nom champ word)
        # On les type comme il faut, on mettra en str + unités dans le merge
        
        # instance._annee:int = annee → Dans constructeur super()

        self._titreFormation:str = "" #   IRIS Sessions (Session) ou FdC V6.1 : C5
        #instance._trigramme_formation:str = trigramme_formation # IRIS Sessions (Trigramme formation) ou FdC V6.1 : C8
        self._date_creationFormation:int = -1  # FdC - V6.1 : C9  # TODO en fait je vais y mettre l'année
        self._rp:str = ""  # Sessions # IRIS Sessions (Nom responsable pédag. + Prénom responsable pédag.)
        self._rt:str = ""
        self._min_participants_cea:int = -1  # FdC - V6.1 : M22
        self._min_participants_ee:int = -1  # FdC - V6.1 : N23 selon CEA intra ou non-CEA ou IRIS Sessions (Min.)
        self._min_participants:int = -1  # FdC - V6.1 : M22 ou N23 selon CEA intra ou non-CEA ou IRIS Sessions (Min.)
        self._prevus_participants:int = -1  # FdC - V6.1 : C17 ou IRIS Sessions (Max.)
        self._max_participants:int = -1  # Prévu + dépass autorisé sur IRIS ou IRIS Sessions (Max. + Dépass. autorisé)

        self._annee_nm1 = self._annee - 1
        self._annee_np1 = self._annee + 1

        self._nb_sessions_nm1:int = -1  # Sessions (str car si multisite il faut spécifier mes différentes valeurs)
        self._nb_sessions_n:int = -1  # Sessions (str car si multisite il faut spécifier mes différentes valeurs)

        self._nb_apprenants_nm1:int = -1  # Sessions (str car si multisite il faut spécifier mes différentes valeurs)
        self._nb_apprenants_n:int = -1  # Sessions (str car si multisite il faut spécifier mes différentes valeurs)

        self._prix_nm1:float = -1  # Extract ventes
        self._prix_n:float = -1  # Extract ventes

        self._satisfactionGlobale_nm1:float = -1  # EvalStat
        self._satisfactionGlobale_n:float = -1  # EvalStat
        self._satisfactionGlobale_n_com:str = -1  # EvalStat

        self._qualiteAnimations_nm1:float = -1  # EvalStat
        self._qualiteAnimations_n:float = -1  # EvalStat
        self._qualiteAnimations_n_com:str = -1  # EvalStat

        self._qualiteMoyensPedagogiques_nm1:float = -1  # EvalStat
        self._qualiteMoyensPedagogiques_n:float = -1  # EvalStat
        self._qualiteMoyensPedagogiques_n_com:str = -1  # EvalStat

       

    @classmethod
    def bilanUnique(cls, formation:Formation_protocol, annee:int, chemin_fdc:Optional[Path]=None) -> None:
        
        instance = cls(formation=formation, annee=annee)

        # On teste la pré-existance du bilan Word sinon on n'exécute pas
        continuer = verifier_existance_fichier(instance.chemin_word_bilan_formation_output)

        if continuer:
            
            # === On charge les éléments dont on va avoir besoin ===
            # IRIS sessions et ventes
            
            # set_iris_chemin_specifique(typeExport="Sessions", chemin=Path(r"C:\..."))  # On redéfinit le chemin d'IRIS Sessions
            instance._iris_sessions = get_iris(typeExport="Sessions")
            instance._iris_ventes = get_iris(typeExport="Ventes")

            # FdC
            if instance.fdc is None:
                instance.fdc = FdC.depuis_chemin(formation=formation, chemin_fdc=chemin_fdc)

            # EvalStat
            # TODO : lesquels ? → Il faudra charger/créer en fonction des besoins après pré-sélection des sessions par l'utilisateur
            
            # Ouverture / création du dataframe de l'extract IRIS sessions filtré selon la période demandée
            # Filtre appliqué dans méthode demande_sessions_a_exclure
            #df_filtre = instance._iris_sessions.df_filtre_periode(trigramme_formation=instance.trigramme_formation, annee=instance._annee)

            # On met à jour df_sessions_filtre selon les sessions que souhaite garder / exclure l'utilisateur
            df_filtre, liste_codes_IRIS_exclus = instance._sessions.demande_sessions_a_exclure(trigramme_formation=instance.trigramme_formation, annee=annee)
            #print("\nÉtat de Excel sessions filtré sur période et trigramme :")
            #pprint(instance._df_sessions_filtre)

            # S'il n'y a plus de session à lire dans _df_sessions_filtre, alors il n'y a plus de raison de faire le bilan
            if len(df_filtre) != 0 :
                # On met à jour l'Excel evalstat de la formation si des sessions demandées par l'utilisateur ne s'y trouvent pas
                instance._maj_evalstat_formation()

                # On calcule les stats
                instance._calculer_stats_criteres()
                #pprint(instance._stats_stagiaires)

                # On construit le bilan de session (bilan de session V3)
                instance._bilanSessionV3()

                # On ouvre le word
                FichierWord.depuisFichier(chemin_fichier=instance._chemin_word_bilan_session_output, charger_contentControl=False, afficherWord=True)

                # On envoie un mail au chef d'unité pour la signature du pdf
                instance._envoyer_mail_chef_unite()


            else:
                vlog.print("Info", f"⚠️ Toutes les sessions sont exclues : il n'y a plus de raison de faire le bilan de session.")
        
        else:
            print("❌  Bilan de session déjà existant → Arrêt du traitement automatique")




        #####
        # Exploitation de l'extract IRIS sessions
        #####

        # Exraire les données de la session la plus récente
        ligne_plus_recente = self._df_sessions_filtre.iloc[0][["Session", "Min.", "Max.", "Dépass. autorisé"]]  # Trouver la ligne avec la date la plus récente
        #print("\nValeurs 'Formation' de la dernière session")
        #print(pd.DataFrame(ligne_plus_recente.transpose()))
        [self._titreFormation, min_participants_sessions, max_participants_sessions, depassementAutorise_participants_sessions] = ligne_plus_recente
        
        # Liste des RP et de leurs lieux
        self._rp = BilanFormation_V3.creer_texte_rp(self._df_sessions_filtre)

        # Nombre de sessions
        self._nb_sessions_nm1 = BilanFormation_V3.creer_texte_nb_sessions(self._df_sessions_filtre_nm1)
        self._nb_sessions_n = BilanFormation_V3.creer_texte_nb_sessions(self._df_sessions_filtre)

        # Nombre d'apprenants
        self._nb_apprenants_nm1 = BilanFormation_V3.creer_texte_nb_apprenants(self._df_sessions_filtre_nm1)
        self._nb_apprenants_n = BilanFormation_V3.creer_texte_nb_apprenants(self._df_sessions_filtre)

        # Pour comparaison données entre les différentes sources
        nouvelle_ligne = pd.DataFrame({
            'Source': ['IRIS sessions'],
            'Titre formation': [self._titreFormation],
            'Min participants': [min_participants_sessions],
            'Cible participants': [max_participants_sessions],
            'Max participants': [max_participants_sessions + depassementAutorise_participants_sessions],
            'Dépassement autorisé': [depassementAutorise_participants_sessions]
        })
        self._recapDonnees = pd.concat([self._recapDonnees, nouvelle_ligne], ignore_index=True)



        #####
        # Exploitation de la fiche de coûts
        #####
        # TODO à virer après les phases de test
        chemin_fdc = Path(chemin_vers_unc(r"P:\FORMATIONS_C\TEL\P05-P06-dossier-conception-referentiel\fiche-de-cout-et-code-de-formation\Fiche de coûts INSTN - TEL - 2025.xlsx"))
        
        # On ouvre la fiche de coûts
        self._fdc = FdC(chemin_fdc=chemin_fdc)       

        # On récupère les informations souhaitées
        self._date_creationFormation = self._fdc.date_creationFormation  # FdC - V6.1 : C9
        self._min_participants_cea = self._fdc.min_participants_cea
        self._min_participants_ee = self._fdc.min_participants_ee
        self._min_participants = self._fdc.min_participants
        self._prevus_participants = self._fdc.prevus_participants
        self._max_participants = self._fdc.max_participants(depassementAutorise=depassementAutorise_participants_sessions)  # Prévu + dépass autorisé sur IRIS ou IRIS Sessions (Max. + Dépass. autorisé)
        

        # Pour comparaison données entre les différentes sources
        nouvelle_ligne = pd.DataFrame({
            'Source': ['FdC'],
            'Titre formation': [self._fdc.nomFormation], 
            'Min participants': [self._min_participants], 
            'Cible participants': [self._prevus_participants], 
            'Max participants': [self._max_participants],
            'Dépassement autorisé': ""
        })
        self._recapDonnees = pd.concat([self._recapDonnees, nouvelle_ligne], ignore_index=True)

        print("\nRécap des données sur les différents fichiers (vérif. incohérence)")
        print(self._recapDonnees)








        #####
        # Ventes
        #####
        """
        Méthodes pour avoir le prix de la session :
        Sur R04301 Ventes : 
        Récupérer pour la session :
           - Date de début
           - N° Session
           - Code IRIS
           - RP
           - Type [Formation Inter, Formation Intra]
           - Lieux Principal
           - Intitulé Client
           - Nb Inscriptions
           - Total HT
           - Type de tarif (Forfait ; Forfait/Pers.)

        On souhaite afficher le tarif EE :
        Si 3 premières lettres intitulé Client = CEA, alors Total HT = Total HT/0.9

        
        Si forfait/pers. alors on calcule le prix unitaire :
        Si forfait/pers., Total HT = Total HT / Nb Inscriptions
        """
        # TODO à virer après les phases de test
        #chemin_IRIS_ventes = Path(chemin_vers_unc(r"R:\_Echanges\VTE\Prog\IRIS\Extracts complets\R04301_Ventes-COMPLET-2026.02.12.xlsx"))

        # On charge le fichier IRIS Ventes si non déjà ouvert
        iris_ventes = get_iris("Ventes", chemin=chemin_IRIS_ventes)  #IRIS.IRIS_traite(nom_typeExport="Ventes", chemin=chemin_IRIS_ventes)
        
        # Filtrer le DataFrame sur le "Trigramme formation" = "TEL" et sur les années n et n-1
        df_filtre = iris_ventes.df[(iris_ventes.df['Trigramme formation'] == self._trigramme_formation) & (iris_ventes.df['Date de début'].dt.year.isin([self._annee-1, self._annee]))]

        # Créer un DataFrame df_travail avec les colonnes spécifiées
        df_travail = df_filtre[['Date de début', 'N° Session', 'Code IRIS', 'RP', 'Type', 'Lieux Principal', 'Intitulé Client', 'Nb Inscriptions', 'Total HT', 'Type tarif']]

        # Ajouter une colonne "CEA" (booléen) pour tester si les 3 premières lettres de "Intitulé Client" = "CEA"
        df_travail['CEA'] = df_travail['Intitulé Client'].str[:3] == 'CEA'

        # Ajouter une colonne "Prix HT EE" (si tarif CEA, alors on divise par 0.9 ; si on est sur un forfait/pers., alors on divise par nombre d'inscriptions)
        df_travail['Prix HT EE'] = df_travail.apply(
            lambda row: (row['Total HT'] / 0.9 if row['CEA'] else row['Total HT']) / row['Nb Inscriptions'] if row['Type tarif'] == 'Forfait/Pers.' else (row['Total HT'] / 0.9 if row['CEA'] else row['Total HT'])
            , axis=1)

        # Arrondir le résultat à l'entier le plus proche
        df_travail['Prix HT EE'] = df_travail['Prix HT EE'].round()

        # Ajouter une colonne "Unité prix" en fonction du "Type de tarif"
        df_travail['Unité prix'] = df_travail['Type tarif'].apply(lambda x: '€ HT (forfait)' if x == 'Forfait' else '€ HT/pers.')


        # Vérification que tous les tarifs sont bien les mêmes
        # J'exclue les lignes si "Prix HT EE" = NaN
        df_travail = df_travail.dropna(subset=['Prix HT EE'])

        # J'exclue les lignes si "Prix HT EE" = 0
        df_travail = df_travail[df_travail['Prix HT EE'] != 0]


        self._prix_nm1 = self._evaluer_prix_annee(df_travail=df_travail, annee=self._annee_nm1)
        self._prix_n = self._evaluer_prix_annee(df_travail=df_travail, annee=self._annee)

        



        #####
        # Exploitation de EvalStat
        #####

        # On récupère la liste des codes IRIS des sessions qui seront traités dans ce bilan

        # On met à jour les ficheirs EvalStat s'ils ne sont pas déjà traités
        fe_evaluations_formation, statuts_csv = EvalStat.depuis_liste_codes_IRIS(self._codes_IRIS, fe_IRIS_sessions=self._sessions.fe)


        self._satisfactionGlobale_nm1:float = -1  # EvalStat
        self._satisfactionGlobale_n:float = -1# EvalStat
        self._satisfactionGlobale_n_com:str = -1# EvalStat

        self._qualiteAnimations_nm1:float = -1  # EvalStat
        self._qualiteAnimations_n:float = -1  # EvalStat
        self._qualiteAnimations_n_com:str = -1  # EvalStat

        self._qualiteMoyensPedagogiques_nm1:float = -1  # EvalStat
        self._qualiteMoyensPedagogiques_n:float = -1  # EvalStat
        self._qualiteMoyensPedagogiques_n_com:str = -1  # EvalStat


        



        #####
        # ? Je ne sais pas où le trouver dans IRIS
        #####
        self._rt:str = ""


    # ===========================
    # ===  METHODES INTERNES  === 
    # ===========================



    def _mergeBilan(self):
        """
        A partir d'un chemin de fichier word avec des champs de fusion, on crée le bilan de formation final en incluant les données à l'intérieur.
        Le fichier output est défini par l'utilisateur

        :param s_word_bilan_input: Chemin du fichier Word contenant les champs de fusion et a employer
        :type s_word_bilan_input: string
        :param s_word_bilan_output: Chemin du fichier Word apres fusion des donnees
        :type s_word_bilan_output: string
        :return: pas de donnee en retour
        :rtype: none

        :Example:

        >>> self.mergeBilan("C:\\Users\\wordIn.docx", "C:\\Users\\wordOut.docx")


        .. seealso:: Rien du tout.
        .. warning:: Rien du tout.
        .. note:: Rien du tout.
        .. todo:: Rien du tout.
        """

        #print(f"BilanFormation lancé avec : trigramme={self._codeFormation}, année={self._annee}")
        document = MailMerge(self._chemin_modele_word_bilan_formation)
        #print(document.get_merge_fields())

        document.merge(
            #Liste des champs de fusion du word
            annee='{:%Y}'.format(self._annee), #instance._annee:int = annee
            titreFormation = f"{self._titreFormation} ({self._trigramme_formation})", #self._titreFormation:str #  Sessions
            #instance._trigramme_formation:str = trigramme_formation
            #self._date_creationFormation:Optional[datetime] = None  # FdC
            #self._rp:str = ""  # Sessions
            #self._rt:str = ""
            #self._min_participants_fdc:int = -1  # FdC
            #self._prevus_participants_fdc:int = -1  # FdC
            #self._max_participants_fdc:int = -1  # FdC

            nb_sessions_nm1 = f"{self.self._nb_sessions_nm1:d}", #self._nb_sessions_nm1:int = -1  # Sessions
            nb_sessions = f"{self.self._nb_sessions:d}", #self._nb_sessions_n:int = -1  # Sessions

            nb_apprenants_nm1 = f"{self.self._nb_apprenants_nm1:d}", #self._nb_apprenants_nm1:int = -1  # Sessions
            nb_apprenants = f"{self.self._nb_apprenants:d}", #self._nb_apprenants_n:int = -1  # Sessions

            #self._prix_nm1:float = -1  # Offre formation / left join avec Sessions sur "Réf. Formation" → On récupère Intitulé de l'offre (des fois plusieurs par trigramme) Type tarif OffreFormation.Montant # Todo rajouter année au traitement / concaténation de OffreFormation
            #self._prix_n:float = -1  # Offre formation / left join avec Sessions sur "Réf. Formation" → On récupère Intitulé de l'offre (des fois plusieurs par trigramme) Type tarif OffreFormation.Montant # Todo rajouter année au traitement / concaténation de OffreFormation

            #self._satisfactionGlobale_nm1:float = -1  # EvalStat
            #self._satisfactionGlobale_n:float = -1  # EvalStat
            #self._satisfactionGlobale_n_com:str = -1  # EvalStat

            #self._qualiteAnimations_nm1:float = -1  # EvalStat
            #self._qualiteAnimations_n:float = -1  # EvalStat
            #self._qualiteAnimations_n_com:str = -1  # EvalStat

            #self._qualiteMoyensPedagogiques_nm1:float = -1  # EvalStat
            #self._qualiteMoyensPedagogiques_n:float = -1  # EvalStat
            #self._qualiteMoyensPedagogiques_n_com:str = -1  # EvalStat

            )
        
        document.write(self._chemin_word_bilan_formation_output)




    @staticmethod
    def _evaluer_prix_annee(df_travail:DataFrame, annee:int) -> str:
        # Année n
        df_travail_n = df_travail[(df_travail['Date de début'].dt.year.isin([annee]))]

        if len(df_travail_n["Prix HT EE"].unique()) == 1:
            prix_annee = f"{df_travail_n.iloc[0]["Prix HT EE"]} {df_travail_n.iloc[0]["Unité prix"]}"
        else :
            print(df_travail_n[['Date de début', 'N° Session','Intitulé Client', 'CEA', 'Nb Inscriptions', 'Total HT', 'Type tarif', 'Prix HT EE', 'Unité prix']])
            # Regrouper les lignes par "Prix HT EE" et "Unité prix"
            df_travail_n_groupe = df_travail_n.groupby(['Prix HT EE', 'Unité prix'])

            # Initialiser une liste pour stocker les lignes de texte
            lignes_texte = []

            # Parcourir chaque groupe
            for (prix, unite), groupe in df_travail_n_groupe:
                # Créer une liste des éléments pour chaque ligne du groupe
                elements = []
                for _, row in groupe.iterrows():
                    element = f"{row['N° Session']}; {row['Date de début']}; {row['Intitulé Client']}; {row['Nb Inscriptions']}; {row['Total HT']}; {row['Type tarif']}"
                    elements.append(element)

                # Créer la ligne de texte pour le groupe
                ligne_texte = f"• {prix} {unite} :\n\t- " + "\n\t- ".join(elements)
                lignes_texte.append(ligne_texte)

            # Joindre toutes les lignes de texte avec des sauts de ligne
            prix_annee = "\n".join(lignes_texte)

        return prix_annee

    @staticmethod
    def creer_texte_nb_sessions(df_sessions:DataFrame) -> str:
        """
        Crée le texte de nb_sessions afin d'arriver à ceci : "6 session(s)\n(3 UEM + 3 UECC))
              
        :param df_sessions: dataframe sessions IRIS à employer (peut être une sous-partie du df original)
        :type df_sessions: DataFrame  
        :return: Un texte correspondant à la description
        :rtype: str
        """
        # Calculer le nombre total de sessions pour tous les lieux
        nombre_sessions_tousLesLieux = df_sessions['Lieu principal'].count()

        # Calculer le nombre de sessions par lieu
        nombre_sessions_par_lieu = df_sessions['Lieu principal'].value_counts()

        # Créer le texte souhaité
        if nombre_sessions_tousLesLieux > 1:
            texte_sessions = "sessions"
        else:
            texte_sessions = "session"

        if len(nombre_sessions_par_lieu) > 1:
            texte_detail = f"\n({' + '.join(f'{count} {config.UNITES[lieu]}' for lieu, count in nombre_sessions_par_lieu.items())})"
        else:
            texte_detail = ""

        texte = f"{nombre_sessions_tousLesLieux} {texte_sessions}{texte_detail}"
        #print(texte)

        return texte

    @staticmethod
    def creer_texte_nb_apprenants(df_sessions:DataFrame) -> str:
        """
        Crée le texte de nb_sessions afin d'arriver à ceci : "16 pers.\n(10 UEM + 6 UECC))
        
        :param df_sessions: dataframe sessions IRIS à employer (peut être une sous-partie du df original)
        :type df_sessions: DataFrame
        :return: Un texte correspondant à la description
        :rtype: str
        """
        # Calculer le nombre total de personnes dans toutes les sessions
        nombre_personnes_total = df_sessions['Nb. Nommés'].sum()

        # Calculer le nombre de personnes par lieu
        nombre_personnes_par_lieu = df_sessions.groupby('Lieu principal')['Nb. Nommés'].sum()

        # Créer le texte souhaité
        texte_detail = f"\n({' + '.join(f'{count} {config.UNITES[lieu]}' for lieu, count in nombre_personnes_par_lieu.items())})" if len(nombre_personnes_par_lieu) > 1 else ""

        texte = f"{nombre_personnes_total} pers.{texte_detail}"
        #print(texte)

        return texte

    @staticmethod
    def creer_texte_rp(df_sessions:DataFrame) -> str:
        """
        Crée le texte des rp afin d'arriver à ceci : "Vincent TESTARD (UEM) ; Antony LEBLED (UECC)"
        
        :param df_sessions: dataframe sessions IRIS à employer (peut être une sous-partie du df original)
        :type df_sessions: DataFrame
        :return: Un texte correspondant à la description
        :rtype: str
        """
        # Pour obtenir la liste des RP et de leurs lieux : grouper et prendre la première occurrence (la plus récente)
        groupe_rp = df_sessions.groupby("Nom responsable pédag.", as_index=False).first()  # Si je ne mets pas le as_index, alors l'index devient le nom du rp et les données ne peuvent plus être filtrées comme des colonnes normales sur ce critère

        # Formater les informations
        infos_rp = (
            groupe_rp["Prénom responsable pédag."] + " " +
            groupe_rp["Nom responsable pédag."] + " (" +
            groupe_rp["Lieu principal"] + ")"
        )

        # Joindre avec des virgules
        return ", ".join(infos_rp)


    # =========================
    # ===  GETTER / SETTER  === 
    # =========================





class BilanFormation_VTE:
    """
    C'est la classe qui contient tous les éléments de ma formation pour mon bilan

    TODO j'en suis là
    Todo Word
    Dans le modèle Word : gérer le lien vers la GED 
    Exploiter EvalStat
    Il y a des trous dans la raquette dans le word de sortie (checkboxes)
    coller des images depuis Excel
    ? Exploiter export formation plutôt que export sessions pour les valeurs par défaut nmin/max...
    """
    def __init__(self, codeFormation:str, annee:int):
        
        self._codeFormation:str = codeFormation
        self._annee:int = annee

        self._sessions_nom_typeExport:str = config.IRIS_SESSIONS._nom_typeExport  #sessions._nom_typeExport #  Provient de la valeur globale sessions
        self._sessions_codeExport:str = config.IRIS_SESSIONS._codeExport #sessions._codeExport #  Provient de la valeur globale sessions
        self._sessions_repertoire:Path = config.IRIS_SESSIONS._output.repertoire #sessions._output.repertoire #  Provient de la valeur globale sessions

        self._chemin_modele_word_bilan_formation = config.CHEMIN_MODELE_WORD_BILAN_FORMATION
        self._chemin_word_bilan_formation_output = config.format_path(config.CHEMIN_WORD_BILAN_FORMATION_OUTPUT, trigramme_formation=codeFormation, annee=annee)


        self._chemin_specsPedagogiques = None

        self._chemin_fdc = None

        self._repertoire_fdc_defaut:Path = config.format_path(config.REPERTOIRE_FDC, trigramme_formation=codeFormation)
        self._repertoire_specsPedagogiques_defaut:Path = config.format_path(config.REPERTOIRE_SPECS, trigramme_formation=codeFormation)
        
        #####
        # Exploitation de l'extract IRIS sessions
        #####

        # J'ouvre un export session de IRIS et load tous ses tableaux structurés dans des DataFrame (inclus dans un FichierExcel)
        chemin_fichier_session = filedialog.askopenfilename(title="Sélectionner l'export " + self._sessions_nom_typeExport + " (" + self._sessions_codeExport + ") Excel à employer", filetype=[("Fichiers Excel", "*.xlsx")], initialdir=self._sessions_repertoire)
        if not chemin_fichier_session:
            log_erreur("click sur cancel du filedialog → Pas de chemin de fichier session")
        self._fe_session = FichierExcel.depuis_fichier(chemin_fichier_session)
        #print(self._fe_session)
        self._df_sessions = self._fe_session._tableaux[self._sessions_nom_typeExport]._df  # Création d'un alias

        # Pour initialiser les valeurs communes, déjà on filtre le Dataframe principal avec le code formation
        #df_filtre = self.__df_sessionsIRIS[(self.__df_sessionsIRIS['Trigramme formation'] == self.__codeFormation) & (self.__df_sessionsIRIS['Année début ses.'] == self.__annee)]
        df_sessions_filtre = self._df_sessions[(self._df_sessions['Trigramme formation'] == self._codeFormation) & (self._df_sessions['Année début ses.'] == self._annee)]
        
        
        # Puis on récupère la dernière ligne pour avoir les valeurs les plus à jour (tri par index)
        [self._titreFormation, self._dureeHeures, self._dureeJours, self._minIRIS, self._maxIRIS, self._depassementAutoriseIRIS] = df_sessions_filtre[["Session", "Durée planif. (H.)", "Durée planif. (J.)", "Min.", "Max.", "Dépass. autorisé"]].iloc[-1]

        # Pour obtenir la liste des RP et de leurs lieux
        self._lieuxFormation = ""
        for rp in df_sessions_filtre["Nom responsable pédag."].unique():
            #print("\n\nRP = " + rp)
            [prenomRP, nomRP, lieuRP] = df_sessions_filtre[self._df_sessions["Nom responsable pédag."] == rp][["Prénom responsable pédag.", "Nom responsable pédag.", "Lieu principal"]].iloc[-1]
            self._lieuxFormation += prenomRP + " " + nomRP + " (" + lieuRP + "), "
        self._lieuxFormation = self._lieuxFormation[:-2]
        #print(self._lieuxFormation)




        #####
        # Exploitation des specs pédagogiques
        #####
        self._chemin_specsPedagogiques = filedialog.askopenfilename(title="Sélectionner les dernières specs pédagogiques", filetype=[("Fichiers PDF", "*.pdf"), ("Documents Word", "*.docx")], initialdir=optimiseCheminRepertoire(self._repertoire_specsPedagogiques_defaut.replace("XXX", self._codeFormation)))
        if not chemin_fichier_session:
            log_erreur("click sur cancel du filedialog → Pas de chemin des specs pédagogiques")
        self._dateSpecs = datetime.fromtimestamp(self._chemin_specsPedagogiques.stat().st_mtime)
        self._sDateSpecs = self._dateSpecs.strftime("%d/%m/%Y")
        #print(self._sDateSpecs)



        #####
        # Exploitation de la fiche de coûts
        #####
        self._chemin_fdc = filedialog.askopenfilename(title="Sélectionner la dernière fiche de coûts", filetype=[("Fichiers Excel", "*.xlsx")], initialdir=optimiseCheminRepertoire(self._repertoire_fdc_defaut.replace("XXX", self._codeFormation)))
        if not chemin_fichier_session:
            log_erreur("click sur cancel du filedialog → Pas de chemin de fiche de coûts")
        self._dateFdC = datetime.fromtimestamp(self._chemin_fdc.stat().st_mtime)
        self._sDateFdC = self._dateFdC.strftime("%d/%m/%Y")
        
        self._df_fdc_infos, self._df_fdc_couts, self._prixVenteRetenuParParticipant, self._dateCreationFormation, self._dureeJours_fdc, self._osThematique, self._nbCible_fcd = lire_fdc(self._chemin_fdc)
        




        #####
        # Exploitation de l'extract IRIS formation TODO
        #####
        #Date création formation
        #Type de reconnaissance (Autre, Certification, Diplôme)
        #Formation habilitante (bool)
        #Reconnu au RNCP
        #Eligible CPF
        #Reconnue au RS


    ### --------------------------------------------------------------------
    #  Méthodes de la classe
    ### --------------------------------------------------------------------
    def mergeBilan(self):
        """
        A partir d'un chemin de fichier word avec des champs de fusion, on crée le bilan de formation final en incluant les données à l'intérieur.
        Le fichier output est défini par l'utilisateur

        :param s_word_bilan_input: Chemin du fichier Word contenant les champs de fusion et a employer
        :type s_word_bilan_input: string
        :param s_word_bilan_output: Chemin du fichier Word apres fusion des donnees
        :type s_word_bilan_output: string
        :return: pas de donnee en retour
        :rtype: none

        :Example:

        >>> self.mergeBilan("C:\\Users\\wordIn.docx", "C:\\Users\\wordOut.docx")


        .. seealso:: Rien du tout.
        .. warning:: Rien du tout.
        .. note:: Rien du tout.
        .. todo:: Rien du tout.
        """

        #print(f"BilanFormation lancé avec : trigramme={self._codeFormation}, année={self._annee}")
        document = MailMerge(self._chemin_modele_word_bilan_formation)
        #print(document.get_merge_fields())

        document.merge(
            annee='{:%Y}'.format(date.today()),
            codeFormation=self._codeFormation,
            titreFormation=self._titreFormation,
            
            lienGED = config.format_path(config.REPERTOIRE_FORMATION, trigramme_formation=self._codeFormation),
            osThematique = self._osThematique,
            lieuxFormation = self._lieuxFormation,
            dureeJours = f"{self._dureeJours:.1f}",
            dureeHeures = f"{self._dureeHeures:.2f}",
            dateCreationFormation = str(self._dateCreationFormation), #Actuellement pris depuis fdc, réaffecter à partir extract IRIS formations

            dateSpecs = self._sDateSpecs,

            dateFdC = self._sDateFdC,
            minFdC = f"{self._df_fdc_couts.loc['Valeur fixée', 'Min participants T3']:d} p.",
            cibleFdC = f"{self._nbCible_fcd:d} p.",
            minIRIS = f"{self._minIRIS:d} p.",
            cibleIRIS = f"{self._maxIRIS:d} p.",
            maxIRIS = f"{self._depassementAutoriseIRIS:d} p.",

            PT1_1 = f"{self._df_fdc_couts.loc['T1', 'Montant cible par participant']:.0f} €/p.",
            PT1_2 = f"{self._df_fdc_couts.loc['T1', 'Montant cible par participant et par jour']:.0f} €/j/p.",
            PT1_3 = f"{self._df_fdc_couts.loc['T1', 'Min participants T1']:d} p.",

            PT3_1 = f"{self._df_fdc_couts.loc['1.1xT3', 'Montant cible par participant']:.0f} €/p.",
            PT3_2 = f"{self._df_fdc_couts.loc['1.1xT3', 'Montant cible par participant et par jour']:.0f} €/j/p.",
            PT3_3 = f"{self._df_fdc_couts.loc['1.1xT3', 'Min participants T1']:d} p.",
            PT3_4 = f"{self._df_fdc_couts.loc['1.1xT3', 'Min participants T3']:d} p.",

            PTR_1 = f"{self._df_fdc_couts.loc['Valeur fixée', 'Montant cible par participant']:.0f} €/p.",
            PTR_2 = f"{self._df_fdc_couts.loc['Valeur fixée', 'Montant cible par participant et par jour']:.0f} €/j/p.",
            PTR_3 = f"{self._df_fdc_couts.loc['Valeur fixée', 'Min participants T1']:d} p.",
            PTR_4 = f"{self._df_fdc_couts.loc['Valeur fixée', 'Min participants T3']:d} p."
            
            )
        
        document.write(self._chemin_word_bilan_formation_output)


    ### --------------------------------------------------------------------
    #  Méthodes externes
    ### --------------------------------------------------------------------
    @staticmethod
    def fenetreBilanFormation():
        def valider_champs(*args):
            trig = entry_trigramme.get().strip()
            annee = entry_annee.get().strip()
            bouton_generer.config(
                state="normal" if len(trig) == 3 and annee.isdigit() and len(annee) == 4 else "disabled"
            )

        def generer_bilan_formation(event=None):
            trigrammeFormation = entry_trigramme.get().strip()
            anneeBilan_str = entry_annee.get().strip()
            try:
                anneeBilan = int(anneeBilan_str)
            except ValueError:
                messagebox.showerror("Erreur", "L'année doit être un entier à 4 chiffres.")
                return
            #BilanFormation(trig, annee)
            print(f"BilanFormation lancé avec : trigramme={trigrammeFormation}, année={anneeBilan}")
            bf = BilanFormation(trigrammeFormation, anneeBilan)
            bf.mergeBilan()
            fenetre.destroy()

        def annuler():
            sys.exit()

        # Création de la fenêtre
        fenetre = tk.Tk()
        fenetre.title("Création d'un bilan de formation")
        fenetre.geometry("350x180")
        fenetre.resizable(False, False)

        # Label + champ pour trigramme
        ttk.Label(fenetre, text="Trigramme formation (3 lettres) :").pack(pady=(10, 0))
        entry_trigramme = ttk.Entry(fenetre)
        entry_trigramme.pack(pady=5)

        # Label + champ pour année
        annee_defaut = str(datetime.now().year - 1)
        ttk.Label(fenetre, text="Année du bilan (4 chiffres) :").pack()
        entry_annee = ttk.Entry(fenetre)
        entry_annee.insert(0, annee_defaut)
        entry_annee.pack(pady=5)

        # Boutons
        frame_boutons = ttk.Frame(fenetre)
        frame_boutons.pack(pady=10)

        bouton_generer = ttk.Button(frame_boutons, text="Générer bilan", state="disabled", command=generer_bilan_formation)
        bouton_generer.grid(row=0, column=0, padx=5)

        bouton_annuler = ttk.Button(frame_boutons, text="Annuler", command=annuler)
        bouton_annuler.grid(row=0, column=1, padx=5)

        # Validation en temps réel
        entry_trigramme.bind("<KeyRelease>", valider_champs)
        entry_annee.bind("<KeyRelease>", valider_champs)

        # Entrée = clic sur bouton générer
        fenetre.bind("<Return>", generer_bilan_formation)

        # Échap = fermeture de la fenêtre
        fenetre.bind("<Escape>", lambda e: fenetre.destroy())

        # Lancer la fenêtre
        fenetre.mainloop()


