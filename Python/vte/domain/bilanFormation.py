from __future__ import annotations
from abc import ABC, abstractmethod
from functools import cached_property
from pathlib import Path
from typing import Any, Callable, Iterable, Optional, Protocol
from mailmerge import MailMerge
import pandas as pd
from pandas import DataFrame

from vte.core import config
from vte.core.iris_referentiel import get_iris, set_iris_chemin_specifique
from vte.domain.fdc import FdC
from vte.domain.iris import IRIS_sessions, IRIS_ventes
from vte.domain.specs import Specs
from vte.utils.utils import chemin_vers_unc
from vte.utils.utils_instn import construire_chemin_config


# ======================================================================================
# PROTOCOLES
# (pour faire passer les informations des objets parents sans ref circulaires)
# ======================================================================================



class Formation_protocol(Protocol):
    """
    Protocol de Formation : permet de simuler une formation en évitant les références circulaires
    """
    @property
    def trigramme_formation(self) -> str: ...

    @property
    def fdc(self) -> Optional[FdC]: ...

    @property
    def specs(self) -> Optional[Specs]: ...

    def ouvrir_fdc(self, chemin_fdc:Optional[Path|str] = None) -> None: ...

    def ouvrir_specs(self, chemin_specs:Optional[Path|str] = None) -> None:  ...

# ======================================================================================
# CLASSE BilanFormation
# ======================================================================================
class BilanFormation:
    # =====================
    # === CONSTRUCTEURS ===
    # =====================
    def __init__(self, formation:Formation_protocol, annee:int):
        # Données de l'utilisateur
        self._formation:Formation_protocol = formation
        self._annee:int = annee



    # =========================
    # === METHODES INTERNES === 
    # =========================
    def creer_texte_rp(self) -> str:
        """
        Crée le texte des rp afin d'arriver à ceci : "Vincent TESTARD (UEM) ; Antony LEBLED (UECC)"
        
        :return: Un texte correspondant à la description
        :rtype: str
        """
        return BilanFormation.creer_texte_rp(df_sessions=self.df_sessions_filtre_annee)



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
            config.UNITES[groupe_rp["Lieu principal"]] + ")"
        )

        # Joindre avec des virgules
        return ", ".join(infos_rp)

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





    # =========================
    # === GETTERS / SETTERS === 
    # =========================
    # Getters / setters liés à BilanSessions
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


    # Liens avec Formation
    @property
    def trigramme_formation(self) -> str:
        return self._formation.trigramme_formation



    # Liens avec IRIS sessions
    @cached_property
    def iris_sessions(self) -> IRIS_sessions:
        return get_iris(typeExport="Sessions")
  
    @cached_property
    def df_sessions_filtre_annee(self) -> DataFrame:
        """
        Renvoie le dataframe de l'extract IRIS filtré VTE selon plusieurs critères :
            - Statut Session != "Annulée" (toujours) ;
            - Nb. Nommés != 0 (toujours) ;
            - l'année du filtre (année n du bilan de formation) ;
            - le trigramme de la formation traité par le bilan (self.trigramme_formation).
        
        :return: extract IRIS VTE filtré
        :rtype: DataFrame
        """
        return self.iris_sessions.df_filtre_periode(trigramme_formation=self.trigramme_formation, annee=self._annee)
      
    @cached_property
    def df_sessions_filtre_annee_nm1(self) -> DataFrame:
        """
        Renvoie le dataframe de l'extract IRIS filtré VTE selon plusieurs critères :
            - Statut Session != "Annulée" (toujours) ;
            - Nb. Nommés != 0 (toujours) ;
            - l'année du filtre (année n-1 du bilan de formation) ;
            - le trigramme de la formation traité par le bilan (self.trigramme_formation).
        
        :return: extract IRIS VTE filtré
        :rtype: DataFrame
        """
        return self.iris_sessions.df_filtre_periode(trigramme_formation=self.trigramme_formation, annee=self._annee-1)
    
    @cached_property
    def df_session_filtre_ligne_plus_recente(self) -> DataFrame:
        return self.df_sessions_filtre_annee.iloc[0]
    


    @cached_property
    def intitule_formation(self) -> str:
        """
        Intitulé de la formation (prend le nom de la ligne la plus récente pour avoir la dernière mise à jour)
        """
        #return self.df_sessions_filtre_annee["Session"].iloc[0]
        return self.df_session_filtre_ligne_plus_recente[["Session"]]

    @cached_property
    def min_participants_sessions(self) -> str:
        return self.df_session_filtre_ligne_plus_recente[["Min."]]

    @cached_property
    def max_participants_sessions(self) -> str:
        return self.df_session_filtre_ligne_plus_recente[["Max."]]

    @cached_property
    def depassementAutorise_participants_sessions(self) -> str:
        return self.df_session_filtre_ligne_plus_recente[["Dépass. autorisé"]]



    # Liens avec IRIS ventes
    @cached_property
    def iris_ventes(self) -> IRIS_ventes:
        return get_iris(typeExport="Ventes")
  
    @cached_property
    def df_ventes_filtre_annee(self) -> DataFrame:
        """
        Renvoie le dataframe de l'extract IRIS filtré VTE selon plusieurs critères :
            - Statut Session != "Annulée" (toujours) ;
            - l'année du filtre (année n du bilan de formation) ;
            - le trigramme de la formation traité par le bilan (self.trigramme_formation).
        
        :return: extract IRIS ventes filtré
        :rtype: DataFrame
        """
        return self.iris_ventes.df_filtre_periode(trigramme_formation=self.trigramme_formation, annee=self._annee)
      
    @cached_property
    def df_ventes_filtre_annee_nm1(self) -> DataFrame:
        """
        Renvoie le dataframe de l'extract IRIS ventes filtré VTE selon plusieurs critères :
            - Statut Session != "Annulée" (toujours) ;
            - l'année du filtre (année n-1 du bilan de formation) ;
            - le trigramme de la formation traité par le bilan (self.trigramme_formation).
        
        :return: extract IRIS ventes VTE filtré
        :rtype: DataFrame
        """
        return self.iris_ventes.df_filtre_periode(trigramme_formation=self.trigramme_formation, annee=self._annee-1)
    


    @cached_property
    def prix_n_str(self) -> str:
        return self.iris_ventes.prix_formation_annee_str(trigramme_formation=self.trigramme_formation, annee=self._annee)
    
    @cached_property
    def prix_nm1_str(self) -> str:
        return self.iris_ventes.prix_formation_annee_str(trigramme_formation=self.trigramme_formation, annee=self._annee-1)



    # Lien avec FdC et IRIS sessions
    @cached_property
    def max_participants_fdc(self) -> int:
        """
        Renvoie le nombre max de participants (= prévu (FdC) + dépassement autorisé (IRIS Sessions) )

        Returns:
            int: le nombre max de participants (= prévu + dépassement autorisé)
        """
        return self._formation.fdc.nb_participants_prevus + self.depassementAutorise_participants_sessions
    


# ======================================================================================
# CLASSE BilanFormation_generateur_word (générique)
# ======================================================================================
class BilanFormation_generateur_word(ABC):
    """
    Classe abstraite pour construire un bilan de formation (toutes versions confondues)

    NE S'APPELLE PAS DIRECTEMENT : UNIQUEMENT VIA SES FILLES
    """

    CHEMIN_MODELE: Path = None  # à surcharger

    # =====================
    # === CONSTRUCTEUR ===
    # =====================
    def __init__(self, bilanFormation:BilanFormation, chemin_fdc:Optional[Path|str] = None):
        self._bilanFormation:BilanFormation = bilanFormation
        self._champs: dict[str, str] = {}
    
    
        # TODO à virer après les phases de test
        chemin_fdc = Path(chemin_vers_unc(r"P:\FORMATIONS_C\TEL\P05-P06-dossier-conception-referentiel\fiche-de-cout-et-code-de-formation\Fiche de coûts INSTN - TEL - 2025.xlsx"))
        self._bilanFormation._formation.ouvrir_fdc(chemin_fdc=chemin_fdc)
        
        # TODO à virer après les phases de test
        chemin_specs = Path(chemin_vers_unc(r"P:\FORMATIONS_C\TEL\P05-P06-dossier-conception-referentiel\specifications-pedagogiques-et-referentiel\P06-Pr01-F01_Specifications-pedagogiques - TEL - 2025.04.pdf"))
        self._bilanFormation._formation.ouvrir_specs(chemin_specs=chemin_specs)

        # TODO à virer après les phases de test
        chemin_IRIS_ventes:Path = Path(chemin_vers_unc(r"R:\_Echanges\VTE\Prog\IRIS\Extracts complets\R04301_Ventes-COMPLET-2026.02.12.xlsx"))
        set_iris_chemin_specifique(typeExport="Ventes", chemin=chemin_IRIS_ventes)
        
        

        # Ce DataFrame liste les principales infos des différentes sources
        self._recapDonnees:DataFrame = DataFrame(columns=['Source', 'Titre formation', 'Min participants', 'Cible participants', 'Max participants', 'Dépassement autorisé'])

    # =========================
    # === METHODES EXTERNES ===
    # =========================
    def construire(self):
        """
        Méthode principale appelée par BilanFormation.

        Calcule/construit les champs puis le fusionne dans le Word
        """
        self._evaluer_recapDonnees()
        self._construire_champs()
        self._fusionner_word()
        #self._post_traitement() # → Dans BilanFormation

    # =========================
    # === METHODES INTERNES === 
    # =========================
    def _evaluer_recapDonnees(self) -> None:
        """
        On a des données similaires (min, max... dans différentes sources.
        Tant qu'à devoir ouvir ces différentes sources, autant vérifier ces cohérences.
        
        Le DataFrame _recapDonnees liste les principales infos des différentes sources
        """
        
        # Données de IRIS sessions
        ligne_IRIS_sessions = DataFrame({
            'Source': ['IRIS sessions'],
            'Titre formation': [self._bilanFormation.intitule_formation],
            'Min participants': [self._bilanFormation.min_participants_sessions],
            'Cible participants': [self._bilanFormation.max_participants_sessions],
            'Max participants': [self._bilanFormation.max_participants_sessions + self._bilanFormation.depassementAutorise_participants_sessions],
            'Dépassement autorisé': [self._bilanFormation.depassementAutorise_participants_sessions]
        })
        self._recapDonnees = pd.concat([self._recapDonnees, ligne_IRIS_sessions], ignore_index=True)


        # Données de la FdC
        ligne_fdc = pd.DataFrame({
            'Source': ['FdC'],
            'Titre formation': [self._bilanFormation._formation.fdc.nom_formation], 
            'Min participants': [self._bilanFormation._formation.fdc.min_participants], 
            'Cible participants': [self._bilanFormation._formation.fdc.nb_participants_prevus], 
            'Max participants': "", #[self._bilanFormation.max_participants_fdc],
            'Dépassement autorisé': ""
        })
        self._recapDonnees = pd.concat([self._recapDonnees, ligne_fdc], ignore_index=True)

        print("\nRécap des données sur les différents fichiers (vérif. incohérence)")
        print(self._recapDonnees)
        
    @abstractmethod
    def _construire_champs(self) -> None:
        """
        Remplit self._champs
        Doit être surchargée dans les classes filles héritées
        """
        pass

    def _fusionner_word(self) -> None:
        """
        Tout le process pour écrire les champs de fusion dans le bilan de sessions Word :
            - ouvrir le word à partir du modèle ;
            - fusionne les champs de fusion ;
            - crée le répertoire pour le bilan si besoin ;
            - sauvegarde le word.
        """
        # Ouvre le document Word à partir du modèle
        document = MailMerge(self.CHEMIN_MODELE)
        #print(document.get_merge_fields())

        # Fusionne les champs de fusion
        document.merge(**self._champs)

        # Chemin de sauvegarde
        chemin = self._bilanFormation.chemin_word_bilan_formation_output

        # On crée le répertoire pour les bilans de session de cette année s'il n'existe pas
        chemin.parent.mkdir(parents=True, exist_ok=True)

        # On écrit le fichier
        document.write(chemin)

    """
    def _post_traitement(self) -> None:
        "
        Peut être surchargé (ouvrir word, mail…)
        "
        # Afficher le word
        FichierWord.depuisFichier(
            chemin_fichier=self._BilanSessions.chemin_word_bilan_output,
            charger_contentControl=False,
            afficherWord=True
        )

        # Préparer le mail pour le chef d'unité
        self._BilanSessions._envoyer_mail_chef_unite()
    """






# ======================================================================================
# CLASSE Bilan_V3 (spécifique V3)
# ======================================================================================
class Bilan_V3(BilanFormation_generateur_word):

    CHEMIN_MODELE = config.CHEMIN_MODELE_WORD_BILAN_FORMATION

    # =======================
    # === PIPELINE METIER === 
    # =======================
    def _construire_champs(self) -> None:
        """
        Pipeline pour calculer et construire les champs de fusion du Word
        """
        # === Données transverses ===
        

        # === On génère le word ===
        self._construire_entete()
        self._construire_commentaires()
        self._construire_stats()


    # =============================
    # === SOUS-PARTIES DU BILAN === 
    # =============================
    def _construire_entete(self) -> None:
        """
        Construit la partie en-tête du bilan de formation
        """
        # Infos génériques
        self._champs["annee"] = self._bilanFormation._annee
        self._champs["titreFormation"] = self._bilanFormation.intitule_formation
        self._champs["trigramme_formation"] = self._bilanFormation.trigramme_formation
        self._champs["date_creationFormation"] = self._bilanFormation._formation.fdc.date_creationFormation.strftime("%d/%m/%Y")

        # Liste des RP et de leurs lieux
        self._champs["rp"] = self._bilanFormation.creer_texte_rp()
        
        self._champs["rt"] = ""

        
        #self._champs["min_participants_cea"] = self._bilanFormation._formation.fdc.min_participants_cea
        #self._champs["min_participants_ee"] = self._bilanFormation._formation.fdc.min_participants_ee
        self._champs["min_participants"] = self._bilanFormation._formation.fdc.min_participants

        self._champs["prevus_participants"] = self._bilanFormation._formation.fdc.nb_participants_prevus
        self._champs["max_participants"] = self._bilanFormation.max_participants_fdc  # Prévu + dépass autorisé sur IRIS ou IRIS Sessions (Max. + Dépass. autorisé)

        self._champs["date_specs"] = self._bilanFormation._formation.specs.date_derniere_modification_str

    def _construire_tableau(self) -> None:
        # Nombre de sessions (IRIS sessions)
        self._nb_sessions_nm1 = BilanFormation.creer_texte_nb_sessions(self._bilanFormation.df_sessions_filtre_annee_nm1)
        self._nb_sessions_n = BilanFormation.creer_texte_nb_sessions(self._bilanFormation.df_sessions_filtre_annee)

        # Nombre d'apprenants (IRIS sessions)
        self._nb_apprenants_nm1 = BilanFormation.creer_texte_nb_apprenants(self._bilanFormation.df_sessions_filtre_annee_nm1)
        self._nb_apprenants_n = BilanFormation.creer_texte_nb_apprenants(self._bilanFormation.df_sessions_filtre_annee)

        # Prix de référence (IRIS ventes)
        self._prix_nm1 = self._bilanFormation.prix_nm1_str
        self._prix_n = self._bilanFormation.prix_n_str

        # Stats evalStat
        # TODO : j'en suis là
        self._satisfactionGlobale_nm1:float = -1  # EvalStat
        self._satisfactionGlobale_n:float = -1  # EvalStat
        self._satisfactionGlobale_n_com:str = -1  # EvalStat

        self._qualiteAnimations_nm1:float = -1  # EvalStat
        self._qualiteAnimations_n:float = -1  # EvalStat
        self._qualiteAnimations_n_com:str = -1  # EvalStat

        self._qualiteMoyensPedagogiques_nm1:float = -1  # EvalStat
        self._qualiteMoyensPedagogiques_n:float = -1  # EvalStat
        self._qualiteMoyensPedagogiques_n_com:str = -1  # EvalStat


    def _construire_commentaires(self):
        """
        Construit la partie commentaires du bilan de sessions
        """
        commentaires = ""

        # Cas avec aucun CSV dispo pour les stats
        if len(self._bilanSessions.liste_codesIRIS_pour_statsCSV) == 0 :
            vlog.print("Info", f"⚠️  Aucun CSV disponible pour ce bilan : les statistiques des stagiaires ne seront pas évaluées.")
            commentaires = f"\n⚠️  Aucun CSV disponible pour ce bilan : les statistiques des stagiaires ne seront pas évaluées.\n"
        
        # Cas avec certains CSV non dispo pour les stats mais pas tous
        elif len(self._bilanSessions.liste_codesIRIS_pour_enTete) != len(self._bilanSessions.liste_codesIRIS_pour_statsCSV) :

            print(self._bilanSessions.df_stagiaires_filtre_statsCSV_1ligne_par_session)
            print()
            print(self._bilanSessions.df_stagiaires_filtre_statsCSV_1ligne_par_session['N° Session'].tolist())

            commentaires = "\n⚠️  Certaines sessions n'ont pas de CSV exploitables pour les statistiques (cf. liste ci-dessous)."
            commentaires += "\n\nDonnées employées pour les statistiques :"
            commentaires += "\n   • Sessions évaluées : "+"".join(f"\n       - {session}" for session in self._bilanSessions.df_stagiaires_filtre_statsCSV_1ligne_par_session['N° Session'].tolist())
            commentaires += "\n   • Nombre d'apprenants sur ces sessions : " + f"{self._nb_apprenants:.0f}"
            commentaires += "\n   • Nombre de stagiaires ayant formulé des retours : " + f"{self._nb_stagiaires_retours:.0f}"

        # Affichage sessions avec pb CSV
        if len(self._bilanSessions.liste_codesIRIS_avec_pb_CSV) > 0:
            commentaires += "\n\nListe des sessions dont les statistiques n'ont pas pu être évaluées :"
            l_codes_IRIS = self._bilanSessions.liste_codesIRIS_avec_pb_CSV

            for statut_evalStat, statut_pourBilan in self._bilanSessions._mapping_statuts.items():
                if (statut_evalStat not in ["Traité", "Exclu - Code IRIS pas dans Extract IRIS sessions"]) :
                    # On filtre par statut
                    codes_statut = self._bilanSessions._get_codesIRIS_par_statut_evalStat(statut_evalStat)

                    # On fait l'intersection avec ceux qui ont pb CSV
                    codes_finaux = [c for c in codes_statut if c in l_codes_IRIS]

                    l_sessions = (
                        self._bilanSessions.df_sessions_filtre_avec_pb_CSV[self._bilanSessions.df_sessions_filtre_avec_pb_CSV['Code IRIS'].isin(codes_finaux)]
                        .dropna(subset=["N° Session"])
                        .drop_duplicates(subset=['Code IRIS'])
                        ["N° Session"]
                        .tolist()
                        )
                    if len(l_sessions) > 0:
                        commentaires += f"\n   • {statut_pourBilan} :" + "".join(f"\n       - {isession}" for isession in l_sessions)

        # Affichage sessions exclues
        if len(self._bilanSessions.liste_codesIRIS_exclus_totalement) > 0:
            commentaires += "\n\nListe des sessions de la période entièrement exclues du bilan :"
            l_codes_IRIS = self._bilanSessions.liste_codesIRIS_exclus_totalement
            l_sessions = (
                self._bilanSessions.df_sessions_filtre_exclus_totalement[self._bilanSessions.df_sessions_filtre_exclus_totalement['Code IRIS'].isin(l_codes_IRIS)]
                .dropna()
                .drop_duplicates(subset=['Code IRIS'])
                ["N° Session"]
                .tolist()
                )
            if len(l_sessions) > 0:
                commentaires += "".join(f"\n       - {isession}" for isession in l_sessions)

        #vlog.print("Info", f"\n{commentaires}")     

        self._champs["commentairesBilan"] = commentaires

    def _construire_stats(self):
        """
        Construit la partie statistiques du bilan de sessions
        """
        # On n'affecte les champs suivants que si des CSV sont disponibles pour les stats
        if len(self._bilanSessions.liste_codesIRIS_pour_statsCSV) > 0 :
            # Satisfaction globale
            self._champs["satisfactionGlobale_moy"] = self._get_stat_avec_format(
                "Satisfaction globale",
                "Moyenne",
                lambda v: f"{v:.1f}/5"
            )
            self._champs["satisfactionGlobale_com"] = self._get_stat_avec_format(
                "Satisfaction globale",
                "Commentaires",
                lambda v: v.replace("_x000D_", "\n"),
                default=""
            )
            """
            try:
                self._satisfactionGlobale_moy = f'{self._bilanSession._stats_stagiaires["Satisfaction globale"]["Moyenne"]:.1f}/5'
            except:
                self._satisfactionGlobale_moy = "Pas de donnée"      
            try:
                self._satisfactionGlobale_com = self._bilanSession._stats_stagiaires["Satisfaction globale"]["Commentaires"].replace("_x000D_", "\n")
            except:
                pass
            """

            # Recommanderiez-vous + commentaires remarques suggestions
            self._champs["recommandation_moy"] = self._get_stat_avec_format(
                "Recommanderiez-vous cette formation ?",
                "Moyenne",
                lambda v: f"{v/5*100:.0f}%"  # (on divise par 5 car on a un booléen stcké sous forme de note sur 5 : 0 = False, 5 = True)
            )
            self._champs["commentairesRemarquesSuggestions_com"] = self._get_stat_avec_format(
                "Commentaires, remarques, suggestions",
                "Commentaires",
                lambda v: v.replace("_x000D_", "\n"),
                default=""
            )
            """
            try:
                self._recommandation_moy = f'{self._bilanSession._stats_stagiaires["Recommanderiez-vous cette formation ?"]["Moyenne"]/5*100:.0f}%'  # (on divise par 5 car on a un booléen stcké sous forme de note sur 5 : 0 = False, 5 = True)
            except:
                self._recommandation_moy = "Pas de donnée"        
            try:
                self._commentairesRemarquesSuggestions_com = self._bilanSession._stats_stagiaires["Commentaires, remarques, suggestions"]["Commentaires"].replace("_x000D_", "\n")
            except:
                pass
            """

            # Notes inférieures à 3
            stats_sous_3 = {  # Dictionnaire pour les critères dont la moyenne est inférieure à 3 et non exclus (critères dans la liste self._CRITERES_A_ENLEVER)
                critere: valeurs
                for critere, valeurs in self._bilanSessions._stats_stagiaires.items()
                if (
                    critere not in self._bilanSessions._CRITERES_A_ENLEVER
                    and valeurs["Moyenne"] is not None
                    and valeurs["Moyenne"] < 3
                )
            }
            self._champs["evalInf3_val"] = f"{len(stats_sous_3)}"
            self._champs["evalInf3_com"] = "\n".join(
                f"• {clef} ({valeurs['Moyenne']:.1f}) :{valeurs['Commentaires'].replace('•', '\n   -').replace('\n\n', '\n')}"
                for clef, valeurs in stats_sous_3.items()
                )
            """
            try:
                self._evalInf3_val = f"{len(stats_sous_3)}"
            except:
                self._evalInf3_val = "Pas de donnée"
            try:
                self._evalInf3_com = "\n".join(f"• {clef} ({valeurs['Moyenne']:.1f}) :{valeurs['Commentaires'].replace('•', '\n   -').replace('\n\n', '\n')}"
                    for clef, valeurs in stats_sous_3.items()
                )

                #self._evalInf3_com = "\n".join(
                #    valeurs["Commentaires"]
                #    for valeurs in stats_sous_3.values()
                #    if valeurs["Commentaires"]
                #).replace("_x000D_", "\n")
            except:
                pass
            """

            # Taux de retour
            self._champs["tauxRetours_val"] = f"{(self._nb_stagiaires_retours/self._nb_apprenants)*100:.0f}%"
            """
            try:
                self._tauxRetours_val = f"{(self._nb_stagiaires_retours/self._nb_apprenants)*100:.0f}%"
            except:
                self._tauxRetours_val = "Pas de donnée"
            """

        else: # Cas : aucun CSV
            defaut = "Aucun CSV dispo"
            self._champs["satisfactionGlobale_moy"] = defaut
            self._champs["satisfactionGlobale_com"] = defaut
            
            self._champs["recommandation_moy"] = defaut
            self._champs["commentairesRemarquesSuggestions_com"] = defaut
            
            self._champs["evalInf3_val"] = defaut
            self._champs["evalInf3_com"] = defaut
            
            self._champs["tauxRetours_val"] = defaut


    # ====================
    # === MÉTHODES GET === 
    # ====================
    def _get_stat(self, critere: str, champ: str, default=None) -> int|float|str|None:
        """
        Récupère une stat depuis le dictionnaire _stats_stagiaires

        :param critere: Critère de la stat à récupérer
        :type critere: str
        :param champ: Champ de ce critère à récupérer (["Nombre", "Moyenne", "Commentaires"])
        :type champ: str
        :param default: Valeur retournée si ce critère n'existe pas. Défaut = None
        :type default: _type_, optional
        :return: la valeur du champ de ce critère (ex. _stats_stagiaires["Satisfaction globale"]["Moyenne"])
        :rtype: int|float|str|None
        """
        stats = self._bilanSessions._stats_stagiaires.get(critere)
        if not stats:
            return default

        valeur = stats.get(champ)
        return valeur if valeur is not None else default

    def _get_stat_avec_format(self, critere: str, champ: str, format:Callable[[Any], str], default:str="Pas de donnée") -> str:
        """
        Récupère et formatte une stat depuis le dictionnaire _stats_stagiaires

        :param critere: Critère de la stat à récupérer
        :type critere: str
        :param champ: _descChamp de ce critère à récupérer (["Nombre", "Moyenne", "Commentaires"])ription_
        :type champ: str
        :param format: format à appliquer (ex. lambda v: f"{v:.1f}/5" ou lambda v: f"{v/5*100:.0f}%" ou lambda v: v.replace("_x000D_", "\n"))
        :type format: Callable[[Any], str]
        :param default: Valeur renvoyée si aucune donnée. Défaut = "Pas de donnée"
        :type default: str, optional
        :return: une statistique formatée en str
        :rtype: str
        """
        valeur = self._get_stat(critere, champ)

        if valeur is None:
            return default

        try:
            return format(valeur)
        except Exception:
            return default





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


