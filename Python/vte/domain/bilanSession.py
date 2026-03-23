from __future__ import annotations
from pathlib import Path
from pprint import pprint
from typing import Optional, Protocol

from pandas import DataFrame
from mailmerge import MailMerge

from vte.core.iris_referentiel import get_iris
from vte.domain.evalStat import EvalStat, EvalStat_formation
from vte.services.evalStat_services import EvalStat_services
from vte.domain.session import Session
from vte.domain.iris import IRIS, IRIS_traite
from vte.core import config
from vte.utils.office import FichierExcel, FichierWord, Mail
from vte.utils.utils import *

# TODO : Pour l'instant c'est une classe de traitement. Le jour où j'ai besoin d'ouvrir un EvalStat pour le lire uniquement, prendre modèle sur IRIS avec des classes de lecture et de traitement

# Test à faire sur TEL période : 2024

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
    def eval(self) -> EvalStat_formation|None: ...
    
    def get_session_par_codeIRIS(self, code_IRIS:int) -> Optional[Session]: ...


# ======================================================================================
# CLASSE BilanSession
# ======================================================================================

class BilanSession:
    """
    C'est la classe qui contient tous les éléments de ma formation pour mon bilan de session V3
    """
    # === VARIABLES PARTAGÉES ENTRE TOUTES LES INSTANCES


    _CRITERES_A_ENLEVER:list[str] = [  # Critères à ne pas retenir pour le calcul des moyennes < 3
        "Comment avez-vous connu cette formation ?", "Avez-vous d'autres besoins de formation ?", "Commentaires, remarques, suggestions", "Recommanderiez-vous cette formation ?"]

    # Dictionnaire pour mapper les statuts aux clés de self._statuts ["Exploités pour les évaluations (CSV présents)", "Exploités pour les évaluations (CSV présents)", "Exclus des évaluations (problème traitement CSV)", "Exclus des évaluations (CSV manquants)", "Exclus des évaluations (CSV vide / aucun retour)", "Exclus entièrement du bilan (non présent dans IRIS / mauvais code)"]
    _mapping_statuts:dict[str, str] = {
        "Traité": "Exploités pour les évaluations (CSV présents et non vides)",                                  # Exploités pour stats initiales → Dans _demande_sessions_a_exclure
        #  "Traité - CSV déjà dans fichier global": "Exploités pour les évaluations (CSV présents et non vides)",   # Exploités pour les stats stagiaires → Dans _maj_evalstat_formation
        "Exclu - Problème lecture CSV": "Exclus des évaluations (problème traitement CSV)",         # Exclus des évaluations car CSV stagiaires manquants → Dans _maj_evalstat_formation
        "Exclu - Fichier non existant": "Exclus des évaluations (CSV manquants)",                   # Exclus des évaluations car problème au traitement des CSV → Dans _maj_evalstat_formation
        "Exclu - CSV vide / Aucun retour": "Exclus des évaluations (CSV vide / aucun retour)",      # Exclus des évaluations car le CSV est vide (i.e. aucun retour d'utilisateur)
        "Exclu - Code IRIS pas dans Extract IRIS sessions": "Exclus entièrement du bilan (non présent dans IRIS / mauvais code)"  # Exclus entièrement du bilan car sessions non réalisées ou mauvais RP (exclus par l'utilisateur) → Dans _maj_evalstat_formation
    }

    # Dictionnaire pour mapper les statuts qui nécessitent de supprimer le code IRIS des stats générales
    _statut_exclus_entierement:tuple[str] = (
        "Exclu - Code IRIS pas dans Extract IRIS sessions",
    )

    # =====================
    # === CONSTRUCTEURS ===
    # =====================
    def __init__(self, formation:Formation_protocol, annee:int, periode:str="Année") -> None:
        """
        Initialise un bilan de session a minima
        
        Un bilan de session est accolé à une formation car on peut avoir plusieurs sessions dans un bilan de sessions.

        :param formation: objet Formation associé à ce bilan
        :type formation: Formation_protocol
        :param annee: Année du bilan
        :type annee: int
        :param periode: période du bilan (appartient à ["Année", "1er semestre", "2nd semestre"]), defaut = "Année"
        :type periode: str, optional
        """
        # --- Variables qui caractérisent le bilan de sessions
        self._formation = formation
        self._annee: int = annee
        self._periode: str = periode  # ["Année", "1er semestre", "2nd semestre"]
        self._codes_IRIS:list[int] = []

        # --- Variables de traitement ---
        self._statuts:dict[str, list] = {clef: [] for clef in self._mapping_statuts.values()}  # Dictionnaire qui liste les codes IRIS selon chaque statut


        self._stats_stagiaires: dict[str, dict[str, int|float|str|None]] = {}  # Dictionnaire des stats des CSV
        """
        dictionnaire de la forme :
        {
            "Nom du critère": {
                "Nombre": ...,
                "Moyenne": ...,
                "Commentaires": ...
            },
            ...
        }
        """

        # --- Liste des champs de fusion du Word [V3] (pour la fonction .mergefields, il faut des str)
        # TODO : voir si je n'ai pas intérêt à faire un dictionnaire d'un part pour tous les str puis les vrais typages pour les valeurs avant str.
        self._titreFormation:str = ""
        self._codeFormation:str = self.trigramme_formation  # Déjà déclaré pour fonctionnement de la classe avec formation
        self._periodeSessionsEvaluees:str = ""  # f"Session {numSession} uniquement ({moisSession} {instance._annee})", f"{self._periode} {self._annee}"
        self._nbSessionsEvaluees:str = ""
        self._numerosSessions:str = ""
        self._nbApprenants:str = ""
        self._rp:str = ""
        self._af:str = ""

        self._commentairesBilan:str = ""

        self._satisfactionGlobale_moy:str = ""
        self._satisfactionGlobale_com:str = ""
        self._recommandation_moy:str = ""
        self._commentairesRemarquesSuggestions_com:str = ""
        self._evalInf3_val:str = ""
        self._evalInf3_com:str = ""
        self._tauxRetours_val:str = ""



    @classmethod   
    def depuis_codesIRIS(cls, formation:Formation_protocol, codes_IRIS:Iterable[int], annee:Optional[int]=None, periode:str="Année") -> BilanSession:
        """
        Génère un bilan de session à partir d'un ou plusieurs codes IRIS (un bilan pour une session ou pour plusieurs sessions (période))

        :param code_IRIS: Code IRIS de la session pour laquelle on souhaite faire le bilan
        :type code_IRIS: int
        :param formation: Objet formation associé à ce bilan
        :type formation: Formation_protocol
        :param annee: Année du bilan (s'il n'est pas donné on l'obtiendra d'IRIS sessions)
        :type annee: Optional[int], optional
        :return: Un objet bilan de session de ce code IRIS
        :rtype: BilanSession
        """
        # On initialise l'instance 
        instance = cls(formation=formation, annee=-1)
        instance.codes_IRIS = codes_IRIS
        instance._periode = periode

        # Si l'année n'est pas donnée, je la récupère d'IRIS sessions
        instance._annee = annee if annee is not None else int(instance.df_sessions_filtre_codesIRIS["Année début ses."].iloc[0])






        # === ON FAIT LES VERIFICATIONS QUI ANNULERAIENT LE TRAITEMENT ===
        continuer = instance.verifier_traitement_bilan()
        if not continuer:
            return
        


        # === TRAITEMENT DU BILAN DE SESSION ===
        # On affecte les autres données de base à la main
        
        #instance._iris_sessions = get_iris(typeExport="Sessions")

        # On met à jour l'Excel evalstat de la formation si la session demandée par l'utilisateur ne s'y trouve pas
        instance._maj_evalstat()

        # On calcule les stats
        instance._calculer_stats_criteres()
        #pprint(instance._stats_stagiaires)

        # On construit le bilan de session (bilan de session V3)
        instance._bilanSessionV3()

        # On ouvre le word
        FichierWord.depuisFichier(chemin_fichier=instance.chemin_word_bilan_output, charger_contentControl=False, afficherWord=True)

        # On envoie un mail au chef d'unité pour la signature du pdf
        instance._envoyer_mail_chef_unite()



        return instance




    # TODO : j'ai fait en sorte que les méthodes soient pas en warning, d'abord finaliser depuis_codesIRIS puis adapter pour que par période marche
    # TODO : A faire ; refactoriser ce qui est mutualisable
    @classmethod   
    def depuis_periode(cls, formation:Formation_protocol, annee:int, periode:str) -> None:
        """
        Permet de générer un bilan de session selon une période qui est l'un de ces éléments : ["1er semestre", "2nd semestre", "Année"]
        Ex : BilanSession.bilanUnique_parPeriode("948", 2024, "Année")
        """

        # On initialise l'instance 
        instance = cls(formation=formation, annee=annee, periode=periode)
        

        # On charge IRIS sessions et on vérifie qu'on a bien une ligne sinon on sort
        # TODO : besoin ?
        
        # On teste la pré-existance du bilan Word
        continuer = verifier_existance_fichier(instance.chemin_word_bilan_output)
        if not continuer:
            print("❌  Bilan de session déjà existant → Arrêt du traitement du bilan par l'utilisateur")
            return

        # Ouverture / création du dataframe de l'extract IRIS sessions filtré selon la période demandée
        #instance._charger_df_sessions_filtre_selon_periode()  
        instance._iris_sessions = get_iris(typeExport="Sessions")

        # On met à jour df_sessions_filtre selon les sessions que souhaite garder / exclure l'utilisateur
        df_filtre, codes_sessions_exclues_par_utilisateur = instance._iris_sessions.demande_sessions_a_exclure()
        #print("\nÉtat de Excel sessions filtré sur période et trigramme :")
        #pprint(instance._df_sessions_filtre)

        # S'il n'y a plus de session à lire dans _df_sessions_filtre, alors il n'y a plus de raison de faire le bilan
        if len(instance.df_sessions_filtre_periode) != 0 :
            # On met à jour l'Excel evalstat de la formation si des sessions demandées par l'utilisateur ne s'y trouvent pas
            instance._maj_evalstat()

            # On calcule les stats
            instance._calculer_stats_criteres()
            #pprint(instance._stats_stagiaires)

            # On construit le bilan de session (bilan de session V3)
            instance._bilanSessionV3()

            # On ouvre le word
            FichierWord.depuisFichier(chemin_fichier=instance.chemin_word_bilan_output, charger_contentControl=False, afficherWord=True)

            # On envoie un mail au chef d'unité pour la signature du pdf
            instance._envoyer_mail_chef_unite()


        else:
            vlog.print("Info", f"⚠️ Toutes les sessions sont exclues : il n'y a plus de raison de faire le bilan de session.")
        








    # =========================
    # === METHODES INTERNES === 
    # =========================
    def _get_codesIRIS_par_statut_evalStat(self, statut_evalStat: str) -> list[int]:
        """
        Retourne une liste des codes IRIS pour un statut donné.

        :param statut_evalStat: statut EvalStat (["Traité", "Traité - Code IRIS déjà dans l'évaluation de la formation", "Traité - CSV déjà dans l'évaluation de la formation", "Exclu - Aucun CSV fourni", "Exclu - Code IRIS pas dans Extract IRIS sessions", "Exclu - Problème lecture CSV", "Exclu - CSV vide / Aucun retour"])
        :type statut_evalStat: str
        :return: une liste des codes IRIS avec ce statut
        :rtype: list[int]
        """
        codes_IRIS = []
        for code_IRIS in self._codes_IRIS:
            session = self._formation.get_session_par_codeIRIS(code_IRIS)
            if session.eval.statut == statut_evalStat:
                codes_IRIS.append(code_IRIS)
        return codes_IRIS

    def _get_codesIRIS_par_statut_pourBilan(self, statut_pourBilan: str) -> list[int]:
        """
        Retourne une liste des codes IRIS pour une clé de mapping donnée (statut pour bilan, i.e. mieux nommés).

        :param statut_pourBilan: Statut "pour bilan" (i.e. mieux nommés) ["Exploités pour les évaluations (CSV présents)", "Exploités pour les évaluations (CSV présents)", "Exclus des évaluations (problème traitement CSV)", "Exclus des évaluations (CSV manquants)", "Exclus des évaluations (CSV vide / aucun retour)", "Exclus entièrement du bilan (non présent dans IRIS / mauvais code)"] 
        :type statut_pourBilan: str
        :return: une liste des codes IRIS avec ce statut
        :rtype: list[int]
        """
        #pprint(self._mapping_statuts)

        statut_evalStat = next((clef for clef, valeur in self._mapping_statuts.items() if valeur == statut_pourBilan), None)
        if statut_evalStat is None:
            return []
        return self._get_codesIRIS_par_statut_evalStat(statut_evalStat)

    def _maj_evalstat(self) -> None:
        """
        Crée ou ouvre les EvalStat de la/les sessions demandées et met à jour le fichier EvalStat de la formation
        Ne s'applique que si des sessions demandées par l'utilisateur ne s'y trouvent pas.
        (on regarde les CSV qui ne sont pas dans le fichier Excel global à partir de la liste df_sessions_filtre['Code IRIS'])
        """

        # On ouvre ou on traite les EvalStats non déjà créés
        EvalStat_services.ouvrir_ou_traiter_evalStat_depuis_liste_codes_IRIS(codes_IRIS=self._codes_IRIS, formation=self._formation)

    def _calculer_stats_criteres(self) -> None: # dict[str, dict[str, int|float|str|None]]:
        """
        Calcule les statistiques (nombre de retours, moyenne retours, agrégation des commentaires) de tous les critères.

        Fait ce traitement pour tous les éléments dont nous avons des CSV (i.e. appartenant à liste_codesIRIS_pour_statsCSV)
        
        Retourne un dictionnaire de la forme :
        {
            "Nom du critère": {
                "Nombre": ...,
                "Moyenne": ...,
                "Commentaires": ...
            },
            ...
        }

        :return: Un dictionnaire de tous les critères
        :rtype: dict
        """

        # S'il n'y a pas de CSV disponibles pour les stats, alors ce n'est pas la peine de faire les stats
        if len(self.liste_codesIRIS_pour_statsCSV) > 0 :
            # On récupère la iste des critères
            liste_criteres = self.df_stagiaires_filtre_statsCSV['Critère'].dropna().unique()


            # On fait les stats pour chaque critère
            for critere in liste_criteres:
                # Dataframe filtré sur ce critère
                df_filtre = self.df_stagiaires_filtre_statsCSV[self.df_stagiaires_filtre_statsCSV['Critère'] == critere]
                
                # Nombre d'éléments avec ce critère
                nb = len(df_filtre)
                # Moyenne de ce critère
                moyenne = float(df_filtre['Note'].mean()) if nb > 0 else None

                # On concatère les commentaires associés
                commentaires_concat = "\n".join(
                    "• " + c.strip()
                    for c in df_filtre['Commentaires'].dropna().astype(str)
                    if c.strip() != ""
                )

                # On met toutes ces données en forme dans _stats_stagiaires
                self._stats_stagiaires[critere] = {
                    "Nombre": nb,
                    "Moyenne": moyenne,
                    "Commentaires": commentaires_concat
                }
        #else:
            #vlog.print("Info", f"⚠️  Aucun CSV disponible pour ce bilan : les statistiques des stagiaires ne seront pas évaluées.")
            #self._commentairesBilan += f"\n⚠️  Aucun CSV disponible pour ce bilan : les statistiques des stagiaires ne seront pas évaluées.\n"

        #return self._stats_stagiaires

    def _envoyer_mail_chef_unite(self, pj:Optional[list[str]] = None):
        """
        Envoie un mail au chef d'unité avec en lien le PDF à signer
        """       
        


        chemin_pdf_bilan_output = self.chemin_word_bilan_output.with_suffix(".pdf")
        #self._CORPS_MAIL_CHEF_UNITE.replace()
        corps_html = remplacer_champs(config.CORPS_MAIL_CHEF_UNITE, [
            ["lien_pdf_bilan", chemin_pdf_bilan_output],
            ["formation", f"{self._titreFormation} ({self.trigramme_formation})"],
            ["periode", minuscule_premiere_lettre(self._periodeSessionsEvaluees)],
        ])

        Mail.creer_mail(
            destinataires=config.ADRESSE_MAIL_CHEF_UNITE,
            sujet=f"Signature bilan de session {self._titreFormation} ({self.trigramme_formation}) : {chemin_pdf_bilan_output.name}",
            corps_html=corps_html,
            pieces_jointes=pj,
            envoyer_mail=False  # envoie directement sans afficher
        ) 


    # ===============================================
    # === MÉTHODES POUR LA V3 DU BILAN DE SESSION === 
    # ===============================================
    def _bilanSessionV3(self) -> None:
        # On construit champs de fusion
        self.__construit_champsFusionV3()

        # On merge les champs de fusion
        self.__mergeBilanV3()

    def __construit_champsFusionV3(self) -> None:
        """
        Calcule puis définit les str des champs de fusion
        """
        #####
        # Affectation des valeurs pour les champs de fusion (ce sont des str)
        #####

        # ===================================
        # === Calculs données transverses ===
        # ===================================
        # Nombre de stagiaires qui ont formulé des retours (provient de eval formation)
        nb_stagiaires_retours = self.df_stagiaires_filtre_statsCSV['NOM Prénom'].nunique()
        # Nombre d'apprenants sur les sessions dont on peut faire les stats CSV (peut provenir de IRIS session ou de eval formation, on prend de df_stagiaire)
        nb_apprenants = int(self.df_stagiaires_filtre_statsCSV_1ligne_par_session['Nb présents'].sum())


        # ======================
        # === Partie en-tête ===
        # ======================
        # On met ici toutes les sessions de la période non excclues par l'utilisateur et qui est dans IRIS sessions

        # Les éléments entre parenthèses (employés lors des stats CSV) sont gérés dans la partie stats
        self._titreFormation = self.intitule_formation
        # self._codeFormation = codeFormation  # (donné en argument)
        #self._periodeSessionsEvaluees = self.periodeSessionsEvaluees  # Evalué plus haut
        self._nbSessionsEvaluees = f"{len(self.df_sessions_filtre_enTete)} session" + ("s" if len(self.df_sessions_filtre_enTete) > 1 else "")  # Valeur toutes les données  
        self._numerosSessions = "\n".join(self.df_sessions_filtre_enTete["N° Session"].dropna().astype(str).unique())
        self._nbApprenants = f"{self.df_sessions_filtre_enTete['Nb. Nommés'].sum()} apprenant" + ("s" if self.df_sessions_filtre_enTete['Nb. Nommés'].sum() > 1 else "")  # Valeur toutes les données
        #self._rp = ", ".join(self.df_sessions_filtre["Trigramme RP"].dropna().astype(str).unique())  # Valeur toutes les données
        self._rp = ", ".join(self.df_sessions_filtre_enTete["Nom responsable pédag."].dropna().astype(str).unique() + " " + self.df_sessions_filtre_enTete["Prénom responsable pédag."].dropna().astype(str).unique())  # Valeur toutes les données
        #self._af = ", ".join(self.df_sessions_filtre["Trigramme AF"].dropna().astype(str).unique())  # Valeur toutes les données
        self._af = ", ".join(self.df_sessions_filtre_enTete["Créée par"].dropna().astype(str).unique())  # Valeur toutes les données



        
        # ===========================
        # === Partie commentaires ===
        # ===========================
        # Cas avec aucun CSV dispo pour les stats
        if len(self.liste_codesIRIS_pour_statsCSV) == 0 :
            vlog.print("Info", f"⚠️  Aucun CSV disponible pour ce bilan : les statistiques des stagiaires ne seront pas évaluées.")
            self._commentairesBilan = f"\n⚠️  Aucun CSV disponible pour ce bilan : les statistiques des stagiaires ne seront pas évaluées.\n"
        
        # Cas avec certains CSV non dispo pour les stats mais pas tous
        elif len(self.liste_codesIRIS_pour_enTete) != len(self. liste_codesIRIS_pour_statsCSV) :

            self._commentairesBilan = "\n⚠️  Certaines sessions n'ont pas de CSV exploitables pour les statistiques (cf. liste ci-dessous)."
            self._commentairesBilan += "\nDonnées employées pour les statistiques :"
            self._commentairesBilan += "\n   • Sessions évaluées : ".join(self.df_stagiaires_filtre_statsCSV_1ligne_par_session["N° Session"])
            self._commentairesBilan += "\n   • Nombre d'apprenants sur ces sessions : " + f" ({nb_apprenants:.0f})"
            self._commentairesBilan += "\n   • Nombre de stagiaires ayant formulé des retours : " + f" ({nb_stagiaires_retours:.0f})"

            #self._nbSessionsEvaluees += f" ({len(self.liste_codesIRIS_pour_statsCSV_uniquement)})"  # Valeur si on ne prend que les données CSV
            #self._numerosSessions += "\n".join("(" + self._df_stagiaires_final_1ligne_session["N° Session"].dropna().astype(str).unique() + ")")  # Valeur si on ne prend que les données CSV
            #self._nbApprenants += f" ({nb_apprenants:.0f})"  # Valeur si on ne prend que les données CSV
            #self._rp += " (" + ", ".join(self._df_stagiaires_final_1ligne_session["Trigramme RP"].dropna().astype(str).unique()) + ")"  # Valeur si on ne prend que les données CSV
            #self._af += " (" + ", ".join(self._df_stagiaires_final_1ligne_session["Trigramme AF"].dropna().astype(str).unique()) + ")"  # Valeur si on ne prend que les données CSV
        
        
        # Affichage sessions avec pb CSV
        if len(self.liste_codesIRIS_avec_pb_CSV) > 0:
            self._commentairesBilan += "\n\nListe des sessions dont les statistiques n'ont pas pu être évaluées :"
            l_codes_IRIS = self.liste_codesIRIS_avec_pb_CSV

            for statut_evalStat, statut_pourBilan in self._mapping_statuts.items():
                if statut_evalStat not in ["Traité", "Exclu - Code IRIS pas dans Extract IRIS sessions"]:
                    l_sessions = (
                        self.df_stagiaires[self.df_stagiaires['Code IRIS'].isin(l_codes_IRIS)]
                        .dropna()
                        .drop_duplicates(subset=['Code IRIS'])
                        ["N° Session"]
                        .tolist()
                        )
                    if len(l_sessions) > 0:
                        self._commentairesBilan += f"\n\n   • {statut_pourBilan} :" + "".join(f"\n       - {isession}" for isession in l_sessions)

        # Affichage sessions exclues
        if len(self.liste_codesIRIS_exclus_totalement) > 0:
            self._commentairesBilan += "\n\nListe des sessions de la période entièrement exclues du bilan :"
            l_codes_IRIS = self.liste_codesIRIS_exclus_totalement
            l_sessions = (
                self.df_stagiaires[self.df_stagiaires['Code IRIS'].isin(l_codes_IRIS)]
                .dropna()
                .drop_duplicates(subset=['Code IRIS'])
                ["N° Session"]
                .tolist()
                )
            if len(l_sessions) > 0:
                self._commentairesBilan += "".join(f"\n       - {isession}" for isession in l_sessions)




        """
        self._commentairesBilan += "\n\nListe des sessions exclues :"
        for statut_evalStat, statut_pourBilan in self._mapping_statuts.items():
            #if statut_evalStat[:5] == "Exclu":
            # TODO : le bilan sort mais j'ai des pb
            l_codes_IRIS = self._get_codesIRIS_par_statut_evalStat(statut_evalStat=statut_evalStat)
            l_sessions = (
                self.df_stagiaires[self.df_stagiaires['Code IRIS'].isin(l_codes_IRIS)]
                .dropna()
                .drop_duplicates(subset=['Code IRIS'])
                ["N° Session"]
                .tolist()
                )
            self._commentairesBilan += f"\n\n   • {statut_pourBilan} :" + "".join(f"\n       - {isession}" for isession in l_sessions)
        """
        vlog.print("Info", f"\n{self._commentairesBilan}")     



        """
        #pprint(self._exploitationBilan)
        self._commentairesBilan += "\n\nListe des sessions avec problèmes:"
        for critere, lsessions in self._exploitationBilan.items():
            if lsessions:
                self._commentairesBilan += f"\n\n   • {critere} :" + "".join(f"\n       - {isession}" for isession in lsessions)
        vlog.print("Info", f"\n{self._commentairesBilan}")      
        """



        # ========================
        # === Partie stats CSV ===
        # ========================
        # On n'affecte les champs suivants que si des CSV sont disponibles pour les stats
        if len(self.liste_codesIRIS_pour_statsCSV) > 0 :

            


            stats_sous_3 = { # Dictionnaire pour les critères dont la moyenne est inférieure à 3 et non exclus (critères dans la liste self._CRITERES_A_ENLEVER)
                critere: valeurs
                for critere, valeurs in self._stats_stagiaires.items()
                if (
                    critere not in self._CRITERES_A_ENLEVER
                    and valeurs["Moyenne"] is not None
                    and valeurs["Moyenne"] < 3
                )
            }


            
            # Si des champs ne sont pas dans le CSV, alors on garde "" qui est déjà définit dans le constructeur
            # Satisfaction globale
            try:
                self._satisfactionGlobale_moy = f'{self._stats_stagiaires["Satisfaction globale"]["Moyenne"]:.1f}/5'
            except:
                self._satisfactionGlobale_moy = "Pas de donnée"      
            try:
                self._satisfactionGlobale_com = self._stats_stagiaires["Satisfaction globale"]["Commentaires"].replace("_x000D_", "\n")
            except:
                pass
            
            # Recommanderiez-vous + commentaires remarques suggestions
            try:
                self._recommandation_moy = f'{self._stats_stagiaires["Recommanderiez-vous cette formation ?"]["Moyenne"]/5*100:.0f}%'  # (on divise par 5 car on a un booléen stcké sous forme de note sur 5 : 0 = False, 5 = True)
            except:
                self._recommandation_moy = "Pas de donnée"        
            try:
                self._commentairesRemarquesSuggestions_com = self._stats_stagiaires["Commentaires, remarques, suggestions"]["Commentaires"].replace("_x000D_", "\n")
            except:
                pass
            
            # Notes inférieures à 3
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
            
            # Taux de retour
            try:
                self._tauxRetours_val = f"{(nb_stagiaires_retours/nb_apprenants)*100:.0f}%"
            except:
                self._tauxRetours_val = "Pas de donnée"

    def __mergeBilanV3(self) -> None:
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

        #print(f"BilanFormation lancé avec : trigramme={self.trigramme_formation}, année={self._annee}")
        document = MailMerge(config.CHEMIN_MODELE_WORD_BILAN_SESSION)
        #print(document.get_merge_fields())

        document.merge(
            titreFormation = self._titreFormation,
            codeFormation = self.trigramme_formation,
            periodeSessionsEvaluees = self._periodeSessionsEvaluees,
            nbSessionsEvaluees = self._nbSessionsEvaluees,
            numerosSessions = self._numerosSessions,
            nbApprenants = self._nbApprenants,
            rp = self._rp,
            af = self._af,
            commentairesBilan = self._commentairesBilan,
            satisfactionGlobale_moy = self._satisfactionGlobale_moy,
            satisfactionGlobale_com = self._satisfactionGlobale_com,
            recommandation_moy = self._recommandation_moy,
            commentairesRemarquesSuggestions_com = self._commentairesRemarquesSuggestions_com,
            evalInf3_val = self._evalInf3_val,
            evalInf3_com = self._evalInf3_com,
            tauxRetours_val = self._tauxRetours_val
            )
        
        # On crée le répertoire pour les bilans de session de cette année s'il n'existe pas
        self.chemin_word_bilan_output.parent.mkdir(parents=True, exist_ok=True)

        # On écrit le fichier
        document.write(self.chemin_word_bilan_output)


    # =========================
    # === METHODES EXTERNES ===
    # =========================

    def verifier_traitement_bilan(self) -> bool:
        """
        Vérifie si l'on doit traiter un bilan de sessions à partir :
            - le bilan Word n'existe pas déjà ;
            - des codes IRIS (bien à 5 caractères + bien présent dans IRIS sessions).

        Pour l'instant on coupe le programme en cas d'échec d'une vérification (la sortie bool=False ne sert à rien car non employée)

        :return: Un bool pour dire si le traitement doit continuer :
            - True on peut traiter l'EvalStat (aucun souci détecté) ; 
            - False il ne faut pas traiter l'EvalStat (code IRIS ou CSV non détectés dans EvalStat + code IRIS présent dans IRIS sessions).
        :rtype: bool
        """
        # Vérification 1 : on vérifie la pré-existance du bilan Word ; si il existe déjà, alors on arrête le traitement
        continuer = verifier_existance_fichier(self.chemin_word_bilan_output)
        if not continuer:
            vlog.log_erreur("❌  Bilan de session déjà existant → Arrêt du traitement du bilan par l'utilisateur")
            
        
        
        # Vérifications 2 : liés à code_IRIS
        for code_IRIS in self._codes_IRIS:
            if code_IRIS is not None:
                # 2.1 : On vérifie que code_iris est bien un entier à 5 chiffres
                est_code_IRIS_valide, _ = IRIS.verifier_code_IRIS(code_IRIS, int)
                if not est_code_IRIS_valide:
                    vlog.log_erreur(f"❌  Code IRIS renseigné non valide : {code_IRIS}")
                    #statut = "Exclu - Code IRIS renseigné non valide"
                    #return False, statut


                # 2.2 : On vérifie si le code_IRIS est bien existant dans l'extract IRIS sessions
                if code_IRIS not in self.df_sessions_filtre_periode["Code IRIS"].values:
                    vlog.log_erreur(f"❌  Code IRIS {code_IRIS} non trouvé dans l’extract IRIS.")
                    #print(f"⚠️  Code IRIS {code_IRIS} non trouvé dans l’extract IRIS.")
                    #statut = "Exclu - Code IRIS pas dans Extract IRIS sessions"
                    #return False, statut
 

        # Si rien n'a arrêté les vérifications, alors tout est OK
        return True



    # =========================
    # === GETTERS / SETTERS === 
    # =========================
    # Getters / setters liés à BilanSession
    @property
    def chemin_word_bilan_output(self) -> Path:
        """
        Renvoie le chemin de sortie du bilan de formation.
        Cette donnée est stockée dans le fichier de config (valeur par défaut).
        
        :return: Chemin de sortie du bilan de formation.
        :rtype: Path
        """
        return config.format_path(config.CHEMIN_WORD_BILAN_SESSION_OUTPUT, trigramme_formation=self.trigramme_formation, annee=self._annee, periode=self.periodeSessionsEvaluees, unite=config.UNITE)

    @property
    def periodeSessionsEvaluees(self) -> str:
        """
        Période de la session dans le cadre d'une session unique (len(codes_IRIS)=1).

        Construction :
            - si un seul code IRIS : f"Session {self.numero_session} uniquement ({self.mois_session} {self._annee})"
            - si plusieurs codes IRIS : f"{self._periode} {self._annee}"
        """
        if self._periodeSessionsEvaluees == "":
            if len(self._codes_IRIS) == 1:  # Cas code IRIS unique
                self._periodeSessionsEvaluees = f"Session {self.numero_session} uniquement ({self.mois_session} {self._annee})"
            elif len(self._codes_IRIS) > 1:
                self._periodeSessionsEvaluees = f"{self._periode} {self._annee}"
            else:
                vlog.log_erreur("J'appelle periodeSessionsEvaluees alors que len(self.codes_IRIS)<=0")
        
        return self._periodeSessionsEvaluees

    @property
    def codes_IRIS(self) -> Iterable[int]:
        return self._codes_IRIS
    
    @codes_IRIS.setter
    def codes_IRIS(self, valeur:int|Iterable[int]) -> None:
        if isinstance(valeur, int):
            self._codes_IRIS.append(valeur)
        elif isinstance(valeur, Iterable) and not isinstance(valeur, str):  # Iterable[int] n’est pas valide dans isinstance
            self._codes_IRIS = valeur
        else:
            vlog.log_erreur("La valeur n'est ni un int ni un Iterable de int (codes_IRIS.setter)")

    @property
    def code_IRIS(self) -> int:
        """
        Le 1er code IRIS (code_IRIS au singulier donc on considère qu'on est sur un bilan contenant un seul code IRIS unique)
        """
        return self._codes_IRIS[0]

    """
    # Ca va me créer de la confusion
    @code_IRIS.setter
    def code_IRIS(self, valeur:int):
        "
        Renseigne un code IRIS (code_IRIS au singulier donc on considère qu'on est sur un bilan contenant un seul code IRIS unique)
        "
        if not self._codes_IRIS: # Cas d'une liste vide
            self._codes_IRIS.append(valeur)
        else:
            self._codes_IRIS[0] = valeur
    """

    @property
    def liste_codesIRIS_exclus_totalement(self) -> list[int]:
        """
        Retourne la liste des codes IRIS qui seront complètement exclus du bilan de sessions (en-tête et CSV).

        Ca correspond aux codes IRIS qui ont été exclus par l'utilisateur ou qui ne sont pas dans IRIS sessions.

        :return: La liste des codes IRIS qui seront complètement exclus du bilan de sessions (en-tête et CSV).
        :rtype: list[int]
        """
        return [
            code_IRIS for code_IRIS in self._codes_IRIS
            if self._formation.get_session_par_codeIRIS(code_IRIS).eval.statut in self._statut_exclus_entierement
        ]

    @property
    def liste_codesIRIS_pour_enTete(self) -> list[int]:
        """
        Retourne la liste des codes IRIS qui seront dans l'en-tête du bilan de sessions.

        Ca correspond aux codes IRIS qui n'ont pas été exclus par l'utilisateur et qui sont dans IRIS sessions.

        :return: La liste des codes IRIS qui seront dans l'en-tête du bilan de sessions.
        :rtype: list[int]
        """
        return [
            code_IRIS for code_IRIS in self._codes_IRIS
            if code_IRIS not in self.liste_codesIRIS_exclus_totalement
        ]
    
    @property
    def liste_codesIRIS_pour_statsCSV(self) -> list[int]:
        """
        Retourne la liste des codes IRIS pour lesquels on peut calculer les stats à partir des CSV.

        Ca correspond aux codes IRIS pour lesquels le CSV est fonctionnel et non vide.

        :return: La liste des codes IRIS pour lesquels on peut calculer les stats à partir des CSV.
        :rtype: list[int]
        """
        return self._get_codesIRIS_par_statut_pourBilan("Exploités pour les évaluations (CSV présents et non vides)")

    @property
    def liste_codesIRIS_avec_pb_CSV(self) -> list[int]:
        """
        Retourne la liste des codes IRIS pour lesquels il y a un problème de CSV.

        Ca correspond aux codes IRIS de l'en-tête "moins" les codes IRIS pour lesquels on a des CSV fonctionnels.

        :return: La liste des codes IRIS pour lesquels il y a un problème de CSV.
        :rtype: list[int]
        """
        return [
            code_IRIS for code_IRIS in self.liste_codesIRIS_pour_enTete
            if code_IRIS not in self.liste_codesIRIS_pour_statsCSV
        ]



    # Liens avec Formation
    @property
    def trigramme_formation(self) -> str:
        return self._formation.trigramme_formation
    
    @property
    def eval_formation(self) -> Optional[EvalStat_formation]:
        return self._formation.eval

    @property
    def eval_fe(self) -> Optional[FichierExcel]:
        return self._formation.eval.fe
     
    @property
    def df_stagiaires(self) -> Optional[DataFrame]:
        return self.eval_formation.fe.get_df_tableau("Stagiaires")
    
    @property
    def df_stagiaires_filtre_statsCSV(self) -> DataFrame:
        """
        df_stagiaires (eval formation) filtré sur les codes IRIS pour lesquels on peut calculer les stats à partir des CSV.

        Ca correspond aux codes IRIS pour lesquels le CSV est fonctionnel et non vide.

        :return: df_stagiaires (eval formation) filtré sur les codes IRIS pour lesquels on peut calculer les stats à partir des CSV.
        :rtype: DataFrame
        """
        return self.df_stagiaires[self.df_stagiaires['Code IRIS'].isin(self.liste_codesIRIS_pour_statsCSV)]
    
    @property
    def df_stagiaires_filtre_statsCSV_1ligne_par_session(self) -> DataFrame:
        return self.df_stagiaires_filtre_statsCSV.drop_duplicates(subset=['N° Session'])



    # Liens avec IRIS 
    @property
    def iris_sessions(self) -> IRIS_traite:
        return get_iris(typeExport="Sessions")
  
  
    @property
    def intitule_formation(self) -> str:
        """
        Intitulé de la formation (prend le nom de la dernière ligne pour avoir la dernière mise à jour)
        """
        return self.df_sessions_filtre_periode["Session"].iloc[-1]

    @property
    def mois_session(self) -> str:
        """
        Mois de la session dans le cadre d'une session unique (len(codes_IRIS)=1)
        """
        return mois_fr_depuis_date(self.df_sessions_filtre_codesIRIS["Date début ses."].iloc[0])

    @property
    def numero_session(self) -> str:
        """
        Numéro de la session dans le cadre d'une session unique (len(codes_IRIS)=1)
        """
        return self.df_sessions_filtre_codesIRIS["N° Session"].iloc[0]


    @property
    def df_sessions_filtre_periode(self) -> DataFrame:
        """
        Renvoie le dataframe de l'extract IRIS filtré VTE selon plusieurs critères :
            - Statut Session != "Annulée" (toujours) ;
            - Nb. Nommés != 0 (toujours) ;
            - l'année de la session (année n du bilan de formation) ;
            - le trigramme de la formation traité par le bilan (self.trigramme_formation).
        
        :return: extract IRIS VTE filtré
        :rtype: DataFrame
        """
        return self.iris_sessions.df_filtre_periode(trigramme_formation=self.trigramme_formation, annee=self._annee, periode=self._periode)

    @property
    def df_sessions_filtre_codesIRIS(self) -> DataFrame:
        """
        Renvoie le dataframe de l'extract IRIS filtré VTE selon plusieurs critères :
            - Code IRIS in self._codes_IRIS
            - Statut Session != "Annulée" (toujours) ;
            - Nb. Nommés != 0 (toujours) ;
            - l'année de la session (année n du bilan de formation) ;
            - le trigramme de la formation traité par le bilan (self.trigramme_formation).
        
        :return: extract IRIS VTE filtré sur codes_IRIS
        :rtype: DataFrame
        """     
        return self.df_sessions_filtre_periode[self.df_sessions_filtre_periode['Code IRIS'].isin(self._codes_IRIS)]

    @property
    def df_sessions_filtre_exclus_totalement(self) -> DataFrame: #_df_stagiaires_final → df_sessions_filtre_stats_generales
        """
        Renvoie le dataframe de l'extract IRIS filtré sur les codes IRIS qui seront complètement exclus du bilan de sessions (en-tête et CSV).

        Ca correspond aux codes IRIS qui ont été exclus par l'utilisateur ou qui ne sont pas dans IRIS sessions.

        :return: dataframe de l'extract IRIS filtré sur les codes IRIS qui seront complètement exclus du bilan de sessions (en-tête et CSV).
        :rtype: DataFrame
        """
        return self.df_sessions_filtre_codesIRIS[self.df_sessions_filtre_codesIRIS['Code IRIS'].isin(self.liste_codesIRIS_pour_enTete)]

    @property
    def df_sessions_filtre_enTete(self) -> DataFrame: #_df_stagiaires_final → df_sessions_filtre_stats_generales
        """
        Renvoie le dataframe de l'extract IRIS filtré sur les codes IRIS exploités pour le bilan (en-tête)

        Ca correspond aux codes IRIS qui n'ont pas été exclus par l'utilisateur et qui sont dans IRIS sessions.

        :return: dataframe de l'extract IRIS filtré sur les codes IRIS exploités pour le bilan (en-tête)
        :rtype: DataFrame
        """
        return self.df_sessions_filtre_codesIRIS[self.df_sessions_filtre_codesIRIS['Code IRIS'].isin(self.liste_codesIRIS_pour_enTete)]

    @property
    def df_sessions_filtre_statsCSV(self) -> DataFrame:
        """
        dataframe de l'extract IRIS filtré sur les codes IRIS pour lesquels on peut calculer les stats à partir des CSV.

        Ca correspond aux codes IRIS pour lesquels le CSV est fonctionnel et non vide.

        :return: dataframe de l'extract IRIS filtré sur les codes IRIS pour lesquels on peut calculer les stats à partir des CSV.
        :rtype: DataFrame
        """
        return self.df_sessions_filtre_codesIRIS[self.df_sessions_filtre_codesIRIS['Code IRIS'].isin(self.liste_codesIRIS_pour_statsCSV)]

    @property
    def df_sessions_filtre_avec_pb_CSV(self) -> DataFrame:
        """
        dataframe de l'extract IRIS filtré sur les codes IRIS pour lesquels il y a un problème de CSV.

        Ca correspond aux codes IRIS de l'en-tête "moins" les codes IRIS pour lesquels on a des CSV fonctionnels.

        :return: dataframe de l'extract IRIS filtré sur les codes IRIS pour lesquels il y a un problème de CSV.
        :rtype: DataFrame
        """
        return self.df_sessions_filtre_codesIRIS[self.df_sessions_filtre_codesIRIS['Code IRIS'].isin(self.liste_codesIRIS_avec_pb_CSV)]

    
class BilanSession