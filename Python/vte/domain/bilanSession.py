from __future__ import annotations
from pathlib import Path
from typing import List, Optional, Protocol

from pandas import DataFrame
from mailmerge import MailMerge

from vte.core.iris_referentiel import get_iris
from vte.domain.evalStat import EvalStat, EvalStat_formation
from vte.services.evalStat_services import EvalStat_services
from vte.domain.iris import IRIS, IRIS_traite
from vte.core import config
from vte.utils.office import FichierExcel, FichierWord, Mail
from vte.utils.utils import *

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
    def trigramme_formation(self) -> str: ...

    @property
    def eval(self) -> EvalStat_formation|None: ...


# ======================================================================================
# CLASSE BilanSession
# ======================================================================================

class BilanSession:
    """
    C'est la classe qui contient tous les éléments de ma formation pour mon bilan de session V3
    """
    # === VARIABLES PARTAGÉES ENTRE TOUTES LES INSTANCES
    _iris_sessions:Optional[IRIS_traite] = None  # Fichier Excel qui contient les extracts IRIS Sessions (ou a minima celles de la période en cours) → Plusieurs bilan peuvent être fait à partir de cet extract, c'est donc une variable de classe


    _CRITERES_A_ENLEVER:list[str] = [  # Critères à ne pas retenir pour le calcul des moyennes < 3
        "Comment avez-vous connu cette formation ?", "Avez-vous d'autres besoins de formation ?", "Commentaires, remarques, suggestions", "Recommanderiez-vous cette formation ?"]

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
        self._codes_IRIS:List[int] = []

        # --- Variables de traitement ---
        self._exploitationBilan:dict[list] = {
            "Exploités pour les stats générales" : [],  # Exploités pour stats initiales → Dans _demande_sessions_a_exclure
            "Exploités pour les évaluations (CSV présents)" : [],  # Exploités pour les stats stagiaires → Dans _maj_evalstat_formation
            "Exclus des évaluations (CSV manquants)" : [],   # Exclus des évaluations car CSV stagiaires manquants → Dans _maj_evalstat_formation
            "Exclus des évaluations (problème traitement CSV)" : [],   # Exclus des évaluations car problème au traitement des CSV → Dans _maj_evalstat_formation
            "Exclus des évaluations (CSV vide / aucun retour)" : [],   # Exclus des évaluations car le CSV est vide (i.e. aucun retour d'utilisateur)
            "Exclus entièrement du bilan (exclus par utilisateur)" : [],  # Exclus entièrement du bilan car sessions non réalisées ou mauvais RP (exclus par l'utilisateur) → Dans _demande_sessions_a_exclure
            "Exclus entièrement du bilan (non présent dans IRIS / mauvais code)" : [],  # Exclus entièrement du bilan car sessions non réalisées ou mauvais RP (exclus par l'utilisateur) → Dans _maj_evalstat_formation
        }

        self._stats_stagiaires: Optional[dict] = None  # Dictionnaire des stats des CSV


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

    # TODO : adapter pour pouvoir prendre plusieurs codes IRIS avec un itérable de int.
    """
    Cette classe sera celle qui sera lancée à partir de "par période" ou "par code IRIS unique" 
    qui auront traité/généré toutes les données d'entrée requises :
    (foramtion, codes_iris, période, année...) → Peut-être mettre dans __init__ ?
    """
    @classmethod   
    def depuis_codesIRIS(cls, formation:Formation_protocol, codes_IRIS:Iterable[int], chemin_IRIS_sessions:Optional[Path]=None) -> BilanSession:
        """
        Génère un bilan de session à partir d'un code IRIS (un bilan pour une session)

        :param formation: Objet formation associé à ce bilan
        :type formation: Formation_protocol
        :param codes_IRIS: Codes IRIS de/des session(s) pour laquelle/lesquells on souhaite faire un bilan
        :type codes_IRIS: Iterable[int]
        :param chemin_IRIS_sessions: chemin IRIS sessions à employer si l'utilisateur en souhaite un différent que celui par défaut
        :type chemin_IRIS_sessions: Optional[Path], optional
        :return: Un objet bilan de session de ce code IRIS
        :rtype: BilanSession
        """
        pass


    @classmethod   
    def depuis_codeIRIS(cls, formation:Formation_protocol, code_IRIS:int, annee:Optional[int]=None, chemin_IRIS_sessions:Optional[Path]=None) -> BilanSession:
        """
        Génère un bilan de session à partir d'un code IRIS (un bilan pour une session)

        :param code_IRIS: Code IRIS de la session pour laquelle on souhaite faire le bilan
        :type code_IRIS: int
        :param formation: Objet formation associé à ce bilan
        :type formation: Formation_protocol
        :param annee: Année du bilan (s'il n'est pas donné on l'obtiendra d'IRIS sessions)
        :type annee: Optional[int], optional
        :param chemin_IRIS_sessions: chemin IRIS sessions à employer si l'utilisateur en souhaite un différent que celui par défaut
        :type chemin_IRIS_sessions: Optional[Path], optional
        :return: Un objet bilan de session de ce code IRIS
        :rtype: BilanSession
        """
        # === ON VERIFIE QU'IRIS SESSIONS CONTIENT BIEN LE CODE IRIS ===
        # On charge IRIS sessions et on vérifie qu'on a bien une ligne sinon on sort
        df_sessions_filtre_codeIRIS = BilanSession.charge_IRIS_session_et_verifie_codeIRIS(code_IRIS=code_IRIS, chemin_IRIS_sessions=chemin_IRIS_sessions)
        if annee is None:
            annee = int(df_sessions_filtre_codeIRIS["Année début ses."].iloc[0])

        # On initialise l'instance 
        instance = cls(formation=formation, annee=annee)


        # === ON FAIT LES AUTRES VERIFICATIONS QUI ANNULERAIENT LE TRAITEMENT ===
        # On vérifie que code_iris est bien un entier à 5 chiffres
        est_code_IRIS_valide, instance.code_IRIS = IRIS.verifier_code_IRIS(code_IRIS, int)
        if not est_code_IRIS_valide:
            vlog.log_erreur(f"Code IRIS en entrée non valide : {instance.code_IRIS} non traité", continuer=True)
            return
    
        # On vérifie la pré-existance du bilan Word ; si il existe déjà, alors on arrête le traitement
        bilan_preexistant = verifier_existance_fichier(instance.chemin_word_bilan_output)
        if bilan_preexistant:
            print("❌  Bilan de session déjà existant → Arrêt du traitement du bilan par l'utilisateur")
            return
        

    

        # === TRAITEMENT DU BILAN DE SESSION ===
        # On affecte les autres données de base à la main
        # instance.code_IRIS = code_IRIS  # Déjà affecté plus haut
        #instance._periode = instance.periode_session  # Déjà mis par défaut à "Année"    
        instance._iris_sessions = get_iris(typeExport="Sessions", chemin=chemin_IRIS_sessions)

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


    # TODO : j'ai fait en sorte que les méthodes soient pas en warning, à tester bilan_unique puis adapter pour que par période marche
    # TODO : A faire ; refactoriser ce qui est mutualisable
    @classmethod   
    def depuis_periode(cls, formation:Formation_protocol, annee:int, periode:str, chemin_IRIS_sessions:Optional[Path]=None) -> None:
        """
        Permet de générer un bilan de session selon une période qui est l'un de ces éléments : ["1er semestre", "2nd semestre", "Année"]
        Ex : BilanSession.bilanUnique_parPeriode("948", 2024, "Année")
        """

        # On initialise l'instance 
        instance = cls(formation=formation, annee=annee, periode=periode)
        

        # On charge IRIS sessions et on vérifie qu'on a bien une ligne sinon on sort
        # TODO : besoin ?
        
        # On teste la pré-existance du bilan Word
        bilan_preexistant = verifier_existance_fichier(instance.chemin_word_bilan_output)
        if bilan_preexistant:
            print("❌  Bilan de session déjà existant → Arrêt du traitement du bilan par l'utilisateur")
            return

        # Ouverture / création du dataframe de l'extract IRIS sessions filtré selon la période demandée
        #instance._charger_df_sessions_filtre_selon_periode()  
        instance._iris_sessions = get_iris(typeExport="Sessions", chemin=chemin_IRIS_sessions)

        # On met à jour df_sessions_filtre selon les sessions que souhaite garder / exclure l'utilisateur
        df_filtre, codes_sessions_exclues_par_utilisateur = instance._iris_sessions.demande_sessions_a_exclure()
        #print("\nÉtat de Excel sessions filtré sur période et trigramme :")
        #pprint(instance._df_sessions_filtre)

        # S'il n'y a plus de session à lire dans _df_sessions_filtre, alors il n'y a plus de raison de faire le bilan
        if len(instance.df_sessions_filtre) != 0 :
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
    def _maj_evalstat(self, chemin_IRIS_sessions:Optional[Path]=None) -> None:
        """
        Crée les EvalStat de la/les sessions demandées et met à jour le fichier EvalStat de la formation
        Ne s'applique que si des sessions demandées par l'utilisateur ne s'y trouvent pas
        (on regarde les CSV qui ne sont pas dans le fichier Excel global à partir de la liste df_sessions_filtre['Code IRIS'])
        """

        # On traite le ou les EvalStats non déjà créés
        # TODO : avec des codes IRIS et sans autrs sauvegardes des chemins des CSV traités par ailleurs, l'utilisateur est potentiellement obligé de sélectionner des CSV déjà traités (i.e. non pertinents) à la main → A améliorer
        # fe_evaluations_formation, statuts_csv = EvalStat.depuis_liste_codes_IRIS(self._codes_IRIS, fe_IRIS_sessions=self._fe_IRIS_sessions)
        statuts_csv = EvalStat_services.traiter_evalStat_depuis_liste_codes_IRIS(codes_IRIS=self._codes_IRIS, chemin_IRIS_sessions=chemin_IRIS_sessions)

        # A ce stade, toutes les valeurs de _codes_IRIS sont sensées être a minima présents dans IRIS avec données pour stats générales
        for code_IRIS in self._codes_IRIS:
            self._exploitationBilan["Exploités pour les stats générales"].append(code_IRIS)
        
        # Cas s'il y a au moins un CSV de lu
        if statuts_csv is not None:



            # === Gestion retour du traitement des CSV ===
            for code_IRIS, donnees in statuts_csv.items():
                if donnees["statut"] in ("Traité", "Exclu - CSV déjà dans fichier global"):
                    self._exploitationBilan["Exploités pour les évaluations (CSV présents)"].append(code_IRIS)

                if donnees["statut"] == "Exclu - Problème lecture CSV":
                    self._exploitationBilan["Exclus des évaluations (problème traitement CSV)"].append(code_IRIS)

                if donnees["statut"] == "Exclu - Fichier non existant":
                    self._exploitationBilan["Exclus des évaluations (CSV manquants)"].append(code_IRIS)

                if donnees["statut"] == "Exclu - CSV vide / Aucun retour":
                    self._exploitationBilan["Exclus des évaluations (CSV vide / aucun retour)"].append(code_IRIS)
                
                if donnees["statut"] == "Exclu - Code IRIS pas dans Extract IRIS sessions":
                    self._exploitationBilan["Exclus entièrement du bilan (non présent dans IRIS / mauvais code)"].append(code_IRIS)
                    # Alors ne pas exploiter pour les stats générales
                    self._exploitationBilan["Exploités pour les stats générales"].remove(code_IRIS)



            # === Préparation des dataframes des évaluations de la formation pour calcul des stats ===
            #self._df_stagiaires = fe_evaluations_formation._tableaux["Stagiaires"]._df  # Création d'un alias pour faciliter le code
            #TODO : est-ce que j'arrive à virer ces df d'ici pour ne pas faire de variables d'instance ou de return avec ces valeurs ?
            # df_stagiaires filtré sur les codes IRIS exploités pour le bilan (infos générales)
            # TODO : besoin _df_stagiaires_final dans _calculer_stats_criteres + __construit_champsFusionV3 ; _df_stagiaires_final_1ligne_session dans __construit_champsFusionV3
            self._df_stagiaires_final = self._df_stagiaires[self._df_stagiaires['Code IRIS'].isin(self._exploitationBilan["Exploités pour les stats générales"])]
            #vlog.print("Info", self._df_stagiaires_final)

            # df_stagiaires avec 1 ligne pour chaque session différente (pour avoir les infos globales issues de IRIS comme le nb d'apprenants)
            self._df_stagiaires_final_1ligne_session = self._df_stagiaires_final.drop_duplicates(subset=['Code IRIS'])  # Ne garde qu'une ligne par Code IRIS (la première rencontrée)
            #vlog.print("Info", self._df_stagiaires_final_1ligne_session)

        # Cas s'il n'y a aucun CSV
        else:
            # Si aucun CSV, alors il faut le traiter manuellement 
            for code_IRIS in self._codes_IRIS:
                self._exploitationBilan["Exclus des évaluations (CSV manquants)"].append(code_IRIS)
            
            #TODO : est-ce que j'arrive à virer ces df d'ici pour ne pas faire de variables d'instance ou de return avec ces valeurs ?
            self._df_stagiaires = None
            self._df_stagiaires_final = None
            self._df_stagiaires_final_1ligne_session = None
            



        # === Afficher les statuts des CSV employés pour le bilan ===
        #pprint(self._exploitationBilan)
        self._commentairesBilan += "\nListe des sessions :"
        for critere, lsessions in self._exploitationBilan.items():
            if lsessions:
                self._commentairesBilan += f"\n   • {critere} :" + "".join(f"\n       - {isession}" for isession in lsessions)
        vlog.print("Info", f"\n{self._commentairesBilan}")      

    def _calculer_stats_criteres(self) -> dict:
        """
        Retourne un dictionnaire de la forme :
        {
            "Nom du critère": {
                "Nombre": ...,
                "Moyenne": ...,
                "Commentaires": ...
            },
            ...
        }
        """      
        self._stats_stagiaires = {}

        # S'il n'y a pas de CSV disponibles pour les stats, alors ce n'est pas la peine de faire les stats
        if len(self._exploitationBilan["Exploités pour les évaluations (CSV présents)"]) != 0 :
            liste_criteres = self._df_stagiaires_final['Critère'].dropna().unique()

            #for critere in liste_criteres:
            #    if critere in self._criteres_a_enlever:
            #        continue

            for critere in liste_criteres:
                df_filtre = self._df_stagiaires_final[self._df_stagiaires_final['Critère'] == critere]
                nb = len(df_filtre)
                moyenne = df_filtre['Note'].mean() if nb > 0 else None

                commentaires_concat = "\n".join(
                    "• " + c.strip()
                    for c in df_filtre['Commentaires'].dropna().astype(str)
                    if c.strip() != ""
                )

                self._stats_stagiaires[critere] = {
                    "Nombre": nb,
                    "Moyenne": moyenne,
                    "Commentaires": commentaires_concat
                }
        else:
            vlog.print("Info", f"⚠️  Aucun CSV disponible pour ce bilan : les statistiques des stagiaires ne seront pas évaluées.")
            self._commentairesBilan += f"\n⚠️  Aucun CSV disponible pour ce bilan : les statistiques des stagiaires ne seront pas évaluées.\n"

        return self._stats_stagiaires

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
        # TODO : pour tous les self.df_sessions_filtre, il y aura probablement qq chose à faire quand j'aurai des codes multiples à cause de l'exclusion manuelle de certaines sessions

        self._titreFormation = self.intitule_formation
        # self._codeFormation = codeFormation  # (donné en argument)
        self._periodeSessionsEvaluees = self.periode_sessions_evaluees  # Evalué plus haut
        self._nbSessionsEvaluees = f"{len(self.df_sessions_filtre)} sessions"  # Valeur toutes les données  
        self._numerosSessions = "\n".join(self.df_sessions_filtre["N° Session"].dropna().astype(str).unique())
        self._nbApprenants = f"{self.df_sessions_filtre['Nb. Nommés'].sum()}"  # Valeur toutes les données
        #self._rp = ", ".join(self.df_sessions_filtre["Trigramme RP"].dropna().astype(str).unique())  # Valeur toutes les données
        self._rp = ", ".join(self.df_sessions_filtre["Nom responsable pédag."].dropna().astype(str).unique() + " " + self.df_sessions_filtre["Prénom responsable pédag."].dropna().astype(str).unique())  # Valeur toutes les données
        #self._af = ", ".join(self.df_sessions_filtre["Trigramme AF"].dropna().astype(str).unique())  # Valeur toutes les données
        self._af = ", ".join(self.df_sessions_filtre["Créée par"].dropna().astype(str).unique())  # Valeur toutes les données


        # On n'affecte les champs suivants que si des CSV sont disponibles pour les stats
        if len(self._exploitationBilan["Exploités pour les évaluations (CSV présents)"]) != 0 :

            # On évalue les données requises pour les stats
            nb_stagiaires_retours = self._df_stagiaires_final['NOM Prénom'].nunique()
            nb_apprenants = self._df_stagiaires_final_1ligne_session['Nb présents'].sum()

            stats_sous_3 = { # Dictionnaire pour les critères dont la moyenne est inférieure à 3 et non exclus (critères dans la liste self._CRITERES_A_ENLEVER)
                critere: valeurs
                for critere, valeurs in self._stats_stagiaires.items()
                if (
                    critere not in self._CRITERES_A_ENLEVER
                    and valeurs["Moyenne"] is not None
                    and valeurs["Moyenne"] < 3
                )
            }

            # On gère les données entre parenthèses s'il y a des sessions exclues d'une manière ou d'une autre
            #if len(self._exploitationBilan["Exploités pour les stats générales"]) != len(self._exploitationBilan["Exploités pour les évaluations (CSV présents)"]) :
            #    self._nbSessionsEvaluees += f" ({len(self._codes_IRIS_communs)})"  # Valeur si on ne prend que les données CSV
            #    self._numerosSessions += "\n".join("(" + self._df_stagiaires_final_1ligne_session["N° Session"].dropna().astype(str).unique() + ")")  # Valeur si on ne prend que les données CSV
            #    self._nbApprenants += f" ({nb_apprenants:.0f})"  # Valeur si on ne prend que les données CSV
            #    #self._rp += " (" + ", ".join(self._df_stagiaires_final_1ligne_session["Trigramme RP"].dropna().astype(str).unique()) + ")"  # Valeur si on ne prend que les données CSV
            #    #self._af += " (" + ", ".join(self._df_stagiaires_final_1ligne_session["Trigramme AF"].dropna().astype(str).unique()) + ")"  # Valeur si on ne prend que les données CSV
            #
            #    self._commentairesBilan = "Les valeurs entre parenthèses dans les statistiques générales sont les données des sessions pour lesquelles nous avons des CSV exploitables.\n" + self._commentairesBilan
            
            # Si des champs ne sont pas dans le CSV, alors on garde "" qui est déjà définit dans le constructeur
            try:
                self._satisfactionGlobale_moy = f'{self._stats_stagiaires["Satisfaction globale"]["Moyenne"]:.1f}/5'
            except:
                self._satisfactionGlobale_moy = "Pas de donnée"
            try:
                self._satisfactionGlobale_com = self._stats_stagiaires["Satisfaction globale"]["Commentaires"].replace("_x000D_", "\n")
            except:
                pass
            try:
                self._recommandation_moy = f'{self._stats_stagiaires["Recommanderiez-vous cette formation ?"]["Moyenne"]/5*100:.0f}%'  # (on divise par 5 car on a un booléen stcké sous forme de note sur 5 : 0 = False, 5 = True)
            except:
                self._recommandation_moy = "Pas de donnée"
            try:
                self._commentairesRemarquesSuggestions_com = self._stats_stagiaires["Commentaires, remarques, suggestions"]["Commentaires"].replace("_x000D_", "\n")
            except:
                pass
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


    # ==========================
    # === METHODES STATIQUES ===
    # ==========================
    @staticmethod
    def charge_IRIS_session_et_verifie_codeIRIS(code_IRIS:int, chemin_IRIS_sessions:Optional[Path]=None) -> DataFrame:
        """
        Charge IRIS Sessions, puis vérifie qu'il contient bienune ligne avec le code IRIS.

        Si IRIS session ne contient pas code IRIS, alors on coupe la procédure.

        Renvoie le dataframe df_session filtré sur le code IRIS

        :param code_IRIS: code IRIS à tester
        :type code_IRIS: int
        :param chemin_IRIS_sessions: chemin IRIS sessions à employer si l'utilisateur en souhaite un différent que celui par défaut
        :type chemin_IRIS_sessions: Optional[Path], optional
        :return: le dataframe df_session filtré sur le code IRIS
        :rtype: DataFrame
        """
        # On charge IRIS sessions
        iris_sessions = get_iris(typeExport="Sessions", chemin=chemin_IRIS_sessions)

        # On filtre le DataFrame d'IRIS avec le code IRIS
        df_sessions_filtre_codeIRIS = iris_sessions.df[iris_sessions.df['Code IRIS'] == code_IRIS]

        # On vérifie qu'on a bien une ligne sinon on sort
        if len(df_sessions_filtre_codeIRIS) < 1:
            print(df_sessions_filtre_codeIRIS)
            vlog.log_erreur(f"Le fichier IRIS Sessions ne contient pas ce code IRIS : {code_IRIS}")   

        return df_sessions_filtre_codeIRIS



    # =========================
    # === GETTERS / SETTERS === 
    # =========================
    
    @property
    def chemin_word_bilan_output(self) -> Path:
        """
        Renvoie le chemin de sortie du bilan de formation.
        Cette donnée est stockée dans le fichier de config (valeur par défaut).
        
        :return: Chemin de sortie du bilan de formation.
        :rtype: Path
        """
        return config.format_path(config.CHEMIN_WORD_BILAN_SESSION_OUTPUT, trigramme_formation=self.trigramme_formation, annee=self._annee, periode=f"{self._periodeSessionsEvaluees}", unite=config.UNITE)

    @property
    def periodeSessionsEvaluees(self):
        if self._periodeSessionsEvaluees is None:
            if len(self.code_IRIS == 1):  # Cas code IRIS unique
                f"Session {self.numero_session} uniquement ({self.mois_session} {self._annee})"
            else:
                self._periodeSessionsEvaluees = f"{self._periode} {self._annee}"

    @property
    def code_IRIS(self) -> int:
        """
        Le 1er code IRIS (code_IRIS au singulier donc on considère qu'on est sur un bilan contenant un seul code IRIS unique)
        """
        return self._codes_IRIS[0]

    @code_IRIS.setter
    def code_IRIS(self, valeur:int):
        """
        Renseigne un code IRIS (code_IRIS au singulier donc on considère qu'on est sur un bilan contenant un seul code IRIS unique)
        """
        if not self._codes_IRIS: # Cas d'une liste vide
            self._codes_IRIS.append(valeur)
        else:
            self._codes_IRIS[0] = valeur

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
    
    # Liens avec IRIS
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
        return self._iris_sessions.df_filtre_periode(trigramme_formation=self.trigramme_formation, annee=self._annee, periode=self._periode)

    @property
    def intitule_formation(self) -> str:
        """
        Intitulé de la formation (prend le nom de la dernière ligne pour avoir la dernière mise à jour)
        """
        return self.df_sessions_filtre["Session"].iloc[-1]

    @property
    def df_sessions_filtre_codeIRIS(self) -> DataFrame:
        """
        Renvoie le dataframe de l'extract IRIS filtré VTE selon plusieurs critères :
            - Code IRIS == self.code_IRIS
            - Statut Session != "Annulée" (toujours) ;
            - Nb. Nommés != 0 (toujours) ;
            - l'année de la session (année n du bilan de formation) ;
            - le trigramme de la formation traité par le bilan (self.trigramme_formation).
        
        :return: extract IRIS VTE filtré
        :rtype: DataFrame
        """
        df_sessions_filtre_codeIRIS = self.df_sessions_filtre[self.df_sessions_filtre['Code IRIS'] == self.code_IRIS]
        #print(self.df_sessions_filtre)

        if len(df_sessions_filtre_codeIRIS) < 1:
            print(df_sessions_filtre_codeIRIS)
            vlog.log_erreur(f"Le fichier IRIS Sessions ne contient pas ce code IRIS : {self.code_IRIS}")        

        return df_sessions_filtre_codeIRIS

    @property
    def mois_session(self) -> str:
        """
        Mois de la session dans le cadre d'une session unique (len(codes_IRIS)=1)
        """
        return mois_fr_depuis_date(self.df_sessions_filtre_codeIRIS["Date début ses."].iloc[0])

    @property
    def numero_session(self) -> str:
        """
        Numéro de la session dans le cadre d'une session unique (len(codes_IRIS)=1)
        """
        return self.df_sessions_filtre_codeIRIS["N° Session"].iloc[0]

    @property
    def periode_sessions_evaluees(self) -> str:
        """
        Période de la session dans le cadre d'une session unique (len(codes_IRIS)=1).

        Construction :
            - si un seul code IRIS : f"Session {self.numero_session} uniquement ({self.mois_session} {self._annee})"
            - si plusieurs codes IRIS : f"{self._periode} {self._annee}"
        """
        if len(self._codes_IRIS) == 1:
            return f"Session {self.numero_session} uniquement ({self.mois_session} {self._annee})"
        elif len(self._codes_IRIS) > 1:
            return f"{self._periode} {self._annee}"
        else:
            vlog.log_erreur("J'appelle periode_sessions_evaluees alors que len(self.codes_IRIS)=0")
    
    
    

class BilanSession_OLD:

    def _charger_IRIS_sessions(cls) -> None:
        """
        Charge le fichier Excel IRIS Sessions si ce n'est pas déjà fait.
        Cette méthode met à jour _fe_IRIS_sessions et _df_sessions.
        """

        # TODO il faudra que je vire ces anciennes méthodes statiques par celles de la classe IRIS : charger_excel_IRIS_VTE
        # Vérifie existance de fe_IRIS sinon on le charge
        cls._fe_IRIS_sessions = IRIS.charger_excel_IRIS_sessions(fe_IRIS_sessions=cls._fe_IRIS_sessions)

        # On convertit les colonnes
        cls._df_IRIS_sessions = IRIS.convertit_types_colonnes_df_sessions(fe_IRIS_sessions=cls._fe_IRIS_sessions)

    

    def _demande_sessions_a_exclure(self) -> None:
        """
        Demande à l'utilisateur les sessions qu'il souhaite exclure du bilan formation
        On retourne un dataframe df_sessions_filtre à jour
        """
        # On prépare le dataframe que l'on va filtrer
        df_sessions_a_traiter = self.df_sessions_filtre

        # Adaptation format date
        df_sessions_a_traiter['Date début ses.'] = pd.to_datetime(df_sessions_a_traiter['Date début ses.']).dt.strftime("%d/%m/%Y")
        df_sessions_a_traiter['Date fin ses.'] = pd.to_datetime(df_sessions_a_traiter['Date fin ses.']).dt.strftime("%d/%m/%Y")
        
        # On affiche à l'utilisateur les sessions et dates et statuts 
        vlog.print("Info", f"\nListe des sessions {self.trigramme_formation} dans {self._iris_sessions.chemin.name} - {self._annee}", style=["jaune"])
        IRIS_traite.affiche_df_moins_de_colonnes(df_sessions_a_traiter)
        
        # On demande à l'utilisateur les sessions qu'il veut exclure
        exclusionSessions = IRIS.demander_liste_codes_IRIS()
        if exclusionSessions:  # si la liste n'est pas vide
            # On trace l'exclusion des sessions
            for session_exclue in exclusionSessions:
                self._exploitationBilan["Exclus entièrement du bilan (exclus par utilisateur)"].append(self.df_sessions_filtre.loc[self.df_sessions_filtre["Code IRIS"] == session_exclue, "N° Session"].iloc[0])

            # On met à jour _df_sessions_filtre en enlevant les sessions exclues
            df_sessions_a_traiter = df_sessions_a_traiter[~df_sessions_a_traiter['Code IRIS'].isin(exclusionSessions)]
        else:
            # la liste est vide, on ne filtre rien, on garde tout
            pass

        # On affiche à l'utilisateur les sessions finalement retenues
        vlog.print("Info", "\nSessions retenues pour le bilan :", style=["jaune"])
        IRIS_traite.affiche_df_moins_de_colonnes(df_sessions_a_traiter)

        # On met à jour _df_sessions_filtre en enlevant les sessions exclues → Plus besoin : on a ça dans _maj_evalstat_formation
        #self._exploitationBilan["Exploités pour les stats générales"] = self.df_sessions_filtre["N° Session"].tolist()

        # On renseigne _codes_IRIS
        self._codes_IRIS = df_sessions_a_traiter["Code IRIS"].tolist()

    def _demande_sessions_a_exclure_VERSION_AVANT_MODIF_VOIR_METHODE_JUSTE_EN_HAUT(self) -> None:
        """
        Demande à l'utilisateur les sessions qu'il souhaite exclure de la période choisie
        On retourne un dataframe df_sessions_filtre à jour
        """
        # On affiche à l'utilisateur les sessions et dates et statuts 
        vlog.print("Info", f"\nListe des sessions {self.trigramme_formation} dans {self._fe_IRIS_sessions._chemin_fichier.name} - {self._periode} {self._annee}", style=["jaune"])

        # Adaptation format date
        self.df_sessions_filtre['Date début ses.'] = pd.to_datetime(self.df_sessions_filtre['Date début ses.']).dt.strftime("%d/%m/%Y")
        self.df_sessions_filtre['Date fin ses.'] = pd.to_datetime(self.df_sessions_filtre['Date fin ses.']).dt.strftime("%d/%m/%Y")

        print(tabulate(
            self.df_sessions_filtre[['Code IRIS', 'Trigramme RP', 'Trigramme AF', 'Date début ses.', 'Date fin ses.', 'Durée réal. (J.)', 'Nb. Présents', 'Statut Session', 'N° Session']], 
            headers='keys', 
            tablefmt='pretty', 
            showindex=False
        ))
        
        # On demande à l'utilisateur les sessions qu'il veut exclure
        exclusionSessions = IRIS.demander_liste_codes_iris()
        if exclusionSessions:  # si la liste n'est pas vide
            # On trace l'exclusion des sessions
            for session_exclue in exclusionSessions:
                self._exploitationBilan["Exclus entièrement du bilan (exclus par utilisateur)"].append(self.df_sessions_filtre.loc[self.df_sessions_filtre["Code IRIS"] == session_exclue, "N° Session"].iloc[0])

            # On met à jour _df_sessions_filtre en enlevant les sessions exclues
            self.df_sessions_filtre = self.df_sessions_filtre[~self.df_sessions_filtre['Code IRIS'].isin(exclusionSessions)]
        else:
            # la liste est vide, on ne filtre rien, on garde tout
            pass

        # On affiche à l'utilisateur les sessions finalement retenues
        vlog.print("Info", "\nSessions retenues pour le bilan :", style=["jaune"])
        print(tabulate(
            self.df_sessions_filtre[['Code IRIS', 'Trigramme RP', 'Trigramme AF', 'Date début ses.', 'Date fin ses.', 'Durée réal. (J.)', 'Nb. Présents', 'Statut Session', 'N° Session']], 
            headers='keys', 
            tablefmt='pretty', 
            showindex=False
        ))

        # On met à jour _df_sessions_filtre en enlevant les sessions exclues → Plus besoin : on a ça dans _maj_evalstat_formation
        #self._exploitationBilan["Exploités pour les stats générales"] = self.df_sessions_filtre["N° Session"].tolist()

        # On renseigne _codes_IRIS
        self._codes_IRIS = self.df_sessions_filtre["Code IRIS"].tolist()

    def _maj_evalstat_formation_BAK(self) -> None:
        """
        Met à jour l'Excel evalstat de la formation si des sessions demandées par l'utilisateur ne s'y trouvent pas
        (on regarde les CSV qui ne sont pas dans le fichier Excel global à partir de la liste df_sessions_filtre['Code IRIS'])

        """
        # On ouvre le fichier Excel global des évaluations de la formation
        timer.debut("Lecture du fichier Excel global des évaluations des stagiaires")
        _es = EvalStat.ouvrir_ou_creer_evaluationsFormation(self.trigramme_formation)
        self._df_stagiaires = self.eval_formation._fe_evaluations_formation._tableaux["Stagiaires"]._df  # Création d'un alias pour faciliter le code
        
        # Retype "Trigramme formation" et "Code IRIS"
        self._df_stagiaires["Trigramme formation"] = self._df_stagiaires["Trigramme formation"].astype(str)
        self._df_stagiaires["Code IRIS"] = (
            pd.to_numeric(self._df_stagiaires["Code IRIS"], errors="coerce")  # "13414" ou 13414.0 → 13414
            .astype("Int64")                               # reste un entier (NaN compatible)
            .astype(str)                                   # enfin en texte propre
        )
        

        #print(f"\nÉtat de Excel évaluations filtré sur période et trigramme ({self.eval_formation._chemin_excel_evaluations_formation}) :")
        #pprint(self._df_stagiaires)

        vlog.ajouter_message("OK", f"✅ Lecture du fichier Excel global des évaluations des stagiaires {self.eval_formation._fe_evaluations_formation.chemin_fichier}")
        timer.fin()



        # --- Gestion des CSV manquants
        # On isole depuis ce fichier les CSV manquants (df_sessions_filtre = sessions demandées par l'utilisateur ; es._df_evaluations_formation = existant dans l'excel global)
        self._code_session_absents = list(set(self.df_sessions_filtre['Code IRIS']) - set(self._df_stagiaires['Code IRIS']))
        #print(set(self.df_sessions_filtre['Code IRIS']))
        #print(set(self._df_stagiaires['Code IRIS']))
        if self._code_session_absents:
            vlog.print("Info", f"🔎 Des codes session sont absents de l'Excel global des évaluations des stagiaires : {self._code_session_absents}")

            # Pour les CSV manquants, on demande à l'utilisateur de sélectionner les CSV à la main
            chemins_csv_a_traiter = {}
            chemins_csv_manquants = {}
            for code_IRIS in self._code_session_absents:
                # L'utilisateur sélectionne le CSV de code_IRIS
                chemin_csv_supp = self.eval_formation._filedialog_csv(code_IRIS=code_IRIS, trigramme_formation=self.trigramme_formation)
                
                if chemin_csv_supp is not None: # CSV sélectionné 
                    chemins_csv_a_traiter[code_IRIS] = chemin_vers_unc(chemin_csv_supp)
                    vlog.print("OK", f"✅ CSV {code_IRIS} : {chemin_csv_supp}")
                else: #CSV manquant
                    # On trace l'exclusion du CSV
                    self._exploitationBilan["Exclus des évaluations (CSV manquants)"].append(self.df_sessions_filtre.loc[self.df_sessions_filtre["Code IRIS"] == code_IRIS, "N° Session"].iloc[0])

                    chemins_csv_manquants[code_IRIS] = f"❌ Pas de CSV disponible pour cette session {code_IRIS}, session exclue."
                    vlog.print("CSV non disponible", f"❌ CSV {code_IRIS} : pas de CSV disponible pour cette session, session exclue.")
            #print(chemins_csv_manquants)


            # On traite les CSV sélectionnés par l'utilisateur
            tuple_csv_stagiaires = tuple(val for val in chemins_csv_a_traiter.values())
            if tuple_csv_stagiaires:
                vlog.print("Info", f"⏳ Traitement des CSV non déjà présents dans {os.path.basename(self.eval_formation._fe_evaluations_formation.chemin_fichier)} :" + "".join(f"\n• {chemin}" for chemin in tuple_csv_stagiaires))
                self.eval_formation = EvalStat.depuis_tuple_csv_stagiaires(tuple_csv_stagiaires=tuple_csv_stagiaires, fe_IRIS_sessions=self._fe_IRIS_sessions)
                self._df_stagiaires = self.eval_formation._fe_evaluations_formation._tableaux["Stagiaires"]._df.copy()  # MàJ de l'alias pour faciliter le code
                self._df_stagiaires["Trigramme formation"] = self._df_stagiaires["Trigramme formation"].astype(str)
                self._df_stagiaires["Code IRIS"] = (
                    pd.to_numeric(self._df_stagiaires["Code IRIS"], errors="coerce")  # "13414" ou 13414.0 → 13414
                    .astype("Int64")                               # reste un entier (NaN compatible)
                    .astype(str)                                   # enfin en texte propre
                )

            #print(self._df_stagiaires)

            # Après traitement des CSV on regarde dans self._df_stagiaires les N° de session
            #self._exploitationBilan["Exploités pour les évaluations (CSV présents)"] = list(set(self.df_sessions_filtre['Code IRIS']) & set(self._df_stagiaires['Code IRIS']))
            #codesIRIS_absents_apres_traitement = list(set(self._code_session_absents) - set(self._df_stagiaires['Code IRIS']) - set(self._exploitationBilan["Exclus des évaluations (CSV manquants)"]))
            #print("codesIRIS_absents_apres_traitement : ", codesIRIS_absents_apres_traitement)
            #print(set(self._code_session_absents))
            #print(set(self._df_stagiaires['Code IRIS']))

            #self._exploitationBilan["Exclus des évaluations (problème traitement CSV)"] = self.df_sessions_filtre[self.df_sessions_filtre['Code IRIS'].isin(codesIRIS_absents_apres_traitement)]['N° Session'].tolist()


            # Affichage selon retours traitement EvalStat
            if (self.eval_formation._chemins_csv_traites or self.eval_formation._chemins_csv_probleme or self.eval_formation._chemins_csv_exclus):
                vlog.print("Info", "\nBilan traitement des CSV :")
            if self.eval_formation._chemins_csv_traites:
                vlog.print("OK", f"   ✅ Chemins traités : " + "".join(f"\n      • {i_csv}" for i_csv in self.eval_formation._chemins_csv_traites))
            if self.eval_formation._chemins_csv_probleme:
                vlog.print("Problème traitement CSV", f"   ❌ Chemins ayant eu des problèmes (exclus) : " + "".join(f"\n      • {i_csv}" for i_csv in self.eval_formation._chemins_csv_probleme))
            if self.eval_formation._chemins_csv_exclus:
                vlog.print("Problème traitement CSV", f"   ❌ Chemins exclus lors du traitement : " + "".join(f"\n      • {i_csv}" for i_csv in self.eval_formation._chemins_csv_exclus))
                
        else:
            vlog.print("OK", f"✅ Tous les CSV sont bien déjà importés dans {self.eval_formation._fe_evaluations_formation.chemin_fichier.name}")



        self._codes_IRIS_communs = list(set(self.df_sessions_filtre['Code IRIS']) & set(self._df_stagiaires['Code IRIS']))
        #codes_IRIS_absents = list(set(self.df_sessions_filtre['Code IRIS']) - set(self._df_stagiaires['Code IRIS']))
        #vlog.print("Info", f"\nIn fine, voici la liste des éléments qui seront :")
        #vlog.print("Info", f"\t• inclus dans le bilan : {self._codes_IRIS_communs}")
        #vlog.print("Info", f"\t• exclus du bilan : {codes_IRIS_absents}")


 
        
        # On remet à jour les alias avec les nouvelles données (après traitement CSV)
        self._df_stagiaires_final = self._df_stagiaires[self._df_stagiaires['Code IRIS'].isin(self._codes_IRIS_communs)]
        #vlog.print("Info", self._df_stagiaires_final)
        self._df_stagiaires_final_1ligne_session = self._df_stagiaires_final.drop_duplicates(subset=['Code IRIS'])  # Ne garde qu'une ligne par Code IRIS (la première rencontrée)
        #vlog.print("Info", self._df_stagiaires_final_1ligne_session)

        # Dernière vérif qu'on a bien tout importé les CSV dans l'excel global
        self._codes_IRIS_absents_fin = list(set(self.df_sessions_filtre['Code IRIS']) - set(self._df_stagiaires_final['Code IRIS']))
        #if not self._codes_IRIS_absents_fin:
            #print(self._code_session_absents_fin)
            #vlog.print("Erreur", f"Erreur il reste encore des disparités avec des CSV non importés qui sont sensés être dans le bilan après traitement : {self._codes_IRIS_absents_fin}")
            #exit()

        # On définit la liste finale des N° session traités
        codesIRIS_avec_CSV = list(set(self.df_sessions_filtre['Code IRIS']) & set(self._df_stagiaires_final['Code IRIS']))
        self._exploitationBilan["Exploités pour les évaluations (CSV présents)"] = self.df_sessions_filtre[self.df_sessions_filtre['Code IRIS'].isin(codesIRIS_avec_CSV)]['N° Session'].tolist()
        self._exploitationBilan["Exclus des évaluations (problème traitement CSV)"] = list(
            set(self._exploitationBilan["Exploités pour les stats générales"])
            - set(self._exploitationBilan["Exploités pour les évaluations (CSV présents)"])
            - set(self._exploitationBilan["Exclus des évaluations (CSV manquants)"])
        )
        #self._exploitationBilan["Exploités pour les évaluations (CSV présents)"] = list(
        #    set(self._exploitationBilan["Exploités pour les stats générales"])
        #    - set(self._exploitationBilan["Exclus des évaluations (CSV manquants)"])
        #    - set(self._exploitationBilan["Exclus des évaluations (problème traitement CSV)"])
        #    )

        #pprint(self._exploitationBilan)
        
        self._commentairesBilan += "\nListe des sessions :"
        for critere, lsessions in self._exploitationBilan.items():
            if lsessions:
                self._commentairesBilan += f"\n   • {critere} :" + "".join(f"\n       - {isession}" for isession in lsessions)
        vlog.print("Info", f"\n{self._commentairesBilan}")        
