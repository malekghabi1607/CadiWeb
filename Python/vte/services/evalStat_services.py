
# ==========================================================================================
# CLASSE EVALSTAT_SERVICES
#
# Pour lancer des fonctions et méthodes faisant appel à EvalStat et classes fonctionnelles mères (ex. : Formation)
# ==========================================================================================
from __future__ import annotations
from collections.abc import Iterable
from pathlib import Path
from colorama import Fore, Style
from typing import Optional, Protocol, Tuple

from colorama import Fore

from vte.core.iris_referentiel import dico_trigrammes_codes_IRIS_a_partir_de_codes_IRIS
#from vte.domain.formation import Formation
from vte.domain.evalStat import EvalStat_session, EvalStat_formation
from vte.domain.session import Session

# ======================================================================================
# PROTOCOLES
# (pour faire passer les informations des objets parents sans ref circulaires)
# ======================================================================================
class Formation_protocol(Protocol):
    """
    Protocol de Formation : permet de simuler une formation en évitant les références circulaires
    """
    @classmethod
    def avec_ajout_sessions(cls, trigramme_formation: str, codes_IRIS:int|Iterable[int]) -> Formation_protocol: ...
    
    @property
    def trigramme_formation(self) -> Optional[str]: ...

    @property
    def sessions(self) -> list[Session]: ...

    @property
    def eval(self) -> Optional[EvalStat_formation]: ...

    @property
    def tuple_codesIRIS_de_sessions(self) -> tuple[int]: ...

    def ajout_sessions(self, codes_IRIS:int|Iterable[int]) -> None: ...





class EvalStat_services:
    """
    Services métier autour des exports IRIS.
    """


    # ======================================
    # === TRAITEMENTS PLUSIEURS EVALSTAT ===
    # ======================================
    @staticmethod
    #def ouvrir_ou_traiter_evalStat_depuis_liste_codes_IRIS(codes_IRIS:Iterable[int], formation:Optional[Formation]=None, ouvrir_dossier:bool=False) -> dict[str, dict[str, str]]:
    def ouvrir_ou_traiter_evalStat_depuis_liste_codes_IRIS(codes_IRIS:Iterable[int], formation:Optional[Formation_protocol]=None, ouvrir_dossier:bool=False) -> None:
        """
        A partir d'un itérable de codes IRIS
        Permet de :
           - générer le fichier d'évaluation de chaque session (via le code IRIS)
           - mettre à jour le fichier d'évaluation de la formation (celui qui concatène les CSV de toutes les sessions) [Il est créé ou on l'append avec les nouvelles valeurs]

        :param codes_IRIS: Itérable des codes IRIS à traiter.
        :type codes_IRIS: Iterable[int]
        :param formation: l'evalFormation pour vérifier si on doit traiter l'EvalStat + le trigramme de la formation qui sert à mieux pointer le répertoire pour sélectrionner le CSV.
        :type formation: Optional[Formation_protocol], Optional
        :param ouvrir_dossier: Ouvre le répertoire de l'EvalStat généré. Defaut = False.
        :type ouvrir_dossier: bool, optional
        :return: Une liste de dictionnaires {codeIRIS,{"fichier": chemin_csv,"statut": statut_csv}}. Je pourrai accéder à la valeur par dico_csv[codeIRIS]["fichier"] ou nom_dico[codeIRIS]["statut"]
        :rtype: dict[str, dict[str, str]]
        """

        # TODO : je pourrai faire une fonction pour générer le dico dico_ouvrir_ou_traiter : ce sera plus facile à lire
        # A partir de codes_IRIS, je dois :
        # créer un dico dico_ouvrir_ou_traiter{formation: sessions} (liste toutes les sessions que je dois soit ouvrir soit traiter)     
        dico_ouvrir_ou_traiter:dict[Formation_protocol, list[Session]] = {}

        # On crée le dictionnaire dico_ouvrir_ou_traiter
        if formation is not None:
            # On vérifie que les sessions sont créées dans formation sinon on les crée
            codes_inexistants_dans_formation = [code_IRIS for code_IRIS in codes_IRIS if code_IRIS not in formation.tuple_codesIRIS_de_sessions]
            formation.ajout_sessions(codes_inexistants_dans_formation)
            
            # On crée le dico (liste toutes les sessions que je dois soit ouvrir soit traiter) 
            #dico_ouvrir_ou_traiter[formation] = [session for session in formation.sessions if session.code_IRIS in codes_a_traiter]
            dico_ouvrir_ou_traiter[formation] = [session for session in formation.sessions if session.code_IRIS]
            
        elif formation is None:  # Si formation n'est pas donné
            # On cherche les trigrammes dans IRIS sessions
            dico_trigrammes_codesIRIS = dico_trigrammes_codes_IRIS_a_partir_de_codes_IRIS(codes_IRIS=codes_IRIS)

            # On crée les formations et on initialise le dico avec
            for trigramme_formation, tcodes_IRIS in dico_trigrammes_codesIRIS.items():
                # On crée l'objet formation et ses sessions
                formation = Formation_protocol.avec_ajout_sessions(trigramme_formation=trigramme_formation, codes_IRIS=tcodes_IRIS)

                # On crée le dico                
                dico_ouvrir_ou_traiter[formation] = [session for session in formation.sessions if session.code_IRIS]

                    
        # On boucle sur les formations
        for formation, sessions in dico_ouvrir_ou_traiter.items():
            print(f"\n\n{Style.BRIGHT}{Fore.RED}Gestion des formations {formation.trigramme_formation}")
            # Création de l'instance formation → Déjà fait en traitant EvalStat session
            #formation.ouvrir_ou_creer_evalStat() # = Formation.avec_ouverture_ou_creation_evalStat(trigramme_formation)
            
            # On crée les EvalStat pour chaque code IRIS de la formation
            for session in sessions:
                session.eval = EvalStat_session.avec_ouverture_ou_traitement(
                    session=session,
                    ecrire_eval_formation=False,  # On sauvegardera après la boucle de traitement
                    ouvrir_dossier=ouvrir_dossier
                )
            
            # On sauvegarde l'évaluation de la formation
            formation.eval.ecrit_et_sauve_df_siModif()

    # TODO : non fonctionnel en l'état -> à retravailler
    @staticmethod
    def ouvrir_ou_traiter_evalStat_depuis_dico_csv(dico_csv:dict[str, dict[int, Path]], ouvrir_dossier:bool=False) -> dict[int, dict[str, str]]:
        """
        A partir d'un dictionnaire CSV stagiaire
        Permet de :
           - générer le fichier d'évaluation de chaque session (via le CSV)
           - mettre à jour le fichier d'évaluation de la formation (celui qui concatène les CSV de toutes les sessions) [Il est créé ou on l'append avec les nouvelles valeurs]
        
        On exclue du traitement les chemin_csv_session qui sont déjà dans l'évaluation de la formation (on considère que le CSV a déjà été traité)

        :param dico_csv: dictionnaire CSV [trigramme_formation][code_IRIS]: chemin_csv généré par EvalStat_session.construire_dictionnaire_trigrammeFormation_codeIRIS_cheminsCSV_depuis_iterableCSV
        :type dico_csv: dict[str, dict[int, Path]]
        :param ouvrir_dossier: Ouvre le répertoire de l'EvalStat généré. Defaut = False.
        :type ouvrir_dossier: bool, optional
        :return: Une liste de dictionnaires {codeIRIS,{"fichier": chemin_csv,"statut": statut_csv}}. Je pourrai accéder à la valeur par nom_dico[codeIRIS]["fichier"] ou nom_dico[codeIRIS]["statut"]
        :rtype: dict[int, dict[str, str]]
        """
        # Dictionnaire des retours du traitement
        statuts_csv:dict[int, dict[str, str]] = {}  # [trigramme_formation][code_IRIS]: chemin_csv
        
        # On boucle sur les trigrammes des formations
        for trigramme_formation in dico_csv.keys():
            print(f"\n\n{Style.BRIGHT}{Fore.RED}Gestion des formations {trigramme_formation}")
            # Création de l'instance formation
            formation = Formation.avec_ouverture_ou_creation_evalStat(trigramme_formation)
            
            # On crée les EvalStat pour chaque code IRIS de la formation
            for code_IRIS, chemin_csv in dico_csv[trigramme_formation].items():
                formation.ajout_sessions(code_IRIS)
                session = formation.sessions[code_IRIS]  # Alias
                session.eval = EvalStat_session.avec_ouverture_ou_traitement(
                    chemin_csv=chemin_csv,
                    ecrire_eval_formation=False,  # On sauvegardera après la boucle de traitement
                    ouvrir_dossier=ouvrir_dossier
                )

                # On met à jour le statut CSV de sortie
                statuts_csv[code_IRIS] = {
                        "fichier": chemin_csv.name,
                        "statut": session.eval.statut
                        }

                #print(f"Fin traitement : {code_IRIS}\t{chemin_csv.name}\t{session.eval.statut_csv}")
            
            # On sauvegarde l'évaluation de la formation
            formation.eval.ecrit_et_sauve_df_siModif()

        return statuts_csv

    # TODO : non fonctionnel en l'état -> à retravailler
    @staticmethod
    def ouvrir_ou_traiter_evalStat_depuis_iterable_de_csv(tuple_csv_stagiaires:Iterable[Path|str], ouvrir_dossier:bool=False) -> dict[str, dict[str, str]]:
        """
        A partir d'un itérable de chemins de CSV stagiaire
        Permet de :
           - générer le fichier d'évaluation de chaque session (via le CSV)
           - mettre à jour le fichier d'évaluation de la formation (celui qui concatène les CSV de toutes les sessions) [Il est créé ou on l'append avec les nouvelles valeurs]
        
        On exclue du traitement les chemin_csv_session qui sont déjà dans l'évaluation de la formation (on considère que le CSV a déjà été traité)

        :param tuple_csv_stagiaires: Itérable des CSV des évaluations à traiter.
        :type tuple_csv_stagiaires: Iterable[Path | str]
        :param ouvrir_dossier: Ouvre le répertoire de l'EvalStat généré. Defaut = False.
        :type ouvrir_dossier: bool, optional
        :return: Une liste de dictionnaires {codeIRIS,{"fichier": chemin_csv,"statut": statut_csv}}. Je pourrai accéder à la valeur par nom_dico[codeIRIS]["fichier"] ou nom_dico[codeIRIS]["statut"]
        :rtype: dict[str, dict[str, str]]
        """
        # Dictionnaire des csv sous une forme qui nous arrange pour le traitement à venir
        dico_csv = EvalStat_session.construire_dictionnaire_trigrammeFormation_codeIRIS_cheminsCSV_depuis_iterableCSV(tuple_csv_stagiaires)

        # Dictionnaire des retours du traitement
        statuts_csv = EvalStat_services.ouvrir_ou_traiter_evalStat_depuis_dico_csv(dico_csv=dico_csv, ouvrir_dossier=ouvrir_dossier)
        
        return statuts_csv


