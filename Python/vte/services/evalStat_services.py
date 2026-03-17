
# ==========================================================================================
# CLASSE EVALSTAT_SERVICES
#
# Pour lancer des fonctions et méthodes faisant appel à EvalStat et classes fonctionnelles mères (ex. : Formation)
# ==========================================================================================
from collections.abc import Iterable
from pathlib import Path
from colorama import Fore, Style
from typing import Optional, Tuple

from colorama import Fore

from vte.domain.formation import Formation
from vte.domain.evalStat import EvalStat_session
from vte.utils.utils_instn import recupere_trig_formation_depuis_chemin


class EvalStat_services:
    """
    Services métier autour des exports IRIS.
    """
    @staticmethod
    def traiter_evalStat_depuis_dico_csv(dico_csv:dict[str, dict[int, Path]], ouvrirDossier:bool=False) -> dict[int, dict[str, str]]:
        """
        A partir d'un dictionnaire CSV stagiaire
        Permet de :
           - générer le fichier d'évaluation de chaque session (via le CSV)
           - mettre à jour le fichier d'évaluation de la formation (celui qui concatène les CSV de toutes les sessions) [Il est créé ou on l'append avec les nouvelles valeurs]
        
        On exclue du traitement les chemin_csv_session qui sont déjà dans l'évaluation de la formation (on considère que le CSV a déjà été traité)

        :param dico_csv: dictionnaire CSV [trigramme_formation][code_IRIS]: chemin_csv généré par EvalStat_session.construire_dictionnaire_trigrammeFormation_codeIRIS_cheminsCSV_depuis_iterableCSV
        :type dico_csv: dict[str, dict[int, Path]]
        :param ouvrirDossier: Ouvre le répertoire de l'EvalStat généré. Defaut = False.
        :type ouvrirDossier: bool, optional
        :return: Une liste de dictionnaires {codeIRIS,{"fichier": chemin_csv,"statut": statut_csv}}. Je pourrai accéder à la valeur par nom_dico[codeIRIS]["fichier"] ou nom_dico[codeIRIS]["statut"]
        :rtype: dict[int, dict[str, str]]
        """
        # Dictionnaire des retours du traitement
        statuts_csv:dict[int, dict[str, str]] = {}  # [trigramme_formation][code_IRIS]: chemin_csv
        
        # On boucle sur les trigrammes des formations
        for trigramme_formation in dico_csv.keys():
            print(f"\n\n{Style.BRIGHT}{Fore.RED}Gestion des formations {trigramme_formation}")
            # Création de l'instance formation
            formation = Formation.avec_ouverture_evalStat(trigramme_formation)
            
            # On crée les EvalStat pour chaque code IRIS de la formation
            for code_IRIS, chemin_csv in dico_csv[trigramme_formation].items():
                formation.ajout_sessions(code_IRIS)
                session = formation.sessions[code_IRIS]
                session.ajout_evalStat_avec_traitement(
                    chemin_csv=chemin_csv,
                    ecrire_eval_formation=False,  # On sauvegardera après la boucle de traitement
                    ouvrirDossier=ouvrirDossier
                )

                # On met à jour le statut CSV de sortie
                statuts_csv[code_IRIS] = {
                        "fichier": chemin_csv.name,
                        "statut": session.eval.statut
                        }

                #print(f"Fin traitement : {code_IRIS}\t{chemin_csv.name}\t{session.eval.statut_csv}")
            
            # On sauvegarde l'évaluation de la formation
            formation.eval.ecritdf_et_sauve_siModif()

        return statuts_csv

    @staticmethod
    def traiter_evalStat_depuis_iterable_de_csv(tuple_csv_stagiaires:Iterable[Path|str], ouvrirDossier:bool=False) -> dict[str, dict[str, str]]:
        """
        A partir d'un itérable de chemins de CSV stagiaire
        Permet de :
           - générer le fichier d'évaluation de chaque session (via le CSV)
           - mettre à jour le fichier d'évaluation de la formation (celui qui concatène les CSV de toutes les sessions) [Il est créé ou on l'append avec les nouvelles valeurs]
        
        On exclue du traitement les chemin_csv_session qui sont déjà dans l'évaluation de la formation (on considère que le CSV a déjà été traité)

        :param tuple_csv_stagiaires: Itérable des CSV des évaluations à traiter.
        :type tuple_csv_stagiaires: Iterable[Path | str]
        :param ouvrirDossier: Ouvre le répertoire de l'EvalStat généré. Defaut = False.
        :type ouvrirDossier: bool, optional
        :return: Une liste de dictionnaires {codeIRIS,{"fichier": chemin_csv,"statut": statut_csv}}. Je pourrai accéder à la valeur par nom_dico[codeIRIS]["fichier"] ou nom_dico[codeIRIS]["statut"]
        :rtype: dict[str, dict[str, str]]
        """
        # Dictionnaire des csv sous une forme qui nous arrange pour le traitement à venir
        dico_csv = EvalStat_session.construire_dictionnaire_trigrammeFormation_codeIRIS_cheminsCSV_depuis_iterableCSV(tuple_csv_stagiaires)

        # Dictionnaire des retours du traitement
        statuts_csv = EvalStat_services.traiter_evalStat_depuis_dico_csv(dico_csv=dico_csv, ouvrirDossier=ouvrirDossier)
        
        return statuts_csv

    @staticmethod
    def traiter_evalStat_depuis_liste_codes_IRIS(codes_IRIS:Iterable[int], formation:Optional[Formation]=None, ouvrirDossier:bool=False) -> dict[str, dict[str, str]]:
        """
        A partir d'un itérable de codes IRIS
        Permet de :
           - générer le fichier d'évaluation de chaque session (via le code IRIS)
           - mettre à jour le fichier d'évaluation de la formation (celui qui concatène les CSV de toutes les sessions) [Il est créé ou on l'append avec les nouvelles valeurs]

        :param codes_IRIS: Itérable des codes IRIS à traiter.
        :type codes_IRIS: Iterable[int]
        :param formation: l'evalFormation pour vérifier si on doit traiter l'EvalStat + le trigramme de la formation qui sert à mieux pointer le répertoire pour sélectrionner le CSV.
        :type formation: Optional[Formation_protocol], Optional
        :param ouvrirDossier: Ouvre le répertoire de l'EvalStat généré. Defaut = False.
        :type ouvrirDossier: bool, optional
        :return: Une liste de dictionnaires {codeIRIS,{"fichier": chemin_csv,"statut": statut_csv}}. Je pourrai accéder à la valeur par nom_dico[codeIRIS]["fichier"] ou nom_dico[codeIRIS]["statut"]
        :rtype: dict[str, dict[str, str]]
        """
        # Dictionnaire des csv sous une forme qui nous arrange pour le traitement à venir ; on fait une première exclusion des EvalStat déjà dans eval formation
        dico_csv = EvalStat_session.construire_dictionnaire_trigrammeFormation_codeIRIS_cheminsCSV_depuis_iterableCodesIRIS(codes_IRIS=codes_IRIS, formation=formation)

        # Dictionnaire des retours du traitement
        statuts_csv = EvalStat_services.traiter_evalStat_depuis_dico_csv(dico_csv=dico_csv, ouvrirDossier=ouvrirDossier)

        return statuts_csv