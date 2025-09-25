from __future__ import annotations
from .office import *

import math

from dataclasses import dataclass
from collections import defaultdict
from tabulate import tabulate

from mailmerge import MailMerge
from tkinter import ttk, messagebox





### --------------------------------------------------------------------
#  Définitions classes et fonctions spécifiques INSTN
### --------------------------------------------------------------------

@dataclass
class InfosExportsIRIS:
    repertoire: str|None
    nom_fichier: str|None
    chemin_fichier: str|None
    nom_onglet: str|None
    nbLignes_avantET: int|None

class PropExportIRIS:
    """C'est une fabrique. Contient toutes les propriétés des fichiers Excel issus des Exports IRIS (structure de configuration d'un export IRIS)"""
    # === Constructeurs ===
    def __init__(self, nom_typeExport:str, codeExport:str, repertoire_input:str, nom_onglet_input:str, nbLignes_avantET_input:str, repertoire_modele:str, nom_fichier_modele:str, repertoire_output:str, nom_fichier_output:str):
        # Type d'export
        self._nom_typeExport = nom_typeExport
        self._codeExport = codeExport

        # Informations input données
        self._input = InfosExportsIRIS(
            repertoire=repertoire_input,
            nom_fichier=None,
            chemin_fichier=None,
            nom_onglet=nom_onglet_input,
            nbLignes_avantET=nbLignes_avantET_input
            )

        # Informations sur le modèle Excel à employer pour remplir l'output
        self._modele = InfosExportsIRIS(
            repertoire=repertoire_modele,
            nom_fichier=nom_fichier_modele,
            chemin_fichier=os.path.join(repertoire_modele, nom_fichier_modele),
            nom_onglet=nom_typeExport,
            nbLignes_avantET=None
            )

        # Informations output
        self._output = InfosExportsIRIS(
            repertoire=repertoire_output,
            nom_fichier=nom_fichier_output,
            chemin_fichier=None,
            nom_onglet=nom_typeExport,
            nbLignes_avantET=None
            )

    # === Affichage ===
    def __str__(self):
        def afficher_infos(nom_section, obj):
            lignes = [f"  {nom_section} :"]
            for champ, valeur in vars(obj).items():
                if valeur is not None:
                    # Appliquer gris clair uniquement sur les détails
                    lignes.append(f"    {Style.DIM}{Fore.LIGHTWHITE_EX}{champ:<17}: {valeur}{Style.RESET_ALL}")
            return "\n".join(lignes)

        type_export_str = f"  Type export : {self._nom_typeExport} ({self._codeExport})"

        return (
            f"PropExportIRIS\n"
            f"{type_export_str}\n"
            f"{afficher_infos('Input', self._input)}\n"
            f"{afficher_infos('Modèle', self._modele)}\n"
            f"{afficher_infos('Output', self._output)}"
        )

class TravauxFichiersIRIS:
    "C'est la classe qui contient l'environnement pour bosser sur des fichiers Exports IRIS"
    def __init__(self, prop:PropExportIRIS, chemins_fichiersInput:str|tuple[str, ...]=None):
        
        # Type d'export
        self._nom_typeExport = prop._nom_typeExport
        self._codeExport = prop._codeExport

        # Informations génériques sur les exports, modèles et output (dépend du type d'export)
        self._input = prop._input  # Informations input données
        self._modele = prop._modele  # Informations sur le modèle Excel à employer pour remplir l'output
        self._output = prop._output  # Informations output

        self._df_tableau = None
        self._df_chemins = None

        # S'il n'y a pas de chemin_fichiersInput de donné, c'est qu'il faut les sélectionner manuellement
        if not chemins_fichiersInput:
            self.choisirFichiers_filedialog()
        else:
            self._chemins_fichiersInput = chemins_fichiersInput

        # self._chemins_fichiersInput doit être un tuple de strings. Si c'est un string c'est qu'un seul fichier a été donné. Alors on convertit en tuple
        if isinstance(self._chemins_fichiersInput, str):
            self._chemins_fichiersInput = (chemins_fichiersInput,)

    @classmethod
    def avecLecture(cls, propExportIRIS:PropExportIRIS, chemins_fichiersInput:str|tuple[str, ...]=None):
        """On lit le/les extract IRIS et on stocke dans self._df_tableau"""
        instance = cls(prop=propExportIRIS, chemins_fichiersInput=chemins_fichiersInput)
        instance.lire_extractIRIS()
        return instance

    @classmethod
    def avecEcritureOutputDefaut(cls, propExportIRIS:PropExportIRIS, chemins_fichiersInput:str|tuple[str, ...]=None):
        
        instance = cls.avecLecture(propExportIRIS, chemins_fichiersInput)

        # Création des paramètres pour l'ouverture du modèle
        chemin_fichier = instance._modele.chemin_fichier
        #nom_onglet = instance._output.nom_onglet

        # Création des paramètres pour l'output
        chemin_fichier_output = os.path.join(instance._output.repertoire, instance._output.nom_fichier[:-5] + "-" + date.today().strftime("%Y.%m.%d") + ".xlsx") #ou f"{datetime.now():%Y.%m.%d}")

        # On ouvre le modèle et tous ses tableaux structurés
        #fe_modele = FichierExcel.depuis_fichier(chemin_fichier=chemin_fichier)
        fe_modele = FichierExcel.depuis_modele(chemin_modele=chemin_fichier, chemin_fichier_sauv=chemin_fichier_output)

        # On copie le DataFrame avec les nouvelles données dans le modèle
        fe_modele._tableaux[instance._nom_typeExport].ecrit_dataFrame_dans_tableauStructure(instance._df_tableau, supprimeDonneesEtRemplace=True)
        
        # On écrit les références des fichiers copiés dans le tableau structuré "Imports"
        instance._df_chemins = pd.DataFrame(instance._chemins_fichiersInput, columns=['Chemin fichier'])
        fe_modele._tableaux["Imports"].ecrit_dataFrame_dans_tableauStructure(df=instance._df_chemins, supprimeDonneesEtRemplace=True)
        
        #On enregistre et on ferme (par précaution car copieformat xlwings sauvegarde)
        fe_modele.save()    
        fe_modele.close()

        return instance
 
     # === Méthodes ===
    def choisirFichiers_filedialog(self):
        # Lister/sélectionner les documents à concaténer
        cheminsExcel = filedialog.askopenfilenames(title="Sélectionner les fichiers " + self._nom_typeExport + " (" + self._codeExport + ") Excel à concaténer", filetype=[("Fichiers Excel", "*.xlsx")], initialdir=self._input.repertoire)
        
        # Gestion du cas où il y a non-sélection de fichiers
        if not cheminsExcel:
            log_erreur("click sur cancel du filedialog → Pas de chemins de fichier")
        self._chemins_fichiersInput = cheminsExcel
    
    def lire_extractIRIS(self):
        """
        Crée le DataFrame pour l'export IRIS. On le stocke dans self.__df_tableau
        On selectionne la bonne methode en fonction du type d'export

        :Example:
        >>> self.lire_extractIRIS()


        .. seealso:: Rien du tout.
        .. warning:: Rien du tout.
        .. note:: Rien du tout.
        .. todo:: Rien du tout.
        """

        # Initialisation : on crée un DataFrame vide pour recevoir (peut-être) des infos que l'on traitera et qui nécessitera d'adjoindre des colonnes à self.__df_tableau
        df_colonnes_sup = None

        # On parcourt le tuple des fichiers à lire
        df_list = [] # Liste des DataFrame qui contiendra chaque fichier Excel séparément
        taille_totale = sum(os.path.getsize(fichier) for fichier in self._chemins_fichiersInput) # Calcul taille totale pour barre de progression

        with tqdm(total=taille_totale, unit='o', unit_scale=True, desc=Fore.CYAN+"Lecture des fichiers Excel" + Style.RESET_ALL) as pbar:
            #for i, ifichier in enumerate((os.path.basename(chemin) for chemin in self._chemins_fichiersInput), 1):
            for i, chemin in enumerate(self._chemins_fichiersInput, 1):
                # Données pour tqdm
                fichier = os.path.basename(chemin)
                taille = os.path.getsize(chemin)  
                pbar.set_postfix(file=fichier, progress=f"{i}/{len(self._chemins_fichiersInput)}")  # Affichage dynamique dans la barre
                
                df = pd.read_excel(chemin, skiprows=self._input.nbLignes_avantET)
                df_list.append(df)  # On ajoute le DataFrame à notre liste de DataFrame
                
                # Mise à jour de la barre avec la taille du fichier
                pbar.update(taille)

        # Concaténation finale (note : toute la fin de la méthode se fait quasi-instantanément)
        self._df_tableau = pd.concat(df_list, ignore_index=True) 

        # Selon le type d'export à traiter, on va faire des traitements spécifiques (extraction d'info des colonnes référence formation ou n° Iris)
        match self._codeExport:
            # Cas Sessions ou Inscriptions ou Ventes (sensiblement comme 'Inscription R04500' mais groupé par Client (pas de détail de chaque stagiaire))
            case "R04110" | "R04500" | "R04301":
                # On extrait / retravaille les informations de la colonne 'N° Session'
                df_colonnes_sup = self._df_tableau['N° Session'].apply(self.extraire_infos_numSessionIRIS)
                # A cause de certains éléments None ou NaN, Pandas type la colonne en float. Je la retype en Int64 qui permet de stocker des NaN avec des entiers, contrairement au type int standard.
                df_colonnes_sup['Année'] = df_colonnes_sup['Année'].astype('Int64')
                #print(df_colonnes_sup)

            # Cas Formations
            case "R0304":
                # On extrait / retravaille les informations de la colonne 'Référence'
                df_colonnes_sup = self._df_tableau['Référence'].apply(self.extraire_infos_referenceFormationIRIS)
                # A cause de certains éléments None ou NaN, Pandas type la colonne en float. Je la retype en Int64 qui permet de stocker des NaN avec des entiers, contrairement au type int standard.
                df_colonnes_sup['Année'] = df_colonnes_sup['Année'].astype('Int64')
                #print(df_colonnes_sup)

        # Ajouter les colonnes supplémentaires au DataFrame principal ssi le DataFrame df_colonnes_sup exite
        if df_colonnes_sup is not None:
            self._df_tableau = pd.concat([self._df_tableau, df_colonnes_sup], axis=1)
    
    def extraire_infos_numSessionIRIS(self, reference:str):
        """
        Fonction pour extraire les colonnes à partir de la colonne 'Référence'. Je dois faire une fonction interne car j'emploie Split qui ne s'applique que sur des string. Je dois donc faire appel à cette fonction ligne par ligne et donc créer une fonction que j'appelle par DataFrame[colonne].apply().

        Exemples de cas à traiter  :
        #S-04934-FI1516-1512-GI_VBE_GBO
        #S-05251-FI1516-1510-AMS-LCH-CLE
        #S-05246-F1516-1510-OPE-HGR-NNO
        #S-04178-FC15-604-SES-CCO

        ANCIENNE METHODE DE TRAITEMENT
        self.__df_tableau['Numéro IRIS'] = self.__df_tableau['N° Session'].astype(str).str[2:7]
        self.__df_tableau['Type formation'] = self.__df_tableau['N° Session'].astype(str).str[8:10]
        self.__df_tableau['Trigramme AF'] = self.__df_tableau['N° Session'].astype(str).str[-3:] #tout sauf 3 derniers caract
        self.__df_tableau['Trigramme RP'] = self.__df_tableau['N° Session'].astype(str).str[-7:-4] #De -7 à -4
        self.__df_tableau['Trigramme formation'] = self.__df_tableau['N° Session'].astype(str).str[-11:-8]        
        """
        # Vérifier que la référence est une chaîne de caractères
        if not isinstance(reference, str):
            #print("Problème : la référence n'est pas une instance : ")
            #print(reference)
            return pd.Series({
                'Trigramme formation': None,
                'Code IRIS': None,
                'Type de formation': None,
                'Année': None,
                'Trigramme RP': None,
                'Trigramme AF': None,
                '3ème élément de la référence': None
            })

        blocs = reference.split('-')

        # Sécurité : vérifier qu'on n'a pas plus de 8 blocs
        if len(blocs) > 8:
            print("Problème : il y a plus de 8 blocs : ")
            print(reference)
            return pd.Series({
                'Trigramme formation': None,
                'Code IRIS': None,
                'Type de formation': None,
                'Année': None,
                'Trigramme RP': None,
                'Trigramme AF': None,
                '3ème élément de la référence': None
            })

        # Traitement commun aux cas 6, 7 et 8 blocs
        # le blocs[0] c'est "S" ça sert à rien
        code_IRIS = blocs[1] if len(blocs) >= 2 else None
        #print(reference)
        if len(blocs) > 2:
            type_formation = blocs[2][:2]
            annee_match = re.search(r'\d+', blocs[2][2:])
            annee = int(annee_match.group()) if annee_match else None
        else :
            type_formation = None
            annee = None

        if len(blocs) > 5:
            trigramme_AF = blocs[-1].rstrip('_')
            trigramme_RP = blocs[-2].rstrip('_')
            trigramme = blocs[-3].rstrip('_')
        else:
            trigramme_RP = None
            trigramme_AF = None
            trigramme = None

        # 3e élément uniquement si on a 7 ou 8 blocs
        troisieme_bloc = '-'.join(blocs[3:-3]) if len(blocs) in [7, 8] else None
        
        return pd.Series({
            'Trigramme formation': trigramme,
            'Code IRIS': code_IRIS,
            'Type de formation': type_formation,
            'Année': annee,
            'Trigramme RP': trigramme_RP,
            'Trigramme AF': trigramme_AF,
            '3ème élément de la référence': troisieme_bloc
        })

    def extraire_infos_referenceFormationIRIS(self, reference):
        """
        Fonction pour extraire les colonnes à partir de la colonne 'Référence'. Je dois faire une fonction interne car j'emploie Split qui ne s'applique que sur des string. Je dois donc faire appel à cette fonction ligne par ligne et donc créer une fonction que j'appelle par DataFrame[colonne].apply().
        """
        
        # Vérifier que la référence est une chaîne de caractères
        if not isinstance(reference, str):
            #print("Problème : la référence n'est pas une instance : ")
            #print(reference)
            return pd.Series({
                'Trigramme formation': None,
                'Type de formation': None,
                'Année': None,
                'Unité de formation': None,
                '3ème élément de la référence': None
            })

        blocs = reference.split('-')

        # Sécurité : vérifier qu'on a au moins 4 blocs
        if len(blocs) < 3:
            print("Problème : il y a moins de 3 blocs : ")
            print(reference)
            return pd.Series({
                'Trigramme formation': None,
                'Type de formation': None,
                'Année': None,
                'Unité de formation': None,
                '3ème élément de la référence': None
            })

        # Traitement commun aux cas 3 et 4+ blocs
        type_formation = blocs[0][:2]
        annee_match = re.search(r'\d+', blocs[0][2:])
        annee = int(annee_match.group()) if annee_match else None

        trigramme = blocs[1].rstrip('_')
        unite_formation = blocs[-1].rstrip('_')

        # 3e élément uniquement si on a 4 blocs ou plus
        troisieme_bloc = '-'.join(blocs[2:-1]) if len(blocs) > 3 else None
        
        return pd.Series({
            'Trigramme formation': trigramme,
            'Type de formation': type_formation,
            'Année': annee,
            'Unité de formation': unite_formation,
            '3ème élément de la référence': troisieme_bloc
        })
   
   
    # === Affichage ===
    def __str__(self):
        if self._chemins_fichiersInput:
            # Pour aff_repertoire et chemin 
            aff_cheminsFichiers = ""
            for cfichier in self._chemins_fichiersInput:
                aff_cheminsFichiers = aff_cheminsFichiers + "\n    " + cfichier
        else:
            aff_cheminsFichiers = "Aucun fichier spécifié"

        # Pour aff_df
        from io import StringIO
        buffer = StringIO()
        if self._df_tableau is not None:
            print(self._df_tableau, file=buffer)
            aff_df = buffer.getvalue()
        else:
            aff_df = "Non défini"

        return (
            f"TravauxFichiersIRIS\n"
            f"  Chemins fichiers à exploiter : {aff_cheminsFichiers}\n"
            f"  df_tableau :\n{aff_df}"
            )

class BilanSessionV3:
    """
    C'est la classe qui contient tous les éléments de ma formation pour mon bilan de session V3

    # TODO j'en suis là
    # Todo Word
    # Dans le modèle Word : gérer le lien vers la GED 
    # Exploiter EvalStat
    # Il y a des trous dans la raquette dans le word de sortie (checkboxes)
    # coller des images depuis Excel
    # ? Exploiter export formation plutôt que export sessions pour les valeurs par défaut nmin/max...

    """
    def __init__(self, codeFormation:str, annee:int, periode:str) -> None:
        self._codeFormation:str = codeFormation
        self._annee:int = annee
        self._periode:str = periode
        self._lieuPrincipal:str = "INSTN Marcoule"

        self._sessions_nom_typeExport = sessions._nom_typeExport #  Provient de la valeur globale sessions
        self._sessions_codeExport = sessions._codeExport #  Provient de la valeur globale sessions
        self._sessions_repertoire = sessions._output.repertoire #  Provient de la valeur globale sessions

        self._fe_sessions:FichierExcel = None # Fichier Excel qui contient les extracts IRIS Sessions (ou a minima celles de la période en cours)

        self._titreFormation:str = None
        self._periodeSessionsEvaluees:str = f"{self._periode} {self._annee}"
        self._nbSessionsEvaluees:int = None
        self._numerosSessions:str = None
        self._nbApprenants:int = None
        self._rp:str = None
        self._af:str = None

        #####
        # Exploitation de l'extract IRIS sessions
        #####

        # J'ouvre un export session de IRIS et load tous ses tableaux structurés dans des DataFrame (inclus dans un FichierExcel)
        chemin_fichier_session = filedialog.askopenfilename(title="Sélectionner l'export " + self._sessions_nom_typeExport + " (" + self._sessions_codeExport + ") Excel à employer", filetype=[("Fichiers Excel", "*.xlsx")], initialdir=self._sessions_repertoire)
        if not chemin_fichier_session:
            vlog.log_erreur("click sur cancel du filedialog → Pas de chemin de fichier session")
        self._fe_sessions = FichierExcel.depuis_fichier(chemin_fichier_session)
        #print(self._fe_session)
        self._df_sessions = self._fe_sessions._tableaux[self._sessions_nom_typeExport]._df  # Création d'un alias
        self._df_sessions['Trigramme formation'] = self._df_sessions['Trigramme formation'].astype(str)
        #print(self._df_sessions.columns.to_list())
        #print(self._df_sessions)
        #for col in self._df_sessions.columns:
        #    print(repr(col))
        #print(self._df_sessions['Trigramme formation'].dtype)
        
        # Pour initialiser les valeurs communes, déjà on filtre le Dataframe principal avec le code formation
        if self._periode == "1er semestre":
            date_debut = pd.Timestamp(f'{annee}-01-01')
            date_fin = pd.Timestamp(f'{annee}-06-30')
        elif self._periode == "2nd semestre":
            date_debut = pd.Timestamp(f'{annee}-07-01')
            date_fin = pd.Timestamp(f'{annee}-12-31')
        elif self._periode == "Année":
            date_debut = pd.Timestamp(f'{annee}-01-01')
            date_fin = pd.Timestamp(f'{annee}-12-31')
        else:
            # Cas par défaut : on ne filtre pas sur la date
            date_debut = None
            date_fin = None

        # Application du pré-filtre avec les 3 critères trigramme, statut session et période
        df_sessions_filtre = self._df_sessions[
            (self._df_sessions['Trigramme formation'] == str(self._codeFormation)) &
            (self._df_sessions['Année début ses.'] == self._annee) &
            (self._df_sessions['Statut Session'] != "Annulée")
        ]
        #print(df_sessions_filtre)
        if date_debut is not None and date_fin is not None:
            df_sessions_filtre = df_sessions_filtre[
                (df_sessions_filtre['Date début ses.'] >= date_debut) &
                (df_sessions_filtre['Date début ses.'] <= date_fin)
            ]   
        #print(df_sessions_filtre)
        # Afficher les sessions et dates et statuts 
        #print(df_sessions_filtre[['Code IRIS', 'Trigramme RP', 'Trigramme AF', 'Date début ses.', 'Date fin ses.', 'Statut \nSession', 'N° Session']].to_string(index=False))
        print(f"Liste des sessions {self._codeFormation} dans {os.path.basename(chemin_fichier_session)} - {self._periode} {self._annee}")
        print(tabulate(
            df_sessions_filtre[['Code IRIS', 'Trigramme RP', 'Trigramme AF', 'Date début ses.', 'Date fin ses.', 'Statut Session', 'N° Session']], 
            headers='keys', 
            tablefmt='pretty', 
            showindex=False
        ))
        
        #On demande à l'utilisateur les sessions qu'il veut exclure
        exclusionSessions = self.demander_entiers()
        if exclusionSessions:  # si la liste n'est pas vide
            df_sessions_filtre = df_sessions_filtre[~df_sessions_filtre['Code IRIS'].isin(exclusionSessions)]
        else:
            # la liste est vide, on ne filtre rien, on garde tout
            pass
 
        print(tabulate(
            df_sessions_filtre[['Code IRIS', 'Trigramme RP', 'Trigramme AF', 'Date début ses.', 'Date fin ses.', 'Statut Session', 'N° Session']], 
            headers='keys', 
            tablefmt='pretty', 
            showindex=False
        ))


        # On définit les mergeField de Word issus de l'exrtract IRIS sessions
        self._titreFormation = df_sessions_filtre[["Session"]].iloc[-1]
        self._nbSessionsEvaluees = len(df_sessions_filtre)
        self._numerosSessions = ", ".join(df_sessions_filtre["N° Session"].astype(str))
        self._nbApprenants = df_sessions_filtre["Nb. Nommés"].sum()
        self._rp = ", ".join(df_sessions_filtre["Trigramme RP"].astype(str))
        self._af = ", ".join(df_sessions_filtre["Trigramme AF"].astype(str))


        #####
        # Exploitation des EvalStat
        #####

        # A partir de la liste df_sessions_filtre['Code IRIS'], on regarde les CSV qui ne sont pas dans le fichier Excel global
        # On ouvre le fichier Excel global des évaluation de la formation
        es = Traiter_evalStat.depuis_fe_evaluations_formation(self._codeFormation)

        # On isole depuis ce fichier les CSV manquants
        code_session_absents = List(set(df_sessions_filtre['Code IRIS']) - set(es._df_formation_stagiaires['Code IRIS']))

        chemins_csv = {}
        for code_IRIS in code_session_absents:
            chemins_csv[code_IRIS] = chemin_vers_unc(es.filedialog_csv(code_IRIS=code_IRIS, trigramme=self._codeFormation))

        

        #TODO : Il serait bien que je gère les sessions sans CSV. Pour l'instant, j'exclue
        # On en fait un tupe en excluant les None
        tuple_csv_stagiaires = tuple(val for val in chemins_csv.values() if val is not None)

        if tuple_csv_stagiaires:
            es = Traiter_evalStat.depuis_tuple_csv_stagiaires(tuple_csv_stagiaires=tuple_csv_stagiaires, fe_sessions=self._fe_sessions)
            

        # On ouvre le fichier Excel EvalStat et on le filtre sur les sessions qui nous intéressent


    @staticmethod
    def demander_entiers(message="Pour exclure des sessions : entrez un ou plusieurs code IRIS (numéro à 5 chiffres) séparés par des espaces (ou rien pour passer) : ") -> List[str]:
        while True:
            entree = input(message).strip()
            if not entree:
                # Pas de saisie => retourner liste vide
                return []
            
            # Séparer les valeurs (espaces ou virgules)
            valeurs = [v.strip() for v in entree.replace(',', ' ').split()]
            return valeurs
            
            # Vérifier que toutes les valeurs sont des entiers
            #try:
            #    entiers = [int(v) for v in valeurs]
            #    return entiers
            #except ValueError:
            #    print("Erreur : veuillez entrer uniquement des nombres entiers, séparés par des espaces ou des virgules.")

class BilanFormation:
    """
    C'est la classe qui contient tous les éléments de ma formation pour mon bilan

    # TODO j'en suis là
    # Todo Word
    # Dans le modèle Word : gérer le lien vers la GED 
    # Exploiter EvalStat
    # Il y a des trous dans la raquette dans le word de sortie (checkboxes)
    # coller des images depuis Excel
    # ? Exploiter export formation plutôt que export sessions pour les valeurs par défaut nmin/max...
    """
    def __init__(self, codeFormation:str, annee:int):
        
        self._codeFormation = codeFormation
        self._annee = annee

        self._sessions_nom_typeExport = sessions._nom_typeExport #  Provient de la valeur globale sessions
        self._sessions_codeExport = sessions._codeExport #  Provient de la valeur globale sessions
        self._sessions_repertoire = sessions._output.repertoire #  Provient de la valeur globale sessions

        self._chemin_specsPedagogiques = None

        self._chemin_fdc = None

        self._repertoire_fdc_defaut = rep_fdc_defaut
        self._repertoire_specsPedagogiques_defaut = rep_specsPedagogiques_defaut
        
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
        self._dateSpecs = time.localtime(os.path.getmtime(self._chemin_specsPedagogiques))
        self._sDateSpecs = f"{self._dateSpecs[2]:02}/{self._dateSpecs[1]:02}/{self._dateSpecs[0]:04}"
        #print(self._sDateSpecs)



        #####
        # Exploitation de la fiche de coûts
        #####
        self._chemin_fdc = filedialog.askopenfilename(title="Sélectionner la dernière fiche de coûts", filetype=[("Fichiers Excel", "*.xlsx")], initialdir=optimiseCheminRepertoire(self._repertoire_fdc_defaut.replace("XXX", self._codeFormation)))
        if not chemin_fichier_session:
            log_erreur("click sur cancel du filedialog → Pas de chemin de fiche de coûts")
        self._dateFdC = time.localtime(os.path.getmtime(self._chemin_fdc))
        self._sDateFdC = f"{self._dateFdC[2]:02}/{self._dateFdC[1]:02}/{self._dateFdC[0]:04}"
        
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
    def mergeBilan(self, chemin_word_bilan_input, chemin_word_bilan_output):
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
        document = MailMerge(chemin_word_bilan_input)
        #print(document.get_merge_fields())

        document.merge(
            annee='{:%Y}'.format(date.today()),
            codeFormation=self._codeFormation,
            titreFormation=self._titreFormation,
            
            lienGED = r'file:///\\\\instnt\\PARTAGE\\FORMATIONS_C\\ACI\\',
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
        
        document.write(chemin_word_bilan_output)

class Traiter_evalStat:
    # Obtenir :
    #   - Taux de recommandation en 2025 :	100%
    #   - Satisfaction sur l’année 2025 :	4,75 sur 5
    #   - Taux de retour (facultatif)

    # Enlever ligne total avant traitement
    # Remise ligne total après fin traitement

    # Aller chercher CSV (sélectionner le fichier à partir du chemin standard)
    # chemin standard vers CSV = "\\Instnt\partage\FORMATIONS_C\"&codeFormation&"\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\CSV\"

    # Importer le CSV dans le modèle Excel (onglet "CSV")
    # Ecrire nb d'inscrits (vient de sessions) + code formation et code IRIS + changer chemin CSV → Voir si je ne peux pas urtiliser des plages nommées
    # coller infos dans tableau travaillé
    #   - depuis colonne H jusqu'à la fin (AG) on copie et on va transposer
    #   - pour toutes les 2 colonnes ça fait critère puis commentaire
    #   - sauf pour "Recommanderiez-vous cette formation ?" (AD) + "Commentaires, remarques, suggestions " (AH)
    #   - transformer valeurs de Recommanderiez-vous cette formation ? et Avez-vous d'autres besoins de formation ? en oui = 5 et non = 0
    # Actualiser les TCD
    # Sauvegarder au bon endroit
    # TODO : pour l'instant à cause de pandas, je casse les segments

    # Possibilité d'importer plusieurs CSV pour une même année
    # Attention : impact sur nombre total de stagiaire (sommer dans sessions)
    
    def __init__(self) -> None:
        
        self._chemin_csv_stagiaires:str|None = None  # Fichier csv EvalStat stagiaire individuel
        self._chemin_excel_stagiaires_output:str|None = None  # Fichier xlsx EvalStat stagiaire individuel qu'on va créer à partir du CSV
        self._fe_stagiaires:FichierExcel|None = None  # Objet contenant les données EvalStat stagiaire individuel

        self._chemin_modeleExcel_stagiaires:str|None = None  # Modèle Excel dans lequel importer le CSV
        
        self._chemin_excel_sessions:str|None = None  # Chemin du fichier Excel qui contient les extracts IRIS Sessions (ou a minima celles de la période en cours) dans lequel on a les informations des sessions (permet de compélter les CSV)
        self._fe_sessions:FichierExcel|None = None  # Fichier Excel qui contient les extracts IRIS Sessions (ou a minima celles de la période en cours)

        self._trigrammeFormation:str|None = None
        self._codeIRIS:int|None = None

        self._chemin_excel_evaluations_formation:Optional[str] = None  # Chemin du fichier Excel qui contient tous les CSV d'évaluation d'une formation
        self._fe_evaluations_formation:Optional[FichierExcel] = None # Fichier Excel qui contient tous les CSV d'évaluation d'une formation
        self._df_formation_stagiaires:Optional[pd.DataFrame] = None # DataFrame de self._fe_evaluations_formation (Alias)

        self._chemins_csv_traites:List[str] = []
        self._chemins_csv_exclus:List[str] = []

        # Valeurs actuellement écrites en dur
        self._chemin_excel_evaluations_defaut = r'\\instnt\partage\FORMATIONS_C\###\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\Evaluation-Stagiaires-Global-###.xlsx'
        self._chemin_modeleExcel_stagiaires = r"C:\Users\vt238770\Documents\_CEA\Prog\Modèles\Evaluation-Stagiaires-Modèle.xlsx"
        self._chemin_excel_sessions = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts complets\R04110_Sessions-COMPLET-2025.08.24.xlsx" 
        #instance._chemin_excel_stagiaires_output = r"C:\Users\vt238770\Documents\_CEA\Prog\Modèles\Evaluation-Stagiaires-testOut.xlsx"

        # === Colonnes du CSV selon traitement à avoir ===
        # Colonnes descriptives à recopier
        self._colonnes_csv_fixes = [
            "Chemin fichier CSV", "Prénom", "Nom", "Entreprise", "Code session"]

        # Colonnes avec note/commentaire en binôme
        self._colonnes_csv_avec_commentaires = [
            "Accueil, organisation et qualité des informations délivrées",
            "Conseils et orientation avant l'inscription",
            "Informations après l'inscription",
            "Accueil à l'arrivée sur site",
            "Prise en compte de vos besoins et attentes",
            "Qualité des animations",
            "Logique d'enchainement des interventions",
            "Qualité des supports de cours utilisés",
            "Qualité des moyens pédagogique",
            "Accès aux outils digitaux",
            "Satisfaction globale",
            "Avez-vous d'autres besoins de formation ?"]

        # Colonnes à valeur texte seule
        self._colonnes_csv_commentaires_seuls = [
            "Comment avez-vous connu cette formation ?",
            "Commentaires, remarques, suggestions"]

        # Colonnes note seule (il se trouve que je vais aussi devoir convertir le booléen)
        self._colonnes_csv_bool = [
            "Recommanderiez-vous cette formation ?"]
        
        
        # === Colonnes de l'extract IRIS Sessions à récupérer ===
        self._colonnes_sessions = [
            "N° Session",
            "Formation",
            "Trigramme formation",
            "Code IRIS",
            "Date début ses.",
            "Année début ses.",
            "Type de formation",
            "Trigramme RP",
            "Trigramme AF",
            "Nb. Présents"]

     
    @classmethod
    def depuis_fe_evaluations_formation(cls, trigramme:str) -> Traiter_evalStat:
        """
        A partir d'un trigramme de foramtion, on ouvre et on charge le fichier excel qui concatène tous les CSV d'une formation 
        """
        # On crée l'instance et on complète les infos avec les valeurs facultatives
        instance = Traiter_evalStat()
        # On définit le chemin vers les évaluations de la formation (le fichier qui va concaténer toutes les évaluation d'une formation)
        instance._chemin_excel_evaluations_formation = instance._chemin_excel_evaluations_defaut.replace("###", trigramme)

        instance.ouvrir_fe_evaluations_formation()

        return instance
        
    @classmethod
    def depuis_chemin_csv_stagiaires(cls, chemin_csv_stagiaires:str, fe_sessions:FichierExcel=None, ouvrirDossier:bool=False, remplace_df:bool=False) -> Traiter_evalStat:
        timer.debut(f"Traitement du CSV {os.path.basename(chemin_csv_stagiaires)}")
        # On créée l'instance
        instance = Traiter_evalStat()
        instance._chemin_csv_stagiaires = chemin_csv_stagiaires
        if fe_sessions is not None:
            instance._fe_sessions = fe_sessions
            # Quand je ferai la jointure plus tard sur "Code IRIS", il faudra que ce soit avec des strings
            instance._fe_sessions._tableaux["Sessions"]._df["Code IRIS"] = instance._fe_sessions._tableaux["Sessions"]._df["Code IRIS"].astype(str)

        # On récupère le trigramme de la formation depuis le chemin du CSV
        instance._trigrammeFormation = instance.recupere_trig_formation_depuis_chemin(chemin_csv_stagiaires)
        
        # Pour le nom de l'Excel output : on reprend le nom du csv et on remplace par xlsx
        instance._chemin_excel_stagiaires_output = os.path.join(os.path.dirname(instance._chemin_csv_stagiaires), os.path.basename(instance._chemin_csv_stagiaires).replace(".csv", ".xlsx"))

        # On génère le fichier Excel du CSV à partir du modèle
        instance._construit_FichierExcel_depuis_CSV(fe_sessions = fe_sessions, remplace_df=remplace_df)

        # Ouverture du dossier à la fin
        if ouvrirDossier:
            ouvrir_dossier(os.path.dirname(instance._chemin_excel_stagiaires_output))

        # On ajoute le chemin au tuple des éléments traités
        instance._chemins_csv_traites.append(chemin_csv_stagiaires)
        
        timer.fin()
        return instance        

    @classmethod
    def depuis_tuple_csv_stagiaires(cls, tuple_csv_stagiaires:Tuple(str), chemin_excel_evaluations_defaut:str=None, chemin_modeleExcel_stagiaires:str=None, chemin_excel_sessions:str=None, fe_sessions:FichierExcel=None, ouvrirDossier:bool=False) -> Traiter_evalStat:
        """
        A partir d'un tuple de chemins de CSV stagiaire (il peut il y avoir plusieurs trigrammes de formations différents)
        Permet de générer :
           - le fichier excel stagiaires de chaque session (via le CSV)
           - le fichier excel stagiaires de chaque formation (celui qui concatène tous les CSV d'une session) [Il est créé ou on l'append avec les nouvelles valeurs]
        
        On exclue du traitement les chemin_csv_session qui sont déjà dans le FE formation (on considère que le CSV a déjà été traité)

        """

        
        # On crée l'instance et on complète les infos avec les valeurs facultatives
        instance = Traiter_evalStat()

        if chemin_excel_evaluations_defaut is not None:
            instance._chemin_excel_evaluations_defaut = chemin_excel_evaluations_defaut
        if chemin_modeleExcel_stagiaires is not None:
            instance._chemin_modeleExcel_stagiaires = chemin_modeleExcel_stagiaires
        if chemin_excel_sessions is not None:
            instance._chemin_excel_sessions = chemin_excel_sessions
        if fe_sessions is not None:
            instance._fe_sessions = fe_sessions
        
        df_formation_csv = None
        df_formation_stagiaires = None
        #dico_sessionsDejaTraitees = None
        #timer = Timer()

        # On récupère les données de Sessions (on en aura besoin plus tard)
        if instance._fe_sessions is None:
            timer.debut("Lecture fichier session")
            instance._fe_sessions = FichierExcel.depuis_fichier(instance._chemin_excel_sessions)
            timer.fin()
        

        # On convertit le tuple de strings en dictionnaire avec les trigrammes formation en clef
        dico_chemins_csv_session = defaultdict(list)  #Dictionnaire spécial : lorsqu’on accède à une clé qui n’existe pas encore, il va automatiquement créer une nouvelle entrée avec une valeur par défaut, ici une liste vide (list())
        for chemin in tuple_csv_stagiaires:
            trigramme = instance.recupere_trig_formation_depuis_chemin(chemin)
            dico_chemins_csv_session[trigramme].append(chemin)
        dico_chemins_csv_session = dict(dico_chemins_csv_session)  # Optionnel : conversion en dict normal


        # Pour chaque trigramme on va traiter chaque session et soit créer soit append le fichier excel global de la formation
        for trigramme, chemins_csv_session in dico_chemins_csv_session.items(): #chemins est la liste des chemins des évaluations pour chaque sessions de ce trigramme formation
            print(f"\n\n{Style.BRIGHT}{Fore.RED}Gestion des formations {trigramme}")

            # On ouvre ou on créée (si inexistant) le fichier Excel qui concatène toutes les sessions d'une formation
            supprimeDonneesEtRemplace_evaluations_formation = instance.ouvrir_ou_creer_evaluationsFormation(trigramme)

            # Pour chaque chemin de session, on crée le fe_stagiaire dédié de la session et on ajoute les lignes de son dataframe au dataframe de fe_evaluations_formation
            for chemin_csv_session in chemins_csv_session:
                print(f"\n{Style.BRIGHT}{Fore.YELLOW}Gestion de la session {chemin_csv_session}")
                traiterCSV = True  # Par défaut, on traite le CSV

                # On vérifie que chemin_csv_session n'est pas déjà dans le fichier session pour savoir si on l'exclue du traitement
                if df_formation_stagiaires is not None:
                    if chemin_csv_session in df_formation_stagiaires["Chemin fichier CSV"].drop_duplicates().tolist():  
                        traiterCSV = False
                        instance._chemins_csv_exclus.append(chemin_csv_session)
                        print(f"Exclusion car csv déjà dans le fichier global : {chemin_csv_session}")
                        vlog.ajouter_message("Exclusion car csv déjà dans le fichier global", chemin_csv_session, style=["orange"])

                # Traitement du CSV
                if traiterCSV:  
                    #timer.debut("Traiter_evalStat.depuis_chemin_csv_stagiaires")
                    traite_csv_session = Traiter_evalStat.depuis_chemin_csv_stagiaires(chemin_csv_session, fe_sessions=instance._fe_sessions, ouvrirDossier=ouvrirDossier, remplace_df=True)

                    #timer.debut("Copie des Dataframe csv et stagiaires")
                    # #Si df_formation_csv est vide, il faut l'initialiser avec le premier df sinon on concatène
                    if df_formation_csv is None:
                        df_formation_csv = traite_csv_session._fe_stagiaires._tableaux["CSV_stagiaires"]._df.copy()
                        df_formation_stagiaires = traite_csv_session._fe_stagiaires._tableaux["Stagiaires"]._df.copy()
                    else:
                        df_formation_csv = pd.concat([df_formation_csv, traite_csv_session._fe_stagiaires._tableaux["CSV_stagiaires"]._df], ignore_index=True)
                        df_formation_stagiaires = pd.concat([df_formation_stagiaires, traite_csv_session._fe_stagiaires._tableaux["Stagiaires"]._df], ignore_index=True)
                    
                    # On ajoute le chemin au tuple des éléments traités
                    instance._chemins_csv_traites.append(chemin_csv_session)
                    vlog.ajouter_message("Fichiers traités", chemin_csv_session, style=["vert"])

            # On concatène, on sauve et on ferme le fe de tous les CSV de la formation
            if instance._chemins_csv_traites :
                print(f"\n\n{Style.BRIGHT}{Fore.RED}Écriture du fichier global des évaluations de la formation {trigramme}")
                timer.debut("Écriture, sauvegarde et fermeture")
                instance._fe_evaluations_formation._tableaux["CSV_stagiaires"].ecrit_dataFrame_dans_tableauStructure(df_formation_csv, supprimeDonneesEtRemplace = supprimeDonneesEtRemplace_evaluations_formation)
                instance._fe_evaluations_formation._tableaux["Stagiaires"].ecrit_dataFrame_dans_tableauStructure(df_formation_stagiaires, supprimeDonneesEtRemplace = supprimeDonneesEtRemplace_evaluations_formation)

                instance._fe_evaluations_formation.save()
                instance._fe_evaluations_formation.close()
                timer.fin()
            
                # Actualisation des TCD
                instance._fe_evaluations_formation.actualiser_TCD()

        return instance

    # ===  (getter / setter) ===
    @property
    def chemins_csv_traites(self):
        return self._chemins_csv_traites


    ######
    #  === Méthodes internes ===
    ######

    def _construit_FichierExcel_depuis_CSV(self, fe_sessions:FichierExcel=None, remplace_df:bool=False):
        ######
        # === Import et traitement du CSV d'evalStat ===
        ######
        # On récupère l'encodage et on importe le CSV dans un DataFrame
        #timer = Timer()
        codage_csv = trouve_encodage_csv(self._chemin_csv_stagiaires)
        #print(f"\nConstruction de l'Excel pour {os.path.basename(self._chemin_csv_stagiaires)} ; codage : {codage_csv}")
        #timer.debut("Import du CSV et traitement du DataFrame")
        df_csv_stagiaires = pd.read_csv(self._chemin_csv_stagiaires, sep=';', encoding=codage_csv)  # Ouverture du CSV et mise dans un DataFrame
        
        # Prise en compte qu'on a plusieurs formats de CSV : on doit traiter des colonnes en + ou - en conséquences
        if "Date de fin" in df_csv_stagiaires.columns:
            # Cas 1 (nouveau format de csv) : supprimer "Date de fin" → Test : r"P:\FORMATIONS_C\54C\P07-bilan-sessions-et-bilan-formation\rapports-sessions-evaluations\2023-06-S14317 UEM\S-14317-FC23-54C-VTE-LRA-Stagiaires.csv"
            df_csv_stagiaires = df_csv_stagiaires.drop(columns=["Date de fin"])
        else:
            # Cas 2 (ancien format de csv) : supprimer la 2e et 3e colonne (indices 1 et 2) → Test : r"P:\FORMATIONS_C\22B\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\S-17606 - 22B - 06-2025\S-17606-FC25-22B-VTE-CAR-Stagiaires.csv"
            df_csv_stagiaires = df_csv_stagiaires.drop(df_csv_stagiaires.columns[[1, 2]], axis=1)
        
        # On rajoute le chemin du CSV en première colonne
        df_csv_stagiaires.insert(0, "Chemin fichier CSV", self._chemin_csv_stagiaires)

        # Définition self._codeIRIS. Sinon Non existant, on récupère le numéro IRIS depuis le CSV (c'est la plus sur)
        if self._codeIRIS is None:
            match = re.search(r"\b\d{5}\b", self._chemin_csv_stagiaires)
            if match:
                self._codeIRIS = match.group(0)
            else:
                self.demander_code_iris()

        # Met à jour ou crée la colonne "Code session" avec self._codeIRIS
        df_csv_stagiaires["Code session"] = self._codeIRIS

        # Mise au format jj/mm/aaaa de la colonne "Date" (si elle existe)
        if "Date" in df_csv_stagiaires.columns:
            try:
                df_csv_stagiaires["Date"] = pd.to_datetime(df_csv_stagiaires["Date"], dayfirst=True, errors="coerce").dt.strftime("%d/%m/%Y")  # dayfirst=True indique que le premier nombre correspond au jour (format jj/mm/aaaa)
            except Exception as e:
                print(f"Erreur de conversion de la colonne Date : {e}")

        # A cause des espaces à la con qui trainent dans les noms des colonnes des CSV, je vais reload le dataframe depuis l'excel que je viens de créer car les colonnes du modèle sont bien nommées
        # Ainsi on sauve ici plutôt qu'à la fin et on reload le DataFrame
        self._fe_stagiaires = FichierExcel.depuis_modele(chemin_modele=self._chemin_modeleExcel_stagiaires, chemin_fichier_sauv=self._chemin_excel_stagiaires_output)
        if remplace_df:
            self._fe_stagiaires._tableaux["CSV_stagiaires"].remplace_df(df_csv_stagiaires)
        self._fe_stagiaires._tableaux["CSV_stagiaires"].ecrit_dataFrame_dans_tableauStructure(df_csv_stagiaires, supprimeDonneesEtRemplace=True)
        self._fe_stagiaires.save(self._chemin_excel_stagiaires_output)
        self._fe_stagiaires.close()
        self._fe_stagiaires = FichierExcel.depuis_fichier(self._chemin_excel_stagiaires_output)
        self._fe_stagiaires._tableaux["CSV_stagiaires"].charge_df()
        df_csv_stagiaires = self._fe_stagiaires._tableaux["CSV_stagiaires"]._df
        
        ######
        # === On crée la seconde partie du DataFrame qui sera dans l'onglet "Stagiaire" ===
        # On va découper le dataframe du CSV selon les différents critères et mettre dans un dataframe qu'on pourra exploiter par un TCD
        ######
        # Nouveau DataFrame à remplir
        df_long = []

        # Parcours des lignes
        for _, row in df_csv_stagiaires.iterrows():
            #print("Ligne en cours : ")
            #print(row)
            base = {col: row[col] for col in self._colonnes_csv_fixes}  # Création des colonnes qui seront répétées à chaque fois
            base["NOM Prénom"] = f"{str(row['Nom']).upper()} {row['Prénom']}".strip()  # Création du champ "NOM Prénom"

            # Cas 1 : colonnes avec note + commentaire associé
            for critere in self._colonnes_csv_avec_commentaires:
                if critere in row:
                    #print(f"'{critere}'")
                    if(critere == "Avez-vous d'autres besoins de formation ?"):  # Il faut changer le booléen en 0 ou 5
                        val = str(row[critere]).strip().lower()
                        row[critere] = 5 if val == "oui" else (0 if val == "non" else None)
                    commentaire_col = row.index[row.index.get_loc(critere) + 1]
                    if(pd.notna(row[critere]) or pd.notna(row.get(commentaire_col, None))) :
                        df_long.append({
                            **base,
                            "Critère": critere,
                            "Note": row[critere],
                            "Commentaires": row.get(commentaire_col, None)
                        })

            # Cas 2 : colonnes texte seules
            for critere in self._colonnes_csv_commentaires_seuls:
                if critere in row:
                    if(pd.notna(row[critere])) :
                        df_long.append({
                            **base,
                            "Critère": critere,
                            "Note": None,
                            "Commentaires": row[critere]
                        })

            # Cas 3 : Note seule + booléens convertis
            for critere in self._colonnes_csv_bool:
                if critere in row:
                    if(pd.notna(row[critere])) :
                        val = str(row[critere]).strip().lower()
                        note = 5 if val == "oui" else (0 if val == "non" else None)
                        df_long.append({
                            **base,
                            "Critère": critere,
                            "Note": note,
                            "Commentaires": None
                        })

        # Construction du DataFrame final
        df_stagiaires = pd.DataFrame(df_long)

        # Réorganise les colonnes pour placer "NOM Prénom" juste après "Nom"
        colonnes = list(df_stagiaires.columns)
        if "NOM Prénom" in colonnes and "Nom" in colonnes:
            colonnes.remove("NOM Prénom")
            index_nom = colonnes.index("Nom")
            colonnes.insert(index_nom + 1, "NOM Prénom")
            df_stagiaires = df_stagiaires[colonnes]

        # Supprime les colonnes "Prénom" et "Nom" devenues inutiles
        df_stagiaires.drop(columns=["Prénom", "Nom"], inplace=True)


        ######
        # === On fait le left join entre df_stagiaires et les données qui proviennent de l'extract IRIS Sessions ===
        ######
        # On récupère les données de Sessions
        if fe_sessions is None:
            timer.debut("Lecture fichier session")
            self._fe_sessions = FichierExcel.depuis_fichier(self._chemin_excel_sessions)
            self._fe_sessions._tableaux["Sessions"]._df["Code IRIS"] = self._fe_sessions._tableaux["Sessions"]._df["Code IRIS"].astype(str)
        else:
            self._fe_sessions = fe_sessions


        # Pour faire le merge, il faut que les colonnes soient de même type (là "Code session" est de type int64 et "Code IRIS" est de type object (souvent des chaînes de caractères)).
        # Comme je ne peux être sûr que tous les "Code IRIS" issu des CSV soient bien convertibles en int (c’est-à-dire pas de chaînes vides, NaN, ou autres caractères non numériques), alors je passe par des strings
        df_stagiaires["Code session"] = df_stagiaires["Code session"].astype(str)
        #self._fe_sessions._tableaux["Sessions"]._df["Code IRIS"] = self._fe_sessions._tableaux["Sessions"]._df["Code IRIS"].astype(str)
        
        # On récupère uniquement les colonnes souhaitées dans une vue pour faciliter le codage (c'est un alias)
        df_sessions_filtre = self._fe_sessions._tableaux["Sessions"]._df[self._colonnes_sessions]
        
        # On fait la jointure entre df_stagiaires et df_sessions_filtre
        #timer.debut("Création du DataFrame Stagiaires (jointure)")
        df_stagiaires = df_stagiaires.merge(
            df_sessions_filtre,
            left_on="Code session",
            right_on="Code IRIS",
            how="left"
        )

        # On réorganise les colonnes : d'abord celles de df_sessions puis celles de df_stagiaires
        colonnes_resultat = (
            df_sessions_filtre.columns.tolist() +  # colonnes de _df_sessions
            [col for col in df_stagiaires.columns if col not in df_sessions_filtre.columns]  # le reste (i.e. celles de df_stagiaires)
        )
        df_stagiaires = df_stagiaires[colonnes_resultat]
        #print(_df_stagiaires)
        
        # On vire "Code session" qui est redondante avec "Code IRIS"
        df_stagiaires.drop(columns=["Code session"], inplace=True)
        
        # On renomme les colonnes
        #df_sessions_filtre.rename(columns={"Date début ses.": "Date"}, inplace=True)
        #df_sessions_filtre.rename(columns={"Année début ses.": "Année"}, inplace=True)

        # On remplace le DataFrame existant par le nouveau
        if remplace_df:
            self._fe_stagiaires._tableaux["Stagiaires"].remplace_df(df_stagiaires)
        
        # On écrit et on sauve
        #timer.debut("On écrit le DataFrame, on met à jour les TCD et on sauve")
        self._fe_stagiaires._tableaux["Stagiaires"].ecrit_dataFrame_dans_tableauStructure(df_stagiaires, supprimeDonneesEtRemplace=True)
        self._fe_stagiaires.save()


        # Màj des TCD
        self._fe_stagiaires.actualiser_TCD()
        #timer.fin()

    def demander_code_iris(self):
        print("TODO : demander_code_iris appelé → Essayer de généraliser avec demander_code")
        def verifier_entree(*args):
            val = entry_code.get()
            bouton_valider.config(state="normal" if val.isdigit() and len(val) == 5 else "disabled")
            bouton_renommer.config(state="normal" if val.isdigit() and len(val) == 5 else "disabled")

        def valider():
            nonlocal code_iris
            self._codeIRIS = int(entry_code.get())
            fenetre.destroy()

        def valider_et_renommer():
            nonlocal code_iris
            self._codeIRIS = int(entry_code.get())
            nouveau_nom = f"S-{code_iris}-Stagiaires.csv"
            nouveau_chemin = os.path.join(
                os.path.dirname(self._chemin_csv_stagiaires),
                nouveau_nom
            )

            try:
                os.rename(self._chemin_csv_stagiaires, nouveau_chemin)
                self._chemin_csv_stagiaires = nouveau_chemin
            except Exception as e:
                tk.messagebox.showerror("Erreur", f"Impossible de renommer le fichier :\n{e}")
                log_erreur(f"impossible de renommer le fichier :\n{e}")
                return  # Ne pas fermer la fenêtre si erreur

            fenetre.destroy()

        def annuler():
            fenetre.destroy()
            log_erreur("code IRIS non renseignée pour " + self._chemin_csv_stagiaires + " → exit()")
            exit()

        code_iris = None

        fenetre = tk.Tk()
        fenetre.title("Code IRIS à renseigner manuellement")
        fenetre.resizable(False, False)
        fenetre.geometry("500x180")
        fenetre.eval('tk::PlaceWindow . center')

        # Fermer avec Échap
        fenetre.bind("<Escape>", lambda e: annuler())
        # Entrée = bouton Valider
        fenetre.bind("<Return>", lambda e: bouton_valider.invoke())

        label_warning = tk.Label(
            fenetre,
            text="⚠ Code IRIS non trouvé automatiquement, veuillez le spécifier manuellement",
            font=("Segoe UI", 10, "bold"),
            fg="orange"
        )
        label_warning.pack(pady=(10, 5))

        label_chemin = tk.Label(
            fenetre,
            text=f"Chemin du fichier source :\n  *{self._chemin_csv_stagiaires}*",
            font=("Segoe UI", 9),
            justify="left",
            wraplength=480
        )
        label_chemin.pack(pady=(0, 10))

        frame_saisie = tk.Frame(fenetre)
        frame_saisie.pack()

        entry_code = tk.Entry(frame_saisie, width=10, justify="center", font=("Segoe UI", 12))
        entry_code.pack()
        entry_code.focus()

        entry_code_var = tk.StringVar()
        entry_code["textvariable"] = entry_code_var
        entry_code_var.trace_add("write", verifier_entree)

        frame_boutons = tk.Frame(fenetre)
        frame_boutons.pack(pady=10)

        bouton_valider = tk.Button(frame_boutons, text="Valider", state="disabled", width=20, command=valider)
        bouton_valider.grid(row=0, column=0, padx=5)

        bouton_renommer = tk.Button(frame_boutons, text="Valider et remplacer le nom du CSV", state="disabled", width=30, command=valider_et_renommer)
        bouton_renommer.grid(row=0, column=1, padx=5)

        bouton_annuler = tk.Button(frame_boutons, text="Annuler", width=10, command=annuler)
        bouton_annuler.grid(row=0, column=2, padx=5)

        fenetre.mainloop()

        return code_iris
    
    def demander_code(self, typeCode) -> int|str:
        # Initialisation en fonction du type de code
        match typeCode:
            case "Code IRIS":
                print("Cas code IRIS")
                nbCaracteres = 5
                chemin_a_tester = self._chemin_csv_stagiaires
            case "Trigramme formation":
                print("Cas trigramme formation")
                nbCaracteres = 3
                chemin_a_tester = self._chemin_csv_stagiaires
            case _:
                log_erreur(f"cas non valide : soit 'IRIS' soit 'Trigramme formation', demandé : {typeCode} → exit()")
                exit()

            
        def verifier_entree(*args):
            val = entry_code.get()
            bouton_valider.config(state="normal" if val.isdigit() and len(val) == nbCaracteres else "disabled")
            if typeCode=="Code IRIS":
                bouton_renommer.config(state="normal" if val.isdigit() and len(val) == nbCaracteres else "disabled")

        def valider():
            nonlocal code # Adaptation : changer en "code"
            code = entry_code.get()            
            fenetre.destroy()

        # Cas IRIS uniquement : ne pas toucher / ne pas afficher si cas Trigramme formation
        def valider_et_renommer():
            nonlocal code
            code = entry_code.get()
            nouveau_nom = f"S-{code}-Stagiaires.csv"
            nouveau_chemin = os.path.join(
                os.path.dirname(self._chemin_csv_stagiaires),
                nouveau_nom
            )

            try:
                os.rename(self._chemin_csv_stagiaires, nouveau_chemin)
                self._chemin_csv_stagiaires = nouveau_chemin
            except Exception as e:
                tk.messagebox.showerror("Erreur", f"Impossible de renommer le fichier :\n{e}")
                log_erreur("Erreur", f"Impossible de renommer le fichier :\n{e}")
                return  # Ne pas fermer la fenêtre si erreur

            fenetre.destroy()

        def annuler():
            fenetre.destroy()
            log_erreur(f"{typeCode} non renseignée pour {chemin_a_tester} → exit()")
            exit()

        code = None

        fenetre = tk.Tk()
        fenetre.title(f"{typeCode} à renseigner manuellement")
        fenetre.resizable(False, False)
        fenetre.geometry("500x180")
        fenetre.eval('tk::PlaceWindow . center')

        # Fermer avec Échap
        fenetre.bind("<Escape>", lambda e: annuler())
        # Entrée = bouton Valider
        fenetre.bind("<Return>", lambda e: bouton_valider.invoke())

        label_warning = tk.Label(
            fenetre,
            text=f"⚠ {typeCode} non trouvé automatiquement.\nVeuillez le spécifier manuellement",
            font=("Segoe UI", 10, "bold"),
            fg="orange"
        )
        label_warning.pack(pady=(10, 5))

        label_chemin = tk.Label(
            fenetre,
            text=f"Chemin du fichier source :\n  *{chemin_a_tester}*",
            font=("Segoe UI", 9),
            justify="left",
            wraplength=480
        )
        label_chemin.pack(pady=(0, 10))

        frame_saisie = tk.Frame(fenetre)
        frame_saisie.pack()

        entry_code = tk.Entry(frame_saisie, width=10, justify="center", font=("Segoe UI", 12))
        entry_code.pack()
        entry_code.focus()

        entry_code_var = tk.StringVar()
        entry_code["textvariable"] = entry_code_var
        entry_code_var.trace_add("write", verifier_entree)

        frame_boutons = tk.Frame(fenetre)
        frame_boutons.pack(pady=10)

        bouton_valider = tk.Button(frame_boutons, text="Valider", state="disabled", width=20, command=valider)
        bouton_valider.grid(row=0, column=0, padx=5)

        if typeCode == "Code IRIS":
            bouton_renommer = tk.Button(frame_boutons, text="Valider et remplacer le nom du CSV", state="disabled", width=30, command=valider_et_renommer)
            bouton_renommer.grid(row=0, column=1, padx=5)

        bouton_annuler = tk.Button(frame_boutons, text="Annuler", width=10, command=annuler)
        bouton_annuler.grid(row=0, column=2, padx=5)

        fenetre.mainloop()

        code = int(code) if code.isdigit() else str(code)
        return code

    def recupere_trig_formation_depuis_chemin(self, chemin:str|None) -> str:
        """
        Extrait un trigramme (3 lettres/chiffres) depuis un chemin, ou le demande à l'utilisateur si introuvable.
        Gère les slashs / et \\ de manière robuste.
        """
        # On récupère le trigramme de la formation depuis le chemin du CSV ou alors on demande à l'utilisateur via tkinter
        trigramme = None

        if chemin:
            # Normalise les slashs pour s'assurer que le chemin est cohérent
            chemin_normalise = chemin.replace("\\", "/")

            # Découpe le chemin en parties
            parties = chemin_normalise.split("/")

            # Recherche un segment de 3 caractères alphanumériques
            for part in parties:
                if re.fullmatch(r"[a-zA-Z0-9]{3}", part):
                    trigramme = part
                    break

            # Si rien trouvé, on demande à l'utilisateur
            if not trigramme:
                trigramme = self.demander_code("Trigramme formation")

        else:
            trigramme = self.demander_code("Trigramme formation")

        #print(instance._trigrammeFormation)

        return trigramme

    def ouvrir_ou_creer_evaluationsFormation(self, trigramme:str) -> bool:
        """
        Ouvre ou crée le fichier Excel d'évaluations globales pour une formation donnée.

        Cette méthode construit le chemin vers le fichier d'évaluations correspondant au 
        trigramme de la formation. Si ce fichier existe, il est ouvert et les données 
        des stagiaires sont chargées dans un DataFrame. Sinon, un nouveau fichier est 
        créé à partir d'un modèle, et les données seront à initialiser.

        Args:
            trigramme (str): Code trigramme de la formation.

        Returns:
                - Un booléen indiquant si les anciennes données doivent être supprimées et remplacées 
                (`True` si nouveau fichier créé, `False` sinon).
        """ 
        # On définit le chemin vers les évaluations de la formation (le fichier qui va concaténer toutes les évaluation d'une formation)
        self._chemin_excel_evaluations_formation = self._chemin_excel_evaluations_defaut.replace("###", trigramme)

        # On vérifie que le répertoire dédié existe sinon on le créée : \\instnt\partage\FORMATIONS_C\###\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations
        os.makedirs(os.path.dirname(self._chemin_excel_evaluations_formation), exist_ok=True)
        
        # On regarde si Evaluation-Stagiaires-Global-XXX.xlsx existe
        if os.path.isfile(self._chemin_excel_evaluations_formation):
            # Alors on l'ouvre
            self.ouvrir_fe_evaluations_formation()
            
            # On récupère la liste des sessions déjà intégrées (liste des codes et des chemins) → Ce sera pour écrire dans le tkinter
            #dico_sessionsDejaTraitees = dict(
            #df_formation_stagiaires[df_formation_stagiaires["Trigramme formation"] == trigramme]     # 1. filtre sur le trigramme
            #.drop_duplicates(subset=["Code IRIS", "Chemin fichier CSV"])  # 2. élimine les doublons
            #[["Code IRIS", "Chemin fichier CSV"]]            # 3. sélection des colonnes
            #.values                               # 4. valeurs du DF
            #    )        

            # Il ne faudra pas supprimer les anciennes données de fe_evaluations_formation
            supprimeDonneesEtRemplace = False  
        else :
            # On créée le fichier : on ouvre le modèle (on sauvera avec le bon nom à la fin)
            fe_evaluations_formation = FichierExcel.depuis_modele(chemin_modele=self._chemin_modeleExcel_stagiaires, chemin_fichier_sauv=self._chemin_excel_evaluations_formation)

            # Il faudra supprimer les anciennes données de fe_evaluations_formation
            supprimeDonneesEtRemplace = True

            df_formation_stagiaires = None
            vlog.ajouter_message("Création EvalStat Global formation", fe_evaluations_formation.chemin_fichier, style=["vert"])

        return fe_evaluations_formation, df_formation_stagiaires, supprimeDonneesEtRemplace

    def ouvrir_fe_evaluations_formation(self) -> None:
        self._fe_evaluations_formation = FichierExcel.depuis_fichier(chemin_fichier=self._chemin_excel_evaluations_formation)
        self._fe_evaluations_formation._tableaux["CSV_stagiaires"].charge_df()
        #self._fe_evaluations_formation._tableaux["Stagiaires"]._df['Trigramme formation'] = self._fe_evaluations_formation._tableaux["Stagiaires"]._df['Trigramme formation'].astype(str)
        self._df_formation_stagiaires = self._fe_evaluations_formation._tableaux["Stagiaires"]._df  # Alias

    def filedialog_csv(self, code_IRIS:int, trigramme:Optional[str] = None) -> str|None:
        root = tk.Tk()
        root.withdraw()  # Ne pas afficher la fenêtre principale

        if (trigramme is None) and (self._chemin_excel_evaluations_formation is not None):
            trigramme = self._chemin_excel_evaluations_formation
        
        if trigramme:
            # On pré-définit le chemin où sont sensés être stockés les CSV d'évaluation des stagiaires
            chemin_repertoire_csv = optimiseCheminRepertoire(os.path.dirname(self._chemin_excel_evaluations_defaut.replace("###", trigramme)))

        while True:
            chemin = filedialog.askopenfilename(title=f"Sélectionner le fichier CSV de la session {code_IRIS}", initialdir=chemin_repertoire_csv or os.getcwd, filetypes=[("Fichiers CSV", "*.csv")])

            if chemin:
                print(f"Fichier sélectionné : {chemin}")
                return chemin  # ✅ Fichier sélectionné → on retourne

            # ❌ Aucun fichier sélectionné → boîte personnalisée
            reponse = self.filedialog_csv_pasDeReponse()

            if reponse == "choisir":
                continue  # 🔁 Re-ouvrir le file dialog
            elif reponse == "absent":
                print("Pas de CSV disponible pour cette session.")
                return None
            elif reponse == "quitter":
                print("Traitement interrompu par l'utilisateur.")
                sys.exit()
            else:
                print("Réponse inattendue. Fermeture.")
                sys.exit()
    
    @staticmethod
    def filedialog_csv_pasDeReponse():
        choix = {} # Astuce : les dictionnaires sont dispo dans les sous-fonctions sans avoir à les déclarer nonlocal (plutôt qu'un str par ex.)

        def choisir_nouveau():
            choix["reponse"] = "choisir"
            fenetre.destroy()

        def pas_disponible():
            choix["reponse"] = "absent"
            fenetre.destroy()

        def quitter():
            choix["reponse"] = "quitter"
            fenetre.destroy()

        fenetre = tk.Tk()
        fenetre.title("Aucun fichier sélectionné")
        fenetre.geometry("400x150")
        fenetre.resizable(False, False)
        fenetre.eval('tk::PlaceWindow . center')  # Centrer la fenêtre

        label = tk.Label(fenetre, text="Aucun fichier n'a été sélectionné.\nQue souhaitez-vous faire ?", pady=20)
        label.pack()

        bouton_frame = tk.Frame(fenetre)
        bouton_frame.pack()

        tk.Button(bouton_frame, text="Pas de CSV pour cette session", width=35, command=pas_disponible).grid(row=0, column=0, padx=5, pady=5)
        tk.Button(bouton_frame, text="Choisir CSV à nouveau", width=25, command=choisir_nouveau).grid(row=1, column=0, padx=5, pady=5)
        tk.Button(bouton_frame, text="Quitter traitement", width=20, command=quitter).grid(row=2, column=0, padx=5, pady=5)

        fenetre.mainloop()
        return choix.get("reponse")

class Traiter_REE:
    def __init__(self):
        # Fichier renseigné par la ressource extérieure
        self._chemin_ficheAdministrative:str = None
        self._word_ficheAdministrative:FichierWord = None  

        # Fichier Excel à remplir pour Laetitia Da Mota (RH INSTN qui s'occupe de rentrer les REE dans IRIS)
        self._chemin_modele_excel_ficheIntervenant:str = r"\\harmonie\INSTN\UEM\_Documents_communs\Formations\Formateurs\P09-Pr01-Qualifier les ressources enseignantes\P09_Pr01_Ta.E_Grille des critères de qualification des compétences_V1.xlsx"
        self._excel_ficheIntervenant:FichierExcel = None  

        self._mail_rh:Mail = None  # Mail à envoyer à la RH INSTN qui s'occupe de rentrer les REE dans IRIS (Laetitita Da Mota)

        self._repertoire_sauvegarde_fichiersREE:str = r"\\harmonie\INSTN\UEM\_Documents_communs\Formations\Formateurs\1.Intervenants - Documents administratifs" # Lieu où sauvegarder les fichiers de l'intervenant

        # Association colonnes excel avec command control Word
        # La préparation de ce ditionnaire peut être faite avec : fe = FichierExcel.depuis_fichier(chemin_fichier=r"T:\_Documents_communs\Formations\Formateurs\P09-Pr01-Qualifier les ressources enseignantes\P09_Pr01_Ta.E_Grille des critères de qualification des compétences_V1.xlsx", charger_df=True) ; fe._tableaux["QualificationsREE"].generer_dictionnaire_depuis_excel()
        self._dict_colExcel_cc:Dict[str, str] = {
            "NOM": "Nom",
            "Pr\u00e9nom": "Prenoms",
            "Dipl\u00f4me ou formation/exp\u00e9rience professionnelle": "Diplome",
            "Dur\u00e9e exp\u00e9rience professionnelle": "DureeExperiencePro",
            "Niveau d'expertise permettant une reconnaissance": "NiveauExpertise",
            "Domaine / Sp\u00e9cialit\u00e9 \nde l'expertise": "DomaineExpertise",
            "ATTRIBUTION Niveau comp\u00e9tence": None,
            "Combien de jours anim\u00e9s, en moyenne par an": "formation_nbJoursAnimes",
            "Combien de jours de formations suivies en p\u00e9dagogie (=animation)": "formation_nbJoursFormationPedagogie",
            "Profils d'apprenants form\u00e9s": "formation_profilApprenants",
            "Taux consolid\u00e9 de la satisfaction des apprenants relativement \u00e0 l'enseignant-formateur consid\u00e9r\u00e9": None,
            "Estimation par le RP de la capacit\u00e9 de l'enseignant-formateur \u00e0 animer \n(fond de salle)": None,
            "Outils num\u00e9riques utilis\u00e9s durant les animations r\u00e9alis\u00e9es\n(serious game, blended-learning\u2026)": "formation_outilsNumeriques",
            "Combien de jours pass\u00e9s en conception de s\u00e9quence de formation, en moyenne par an": "IngPedago_nbJoursConception",
            "Combien de jours de formations suivies en ing\u00e9nierie p\u00e9dagogique (=conception de s\u00e9quences de formation)": "IngPedago_nbJoursFormationIngPedago",
            "Estimation par le RP de la conception de la s\u00e9quence en fonction des objectifs p\u00e9dagogiques fournis par le RP\n(fond de salle, analyse des supports fournis)": None,
            "Estimation par le RP de la pertinence de l'\u00e9valuation des acquis r\u00e9alis\u00e9e par l'enseignant-formateur sur sa s\u00e9quence\n(analyse de la progression des apprenants : tests avant/apr\u00e8s)": None,
            "Estimation par le RP de l'utilisation des m\u00e9thodes actives\n(\u00e9tudes de cas, r\u00e9solution de probl\u00e8mes, classes invers\u00e9es, travaux de groupes\u2026)": None,
            "Combien d'ann\u00e9es d'exp\u00e9rience en conception de dispositifs de formations\n(=cr\u00e9ation et coordination)": "IngFormation_nbJoursConception",
            "Combien de jours de formations suivies en ing\u00e9nierie de formation\n(=conception de dispositifs de formation)": "IngFormation_nbJoursFormationIngFormation",
            "Estimation par le chef de projet ou le CUE de la complexit\u00e9 des pr\u00e9c\u00e9dents dispositifs de formation con\u00e7us": None,
            "Profil des apprenants des dispositifs de formations prc\u00e9demment con\u00e7us": "IngFormation_profilApprenants",
            "Estimation par le chef de projet ou le CUE de l'\u00e9valuation des acquis r\u00e9alis\u00e9 dans le dispositif de formation\n(mesure de la progression des apprenants=estimation de la qualit\u00e9 du dispositif de formation)": None,
            "Combien d'ann\u00e9es d'exp\u00e9rience en tant que tuteur acad\u00e9mique": "IngFormation_nbAnneesTuteur",
            "Combien de r\u00e9f\u00e9rentiels d'activit\u00e9, de comp\u00e9tence et d'\u00e9valuation r\u00e9alis\u00e9s": "IngFormation_nbAnneesTuteur",
            "Combien de jours de formations suivies en ing\u00e9nierie de comp\u00e9tences": "IngCompetences_nbReferentiels",
            "Estimation par la cellule p\u00e9dagogique de DPF de la complexit\u00e9 des pr\u00e9c\u00e9dentes r\u00e9alisations de l'ing\u00e9nieur/consultant en ing\u00e9nierie de comp\u00e9tences \n(complexit\u00e9 du m\u00e9tier et de son environnement : risques, r\u00e9glementation...)": "IngCompetences_nbJoursFormationIngCompetences",
            "ATTRIBUTION Niveau comp\u00e9tence ": None,
            "Evaluation CECRL ou \u00e9quivalence TOEIC, TOEFL": "ResultatLangue2",
            "ATTRIBUTION Niveau comp\u00e9tence  ": None,
            "Curriculum vitae": None,
        }



        # Fichier Excel qui liste les AI des intervenants
        self._chemin_excel_AI:str = r"\\harmonie\INSTN\UEM\_Documents_communs\Formations\Formateurs\2.AI\Liste AI formateurs.xlsx"
        self._excel_AI:FichierExcel = None
        self._onglet_excel_AI:str = "Intervenants" #Ligne ET = 3, 1èreligne = 4
        self._tableauStructure_AI:str = "ListeIntervenants"
        nbLignes_avantET_AI:int = 2 # Si tableau structuré, normalement on n'en a pas besoin

        # Fichier Excel qui liste des intervenants qui centralise les coordonnées 
        self._chemin_excel_listeIntervenants:str = r"\\harmonie\INSTN\UEM\_Echanges\VTE\Planning UEM.xlsm"
        self._excel_listeIntervenants:FichierExcel = None
        self._onglet_excel_listeIntervenants:str = "Liste intervenants" #Ligne ET = 4, 1èreligne = 5
        self._tableauStructure_listeIntervenant:str = "ListeIntervenants"
        nbLignes_avantET_listeIntervenants:int = 3  # Si tableau structuré, normalement on n'en a pas besoin
        


        # === Début code ===

        # On ouvre le word et on charge tous les command control (filedialog depuis "Download"). On le ferme
        self._word_ficheAdministrative = FichierWord.depuisFichier()
        #print(self._word_ficheAdministrative)

        # On crée le répertoire dans le répertoire des REE s'il n'existe pas (ou assimilé) (NOM Prénom (Société - AAAA))
        self._repertoire_sauvegarde_fichiersREE += f"\\{self._word_ficheAdministrative.cc['Nom'].upper()} {self._word_ficheAdministrative.cc['Prenoms'].title()} ({self._word_ficheAdministrative.cc['RaisonSociale'] if self._word_ficheAdministrative.cc['RaisonSociale'] != 'Raison sociale employeur principal' else 'CEA'} - {datetime.now().year})"
        #print(self._repertoire_sauvegarde_fichiersREE)
        #os.makedirs(self._repertoire_sauvegarde_fichiersREE, exist_ok=True)

        # On sélectionne tous les fichiers de la REE et on les déplace dans le répertoire idoine
        #self.deplacer_fichiers(self._repertoire_sauvegarde_fichiersREE)

        # On ouvre le fichier Excel à remplir pour Laetitia Da Mota (c'est un modèle, on le collera avec le bon nom dans le répertoire idoine)
        fichier_sauvegarde_fichiersREE = os.path.join(self._repertoire_sauvegarde_fichiersREE, os.path.basename(self._chemin_modele_excel_ficheIntervenant))
        self._excel_ficheIntervenant = FichierExcel.depuis_modele(self._chemin_modele_excel_ficheIntervenant, fichier_sauvegarde_fichiersREE, charger_df=True)
        #print(self._excel_ficheIntervenant._tableaux["QualificationsREE"]._df)

        # On écrit le dataframe du tableau QualificationsREE avec les données de l'intervenant provenant du word
        df_REE = self._excel_ficheIntervenant._tableaux["QualificationsREE"]._df  # Alias
        nouvelle_ligne = {}
        for col_df, cc_key in self._dict_colExcel_cc.items():
            if cc_key is None:
                # Pas de clé correspondante => valeur vide dans la DataFrame
                nouvelle_ligne[col_df] = None
            else:
                # Récupérer la valeur dans le dictionnaire Word, ou None si la clé absente
                valeur = self._word_ficheAdministrative._cc.get(cc_key, None)
                nouvelle_ligne[col_df] = convertir_si_possible(valeur)
                #print(valeur, type(convertir_si_possible(valeur)))

        # Ajouter la nouvelle ligne au DataFrame
        # TODO : non pas sûr
        df_REE = pd.concat([df_REE, pd.DataFrame([nouvelle_ligne])], ignore_index=True)

        # On pré-rempli le fichier Excel fiche intervenant grâce aux CC et au dictionnaire
        self._excel_ficheIntervenant._tableaux["QualificationsREE"].ecrit_dataFrame_dans_tableauStructure(df_REE, remplace_df_par_nouveau=True)

        # On sauve la fiche intervenant
        #self._excel_ficheIntervenant.save()
        #self._excel_ficheIntervenant.close()

        # On ouvre l'Excel et le Word pour comparaison et adaptations manuelles

        # Dès que l'Excel est fermé, on prépare le mail pour Laetitia

        # On met à jour le fichier Excel Liste AI formateurs.xlsx : onglet intervenant, on cherche et remplace la date de validité de l'attestation employeur sinon nouvelle ligne (recopier formule + format)
        # On met à jour le fichier Excel  avec la liste des intervenants :  on cherche et remplace les données mail, tel, Ville, la date de validité de l'attestation employeur... sinon nouvelle ligne (recopier formule + format)

        
    def deplacer_fichiers(self, destination: str = None) -> None:
        """
        Ouvre un dialogue pour sélectionner des fichiers, puis les déplace vers un dossier choisi.

        Args:
            destination (str, optional): Chemin du dossier de destination.
                                        Si None, un dialogue s'ouvrira pour le choisir.
        """

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
        for fichier in fichiers:
            nom_fichier = os.path.basename(fichier)
            chemin_destination = os.path.join(destination, nom_fichier)

            try:
                shutil.move(fichier, chemin_destination)
                print(f"✅ Déplacé : {nom_fichier}")
            except Exception as e:
                print(f"❌ Erreur avec {nom_fichier} : {e}")

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

    Validation :
       - mail de premier contact (à partir d'un modèle .msg) → Oui mais TODO signature à enlever
       - préparer les RDV Outlook pour les entretiens (à partir de modèles .oft) ;
       - faire les mails de relance pour des documents ou Studea.

    """
    
    @dataclass
    class PropEntretien:
        chemin_modele: str
        duree: timedelta
        sujet: str
        date_debut:datetime
        categorie: str = "FI"
    
    # Fichier Excel avec les informations des étudiants
    #_chemin_fichier_etudiants:str
    _prefixe_sujet:str
    
    _fe_etudiants:FichierExcel
    _df_etudiants:DataFrame
    
    _chemin_modele_mail_priseContact:Optional[str] = None
    _ficheEvaluation1:Optional[str] = None
    _ficheEvaluation2:Optional[str] = None

    _envoyer_mail:bool = False
    _entretiens:List[PropEntretien] = []
    _mail_responsables_univ:Optional[str] = None


    # === Initialisations statiques ===
    # Colonnes du fichier etudiant à récupérer
    _COLONNES_FE_ETUDIANTS = [
        "Cursus", 
        "Nom", "Prénom", "Mail apprenti", "Téléphone apprenti", 
        "Entreprise ", "Lieu entreprise", "Nom TE", "Prénom TE", "Mail TE", "Téléphone TE", "Ma fonction de suivi de l'alternant", 
        "Engagement des parties", "Entretien initial", "1ère visite", "2ème visite", "Fiche évaluation 1", "Fiche évaluation 2"]


    def __init__(self, chemin_fichier_etudiants:str, nom_onglet:str) -> None:

        # Fichier Excel avec les informations des étudiants
        self._chemin_fichier_etudiants = chemin_fichier_etudiants

        # Par défaut, il faut charger l'Excel des coordonnées des étudiants
        self._fe_etudiants = FichierExcel.depuis_fichier(
            chemin_fichier = chemin_fichier_etudiants,
            nom_onglet = nom_onglet
        )

        # On réduit le DataFrame aux informations qui nous sont utiles
        self._df_etudiants = self._fe_etudiants._tableaux[nom_onglet]._df[self._COLONNES_FE_ETUDIANTS] # On ne garde que les colonnes qui nous intéressent mais attention ça reste une vue dont les modifications affectent le dataframe initial
        self._df_etudiants = self._df_etudiants[self._df_etudiants["Ma fonction de suivi de l'alternant"] == "Tuteur"]

    @classmethod
    def UGA(cls) -> Traiter_contactsApprentis:
        """
        """

        # === Initialisations statiques ===
        prefixe_sujet = "Master IN - Suivi d'alternance"
        
        nom_onglet = "Etudiants"
        envoyer_mail = False

        chemin_fichier_etudiants =  r"\\instnt\partage\FORMATIONS_I\GDRA - Master IN (parcours ADIN-GDRA-SN)\2025-2026\1-dossier etudiants\Master ADIN - 2025_2026.xlsx"
        ficheEvaluation1 =          r"\\instnt\partage\FORMATIONS_I\GDRA - Master IN (parcours ADIN-GDRA-SN)\2024-2025\9-stages\Fiche_Evaluation_alternance_M2_Ingénierie_Nucléaire_janvier.docx"
        ficheEvaluation2 =          r"\\instnt\partage\FORMATIONS_I\GDRA - Master IN (parcours ADIN-GDRA-SN)\2024-2025\9-stages\Fiche_Evaluation_alternance_M2_Ingénierie_Nucléaire_aout.docx"
        chemin_modele_mail_priseContact = r"C:\Users\vt238770\Documents\_CEA\Modèles adaptés\Mails\Master IN\Master IN - Tutorat en entreprise - Prise de contact.msg"
        
        mail_responsables_univ = "master-in-responsables@univ-grenoble-alpes.fr"

        prise_de_fonction = cls.PropEntretien(
            chemin_modele = r"C:\Users\vt238770\Documents\_CEA\Modèles adaptés\Mails\Master IN\Master IN - Suivi d'alternance - Entretien.oft",
            duree = timedelta(hours=0, minutes=45),
            sujet = "Entretien de prise de fonction",
            date_debut = RDV_Outlook.get_lundi_depuis_num_semaine(numero_semaine=41) # (Autour du 6 octobre : dernière semaine de la première période en entreprise → A faire avant mi-novembre)
        )
        
        premiere_visite = cls.PropEntretien(
            chemin_modele = r"C:\Users\vt238770\Documents\_CEA\Modèles adaptés\Mails\Master IN\Master IN - Suivi d'alternance - Entretien.oft",
            duree = timedelta(hours=1),
            sujet = "1ère visite en entreprise",
            date_debut = RDV_Outlook.get_lundi_depuis_num_semaine(numero_semaine=51) # (Autour du 15 décembre : dernière semaine avant vacances Noël et reprise école → A faire avant mi-janvier)
        )

        deuxieme_visite = cls.PropEntretien(
            chemin_modele = r"C:\Users\vt238770\Documents\_CEA\Modèles adaptés\Mails\Master IN\Master IN - Suivi d'alternance - Entretien.oft",
            duree = timedelta(hours=1),
            sujet = "2ème visite en entreprise",
            date_debut = RDV_Outlook.get_lundi_depuis_num_semaine(numero_semaine=17) # (Autour du 20 avril : dernière semaine reprise école → A faire avant fin mai)
        )

        # Initialisation de l'instance
        instance = cls(chemin_fichier_etudiants, nom_onglet)

        instance._prefixe_sujet = prefixe_sujet
        instance._chemin_modele_mail_priseContact = chemin_modele_mail_priseContact
        instance._envoyer_mail = envoyer_mail
        instance._mail_responsables_univ = mail_responsables_univ

        instance._entretiens.append(prise_de_fonction)
        instance._entretiens.append(premiere_visite)
        instance._entretiens.append(deuxieme_visite)

        instance._ficheEvaluation1 = ficheEvaluation1
        instance._ficheEvaluation2 = ficheEvaluation2

        # Tests :
        instance._df_etudiants = instance._df_etudiants.head(1)
        
        return instance

                

    def contactInitial(self, chemin_modele_mail_priseContact:Optional(str)=None) -> None:
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
    def creer_html_mail_relance(choix:int, ligne:pd.Series, lettreInterlocuteur:str) -> str|None:
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
        

    def creer_html_mail_ficheEvaluation(self, ligne:pd.Series, date:date) -> str:
        
        periode = self.periode_scolaire_UGA()
        date_str = date.strftime('%d %B')     

        corps_html = """
        <html>
        <body>
            <p>Bonjour,</p>
            <p>Dans le cadre de la notation de """ + f'{ligne["Prénom"]} {ligne["Nom"]}' + f" lors de sa prochaine soutenance de {periode}, " + """est-ce que vous pouvez compléter et signer le document en PJ s'il vous plait ?</p>
            <p>Il sera à retourner à moi-même en mettant en copie""" + f"{self._mail_responsables_univ}" + f"pour le <span style='color: red; font-weight: bold;'>{date_str} au plus tard</span> " + """.</p>
            <p>Je vous envoie un avis de rdv pour faire office de pense-bête. Il sera mis au """ + f"{date_str}" + """ avec un rappel une semaine avant (mais en mode « disponible » histoire de ne pas bloquer le créneau sur vos agendas donc vous pouvez l’accepter dans risque).</p>
            """

        if periode == "fin d'année" :
            corps_html += """<p>Si vous êtes en vacances durant cette période, veillez me l’envoyer avant de profiter de votre repos mérité !</p>"""

        corps_html += """
                <p>Je vous remercie, passez une excellente fin de journée,</p>
            </body>
            </html>"""
       
        return corps_html

    def creer_mail_ficheEvaluation(self, ligne:pd.Series, date:date):
        corps_html = self.creer_html_mail_ficheEvaluation(ligne, date)
        pieces_jointes = self.defini_pj_ficheEvaluation_UGA()
        sujet = self._prefixe_sujet + " - Fiche d'évaluation à compléter et retourner"
        
        Mail.creer_mail(
            destinataires = ligne["Mail TE"],
            sujet = sujet,
            corps_html = corps_html,
            pieces_jointes = pieces_jointes,
            envoyer_mail = self._envoyer_mail
        )

        # TODO : Faire envoi d'un RDV à date donnée avec PJ et rappel une semaine avant
        # Adapter l'appli une fois fini



    def defini_pj_ficheEvaluation_UGA(self) -> str:
        """
        Sélectionne et renvoei le chemin de la bonne fiche d'évaluation en fonction de la période (mi-année ou fin d'année)
        """
        periode = self.periode_scolaire_UGA()

        if (periode == "mi-année"):
            #print("On est entre septembre et février (inclus)")
            pieces_jointes = self._ficheEvaluation1
        else:
            #print("On est entre mars et août")
            pieces_jointes = self._ficheEvaluation2
        
        return pieces_jointes

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
    def periode_scolaire_UGA() -> str:
        """
        Renvoie "mi-année" si date actuelle entre septembre et février.
        Renvoie "fin d'année" sinon
        """
        # Période à définir pour la fiche d'évaluation à employer
        mois = date.today().month  # 1=janvier, ..., 12=décembre

        if ((9 <= mois) and (mois <= 2)):  # septembre (9) → février (2)
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




### --------------------------------------------------------------------
#  Fonctions globales INSTN
### --------------------------------------------------------------------

def initialisationListeFichiersExportsIRIS():
    """
    Initialise les tuples des fichiers Excel (extracttions d'IRIS) à concaténer / traiter

    :Example:

    >>> initialisationListeFichiersExportsIRIS()

    .. seealso:: Rien du tout.
    .. warning:: Rien du tout.
    .. note:: Rien du tout.
    .. todo:: Rien du tout.
    """
    tSessions = (
        'R04110_Sessions-2011 à 2014 FINAL.xlsx',
        'R04110_Sessions-2015 FINAL.xlsx',
        'R04110_Sessions-2016 FINAL.xlsx',
        'R04110_Sessions-2017 FINAL.xlsx',
        'R04110_Sessions-2018 FINAL.xlsx',
        'R04110_Sessions-2019 FINAL.xlsx',
        'R04110_Sessions-2020 FINAL.xlsx',
        'R04110_Sessions-2021 FINAL.xlsx',
        'R04110_Sessions-2022 FINAL.xlsx',
        'R04110_Sessions-2023 FINAL.xlsx',
        'R04110_Sessions-2024 FINAL.xlsx',
        'R04110_Sessions-2025 au 2025.08.07.xlsx')

    tFormations = (
        "R0304_Ref_Formation-Listedesformations-2025.06.06.xlsx", )

    tVentes = (
        'R04301_Sessions-Ventes-FC2020 FINAL.xlsx',
        'R04301_Sessions-Ventes-FC2021 FINAL.xlsx',
        'R04301_Sessions-Ventes-FC2022 FINAL.xlsx',
        'R04301_Sessions-Ventes-FC2023 FINAL.xlsx',
        'R04301_Sessions-Ventes-FC2024 FINAL.xlsx',
        'R04301_Sessions-Ventes-filtre sur FC-2025 au 2025-02-05 LG.xlsx',
        'R04301_Sessions-Ventes-filtre sur FC-2025 au 2025-03-03 LG.xlsx',
        'R04301_Sessions-Ventes-filtre sur FC-2025 au 2025-04-01 LG.xlsx',
        'R04301_Sessions-Ventes-filtre sur FC-2025 au 2025-05-12 LG.xlsx',
        'R04301_Sessions-Ventes-filtre sur FC-2025 au 2025-06-02 LG.xlsx',
        'R04301_Sessions-Ventes-filtre sur FC-2025 au 2025-07-01 LG.xlsx',
        'R04301_Sessions-Ventes-filtre sur FC-2025 au 2025-08-01 LG.xlsx')
        
    tInscriptions = (
        'R04500_Sessions-Inscriptions-FC2020 FINAL.xlsx',
        'R04500_Sessions-Inscriptions-FC2021 FINAL.xlsx',
        'R04500_Sessions-Inscriptions-FC2022 FINAL.xlsx',
        'R04500_Sessions-Inscriptions-FC2023 FINAL.xlsx',
        'R04500_Sessions-Inscriptions-FC2024 FINAL.xlsx',
        'R04500_Sessions-Inscriptions-filtre sur FC-2025 au 2025-02-05.xlsx',
        'R04500_Sessions-Inscriptions-filtre sur FC-2025 au 2025-03-03.xlsx',
        'R04500_Sessions-Inscriptions-filtre sur FC-2025 au 2025-04-01.xlsx',
        'R04500_Sessions-Inscriptions-filtre sur FC-2025 au 2025-05-12.xlsx',
        'R04500_Sessions-Inscriptions-filtre sur FC-2025 au 2025-06-02.xlsx',
        'R04500_Sessions-Inscriptions-filtre sur FC-2025 au 2025-07-01.xlsx',
        'R04500_Sessions-Inscriptions-filtre sur FC-2025 au 2025-08-01.xlsx')

    return {
        "Sessions" : tSessions,
        "Formations" : tFormations,
        "Ventes" : tVentes,
        "Inscriptions" : tInscriptions
    }

def mettreAJourTousLesExportsIRIS_auto(tuple_types:Tuple(str)):
    """
    Permet de créer un seul fichier Excel à partir de plusieurs exports d'IRIS.
    Les fichiers à traités sont initialisés par la fonction initialisationListeFichiersExportsIRIS() qui permet à l'utilisateur de tout lister à la main (ça peut être plus pratique dans certains cas afin d'éviter de passer par une sélection manuelle)
    
    Les fichiers output sont des modèles avec les mêmes colonnes que les extracts d'IRIS mais avec de meilleures formes (format, couleurs...) + des colonnes adjointes à la fin pour extraire et séparer les infos du n° de session ou de la référence de la formation (ex. : trigramme formation, trigramme RP, trigramme AF...)

    Les tuples des fichiers Excel à traiter sont enregistrés dans des instances de TravauxFichiersIRIS et on emploie les méthodes de cette classe

    :param tuple_types: tuple de strings avec les noms des extracts 
    :type donnees: Tuple[str, ...]

    :Example:
    
    >>> mettreAJourTousLesExportsIRIS_auto(("Formations", ))
    >>> mettreAJourTousLesExportsIRIS_auto(("Formations", "Sessions", "Ventes", "Inscriptions"))


    .. seealso:: Rien du tout.
    .. warning:: Si une seule valeur pour tuple_types, bien mettre sous cette forme : ("Formations",) car sans la virgule Python interprête juste comme un string
    .. note:: Rien du tout.
    .. todo:: Rien du tout.
    """
    # On initialise tous les chemins des fichiers à concaténer
    dict_DE_IRIS["Fichiers"] = initialisationListeFichiersExportsIRIS()

    # On traite à la suite
    for clef in dict_DE_IRIS["CodesExports"].keys():
        if clef in tuple_types:
            print(Style.BRIGHT + Fore.YELLOW + "\nTraitement des exports " + clef)
            fichiers_formates = ["\n\t" + f for f in dict_DE_IRIS["Fichiers"][clef]]
            print("Liste des fichiers traités : " + ", ".join(fichiers_formates))
            chemins_fichiersInput = tuple(os.path.join(dict_DE_IRIS["PropExportIRIS"][clef]._input.repertoire, nom) for nom in dict_DE_IRIS["Fichiers"][clef])
            TravauxFichiersIRIS.avecEcritureOutputDefaut(dict_DE_IRIS["PropExportIRIS"][clef], chemins_fichiersInput=chemins_fichiersInput)

def mettreAJourTousLesExportsIRIS_fileDialog(tuple_types:Tuple[str, ...]):
    """
    Permet de créer un seul fichier Excel à partir de plusieurs exports d'IRIS.
    Les fichiers à traités sont sélectionnés à la suite par l'utilisateur à travers un filedialog dans l'ordre du doctionnaire d'entrée, puis tous les fichiers sont traités successivement après.
    
    Les fichiers output sont des modèles avec les mêmes colonnes que les extracts d'IRIS mais avec de meilleures formes (format, couleurs...) + des colonnes adjointes à la fin pour extraire et séparer les infos du n° de session ou de la référence de la formation (ex. : trigramme formation, trigramme RP, trigramme AF...)

    Les tuples des fichiers Excel à traiter sont enregistrés dans des instances de ExtractIRIS et on emploie les méthodes de cette classe

    :param tuple_types: tuple de strings avec les noms des extracts 
    :type donnees: Tuple[str, ...]

    :Example:
    >>> mettreAJourTousLesExportsIRIS_fileDialog(("Formations", ))
    >>> mettreAJourTousLesExportsIRIS_fileDialog(("Formations", "Sessions", "Ventes", "Inscriptions"))


    .. seealso:: Rien du tout.
    .. warning:: Si une seule valeur pour tuple_types, bien mettre sous cette forme : ("Formations",) car sans la virgule Python interprête juste comme un string
    .. note:: Rien du tout.
    .. todo:: Rien du tout.
    """
    # On fait choisir les fichiers à l'utilisateur
    dict_traitements = {}
    for clef, codeIRIS in dict_DE_IRIS["CodesExports"].items():
        if clef in tuple_types:
            traitement = TravauxFichiersIRIS(dict_DE_IRIS["PropExportIRIS"][clef])
            dict_traitements[clef] = traitement

    # On traite à la suite
    for clef in dict_traitements.keys():
        print(Style.BRIGHT + Fore.YELLOW + "\nTraitement des exports " + clef)
        fichiers_formates = ["\n\t" + f for f in dict_traitements[clef]._chemins_fichiersInput]
        print("Liste des fichiers traités : " + ", ".join(fichiers_formates))
        #print("Liste des fichiers traités : " + ", ".join(dictEI[clef].nomsFichiers_exportIRIS))
        #dict_traitements[clef].lire_extractIRIS()
        #dict_traitements[clef].ecrit_dataFrame_dans_tableauStructure(df=instance._df_chemins, supprimeDonneesEtRemplace=True)
        TravauxFichiersIRIS.avecEcritureOutputDefaut(dict_DE_IRIS["PropExportIRIS"][clef], chemins_fichiersInput=dict_traitements[clef]._chemins_fichiersInput)

def lire_fdc(chemin_fdc):

    """
    A partir d'un chemin de fichier Excel (qui se doit d'etre une fiche de couts, on retourne plusieurs dataframes.

    :param s_fdc: Chemin du fichier Excel a ouvrir (se doit d'etre une fiche de coûts)
    :type s_fdc: string
    :return: Un DataFrame de l'extract IRIS avec ajouts de colonnes (on a extrait les informations de la colonne 'N° Session' par decoupage)
    :rtype: DataFrame

    :Example:

    >>> DataFrame.df_sessionsIRIS = lire_fdc("C:\\Users\\fichier.xlsx")


    .. seealso:: Rien du tout.
    .. warning:: Rien du tout.
    .. note:: Rien du tout.
    .. todo:: Rien du tout.
    """

    nomOnglet_fdc = "Fiche de coûts"

    # Premier tableau : B5:C21 (informations génériques)
    df_fdc_infos = pd.read_excel(chemin_fdc, sheet_name=nomOnglet_fdc, usecols=[1, 2], names=["Critere", "Valeur"], header=3, nrows=17)
    #print(df_fdc_infos)
    dateCreationFormation = df_fdc_infos.iloc[4, 1] #C9
    if isinstance(dateCreationFormation, int):
        dateCreationFormation = dateCreationFormation
    elif hasattr(dateCreationFormation, 'year'):
        dateCreationFormation = dateCreationFormation.year
    else:
        print(f"Valeur inattendue pour une année : {dateCreationFormation} (type {type(dateCreationFormation)})")
        dateCreationFormation = 1900

    dureeJours_fdc = df_fdc_infos.iloc[6, 1] #C11
    osThematique = df_fdc_infos.iloc[8, 1] #C13
    nbCible_fcd = df_fdc_infos.iloc[12, 1] #C17

    # Prix de vente défini par le RP (cellule J22)
    prixVenteRetenuParParticipant = pd.read_excel(chemin_fdc, sheet_name=nomOnglet_fdc, usecols=[9], names=["Valeur"], skiprows=20, nrows=1).iloc[0,0]
    
    # Deuxième tableau : tableau des coûts (tout compris) et des prix par personne (T1, T2 et T3) : ref K35:N35
    df_fdc_couts = pd.read_excel(chemin_fdc, sheet_name=nomOnglet_fdc, usecols=[10, 11, 12, 13, 14, 15], names=["T1", "T3", "T2", "0.9xT3", "1.1xT3", "Valeur fixée"], header=33, nrows=2).transpose() #Grosse astuce : je mets 2 lignes de plus pour affecter les noms plus facilement et je les recalculerai après
    df_fdc_couts.rename(columns={0: "Couts fixes et variables nb cible"}, inplace=True) # Pour changer le nom de la colonne après transposition. Coûts prix fixes et prix
    df_fdc_couts.head() # Requis pour MàJ le nom de la colonne après transposition

    df_fdc_couts.loc["0.9xT3", "Couts fixes et variables nb cible"] = df_fdc_couts.loc["T3", "Couts fixes et variables nb cible"] * 0.9
    df_fdc_couts.loc["1.1xT3", "Couts fixes et variables nb cible"] = df_fdc_couts.loc["T3", "Couts fixes et variables nb cible"] * 1.1
    df_fdc_couts.loc["Valeur fixée", "Couts fixes et variables nb cible"] = prixVenteRetenuParParticipant / 1.1 * nbCible_fcd # Astuce : comme c'est une valeur que je n'ai pas, je fais le calcul inverse que pour avoir le montant par session avec aleas


    # On ajoute les colonnes montants cibles avec calculs
    l1 = []
    l2 = []
    l3 = []
    for index, row in df_fdc_couts.iterrows():
        l1.append(row["Couts fixes et variables nb cible"] * 1.1)
        l2.append(row["Couts fixes et variables nb cible"] * 1.1 / nbCible_fcd)
        l3.append(row["Couts fixes et variables nb cible"] * 1.1 / nbCible_fcd / dureeJours_fdc)
    df_fdc_couts["Montant cible par session avec aléas"] = l1
    df_fdc_couts["Montant cible par participant"] = l2
    df_fdc_couts["Montant cible par participant et par jour"] = l3


    # On ajoute les colonnes pour calcul nb participants
    l1 = []
    l2 = []
    l3 = []
    for index, row in df_fdc_couts.iterrows():
        if row["Montant cible par participant"] != 0 :
            l1.append(math.ceil(df_fdc_couts.loc["T1", "Montant cible par session avec aléas"] / row["Montant cible par participant"]))
            l2.append(math.ceil(df_fdc_couts.loc["T2", "Montant cible par session avec aléas"] / row["Montant cible par participant"]))
            l3.append(math.ceil(df_fdc_couts.loc["T3", "Montant cible par session avec aléas"] / row["Montant cible par participant"]))
        else :
            l1.append(0)
    df_fdc_couts["Min participants T1"] = l1
    df_fdc_couts["Min participants T2"] = l2
    df_fdc_couts["Min participants T3"] = l3 

    print(df_fdc_couts)

    return df_fdc_infos, df_fdc_couts, prixVenteRetenuParParticipant, dateCreationFormation, dureeJours_fdc, osThematique, nbCible_fcd

def fenetreBilan():
    def valider_champs(*args):
        trig = entry_trigramme.get().strip()
        annee = entry_annee.get().strip()
        bouton_generer.config(
            state="normal" if len(trig) == 3 and annee.isdigit() and len(annee) == 4 else "disabled"
        )

    def generer_bilan(event=None):
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
        bf.mergeBilan(chemin_word_bilan_input, chemin_word_bilan_output)
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

    bouton_generer = ttk.Button(frame_boutons, text="Générer bilan", state="disabled", command=generer_bilan)
    bouton_generer.grid(row=0, column=0, padx=5)

    bouton_annuler = ttk.Button(frame_boutons, text="Annuler", command=annuler)
    bouton_annuler.grid(row=0, column=1, padx=5)

    # Validation en temps réel
    entry_trigramme.bind("<KeyRelease>", valider_champs)
    entry_annee.bind("<KeyRelease>", valider_champs)

    # Entrée = clic sur bouton générer
    fenetre.bind("<Return>", generer_bilan)

    # Échap = fermeture de la fenêtre
    fenetre.bind("<Escape>", lambda e: fenetre.destroy())

    # Lancer la fenêtre
    fenetre.mainloop()



### --------------------------------------------------------------------
#  Initialisations variables globales communes
### --------------------------------------------------------------------

# Pour couleur barres de progression
colorama.init(autoreset=True)

# Pour chrono des fonctions
timer = Timer()

# Pour message de sortie applis externes
vlog = Vlog()










### --------------------------------------------------------------------
#  Initialisations pour mises à jour des Extracts IRIS
### --------------------------------------------------------------------

# ==== Initialisation exports IRIS ====
dictCodesIRIS = {
        "Sessions" : "R04110",
        "Formations" : "R0304",
        "Ventes" : "R04301",
        "Inscriptions" : "R04500",
    }

sessions = PropExportIRIS(
    nom_typeExport = "Sessions",
    codeExport = "R04110",

    repertoire_input = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts originaux",
    nom_onglet_input = "Data",
    nbLignes_avantET_input = 1,

    repertoire_modele = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Modèles",
    nom_fichier_modele = "R04110_Sessions-Modèle.xlsx",

    repertoire_output = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts complets",
    nom_fichier_output = "R04110_Sessions-COMPLET.xlsx"
    )

formations = PropExportIRIS(
    nom_typeExport = "Formations",
    codeExport = "R0304",

    repertoire_input = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts originaux",
    nom_onglet_input = "Data",
    nbLignes_avantET_input = 0,

    repertoire_modele = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Modèles",
    nom_fichier_modele = "R0304_Formations-Modèle.xlsx",

    repertoire_output = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts complets",
    nom_fichier_output = "R0304_Formations-COMPLET.xlsx"
    )

ventes = PropExportIRIS(
    nom_typeExport = "Ventes",
    codeExport = "R04301",

    repertoire_input = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts originaux", #Car il y a des petits bugs sur certains CSV
    nom_onglet_input = "Data",
    nbLignes_avantET_input = 1,

    repertoire_modele = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Modèles",
    nom_fichier_modele = "R04301_Ventes-Modèle.xlsx",

    repertoire_output = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts complets",
    nom_fichier_output = "R04301_Ventes-COMPLET.xlsx"
    )

inscriptions = PropExportIRIS(
    nom_typeExport = "Inscriptions",
    codeExport = "R04500",

    repertoire_input = r"\\instnt\HOME\REFERENC\IRIS - rapports de synthese",
    nom_onglet_input = "Data",
    nbLignes_avantET_input = 1,

    repertoire_modele = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Modèles",
    nom_fichier_modele = "R04500_Inscriptions-Modèle.xlsx",

    repertoire_output = r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts complets",
    nom_fichier_output = "R04500_Inscriptions-COMPLET.xlsx"
    )

dict_DE_IRIS = {}
dict_DE_IRIS["CodesExports"] = {
        "Sessions" : "R04110",
        "Formations" : "R0304",
        "Ventes" : "R04301",
        "Inscriptions" : "R04500",
    }
dict_DE_IRIS["PropExportIRIS"] = {
        "Sessions" : sessions, 
        "Formations" : formations,
        "Ventes" : ventes,
        "Inscriptions" : inscriptions}




### --------------------------------------------------------------------
#   Initialisations pour création bilans pédagogiques
### --------------------------------------------------------------------

# ==== Initialisation variables utilisateur ====
# Initialisation des chemins des répertoires
#rep_gedMiroir = r"\\instnt\partage\FORMATIONS_C"
rep_fdc_defaut = r"\\instnt\partage\FORMATIONS_C\XXX\P05-P06-dossier-conception-referentiel\fiche-de-cout-et-code-de-formation"
rep_specsPedagogiques_defaut = r"\\instnt\partage\FORMATIONS_C\XXX\P05-P06-dossier-conception-referentiel\specifications-pedagogiques-et-referentiel"


# Modèle du bilan à remplir
chemin_word_bilan_input = r'C:\Users\vt238770\Documents\_CEA\Prog\Modèles\Bilan formation.docx'

# Bilan en sortie après remplissage
chemin_word_bilan_output = r'C:\Users\vt238770\Documents\_CEA\Prog\Modèles\Bilan formation - output.docx'

# Fiche de coûts
#chemin_fdc = r'\\instnt\partage\FORMATIONS_C\948\P05-P06-dossier-conception-referentiel\fiche-de-cout-et-code-de-formation\Fiche de coûts - 948 - Elaboration de scénarios de DEM - 2025.01.24.xlsx'

# Specs pédagogiques
#chemin_specsPedagogiques = r'\\instnt\partage\FORMATIONS_C\948\P05-P06-dossier-conception-referentiel\specifications-pedagogiques-et-referentiel\P06-Pr01-F01_Specifications-pedagogiques - 948 - 2025.04.pdf'


"""
from __future__ import annotations
from .office import *

import warnings
warnings.filterwarnings("ignore", message="Slicer List extension is not supported and will be removed")

from fileinput import filename
from typing import Dict, List, Tuple, Optional

import pandas as pd
import pandas as DataFrame

import win32com.client
import extract_msg
import os
import shutil
import sys
import re
import time
import math

from datetime import date, datetime, timedelta, time
from dataclasses import dataclass
from collections import defaultdict

from tabulate import tabulate

from tqdm import tqdm
import colorama
from colorama import Fore, Style

from mailmerge import MailMerge

import tkinter as tk
from tkinter import filedialog, ttk, messagebox
"""