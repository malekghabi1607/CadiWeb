from __future__ import annotations

import math
from pathlib import Path
from typing import Optional, Protocol

from vte.core import config
from vte.utils.utils import *
from vte.utils.office import FichierExcel
from vte.utils.utils_instn import recupere_trig_formation_depuis_chemin

# TODO : Pour l'instant c'est une classe de traitemnt. Le jour où j'ai besoin d'ouvrir un EvalStat pour le lire uniquement, prendre modèle sur IRIS avec des classes de lecture et de traitement

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


# ======================================================================================
# CLASSE FDC
# ======================================================================================
class FdC:
    """
    Classe qui gère l'ouverture d'une fiche de coûts INSTN et l'accès à ses différentes valeurs
    """

    
    # === VARIABLES DE CLASSE COMMUNES À TOUTES LES INSTANCES ===
    _NOM_ONGLET_TABLEAU = "Fiche de coûts"




    # ====================
    # === CONSTRUCTEUR ===
    # ====================
    def __init__(self, formation: Optional[Formation_protocol] = None, fe: Optional[FichierExcel] = None) -> None:
        """
        Initialise une instance FdC.

        :param formation: l'instance de Formation pour la fiche de coûts (nécessaire uniquement pour facilite la sélection du fichier de FdC (pré-sélection répertoire))
        :type formation: Optional[Formation_protocol], optional
        :param fe: Objet FichierExcel de la fiche de coûts (contient le chemin de la FdC).
        :type fe: Optional[FichierExcel], optional
        """
        # Variables propres à la FdC
        self._formation:Optional[Formation_protocol] = formation  # Trigramme de la formation ; nécessaire uniquement pour facilite la sélection du fichier de FdC (pré-sélection répertoire)
        self._fe: Optional[FichierExcel] = fe  # Objet Excel contenant la fiche de coûts

        # Chargement de la FdC
        self._charger_fdc()

        
    # =========================
    # === METHODES INTERNES ===
    # =========================
    def _choisir_fdc(self) -> Path | None:
        """
        Ouvre un filedialog pour demander à l'utilisateur de sélectionner une fiche de coûts.
        On pointe au mieux sur le répertoire des FdC pour la boîte de dialogue.
        
        :return: le chemin de la fiche de coûts
        :rtype: Path | None
        """

        return choisir_fichier(titre=f"Sélectionner la fiche de coûts à employer.",
                        types_fichiers=[("Fichiers Excel", "*.xlsx")],
                        dossier_initial=self.dossier_plan_classement,
                        texte_bouton_choisir=f"Choisir FdC à nouveau"
                        )

    def _charger_fdc(self):
        """
        Charge l'Excel de la fiche de coûts dans l'instance.
        Si aucun fichier Excel n'est dans l'instance (i.e. pas de chemin pour la FdC), alors on ouvre un filedialog
        """

        if self._fe is not None:
            #  Soit le FichierExcel et son onglet (objet TableauExcel) existent déjà : pas besoin de le recharger
            if self.nom_onglet in self._fe.tableaux:
                pass
            else:  # Sinon _fe a été initialisé à minima (juste le chemin) et il faut le charger
                timer.debut("Chargement fiche de coûts")
                self._fe = FichierExcel.depuis_fichier(chemin_fichier=self.chemin, nom_onglet=self.nom_onglet)
                timer.fin()
        
        elif self._fe is None:  # Si aucun fichier Excel n'est dans l'instance (i.e. pas de chemin pour la FdC), alors on ouvre un filedialog
            chemin = self._choisir_fdc()
            if chemin is not None:
                timer.debut("Chargement fiche de coûts")
                self._fe = FichierExcel.depuis_fichier(chemin_fichier=chemin, nom_onglet=self.nom_onglet)
                timer.fin()
            else:
                vlog.log_erreur("Le fichier FdC n'a pas été sélectionné")
        
        else: # Pas besoin de le charger
            vlog.log_erreur("Je suis sorti des conditions sans avoir chargé ma FdC")
       
        

    


    # =========================
    # === METHODES EXTERNES ===
    # =========================

    def max_participants(self, depassementAutorise:int) -> int:
        """
        Renvoie le nombre max de participants (= prévu + dépassement autorisé)

        Args:
            depassementAutorise (int): nb de dépassement autorisé

        Returns:
            int: le nombre max de participants (= prévu + dépassement autorisé)
        """
        return self.prevus_participants + depassementAutorise



    # =========================
    # === GETTERS / SETTERS ===
    # =========================
    @property
    def trigramme_formation(self) -> str|None:
        """
        Renvoie le trigramme de la formation.
        On la lit soit :
           - depuis l'instance Formation ;
           - en l'extrayant depuis le chemin du fichier Excel de la fiche de coûts
        
        Si _formation et _fe sont None les 2, alors on renvoie None

        :return: le trigramme de la formation
        :rtype: int|None
        """
        if self._formation is not None:
            return self._formation.trigramme_formation
        elif self._fe is not None:
            return recupere_trig_formation_depuis_chemin(self._fe.chemin_fichier)
        else :
            return None
    
    @property
    def chemin(self) -> Path:
        """
        Renvoie le chemin de la fiche de coûts.
        
        :return: le chemin de la fiche de coûts
        :rtype: Path
        """
        return self._fe.chemin_fichier

    @property
    def nom_onglet(self) -> str:
        return self._NOM_ONGLET_TABLEAU

    @property
    def tableau(self) -> FichierExcel._TableauExcel:
        #fe_fdc._tableaux["Fiche de coûts"]
        return self._fe.tableaux[self.nom_onglet]

    @property
    def date_derniere_modification(self) -> str:
        """
        Extrait la date de la fiche de coût à partir de la dernière date de modif (valeur système windows)
        
        :return: la date de la fiche de coût à partir de la dernière date de modif (valeur système windows)
        :rtype: str
        """
        dateFdC = datetime.fromtimestamp(self.chemin.stat().st_mtime)
        sDateFdC = dateFdC.strftime("%d/%m/%Y")
        return sDateFdC

    @property
    def dossier_plan_classement(self) -> Path:
        return config.format_path(config.REPERTOIRE_FDC, trigramme_formation=self.trigramme_formation)


    # ========================================================
    # === GETTERS données FdC (lié à la version de la FdC) ===
    # ========================================================
    @property
    def nom_formation(self) -> str:
        """
        Renvoie le nom de la formation (C5)
        
        :return: le nom de la formation tel que donné dans la fiche de coûts
        :rtype: str
        """
        return self.tableau["C5"]

    @property
    def nb_participants_prevus(self) -> int:
        """
        Renvoie le nombre de participants prévus (C17)
        
        :return: le nombre de participants prévus tel que donné dans la fiche de coûts
        :rtype: int
        """
        return self.tableau["C17"]

    @property
    def annee_creationFormation(self) -> int:
        """
        Retourne l'année de conception de la formation (C9).
        
        :return: l'année de conception de la formation
        :rtype: int
        """
        anneeCreationFormation = self.tableau["C9"]

        if isinstance(anneeCreationFormation, int):
            anneeCreationFormation = anneeCreationFormation
        elif hasattr(anneeCreationFormation, 'year'):
            anneeCreationFormation = anneeCreationFormation.year
        else:
            print(f"Valeur inattendue pour une année : {anneeCreationFormation} (type {type(anneeCreationFormation)})")
            anneeCreationFormation = 1900
        return anneeCreationFormation

    @property
    def min_participants_cea(self) -> int:
        """
        Renvoie le nombre min de participants prévus pour les CEA (M22)
        
        :return: le nombre min de participants tel que donné dans la fiche de coûts
        :rtype: int
        """
        return self.tableau["M22"]
    
    @property
    def min_participants_ee(self) -> int:
        """
        Renvoie le nombre min de participants prévus pour les EE (N23)
        
        :return: le nombre min de participants tel que donné dans la fiche de coûts
        :rtype: int
        """
        return self.tableau["N23"]
    
    @property
    def min_participants(self) -> str:
        """
        Renvoie le nombre min de participants prévus pour les EE et CEA.
        
        :return: le nombre min de participants EE et CEA tel que donné dans la fiche de coûts
        :rtype: str
        """
        return f"{self.min_participants_ee} pers. (EE) / {self.min_participants_cea} pers. (CEA)"
    

    # ==========================================================================
    # === ANCIENNE METHODE DE LECTURE DE LA FDC (lié à la version de la FdC) ===
    # ==========================================================================
    def lire_fdc_bak(self):

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



        """
        # On charge le fichier fdc si non déjà ouvert
        self._fe_IRIS_sessions = IRIS.charger_excel_IRIS_sessions(fe_IRIS_sessions=self._fe_IRIS_sessions, chemin_IRIS_sessions=chemin_IRIS_sessions)



        self._chemin_fdc = filedialog.askopenfilename(title="Sélectionner la dernière fiche de coûts", filetype=[("Fichiers Excel", "*.xlsx")], initialdir=optimiseCheminRepertoire(self._repertoire_fdc_defaut.replace("XXX", self._codeFormation)))
        if not chemin_fichier_session:
            log_erreur("click sur cancel du filedialog → Pas de chemin de fiche de coûts")
        self._dateFdC = datetime.fromtimestamp(self._chemin_fdc.stat().st_mtime)
        self._sDateFdC = self._dateFdC.strftime("%d/%m/%Y")
        
        self._df_fdc_infos, self._df_fdc_couts, self._prixVenteRetenuParParticipant, self._dateCreationFormation, self._dureeJours_fdc, self._osThematique, self._nbCible_fcd = lire_fdc(self._chemin_fdc)
        """

        nomOnglet_fdc = "Fiche de coûts"

        # Premier tableau : B5:C21 (informations génériques)
        df_fdc_infos = pd.read_excel(self._chemin_fdc, sheet_name=nomOnglet_fdc, usecols=[1, 2], names=["Critere", "Valeur"], header=3, nrows=17)
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
        prixVenteRetenuParParticipant = pd.read_excel(self._chemin_fdc, sheet_name=nomOnglet_fdc, usecols=[9], names=["Valeur"], skiprows=20, nrows=1).iloc[0,0]
        
        # Deuxième tableau : tableau des coûts (tout compris) et des prix par personne (T1, T2 et T3) : ref K35:N35
        df_fdc_couts = pd.read_excel(self._chemin_fdc, sheet_name=nomOnglet_fdc, usecols=[10, 11, 12, 13, 14, 15], names=["T1", "T3", "T2", "0.9xT3", "1.1xT3", "Valeur fixée"], header=33, nrows=2).transpose() #Grosse astuce : je mets 2 lignes de plus pour affecter les noms plus facilement et je les recalculerai après
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


