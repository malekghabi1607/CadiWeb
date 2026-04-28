from __future__ import annotations

from abc import ABC, abstractmethod
from functools import cached_property
import math
from pathlib import Path
from typing import Optional, Protocol
from datetime import date, datetime

from vte.core import config
from vte.core.fichierGenerique import FichierGenerique
from vte.utils.utils import *
from vte.utils.office import FichierExcel
from vte.utils.utils_instn import construire_chemin_config, recupere_trig_formation_depuis_chemin
from openpyxl import load_workbook

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
# VARIABLES DE MODULE
# ======================================================================================
# ALLER A LA FIN DU MODULE POUR REDEFINIR LE REGISTRY DES LECTEURS FDC



# ======================================================================================
# CLASSE FDC – ENTREE PUBLIQUE
# ======================================================================================
class FdC:
    """
    Point d'entrée unique pour manipuler une FdC
    """

    def __init__(self, lecteur: FdC_Lecteur):
        """
        Initialise une instance FdC.

        :param lecteur:
            Implémentation concrète du lecteur de FdC correspondant à la version du fichier.
            Ce paramètre est interne et ne doit pas être fourni manuellement.
        :type lecteur: Classe abstraite _FdCBase
        """
        # Implémentation concrète du lecteur de FdC correspondant à la version du fichier.
        self._lecteur:FdC_Lecteur = lecteur

    @classmethod
    def ouvrir(cls, chemin: Optional[Path] = None, formation: Optional[Formation_protocol] = None, fe: Optional[FichierExcel] = None) -> FdC:
        """
        Initialise et ouvre une instance FdC.

        (Redéfinit la méthode de FichierGenerique)

        :param chemin: chemin à employer pour la fiche de coût (supplante le chemin par défaut de la config)
        :type chemin: Path
        :param formation: l'instance de Formation pour la fiche de coûts (nécessaire uniquement pour facilite la sélection du fichier de FdC (pré-sélection répertoire))
        :type formation: Optional[Formation_protocol], optional
        :param fe: Objet FichierExcel de la fiche de coûts (contient le chemin de la FdC).
        :type fe: Optional[FichierExcel], optional
        """
        #if chemin is None:
        #    chemin = cls._choisir_fdc(formation)
        
        lecteur = FdC_Lecteur.ouvrir(
            chemin=chemin,
            formation=formation,
            fe=fe
        )
        return cls(lecteur)

    def __getattr__(self, item):
        """
        Délègue dynamiquement l'accès aux attributs et méthodes vers le lecteur interne.

        Permet d'exposer directement les propriétés métier de la FdC (ex: nom_formation,
        nb_participants_prevus, etc.) sans avoir à les redéfinir dans la classe FdC.

        Ainsi on peut appeler fdc._lecteur.nom_formation en écrivant fdc.nom_formation 

        :param item: Nom de l'attribut ou de la méthode demandé
        :type item: str
        :return: Attribut ou méthode du lecteur interne correspondant
        :raises AttributeError: si l'attribut n'existe pas dans le lecteur
        """
        return getattr(self._lecteur, item)



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
        return self.nb_participants_prevus + depassementAutorise







# ======================================================================================
# CLASSE FDC – CLASSE ABSTRAITE INTERNE
# ======================================================================================
class FdC_Lecteur(FichierGenerique, ABC):
    """
    Classe abstraite interne servant de modèle aux différentes versions de la FdC
    """
    
    _NOM_ONGLET_TABLEAU = "Fiche de coûts"  # Pour l'instant c'est toujours le même nom. Plus tard au besoin ce sera à surcharger dans les classes concrètes

    def __init__(self, chemin:Optional[Path], formation: Optional[Formation_protocol] = None, fe: Optional[FichierExcel] = None):
        """
        Initialise et ouvre une instance FdC.

        :param chemin: chemin à employer pour la fiche de coût (supplante le chemin par défaut de la config)
        :type chemin: Path
        :param formation: l'instance de Formation pour la fiche de coûts (nécessaire uniquement pour facilite la sélection du fichier de FdC (pré-sélection répertoire))
        :type formation: Optional[Formation_protocol], optional
        :param fe: Objet FichierExcel de la fiche de coûts (contient le chemin de la FdC).
        :type fe: Optional[FichierExcel], optional
        """
        super().__init__(chemin)

        # Variables propres à la FdC
        self._formation:Optional[Formation_protocol] = formation  # Trigramme de la formation ; nécessaire uniquement pour facilite la sélection du fichier de FdC (pré-sélection répertoire)
        self._fe: Optional[FichierExcel] = fe  # Objet Excel contenant la fiche de coûts

        # TODO : màj charger_fdc : ajouter méthode générique pour essayer d'aller chercher le fichier automatiquement et si non trouvé, alors l'utilisateur pointe
        # TODO : màj charger_fdc : Il faudra que je prenne en compte que si chemin est donné, alors ça supplante le reste
        #self._charger_fdc()  # chargement que si on en a besoin : mis dans self.fe
        #self._fe = FichierExcel.depuis_fichier(chemin_fichier=chemin, nom_onglet=self._NOM_ONGLET_TABLEAU)


    @classmethod
    def ouvrir(cls, chemin: Path, formation: Optional[Formation_protocol] = None, fe: Optional[FichierExcel] = None) -> FdC_Lecteur:
        for lecteur_cls in LECTEURS_FDC:
            if lecteur_cls.est_compatible(chemin):
                return lecteur_cls(
                    chemin=chemin,
                    formation=formation,
                    fe=fe
                )

        vlog.log_erreur("Aucun lecteur FdC compatible trouvé pour {chemin}")


    # =========================
    # === METHODES INTERNES ===
    # =========================
    def _resoudre_chemin(self) -> Path:
        """
        Permet de savoir quel chemin employer pour la fiche de coûts. Ordre de priorité :
            - chemin donné par l'utilisateur ;
            - fichier Excel ;
            - on ouvre un filedialog

        :return: Le chemin de l'Excel de la FdC à employer
        :rtype: Path
        """
        # 1) Cas prioritaire : si on a un chemin donné par l'utilisateur, alors c'est ce chemin qui fait foi
        if self.chemin:
            return self.chemin
        
        # 2) Si on a un Excel, alors c'est ce chemin qui fait foi
        if self._fe: # Attention : ne pas faire appel à self.chemin_fe car sinon on va lancer _charger_fdc or on est en train de résoudre le chemin là (antécédent)
            return self._fe.chemin_fichier

        # 3) On ouvre un filedialog
        chemin = self._choisir_fdc()
        if chemin is None:
            vlog.log_erreur("Le fichier FdC n'a pas été sélectionné")
        
        return chemin
    
    def _charger_fdc(self):
        """
        Charge l'Excel de la fiche de coûts dans l'instance.
        Si aucun fichier Excel n'est dans l'instance (i.e. pas de chemin pour la FdC), alors on ouvre un filedialog
        """
        chemin_fdc = self._resoudre_chemin()

        # On ouvre le fichier Excel avec le chemin
        timer.debut("Chargement fiche de coûts")
        self._fe = FichierExcel.depuis_fichier(chemin_fichier=chemin_fdc, nom_onglet=self.nom_onglet)
        timer.fin()



    # =========================
    # === METHODES EXTERNES ===
    # =========================
    @staticmethod
    def construire_chemin_repertoire_fdc(trigramme_formation:Optional[str]=None) -> Path:
        """
        Construit le chemin du répertoire de la fiche de coûts de la formation (à partir des données de la config REPERTOIRE_FDC) :
            - le chemin est transformé en unc (s'il y a un raccourci lecteur réseau sur le poste de l'utilisateur on transforme en chemin réseau complet) ;
            - l'utilisateur peut optimiser le chemin (chemin le plus long entre l'attendu et ce qui existe).

        :param trigramme_formation: Trigramme de la formation. Défaut = None
        :type trigramme_formation: Optional[str], optional

        :return: Le chemin de sortie du bilan de sessions.
        :rtype: Path
        """
        return construire_chemin_config(
            chemin_a_completer = config.REPERTOIRE_FDC,
            trigramme_formation = trigramme_formation
        )





    # ===========
    # === IHM ===
    # ===========
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




    # =========================
    # === GETTERS / SETTERS ===
    # =========================
    @cached_property
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
        elif self.fe is not None:
            return recupere_trig_formation_depuis_chemin(self.chemin_fe)
        else :
            return None

    @property
    def fe(self) -> FichierExcel :
        """
        Retourne l'objet FichierExcel de la Fiche de coûts.

        On ne charge la FdC que si l'on a besoin de _fe.

        :return: Objet FichierExcel de la Fiche de coûts
        :rtype: FichierExcel
        """
        if self._fe is None:
            self._charger_fdc()
        return self._fe

    @property
    def chemin_fe(self) -> Path:
        """
        Renvoie le chemin de la fiche de coûts à partir de l'objet Fichier Excel.

        Cette méthode est vouée à être retirée petit à petit pour remplacer la référence du chemin par la valeur de FichierGenerique.
        
        :return: le chemin de la fiche de coûts à partir de l'objet Fichier Excel
        :rtype: Path
        """
        # TODO : A virer petit à petit pour faire une ref à FichierGenerique.chemin
        return self.fe.chemin_fichier
        
    @property
    def tableau(self) -> FichierExcel._TableauExcel|None:
        """
        Retourne le tableau principal de la FdC (Objet TableauExcel)

        :return: Tableau principal de la FdC
        :rtype: FichierExcel._TableauExcel
        """
        """
        # fe toujours chargé car dans le init. Normalement plus besoin de ce test
        if self._fe is not None:
            return self._fe.tableaux[self._NOM_ONGLET_TABLEAU]
        else:
            vlog.log_erreur("Fichier Excel non chargé", continuer=True)
            return None
        """
        
        return self.fe.tableaux[self._NOM_ONGLET_TABLEAU]

    @property
    def nb_participants_min(self) -> str:
        """
        Renvoie le nombre min de participants prévus pour les EE et CEA.
        
        :return: le nombre min de participants EE et CEA tel que donné dans la fiche de coûts
        :rtype: str
        """
        return f"{self.nb_participants_min_ee} pers. (EE) / {self.nb_participants_min_cea} pers. (CEA)"

    @cached_property
    def dossier_plan_classement(self) -> Path:
        """
        Retourne le dossier de la FdC tel quel défini par le process qualité (i.e. valeur dans le fichier config).

        :return: Dossier de la FdC tel quel défini par le process qualité
        :rtype: Path
        """
        return FdC_Lecteur.construire_chemin_repertoire_fdc(trigramme_formation=self.trigramme_formation)

    @property
    def version(self) -> str | None:
        """
        Extrait la version de la FdC à partir de la cellule A1.

        Exemple attendu :
        "Fiche de coûts INSTN - V1.4 du 25/01/2019"

        :return: version extraite (ex: "V1.4 du 25/01/2019") ou None si non trouvée
        :rtype: str | None
        """
        valeur = str(self.tableau["A1"])

        match = re.search(r"(V\d+(?:\.\d+)?\s*du\s*\d{2}/\d{2}/\d{4})", valeur)

        return match.group(1) if match else None
    
    # === CONTRAT METIER (doit être dans les classes concrètes) ===
    @property
    @abstractmethod
    def nom_onglet(self) -> str:
        """
        Nom de l'onglet principal de la FdC

        :return: Le nom de l'onglet principal de la FdC
        :rtype: str
        """
        pass

    @property
    @abstractmethod
    def nom_formation(self) -> str:
        """
        Renvoie le nom de la formation (C5)
        
        :return: le nom de la formation tel que donné dans la fiche de coûts
        :rtype: str
        """
        pass

    @property
    @abstractmethod
    def annee_creationFormation(self) -> int:
        """
        Retourne l'année de conception de la formation (C9).
        
        :return: l'année de conception de la formation
        :rtype: int
        """
        pass

    @property
    @abstractmethod
    def date_creationFormation(self) -> date:
        """
        Retourne la date de conception de la formation (C9).
        
        Si la valeur est :
        - une date/datetime → retourne la date
        - un objet avec attribut year → retourne le 01/01/year
        - un entier → considéré comme une année → retourne le 01/01/année
        - une chaîne → tentative de conversion
        - sinon → retourne le 01/01/1900
        
        :return: la date de conception de la formation
        :rtype: date
        """
        pass

    @property
    @abstractmethod
    def nb_participants_prevus(self) -> int:
        """
        Renvoie le nombre de participants prévus (C17)
        
        :return: le nombre de participants prévus tel que donné dans la fiche de coûts
        :rtype: int
        """
        pass

    @property
    @abstractmethod
    def nb_participants_min_cea(self) -> int:
        """
        Renvoie le nombre min de participants prévus pour les CEA (M22)
        
        :return: le nombre min de participants tel que donné dans la fiche de coûts
        :rtype: int
        """
        pass
    
    @property
    @abstractmethod
    def nb_participants_min_ee(self) -> int:
        """
        Renvoie le nombre min de participants prévus pour les EE (N23)
        
        :return: le nombre min de participants tel que donné dans la fiche de coûts
        :rtype: int
        """
        pass
    
    

# ======================================================================================
# FDC VERSION 6.1 – IMPLEMENTATION DE LA CLASSE ABSTRAITE
# ======================================================================================
class FdC_Lecteur_V6_1(FdC_Lecteur):

    @classmethod
    # =========================
    # === METHODES EXTERNES ===
    # =========================
    def est_compatible(cls, chemin: Path) -> bool:
        """
        Permet de connaître la version du fichier

        :param chemin: _description_
        :type chemin: Path
        :return: _description_
        :rtype: bool
        """
        #fe = FichierExcel.depuis_fichier(chemin, nom_onglet="Fiche de coûts")
        #valeur = fe.tableaux["Fiche de coûts"]["A1"]
        #return "V6.1" in str(valeur)
        


    
        # exemple simple
        #return "V1" in chemin.name

        #wb = load_workbook(chemin, read_only=True, data_only=True)
        #ws = wb.active  # ou nom connu
        #valeur = ws["A1"].value

        #return valeur and "V6.1" in str(valeur)

        valeur = FichierExcel.valeur_cellule_ouverture_rapide(
            chemin_fichier=chemin, 
            nom_onglet=cls._NOM_ONGLET_TABLEAU, 
            cellule="A1"
            )
        #return valeur and "V6.1" in str(valeur)
        #return "V6.1" in str(valeur)
        return isinstance(valeur, str) and "V6.1" in valeur
    


    # =========================
    # === GETTERS / SETTERS ===
    # =========================
    @property
    def nom_onglet(self) -> str:
        return self._NOM_ONGLET_TABLEAU

    @property
    def nom_formation(self) -> str:
        """
        Renvoie le nom de la formation (C5)
        
        :return: le nom de la formation tel que donné dans la fiche de coûts
        :rtype: str
        """
        return self.tableau["C5"]

    @property
    def annee_creationFormation(self) -> int:
        """
        Retourne l'année de conception de la formation (C9).
        
        :return: l'année de conception de la formation
        :rtype: int
        """
        return convertir_dateFormatIndefini_annee(self.tableau["C9"])

    @property
    def date_creationFormation(self) -> date:
        """
        Retourne la date de conception de la formation (C9).
        
        Si la valeur est :
        - une date/datetime → retourne la date
        - un objet avec attribut year → retourne le 01/01/year
        - un entier → considéré comme une année → retourne le 01/01/année
        - une chaîne → tentative de conversion
        - sinon → retourne le 01/01/1900
        
        :return: la date de conception de la formation
        :rtype: date
        """
        return convertir_dateFormatIndefini_date(self.tableau["C9"])

    @property
    def nb_participants_prevus(self) -> int:
        """
        Renvoie le nombre de participants prévus (C17)
        
        :return: le nombre de participants prévus tel que donné dans la fiche de coûts
        :rtype: int
        """
        return self.tableau["C17"]

    @property
    def nb_participants_min_cea(self) -> int:
        """
        Renvoie le nombre min de participants prévus pour les CEA (M22)
        
        :return: le nombre min de participants tel que donné dans la fiche de coûts
        :rtype: int
        """
        return self.tableau["M22"]
    
    @property
    def nb_participants_min_ee(self) -> int:
        """
        Renvoie le nombre min de participants prévus pour les EE (N23)
        
        :return: le nombre min de participants tel que donné dans la fiche de coûts
        :rtype: int
        """
        return self.tableau["N23"]

# ======================================================================================
# FDC VERSION DEFAUT – IMPLEMENTATION DE LA CLASSE ABSTRAITE
# ======================================================================================
class FdC_Lecteur_Defaut(FdC_Lecteur):
    # =========================
    # === METHODES INTERNES ===
    # =========================
    @classmethod
    def est_compatible(cls, chemin):
        return True  # fallback

    # =========================
    # === GETTERS / SETTERS ===
    # =========================
    @property
    def nom_onglet(self) -> str:
        raise NotImplementedError("Version non reconnue")

    @property
    def nom_formation(self) -> str:
        raise NotImplementedError("Version non reconnue")

    @property
    def annee_creationFormation(self) -> int:
        raise NotImplementedError("Version non reconnue")

    @property
    def date_creationFormation(self) -> date:
        raise NotImplementedError("Version non reconnue")

    @property
    def nb_participants_prevus(self) -> int:
        raise NotImplementedError("Version non reconnue")

    @property
    def nb_participants_min_cea(self) -> int:
        raise NotImplementedError("Version non reconnue")
    
    @property
    def nb_participants_min_ee(self) -> int:
        raise NotImplementedError("Version non reconnue")
    
    




# ======================================================================================
# VARIABLES DE MODULE
# ======================================================================================
LECTEURS_FDC = [
    FdC_Lecteur_V6_1,
    FdC_Lecteur_Defaut,  # toujours en dernier
]










# ======================================================================================
# CLASSE FDC - OLD
# ======================================================================================
class FdC_OLD:
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
        self._charger_fdc()  # TODO : ❌ supprimé (remplacé par constructeur)

    @classmethod
    def depuis_chemin(cls, formation: Optional[Formation_protocol] = None, chemin_fdc: Optional[Path|str] = None) -> FdC:
        """
        Initialise une instance FdC à partir du chemin de la FdC.

        Si aucun chemin n'est donné, on ouvrira un filedialog.

        :param formation: l'instance de Formation pour la fiche de coûts (nécessaire uniquement pour facilite la sélection du fichier de FdC (pré-sélection répertoire))
        :type formation: Optional[Formation_protocol], optional
        :param chemin_fdc: chemin de la fiche de coûts. Défaut = None
        :type chemin_fdc: Optional[Path|str], optional
        """

        fe = FichierExcel(chemin_fichier=convertir_chemin_en_path(chemin_fdc))
        instance = FdC(formation=formation, fe=fe)

        return instance
        
        
    # =========================
    # === METHODES INTERNES ===
    # =========================
    def _choisir_fdc(self) -> Path | None:  # AA
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

    def _charger_fdc(self):  # TODO : ❌ supprimé (remplacé par constructeur)
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

    def max_participants(self, depassementAutorise:int) -> int: # AA
        """
        Renvoie le nombre max de participants (= prévu + dépassement autorisé)

        Args:
            depassementAutorise (int): nb de dépassement autorisé

        Returns:
            int: le nombre max de participants (= prévu + dépassement autorisé)
        """
        return self.prevus_participants + depassementAutorise


    @staticmethod
    def construire_chemin_repertoire_fdc(trigramme_formation:Optional[str]=None) -> Path:  # AA
        """
        Construit le chemin du répertoire de la fiche de coûts de la formation (à partir des données de la config REPERTOIRE_FDC) :
            - le chemin est transformé en unc (s'il y a un raccourci lecteur réseau sur le poste de l'utilisateur on transforme en chemin réseau complet) ;
            - l'utilisateur peut optimiser le chemin (chemin le plus long entre l'attendu et ce qui existe).

        :param trigramme_formation: Trigramme de la formation. Défaut = None
        :type trigramme_formation: Optional[str], optional

        :return: Le chemin de sortie du bilan de sessions.
        :rtype: Path
        """
        return construire_chemin_config(
            chemin_a_completer = config.REPERTOIRE_FDC,
            trigramme_formation = trigramme_formation
        )



    # =========================
    # === GETTERS / SETTERS ===
    # =========================
    @cached_property
    def trigramme_formation(self) -> str|None: # AA
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
    def chemin(self) -> Path:  # AA
        """
        Renvoie le chemin de la fiche de coûts.
        
        :return: le chemin de la fiche de coûts
        :rtype: Path
        """
        return self._fe.chemin_fichier

    @property
    def nom_onglet(self) -> str:  # AA
        return self._NOM_ONGLET_TABLEAU

    @property
    def tableau(self) -> FichierExcel._TableauExcel:  # AA
        #fe_fdc._tableaux["Fiche de coûts"]
        return self._fe.tableaux[self.nom_onglet]

    @property
    def date_derniere_modification_windows(self) -> str:  # AA
        """
        Extrait la date de la fiche de coût à partir de la dernière date de modif (valeur système windows)
        
        :return: la date de la fiche de coût à partir de la dernière date de modif (valeur système windows)
        :rtype: str
        """
        dateFdC = datetime.fromtimestamp(self.chemin.stat().st_mtime)
        sDateFdC = dateFdC.strftime("%d/%m/%Y")
        return sDateFdC

    @cached_property
    def dossier_plan_classement(self) -> Path:  # AA
        return FdC.construire_chemin_repertoire_fdc(trigramme_formation=self.trigramme_formation)


    # ========================================================
    # === GETTERS données FdC (lié à la version de la FdC) ===
    # ========================================================
    @property
    def nom_formation(self) -> str:  # AA
        """
        Renvoie le nom de la formation (C5)
        
        :return: le nom de la formation tel que donné dans la fiche de coûts
        :rtype: str
        """
        return self.tableau["C5"]

    @property
    def nb_participants_prevus(self) -> int:  # AA
        """
        Renvoie le nombre de participants prévus (C17)
        
        :return: le nombre de participants prévus tel que donné dans la fiche de coûts
        :rtype: int
        """
        return self.tableau["C17"]

    @property
    def annee_creationFormation(self) -> int:  # AA
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
    def date_creationFormation(self) -> date:  # AA
        """
        Retourne la date de conception de la formation (C9).
        
        Si la valeur est :
        - une date/datetime → retourne la date
        - un objet avec attribut year → retourne le 01/01/year
        - un entier → considéré comme une année → retourne le 01/01/année
        - une chaîne → tentative de conversion
        - sinon → retourne le 01/01/1900
        
        :return: la date de conception de la formation
        :rtype: date
        """
        date_excel = self.tableau["C9"]

        if isinstance(date_excel, datetime):
            return date_excel.date()

        if isinstance(date_excel, date):
            return date_excel

        if hasattr(date_excel, "year"):
            return date(date_excel.year, 1, 1)

        if isinstance(date_excel, int):
            return date(date_excel, 1, 1)

        if isinstance(date_excel, str):
            try:
                # tentative simple ISO (YYYY-MM-DD)
                return datetime.fromisoformat(date_excel).date()
            except Exception:
                try:
                    # tentative format FR courant
                    return datetime.strptime(date_excel, "%d/%m/%Y").date()
                except Exception:
                    try:
                        # si c'est juste une année en string
                        annee = int(date_excel)
                        return date(annee, 1, 1)
                    except Exception:
                        pass

        print(f"Valeur inattendue pour une date : {date_excel} (type {type(date_excel)})")
        return date(1900, 1, 1)

    @property
    def min_participants_cea(self) -> int:  # AA
        """
        Renvoie le nombre min de participants prévus pour les CEA (M22)
        
        :return: le nombre min de participants tel que donné dans la fiche de coûts
        :rtype: int
        """
        return self.tableau["M22"]
    
    @property
    def min_participants_ee(self) -> int:  # AA
        """
        Renvoie le nombre min de participants prévus pour les EE (N23)
        
        :return: le nombre min de participants tel que donné dans la fiche de coûts
        :rtype: int
        """
        return self.tableau["N23"]
    
    @property
    def min_participants(self) -> str:  # AA
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
        self._fe_IRIS_sessions = IRIS.charger_excel_IRIS_sessions(fe_IRIS_sessions=self._fe_IRIS_sessions)



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


