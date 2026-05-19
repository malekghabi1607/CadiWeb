from abc import ABC, abstractmethod
from pathlib import Path
from typing import Optional
from datetime import date, datetime

class FichierGenerique(ABC):
    """
    Classe de base pour tous les fichiers versionnés.
    Gère :
    - le chemin
    - la détection de version
    - le routage vers la bonne sous-classe
    """

    def __init__(self, chemin:Optional[Path]):
        self._chemin = chemin

        # TODO
        self._dateDocument:Optional[date] = None
        self._version:Optional[str] = None

    @property
    def chemin(self) -> Optional[Path]:
        return self._chemin
    
    @chemin.setter
    def chemin(self, valeur:Path) -> None:
        """
        Attention, le setter n'est pas dans l'objet fichier (ex. FdC), mais dans son lecteur. Si je fais fdc.chemin = Path("C:/aa.txt") ça ne marchera pas → Il faut déleguer le setter à l'objet (ici : FdC)

        :param valeur: _description_
        :type valeur: Path
        """
        self._chemin = valeur

    # TODO : rajouter les date et version dans classmethod ; peut-être faire cached property et renseigner la variable d'instance au 1er appel ; au second on lit cette variable

    # === FACTORY PRINCIPALE ===
    @classmethod
    def ouvrir(cls, chemin: Path):
        """
        Détecte la version et instancie la bonne sous-classe
        """
        for sous_classe in cls.__subclasses__():
            if sous_classe.est_compatible(chemin):
                return sous_classe(chemin)

        raise ValueError(f"Aucune version compatible pour : {chemin}")





    # === Affichage ===
    def __str__(self) -> str:
        str = (
            f"\n📁 Propriétés du fichier :\n"
            f"  Fichier : {self.chemin.name or 'non défini'}\n"
            f"  Répertoire : {self.chemin.parent or 'non défini'}\n"
            f"  Version : {self.version or 'non défini'}\n"
            f"  Date dernière modif windows : {(self.date_derniere_modification_windows or 'non défini') if self.chemin.exists() else "FICHIER N'EXISTE PAS"}\n"
        )

        return str




    # === GETTERS ===
    @property
    def date_derniere_modification_windows(self) -> str:
        """
        Extrait la date du fichier à partir de la dernière date de modif (valeur système windows)
        
        :return: la date du fichier à partir de la dernière date de modif (valeur système windows)
        :rtype: str
        """
        datetime_fichier = datetime.fromtimestamp(self.chemin.stat().st_mtime)
        return datetime_fichier.strftime("%d/%m/%Y")


    # === CONTRAT ===
    @classmethod
    @abstractmethod
    def est_compatible(cls, chemin: Path) -> bool:
        pass

    @classmethod
    @abstractmethod
    def version(self) -> str:
        """
        Version métier du fichier (ex: 'V6.1 du 17/01/2025' pour la FdC)
        """
        pass