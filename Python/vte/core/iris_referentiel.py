from __future__ import annotations
from collections.abc import Iterable
from pathlib import Path
from typing import Dict, Optional

from vte.domain.iris import IRIS, IRIS_traite



# ======================================================================================
# REFERENTIEL IRIS (chargé une seule fois)
# Référentiel partagé, chargé une seule fois (configs, constantes lourdes)
# ======================================================================================


# ======================================================================================
# CACHE INTERNE
# ======================================================================================

_cache: Dict[str, IRIS_traite] = {}
_chemins_specifiques: Dict[str, Optional[Path]] = {}


# ======================================================================================
# API PUBLIQUE
# ======================================================================================

# === FONCTIONS DE BASE ===
def get_iris(typeExport: str, chemin:Optional[Path]=None) -> IRIS_traite:
    """
    Retourne une instance IRIS chargée pour un type donné.
    Si elle n'existe pas encore, elle est créée et mise en cache.

    :param typeExport: Type de l'export iris à considérer. Doit être dans cette liste : ["Sessions", "Formations", "Ventes", "Insciptions"]
    :type typeExport: str
    :param chemin: chemin de l'export IRIS. Si renseigné, alors on n'emploiera pas celui par défaut (plus récent dans le répertoire idoine). Défaut = None
    :type chemin: Optional[Path]
    :return: le fichier IRIS traité
    :rtype: IRIS_traite
    """
    # On affecte chemin si l'utilisateur en a spécifié un différent à un moment
    if chemin is None:
        chemin = _chemins_specifiques.get(typeExport)  # Si aucun chemin spécifique n'a été rentré, alors ça renverra None et on retrouvera le cas par défaut

    # On charge IRIS selon le type d'export s'il n'a pas déjà été chargé auparavant
    if typeExport not in _cache:
        iris = IRIS_traite(typeExport=typeExport, chemin=chemin)
        _cache[typeExport] = iris

    return _cache[typeExport]

def reload_iris(type_export: str, chemin:Optional[Path]=None) -> IRIS_traite:
    """
    Recharge le cache
    """
    _cache.pop(type_export, None)
    return get_iris(typeExport=type_export, chemin=chemin)

def clear_iris() -> None:
    """
    Vide le cache (utile pour tests).
    """
    _cache.clear()

def set_iris_chemin_specifique(typeExport: str, chemin: Path) -> None:
    """
    Affecte un chemin spécifique à un type d'export

    :param typeExport: Type de l'export iris à considérer. Doit être dans cette liste : ["Sessions", "Formations", "Ventes", "Insciptions"]
    :type typeExport: str
    :param chemin: Chemin spécifique à affecter
    :type chemin: Path
    """
    _chemins_specifiques[typeExport] = chemin



# === FONCTIONS DE BASE ===
def dico_trigrammes_codes_IRIS_a_partir_de_codes_IRIS(codes_IRIS:int|Iterable[int], chemin:Optional[Path]=None) -> dict[str:tuple(int)]:
    """
    A partir d'un itérable de codes IRIS, je crée un dictionnaire {trigrammes: tuple(codes_IRIS)}

    :param codes_IRIS: Codes IRIS dont il faut récupérer le trigramme
    :type codes_IRIS: int | Iterable[int]
    :param chemin: Si chemin spécifique pour IRIS sessions, défaut = None
    :type chemin: Optional[Path], optional
    :return: dictionnaire {trigrammes: tuple(codes_IRIS)}
    :rtype: dict[str:tuple(int)
    """
    if isinstance(codes_IRIS, int):
        codes_IRIS = (codes_IRIS, )
    
    df = get_iris(typeExport="Sessions", chemin=chemin).df  # Alias

    # Possible avant .tolist() : .unique() pour valeurs uniques ; .dropna() pour gérer valeurs manquantes
    #return df[df["Code IRIS"].isin(codes_IRIS)]["Trigramme formation"].unique().tolist()
    # TODO : il semble que ❌ Tu ignores codes_IRIS → tu prends TOUT
    # TODO FIX :
    # TODO df_filtre = df[df["Code IRIS"].isin(codes_IRIS)]
    # TODO return df_filtre.groupby("Trigramme formation")["Code IRIS"].apply(tuple).to_dict()
    #return df.groupby("Trigramme formation")["Code IRIS"].apply(tuple).to_dict()

    # Essai de correction
    df_filtre = df[df["Code IRIS"].isin(codes_IRIS)]
    return df_filtre.groupby("Trigramme formation")["Code IRIS"].apply(tuple).to_dict()




# Emploi :
#iris_sessions = get_iris("Sessions")
#df = iris_sessions.df

# Si chemin spécifique :
#set_iris_path("Sessions", chemin_test)