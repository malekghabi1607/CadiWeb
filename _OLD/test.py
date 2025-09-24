import re, os
from typing import Tuple, List

def recupere_trigFormation_fichier_chemin_depuis_chemin(chemin:Tuple(str)) -> str:
    # On récupère le trigramme de la formation depuis le chemin du CSV ou alors on demande à l'utilisateur via tkinter
    trigramme = None

    # Partie trigramme
    match = re.search(r"\\([a-zA-Z0-9]{3})\\", chemin)
    if match:
        trigramme = match.group(1)
    else:
        print("Erreur dans recupere_trigFormation_fichier_chemin_depuis_chemin : le chemin ne contient pas le trigramme de la formation.")
        trigramme = "Err"

    # Partie 


    return trigramme



def recupere_trigFormation_fichier_chemin_depuis_chemin(chemins: Tuple[str]) -> List[Tuple[str, str, str]]:
    resultats = []
    for chemin in chemins:
        
        # Partie trigramme
        match = re.search(r"\\([a-zA-Z0-9]{3})\\", chemin)
        if match:
            trigramme = match.group(1)
        else:
            print(f"Erreur dans recupere_trigFormation_fichier_chemin_depuis_chemin : le chemin ne contient pas le trigramme de la formation {chemin}")
            trigramme = "Err"
        # On rajoute une ligne de résultats
        resultats.append((trigramme, os.path.basename(chemin), chemin))
    return resultats
