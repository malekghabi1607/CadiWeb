import pytest
import sys
from pathlib import Path
from pprint import *

from vte.domain.formation import Formation
from vte.domain.bilanSession import BilanSession
from vte.utils.utils import backup_fichier_test, restore_nom_fichier_test

sys.path.append(str(Path(__file__).resolve().parents[1]))
from conftest import tel_data

def verif_initial_bilanSession():
    # Backups de mon environnement de travail
    #backup_fichier_test(tel["chemin_eval_session"])  # Eval session

    tel = tel_data()
    formation = Formation(tel["trigramme_formation"])
    
    BilanSession.depuis_codeIRIS(formation=formation, code_IRIS=tel["code_IRIS"], chemin_IRIS_sessions=tel["chemin_IRIS_sessions"])

    
    # Restauration de mon environnement de travail
    #restore_nom_fichier_test(tel["chemin_eval_session"])

def main():
    verif_initial_bilanSession()

if __name__ == "__main__":
    main()

