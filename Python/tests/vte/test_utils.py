import pytest

from vte.utils import *

# ---------- TESTS ----------

"""def test_optimiseCheminRepertoire(path_in:str):
    
    assert fx.nom_fichier == "test_pandas.xlsx"
    assert fx.wb is not None"""

@pytest.mark.skipif(os.name != "nt", reason="Windows only")
def test_chemin_vers_unc_local_path():
    # chemin local standard
    path_input = r"P:\FORMATIONS_C"
    result = chemin_vers_unc(path_input)
    print(result)
    # Comme il n'y a pas de lecteur réseau réel, la fonction devrait retourner le path normalisé
    assert os.path.normpath(path_input) == result or result.startswith("\\\\")  # UNC éventuel

def test_optimiseCheminRepertoire():
    chemin = r"C:\Windows\System32\drivers\etc\salut\toto\maman"
    result = optimiseCheminRepertoire(chemin)
    print(result)
    assert r"C:\Windows\System32\drivers\etc" == result


# ============================================================
# === TESTS MANUELS (rendu graphique, interaction, etc.) ===
# ============================================================

# Lancer avec python -m pytest -m manual
# python -m pytest -k test_choisir_fichier_manuel

@pytest.mark.manual
def test_choisir_fichier_manuel():
    """
    Test manuel : permet de vérifier visuellement la fenêtre de sélection
    et les boutons de dialogue.
    """

    choisir_fichier(titre="Test manuel - Sélection de fichier")