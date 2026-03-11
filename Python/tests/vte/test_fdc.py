from vte.domain.formation import Formation
from vte.domain.fdc import *

import pytest

# Répertoire des fichiers pour les tests
DATA_DIR = Path(__file__).parent / "data"



"""
# Données constantes des tests
chemin_fdc= DATA_DIR / r"Fiche de coûts V6.1 - TEL.xlsx"
nom_onglet = "Fiche de coûts"
trigramme_formation = "TEL"


formation = Formation(trigramme_formation)
fe_avecChemin = FichierExcel(chemin_fichier=chemin_fdc)
fe_complet = FichierExcel.depuis_fichier(chemin_fichier=chemin_fdc, nom_onglet=nom_onglet)

# Cas 1 - OK mais relou car popup
#fdc = FdC(formation=formation)
#print(fdc.min_participants_cea)


# Cas 2
# fdc = FdC(formation=formation, fe=fe_avecChemin)
# print(fdc.min_participants_cea)


# Cas 3
# fdc = FdC(formation=formation, fe=fe_complet)
# print(fdc.min_participants_cea)

"""





# ---------- FIXTURES (données d'entrée aux tests) ----------
@pytest.fixture
def chemin_fdc():
    return Path(
        DATA_DIR / r"Fiche de coûts V6.1 - TEL.xlsx"
    )

@pytest.fixture
def fe_factory(chemin_fdc):
    """
    Liste de fichiers Excel à tester.
    Chaque fe_factory représente un cas réel rencontré.
    """

    return [
        None,
        FichierExcel(chemin_fichier=chemin_fdc),
        FichierExcel.depuis_fichier(
            chemin_fichier=chemin_fdc,
            nom_onglet="Fiche de coûts"
        ),
    ]

# ---------- TESTS ----------
"""
# Test unitaire
def test_avec_chemin_fdc(formation_tel, chemin_fdc):
    fe = FichierExcel(chemin_fichier=chemin_fdc)
    fdc = FdC(formation=formation_tel, fe=fe)

    assert fdc.min_participants_cea == 3

"""

# Test paramétré (même test mais avec arguments différents) :
@pytest.mark.parametrize("ife", range(3))
def test_constructeur(fe_factory, tel, ife):
    formation = Formation(tel["trigramme_formation"])  # Récupéré depuis conftest
    fe = fe_factory[ife]
    #print(fe)
    fdc = FdC(formation=formation, fe=fe)

    assert fdc.min_participants_cea == 3


"""
C'était une V1 avec fe_factory décalré directement
@pytest.mark.parametrize(
    "fe_factory",
    [
        lambda chemin: None,
        lambda chemin: FichierExcel(chemin_fichier=chemin),
        lambda chemin: FichierExcel.depuis_fichier(
            chemin_fichier=chemin,
            nom_onglet="Fiche de coûts"
        ),
    ],
)
def test_constructeur(fe_factory, formation_tel, chemin_fdc):
    fe = fe_factory(chemin_fdc)
    fdc = FdC(formation=formation_tel, fe=fe)

    assert fdc.min_participants_cea == 3
"""

# Lancer les tests :
#    - pytest : tous les tests depuis la racine du projet
#    - pytest -v : plus verbeux
#    - cyblé : python -m pytest -v tests/vte/test_fdc.py






