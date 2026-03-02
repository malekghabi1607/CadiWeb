import pytest
from pathlib import Path

from vte.domain.iris import *

DATA_DIR = Path(__file__).parent / "data"


# Lancer les tests :
#    - pytest : tous les tests depuis la racine du projet
#    - pytest -v : plus verbeux
#    - ciblé : python -m pytest -v tests/vte/test_iris.py
#    - ciblé avec les print: python -m pytest -s -v tests/vte/test_iris.py
#    - ciblé et unitaire : python -m pytest -s -v tests/vte/test_iris.py::test_iris_natif_initialisation_et_chargement_df_un_fichier_input


# ---------- FIXTURES ----------

@pytest.fixture
def iris_natif_session():
    return {
        "chemin": DATA_DIR / r"R04110_Sessions-2024 FINAL (extract IRIS natif).xlsx",
        "chemins" : (
            DATA_DIR / r"R04110_Sessions-2024 FINAL (extract IRIS natif).xlsx",
            DATA_DIR / r"R04110_Sessions-2025 FINAL (extract IRIS natif).xlsx"
        ),
        "typeExport": "Sessions",
        "chemin_output" : Path(r"C:\Users\vt238770\Downloads\test.xlsx")
    }



# ---------- TESTS ----------
def test_iris_natif_initialisation(iris_natif_session):
    iris = IRIS_natif(
        typeExport=iris_natif_session["typeExport"]
        )
    #print("\n")
    #print(iris.typeExport) # = "Session"
    assert iris.typeExport == iris_natif_session["typeExport"]

def test_iris_natif_initialisation_et_chargement_df_un_fichier_input(iris_natif_session):
    iris = IRIS_natif(
        typeExport=iris_natif_session["typeExport"]
        )
    df, tuple_chemins = iris._charger_iris_natifs(iris_natif_session["chemin"])
    #print("\n")
    #print(df)
    #print(df["N° Session"].iloc[-1])
    assert df["N° Session"].loc[0] == "S-17804-FI2425-REN-0324-ADM-ABA"

def test_iris_natif_initialisation_et_chargement_df_plusieurs_fichiers_input(iris_natif_session):
    iris = IRIS_natif(
        typeExport=iris_natif_session["typeExport"]
        )
    df, tuple_chemins = iris._charger_iris_natifs(iris_natif_session["chemins"])

    tuple_in = tuple(Path(i).name for i in iris_natif_session["chemins"])
    tuple_out = tuple(i.name for i in tuple_chemins)
    #print("\n")
    #print(tuple_in)
    #print(tuple_out)

    assert tuple_in == tuple_out
    assert df["N° Session"].iloc[-1] == "S-15820-FC25-46C-DHE-ACD"

def test_iris_natif_avec_creationExport(iris_natif_session):
    iris = IRIS_natif.avec_traitement(
        typeExport=iris_natif_session["typeExport"],
        chemins_fichiersInput=iris_natif_session["chemins"],
        chemin_fichier_sauv=iris_natif_session["chemin_output"]
        )
    
    #print("\n")
    #print(iris.fe.chemin_fichier)

    assert iris.df["N° Session"].iloc[-1] == "S-15820-FC25-46C-DHE-ACD"
    assert iris.fe.chemin_fichier == iris_natif_session["chemin_output"]
    assert len(iris_natif_session["chemins"]) == len(iris.tableau_fichiers_importes.df)
