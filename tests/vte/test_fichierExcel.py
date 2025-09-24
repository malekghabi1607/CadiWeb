from vte.office import *  # <-- adapte avec ton vrai nom de module

import pytest
from pathlib import Path

DATA_DIR = Path(__file__).parent / "data"

# ---------- FIXTURES ----------

@pytest.fixture
def excel_complexe():
    return {
        "chemin": DATA_DIR / r"S-15697-FC24-948-VTE-VCA-Stagiaires.xlsx",
        "nom_onglet": "Stagiaires",
        "nb_tableaux": 3
    }
    
@pytest.fixture
def excel_nonStruct():
    return {
        "chemin": DATA_DIR / r"Tableau_nonStruct.xlsx",
        "nom_onglet": "Feuil1",
        "nbLignes_avantET": 2,
        "ref": "A3:C5",
        "columns": ['A', 'B', 'Formule']
    }

@pytest.fixture
def excel_extractIRIS():
    return {
        "chemin": DATA_DIR / r"R04110_Sessions-2024 FINAL (extract IRIS natif).xlsx",
        "nom_onglet": "Data",
        "nbLignes_avantET": 1,
        "ref": "A2:AU1329",
        "nom_last_col": "Nb. Absents"
    }

 #fe = FichierExcel.depuis_fichier(r"C:\Users\vt238770\Documents\_CEA\_Formations\Extracts IRIS - Faits\Extracts originaux\R04110_Sessions-2015 FINAL.xlsx")
    #fe.charger_tableau("Data", nbLignes_avantET=1)
    #print(fe._tableaux["Data"])
    #print(fe)

# ---------- TESTS ----------

def test_depuis_fichier_tousTableauxStructures(excel_complexe):
    fe = FichierExcel.depuis_fichier(excel_complexe["chemin"])
    print(fe)
    assert fe.nom_fichier == os.path.basename(excel_complexe["chemin"])
    assert fe.wb is not None
    assert fe.get_tableau(excel_complexe["nom_onglet"]).nom_tableau == excel_complexe["nom_onglet"]
    assert len(fe.tableaux) == excel_complexe["nb_tableaux"]


def test_depuis_fichier_tableauSpecifique(excel_complexe):
    fe = FichierExcel.depuis_fichier(excel_complexe["chemin"], nom_onglet=excel_complexe["nom_onglet"])
    print(fe)
    assert fe.nom_fichier == os.path.basename(excel_complexe["chemin"])
    assert fe.wb is not None
    assert fe.get_tableau(excel_complexe["nom_onglet"]).nom_tableau == excel_complexe["nom_onglet"]
    assert len(fe.tableaux) == 1


def test_depuis_fichier_tableauSpecifique_nonStruct(excel_nonStruct):
    fe = FichierExcel.depuis_fichier(excel_nonStruct["chemin"], nom_onglet=excel_nonStruct["nom_onglet"], nbLignes_avantET=excel_nonStruct["nbLignes_avantET"])
    #print(fe)

    tab = fe.get_tableau(excel_nonStruct["nom_onglet"])
    print(tab)
    print(tab.df.columns)
    
    assert fe.nom_fichier == os.path.basename(excel_nonStruct["chemin"])
    assert fe.wb is not None
    assert len(fe.tableaux) == 1

    assert tab.nom_tableau == excel_nonStruct["nom_onglet"]
    assert tab.ref_tableau == excel_nonStruct["ref"]
    assert tab.df.columns.tolist() == excel_nonStruct["columns"]


def test_depuis_fichier_extractIRIS(excel_extractIRIS):
    fe = FichierExcel.depuis_fichier(excel_extractIRIS["chemin"], nom_onglet=excel_extractIRIS["nom_onglet"], nbLignes_avantET=excel_extractIRIS["nbLignes_avantET"])
    #print(fe)

    tab = fe.get_tableau(excel_extractIRIS["nom_onglet"])
    print(tab)
    print(tab.df.columns)
    
    assert fe.nom_fichier == os.path.basename(excel_extractIRIS["chemin"])
    assert fe.wb is not None
    assert len(fe.tableaux) == 1

    assert tab.nom_tableau == excel_extractIRIS["nom_onglet"]
    assert tab.df.columns[-1] == excel_extractIRIS["nom_last_col"]
    assert tab.ref_tableau == excel_extractIRIS["ref"]

"""
def test_ecriture_dataframe_dans_tableau(fichier_excel_pandas, tmp_path):
    fx = FichierExcel.depuis_fichier(fichier_excel_pandas)
    tableau = fx.ajouter_tableau_normal("Sheet1", nb_lignes_avant_et=0)

    new_df = pd.DataFrame({"col1": [10, 20], "col2": ["x", "y"]})
    tableau.remplace_df(new_df)

    # Sauvegarde et recharge pour vérifier persistance
    new_path = tmp_path / "out.xlsx"
    fx.save(new_path)

    fx2 = FichierExcel.depuis_fichier(new_path)
    tableau2 = fx2.get_tableau(list(fx2.tableaux.keys())[0])
    assert tableau2._df.equals(new_df)

    """
