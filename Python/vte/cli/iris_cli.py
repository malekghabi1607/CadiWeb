from vte.domain.iris import *

DATA_TEST_DIR = Path(r"C:\Users\vt238770\Documents\_CEA\Prog\Python\tests\vte\data")

# Données des tests                     
iris_natif_session = {
        "chemin": DATA_TEST_DIR / r"R04110_Sessions-2024 FINAL (extract IRIS natif).xlsx",
        "chemins" : (
            DATA_TEST_DIR / r"R04110_Sessions-2024 FINAL (extract IRIS natif).xlsx",
            DATA_TEST_DIR / r"R04110_Sessions-2025 FINAL (extract IRIS natif).xlsx"
        ),
        "typeExport": "Sessions",
        "chemin_output" : Path(r"C:\Users\vt238770\Downloads\test.xlsx")
    }



def creer_objet_IRIS_natif():
    iris = IRIS_natif(
        typeExport="Session",
        chemin=r"C:\Users\vt238770\Documents\_CEA\Prog\Python\tests\vte\data\R04110_Sessions-2024 FINAL (extract IRIS natif).xlsx")
    print(iris.typeExport) # = "Session"

def iris_natif_initialisation_et_chargement_df_plusieurs_fichiers_input(iris_natif_session):
    iris = IRIS_natif(
        typeExport=iris_natif_session["typeExport"]
        )
    df, tuple_chemins = iris._charger_iris_natifs(iris_natif_session["chemins"])

    tuple_in = tuple(Path(i).name for i in iris_natif_session["chemins"])
    tuple_out = tuple(i.name for i in tuple_chemins)
    print("\n")
    print(tuple_in)
    print(tuple_out)

def iris_natif_avec_creationExport(iris_natif_session):
    iris = IRIS_natif.avec_traitement(
        typeExport=iris_natif_session["typeExport"],
        chemins_fichiersInput=iris_natif_session["chemins"],
        chemin_fichier_sauv=iris_natif_session["chemin_output"]
        )
    
    print("\n")
    print(iris.fe.chemin_fichier)
    assert iris.df["N° Session"].iloc[-1] == "S-15820-FC25-46C-DHE-ACD"
    assert iris.fe.chemin_fichier == iris_natif_session["chemin_output"]

def iris_natif_avec_creationExport(iris_natif_session):
    iris = IRIS_natif.avec_traitement(
        typeExport=iris_natif_session["typeExport"],
        chemins_fichiersInput=iris_natif_session["chemins"],
        chemin_fichier_sauv=iris_natif_session["chemin_output"]
        )
    print(iris.df["N° Session"])
    print(iris.df["N° Session"].iloc[-1] == "S-15820-FC25-46C-DHE-ACD")
    print(iris.fe.chemin_fichier == iris_natif_session["chemin_output"])
    print(len(iris_natif_session["chemins"]) == len(iris.tableau_fichiers_importes.df))

def main():
    iris_natif_avec_creationExport(iris_natif_session)

if __name__ == "__main__":
    #test07()
    main()





"""
# Par ChatGPT
def charger_sessions_IRIS():
    chemin = choisir_fichier_IRIS("Sessions")
    iris = IRISReferentiel.get("Sessions", chemin)
    iris.charger()
    return iris


def maj_exports_IRIS(typesExports, depuis_config=False):
    IRISServices.concatener_plusieurs_types(typesExports, depuis_config)
"""