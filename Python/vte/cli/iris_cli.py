from vte.domain.iris import *
from vte.services.iris_services import *


# ======================================================================================
# === TESTS
# ======================================================================================

DATA_TEST_DIR = Path(r"C:\Users\vt238770\Documents\_CEA\Prog\Python\tests\vte\data")

# Données des tests                     
de_iris_natif_session = {
        "chemin": DATA_TEST_DIR / r"R04110_Sessions-2024 FINAL (extract IRIS natif).xlsx",
        "chemins" : (
            DATA_TEST_DIR / r"R04110_Sessions-2024 FINAL (extract IRIS natif).xlsx",
            DATA_TEST_DIR / r"R04110_Sessions-2025 FINAL (extract IRIS natif).xlsx"
        ),
        "typeExport": "Sessions",
        "chemin_output" : Path(r"C:\Users\vt238770\Downloads\test.xlsx")
    }

de_iris_traite_session = {
        # Données d'entrée
        "chemin": DATA_TEST_DIR / r"R04110_Sessions-COMPLET-(IRIS traité).xlsx",
        "typeExport": "Sessions",
        "fe": FichierExcel.depuis_fichier(chemin_fichier=DATA_TEST_DIR / r"R04110_Sessions-COMPLET-(IRIS traité).xlsx"),
        "fe_vide": FichierExcel(),   # Fichier Excel vide

        # Données de sortie
        "Dernier N° Session": "S-18574-FI2527-2512-QPR-CEC-JPE",
        "Nb fichiers import": 2
    }

# === IRIS NATIFS ===
def creer_objet_IRIS_natif():
    iris = IRIS_natif(
        typeExport="Session",
        chemin=r"C:\Users\vt238770\Documents\_CEA\Prog\Python\tests\vte\data\R04110_Sessions-2024 FINAL (extract IRIS natif).xlsx")
    print(iris.typeExport) # = "Session"

def iris_natif_initialisation_et_chargement_df_plusieurs_fichiers_input(de_iris_natif_session):
    iris = IRIS_natif(
        typeExport=de_iris_natif_session["typeExport"]
        )
    df, tuple_chemins = iris._charger_iris_natifs(de_iris_natif_session["chemins"])

    tuple_in = tuple(Path(i).name for i in de_iris_natif_session["chemins"])
    tuple_out = tuple(i.name for i in tuple_chemins)
    print("\n")
    print(tuple_in)
    print(tuple_out)

def iris_natif_avec_creationExport(de_iris_natif_session):
    iris = IRIS_natif.avec_traitement(
        typeExport=de_iris_natif_session["typeExport"],
        chemins_fichiersInput=de_iris_natif_session["chemins"],
        chemin_fichier_sauv=de_iris_natif_session["chemin_output"]
        )
    print(iris.df["N° Session"])
    print(iris.df["N° Session"].iloc[-1] == "S-15820-FC25-46C-DHE-ACD")
    print(iris.fe.chemin_fichier == de_iris_natif_session["chemin_output"])
    print(len(de_iris_natif_session["chemins"]) == len(iris.tableau_fichiers_importes.df))

def iris_natif_avec_creationExport_sans_fichiersInput(de_de_iris_natif_session):
    iris = IRIS_natif.avec_traitement(
        typeExport=de_iris_natif_session["typeExport"],
        chemin_fichier_sauv=de_iris_natif_session["chemin_output"]
        )
    print(iris.df["N° Session"])
    print(iris.df["N° Session"].iloc[-1] == "S-15820-FC25-46C-DHE-ACD")
    print(iris.fe.chemin_fichier == de_iris_natif_session["chemin_output"])



# === IRIS NATIFS ===
def iris_traite_methode_chemin_IRIS_plus_recent(de_iris_traite_session):
    iris = IRIS_traite(
    typeExport=de_iris_traite_session["typeExport"],
    fe=de_iris_traite_session["fe_vide"]
    )
    print(iris._chemin_IRIS_traite_plus_recent())

def iris_traite_charge_avec_fe_vide(de_iris_traite_session):
    iris = IRIS_traite(
        typeExport=de_iris_traite_session["typeExport"],
        fe=de_iris_traite_session["fe_vide"]
    )

    print(iris.fe)
    print(iris.fe.chemin_fichier)  # = None
    print(iris.fe.chemin_fichier is None)

def iris_traite_charge_avec_fe(de_iris_traite_session):
    iris = IRIS_traite(
        typeExport=de_iris_traite_session["typeExport"],
        fe=de_iris_traite_session["fe"]
    )

    print(iris.fe)
    print(iris.fe.chemin_fichier)  # = de_iris_traite_session["chemin"]
    print(iris.fe.chemin_fichier == de_iris_traite_session["chemin"])

def iris_traite_charge_avec_chemin(de_iris_traite_session):
    iris = IRIS_traite(
        typeExport=de_iris_traite_session["typeExport"],
        chemin=de_iris_traite_session["chemin"]
    )

    print(iris.fe)
    print(iris.fe.chemin_fichier)  # = de_iris_traite_session["chemin"]
    print(iris.fe.chemin_fichier == de_iris_traite_session["chemin"])  # OK
    print()
    print(iris.df["N° Session"])
    print(iris.df["N° Session"].iloc[-1] == de_iris_traite_session["Dernier N° Session"])
    print(iris.fe.chemin_fichier == de_iris_traite_session["chemin"])
    print(len(iris.tableau_fichiers_importes.df) == de_iris_traite_session["Nb fichiers import"])

def iris_traite_charge_sans_chemin(de_iris_traite_session):
        iris = IRIS_traite(
            typeExport=de_iris_traite_session["typeExport"]
        )
        print(iris.fe)
        print(iris.fe.chemin_fichier)  # = de_iris_traite_session["chemin"]
        print(iris.fe.chemin_fichier == de_iris_traite_session["chemin"])  # OK
        print()
        print(iris.df["N° Session"])
        print(iris.df["N° Session"].iloc[-1] == de_iris_traite_session["Dernier N° Session"])
        print(iris.fe.chemin_fichier == de_iris_traite_session["chemin"])
        print(len(iris.tableau_fichiers_importes.df) == de_iris_traite_session["Nb fichiers import"])


# ======================================================================================
# === TRAITEMENT tous les IRIS
# ======================================================================================
def traite_tous_extract_IRIS_depuis_config():
    IRIS_services.concatener_plusieursTypes(
        typesExports=("Sessions", "Formations", "Ventes", "Inscriptions"),
        depuis_config=True
    )

# ======================================================================================
# === TESTS
# ======================================================================================

def main():
    #iris_traite_methode_chemin_IRIS_plus_recent(de_iris_traite_session)
    traite_tous_extract_IRIS_depuis_config()

    #IRIS_natif.avec_traitement("Sessions")

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