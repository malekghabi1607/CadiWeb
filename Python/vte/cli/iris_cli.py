from vte.services.iris_services import *


# ======================================================================================
# === TRAITEMENT tous les IRIS
# ======================================================================================
def traite_tous_extract_IRIS_depuis_config():
    IRIS_services.concatener_plusieursTypes(
        #typesExports=("Sessions", "Formations", "Ventes", "Inscriptions"),
        typesExports=("Sessions"),
        depuis_config=True
    )

# ======================================================================================
# === MAIN
# ======================================================================================

def main():
    #iris_traite_methode_chemin_IRIS_plus_recent(de_iris_traite_session)
    traite_tous_extract_IRIS_depuis_config()

    #IRIS_natif.avec_traitement("Sessions")

if __name__ == "__main__":
    #test07()
    main()