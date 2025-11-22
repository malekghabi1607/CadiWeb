from vte.instn import *


### --------------------------------------------------------------------
#  EvalStat
### --------------------------------------------------------------------

chemin_csv_session = Path(r"\\INSTNT\partage\FORMATIONS_C\948\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2022-11-S-12995 UEM\S-12995-FC22-948-VTE-SNA-Stagiaires.csv")
chemin_excel_session = chemin_csv_session.with_suffix(".xlsx")

# 948
chemin_excel_formation_948 = Path(r"\\INSTNT\partage\FORMATIONS_C\948\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\Evaluation-Stagiaires-Global-948.xlsx")
tuple_csv_stagiaires_948 = (
    r"\\INSTNT\partage\FORMATIONS_C\948\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2021-11-S-11369 UEM\EVALSTAT\S-11369-FC21-948-VLE-SNA-Stagiaires.csv",
    r"\\INSTNT\partage\FORMATIONS_C\948\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2022-11-S-12995 UEM\S-12995-FC22-948-VTE-SNA-Stagiaires.csv",
    r"\\INSTNT\partage\FORMATIONS_C\948\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2024-01-S-15697 UEM\S-15697-FC24-948-VTE-VCA-Stagiaires.csv"
    )
liste_codes_iris_948 = [11369, 12995, 15697]


# 22B
chemin_excel_formation_22B = Path(r"\\INSTNT\partage\FORMATIONS_C\22B\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\Evaluation-Stagiaires-Global-22B.xlsx")
tuple_csv_stagiaires_948_22B = (
    r"\\INSTNT\partage\FORMATIONS_C\948\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2021-11-S-11369 UEM\EVALSTAT\S-11369-FC21-948-VLE-SNA-Stagiaires.csv",
    r"\\INSTNT\partage\FORMATIONS_C\22B\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\S-13363-FC22-22B-JVI-LRA-Stagiaires.csv",
    r"\\INSTNT\partage\FORMATIONS_C\948\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2022-11-S-12995 UEM\S-12995-FC22-948-VTE-SNA-Stagiaires.csv",
    r"\\INSTNT\partage\FORMATIONS_C\22B\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\S-17606 - 22B - 06-2025\S-17606-FC25-22B-VTE-CAR-Stagiaires.csv",
    r"\\INSTNT\partage\FORMATIONS_C\948\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2024-01-S-15697 UEM\S-15697-FC24-948-VTE-VCA-Stagiaires.csv"
    )
liste_codes_iris_948_22B = [11369, 13363, 12995, 17606, 15697]

#####
# === Créer un EvalStat d'une session
#####

### Cas 1 — Si fichier session inexistant & excel eval formation inexistant
def test_cas1() :
    chemin_excel_session.unlink(missing_ok=True)  # Suppression excel session
    chemin_excel_formation_948.unlink(missing_ok=True)  # Suppression excel formation
    EvalStat.depuis_chemin_csv_evaluations_stagiaires(chemin_csv_session)
    # OK

### Cas 2 — Si fichier session existant & excel eval formation inexistant
def test_cas2() :
    #test_cas1()
    chemin_excel_formation_948.unlink(missing_ok=True)  # Suppression excel formation
    EvalStat.depuis_chemin_csv_evaluations_stagiaires(chemin_csv_session)
    # OK


### Cas 3 — Si fichier session existant & excel eval formation existant
def test_cas3() :
    #test_cas1()
    EvalStat.depuis_chemin_csv_evaluations_stagiaires(chemin_csv_session)




# === Depuis Tuple stagiaires ===
### Cas 1 — Même trigramme : si fichiers sessions inexistants & excel eval formation inexistant
def test_cas4() :
    # Suppression excels sessions
    for chemin_csv_session in tuple_csv_stagiaires_948 :
        chemin_csv_session = Path(chemin_csv_session)
        chemin_excel_session = chemin_csv_session.with_suffix(".xlsx")
        chemin_excel_session.unlink(missing_ok=True)  

    chemin_excel_formation_948.unlink(missing_ok=True)  # Suppression excel formation

    EvalStat.depuis_tuple_csv_stagiaires(tuple_csv_stagiaires=tuple_csv_stagiaires_948)
    
### Cas 2 — Même trigramme : si fichiers sessions existants & excel eval formation existant
def test_cas5() :

    EvalStat.depuis_tuple_csv_stagiaires(tuple_csv_stagiaires=tuple_csv_stagiaires_948)
    
### Cas 3 — Trigrammes différents, ordre random : si fichiers sessions inexistants & excel eval formation inexistant
def test_cas6() :
    # Suppression excels sessions
    for chemin_csv_session in tuple_csv_stagiaires_948_22B :
        chemin_csv_session = Path(chemin_csv_session)
        chemin_excel_session = chemin_csv_session.with_suffix(".xlsx")
        chemin_excel_session.unlink(missing_ok=True)  

    chemin_excel_formation_948.unlink(missing_ok=True)  # Suppression excel formation
    chemin_excel_formation_22B.unlink(missing_ok=True)  # Suppression excel formation

    EvalStat.depuis_tuple_csv_stagiaires(tuple_csv_stagiaires=tuple_csv_stagiaires_948_22B)



# === Depuis Tuple stagiaires ===
# Cas 1 — Trigrammes différents, ordre random : si fichiers sessions inexistants & excel eval formation inexistant
def test_cas7():
    """
    On a déjà fait des tests avant : on reprend le cas 6 mais on ne delete que les formation 948 (i.e. tuple) 
    """
    # Suppression excels sessions
    for chemin_csv_session in tuple_csv_stagiaires_948 :
        chemin_csv_session = Path(chemin_csv_session)
        chemin_excel_session = chemin_csv_session.with_suffix(".xlsx")
        chemin_excel_session.unlink(missing_ok=True)  
    
    chemin_excel_formation_948.unlink(missing_ok=True)  # Suppression excel formation

    EvalStat.depuis_liste_codes_IRIS(liste_codes_IRIS=liste_codes_iris_948_22B)

# === Popup ===

if __name__ == "__main__":
    test_cas6()










# === Fichier CSV individuel
#es = Traiter_evalStat()
#es = Traiter_evalStat.depuis_chemin_csv_stagiaires(r"\\instnt\partage\FORMATIONS_C\22B\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\S-17606 - 22B - 06-2025\S-17606-FC25-22B-VTE-CAR-Stagiaires.csv")
#es = Traiter_evalStat.depuis_chemin_csv_stagiaires(r"P:\FORMATIONS_C\778\P07-bilan-sessions-et-bilan-formation\2024-Bilans 778\Evaluations 778(2024.11)\S-15715-FC24-778-VMO-VCA-Stagiaires.csv")
#es = Traiter_evalStat.depuis_chemin_csv_stagiaires(r"\\INSTNT\partage\FORMATIONS_C\948\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2024-01-S-15697 UEM\S-15697-FC24-948-VTE-VCA-Stagiaires.csv")

"""
tuple_csv_stagiaires = (
    r"\\INSTNT\partage\FORMATIONS_C\948\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2021-11-S-11369 UEM\EVALSTAT\S-11369-FC21-948-VLE-SNA-Stagiaires.csv",
    r"\\INSTNT\partage\FORMATIONS_C\948\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2022-11-S-12995 UEM\S-12995-FC22-948-VTE-SNA-Stagiaires.csv",
    r"\\INSTNT\partage\FORMATIONS_C\948\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2024-01-S-15697 UEM\S-15697-FC24-948-VTE-VCA-Stagiaires.csv"
    )
"""

#es = Traiter_evalStat.depuis_tuple_csv_stagiaires(tuple_csv_stagiaires=tuple_csv_stagiaires)








