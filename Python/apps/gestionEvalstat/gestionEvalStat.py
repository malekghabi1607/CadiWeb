from vte.instn import *
#from vte.ihm_console import IHM_console
#from vte.ihm_tkinter import IHMTkinter

def main():
    # === Importer données d'un seul fichier IRIS dans un seul Dataframe
    #es = EvalStat.depuis_chemin_csv_stagiaires(r"\\instnt\partage\FORMATIONS_C\22B\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\S-17606 - 22B - 06-2025\S-17606-FC25-22B-VTE-CAR-Stagiaires.csv")  # OK



    # === Importer un ou plusieurs fichier(s) IRIS dans un seul Dataframe : pour chaque trigramme on va traiter chaque session et soit créer soit append le fichier excel global de la formation
    #tuple_csv_stagiaires = (
    #    r"\\INSTNT\partage\FORMATIONS_C\948\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2021-11-S-11369 UEM\EVALSTAT\S-11369-FC21-948-VLE-SNA-Stagiaires.csv",
    #    r"\\INSTNT\partage\FORMATIONS_C\948\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2022-11-S-12995 UEM\S-12995-FC22-948-VTE-SNA-Stagiaires.csv",
    #    r"\\INSTNT\partage\FORMATIONS_C\948\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2024-01-S-15697 UEM\S-15697-FC24-948-VTE-VCA-Stagiaires.csv"
    #    )

    #es = EvalStat.depuis_tuple_csv_stagiaires(tuple_csv_stagiaires=tuple_csv_stagiaires)  # OK

    print("")




def test():
    # === Importer données d'un seul fichier IRIS dans un seul Dataframe
    #es = EvalStat.depuis_chemin_csv_stagiaires(r"\\instnt\partage\FORMATIONS_C\22B\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\S-17606 - 22B - 06-2025\S-17606-FC25-22B-VTE-CAR-Stagiaires.csv")  # OK

    # === Importer un ou plusieurs fichier(s) IRIS dans un seul Dataframe : pour chaque trigramme on va traiter chaque session et soit créer soit append le fichier excel global de la formation
    tuple_csv_stagiaires = (
        r"\\INSTNT\partage\FORMATIONS_C\948\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2021-11-S-11369 UEM\EVALSTAT\S-11369-FC21-948-VLE-SNA-Stagiaires.csv",
        r"\\INSTNT\partage\FORMATIONS_C\948\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2022-11-S-12995 UEM\S-12995-FC22-948-VTE-SNA-Stagiaires.csv",
        r"\\INSTNT\partage\FORMATIONS_C\948\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2024-01-S-15697 UEM\S-15697-FC24-948-VTE-VCA-Stagiaires.csv"
        )

    es = EvalStat.depuis_tuple_csv_stagiaires(tuple_csv_stagiaires=tuple_csv_stagiaires)
    print()



    

if __name__ == "__main__":
    test()
    #main()




### --------------------------------------------------------------------
#  EvalStat
### --------------------------------------------------------------------

# === Fichier CSV individuel
#es = Traiter_evalStat()
#es = Traiter_evalStat.depuis_chemin_csv_stagiaires(r"\\instnt\partage\FORMATIONS_C\22B\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\S-17606 - 22B - 06-2025\S-17606-FC25-22B-VTE-CAR-Stagiaires.csv")
#es = Traiter_evalStat.depuis_chemin_csv_stagiaires(r"P:\FORMATIONS_C\778\P07-bilan-sessions-et-bilan-formation\2024-Bilans 778\Evaluations 778(2024.11)\S-15715-FC24-778-VMO-VCA-Stagiaires.csv")
#es = Traiter_evalStat.depuis_chemin_csv_stagiaires(r"\\INSTNT\partage\FORMATIONS_C\948\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2024-01-S-15697 UEM\S-15697-FC24-948-VTE-VCA-Stagiaires.csv")

tuple_csv_stagiaires = (
    r"\\INSTNT\partage\FORMATIONS_C\948\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2021-11-S-11369 UEM\EVALSTAT\S-11369-FC21-948-VLE-SNA-Stagiaires.csv",
    r"\\INSTNT\partage\FORMATIONS_C\948\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2022-11-S-12995 UEM\S-12995-FC22-948-VTE-SNA-Stagiaires.csv",
    r"\\INSTNT\partage\FORMATIONS_C\948\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2024-01-S-15697 UEM\S-15697-FC24-948-VTE-VCA-Stagiaires.csv"
    )

#es = Traiter_evalStat.depuis_tuple_csv_stagiaires(tuple_csv_stagiaires=tuple_csv_stagiaires)





