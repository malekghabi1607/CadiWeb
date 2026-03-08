"""
Mode op

Vous pouvez soit traiter des EvalStat :
   - soit par session en indiquant des codes IRIS à 5 chiffres (rensigner liste_codes_IRIS)
   - soit par chemins de CSV (renseigner tuple_csv_stagiaires)
"""

liste_codes_IRIS = [12995, 15697]

# 948 (tous)
tuple_csv_stagiaires_948 = (
    r"P:\FORMATIONS_C\948\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2021-11-S-11369 UEM\EVALSTAT\S-11369-FC21-948-VLE-SNA-Stagiaires.csv",
    r"P:\FORMATIONS_C\948\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2022-11-S-12995 UEM\S-12995-FC22-948-VTE-SNA-Stagiaires.csv",
    r"P:\FORMATIONS_C\948\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2024-01-S-15697 UEM\S-15697-FC24-948-VTE-VCA-Stagiaires.csv",
    # 15717 - Pas de CSV stagiaire (bug QR code mais tour de table)
    r"P:\FORMATIONS_C\948\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2025-12-S17139 UEM\S-17139-FC25-948-VTE-VCA-Stagiaires.csv",
)

# TEL (tous)
tuple_csv_stagiaires_TEL = (
    # 11090 - Pas de csv stagiaires (mais pdf existe : à sortir si besoin)
    r"P:\FORMATIONS_C\TEL\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\S-12766-rapports-session-evaluations\S-12766-FC22-TEL-JVI-LRA-Stagiaires.csv",
    # 12768 - Pas de csv stagiaire (mais pdf existe : à sortir si besoin)
    r"P:\FORMATIONS_C\TEL\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\S-13414-rapports-session-evaluations\S-13414-FC22-TEL-JVI-LRA-Stagiaires.csv",
    r"P:\FORMATIONS_C\TEL\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2024 csv\S-15942-FC24-TEL-JVI-ACD-Stagiaires.csv",
    r"P:\FORMATIONS_C\TEL\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2024 csv\S-16161-FC24-TEL-VTE-ACD-Stagiaires.csv",
    r"P:\FORMATIONS_C\TEL\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\S-16411-rapports-session-evaluations\S-16411-FC25-TEL-VTE-CAR-Stagiaires.csv",
)

# 22B (tous)
tuple_csv_stagiaires_22B = (
    r"P:\FORMATIONS_C\22B\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\S-13363-FC22-22B-JVI-LRA-Stagiaires.csv",
    r"P:\FORMATIONS_C\22B\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\S-17606 - 22B - 06-2025\S-17606-FC25-22B-VTE-CAR-Stagiaires.csv",
)




tuple_csv_stagiaires = tuple_csv_stagiaires_948 + tuple_csv_stagiaires_TEL + tuple_csv_stagiaires_22B