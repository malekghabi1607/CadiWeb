"""
Configuration de secours EvalStat.

Le web detecte automatiquement les CSV stagiaires dans FORMATIONS_C.
Ce fichier sert seulement de secours si le lecteur reseau n'est pas accessible.
Ne pas ajouter les nouvelles sessions ici : il suffit de deposer le CSV dans le
dossier de la formation.
"""

codes_iris = [12995, 15697]

csv_stagiaires = [
   
    {
        "formation": "948",
        "code_iris": 12995,
        "enabled": True,
        "path": r"P:\FORMATIONS_C\948\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2022-11-S-12995 UEM\S-12995-FC22-948-VTE-SNA-Stagiaires.csv",
    },
    {
        "formation": "948",
        "code_iris": 15697,
        "enabled": True,
        "path": r"P:\FORMATIONS_C\948\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2024-01-S-15697 UEM\S-15697-FC24-948-VTE-VCA-Stagiaires.csv",
    },
    {
        "formation": "948",
        "code_iris": 17139,
        "enabled": True,
        "path": r"P:\FORMATIONS_C\948\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2025-12-S17139 UEM\S-17139-FC25-948-VTE-VCA-Stagiaires.csv",
    },
    {
        "formation": "TEL",
        "code_iris": 12766,
        "enabled": True,
        "path": r"P:\FORMATIONS_C\TEL\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\S-12766-rapports-session-evaluations\S-12766-FC22-TEL-JVI-LRA-Stagiaires.csv",
    },
    {
        "formation": "TEL",
        "code_iris": 13414,
        "enabled": True,
        "path": r"P:\FORMATIONS_C\TEL\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\S-13414-rapports-session-evaluations\S-13414-FC22-TEL-JVI-LRA-Stagiaires.csv",
    },
    {
        "formation": "TEL",
        "code_iris": 15942,
        "enabled": True,
        "path": r"P:\FORMATIONS_C\TEL\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2024 csv\S-15942-FC24-TEL-JVI-ACD-Stagiaires.csv",
    },
    {
        "formation": "TEL",
        "code_iris": 16161,
        "enabled": True,
        "path": r"P:\FORMATIONS_C\TEL\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2024 csv\S-16161-FC24-TEL-VTE-ACD-Stagiaires.csv",
    },
    {
        "formation": "TEL",
        "code_iris": 16411,
        "enabled": True,
        "path": r"P:\FORMATIONS_C\TEL\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\S-16411-rapports-session-evaluations\S-16411-FC25-TEL-VTE-CAR-Stagiaires.csv",
    },
    {
        "formation": "22B",
        "code_iris": 13363,
        "enabled": True,
        "path": r"P:\FORMATIONS_C\22B\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\S-13363-FC22-22B-JVI-LRA-Stagiaires.csv",
    },
    {
        "formation": "22B",
        "code_iris": 17606,
        "enabled": True,
        "path": r"P:\FORMATIONS_C\22B\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\S-17606 - 22B - 06-2025\S-17606-FC25-22B-VTE-CAR-Stagiaires.csv",
    },
]

# Compatibilite avec le menu console existant.
liste_codes_IRIS = codes_iris
tuple_csv_stagiaires = tuple(
    item["path"]
    for item in csv_stagiaires
    if item.get("enabled", True)
)
