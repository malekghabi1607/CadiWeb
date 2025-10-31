"""
Mode op

Vous pouvez soit traiter des bilans de session :
   - soit par session [liste_codes_IRIS] en indiquant des codes IRIS à 5 chiffres
   - soit par période (1er semestre, 2nd semestre ou annuel) en indiquant :
        ¤ un trigramme formation en majuscule (ex. : TEL)
        ¤ une année à 4 chiffres (ex. : 2025)
        ¤ une periode = [1er semestre, 2nd semestre, Année] (ex. : 2nd semestre)

Pour chacun de ces cas, vous pouvez :
   - soit les faire un par un ;

"""

liste_codes_IRIS = [16411, 17343, 17173]

liste_periodes = [
    ("TEL", 2023, "Année"),
    ("TEL", 2024, "1er semestre"),
    ("TEL", 2024, "2nd semestre"),
    ("TEL", 2025, "1er semestre"),
]



####
# PAR PERIODE
####
# periode = ["1er semestre", "2nd semestre", "Année"]

# === Bilan unique : ok ===
#BilanSession.bilanUnique("948", 2024, "Année")

# === Plusieurs bilans : ok ===
#bilans_a_traiter = [
#    ("TEL", 2022, "Année"),
#    ("TEL", 2023, "Année"),
#    ("TEL", 2024, "Année"),
#]
#BilanSession.plusieursBilans(bilans_a_traiter)

# === Bilan par code IRIS ===
#BilanSession.bilanUnique_parCodeIRIS(17343)
#BilanSession.bilanUnique_parCodeIRIS(16411)

#BilanSession.plusieursBilans_parCodeIRIS([17343, 16411])