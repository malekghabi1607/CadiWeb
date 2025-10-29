###############
# Utilisation #
###############
"""
Toute ligne commençant par un slash (#) n'est pas exécutée.

Vous pouvez soit traiter des bilans de session :
   - soit par session (en indiquant un code IRIS à 5 chiffres)
   - soit par période (1er semestre, 2nd semestre ou annuel)

Pour chacun de ces cas, vous pouvez :
   - soit les faire un par un ;

"""





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