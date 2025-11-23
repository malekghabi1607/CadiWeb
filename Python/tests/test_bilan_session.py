from vte.instn import *


### --------------------------------------------------------------------
#  BilanSession
### --------------------------------------------------------------------

#####
# === Traiter bilans pour des sessions uniques (une seule session par bilan)
#####
### Cas 1 — Bilan unique avec code IRIS unique, bilan csv existant
def test_cas1() :
    BilanSession.bilanUnique_parCodeIRIS(16411)  # TEL octobre 2025 : bilan avec CSV
    
### Cas 2 — Bilan unique avec code IRIS unique, bilan csv inexistant
def test_cas2() :
    BilanSession.bilanUnique_parCodeIRIS(17343)  # TEL mars 2025 : bilan avec CSV manquant

""" 
BilanSession.plusieursBilans_parCodeIRIS

chargement_config_demander_verif_utilisateur
BilanSession.plusieursBilans_parCodeIRIS_fichierConfig

BilanSession.bilanUnique_parPeriode
BilanSession.bilanUnique_parPeriode("948", 2024, "Année")

chargement_config_demander_verif_utilisateur
BilanSession.plusieursBilans_parPeriode
"""


if __name__ == "__main__":
    test_cas2()