from vte.instn import *

# TODO ⚠️ déplacer les bilans originaux en cas de tests !!! ⚠️

### --------------------------------------------------------------------
#  BilanSession
### --------------------------------------------------------------------

#####
# === Traiter bilans pour des sessions uniques (une seule session par bilan)
#####
# TODO ⚠️ déplacer les bilans originaux en cas de tests !!! ⚠️
### Cas 1 — Bilan unique avec code IRIS unique, bilan csv existant
def test_cas1() :
    BilanSession.bilanUnique_parCodeIRIS(16411)  # TEL octobre 2025 : bilan avec CSV
    
### Cas 2 — Bilan unique avec code IRIS unique, bilan csv inexistant
def test_cas2() :
    BilanSession.bilanUnique_parCodeIRIS(17343)  # TEL mars 2025 : bilan avec CSV manquant
    
### Cas 3 — Plusieurs bilans avec codes IRIS unique, bilan csv existant / inexistant
def test_cas3() :
    BilanSession.plusieursBilans_parCodeIRIS([17343, 16411])  # TEL mars et octobre 2025

### Cas 4 — Plusieurs bilans avec codes IRIS unique, via le fichier de config
def test_cas4() :
    BilanSession.plusieursBilans_parCodeIRIS_fichierConfig()



#####
# === Traiter bilans pour des périodes (plusieurs sessions par bilan)
#####
# TODO ⚠️ déplacer les bilans originaux en cas de tests !!! ⚠️
### Cas 5 — Période annuelle
def test_cas5():
    BilanSession.bilanUnique_parPeriode("948", 2024, "Année")

### Cas 6 — Période annuelle
def test_cas6():
    chargement_config_demander_verif_utilisateur(
    nom_variable = "liste_periodes",
    fonction_execution = BilanSession.plusieursBilans_parPeriode,
)

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
    #TODO ⚠️ déplacer les bilans originaux en cas de tests !!! ⚠️
    test_cas6()