"""
Mode op

Permet de lister les chemins des fichiers IRIS à concaténer dans le fichier global.
Les sessions et formations restent renseignées manuellement.
Les ventes et inscriptions peuvent être générées automatiquement depuis le dossier
des extracts originaux après la mise à jour GED.

"""

from vte.core.config import REPERTOIRE_EXTRACT_IRIS_LOCAL
from vte.services.iris_dumps_services import construire_liste_fichiers_type


_tSessions = (
    'R04110_Sessions-2011 à 2014 FINAL.xlsx',
    'R04110_Sessions-2015 FINAL.xlsx',
    'R04110_Sessions-2016 FINAL.xlsx',
    'R04110_Sessions-2017 FINAL.xlsx',
    'R04110_Sessions-2018 FINAL.xlsx',
    'R04110_Sessions-2019 FINAL.xlsx',
    'R04110_Sessions-2020 FINAL.xlsx',
    'R04110_Sessions-2021 FINAL.xlsx',
    'R04110_Sessions-2022 FINAL.xlsx',
    'R04110_Sessions-2023 FINAL.xlsx',
    'R04110_Sessions-2024 FINAL.xlsx',
    'R04110_Sessions-2025 FINAL.xlsx',
    'R04110_Sessions-2026 au 2026.05.20.xlsx',
)

_tFormations = (
    "R0304_Ref_Formation-Listedesformations-2026.05.20.xlsx",
)

_tVentes = construire_liste_fichiers_type(
    REPERTOIRE_EXTRACT_IRIS_LOCAL,
    "R04301",
)

_tInscriptions = construire_liste_fichiers_type(
    REPERTOIRE_EXTRACT_IRIS_LOCAL,
    "R04500",
)



_colonnes_modele_inscriptions = [
    "N° Session", 
    "Intitulé Session", 
    "Trigramme formation", 
    "Code IRIS", 
    "Type de formation", 
    "Année", 
    "Trigramme RP", 
    "Trigramme AF", 
    "3ème élément de la référence", 
    "Statut Session", 
    "Lieu Session", 
    "Modalité", 
    "Type", 
    "Resp pédagogique", 
    "Affectation RP", 
    "Organisatrice", 
    "Domaine parent", 
    "Participants MIN", 
    "Participants MAX", 
    "Durée (H) Session", 
    "Durée (J) Session", 
    "Date Début Session", 
    "Année Début Session", 
    "Mois Début Session", 
    "Date Fin Session", 
    "Année Fin Session", 
    "Mois Fin Session", 
    "Gestionnaire Session", 
    "Lieu de formation", 
    "Chef de projet", 
    "Code Formation", 
    "Ref. Formation", 
    "Intitulé Formation", 
    "Spécialité Formation", 
    "Code Domaine", 
    "Ref. Domaine", 
    "Domaine", 
    "Ref. Domaine principal", 
    "Domaine principal", 
    "Ref. Org. Facturation", 
    "Intitulé Org. Facturation", 
    "Dossier N°", 
    "Statut Dossier", 
    "N°Cde", 
    "Statut Cde", 
    "Organisme", 
    "Commercial", 
    "Client", 
    "Secteur d'activité", 
    "Catégorie  client", 
    "Fidélité", 
    "Autre critère", 
    "Contact Client", 
    "Mail contact client", 
    "Fonction contact client", 
    "Civilité stagiaire", 
    "Nom Stagiaire", 
    "Prénom Stagiaire", 
    "Sexe Stagiaire", 
    "Age Stagiaire", 
    "Date de naissance", 
    "Mail Stagiaire", 
    "Nationalité Stagiaire", 
    "Etablissement Stagiaire", 
    "SIRET", 
    "Entité Juridique Stagiaire", 
    "Affect. CEA / Société", 
    "CSP  Stagiaire", 
    "Contrat Stagiaire", 
    "Fonction Stagiaire", 
    "Référence Stagiaire", 
    "Motif annulation", 
    "Statut de la qualif.", 
    "Pédagogie terminée", 
    "Financier terminé", 
    "Planifiée (H)", 
    "Qualifiée (H)", 
    "Réalisée (H)", 
    "Planifiée (J)", 
    "Qualifiée (J)", 
    "Réalisée (J)", 
    "Prévu", 
    "Réalisé", 
    "Facturé", 
    "A facturer", 
    "Réglé"
]


# === Ecrire nouveaux extracts IRIS complets (qui concatène plusieurs extracts individuels) ===
# Procédure :
# IRIS : 
#    - faire export sessions pour depuis le début de l'année en cours jusqu'à 2030 (05_Planification ; puis R04110)
#    - faire export formations (01_Référentiel puis R0304_listeFormations)
# COPIER-COLLER ventes depuis référence (GED) sans écraser car petits bugs sur en-têtes de certains fichiers