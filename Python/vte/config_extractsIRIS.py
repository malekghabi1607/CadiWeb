"""
Mode op

Permet de lister les chemins des CSV à concaténer dans le fichier global

"""

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
    'R04110_Sessions-2025 au 2025.11.29.xlsx')

_tFormations = (
    "R0304_Ref_Formation-Listedesformations-2025.11.29.xlsx", )

_tVentes = (
    'R04301_Sessions-Ventes-FC2020 FINAL.xlsx',
    'R04301_Sessions-Ventes-FC2021 FINAL.xlsx',
    'R04301_Sessions-Ventes-FC2022 FINAL.xlsx',
    'R04301_Sessions-Ventes-FC2023 FINAL.xlsx',
    'R04301_Sessions-Ventes-FC2024 FINAL.xlsx',
    'R04301_Sessions-Ventes-filtre sur FC-2025 au 2025-02-05 LG.xlsx',
    'R04301_Sessions-Ventes-filtre sur FC-2025 au 2025-03-03 LG.xlsx',
    'R04301_Sessions-Ventes-filtre sur FC-2025 au 2025-04-01 LG.xlsx',
    'R04301_Sessions-Ventes-filtre sur FC-2025 au 2025-05-12 LG.xlsx',
    'R04301_Sessions-Ventes-filtre sur FC-2025 au 2025-06-02 LG.xlsx',
    'R04301_Sessions-Ventes-filtre sur FC-2025 au 2025-07-01 LG.xlsx',
    'R04301_Sessions-Ventes-filtre sur FC-2025 au 2025-08-01 LG.xlsx',
    'R04301_Sessions-Ventes-filtre sur FC-2025 au 2025-09-01 LG.xlsx',
    'R04301_Sessions-Ventes-filtre sur FC-2025 au 2025-10-01 LG.xlsx',
    'R04301_Sessions-Ventes-filtre sur FC-2025 au 2025-11-03 LG.xlsx')
    
_tInscriptions = (
    'R04500_Sessions-Inscriptions-FC2020 FINAL.xlsx',
    'R04500_Sessions-Inscriptions-FC2021 FINAL.xlsx',
    'R04500_Sessions-Inscriptions-FC2022 FINAL.xlsx',
    'R04500_Sessions-Inscriptions-FC2023 FINAL.xlsx',
    'R04500_Sessions-Inscriptions-FC2024 FINAL.xlsx',
    'R04500_Sessions-Inscriptions-filtre sur FC-2025 au 2025-02-05.xlsx',
    'R04500_Sessions-Inscriptions-filtre sur FC-2025 au 2025-03-03.xlsx',
    'R04500_Sessions-Inscriptions-filtre sur FC-2025 au 2025-04-01.xlsx',
    'R04500_Sessions-Inscriptions-filtre sur FC-2025 au 2025-05-12.xlsx',
    'R04500_Sessions-Inscriptions-filtre sur FC-2025 au 2025-06-02.xlsx',
    'R04500_Sessions-Inscriptions-filtre sur FC-2025 au 2025-07-01.xlsx',
    'R04500_Sessions-Inscriptions-filtre sur FC-2025 au 2025-08-01.xlsx',
    'R04500_Sessions-Inscriptions-filtre sur FC-2025 au 2025-09-01.xlsx',
    'R04500_Sessions-Inscriptions-filtre sur FC-2025 au 2025-10-01.xlsx',
    'R04500_Sessions-Inscriptions-filtre sur FC-2025 au 2025-11-03.xlsx')

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