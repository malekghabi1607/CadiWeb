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
    "R0304_Ref_Formation-Listedesformations-2026.05.20.xlsx", )

_tVentes = (
    'R04301_Sessions-Ventes-FC2020 FINAL.xlsx',
    'R04301_Sessions-Ventes-FC2021 FINAL.xlsx',
    'R04301_Sessions-Ventes-FC2022 FINAL.xlsx',
    'R04301_Sessions-Ventes-FC2023 FINAL.xlsx',
    'R04301_Sessions-Ventes-FC2024 FINAL.xlsx',
    'R04301_Sessions-Ventes-FC2025 FINAL.xlsx',
    'R04301_Sessions-Ventes-filtre sur FC2026 au 2026-04-01 LG.xlsx',
    )
    
_tInscriptions = (
    'R04500_Sessions-Inscriptions-FC2020 FINAL.xlsx',
    'R04500_Sessions-Inscriptions-FC2021 FINAL.xlsx',
    'R04500_Sessions-Inscriptions-FC2022 FINAL.xlsx',
    'R04500_Sessions-Inscriptions-FC2023 FINAL.xlsx',
    'R04500_Sessions-Inscriptions-FC2024 FINAL.xlsx',
    'R04500_Sessions-Inscriptions-FC2025 FINAL.xlsx',
    'R04500_Sessions-Inscriptions-filtre sur FC2026 au 2026-04-01.xlsx',
    )