from vte.utils.utils import *
from vte.domain.formation import Formation
from vte.domain.bilanSessions import *
#test = BilanFormation_V3("TEL", 2024) # Complet (sessions UEM + UECC)
#test = BilanFormation_V3("948", 2024) # Bon test car 2023 n'a rien


#chemin_fdc = Path(chemin_vers_unc(r"P:\FORMATIONS_C\TEL\P05-P06-dossier-conception-referentiel\fiche-de-cout-et-code-de-formation\Fiche de coûts INSTN - TEL - 2025.xlsx"))
#fdc = FdC(chemin_fdc=chemin_fdc)

trigramme_formation = "TEL"

#formation = Formation(trigramme_formation=trigramme_formation)
#formation.ajout_sessions(15697)

#print(formation.sessions)

print(
    BilanSessions.construire_chemin_word_bilan_sessions_output(
        trigramme_formation="TEL",
        annee=2025,
        periode_pour_titre="Année 2025",
        unite="UEM"
    )
)
