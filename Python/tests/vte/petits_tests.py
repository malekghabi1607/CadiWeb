from vte.utils import *
from vte.instn import *
from vte.formation import Formation
#test = BilanFormation_V3("TEL", 2024) # Complet (sessions UEM + UECC)
#test = BilanFormation_V3("948", 2024) # Bon test car 2023 n'a rien


#chemin_fdc = Path(chemin_vers_unc(r"P:\FORMATIONS_C\TEL\P05-P06-dossier-conception-referentiel\fiche-de-cout-et-code-de-formation\Fiche de coûts INSTN - TEL - 2025.xlsx"))
#fdc = FdC(chemin_fdc=chemin_fdc)

trigramme_formation = "TEL"

formation = Formation(trigramme_formation=trigramme_formation)
formation.ajout_session(15697)

print(formation.sessions)
