from vte.domain.formation import Formation
from vte.evalStat import *

trigramme_formation = "TEL"

formation = Formation(trigramme_formation=trigramme_formation)
formation.ajout_sessions((16411, 16161))

print(formation.sessions)