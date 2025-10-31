from vte.utils import *

test_path = Path(r"R:\_Echanges\VTE\Prog\Modèles")
test_str = r"R:\_Echanges\VTE\Prog\Modèles"

print(chemin_vers_unc(test_path))
print(chemin_vers_unc(test_str))