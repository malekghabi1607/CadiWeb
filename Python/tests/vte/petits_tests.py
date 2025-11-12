from vte.utils import *

choisir_fichier(titre="Sélectionner un fichier CSV",
                types_fichiers=[("Fichiers CSV", "*.csv")],
                dossier_initial=Path(r"P:\FORMATIONS_C\RH3\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2024"),
                obligatoire=False,
                texte_bouton_aucun="Fichier CSV inexistant"
                )