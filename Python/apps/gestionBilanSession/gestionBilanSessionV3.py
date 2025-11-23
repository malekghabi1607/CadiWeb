from vte.instn import *
from vte.ihm_console import IHM_console

from colorama import init
init(autoreset=True)


MODE = "console"

MENUS = {
    "Traiter bilans pour des sessions uniques (une seule session par bilan)": {
        "Un seul bilan à traiter par code IRIS (manuel)": {
            "action": BilanSession.bilanUnique_parCodeIRIS,
            "kwargs": {},
            "demander": []
        },
        "Plusieurs bilans à traiter par codes IRIS (manuel)": {
            "action": BilanSession.plusieursBilans_parCodeIRIS,
            "kwargs": {},
            "demander": [],
            "indications":"Entrez les codes IRIS séparés par des virgules (ex. : 16411, 17343)"
        },
        "Plusieurs bilans à traiter par codes IRIS (depuis user_config.py)": {
            "action": chargement_config_demander_verif_utilisateur,
            "kwargs": {"nom_variable": "liste_codes_IRIS", 'fonction_execution':BilanSession.plusieursBilans_parCodeIRIS_fichierConfig},
            "demander": []
            #"indications":"Entrez les codes IRIS séparés par des virgules (ex. : 16411, 17343)"
        },
    },

    "Traiter pour une période entière (possiblement plusieurs sessions sur un bilan)": {
        "Période unique (manuel)": {
            "action": BilanSession.bilanUnique_parPeriode,  # BilanSession.bilanUnique_parPeriode("948", 2024, "Année")
            "kwargs": {},
            "demander": [],
            "indications":"codeFormation = trigramme IRIS en majuscule si besoin (ex. : TEL)\n"
                        + "annee = année à 4 chiffres (ex. : 2025)\n"
                        + "periode = [1er semestre, 2nd semestre, Année] (ex. : 2nd semestre)"
        },
        "Plusieurs périodes (depuis user_config.py)": {
            "action": chargement_config_demander_verif_utilisateur,
            "kwargs": {"nom_variable": "liste_periodes", 'fonction_execution':BilanSession.plusieursBilans_parPeriode},
            "demander": [],
            "indications":"Entrez les codes IRIS séparés par des virgules (ex. : 16411, 17343)"
        },
    }
}


def main():
    
    vlog.print("Info", "*************\nBienvenue dans le script pour générer des bilans de sessions.\n*************", style=["vert clair"])

    ihm = IHM_console(MENUS)
    ihm.afficher_menu()

    




def backup_bilans_novembre_2025():
    # === Plusieurs bilans ===
    #bilans_a_traiter = [
        #("TEL", 2022, "Année"), OK
        #("TEL", 2023, "Année"), OK
        #("TEL", 2024, "Année"), OK
        #("TEL", 2025, "Année"), #A REFAIRE MANQUE CSV

        #("19C", 2022, "Année"), OK
        
        #("22B", 2025, "Année"), OK
        
        #("35C", 2021, "Année"), OK
        #("35C", 2022, "Année"), OK
        #("35C", 2023, "Année"), OK
        #("35C", 2024, "Année"), ANNULEES
        
        #("49C", 2024, "Année"), Pas de session
        
        #("54C", 2022, "Année"), OK
        #("54C", 2023, "Année"), OK
        #("54C", 2024, "Année"), ANNULEES
        
        #("66B", 2022, "Année"), OK
        #("66B", 2023, "Année"), OK
        #("66B", 2024, "Année"), ANNULEES
        
        #("778", 2022, "Année"), OK
        
        #("79B", 2022, "Année"), OK
        #("79B", 2023, "Année"), OK
        #("79B", 2024, "Année"), ANNULEES
        
        #("80B", 2022, "Année"), OK
        #("80B", 2023, "Année"), OK
        #("80B", 2024, "Année"), OK
        
        #("812", 2022, "Année"), ANNULEE

        #("86A", 2022, "Année"), OK

        #("878", 2022, "Année"), OK

        #("948", 2022, "Année"), OK
        #("948", 2023, "Année"), Pas de session
        #("948", 2024, "Année"),
        #("948", 2025, "Année"), PAS ENCORE FAIT

        #("949", 2022, "Année"), OK

        #("996", 2022, "Année"), Pas de session
    #]
    #BilanSession.plusieursBilans(bilans_a_traiter)

    pass



    

if __name__ == "__main__":
    #test()
    main()




