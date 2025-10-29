from vte.instn import *
from vte.ihm_console import IHM_console
#from vte.ihm_tkinter import IHMTkinter


MODE = "console"


MENUS = {
    "Traiter à partir de codes IRIS": {
        "Code IRIS unique": {
            "action": BilanSession.bilanUnique_parCodeIRIS,
            "kwargs": {},
            "demander": []
        },
        "Plusieurs codes IRIS": {
            "action": BilanSession.plusieursBilans_parCodeIRIS,
            "kwargs": {},
            "demander": []
        },
    },

    "Traiter pour une période entière (possiblement plusieurs sessions sur un bilan)": {
        "Code IRIS unique": {
            "action": BilanSession.bilanUnique_parPeriode,  # BilanSession.bilanUnique_parPeriode("948", 2024, "Année")
            "kwargs": {},
            "demander": []
        },
        "Plusieurs codes IRIS": {
            "action": BilanSession.plusieursBilans_parPeriode,
            "kwargs": {},
            "demander": []
        },
    }
}

def main():
    
    vlog.print("Info", "Bienvenue dans le script pour générer des bilans de sessions.", style=["vert clair"])

    print("\nSi vous souhaitez faire un bilan pour une seule session à part")


    ihm = IHM_console(MENUS)
    ihm.afficher_menu()

    




def test():
    # === Bilan unique : ok ===
    # periode = ["1er semestre", "2nd semestre", "Année"]
    #BilanSession.bilanUnique("948", 2024, "Année")
    #BilanSession.bilanUnique("TEL", 2022, "Année")

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


    # === Bilan unique par code IRIS ===
    #BilanSession.bilanUnique_parCodeIRIS(17343)
    #BilanSession.bilanUnique_parCodeIRIS(16411)

    #BilanSession.plusieursBilans_parCodeIRIS([17343, 16411])

    pass



    

if __name__ == "__main__":
    #test()
    main()




