from vte.instn import *
#from vte.ihm_console import IHM_console
#from vte.ihm_tkinter import IHMTkinter

def main():
    
    # periode = ["1er semestre", "2nd semestre", "Année"]

    # === Bilan unique : ok ===
    #BilanSession.bilanUnique("948", 2024, "Année")
    


    pass




def test():
    # === Bilan unique : ok ===
    # periode = ["1er semestre", "2nd semestre", "Année"]
    #BilanSession.bilanUnique("948", 2024, "Année")
    #BilanSession.bilanUnique("TEL", 2022, "Année")

    # === Plusieurs bilans ===
    bilans_a_traiter = [
        #("TEL", 2022, "Année"), OK
        #("TEL", 2023, "Année"), OK
        #("TEL", 2024, "Année"), OK
        #("TEL", 2025, "Année"), MANQUE CSV

        ("19C", 2022, "Année"),
        
        ("22B", 2025, "Année"),
        
        ("35C", 2021, "Année"),
        ("35C", 2022, "Année"),
        ("35C", 2023, "Année"),
        ("35C", 2024, "Année"),
        
        ("49C", 2024, "Année"),
        
        ("54C", 2022, "Année"),
        ("54C", 2023, "Année"),
        ("54C", 2024, "Année"),
        
        ("66B", 2022, "Année"),
        ("66B", 2023, "Année"),
        ("66B", 2024, "Année"),
        
        ("778", 2022, "Année"),
        
        ("79B", 2022, "Année"),
        ("79B", 2023, "Année"),
        ("79B", 2024, "Année"),
        
        ("80B", 2022, "Année"),
        ("80B", 2023, "Année"),
        ("80B", 2024, "Année"),
        
        ("812", 2022, "Année"),

        ("86A", 2022, "Année"),

        ("878", 2022, "Année"),

        ("948", 2022, "Année"),
        ("948", 2023, "Année"),
        ("948", 2024, "Année"),
        #("948", 2025, "Année"), PAS ENCORE FAIT

        ("949", 2022, "Année"),

        ("996", 2022, "Année"),
    ]
    BilanSession.plusieursBilans(bilans_a_traiter)

    pass



    

if __name__ == "__main__":
    test()

    # J'ouvre un export session de IRIS et load tous ses tableaux structurés dans des DataFrame (inclus dans un FichierExcel)
    # Pour l'instant je récupère le plus récent



    #main()



# === bilans de session
#bs = BilanSessionV3("948", 2024, "Année")



