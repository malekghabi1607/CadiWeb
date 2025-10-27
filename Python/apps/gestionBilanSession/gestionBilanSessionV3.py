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

    # === Plusieurs bilans ===
    bilans_a_traiter = [
        ("TEL", 2022, "Année"),
        ("TEL", 2023, "Année"),
        ("TEL", 2024, "Année"),
        ("TEL", 2025, "Année"),

        ("22B", 2022, "Année"),
        ("22B", 2023, "Année"),
        ("22B", 2024, "Année"),
        ("22B", 2025, "Année"),
        
        ("948", 2022, "Année"),
        ("948", 2023, "Année"),
        ("948", 2024, "Année"),
        ("948", 2025, "Année"),

        ("19C", 2022, "Année"),
        ("19C", 2023, "Année"),
        ("19C", 2024, "Année"),
        ("19C", 2025, "Année"),
        
        ("70B", 2022, "Année"),
        ("19C", 2023, "Année"),
        ("19C", 2024, "Année"),
        ("19C", 2025, "Année"),
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



