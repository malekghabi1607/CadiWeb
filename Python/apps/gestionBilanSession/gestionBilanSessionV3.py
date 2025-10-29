from vte.instn import *
#from vte.ihm_console import IHM_console
#from vte.ihm_tkinter import IHMTkinter

def main():
    
    

    ####
    # PAR PERIODE
    ####
    # periode = ["1er semestre", "2nd semestre", "Année"]

    # === Bilan unique : ok ===
    #BilanSession.bilanUnique("948", 2024, "Année")
    
    # === Plusieurs bilans : ok ===
    bilans_a_traiter = [
        ("TEL", 2022, "Année"),
        ("TEL", 2023, "Année"),
        ("TEL", 2024, "Année"),
    ]
    #BilanSession.plusieursBilans(bilans_a_traiter)





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
        ("TEL", 2025, "Année"), #A REFAIRE MANQUE CSV

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
    ]
    #BilanSession.plusieursBilans(bilans_a_traiter)


    # === Bilan unique par code IRIS ===
    #BilanSession.bilanUnique_parCodeIRIS(17343)
    #BilanSession.bilanUnique_parCodeIRIS(16411)

    test = """
        <p>Bonjour Florent,</p>
        <p>Est-ce que tu peux signer le bilan de session ci-dessous stp.\nLien du bilan de session : <a href="{lien_pdf_bilan}">{lien_pdf_bilan}</a></p> 
        <p>Il concerne la formation {formation} : {periode}.</p> 
        <p>Je te remercie, passe une excellente journée,</p>
    """

    print(remplacer_champs(test, [
            ["lien_pdf_bilan", "a"],
            ["formation", "b"],
            ["periode", "c"],
        ]))

    pass



    

if __name__ == "__main__":
    test()

    # J'ouvre un export session de IRIS et load tous ses tableaux structurés dans des DataFrame (inclus dans un FichierExcel)
    # Pour l'instant je récupère le plus récent



    #main()



# === bilans de session
#bs = BilanSessionV3("948", 2024, "Année")



