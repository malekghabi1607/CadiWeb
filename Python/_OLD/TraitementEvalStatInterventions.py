#from ast import Index
#from operator import index
#from urllib.parse import non_hierarchical
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo


import pandas as pd
import re #expressions régulières
import shutil #manipulation de fichiers (copie...)





### --------------------------------------------------------------------
#  Initialisation fonctions utilisateur
### --------------------------------------------------------------------

def concatene_CSV(listeCSV) :
    """Fonction qui boucle sur les fichiers CSV indiqués dans la liste listeCSV et qui concatène les CSV dans un seul dataframe."""
    
    df_out = pd.DataFrame() #Dataframe temporaire (c'est lui qui réceptionnera la lecture brute des CSV)
    
    for iCSV in listeCSV :
        #On récupère le trigramme de la formation et id session
        m = rechercheTrigrammeFormation.match(iCSV)
        #print(m.group(1), m.group(2))

        #On récupère les données du CSV
        data=pd.read_csv(iCSV, sep=';', encoding='latin-1')
        df_out = pd.concat([df_out, data])
    df_out.reset_index(drop=True) #Permet de supprimer les indices et de les réinitialiser. Normalement, cette fonction crée une nouvelle colonne avec les anciens indices ; avec l'option drop, ça permet de virer complètement ces anciens indices
    #print(df_interventions_temp)
    return(df_out)

def mk_df_interventions(fichier_csv):
    """
        A partir d'un dataframe CSV des interventions eval&go, on en fait un dataframe qui est réarrangé
 
        ***Description plus complète à faire
 
        :param fichier_csv: The first number to add
        :type fichier_csv: string
        :return: Un DataFrame qui est formaté selon ce que l'on souhaite pour exporter vers excel (i.e. structure différente que le CSV)
        :rtype: DataFrame
 
        :Example:
 
        >>> DataFrame.df = mk_df_interventions("fichier.csv")

 
        .. seealso:: Rien du tout.
        .. warning:: Rien du tout.
        .. note:: Rien du tout.
        .. todo:: Rien du tout.
    """
    data_csv = pd.DataFrame() #Dataframe temporaire (c'est lui qui réceptionnera la lecture brute des CSV)
    df_interventions = pd.DataFrame(columns = enTete_interventions) #Dataframe final des interventions et qui sera exporté vers Excel
    df = pd.DataFrame(columns = enTete_interventions) #Dataframe final des interventions et qui sera exporté vers Excel

    #On récupère le trigramme de la formation et id session à partir du chemin du fichier CSV
    resultat_rech = rechercheTrigrammeFormation.match(fichier_csv)
    #print(resultat_rech.group(1), resultat_rech.group(2))

    #On récupère les données du CSV
    data_csv=pd.read_csv(fichier_csv, sep=';', encoding='latin-1', index_col=None) #df = pd.read_csv('addresses.csv', index_col=None,  header=None)

    #On adapte les données pour avoir la forme souhaitée (ici : interventions)
    for index, iLigne in data_csv.iterrows() :
        #print("index = ",index)
        #print(iLigne)

        df = {
                "id":index, 
                "Formation":resultat_rech.group(1), 
                "Session": resultat_rech.group(2), 
                "Date": iLigne["Date"], 
                "Nom et prénom": iLigne["NOM et Prénom"], 
                "Société": iLigne["Société "], 
                "Cours suivi": iLigne["Cours suivi"], 
                "Critère": "NA", 
                "Valeur": "NA"
                } 

        for iCritere in listeCriteres_interventions:
            #print("iCritere = ", iCritere)
            df["Critère"] = iCritere
            df["Valeur"] = iLigne[iCritere]
            df_interventions = df_interventions._append(df, ignore_index = True)


    #print(df_interventions)
    #wait = input("Sortie de la fonction mk_df_interventions")

    return(df_interventions)




def mk_df_sessions(fichier_csv):
    """
        A partir d'un dataframe CSV des sessions eval&go (nommé stagiaires), on en fait un dataframe qui est réarrangé
 
        ***Description plus complète à faire
 
        :param fichier_csv: The first number to add
        :type fichier_csv: string
        :return: Un DataFrame qui est formaté selon ce que l'on souhaite pour exporter vers excel (i.e. structure différente que le CSV)
        :rtype: DataFrame
 
        :Example:
 
        >>> DataFrame.df = mk_df_interventions("fichier.csv")

 
        .. seealso:: Rien du tout.
        .. warning:: Rien du tout.
        .. note:: Rien du tout.
        .. todo:: Rien du tout.
    """
    data_csv = pd.DataFrame() #Dataframe temporaire (c'est lui qui réceptionnera la lecture brute des CSV)
    df_sessions = pd.DataFrame(columns = enTete_sessions) #Dataframe final des interventions et qui sera exporté vers Excel
    df = pd.DataFrame(columns = enTete_sessions) #Dataframe final des interventions et qui sera exporté vers Excel

    #On récupère le trigramme de la formation et id session à partir du chemin du fichier CSV
    resultat_rech = rechercheTrigrammeFormation.match(fichier_csv)
    #print(resultat_rech.group(1), resultat_rech.group(2))

    #On récupère les données du CSV
    data_csv=pd.read_csv(fichier_csv, sep=';', encoding='latin-1', index_col=None) #df = pd.read_csv('addresses.csv', index_col=None,  header=None)

    #On adapte les données pour avoir la forme souhaitée (ici : interventions)
    for index, iLigne in data_csv.iterrows() :
        #print("index = ",index)
        #print(iLigne)
        #print(iLigne.dtypes)

        

        df = {
                "id":index, 
                "Formation":resultat_rech.group(1), 
                "Session": resultat_rech.group(2), 
                "Date": iLigne["Date"], 
                "Prénom": iLigne.iloc[3], #On ne peut pas utiliser "Prénom" car il est déjà employé et vide
                "Nom": iLigne.iloc[4],
                "Société": iLigne["Entreprise"], 
                "Code session": iLigne["Code session"], 
                "Critère": "NA", 
                "Valeur": "NA",
                "Commentaires": "NA"
                } 

        for iCritere in listeCriteres_sessions:
            #print("iCritere = ", iCritere)
            df["Critère"] = iCritere
            df["Valeur"] = iLigne[iCritere]
            
            if iCritere in listeCriteres_sessions_sansCommentaires:
                df["Commentaires"] = ""
            else:
                #On recherche l'ID de la colonne de ce critère pour y mettre un +1 après (nous voulons la colonne qui est juste après)
                idCol = data_csv.columns.get_loc(iCritere)+1
                df["Commentaires"] = iLigne.iloc[idCol]

    
            df_sessions = df_sessions._append(df, ignore_index = True)

    #print(df_sessions)

    return(df_sessions)



def writeInStructuredRef(df, nom_ws, nom_tableauStructure, lettreFinTableau):
    ws = wb[nom_ws]
    table = ws._tables[nom_tableauStructure]

    #On redéfinit le dimensionnement du tableau
    table.ref = "A1:{}{}".format(lettreFinTableau, len(df)+1) #+1 car on a la ligne d'en-tête
    ws.delete_rows(idx=2, amount=1) #On supprime la ligne vide du tableau structuré

    #On écrit toutes les autres lignes une par une
    for il in df.itertuples(): #for r in dataframe_to_rows(df, header=True, index=False):
        ws.append(il[1:-1])










### --------------------------------------------------------------------
#  Initialisation variables utilisateur
### --------------------------------------------------------------------

rechercheTrigrammeFormation = re.compile(r'^[^\\]*\\[^\\]*\\(\w\w\w)\\.*S-(\w{5}).*')


s_classeurIni = r'C:\Users\vt238770\Documents\_CEA\EXCEL - Tests ou Backup\Stats-Eval&Go2.xlsx'
s_classeurDestination = r'C:\Users\vt238770\Documents\_CEA\EXCEL - Tests ou Backup\Stats-Eval&Go2 - Copie.xlsx'
s_classeurDestination2 = r'C:\Users\vt238770\Documents\_CEA\EXCEL - Tests ou Backup\output2.xlsx'


s_csv = r'C:\Users\vt238770\Downloads\Classeur1.csv'
s_csv2 = r'"C:\Users\vt238770\Downloads\Classeur2.csv"'







s_ws_interventions = r'InterventionsParCritere' #Nom du classeur
s_tableauStructure_interventions = s_ws_interventions #Nom du tableau structuré


listeCriteres_interventions = ["Intérêt du sujet et méthodes", "Adéquation contenu/durée", "Qualité de l'animation ", "Qualité des supports de présentation et de cours", "Atteintes des objectifs liées à la séquence (prise en compte de mes besoins et attentes)", "Commentaires "] #Liste des critères du CSV interventions
enTete_interventions = ["Formation", "Session", "Date", "Nom et prénom", "Société", "Cours suivi", "Critère", "Valeur"] #["id"] + list(df_destination.columns)
df_interventions = pd.DataFrame(columns = enTete_interventions) #Dataframe qui sera exporté vers Excel
#df_interventions_temp = pd.DataFrame()

liste_fichiersInterventions = [] #Liste des fichiers interventions à lire
#liste_fichiersInterventions.append(s_csv)
#liste_fichiersInterventions.append(s_csv2)
liste_fichiersInterventions.append(r'P:\FORMATIONS_C\35C\P07-bilan-sessions-et-bilan-formation\rapports-sessions-evaluations\2021-12-S-12664 UEM\S-12664-FC21-35C-VTE-SNA-Interventions.csv')
liste_fichiersInterventions.append(r'P:\FORMATIONS_C\35C\P07-bilan-sessions-et-bilan-formation\rapports-sessions-evaluations\2022-03-S-13172 UEM\S-13172-FC22-35C-VTE-SNA-Interventions.csv')
liste_fichiersInterventions.append(r'P:\FORMATIONS_C\35C\P07-bilan-sessions-et-bilan-formation\rapports-sessions-evaluations\2022-11-S-13362 UEM\S-13362-FC22-35C-VTE-SNA-Interventions.csv')
liste_fichiersInterventions.append(r'P:\FORMATIONS_C\54C\P07-bilan-sessions-et-bilan-formation\rapports-sessions-evaluations\2022-06-S-13165  UEM\S-13165-FC22-54C-VTE-SNA-Interventions.csv')
liste_fichiersInterventions.append(r'P:\FORMATIONS_C\54C\P07-bilan-sessions-et-bilan-formation\rapports-sessions-evaluations\2022-11-S-13720 UES\S-13720-FC22-54C-YBU-CVI-interventions.csv')
liste_fichiersInterventions.append(r'P:\FORMATIONS_C\54C\P07-bilan-sessions-et-bilan-formation\rapports-sessions-evaluations\2023-06-S14317 UEM\S-14317-FC23-54C-VTE-LRA-Interventions.csv')
liste_fichiersInterventions.append(r'P:\FORMATIONS_C\79B\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2021-09-S-13361 UEM\S-11361-FC21-79B-PDX-SNA-Interventions.csv')
liste_fichiersInterventions.append(r'P:\FORMATIONS_C\79B\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2022-03-S-13168 UEM\S-13168-FC22-79B-PDX-SNA-Interventions.csv')
liste_fichiersInterventions.append(r'P:\FORMATIONS_C\79B\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2022-09-S-13169 UEM\S-13169-FC22-79B-VTE-SNA-Interventions.csv')
liste_fichiersInterventions.append(r'P:\FORMATIONS_C\80B\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2022-09-S-13171 UEM\S-13171-FC22-80B-VTE-SNA-Interventions.csv')
liste_fichiersInterventions.append(r'P:\FORMATIONS_C\80B\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\S-11358-FC21-80B-PDX-SNA\S-11358-FC21-80B-PDX-SNA-Interventions.csv')
liste_fichiersInterventions.append(r'P:\FORMATIONS_C\80B\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\S-11365-FC21-80B-PDX-SNA\CSV\S-11365-FC21-80B-PDX-SNA-Interventions.csv')




s_ws_sessions = r'SessionsParCritere' #Nom du classeur
s_tableauStructure_sessions = s_ws_sessions #Nom du tableau structuré

listeCriteres_sessions = ["Comment avez-vous connu cette formation ?", "Accueil, organisation et qualité des informations délivrées", "Conseils et orientation avant l'inscription ", "Informations après l'inscription ", "Accueil à l'arrivée sur site ", "Prise en compte de vos besoins et attentes", "Qualité des animations", "Logique d'enchainement des interventions", "Qualité des supports de cours utilisés", "Qualité des moyens pédagogique", "Accès aux outils digitaux", "Satisfaction globale", "Recommanderiez-vous cette formation ?", "Avez-vous d'autres besoins de formation ? ", "Lesquels ?", "Commentaires, remarques, suggestions "] #Liste des critères du CSV stagiaires
listeCriteres_sessions_sansCommentaires = ["Comment avez-vous connu cette formation ?", "Recommanderiez-vous cette formation ?", "Avez-vous d'autres besoins de formation ? ", "Lesquels ?", "Commentaires, remarques, suggestions "] #Liste des critères du CSV stagiaires qui ne nécessitent pas de champ "Commentaires"
enTete_sessions = ["Formation", "Session", "Date", "Prénom", "Nom", "Société", "Code session", "Critère", "Valeur", "Commentaires"]
df_sessions = pd.DataFrame(columns = enTete_sessions) #Dataframe qui sera exporté vers Excel
#df_sessions_temp = pd.DataFrame()

#On initialise la liste des fichiers à traiter
liste_fichiersSessions = []
liste_fichiersSessions.append(r'P:\FORMATIONS_C\35C\P07-bilan-sessions-et-bilan-formation\rapports-sessions-evaluations\2021-12-S-12664 UEM\S-12664-FC21-35C-VTE-SNA-Stagiaires.csv')
liste_fichiersSessions.append(r'P:\FORMATIONS_C\35C\P07-bilan-sessions-et-bilan-formation\rapports-sessions-evaluations\2022-03-S-13172 UEM\S-13172-FC22-35C-VTE-SNA-Stagiaires.csv')
liste_fichiersSessions.append(r'P:\FORMATIONS_C\35C\P07-bilan-sessions-et-bilan-formation\rapports-sessions-evaluations\2022-11-S-13362 UEM\S-13362-FC22-35C-VTE-SNA-Stagiaires.csv')
liste_fichiersSessions.append(r'P:\FORMATIONS_C\54C\P07-bilan-sessions-et-bilan-formation\rapports-sessions-evaluations\2022-06-S-13165  UEM\S-13165-FC22-54C-VTE-SNA-Stagiaires.csv')
liste_fichiersSessions.append(r'P:\FORMATIONS_C\54C\P07-bilan-sessions-et-bilan-formation\rapports-sessions-evaluations\2022-11-S-13720 UES\S-13720-FC22-54C-YBU-CVI-stagiaires.csv')
liste_fichiersSessions.append(r'P:\FORMATIONS_C\54C\P07-bilan-sessions-et-bilan-formation\rapports-sessions-evaluations\2023-06-S14317 UEM\S-14317-FC23-54C-VTE-LRA-Stagiaires.csv')
liste_fichiersSessions.append(r'P:\FORMATIONS_C\79B\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2021-09-S-13361 UEM\S-11361-FC21-79B-PDX-SNA-Stagiaires.csv')
liste_fichiersSessions.append(r'P:\FORMATIONS_C\79B\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2022-03-S-13168 UEM\S-13168-FC22-79B-PDX-SNA-Stagiaires.csv')
liste_fichiersSessions.append(r'P:\FORMATIONS_C\79B\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2022-09-S-13169 UEM\S-13169-FC22-79B-VTE-SNA-Stagiaires.csv')
liste_fichiersSessions.append(r'P:\FORMATIONS_C\79B\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2023-09-S-14585 UES\S-14585-FC23-79B-OCO-CPI-stagiaires.csv')
liste_fichiersSessions.append(r'P:\FORMATIONS_C\80B\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\2022-09-S-13171 UEM\S-13171-FC22-80B-VTE-SNA-Stagiaires.csv')
liste_fichiersSessions.append(r'P:\FORMATIONS_C\80B\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\S-11358-FC21-80B-PDX-SNA\S-11358-FC21-80B-PDX-SNA-Stagiaires.csv')
liste_fichiersSessions.append(r'P:\FORMATIONS_C\80B\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\S-11365-FC21-80B-PDX-SNA\CSV\S-11365-FC21-80B-PDX-SNA-Stagiaires.csv')
liste_fichiersSessions.append(r'P:\FORMATIONS_C\80B\P07-bilan-sessions-et-bilan-formation\rapports-sessions-CSV-evaluations\S-14517-FC23-80B-OCO-CVI\S-14517-FC23-80B-OCO-CVI-stagiaires.csv')








### --------------------------------------------------------------------
#  Début code
### --------------------------------------------------------------------


for iFichier in liste_fichiersInterventions :
    #On concatène à df_interventions le nouveau CSV traité
    df_interventions = pd.concat([df_interventions, mk_df_interventions(iFichier)])
    #print(df_interventions)
    #wait = input("Fin de la boucle dans le corps principal")
df_interventions.reset_index(drop=True) #Permet de supprimer les indices et de les réinitialiser. Normalement, cette fonction crée une nouvelle colonne avec les anciens indices ; avec l'option drop, ça permet de virer complètement ces anciens indices





for iFichier in liste_fichiersSessions :
    #On concatène à df_interventions le nouveau CSV traité
    df_sessions = pd.concat([df_sessions, mk_df_sessions(iFichier)])
    #print(df_interventions)
    #wait = input("Fin de la boucle dans le corps principal")
df_sessions.reset_index(drop=True) #Permet de supprimer les indices et de les réinitialiser. Normalement, cette fonction crée une nouvelle colonne avec les anciens indices ; avec l'option drop, ça permet de virer complètement ces anciens indices


#df_sessions = pd.concat([df_sessions, mk_df_sessions(liste_fichiersSessions[0])])



#On copie le classeur
shutil.copy(s_classeurIni, s_classeurDestination)

#On ouvre le classeur
wb = load_workbook(filename = s_classeurDestination)

#On écrit les interventions dans l'Excel
writeInStructuredRef(df_interventions, s_ws_interventions, s_tableauStructure_interventions, "H")

#On écrit les sessions dans l'Excel
writeInStructuredRef(df_sessions, s_ws_sessions, s_tableauStructure_sessions, "J")

#On enregistre et on ferme
wb.save(filename=s_classeurDestination)
wb.close()







#tableRef = 0
##print(enumerate(ws._tables))
#for i, table in enumerate(ws._tables):
#    print(i)
#    print(table)
#    if table == s_tableauStructure_interventions:
#        tableRef = i
#        print(tableRef)



#resTable = Table(displayName="Data", ref="A1:{}{}".format(colnum_string(maxRef[0]), maxRef[1]))
#resTable = Table(displayName="Data", ref="A1:H50")
#resTable.tableStyleInfo = style
