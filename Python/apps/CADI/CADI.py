import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import tkinter.font as tkfont

import threading

import os, json, copy, re
from datetime import datetime
from typing import List, Tuple

#import INSTN
from vte.instn import *
#from INSTN import Traiter_evalStat, FichierExcel, chemin_vers_unc, vlog


# Todo CADI : faire un truc pour la lecture sessions qui peut changer (dernière version de GLOBAL)
#def etat_boutons_traitement(state):
#    # Désactive ou active les 3 boutons
#    for btn in (btn_traiter_eval, bilan_button, iris_button):
#        btn.config(state=state)

#def chargements_initiaux():
#    fe_sessions = FichierExcel.depuis_fichier(config["Extractions d'IRIS"]["Fichier sessions R04110"])



# Au lancement, désactiver les 3 boutons de traitement
#update_status("Lecture fichier session")
#etat_boutons_traitement('disabled')

#fe_sessions = None
# Lancer le chargement en thread séparé
#threading.Thread(target=chargements_initiaux, daemon=True).start()

#print(fe_sessions)

# Quand c'est fini, on réactive les boutons, mais dans le thread Tkinter !
#etat_boutons_traitement('normal')
#update_status("Prêt.")



#fichiers = tuple(map(chemin_vers_unc, filedialog.askopenfilenames(filetypes=[("CSV files", "*.csv")])))
#print(fichiers)


#Tu peux maintenant :
#Afficher des messages dans la barre de statut avec update_status("ton message"),
#Suivre les étapes de traitement en direct pendant les clics,
#Ajouter des appels à update_status(...) dans ton futur code métier.


class ApplicationCADI(tk.Tk):
    FICHIER_CONFIG = "config.json"

    CONFIG_PAR_DEFAUT = {
        "Extractions d'IRIS": {
            "Fichier sessions R04110": "",
            "Fichier formations R0304": "",
            "Fichier inscriptions R04500": "",
            "Fichier ventes R04301": ""
        },
        "EvalStat": {
            "Sauvegarde des évaluations d'une formation": "",
            "Modèle Excel des évaluations" : ""
        },
        "Bilans": {
            "Répertoire par défaut fiche de coûts": "",
            "Répertoire par défaut specs pédagogiques": "",
            "Modèle Word du bilan": ""
        }
    }

    def __init__(self):
        super().__init__()
        self.title("CADI - Création Automatique de Documents INSTN")
        self.geometry("800x600")

        # Données
        self.config = self.charger_config()
        self.fichiers_evalstat = {}  # {chemin: (formation, nom_fichier)}
        self.lecture_initiale_fe_sessions_fini = False
        self.fe_sessions:FichierExcel = None

        # UI principale
        # Styles
        self.creer_style_widgets()

        # Organisation des sections
        self.cadre_principal = ttk.Frame(self, padding=10)
        self.cadre_principal.pack(fill="both", expand=True)

        self.creer_section_evalstat()
        self.creer_section_bilan()
        self.creer_section_iris()
        self.creer_barre_boutons_bas()
        self.creer_barre_statut()

        self.maj_statut("Prêt.")

        # Événement : quitter avec Échap
        self.bind("<Escape>", lambda e: self.quit())

        # Todo : faire un truc pour la lecture sessions qui peut changer (dernière version de GLOBAL)

        # === Chargement de du fichier session en tâche de fond ===
        # Au lancement, désactiver les 3 boutons de traitement
        self.maj_statut("Lecture fichier session...")
        self.etat_boutons_traitement('disabled')

        # Lancer le chargement en thread séparé
        threading.Thread(target=self.charger_feSessions, daemon=True).start()

        

        
        

    # === FONCTIONS VTE ===
    def charger_feSessions(self):
        self.fe_sessions = FichierExcel.depuis_fichier(self.config["Extractions d'IRIS"]["Fichier sessions R04110"])

        # Retour dans le thread principal Tkinter :
        self.after(0, self.fin_chargement_feSessions)
    
    def fin_chargement_feSessions(self):
        # Quand c'est fini, on réactive les boutons, mais dans le thread Tkinter !
        # print(self.fe_sessions)
        self.lecture_initiale_fe_sessions_fini = True
        self.etat_boutons_traitement('normal')
        self.maj_statut("Prêt.")

    def etat_boutons_traitement(self, state):
        # Désactive ou active les 3 boutons
        for btn in (self.bouton_traiter_eval, self.bouton_bilan, self.bouton_traiter_iris):
            btn.config(state=state)

    # === CONFIGURATION ===

    def charger_config(self):
        if not os.path.exists(self.FICHIER_CONFIG):
            self.sauvegarder_config(self.CONFIG_PAR_DEFAUT)
            return copy.deepcopy(self.CONFIG_PAR_DEFAUT)
        with open(self.FICHIER_CONFIG, "r", encoding="utf-8") as f:
            return json.load(f)

    def sauvegarder_config(self, config):
        with open(self.FICHIER_CONFIG, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=4, ensure_ascii=False)

    # === STYLE ===

    def creer_style_widgets(self):
        style = ttk.Style()
        style.configure("Danger.TButton", foreground="red")
        style.configure("Vert.TButton", foreground="green")
        style.configure("CadreBordé.TFrame", borderwidth=1, relief="solid")

    # === CONSTRUCTION DE L'INTERFACE ===

    def creer_section_evalstat(self):
        cadre_eval = ttk.LabelFrame(self.cadre_principal, text="Traiter des évaluations EvalStat")
        cadre_eval.pack(fill="x", padx=10, pady=5)

        # --- Boutons du haut ---
        cadre_boutons = ttk.Frame(cadre_eval)
        cadre_boutons.pack(fill="x", padx=5, pady=(5, 0))

        self.bouton_ajouter_fichiers = ttk.Button(
            cadre_boutons, text="Ajouter fichier(s)", command=self.ajouter_fichiers_evalstat
        )
        self.bouton_ajouter_fichiers.pack(side="left")

        self.bouton_supprimer_selection = ttk.Button(
            cadre_boutons, text="Supprimer sélection", style="Danger.TButton",
            command=self.supprimer_selection, state="disabled"
        )
        self.bouton_supprimer_selection.pack(side="left", padx=(10, 0))

        # --- Tableau Treeview ---
        cadre_tableau_externe = ttk.Frame(cadre_eval, style="CadreBordé.TFrame")
        cadre_tableau_externe.pack(fill="both", expand=True, padx=5, pady=5)

        self.arbre_eval = ttk.Treeview(
            cadre_tableau_externe,
            columns=("formation", "fichier", "chemin"),
            show="headings",
            selectmode="extended",
            height=2
        )
        self.arbre_eval.pack(side="left", fill="both", expand=True)

        # Scrollbars
        scroll_y = ttk.Scrollbar(cadre_tableau_externe, orient="vertical", command=self.arbre_eval.yview)
        scroll_y.pack(side="right", fill="y")
        scroll_x = ttk.Scrollbar(cadre_tableau_externe, orient="horizontal", command=self.arbre_eval.xview)
        scroll_x.pack(side="bottom", fill="x")

        self.arbre_eval.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        # En-têtes et colonnes
        self.arbre_eval.heading("formation", text="Formation")
        self.arbre_eval.heading("fichier", text="Fichier CSV")
        self.arbre_eval.heading("chemin", text="Chemin complet")

        self.arbre_eval.column("formation", width=150, anchor="center", stretch=False)
        self.arbre_eval.column("fichier", width=250, stretch=False)
        self.arbre_eval.column("chemin", width=600, stretch=False)

        self.arbre_eval.bind("<<TreeviewSelect>>", self.maj_etat_boutons_eval)

        # Scroll molette
        self.arbre_eval.bind("<MouseWheel>", self._scroll_vertical)
        self.arbre_eval.bind("<Button-4>", lambda e: self.arbre_eval.yview_scroll(-1, "units"))  # Linux
        self.arbre_eval.bind("<Button-5>", lambda e: self.arbre_eval.yview_scroll(1, "units"))   # Linux
        self.arbre_eval.bind("<Shift-MouseWheel>", self._scroll_horizontal)

        # --- Bouton "Traiter" ---
        cadre_bas = ttk.Frame(cadre_eval)
        cadre_bas.pack(fill="x", padx=5, pady=(0, 5))

        self.bouton_traiter_eval = ttk.Button(
            cadre_bas, text="Traiter EvalStat", style="Vert.TButton",
            command=self.traiter_evalstat, state="disabled"
        )
        self.bouton_traiter_eval.pack(side="right")

    def creer_section_bilan(self):
        cadre_bilan = ttk.LabelFrame(self.cadre_principal, text="Générer un bilan de formation")
        cadre_bilan.pack(fill="x", padx=5, pady=5)

        self.var_annee_bilan = tk.IntVar(value=datetime.now().year - 1)
        self.var_trigramme_bilan = tk.StringVar()

        ttk.Label(cadre_bilan, text="Année du bilan :").pack(anchor="w")
        ttk.Entry(cadre_bilan, textvariable=self.var_annee_bilan).pack(anchor="w", pady=2)

        ttk.Label(cadre_bilan, text="Trigramme formation :").pack(anchor="w")
        ttk.Entry(cadre_bilan, textvariable=self.var_trigramme_bilan).pack(anchor="w", pady=2)

        self.bouton_bilan = ttk.Button(
            cadre_bilan,
            text="Générer bilan",
            style="Vert.TButton",
            command=self.generer_bilan,
            state="disabled"
        )
        self.bouton_bilan.pack(anchor="e", pady=5)

        self.var_annee_bilan.trace_add("write", self.verifier_bilan_valide)
        self.var_trigramme_bilan.trace_add("write", self.verifier_bilan_valide)

    def creer_section_iris(self):
        cadre_iris = ttk.LabelFrame(self.cadre_principal, text="Mettre à jour fichiers export d'IRIS")
        cadre_iris.pack(fill="x", padx=5, pady=5)

        self.vars_iris = {
            "Sessions": tk.BooleanVar(value=True),
            "Formations": tk.BooleanVar(value=True),
            "Ventes": tk.BooleanVar(value=True),
            "Inscriptions": tk.BooleanVar(value=True),
        }

        for texte, var in self.vars_iris.items():
            ttk.Checkbutton(cadre_iris, text=texte, variable=var).pack(anchor="w")

        self.bouton_traiter_iris = ttk.Button(
            cadre_iris,
            text="Traiter extracts",
            style="Vert.TButton",
            command=self.traiter_extraits_iris
        )
        self.bouton_traiter_iris.pack(anchor="e", pady=5)


    # === COMPORTEMENTS ÉVALSTAT ===

    def maj_etat_boutons_eval(self, event=None):
        nb_items = len(self.arbre_eval.get_children())
        sel = self.arbre_eval.selection()
        if self.lecture_initiale_fe_sessions_fini:
            self.bouton_traiter_eval.state(["!disabled"] if nb_items > 0 else ["disabled"])
        else:
            self.maj_statut("Lecture fichier session...")
        self.bouton_supprimer_selection.state(["!disabled"] if sel else ["disabled"])

    def ajouter_fichiers_evalstat(self):
        #fichiers = filedialog.askopenfilenames(filetypes=[("CSV files", "*.csv")])
        fichiers = tuple(map(chemin_vers_unc, filedialog.askopenfilenames(filetypes=[("CSV files", "*.csv")])))
        if not fichiers:
            return
        infos = self.extraire_infos_fichiers(fichiers)
        for formation, nom_fichier, chemin in infos:
            if chemin not in self.fichiers_evalstat:
                self.fichiers_evalstat[chemin] = (formation, nom_fichier)
                self.arbre_eval.insert("", "end", iid=chemin, values=(formation, nom_fichier, chemin))
        self.ajuster_largeur_colonnes()
        self.maj_etat_boutons_eval()
        self.maj_statut(f"{len(infos)} fichier(s) ajouté(s).")

    def supprimer_selection(self):
        for item_id in self.arbre_eval.selection():
            self.arbre_eval.delete(item_id)
            self.fichiers_evalstat.pop(item_id, None)
        self.maj_etat_boutons_eval()
        self.maj_statut("Sélection supprimée.")

    def traiter_evalstat(self):
        self.maj_statut("Traitement EvalStat en cours...")

        try:
            # On lance le traitement
            # TODO : faire un truc pour voir la progression
            evalstat = Traiter_evalStat.depuis_tuple_csv_stagiaires(
                tuple_csv_stagiaires=tuple(self.fichiers_evalstat.keys()),
                chemin_excel_evaluations_defaut=self.config["EvalStat"]["Sauvegarde des évaluations d'une formation"],
                chemin_modeleExcel_stagiaires=self.config["EvalStat"]["Modèle Excel des évaluations"],
                #chemin_excel_sessions=self.config["Extractions d'IRIS"]["Fichier sessions R04110"],
                fe_sessions=self.fe_sessions,
                ouvrirDossier=False
            )

            # Supprimer les fichiers traités de l'affichage
            for item_id in self.arbre_eval.get_children():
                chemin = self.arbre_eval.item(item_id, "values")[2]
                if chemin in evalstat.chemins_csv_traites:
                    self.arbre_eval.delete(item_id)
                    self.fichiers_evalstat.pop(item_id, None)

            self.maj_statut("Traitement EvalStat terminé.")
            self.maj_etat_boutons_eval()

            if vlog.dict_messages:
                vlog.afficher_popup("Rapport du traitement des EvalStat")

        except Exception as e:
            self.maj_statut(f"Erreur : {e}")

    def extraire_infos_fichiers(self, chemins: Tuple[str]) -> List[Tuple[str, str, str]]:
        resultats = []
        for chemin in chemins:
            match = re.search(r"\\([a-zA-Z0-9]{3})\\", chemin)
            trig = match.group(1) if match else "ERR"
            nom_fichier = os.path.basename(chemin)
            resultats.append((trig, nom_fichier, chemin))
        return resultats

    def ajuster_largeur_colonnes(self):
        style = ttk.Style()
        police_nom = style.lookup("Treeview", "font")
        police = tkfont.Font(font=police_nom) if police_nom else tkfont.nametofont("TkDefaultFont")
        for col in self.arbre_eval["columns"]:
            largeur_max = police.measure(self.arbre_eval.heading(col)["text"])
            for item in self.arbre_eval.get_children():
                texte = self.arbre_eval.set(item, col)
                largeur_texte = police.measure(texte)
                largeur_max = max(largeur_max, largeur_texte)
            self.arbre_eval.column(col, width=largeur_max + 10)

    # Scroll personnalisé ÉVALSTAT

    def _scroll_vertical(self, event):
        self.arbre_eval.yview_scroll(int(-1 * (event.delta / 120)), "units")
        return "break"

    def _scroll_horizontal(self, event):
        if event.state & 0x0001:  # Shift
            fraction = self.arbre_eval.xview()[0]
            nouvelle_valeur = max(0.0, min(1.0, fraction - event.delta / 1000))
            self.arbre_eval.xview_moveto(nouvelle_valeur)
            return "break"

    # === PARTIE BILAN ===
    def verifier_bilan_valide(self, *args):
        annee_valide = self.var_annee_bilan.get() > 2000
        trig_valide = len(self.var_trigramme_bilan.get()) == 3
        etat = "normal" if annee_valide and trig_valide else "disabled"
        self.bouton_bilan.config(state=etat)

    def generer_bilan(self):
        annee = self.var_annee_bilan.get()
        trig = self.var_trigramme_bilan.get().upper()
        self.maj_statut(f"Génération du bilan pour {trig} ({annee}) en cours...")
        # 👉 Insère ici ton vrai traitement de génération
        self.maj_statut("Génération du bilan terminée.")
    
    # === PARTIE EXPORTS IRIS ===
    def traiter_extraits_iris(self):
        selectionnes = [texte for texte, var in self.vars_iris.items() if var.get()]
        self.maj_statut(f"Traitement des extraits IRIS ({', '.join(selectionnes)}) en cours...")
        # 👉 Traitement réel ici
        self.maj_statut("Mise à jour des extraits IRIS terminée.")

    # === COMPORTEMENTS PARTIE BASSE : SORTIR et PARAMETRES  ===
    def creer_barre_boutons_bas(self):
        barre = ttk.Frame(self)
        barre.pack(side="bottom", fill="x", pady=5)

        barre.columnconfigure(0, weight=1)
        barre.columnconfigure(1, weight=1)
        barre.columnconfigure(2, weight=1)

        bouton_sortir = ttk.Button(barre, text="Sortir", command=self.quitter_application)
        bouton_sortir.grid(row=0, column=1)

        bouton_parametres = ttk.Button(barre, text="⚙️", command=self.ouvrir_fenetre_parametres)
        bouton_parametres.grid(row=0, column=2, sticky="e", padx=5)

        self.bind("<Escape>", lambda e: self.quitter_application())

    def quitter_application(self):
        self.quit()

    # === BARRE DE STATUT ===

    def creer_barre_statut(self):
        self.var_statut = tk.StringVar(value="Prêt.")

        self.cadre_statut = ttk.Frame(self, relief="sunken", borderwidth=1)
        self.cadre_statut.pack(side="bottom", fill="x")

        cadre_interne = ttk.Frame(self.cadre_statut)
        cadre_interne.pack(fill="x", padx=6, pady=3)

        self.label_statut = ttk.Label(cadre_interne, textvariable=self.var_statut, anchor="w", justify="left")
        self.label_statut.pack(side="left", fill="x", expand=True)

        self.barre_progression = ttk.Progressbar(cadre_interne, mode="determinate", maximum=100, length=150)
        self.barre_progression.pack(side="right", padx=5)
        self.barre_progression.pack_forget()

    def maj_statut(self, message: str, progression: int | None = None):
        self.var_statut.set(message)
        if progression is None:
            self.barre_progression["value"] = 0
            self.barre_progression.pack_forget()
        else:
            self.barre_progression["value"] = progression
            self.barre_progression.pack()
        self.cadre_statut.update_idletasks()




    # === FENETRE PARAMETRES ===
    def ouvrir_fenetre_parametres(self):
        fenetre = tk.Toplevel(self)
        fenetre.title("Paramètres")
        fenetre.geometry("1250x500")

        config_temp = copy.deepcopy(self.config)

        cadre_principal = ttk.Frame(fenetre)
        cadre_principal.pack(fill="both", expand=True)

        canvas = tk.Canvas(cadre_principal)
        scrollbar = ttk.Scrollbar(cadre_principal, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)

        cadre_scrollable = ttk.Frame(canvas)
        fenetre_interne = canvas.create_window((0, 0), window=cadre_scrollable, anchor="nw")

        def ajuster_largeur(event):
            canvas.itemconfig(fenetre_interne, width=event.width)

        canvas.bind("<Configure>", ajuster_largeur)

        def mise_a_jour_scrollregion(event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        cadre_scrollable.bind("<Configure>", mise_a_jour_scrollregion)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        def construire_section(parent, nom_section):
            cadre = ttk.LabelFrame(parent, text=nom_section)
            cadre.pack(fill="x", padx=10, pady=10, anchor="n")

            for cle in config_temp[nom_section]:
                ligne = ttk.Frame(cadre)
                ligne.pack(fill="x", padx=10, pady=2)
                ligne.columnconfigure(1, weight=1)

                ttk.Label(ligne, text=cle, width=45).grid(row=0, column=0, sticky="w")

                var = tk.StringVar(value=config_temp[nom_section][cle])
                entree = ttk.Entry(ligne, textvariable=var)
                entree.grid(row=0, column=1, sticky="ew", padx=5)

                def choisir_chemin(v=var, section=nom_section, k=cle):
                    initial_dir = os.path.dirname(v.get()) or "."
                    if "répertoire" in k.lower():
                        selection = filedialog.askdirectory(initialdir=initial_dir)
                    else:
                        selection = filedialog.askopenfilename(initialdir=initial_dir)
                    if selection:
                        v.set(selection)
                        config_temp[section][k] = selection
                        self.maj_statut(f"Modifié : {k}")
                        fenetre.lift()
                        fenetre.focus_force()

                ttk.Button(ligne, text="📁", command=choisir_chemin).grid(row=0, column=2)

                def maj_config(*args, v=var, section=nom_section, k=cle):
                    config_temp[section][k] = v.get()
                    self.maj_statut(f"Modifié : {k}")

                var.trace_add("write", maj_config)

        for section in config_temp:
            construire_section(cadre_scrollable, section)

        cadre_boutons = ttk.Frame(fenetre)
        cadre_boutons.pack(fill="x", pady=10)

        def annuler():
            self.maj_statut("Modifications annulées.")
            fenetre.destroy()

        def sauver_et_fermer():
            self.config = config_temp
            self.sauvegarder_config(self.config)
            self.maj_statut("Paramètres sauvegardés.")
            fenetre.destroy()

        ttk.Button(cadre_boutons, text="Sauver", style="Vert.TButton", command=sauver_et_fermer).pack(side="left", padx=10, expand=True)
        ttk.Button(cadre_boutons, text="Annuler", style="Rouge.TButton", command=annuler).pack(side="left", padx=10, expand=True)





if __name__ == "__main__":
    app = ApplicationCADI()
    app.mainloop()
