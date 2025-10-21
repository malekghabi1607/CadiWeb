import inspect
import tkinter as tk
from tkinter import simpledialog, messagebox

class IHMTkinter:
    def __init__(self, menus, ca):
        self.menus = menus
        self.ca = ca
        self.root = tk.Tk()
        self.root.title("Gestion Tutorat")

    def demander_arguments(self, action):
        sig = inspect.signature(action)
        kwargs = {}
        for nom, param in sig.parameters.items():
            type_attendu = param.annotation if param.annotation != inspect._empty else str
            val = simpledialog.askstring("Entrée requise", f"{nom} ({type_attendu.__name__}) :")
            if val is None:
                return None
            try:
                kwargs[nom] = type_attendu(val)
            except ValueError:
                messagebox.showerror("Erreur", f"Valeur invalide pour {nom}.")
                return None
        return kwargs

    def executer_action(self, action):
        kwargs = self.demander_arguments(action)
        if kwargs is not None:
            action(**kwargs)
            messagebox.showinfo("Succès", "Action exécutée avec succès")

    def afficher_menu(self, menu=None):
        if menu is None:
            menu = self.menus

        win = tk.Toplevel(self.root)
        win.title("Menu")

        for cle, action in menu.items():
            if isinstance(action, dict):
                tk.Button(win, text=cle, command=lambda m=action: self.afficher_menu(m)).pack(pady=2)
            else:
                tk.Button(win, text=cle, command=lambda a=action: self.executer_action(a)).pack(pady=2)

    def lancer(self):
        tk.Button(self.root, text="Ouvrir menu", command=lambda: self.afficher_menu()).pack(pady=20)
        self.root.mainloop()