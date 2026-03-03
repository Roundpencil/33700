import argparse
import os
import json
import pickle
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk

def print_gui(args=None):

    root = tk.Tk()
    root.title("La moulinette 33700 de l'af2m")
    root.grid()

    if not args:
        args = argparse.Namespace(dossier_sortie=".", fichier_entree=".", db_path=".")
    dossier_sortie_var = tk.StringVar(value=getattr(args, "dossier_sortie", ""))
    filepath_var = tk.StringVar(value=getattr(args, "fichier_entree", ""))
    db_path_var = tk.StringVar(value=getattr(args, "db_path", ""))

    def browse_directory(var: tk.StringVar):
        dirname = filedialog.askdirectory()
        if dirname:
            var.set(dirname)

    def browse_file(var: tk.StringVar):
        filename = filedialog.askopenfilename()
        if filename:
            var.set(filename)

    ttk.Label(root, text="Fichier csv entrée :").grid(row=0, column=0, sticky="e", padx=4, pady=4)
    ttk.Entry(root, textvariable=filepath_var, width=50).grid(row=0, column=1, sticky="we", padx=4, pady=4)
    ttk.Button(root, text="Parcourir", command=lambda: browse_file(filepath_var)).grid(row=0, column=2, padx=4, pady=4)

    ttk.Label(root, text="Base de donnée :").grid(row=1, column=0, sticky="e", padx=4, pady=4)
    ttk.Entry(root, textvariable=db_path_var, width=50).grid(row=1, column=1, sticky="we", padx=4, pady=4)
    ttk.Button(root, text="Parcourir", command=lambda: browse_file(db_path_var)).grid(row=1, column=2, padx=4, pady=4)

    ttk.Label(root, text="Dossier de sortie :").grid(row=2, column=0, sticky="e", padx=4, pady=4)
    ttk.Entry(root, textvariable=dossier_sortie_var, width=50).grid(row=2, column=1, sticky="we", padx=4, pady=4)
    ttk.Button(root, text="Parcourir", command=lambda: browse_directory(dossier_sortie_var)).grid(row=2, column=2, padx=4, pady=4)

    ttk.Button(root, text="Convertir!", command=lambda: killandreturn()).grid(row=8, column=1, pady=10)

    def killandreturn():
        args.dossier_sortie = dossier_sortie_var.get()
        args.fichier_entree = filepath_var.get()
        args.db_path = db_path_var.get()
        root.destroy()

    root.mainloop()
    return args

    #
    # def create_widgets(self):
    #     self.lbl1 = tk.Label(self, text="1. Choisissez le fichier d'entrée CSV : (aucun fichier)")
    #     self.lbl1.grid(row=0, column=0, sticky='w')
    #
    #     self.file_button = tk.Button(self, text="Choisir fichier", command=self.load_file)
    #     self.file_button.grid(row=0, column=1, sticky='w')
    #
    #     self.lbl2 = tk.Label(self, text=f"2. Choisissez le dossier de sortie (facultatif) : {self.outdir}")
    #     self.lbl2.grid(row=1, column=0, sticky='w')
    #
    #     self.dir_button = tk.Button(self, text="Choisir dossier", command=self.load_dir)
    #     self.dir_button.grid(row=1, column=1, sticky='w')
    #
    #     self.convert_button = tk.Button(self, text="Convertir !", command=self.convert)
    #     self.convert_button.grid(row=2, column=0, columnspan=2)
    #
    # def load_file(self):
    #     self.filepath = filedialog.askopenfilename(filetypes=[("Fichiers CSV", "*.csv")])
    #     self.lbl1.config(
    #         text=f"1. Choisissez le fichier d'entrée CSV : {self.filepath if self.filepath else '(aucun fichier)'}")
    #
    # def load_dir(self):
    #     self.outdir = filedialog.askdirectory()
    #     self.lbl2.config(text=f"2. Choisissez le dossier de sortie (facultatif) : {self.outdir}")
    #     self.save_last_dir(self.outdir)

if __name__ == "__main__":
    args = print_gui()
    print(args)