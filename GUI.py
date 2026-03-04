import argparse
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk

def print_gui(args=None):

    root = tk.Tk()
    root.title("La moulinette 33700 de l'af2m")
    root.grid()

    if not args:
        args = argparse.Namespace(dossier_sortie=".", fichier_entree=".", db_path=".",
                                  ajouter_operateurs=True, score_phising=True)

    print(args)
    dossier_sortie_var = tk.StringVar(value=getattr(args, "dossier_sortie", ""))
    filepath_var = tk.StringVar(value=getattr(args, "fichier_entree", ""))
    db_path_var = tk.StringVar(value=getattr(args, "db_path", ""))
    ajouter_operateurs_var = tk.BooleanVar(value=getattr(args, "ajouter_operateurs", True))
    score_phising_var = tk.BooleanVar(value=getattr(args, "score_phising", True))
    analyser_si_isa_var = tk.BooleanVar(value=getattr(args, "analyser_si_isa", True))
    contenu_xls_var = tk.StringVar(value=getattr(args, "contenu_xls", "complet"))

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

    ttk.Checkbutton(
        root,
        text="Ajouter les opérateurs des numéros de rebond",
        variable=ajouter_operateurs_var
    ).grid(row=3, column=1, sticky="w", padx=4, pady=4)

    ttk.Checkbutton(
        root,
        text="Calculer le score phishing",
        variable=score_phising_var
    ).grid(row=4, column=1, sticky="w", padx=4, pady=4)

    ttk.Checkbutton(
        root,
        text="Ajouter le type de protection",
        variable=analyser_si_isa_var
    ).grid(row=5, column=1, sticky="w", padx=4, pady=4)

    ttk.Label(root, text="Contenu du fichier xls généré").grid(row=6, column=0, sticky="w", padx=4, pady=4)

    ttk.Radiobutton(
        root,
        text="Créer un fichier complet",
        variable=contenu_xls_var,
        value="complet"
    ).grid(row=6, column=1, sticky="w", padx=4, pady=4)

    ttk.Radiobutton(
        root,
        text="Exporter uniquement les nouvelles lignes",
        variable=contenu_xls_var,
        value="nouveaux"
    ).grid(row=7, column=1, sticky="w", padx=4, pady=4)

    ttk.Button(root, text="Convertir!", command=lambda: killandreturn()).grid(row=8, column=1, pady=10)

    def killandreturn():
        args.dossier_sortie = dossier_sortie_var.get()
        args.fichier_entree = filepath_var.get()
        args.db_path = db_path_var.get()
        args.ajouter_operateurs = ajouter_operateurs_var.get()
        args.score_phising = score_phising_var.get()
        args.analyser_si_isa = analyser_si_isa_var.get()
        args.contenu_xls = contenu_xls_var.get()
        root.destroy()

    root.mainloop()
    return args


if __name__ == "__main__":
    args = print_gui()
    print(args)