import tkinter as tk
from tkinter import filedialog
import argparse
import os
import json

CONFIG_FILE = 'config.json'

def print_GUI(args):
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r') as f:
            config = json.load(f)
        args.csv_input = config.get('csv_input', args.csv_input)
        args.output_dir = config.get('output_dir', args.output_dir)
        args.mode = config.get('mode', args.mode)
        args.db_path = config.get('db_path', args.db_path)

    root = tk.Tk()
    root.title("Configurer les options")

    csv_input_var = tk.StringVar(value=args.csv_input)
    output_dir_var = tk.StringVar(value=args.output_dir)
    mode_var = tk.StringVar(value=args.mode)
    db_path_var = tk.StringVar(value=args.db_path)

    def browse_file(var):
        filename = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"), ("Tous les fichiers", "*.*")])
        if filename:
            var.set(filename)

    def browse_directory(var):
        dirname = filedialog.askdirectory()
        if dirname:
            var.set(dirname)

    def update_db_field():
        if mode_var.get() == "af2m":
            db_entry.config(state='normal')
            db_button.config(state='normal')
        else:
            db_entry.config(state='disabled')
            db_button.config(state='disabled')
            db_path_var.set('')

    def on_ok():
        if not csv_input_var.get():
            messagebox.showerror("Erreur", "Veuillez sélectionner un fichier CSV.")
            return
        if not output_dir_var.get():
            messagebox.showerror("Erreur", "Veuillez sélectionner un dossier de sortie.")
            return
        if not mode_var.get():
            messagebox.showerror("Erreur", "Veuillez choisir un mode (opérateurs ou af2m).")
            return
        if mode_var.get() == "af2m" and not db_path_var.get():
            messagebox.showerror("Erreur", "Veuillez sélectionner un fichier base de données pour le mode AF2M.")
            return
        root.quit()

    # --- Interface graphique ---
    tk.Label(root, text="Fichier CSV d'entrée:").grid(row=0, column=0, sticky='e')
    tk.Entry(root, textvariable=csv_input_var, width=50).grid(row=0, column=1)
    tk.Button(root, text="Parcourir", command=lambda: browse_file(csv_input_var)).grid(row=0, column=2)

    tk.Label(root, text="Dossier de sortie:").grid(row=1, column=0, sticky='e')
    tk.Entry(root, textvariable=output_dir_var, width=50).grid(row=1, column=1)
    tk.Button(root, text="Parcourir", command=lambda: browse_directory(output_dir_var)).grid(row=1, column=2)

    tk.Label(root, text="Mode:").grid(row=2, column=0, sticky='e')
    tk.Radiobutton(root, text="Opérateurs", variable=mode_var, value='operateurs', command=update_db_field).grid(row=2, column=1, sticky='w')
    tk.Radiobutton(root, text="AF2M", variable=mode_var, value='af2m', command=update_db_field).grid(row=2, column=1, sticky='e')

    # Champ Base de données (activé seulement en mode af2m)
    tk.Label(root, text="Fichier base de données:").grid(row=3, column=0, sticky='e')
    db_entry = tk.Entry(root, textvariable=db_path_var, width=50, state='disabled')
    db_entry.grid(row=3, column=1)
    db_button = tk.Button(root, text="Parcourir", command=lambda: browse_file(db_path_var), state='disabled')
    db_button.grid(row=3, column=2)

    tk.Button(root, text="OK", command=on_ok).grid(row=4, column=1)

    # Appliquer l'état initial selon le mode
    update_db_field()

    root.mainloop()

    args.csv_input = csv_input_var.get()
    args.output_dir = output_dir_var.get()
    args.mode = mode_var.get()
    args.db_path = db_path_var.get()

    root.destroy()

    config = {
        'csv_input': args.csv_input,
        'output_dir': args.output_dir,
        'mode': args.mode,
        'db_path': args.db_path
    }
    with open(CONFIG_FILE, 'w') as f:
        json.dump(config, f)

    return args


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Script avec interface graphique ou CLI pour choisir des paramètres.")
    parser.add_argument('--console', action='store_true', help="Utiliser uniquement l'interface en ligne de commande")

    parser.add_argument('--csv_input', type=str, default='', help="Chemin vers le fichier CSV d'entrée")
    parser.add_argument('--output_dir', type=str, default='', help="Dossier de sortie")
    parser.add_argument('--mode', type=str, choices=['operateurs', 'af2m'], default='', help="Mode à utiliser")
    parser.add_argument('--db_path', type=str, default='',
                        help="Chemin vers le fichier base de données (obligatoire si mode af2m)")

    args = parser.parse_args()

    if not args.console:
        args = print_GUI(args)
    else:
        missing_fields = []
        if not args.csv_input:
            missing_fields.append("csv_input")
        if not args.output_dir:
            missing_fields.append("output_dir")
        if not args.mode:
            missing_fields.append("mode")
        if args.mode == "af2m" and not args.db_path:
            missing_fields.append("db_path (requis pour le mode af2m)")

        if missing_fields:
            print(f"Erreur : les champs suivants sont obligatoires : {', '.join(missing_fields)}")
            exit(1)

    # ✅ Utilisation des valeurs collectées
    print("Fichier CSV :", args.csv_input)
    print("Dossier de sortie :", args.output_dir)
    print("Mode sélectionné :", args.mode)

