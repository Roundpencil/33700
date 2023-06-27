import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import numpy as np
import re
import os
import pickle


class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.grid()
        self.outdir = self.load_last_dir()
        self.create_widgets()
        self.filepath = ""
        self.master.title("La moulinette 33700 de l'af2m")


    def create_widgets(self):
        self.lbl1 = tk.Label(self, text="1. Choisissez le fichier d'entrée CSV : (aucun fichier)")
        self.lbl1.grid(row=0, column=0, sticky='w')

        self.file_button = tk.Button(self, text="Choisir fichier", command=self.load_file)
        self.file_button.grid(row=0, column=1, sticky='w')

        self.lbl2 = tk.Label(self, text=f"2. Choisissez le dossier de sortie (facultatif) : {self.outdir}")
        self.lbl2.grid(row=1, column=0, sticky='w')

        self.dir_button = tk.Button(self, text="Choisir dossier", command=self.load_dir)
        self.dir_button.grid(row=1, column=1, sticky='w')

        self.convert_button = tk.Button(self, text="Convertir !", command=self.convert)
        self.convert_button.grid(row=2, column=0, columnspan=2)

    def load_file(self):
        self.filepath = filedialog.askopenfilename(filetypes=[("Fichiers CSV", "*.csv")])
        self.lbl1.config(
            text=f"1. Choisissez le fichier d'entrée CSV : {self.filepath if self.filepath else '(aucun fichier)'}")

    def load_dir(self):
        self.outdir = filedialog.askdirectory()
        self.lbl2.config(text=f"2. Choisissez le dossier de sortie (facultatif) : {self.outdir}")
        self.save_last_dir(self.outdir)

    def load_last_dir(self):
        try:
            with open("last_dir.pkl", "rb") as f:
                return pickle.load(f)
        except (FileNotFoundError, EOFError):
            default_dir = os.path.join(os.path.dirname(__file__), "Fichiers sortie Excel")
            os.makedirs(default_dir, exist_ok=True)  # Creates the default directory if it doesn't exist
            return default_dir

    def save_last_dir(self, path):
        with open("last_dir.pkl", "wb") as f:
            pickle.dump(path, f)

    def convert(self):  # sourcery skip: use-named-expression
        if self.filepath:
            # try:
                # Load and process the CSV
                # df = pd.read_csv(self.filepath, delimiter=';', encoding='ISO-8859-1')
                df = pd.read_csv(self.filepath, delimiter=';', encoding='ISO-8859-1', dtype={'EMETTEUR': str, 'ALIAS_SIGNALANT': str})
                df.replace({re.compile(r'[\x01-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]'): ''}, regex=True, inplace=True)
                df['ALIAS_SIGNALANT'] = df['ALIAS_SIGNALANT'].astype(str).replace('nan', '').str.rstrip('.0')

                # Charger la liste OADC
                try:
                    df_oadc = pd.read_csv('liste_oadc.csv', dtype=str)
                    oadc_list = df_oadc['OADC'].tolist()
                except FileNotFoundError:
                    oadc_list = []

                print(oadc_list)

                # Chercher dans la colonne EMETTEUR
                df['expediteur_nettoye'] = ""
                df['typologie_expediteur'] = ""
                # df['EMETTEUR'] = df['EMETTEUR'].replace('nan', '')
                for i, row in df.iterrows():
                    emetteur = row['EMETTEUR']
                    print(f"emetteur en cours : {emetteur}", end=' ')
                    if pd.isna(emetteur):
                        df.at[i, 'expediteur_nettoye'] = ""
                        df.at[i, 'typologie_expediteur'] = "Non identifié"
                        continue
                    emetteur_sans_espace = emetteur.replace(' ', '')  # Retire tous les espaces de la chaîne

                    # match = re.search(r'33\d{9}|0\d{9}|\d{9}|118\d{6}|\d{5}|\d{4}', emetteur_sans_espace)
                    # match = re.search(r'33\d{9}|0\d{9}|\d{9}|118\d{6}|\d{5}|\d{4}|\d{14}|33\d{12}|\d{13}',
                    match = re.search(r'33\d{13}|0\d{13}|\d{13}|33\d{9}|0\d{9}|\d{9}|118\d{6}|\d{5}|\d{4}',
                                      emetteur_sans_espace)
                    if match:
                        numero = normaliser_numero(match.group())
                        df.at[i, 'expediteur_nettoye'] = numero
                        df.at[i, 'typologie_expediteur'] = typologie_emetteur(numero)
                        print(f" - numero nettoyé = {numero} - typologie : {typologie_emetteur(numero)}")
                    elif emetteur in oadc_list:
                        df.at[i, 'typologie_expediteur'] = "OADC"
                        df.at[i, 'expediteur_nettoye'] = emetteur
                        print(f" - numero nettoyé = {emetteur} : typologie : OADC")

                    else:
                        df.at[i, 'typologie_expediteur'] = "Non identifié"
                        print(f" - non identifié = {emetteur}")


                # Save to Excel
                outfile = os.path.join(self.outdir, os.path.basename(self.filepath).split('.')[0] + '.xlsx')
                df.to_excel(outfile, index=False, engine='openpyxl')

                # Inform the user
                messagebox.showinfo("Succès", "Conversion réussie!")
            # except Exception as e:
            #     messagebox.showerror("Erreur", str(e))
        else:
            messagebox.showerror("Erreur", "Veuillez choisir un fichier d'entrée CSV")

def normaliser_numero(numero):
    if len(numero) == 15 and numero.startswith("33"):
        return "0" + numero[2:]
    elif len(numero) == 13:
        return "0" + numero
    if len(numero) == 11 and numero.startswith("33"):
        return "0" + numero[2:]
    elif len(numero) == 9:
        return "0" + numero
    else:
        return numero

def typologie_emetteur(numero):
    if len(numero) == 14 and (numero.startswith("06") or numero.startswith("07")):
        return "M2M"
    if len(numero) == 10 and numero.startswith("0"):
        if numero[:2] == "09":
            return "'09"
        elif numero[:2] in ["06", "07"]:
            return "MSISDN"
        elif numero[:2] == "08":
            return "SVA"
        else:
            return "Géographique"
    elif len(numero) == 5:
        if numero == "33700":
            return "33700"
        elif numero[:2] in ["36", "37", "38"]:
            return "Shortcode BM"
        elif "30" <= numero[:2] <= "94":
            return "SMS+"
        elif numero[:1] in ["1", "2"]:
            return "Numéro opérateur"
        else:
            return "Autres ABDCE"
    else:
        return "Non identifié"

root = tk.Tk()
app = Application(master=root)
app.mainloop()
