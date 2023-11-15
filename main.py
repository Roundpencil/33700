import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import numpy as np
import re
import os
import pickle
import chardet

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
            # df = pd.read_csv(self.filepath, delimiter=';', encoding='ISO-8859-1', dtype={'EMETTEUR': str, 'ALIAS_SIGNALANT': str})
            df = pd.read_csv(self.filepath, delimiter=';', encoding='ISO-8859-1', dtype=str)
            df.replace({re.compile(r'[\x01-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]'): ''}, regex=True, inplace=True)
            df['ALIAS_SIGNALANT'] = df['ALIAS_SIGNALANT'].astype(str).replace('nan', '')
            # df['ALIAS_SIGNALANT'] = df['ALIAS_SIGNALANT'].astype(str).replace('nan', '').str.rstrip('.0')

            # Charger la liste OADC
            try:
                # Détecter l'encodage du fichier
                with open('liste_oadc.csv', 'rb') as f:
                    result = chardet.detect(f.read())

                # Lire le fichier avec l'encodage détecté
                encoding = result['encoding']
                df_oadc = pd.read_csv('liste_oadc.csv', delimiter=';', encoding=encoding, dtype=str)
                # df_oadc = pd.read_csv('liste_oadc.csv', delimiter=';', dtype=str)

                oadc_list = df_oadc['OADC'].tolist()
            except FileNotFoundError:
                oadc_list = []

            print(oadc_list)

            # Chercher dans la colonne EMETTEUR
            df['expediteur_nettoye'] = ""
            df['typologie_expediteur'] = ""
            df['rebond_nettoye'] = ""
            df['typologie_rebond'] = ""
            df['date_requalifiee'] = ""
            df['categorie_no_cible'] = ""
            df['mois'] = ""

            # identification des opérateurs de l'éxpéditeur
            majnum = pd.read_excel('MAJNUM.xls')

            # Détecter l'encodage du fichier
            with open('identifiants_CE.csv', 'rb') as f:
                result = chardet.detect(f.read())

            # Lire le fichier avec l'encodage détecté
            encoding = result['encoding']
            identifiants_CE = pd.read_csv('identifiants_CE.csv', delimiter=';', encoding=encoding, dtype=str)

            df['operateur_arcep'] = ""

            # changer d'ordre des colonnes
            column_order = ['DATE_SIGNALEMENT', 'MESSAGE', 'EMETTEUR', 'ALIAS_SIGNALANT', 'NUMERO_REBOND_SIGNAL',
                            'OPERATEUR_SIGNALANT', 'URL_REBOND_SIGNALE', 'date_requalifiee', 'expediteur_nettoye',
                            'typologie_expediteur', 'operateur_arcep', 'typologie_rebond', 'categorie_no_cible',
                            'categorie_no_cible', 'mois',
                            'DATE_RECEPTION',
                            'MOIS_RECEPTION',
                            'ANALYSE_STOP', 'TYPE_EMETTEUR']

            # Check if all columns in column_order are present in df.columns
            if set(column_order).issubset(df.columns):
                remaining_columns = [col for col in df.columns if col not in column_order]
                new_order = column_order + remaining_columns
                df = df[new_order]
            else:
                print("Some columns are missing from the dataframe")

            # df['EMETTEUR'] = df['EMETTEUR'].replace('nan', '')
            for i, row in df.iterrows():
                #extraction de la date
                date_full = row['DATE_SIGNALEMENT']
                df.at[i, 'date_requalifiee'] = date_full[:10]
                df.at[i, 'mois'] = date_full[:7]

                # extraction de l'émetteur et
                emetteur = row['EMETTEUR']
                print(f"emetteur en cours : {emetteur}", end=' ')
                if pd.isna(emetteur):
                    df.at[i, 'expediteur_nettoye'] = ""
                    df.at[i, 'typologie_expediteur'] = "Non identifié"
                else:
                    numero = extraire_numero_de_texte(emetteur)
                    if numero:
                        # numero = normaliser_numero(match.group())
                        df.at[i, 'expediteur_nettoye'] = numero
                        df.at[i, 'typologie_expediteur'] = typologie_numero(numero)
                        print(f" - numero nettoyé = {numero} - typologie : {typologie_numero(numero)}")

                        #identificaiton opérateur
                        operateur = trouver_operateur(numero, majnum, identifiants_CE)
                        print(f'Opérateur trouve : {operateur}')
                        df.at[i, 'operateur_arcep'] = operateur
                    elif emetteur in oadc_list:
                        df.at[i, 'typologie_expediteur'] = "OADC"
                        df.at[i, 'expediteur_nettoye'] = emetteur
                        print(f" - numero nettoyé = {emetteur} : typologie : OADC")

                    else:
                        df.at[i, 'typologie_expediteur'] = "Non identifié"
                        print(f" - non identifié = {emetteur}")

                # extraction du numéro de rebond du message
                texte_message = row['MESSAGE']
                if pd.isna(texte_message):
                    df.at[i, 'rebond_nettoye'] = ""
                    df.at[i, 'typologie_rebond'] = "Aucun"
                else:
                    numero_rebond = extraire_numero_de_texte(texte_message)
                    # message_nettoye = texte_message.replace(' ', '').replace('.', '').replace('-', '')
                    # match = re.search(r'33\d{13}|0\d{13}|33\d{9}|0\d{9}|118\d{6}|\d{5}|\d{4}',
                    #                   message_nettoye)
                    # if match:
                    if numero_rebond:
                        typologie_rebond = typologie_numero(numero_rebond)
                        # numero = normaliser_numero(match.group())
                        df.at[i, 'rebond_nettoye'] = numero_rebond
                        df.at[i, 'typologie_rebond'] = typologie_rebond
                        print(f" - rebond nettoyé = {numero_rebond} - typologie : {typologie_rebond}")

                # print(row['URL_REBOND_SIGNALE'])
                if pd.isna(row['URL_REBOND_SIGNALE']):
                    # df.at[i, 'categorie_no_cible'] = row['typologie_rebond']
                    df.at[i, 'categorie_no_cible'] = df.at[i, 'typologie_rebond']
                else:
                    df.at[i, 'categorie_no_cible'] = 'URL'



                # Save to Excel
            outfile = os.path.join(self.outdir, os.path.basename(self.filepath).split('.')[0] + '.xlsx')
            df.to_excel(outfile, index=False, engine='openpyxl')

            # Inform the user
            messagebox.showinfo("Succès", "Conversion réussie!")
        # except Exception as e:
        #     messagebox.showerror("Erreur", str(e))
        else:
            messagebox.showerror("Erreur", "Veuillez choisir un fichier d'entrée CSV")


def extraire_numero_de_texte(texte_source):
    # sourcery skip: use-named-expression
    # source_sans_espace = texte_source.replace(' ', '')  # Retire tous les espaces de la chaîne
    source_sans_espace = texte_source.replace(' ', '').replace('.', '').replace('-', '')  # Retire tous les séparateurs usuels de la chaine
    # match = re.search(r'33\d{9}|0\d{9}|\d{9}|118\d{6}|\d{5}|\d{4}', emetteur_sans_espace)
    # match = re.search(r'33\d{9}|0\d{9}|\d{9}|118\d{6}|\d{5}|\d{4}|\d{14}|33\d{12}|\d{13}',
    # match = re.search(r'33\d{13}|0\d{13}|\d{13}|33\d{9}|0\d{9}|\d{9}|118\d{3}|\d{5}|\d{4}',
    #                   source_sans_espace)
    match = re.search(r'(00)?33700\d{10}|0700\d{10}|700\d{10}|(00)?33\d{9}|0\d{9}|\d{9}|118\d{3}|\d{5}|\d{4}',
                      source_sans_espace)
    if not match:
        return extraire_numero_international(source_sans_espace)

    numero_brut_trouve = match.group()

    matches_consecutifs = re.findall(r'\d+', source_sans_espace)
    tailles_consecutifs = [len(match) for match in matches_consecutifs]

    if len(numero_brut_trouve) not in tailles_consecutifs:
        return extraire_numero_international(source_sans_espace)

    return normaliser_numero(numero_brut_trouve)


def extraire_numero_international(source_sans_espace):
    match_international = re.search(r'\+\d{5,}|00\d{5,}', source_sans_espace)
    return match_international.group() if match_international else None


def normaliser_numero(numero):
    if numero.startswith('00'):
        numero = numero[2:]

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


def typologie_numero(numero):
    if numero.startswith('+') or numero.startswith('00'):
        return "International"
    if len(numero) == 14 and numero.startswith("0700"):
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
    elif len(numero) == 6 and numero.startswith('118'):
        return 'SVA'
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
    elif len(numero) == 4:
        if numero.startswith('3'):
            return 'SVA'
        elif numero[0] in ['1', '2']:
            return "Numéro opérateur"
    else:
        return "Non identifié"


# Fonction pour trouver l'opérateur
def trouver_operateur(numero, majnum, majrio):
    print(f"\n\n numéro en cours : {numero} : ")
    try:
        # Tentez de convertir 'numero' en entier
        numero = int(numero)
    except ValueError:
        # print(f"Le numéro {numero} n'est pas un nombre, il n'a pas d'opérateur")
        return ''

    try:
        tranche = majnum[(majnum['Tranche_Debut'] <= numero) & (majnum['Tranche_Fin'] >= numero)]
        print(tranche)
        if not tranche.empty:
            mnemo = tranche['Mnémo'].iloc[0]
            operateur = majrio[majrio['CODE_OPERATEUR'] == mnemo]['IDENTITE_OPERATEUR'].iloc[0]
            # print(f'opérateur trouvé pour le numéro {numero}: {operateur}')
            return operateur
        else:
            return 'Inconnu'
    except Exception as e:
        print(f'erreur durant trouver_operateur : {e}')
        return ''

root = tk.Tk()
app = Application(master=root)
app.mainloop()
