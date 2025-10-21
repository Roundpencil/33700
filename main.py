import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import numpy as np
import re
import os
import pickle
import chardet
import sqlite3
from datetime import datetime
import threading
import queue
# Dépendance nécessaire pour l'export Excel, à installer via : pip install openpyxl
import openpyxl


class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.grid()
        self.outdir = self.load_last_dir()
        self.create_widgets()
        self.filepath = ""
        self.master.title("La moulinette 33700 de l'af2m - Version DB")
        self.db_path = None
        self.setup_database_path()

        # Pour le traitement asynchrone
        self.processing_queue = queue.Queue()
        self.is_processing = False

    def setup_database_path(self):
        """Configure le chemin de la base de données et vérifie son existence"""
        self.db_path = os.path.join(self.outdir, "moulinette_data.db")
        self.update_db_status()

    def update_db_status(self):
        """Met à jour l'affichage du statut de la base de données"""
        if os.path.exists(self.db_path):
            try:
                conn = sqlite3.connect(self.db_path)
                cursor = conn.cursor()
                cursor.execute("SELECT COUNT(*) FROM signalements")
                count = cursor.fetchone()[0]
                conn.close()
                status = f"Base existante : {count:,} enregistrements"
            except sqlite3.Error:  # Correction : Être spécifique sur l'erreur
                status = "Base existante (erreur lecture)"
        else:
            status = "Base de données à créer"

        self.db_status_label.config(text=f"Statut DB : {status}")

    def create_database(self):
        """Crée une nouvelle base de données"""
        if os.path.exists(self.db_path):
            if not messagebox.askyesno("Confirmation",
                                       "Une base de données existe déjà. La remplacer ?"):
                return
            os.remove(self.db_path)

        self.init_database()
        messagebox.showinfo("Succès", "Base de données créée avec succès !")
        self.update_db_status()

    def init_database(self):
        """Initialise la base de données avec les tables nécessaires"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        # Table principale pour les données traitées
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS signalements (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date_signalement TEXT,
                message TEXT,
                emetteur TEXT,
                alias_signalant TEXT,
                numero_rebond_signal TEXT,
                operateur_signalant TEXT,
                url_rebond_signale TEXT,
                date_requalifiee TEXT,
                expediteur_nettoye TEXT,
                typologie_expediteur TEXT,
                operateur_arcep TEXT,
                rebond_nettoye TEXT,
                typologie_rebond TEXT,
                categorie_no_cible TEXT,
                mois TEXT,
                date_reception TEXT,
                mois_reception TEXT,
                analyse_stop TEXT,
                type_emetteur TEXT,
                type_protection TEXT,
                opr_arcep_rebond TEXT,
                fichier_source TEXT,
                date_import TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(date_signalement, emetteur, message) ON CONFLICT IGNORE
            )
        ''')

        # Table de métadonnées pour tracer les fichiers traités
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS fichiers_traites (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nom_fichier TEXT UNIQUE,
                chemin_complet TEXT,
                taille_fichier INTEGER,
                nb_lignes_traitees INTEGER,
                date_traitement TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                statut TEXT DEFAULT 'Terminé'
            )
        ''')

        # Index pour améliorer les performances
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_date_signalement ON signalements(date_signalement)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_expediteur ON signalements(expediteur_nettoye)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_mois ON signalements(mois)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_fichier_source ON signalements(fichier_source)')

        conn.commit()
        conn.close()

    def create_widgets(self):
        # ... (Le code des widgets reste inchangé) ...
        # Section fichier d'entrée
        self.lbl1 = tk.Label(self, text="1. Choisissez le fichier d'entrée CSV : (aucun fichier)")
        self.lbl1.grid(row=0, column=0, columnspan=2, sticky='w', padx=5, pady=5)

        self.file_button = tk.Button(self, text="Choisir fichier", command=self.load_file)
        self.file_button.grid(row=0, column=2, sticky='w', padx=5, pady=5)

        # Section dossier de base de données
        self.lbl2 = tk.Label(self, text=f"2. Dossier de base de données : {self.outdir}")
        self.lbl2.grid(row=1, column=0, columnspan=2, sticky='w', padx=5, pady=5)

        self.dir_button = tk.Button(self, text="Changer dossier", command=self.load_dir)
        self.dir_button.grid(row=1, column=2, sticky='w', padx=5, pady=5)

        # Section statut de la base de données
        self.db_status_label = tk.Label(self, text="Statut DB : Vérification...")
        self.db_status_label.grid(row=2, column=0, columnspan=2, sticky='w', padx=5, pady=5)

        db_button_frame = tk.Frame(self)
        db_button_frame.grid(row=2, column=2, sticky='w', padx=5, pady=5)

        self.create_db_button = tk.Button(db_button_frame, text="Créer DB",
                                          command=self.create_database, bg='lightgreen')
        self.create_db_button.pack(side=tk.LEFT, padx=2)

        self.backup_db_button = tk.Button(db_button_frame, text="Sauvegarder",
                                          command=self.backup_database)
        self.backup_db_button.pack(side=tk.LEFT, padx=2)

        # Section options de traitement
        self.lbl3 = tk.Label(self, text="3. Options de traitement :")
        self.lbl3.grid(row=3, column=0, sticky='w', padx=5, pady=5)

        options_frame = tk.Frame(self)
        options_frame.grid(row=4, column=0, columnspan=3, sticky='ew', padx=5)

        tk.Label(options_frame, text="Taille de lot CSV :").grid(row=0, column=0, sticky='w', padx=5)
        self.batch_size_var = tk.StringVar(value="10000")
        tk.Entry(options_frame, textvariable=self.batch_size_var, width=10).grid(row=0, column=1, sticky='w')

        tk.Label(options_frame, text="Taille insertion DB :").grid(row=0, column=2, sticky='w', padx=5)
        self.sql_batch_size_var = tk.StringVar(value="100")
        tk.Entry(options_frame, textvariable=self.sql_batch_size_var, width=10).grid(row=0, column=3, sticky='w')

        self.skip_duplicates_var = tk.BooleanVar(value=True)
        tk.Checkbutton(options_frame, text="Ignorer les doublons",
                       variable=self.skip_duplicates_var).grid(row=1, column=0, columnspan=2, sticky='w', padx=5)

        # Barre de progression
        self.progress = ttk.Progressbar(self, mode='indeterminate')
        self.progress.grid(row=5, column=0, columnspan=3, sticky='ew', padx=5, pady=5)

        self.status_label = tk.Label(self, text="Prêt")
        self.status_label.grid(row=6, column=0, columnspan=3, sticky='w', padx=5)

        # Boutons principaux
        button_frame = tk.Frame(self)
        button_frame.grid(row=7, column=0, columnspan=3, pady=10)

        self.convert_button = tk.Button(button_frame, text="Traiter et Stocker",
                                        command=self.start_processing, bg='lightblue')
        self.convert_button.pack(side=tk.LEFT, padx=5)

        self.export_button = tk.Button(button_frame, text="Exporter vers Excel",
                                       command=self.export_to_excel)
        self.export_button.pack(side=tk.LEFT, padx=5)

        self.stats_button = tk.Button(button_frame, text="Voir Statistiques",
                                      command=self.show_stats)
        self.stats_button.pack(side=tk.LEFT, padx=5)

        self.files_button = tk.Button(button_frame, text="Fichiers traités",
                                      command=self.show_processed_files)
        self.files_button.pack(side=tk.LEFT, padx=5)

        self.clear_button = tk.Button(button_frame, text="Vider DB",
                                      command=self.clear_database, bg='lightcoral')
        self.clear_button.pack(side=tk.LEFT, padx=5)

    def insert_data_in_batches(self, conn, df, table_name):
        """OPTIMISATION: Insère les données par lots calculés pour éviter l'erreur 'too many SQL variables'"""
        if df.empty:
            return 0

        total_rows = len(df)
        columns = df.columns.tolist()
        num_columns = len(columns)

        # La limite par défaut de SQLite est 999. On calcule la taille de lot max pour s'y conformer.
        sql_batch_size = min(int(self.sql_batch_size_var.get()), 999 // num_columns if num_columns > 0 else 999)
        if sql_batch_size == 0: sql_batch_size = 1  # Au cas où il y aurait plus de 999 colonnes.

        placeholders = ','.join(['?' for _ in columns])
        column_names = ','.join(columns)

        query = f"INSERT OR IGNORE INTO {table_name} ({column_names}) VALUES ({placeholders})"

        cursor = conn.cursor()
        inserted_count = 0

        for start_idx in range(0, total_rows, sql_batch_size):
            end_idx = min(start_idx + sql_batch_size, total_rows)
            batch_df = df.iloc[start_idx:end_idx]

            data_tuples = [tuple(None if pd.isna(val) else val for val in row) for row in
                           batch_df.itertuples(index=False)]

            try:
                cursor.executemany(query, data_tuples)
                inserted_count += cursor.rowcount
                conn.commit()
                self.processing_queue.put(('status', f"Inséré en DB : {inserted_count:,} lignes"))
            except sqlite3.Error as e:
                print(f"Erreur lors de l'insertion du lot {start_idx}-{end_idx}: {e}")
                conn.rollback()  # Annuler la transaction de lot en échec
                continue

        return inserted_count

    # ... (les fonctions backup_database, load_file, check_file_already_processed, load_dir, load_last_dir, save_last_dir, start_processing restent majoritairement inchangées) ...
    def backup_database(self):
        """Crée une sauvegarde de la base de données"""
        if not os.path.exists(self.db_path):
            messagebox.showwarning("Attention", "Aucune base de données à sauvegarder")
            return

        backup_path = filedialog.asksaveasfilename(
            defaultextension=".db",
            filetypes=[("Base de données", "*.db")],
            initialname=f"moulinette_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db"
        )

        if backup_path:
            try:
                import shutil
                shutil.copy2(self.db_path, backup_path)
                messagebox.showinfo("Succès", f"Sauvegarde créée : {backup_path}")
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur lors de la sauvegarde : {str(e)}")

    def load_file(self):
        self.filepath = filedialog.askopenfilename(filetypes=[("Fichiers CSV", "*.csv")])
        display_name = os.path.basename(self.filepath) if self.filepath else "(aucun fichier)"
        self.lbl1.config(text=f"1. Choisissez le fichier d'entrée CSV : {display_name}")

        if self.filepath and os.path.exists(self.db_path):
            self.check_file_already_processed()

    def check_file_already_processed(self):
        """Vérifie si le fichier a déjà été traité"""
        filename = os.path.basename(self.filepath)

        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute(
                "SELECT nom_fichier, taille_fichier, nb_lignes_traitees, date_traitement, statut FROM fichiers_traites WHERE nom_fichier = ?",
                (filename,))
            result = cursor.fetchone()
            conn.close()

            if result:
                message = f"Le fichier '{filename}' a déjà été traité le {result[3]}.\n"
                message += f"Lignes traitées : {result[2]}\n\n"
                message += "Voulez-vous le traiter à nouveau ?"

                if not messagebox.askyesno("Fichier déjà traité", message):
                    self.filepath = ""
                    self.lbl1.config(text="1. Choisissez le fichier d'entrée CSV : (aucun fichier)")
        except sqlite3.Error as e:
            print(f"Erreur lors de la vérification du fichier : {e}")

    def load_dir(self):
        new_dir = filedialog.askdirectory()
        if new_dir:
            self.outdir = new_dir
            self.lbl2.config(text=f"2. Dossier de base de données : {self.outdir}")
            self.save_last_dir(self.outdir)
            self.db_path = os.path.join(self.outdir, "moulinette_data.db")
            self.update_db_status()

    def load_last_dir(self):
        try:
            with open("last_dir.pkl", "rb") as f:
                return pickle.load(f)
        except (FileNotFoundError, EOFError, pickle.UnpicklingError):
            default_dir = os.path.join(os.path.dirname(__file__), "Fichiers_Sortie")
            os.makedirs(default_dir, exist_ok=True)
            return default_dir

    def save_last_dir(self, path):
        with open("last_dir.pkl", "wb") as f:
            pickle.dump(path, f)

    def start_processing(self):
        """Lance le traitement dans un thread séparé"""
        if not self.filepath:
            messagebox.showerror("Erreur", "Veuillez choisir un fichier d'entrée CSV")
            return

        if not os.path.exists(self.db_path):
            if messagebox.askyesno("Base inexistante",
                                   "La base de données n'existe pas. La créer maintenant ?"):
                self.init_database()
                self.update_db_status()
            else:
                return

        if self.is_processing:
            messagebox.showwarning("Traitement", "Un traitement est déjà en cours")
            return

        self.is_processing = True
        self.progress.start()
        self.convert_button.config(state='disabled')

        thread = threading.Thread(target=self.process_file, daemon=True)
        thread.start()

        self.master.after(100, self.check_processing_status)

    def process_file(self):
        """Traite le fichier par lots et stocke en base"""
        try:
            batch_size = int(self.batch_size_var.get())
            filename = os.path.basename(self.filepath)
            file_size = os.path.getsize(self.filepath)

            ref_data = self.load_reference_data()
            if ref_data is None:
                self.processing_queue.put(
                    ('error', "Erreur critique : un fichier de référence est manquant. Traitement annulé."))
                return

            total_lines = sum(1 for _ in open(self.filepath, 'r', encoding='ISO-8859-1', errors='ignore')) - 1

            self.processing_queue.put(('status', f"Traitement de {total_lines:,} lignes par lots de {batch_size:,}"))

            processed_count = 0
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()

            cursor.execute('''
                INSERT OR REPLACE INTO fichiers_traites 
                (nom_fichier, chemin_complet, taille_fichier, nb_lignes_traitees, statut)
                VALUES (?, ?, ?, 0, 'En cours')
            ''', (filename, self.filepath, file_size))
            file_id = cursor.lastrowid
            conn.commit()

            # --- CORRECTION & OPTIMISATION ---
            # 1. On exécute la requête HORS de la boucle pour plus d'efficacité.
            # 2. On utilise c[1] pour accéder au nom de la colonne dans le tuple.
            cursor.execute('PRAGMA table_info(signalements)')
            db_cols = [c[1] for c in cursor.fetchall()]

            for chunk_df in pd.read_csv(self.filepath, delimiter=';', encoding='ISO-8859-1',
                                        dtype=str, chunksize=batch_size, on_bad_lines='warn'):
                processed_chunk = self.process_chunk(chunk_df, ref_data, filename)

                # Sélectionner uniquement les colonnes dont le nom existe dans la table de destination.
                # Cette approche est plus robuste si le CSV contient des colonnes superflues.
                final_df = processed_chunk[[col for col in processed_chunk.columns if col in db_cols]]

                self.insert_data_in_batches(conn, final_df, 'signalements')

                processed_count += len(chunk_df)
                progress_pct = (processed_count / total_lines) * 100
                self.processing_queue.put(
                    ('status', f"Traité : {processed_count:,}/{total_lines:,} ({progress_pct:.1f}%)"))

            cursor.execute('''
                UPDATE fichiers_traites 
                SET nb_lignes_traitees = ?, statut = 'Terminé', date_traitement = CURRENT_TIMESTAMP
                WHERE id = ?
            ''', (processed_count, file_id))

            conn.commit()
            conn.close()

            self.processing_queue.put(('complete', f"Traitement terminé ! {processed_count:,} lignes traitées."))

        except Exception as e:
            # Tenter de marquer le fichier comme ayant échoué
            try:
                conn = sqlite3.connect(self.db_path)
                cursor = conn.cursor()
                cursor.execute("UPDATE fichiers_traites SET statut = 'Erreur' WHERE nom_fichier = ?", (filename,))
                conn.commit()
                conn.close()
            except Exception:
                pass

            self.processing_queue.put(('error', f"Une erreur majeure est survenue: {e}"))
    #todo : valider que l'exel qui sort est ok avec celui eixstant :
    # ordre des colonnes
    # complètude des informations d'entrées prises en compte dans la base (colonne absentes dans le fichier de sortie)
    # regarder si on peut faire plusieurs formats de sortie
    # todo : voir si on peut faire les stats directement depuis la bdd
    # todo : voir si on peut chunker la bdd pour éviter la base trop grosse

    def process_chunk(self, df, ref_data, filename):
        """OPTIMISATION: Traite un chunk de données en utilisant df.apply pour la performance."""
        df = df.copy()
        df.replace({re.compile(r'[\x01-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]'): ''}, regex=True, inplace=True)

        if 'ALIAS_SIGNALANT' in df.columns:
            df['ALIAS_SIGNALANT'] = df['ALIAS_SIGNALANT'].astype(str).replace('nan', '')

        # Application de la logique de traitement à chaque ligne de manière optimisée
        new_cols_df = df.apply(self.process_row, axis=1, result_type='expand', args=(ref_data,))

        # Renommer les colonnes du nouveau DataFrame
        new_cols_df.columns = [
            'date_requalifiee', 'mois', 'expediteur_nettoye', 'typologie_expediteur', 'operateur_arcep',
            'type_protection', 'rebond_nettoye', 'typologie_rebond', 'opr_arcep_rebond', 'categorie_no_cible'
        ]

        # Joindre les colonnes calculées au DataFrame original
        df = pd.concat([df, new_cols_df], axis=1)
        df['fichier_source'] = filename

        # Renommer les colonnes pour correspondre au schéma DB
        column_mapping = {
            'DATE_SIGNALEMENT': 'date_signalement', 'MESSAGE': 'message', 'EMETTEUR': 'emetteur',
            'ALIAS_SIGNALANT': 'alias_signalant', 'NUMERO_REBOND_SIGNAL': 'numero_rebond_signal',
            'OPERATEUR_SIGNALANT': 'operateur_signalant', 'URL_REBOND_SIGNALE': 'url_rebond_signale',
            'DATE_RECEPTION': 'date_reception', 'MOIS_RECEPTION': 'mois_reception',
            'ANALYSE_STOP': 'analyse_stop', 'TYPE_EMETTEUR': 'type_emetteur'
        }
        df.rename(columns=column_mapping, inplace=True)
        return df

    def process_row(self, row, ref_data):
        """OPTIMISATION: Traite une ligne (une Series) et retourne une Series avec les nouvelles colonnes."""
        oadc_list, base_oadc_interdits, majnum, identifiants_CE = ref_data

        # Initialisation des valeurs de retour
        date_requalifiee, mois = "", ""
        expediteur_nettoye, typologie_expediteur, operateur_arcep, type_protection = "", "Non identifié", "", ""
        rebond_nettoye, typologie_rebond, opr_arcep_rebond = "", "Aucun", ""
        categorie_no_cible = ""

        # 1. Traitement de la date
        date_full = row.get('DATE_SIGNALEMENT')
        if pd.notna(date_full):
            date_requalifiee = str(date_full)[:10]
            mois = str(date_full)[:7]

        # 2. Traitement de l'émetteur
        emetteur = row.get('EMETTEUR')
        if pd.notna(emetteur):
            numero = extraire_numero_de_texte(str(emetteur))
            if numero:
                expediteur_nettoye = numero
                typologie_expediteur = typologie_numero(numero)
                operateur_arcep = trouver_operateur(numero, majnum, identifiants_CE)
            elif str(emetteur).lower() in oadc_list:
                expediteur_nettoye = str(emetteur)
                typologie_expediteur = "OADC"
                type_protection = trouver_interdiction(str(emetteur), base_oadc_interdits)
            else:
                expediteur_nettoye = str(emetteur)  # Conserver la valeur originale

        # 3. Traitement du message pour le rebond
        texte_message = row.get('MESSAGE')
        if pd.notna(texte_message):
            numero_rebond = extraire_numero_de_texte(str(texte_message))
            if numero_rebond:
                rebond_nettoye = numero_rebond
                typologie_rebond = typologie_numero(numero_rebond)
                opr_arcep_rebond = trouver_operateur(numero_rebond, majnum, identifiants_CE)

        # 4. Catégorie no cible
        if pd.notna(row.get('URL_REBOND_SIGNALE')):
            categorie_no_cible = 'URL'
        else:
            categorie_no_cible = typologie_rebond

        return pd.Series([date_requalifiee, mois, expediteur_nettoye, typologie_expediteur, operateur_arcep,
                          type_protection, rebond_nettoye, typologie_rebond, opr_arcep_rebond, categorie_no_cible])

    def load_reference_data(self):
        """ROBUSTESSE: Charge les données de référence en gérant les fichiers manquants."""
        try:
            with open('liste_oadc.csv', 'rb') as f:
                encoding = chardet.detect(f.read())['encoding']
            df_oadc = pd.read_csv('liste_oadc.csv', delimiter=';', encoding=encoding, dtype=str)
            oadc_list = [str(v).lower() for v in df_oadc['OADC'].dropna().tolist()]
        except FileNotFoundError:
            messagebox.showwarning("Fichier manquant",
                                   "'liste_oadc.csv' est introuvable. La typologie OADC sera désactivée.")
            oadc_list = []

        try:
            with open('oadc_sensibles.csv', 'rb') as f:
                encoding = chardet.detect(f.read())['encoding']
            base_oadc_interdits = pd.read_csv('oadc_sensibles.csv', delimiter=';', encoding=encoding, dtype=str)
            base_oadc_interdits['OADC INTERDIT'] = base_oadc_interdits['OADC INTERDIT'].str.lower()
        except FileNotFoundError:
            messagebox.showwarning("Fichier manquant",
                                   "'oadc_sensibles.csv' est introuvable. La protection OADC sera désactivée.")
            base_oadc_interdits = pd.DataFrame(columns=['OADC INTERDIT', "TYPE D'INTERDICTION"])

        try:
            majnum = pd.read_excel('MAJNUM.xls')
        except FileNotFoundError:
            messagebox.showerror("Fichier critique manquant",
                                 "'MAJNUM.xls' est introuvable. Le traitement ne peut continuer.")
            return None  # Erreur bloquante

        try:
            with open('identifiants_CE.csv', 'rb') as f:
                encoding = chardet.detect(f.read())['encoding']
            identifiants_CE = pd.read_csv('identifiants_CE.csv', delimiter=';', encoding=encoding, dtype=str)
        except FileNotFoundError:
            messagebox.showerror("Fichier critique manquant",
                                 "'identifiants_CE.csv' est introuvable. Le traitement ne peut continuer.")
            return None  # Erreur bloquante

        return oadc_list, base_oadc_interdits, majnum, identifiants_CE

    # ... (check_processing_status, show_processed_files, export_to_excel, show_stats, clear_database restent majoritairement inchangées) ...
    def check_processing_status(self):
        """Vérifie le statut du traitement"""
        try:
            while True:
                message_type, message = self.processing_queue.get_nowait()

                if message_type == 'status':
                    self.status_label.config(text=message)
                elif message_type == 'complete':
                    self.progress.stop()
                    self.convert_button.config(state='normal')
                    self.is_processing = False
                    self.status_label.config(text=message)
                    self.update_db_status()
                    messagebox.showinfo("Succès", message)
                    return
                elif message_type == 'error':
                    self.progress.stop()
                    self.convert_button.config(state='normal')
                    self.is_processing = False
                    self.status_label.config(text="Erreur lors du traitement")
                    messagebox.showerror("Erreur", message)
                    return

        except queue.Empty:
            pass

        if self.is_processing:
            self.master.after(100, self.check_processing_status)

    def show_processed_files(self):
        """Affiche la liste des fichiers traités"""
        if not os.path.exists(self.db_path):
            messagebox.showwarning("Attention", "Aucune base de données trouvée")
            return

        try:
            conn = sqlite3.connect(self.db_path)
            files_df = pd.read_sql_query(
                "SELECT nom_fichier, taille_fichier, nb_lignes_traitees, date_traitement, statut FROM fichiers_traites ORDER BY date_traitement DESC",
                conn)
            conn.close()

            if files_df.empty:
                messagebox.showinfo("Information", "Aucun fichier traité trouvé")
                return

            files_window = tk.Toplevel(self.master)
            files_window.title("Fichiers traités")
            files_window.geometry("800x400")

            tree = ttk.Treeview(files_window, columns=('Fichier', 'Taille (KB)', 'Lignes', 'Date', 'Statut'),
                                show='headings')

            for col in tree['columns']:
                tree.heading(col, text=col)
                tree.column(col, width=150)

            for _, row in files_df.iterrows():
                tree.insert('', 'end', values=(
                    row['nom_fichier'],
                    f"{row['taille_fichier'] / 1024:.1f}" if pd.notna(row['taille_fichier']) else 'N/A',
                    f"{row['nb_lignes_traitees']:,}" if pd.notna(row['nb_lignes_traitees']) else '0',
                    row['date_traitement'],

                    row['statut']
                ))

            tree.pack(fill='both', expand=True, padx=10, pady=10)

        except sqlite3.Error as e:
            messagebox.showerror("Erreur", f"Erreur de base de données : {str(e)}")

    def export_to_excel(self):
        """Exporte les données de la base vers Excel avec filtres"""
        if not os.path.exists(self.db_path):
            messagebox.showwarning("Attention", "Aucune base de données trouvée")
            return
        ExportDialog(self.master, self.db_path)

    def show_stats(self):
        """Affiche les statistiques de la base"""
        if not os.path.exists(self.db_path):
            messagebox.showwarning("Attention", "Aucune base de données trouvée")
            return

        try:
            conn = sqlite3.connect(self.db_path)
            stats_query = """
            SELECT 
                COUNT(*) as total_records, COUNT(DISTINCT fichier_source) as total_files,
                COUNT(DISTINCT mois) as total_months, MIN(date_requalifiee) as first_date,
                MAX(date_requalifiee) as last_date
            FROM signalements
            """
            stats = pd.read_sql_query(stats_query, conn)
            conn.close()

            stats_text = f"""Statistiques de la base de données:
    Total d'enregistrements: {stats['total_records'].iloc[0]:,}
    Fichiers traités: {stats['total_files'].iloc[0]}
    Mois couverts: {stats['total_months'].iloc[0]}
    Première date: {stats['first_date'].iloc[0] or 'N/A'}
    Dernière date: {stats['last_date'].iloc[0] or 'N/A'}
    Taille de la base: {os.path.getsize(self.db_path) / (1024 * 1024):.1f} MB"""
            messagebox.showinfo("Statistiques", stats_text)
        except sqlite3.Error as e:
            messagebox.showerror("Erreur", f"Impossible de lire les statistiques : {e}")

    def clear_database(self):
        """Vide la base de données"""
        if not os.path.exists(self.db_path):
            messagebox.showwarning("Attention", "Aucune base de données trouvée")
            return

        if messagebox.askyesno("Confirmation",
                               "Êtes-vous sûr de vouloir vider complètement la base de données ? Cette action est irréversible."):
            try:
                conn = sqlite3.connect(self.db_path)
                conn.execute("DELETE FROM signalements")
                conn.execute("DELETE FROM fichiers_traites")
                conn.execute("VACUUM")
                conn.commit()
                conn.close()
                self.update_db_status()
                messagebox.showinfo("Succès", "Base de données vidée !")
            except sqlite3.Error as e:
                messagebox.showerror("Erreur", f"Impossible de vider la base de données : {e}")


# ... (La classe ExportDialog reste majoritairement inchangée) ...
class ExportDialog(tk.Toplevel):
    def __init__(self, parent, db_path):
        super().__init__(parent)
        self.db_path = db_path
        self.title("Exporter vers Excel")
        self.geometry("400x300")
        self.resizable(False, False)

        self.date_from = tk.StringVar()
        self.date_to = tk.StringVar()
        self.limit_var = tk.StringVar(value="1048575")  # Limite Excel

        self.create_export_widgets()

    def create_export_widgets(self):
        main_frame = tk.Frame(self, padx=10, pady=10)
        main_frame.pack(fill='both', expand=True)

        tk.Label(main_frame, text="Filtres d'export :").grid(row=0, column=0, columnspan=2, sticky='w', pady=(0, 10))

        tk.Label(main_frame, text="Date de début (YYYY-MM-DD):").grid(row=1, column=0, sticky='w')
        tk.Entry(main_frame, textvariable=self.date_from).grid(row=1, column=1, padx=5, pady=2)

        tk.Label(main_frame, text="Date de fin (YYYY-MM-DD):").grid(row=2, column=0, sticky='w')
        tk.Entry(main_frame, textvariable=self.date_to).grid(row=2, column=1, padx=5, pady=2)

        tk.Label(main_frame, text="Limite d'enregistrements:").grid(row=3, column=0, sticky='w')
        tk.Entry(main_frame, textvariable=self.limit_var).grid(row=3, column=1, padx=5, pady=2)

        button_frame = tk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=2, pady=(20, 0))

        tk.Button(button_frame, text="Exporter", command=self.export, bg="lightblue").pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Annuler", command=self.destroy).pack(side=tk.LEFT, padx=5)

    def export(self):
        try:
            query = "SELECT * FROM signalements WHERE 1=1"
            params = []

            if self.date_from.get():
                query += " AND date_requalifiee >= ?"
                params.append(self.date_from.get())

            if self.date_to.get():
                query += " AND date_requalifiee <= ?"
                params.append(self.date_to.get())

            query += " ORDER BY date_requalifiee DESC"

            limit = int(self.limit_var.get())
            if limit > 0:
                query += f" LIMIT {limit}"

            conn = sqlite3.connect(self.db_path)
            df = pd.read_sql_query(query, conn, params=tuple(params))
            conn.close()

            if df.empty:
                messagebox.showwarning("Aucune donnée", "Aucune donnée ne correspond aux critères de filtre.",
                                       parent=self)
                return

            filename = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Fichiers Excel", "*.xlsx")],
                initialfile=f"export_{datetime.now().strftime('%Y%m%d')}.xlsx"
            )

            if filename:
                df.to_excel(filename, index=False, engine='openpyxl')
                messagebox.showinfo("Succès", f"Export terminé ! {len(df):,} lignes exportées.", parent=self)
                self.destroy()

        except ValueError:
            messagebox.showerror("Erreur", "La limite doit être un nombre entier.", parent=self)
        except sqlite3.Error as e:
            messagebox.showerror("Erreur d'export", f"Erreur de base de données : {str(e)}", parent=self)
        except Exception as e:
            messagebox.showerror("Erreur d'export", f"Une erreur inattendue est survenue : {str(e)}", parent=self)


# Fonctions utilitaires (robustesse améliorée)
def extraire_numero_de_texte(texte_source):
    if not isinstance(texte_source, str) or not texte_source:
        return None
    source_sans_espace = texte_source.replace(' ', '').replace('.', '').replace('-', '')
    # Regex amélioré pour capturer des cas limites
    match = re.search(
        r'(?:00)?33700\d{10}|0700\d{10}|700\d{10}|(?:\+|00)33[67]\d{8}|0[67]\d{8}|118\d{3}|3\d{3}|[1-9]\d{4}',
        source_sans_espace)
    if not match:
        return extraire_numero_international(source_sans_espace)
    return normaliser_numero(match.group(0))


def extraire_numero_international(source_sans_espace):
    match_international = re.search(r'(\+\d{8,15})|(00\d{8,15})', source_sans_espace)
    return match_international.group() if match_international else None


def normaliser_numero(numero):
    if numero.startswith('00'):
        numero = numero[2:]
    if numero.startswith('+'):
        numero = numero[1:]
    if len(numero) == 11 and numero.startswith("33"):
        return "0" + numero[2:]
    elif len(numero) == 9 and not numero.startswith("0"):
        return "0" + numero
    return numero


def typologie_numero(numero):
    if not isinstance(numero, str): return "Non identifié"
    # ... (logique inchangée mais plus robuste grâce aux vérifications en amont)
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
        else:
            return "Autres ABDCE"
    elif len(numero) == 4 and numero.startswith('3'):
        return 'SVA'
    else:
        return "Non identifié"


def trouver_operateur(numero, majnum, identifiants_CE):
    """ROBUSTESSE: Correction du paramètre et ajout de gardes contre les erreurs."""
    if majnum.empty or identifiants_CE.empty:
        return ''
    try:
        numero_int = int(numero)
    except (ValueError, TypeError):
        return ''

    tranche = majnum[(majnum['Tranche_Debut'] <= numero_int) & (majnum['Tranche_Fin'] >= numero_int)]
    if not tranche.empty:
        mnemo = tranche['Mnémo'].iloc[0]
        operateur_row = identifiants_CE[identifiants_CE['CODE_OPERATEUR'] == mnemo]
        if not operateur_row.empty:
            return operateur_row['IDENTITE_OPERATEUR'].iloc[0]
    return 'Inconnu'


def trouver_interdiction(oadc, base_oadc_interdits):
    """ROBUSTESSE: Ajout de gardes contre les erreurs."""
    if base_oadc_interdits.empty:
        return 'non protégé'
    try:
        ligne = base_oadc_interdits[base_oadc_interdits['OADC INTERDIT'] == oadc.lower()]
        if not ligne.empty:
            return ligne["TYPE D'INTERDICTION"].iloc[0].lower().strip()
    except (AttributeError, IndexError):
        return 'non protégé'
    return 'non protégé'


if __name__ == "__main__":
    root = tk.Tk()
    app = Application(master=root)
    app.mainloop()