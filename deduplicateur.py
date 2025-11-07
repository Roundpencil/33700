import chardet
import pandas as pd
# import re
import regex as re
import unicodedata
import os
import json
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext
from collections import Counter, defaultdict
import traceback

# --- LISTE DES MOTS-VIDES PAR DÉFAUT ---
# Une liste de base modifiable via l'interface.
DEFAULT_FRENCH_STOP_WORDS = {
    "a", "afin", "ai", "ainsi", "alors", "au", "aucun", "aussi", "autre", "aux",
    "avec", "avoir", "bon", "car", "ce", "ceci", "cela", "ces", "cette", "ceux",
    "chaque", "ci", "comme", "comment", "dans", "de", "des", "deux", "donc", "dont",
    "du", "elle", "en", "et", "etre", "eux", "faire", "il", "ils", "je", "jusque",
    "la", "le", "les", "leur", "leurs", "ma", "mais", "me", "mes", "mon", "ne",
    "nos", "notre", "nous", "ou", "où", "par", "pas", "plus", "pour", "qu", "que",
    "qui", "sa", "sans", "se", "ses", "si", "son", "sont", "sur", "ta", "te", "tes",
    "toi", "ton", "tous", "tout", "tu", "un", "une", "vos", "votre", "vous", "y"
}


# --- CŒUR DE L'APPLICATION (CLASSE PRINCIPALE) ---
class WordFrequencyAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Analyseur de Fréquence de Mots")
        self.root.geometry("800x650")

        # Variables Tkinter
        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.sheet_name = tk.StringVar(value="data")
        self.status_text = tk.StringVar(value="Prêt.")

        self.create_widgets()
        self.load_stop_words()

    def create_widgets(self):
        """Crée et positionne tous les éléments de l'interface."""
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- Section Fichiers ---
        files_frame = ttk.LabelFrame(main_frame, text="1. Fichiers", padding="10")
        files_frame.pack(fill=tk.X, expand=True, pady=(0, 10))

        ttk.Label(files_frame, text="Fichier d'entrée (csv):").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(files_frame, textvariable=self.input_path, width=70).grid(row=0, column=1, sticky="ew", padx=5)
        ttk.Button(files_frame, text="Parcourir...", command=self.select_input_file).grid(row=0, column=2, padx=5)

        ttk.Label(files_frame, text="Fichier de sortie (CSV/XLSX):").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(files_frame, textvariable=self.output_path, width=70).grid(row=1, column=1, sticky="ew", padx=5)
        ttk.Button(files_frame, text="Enregistrer sous...", command=self.select_output_file).grid(row=1, column=2,
                                                                                                  padx=5)
        files_frame.columnconfigure(1, weight=1)

        # --- Section Mots à exclure ---
        stopwords_frame = ttk.LabelFrame(main_frame, text="2. Mots à exclure (un par ligne)", padding="10")
        stopwords_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        self.stopwords_text = scrolledtext.ScrolledText(stopwords_frame, wrap=tk.WORD, height=10)
        self.stopwords_text.pack(fill=tk.BOTH, expand=True)

        stopwords_buttons = ttk.Frame(main_frame)
        stopwords_buttons.pack(fill=tk.X, pady=(0, 10))
        ttk.Button(stopwords_buttons, text="Sauvegarder cette liste", command=self.save_stop_words).pack(side=tk.LEFT)
        ttk.Button(stopwords_buttons, text="Réinitialiser la liste", command=self.reset_stop_words).pack(side=tk.LEFT,
                                                                                                         padx=10)

        # --- Section Lancement ---
        action_frame = ttk.LabelFrame(main_frame, text="3. Lancement de l'analyse", padding="10")
        action_frame.pack(fill=tk.X, expand=True, pady=10)

        self.run_button = ttk.Button(action_frame, text="Lancer l'Analyse", command=self.start_analysis_thread,
                                     style="Accent.TButton")
        self.run_button.pack(pady=10)

        self.progress_bar = ttk.Progressbar(action_frame, orient="horizontal", mode="determinate")
        self.progress_bar.pack(fill=tk.X, expand=True, pady=5)

        ttk.Label(action_frame, textvariable=self.status_text).pack(side=tk.LEFT)

        style = ttk.Style()
        style.configure("Accent.TButton", font=('Helvetica', 12, 'bold'), foreground='green')

    # --- Fonctions de gestion des fichiers ---
    def select_input_file(self):
        path = filedialog.askopenfilename(title="Sélectionner le fichier Excel d'entrée",
                                          filetypes=[("Fichiers Csv", "*.csv")])
        if path:
            self.input_path.set(path)
            # Suggérer un nom pour le fichier de sortie
            base, _ = os.path.splitext(path)
            self.output_path.set(f"{base}_frequence.csv")

    def select_output_file(self):
        path = filedialog.asksaveasfilename(title="Choisir le nom du fichier de sortie", defaultextension=".csv",
                                            filetypes=[("Fichiers CSV", "*.csv")])
        if path:
            self.output_path.set(path)

    # --- Fonctions de gestion des mots-vides ---
    def load_stop_words(self):
        """Charge la liste des mots-vides depuis un fichier JSON, sinon utilise la liste par défaut."""
        try:
            if os.path.exists("mots_vides.json"):
                with open("mots_vides.json", "r", encoding="utf-8") as f:
                    words = set(json.load(f))
            else:
                words = DEFAULT_FRENCH_STOP_WORDS
            self.stopwords_text.delete(1.0, tk.END)
            self.stopwords_text.insert(tk.END, "\n".join(sorted(list(words))))
        except Exception as e:
            messagebox.showerror("Erreur de chargement", f"Impossible de charger la liste des mots-vides : {e}")

    def save_stop_words(self):
        """Sauvegarde la liste actuelle des mots-vides dans un fichier JSON."""
        words = self.stopwords_text.get(1.0, tk.END).strip().split("\n")
        # Filtre les lignes vides
        word_set = {word.strip() for word in words if word.strip()}
        try:
            with open("mots_vides.json", "w", encoding="utf-8") as f:
                json.dump(list(word_set), f, indent=2, ensure_ascii=False)
            messagebox.showinfo("Succès", "La liste des mots à exclure a été sauvegardée dans 'mots_vides.json'.")
        except Exception as e:
            messagebox.showerror("Erreur de sauvegarde", f"Impossible de sauvegarder la liste : {e}")

    def reset_stop_words(self):
        """Réinitialise la liste des mots-vides à sa valeur par défaut."""
        self.stopwords_text.delete(1.0, tk.END)
        self.stopwords_text.insert(tk.END, "\n".join(sorted(list(DEFAULT_FRENCH_STOP_WORDS))))

    # --- Fonctions d'analyse (avec threading) ---
    def start_analysis_thread(self):
        """Prépare et lance l'analyse dans un thread séparé pour ne pas bloquer l'IHM."""
        input_file = self.input_path.get()
        output_file = self.output_path.get()
        if not input_file or not output_file:
            messagebox.showerror("Champs manquants", "Veuillez spécifier les fichiers d'entrée et de sortie.")
            return

        self.run_button.config(state=tk.DISABLED)
        self.progress_bar["value"] = 0

        raw_stop_words = self.stopwords_text.get(1.0, tk.END).strip().split("\n")
        stop_words = {word.strip() for word in raw_stop_words if word.strip()}

        # Lancement du thread
        analysis_thread = threading.Thread(target=self.run_analysis, args=(input_file, output_file, stop_words),
                                           daemon=True)
        analysis_thread.start()

    def update_status(self, value, text):
        """Met à jour la barre de progression et le texte de statut depuis n'importe quel thread."""
        self.root.after(0, self.progress_bar.config, {"value": value})
        self.root.after(0, self.status_text.set, text)

    def run_analysis(self, csv_path, output_path, stop_words_set):
        """Logique principale d'analyse de fréquence (exécutée dans le thread)."""
        try:
            self.update_status(5, "Préparation...")

            # Normaliser les mots-vides pour une comparaison efficace
            processed_stop_words = set()
            for word in stop_words_set:
                nfkd_form = unicodedata.normalize('NFKD', word.lower())
                normalized_stop_word = "".join([c for c in nfkd_form if not unicodedata.combining(c)])
                processed_stop_words.add(normalized_stop_word)

            normalized_word_counts = Counter()
            word_variations = defaultdict(set)

            self.update_status(10, f"Lecture du fichier '{os.path.basename(csv_path)}'...")

            c_size = 5000
            with open(csv_path, 'rb') as f:
                result = chardet.detect(f.read(10000))
                # total_chunks = sum(1 for _ in open(csv_path, 'r')) - 1
                total_rows = sum(1 for _ in f) - 1
                total_chunks = (total_rows // c_size) + 1

            encoding_detected = result['encoding']

            print(f"Encodage détecté : {encoding_detected}")
            # reader = pd.read_csv(csv_path, chunksize=c_size, encoding=encoding_detected)
            # for i, df_chunk in enumerate(reader):
            #     progress = 10 + int((i / total_chunks) * 80)
            #     self.update_status(progress, f"Traitement du bloc {i + 1}/{total_chunks}...")
            #
            #     for text_cell in df_chunk.iloc[:, 0].dropna().astype(str):
            #         text = text_cell.lower().replace('\n', ' ')
            #         text = re.sub(r'[^\p{L}\s]', ' ', text, flags=re.UNICODE)
            #
            #         words = text.split()
            #         for original_word in words:
            #             nfkd_form = unicodedata.normalize('NFKD', original_word)
            #             normalized_word = "".join([c for c in nfkd_form if not unicodedata.combining(c)])
            #
            #             if normalized_word and normalized_word not in processed_stop_words:
            #                 normalized_word_counts[normalized_word] += 1
            #                 word_variations[normalized_word].add(original_word)

            #remplacement de la lecture par un fichier txt
            with open(csv_path, 'r', encoding=encoding_detected, errors='ignore') as f:
                # Ignorer la première ligne (en-tête)
                next(f, None)

                line_count = 0
                chunk_num = 0

                for line in f:
                    line_count += 1

                    # Mise à jour de la progression tous les c_size lignes
                    if line_count % c_size == 0:
                        chunk_num += 1
                        progress = 10 + int((chunk_num / total_chunks) * 80)
                        self.update_status(progress, f"Traitement de la ligne {line_count}...")

                    # Traiter la ligne comme du texte brut (ignorer la structure CSV)
                    # On retire juste les guillemets éventuels en début/fin
                    text = line.strip().strip('"').lower().replace('\n', ' ')
                    text = re.sub(r'[^\p{L}\s]', ' ', text, flags=re.UNICODE)

                    words = text.split()
                    for original_word in words:
                        nfkd_form = unicodedata.normalize('NFKD', original_word)
                        normalized_word = "".join([c for c in nfkd_form if not unicodedata.combining(c)])

                        if normalized_word and normalized_word not in processed_stop_words:
                            normalized_word_counts[normalized_word] += 1
                            word_variations[normalized_word].add(original_word)

            # with pd.ExcelFile(csv_path) as xls:
            #     reader = pd.read_csv(xls, chunksize=5000)
            #
            #     # Pour la barre de progression, on estime le nombre de blocs
            #     total_rows = xls.book[self.sheet_name.get()].max_row
            #     total_chunks = (total_rows // 5000) + 1
            #
            #     for i, df_chunk in enumerate(reader):
            #         progress = 10 + int((i / total_chunks) * 80)
            #         self.update_status(progress, f"Traitement du bloc {i + 1}/{total_chunks}...")
            #
            #         for text_cell in df_chunk.iloc[:, 0].dropna().astype(str):
            #             text = text_cell.lower().replace('\n', ' ')
            #             text = re.sub(r'[^\p{L}\s]', ' ', text, flags=re.UNICODE)
            #
            #             words = text.split()
            #             for original_word in words:
            #                 nfkd_form = unicodedata.normalize('NFKD', original_word)
            #                 normalized_word = "".join([c for c in nfkd_form if not unicodedata.combining(c)])
            #
            #                 if normalized_word and normalized_word not in processed_stop_words:
            #                     normalized_word_counts[normalized_word] += 1
            #                     word_variations[normalized_word].add(original_word)

            self.update_status(90, "Agrégation et sauvegarde...")

            output_rows = []
            for normalized_word, count in normalized_word_counts.most_common(1000000):
                variations_str = ', '.join(sorted(list(word_variations[normalized_word])))
                output_rows.append([normalized_word, variations_str, count])

            output_df = pd.DataFrame(output_rows, columns=['Mot Normalisé', 'Variations Trouvées', 'Fréquence'])

            _, ext = os.path.splitext(output_path)
            if ext.lower() == '.csv':
                output_df.to_csv(output_path, index=False, encoding='utf-8-sig')
            else:
                output_df.to_excel(output_path, index=False, engine='openpyxl')

            self.update_status(100, f"Analyse terminée ! Fichier '{os.path.basename(output_path)}' créé.")
            messagebox.showinfo("Succès",
                                f"Analyse terminée avec succès.\nLe fichier a été enregistré ici :\n{output_path}")

        except Exception as e:
            self.update_status(0, f"Erreur : {e}")
            print(f"Une erreur est survenue : {e}")
            print("Traceback complet :")
            tb_str = traceback.format_exc()
            print(tb_str)
            messagebox.showerror("Erreur d'analyse", f"Une erreur est survenue : {e}")
        finally:
            self.root.after(0, self.run_button.config, {"state": tk.NORMAL})


# --- POINT D'ENTRÉE DU PROGRAMME ---
if __name__ == "__main__":
    root = tk.Tk()
    app = WordFrequencyAnalyzerApp(root)
    root.mainloop()