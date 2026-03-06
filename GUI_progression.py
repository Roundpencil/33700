import threading
import tkinter as tk
import time
from tkinter.ttk import Progressbar
from datetime import datetime, timedelta

class FenetreProgression:

    def __init__(self):

        self.root = tk.Tk()
        self.root.title("Progression")

        self.progress_bar_var = tk.DoubleVar(value=0)
        self.heure_de_fin_var = tk.StringVar()
        self.duree_var = tk.StringVar()

        self.start_time = datetime.now()

        Progressbar(
            self.root,
            orient="horizontal",
            length=300,
            mode='determinate',
            variable=self.progress_bar_var
        ).grid(row=0, column=0, columnspan=3, pady=15, sticky="nsew", padx=30)

        tk.Label(self.root, text="Durée estimée :").grid(
            row=1, column=0, padx=5, pady=10, sticky="e"
        )

        tk.Label(self.root, textvariable=self.duree_var).grid(
            row=1, column=1, pady=10, sticky="w"
        )

        tk.Label(self.root, text="Heure de fin estimée :").grid(
            row=2, column=0, padx=5, pady=10, sticky="e"
        )

        tk.Label(self.root, textvariable=self.heure_de_fin_var).grid(
            row=2, column=1, pady=10, sticky="w"
        )

        # bouton fermer (désactivé au départ)
        self.bouton_fermer = tk.Button(
            self.root,
            text="Fermer",
            state="disabled",
            command=self.root.destroy
        )
        self.bouton_fermer.grid(row=3, column=0, columnspan=3, pady=20)

    def update_estimated_times(self, iteration, iterations):
        elapsed_time = datetime.now() - self.start_time

        if iteration > 0:
            avg_time_per_iteration = elapsed_time / iteration
            estimated_total_duration = avg_time_per_iteration * iterations
        else:
            estimated_total_duration = timedelta(0)

        estimated_total_seconds = int(estimated_total_duration.total_seconds())

        # durée totale estimée
        hours = estimated_total_seconds // 3600
        minutes = (estimated_total_seconds % 3600) // 60
        seconds = estimated_total_seconds % 60

        self.duree_var.set(
            f"{hours:02} heures, {minutes:02} minutes, {seconds:02} secondes"
        )

        # heure de fin estimée
        end_time = self.start_time + estimated_total_duration
        self.heure_de_fin_var.set(end_time.strftime("%H:%M:%S"))

    # def set_progress(self, value):
    #     """Met à jour la progression (0 à 100)."""
    #     self.progress_bar_var.set(value)
    #
    #     if value >= 100:
    #         self.bouton_fermer.config(state="normal")

    def update_progress_bar(self, iteration, iterations):
        self.progress_bar_var.set((iteration/iterations)*100)  # 50 %

    def observateur(self, iteration, iterations):
        self.update_progress_bar(iteration, iterations)
        self.update_estimated_times(iteration, iterations)
        if iteration == iterations:
            self.bouton_fermer.config(state="normal")
        self.root.update_idletasks()

    def set_duree(self, texte):
        self.duree_var.set(texte)

    def set_heure_fin(self, texte):
        self.heure_de_fin_var.set(texte)

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    fenetre = FenetreProgression()


    def simulation():
        for i in range(101):
            time.sleep(0.05)
            fenetre.observateur(i, 100)
            # fenetre.set_progress(i)
            # fenetre.update_progress_bar(i, 101)


    threading.Thread(target=simulation).start()

    fenetre.run()