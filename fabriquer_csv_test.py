import pandas as pd
import os

# Chemin du fichier source
fichier_source = 'C:/OD/OneDrive - AFMM/data 33700/rapport_total_202512.csv'

# Lire le CSV avec vos paramètres
df = pd.read_csv(
    fichier_source,
    delimiter=';',
    encoding='ISO-8859-1',
    dtype=str
)

# Garder les 100 premières lignes
df_raccourci = df.head(100)

# Construire le nom du fichier de sortie
nom, ext = os.path.splitext(fichier_source)
fichier_sortie = f"{nom}_raccourci{ext}"

# Exporter
df_raccourci.to_csv(
    fichier_sortie,
    sep=';',
    encoding='ISO-8859-1',
    index=False
)

print(f"Fichier créé : {fichier_sortie}")