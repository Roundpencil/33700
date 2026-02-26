import pandas as pd
import os
import unicodedata


def supprimer_accents(texte):
    """
    Supprime les accents d'une chaîne de caractères.
    """
    if pd.isna(texte):
        return ""
    texte = str(texte)
    texte = unicodedata.normalize("NFD", texte)
    texte = "".join(c for c in texte if unicodedata.category(c) != "Mn")
    return texte


def filtrer_messages(fichier_entree, liste_strings, fichier_sortie=None):
    """
    Filtre un fichier Excel en conservant uniquement les lignes dont la colonne
    'message' contient au moins une des strings fournies.

    La recherche est :
    - insensible à la casse
    - insensible aux accents
    """

    # Lecture du fichier Excel
    df = pd.read_excel(fichier_entree)

    if "Message" not in df.columns:
        raise ValueError("La colonne 'message' n'existe pas dans le fichier.")

    # Normalisation des mots-clés
    mots_cles_normalises = [
        supprimer_accents(mot).lower() for mot in liste_strings
    ]

    # Normalisation de la colonne message
    messages_normalises = (
        df["Message"]
        .astype(str)
        .apply(supprimer_accents)
        .str.lower()
    )

    # Filtrage : au moins un mot-clé présent
    masque = messages_normalises.apply(
        lambda message: any(mot in message for mot in mots_cles_normalises)
    )

    df_filtre = df[masque]

    # Définition du fichier de sortie
    if fichier_sortie is None:
        base, ext = os.path.splitext(fichier_entree)
        fichier_sortie = f"{base}_filtre{ext}"

    # Export
    df_filtre.to_excel(fichier_sortie, index=False)

    print(f"Fichier filtré créé : {fichier_sortie}")

if __name__ == "__main__":
    fichier = "donnees.xlsx"
    mots_cles = ["Crédit Mutuel de Bretagne",
                "CMB",
                "Crédit Mutuel du Sud Ouest",
                "Crédit Mutuel du Sud-Ouest",
                "CMSO",
                "Fortuneo",
                "Arkéa Banque Entreprises & Institutionnels",
                "ABEI",
                "Arkéa",
                "Meia",
                "Financo",
                "Arkéa Banque Privée",
                "ABP"]



    filtrer_messages(fichier, mots_cles)