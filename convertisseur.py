import os

import chardet
import pandas as pd
import numpy as np
import re

import unicodedata
from pandas import DataFrame
import requests
from pathlib import Path

# todo QOL:
#  vérifier au lancement si on est dans les nouvelles lignes:
#   si oui afficher message : voulez-vous refaire traitement alors qu'un export suffit
#  vérifier si on a vraiment besoin de la colonne id ou si l'heure du signalement suffit

# todo, moins urgent :
#  ajouter une fonction + GUI pour faire voiture balais sur la base à postériori de la génération (traitements non effectués) + proposer de faire tourner pendant X heures pour éviter boucle infinie
#  ajouter l'ajout des pages avec les données calculées / chiffres automatiquement dans un onlget de l'excel (voire les graphes si on peut faire cela...)
#  permettre de relancer un calcul smishing / opérateurs qui écrase l'ancien
#  ajouter options pour générer rapport sur d'autre dates que les 3 derniers mois

# todo :
#  ajouter GUI pour réquisition

def observateur_basique(iteration, iterations):
    print(f"Itération {iteration} sur {iterations}")


def download_csv(url: str):
    filename = url.split("/")[-1]  # nom du fichier depuis l'URL
    path = Path.cwd() / filename

    print(f"path = {path}")

    response = requests.get(url)
    response.raise_for_status()  # erreur si téléchargement échoue

    with open(path, "wb") as f:
        f.write(response.content)

    print(f"Fichier téléchargé : {path}")


def enrichir(df: DataFrame,
             avec_arcep_rebond=True, calculer_phishing=False, analyser_si_isa=False, format_etendu=False,
             liste_oadc_csv='liste_oadc.csv', oadc_sensibles_csv='oadc_sensibles.csv',
             identifiants_ce_csv='identifiants_ce.csv', mots_clefs_phising_csv="mots_clefs_phising.csv",
             tous_mots_phishing=False, majnum_csv='MAJNUM.csv', observateur=observateur_basique):
    nombre_lignes = len(df)
    # Charger la liste OADC
    try:
        # Détecter l'encodage du fichier
        with open(liste_oadc_csv, 'rb') as f:
            result = chardet.detect(f.read())

        # Lire le fichier avec l'encodage détecté
        encoding = result['encoding']
        df_oadc = pd.read_csv(liste_oadc_csv, delimiter=';', encoding=encoding, dtype=str)
        # df_oadc = pd.read_csv('liste_oadc.csv', delimiter=';', dtype=str)

        # oadc_list = df_oadc['OADC'].tolist()
        oadc_list = [str(value).lower() for value in df_oadc['OADC'].tolist()]
    except FileNotFoundError:
        oadc_list = []

    print(oadc_list)

    # Charger la liste OADC sensibles
    try:
        # Détecter l'encodage du fichier
        with open(oadc_sensibles_csv, 'rb') as f:
            result = chardet.detect(f.read())

        # Lire le fichier avec l'encodage détecté
        encoding = result['encoding']
        base_oadc_interdits = pd.read_csv(oadc_sensibles_csv, delimiter=';', encoding=encoding, dtype=str)
        print(base_oadc_interdits.columns)
        # Convert the 'OADC INTERDITS' column to lowercase
        base_oadc_interdits['OADC INTERDIT'] = base_oadc_interdits['OADC INTERDIT'].str.lower()
    except FileNotFoundError:
        base_oadc_interdits = None

    try:
        with open(mots_clefs_phising_csv, encoding="cp850") as f:
            liste_mots_clefs_phising = [
                line.strip()
                for line in f
                if line.strip()
            ]

    except FileNotFoundError:
        liste_mots_clefs_phising = []

    print("liste mots clefs phihising : ")
    print(liste_mots_clefs_phising)

    # créer un code traitement
    code_traitement = ''
    code_traitement += 'O' if avec_arcep_rebond else ''
    code_traitement += 'P' if calculer_phishing else ''
    code_traitement += 'I' if analyser_si_isa else ''
    code_traitement += 'T' if tous_mots_phishing else ''
    df['traitements'] = code_traitement

    # identification des opérateurs de l'éxpéditeur
    # majnum = pd.read_excel('MAJNUM.xls')

    # Détecter l'encodage du fichier
    with open(majnum_csv, 'rb') as f:
        result = chardet.detect(f.read())

    # Lire le fichier avec l'encodage détecté
    encoding = result['encoding']
    majnum = pd.read_csv(majnum_csv, delimiter=';', encoding=encoding, dtype=str)
    # s'assurer que les ranches sont bien lues
    majnum['Tranche_Debut'] = pd.to_numeric(majnum['Tranche_Debut'])
    majnum['Tranche_Fin'] = pd.to_numeric(majnum['Tranche_Fin'])

    # Détecter l'encodage du fichier
    with open(identifiants_ce_csv, 'rb') as f:
        result = chardet.detect(f.read())

    # Lire le fichier avec l'encodage détecté
    encoding = result['encoding']
    identifiants_ce = pd.read_csv(identifiants_ce_csv, delimiter=';', encoding=encoding, dtype=str)

    df['operateur_arcep'] = ""

    # df['EMETTEUR'] = df['EMETTEUR'].replace('nan', '')
    for i, row in df.iterrows():
        if i % 1000 == 0:
            observateur(i, nombre_lignes + 1)

        # extraction de la date
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

                # identificaiton opérateur
                operateur = trouver_operateur(numero, majnum, identifiants_ce)
                print(f'Opérateur trouve : {operateur}')
                df.at[i, 'operateur_arcep'] = operateur
            elif emetteur.lower() in oadc_list:
                df.at[i, 'typologie_expediteur'] = "OADC"
                df.at[i, 'expediteur_nettoye'] = emetteur
                print(f" - numero nettoyé = {emetteur} : typologie : OADC")
                if analyser_si_isa:
                    df.at[i, 'type_protection'] = trouver_interdiction(emetteur, base_oadc_interdits)

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

            if numero_rebond:
                typologie_rebond = typologie_numero(numero_rebond)
                # numero = normaliser_numero(match.group())
                df.at[i, 'rebond_nettoye'] = numero_rebond
                df.at[i, 'typologie_rebond'] = typologie_rebond
                if avec_arcep_rebond:
                    operateur = trouver_operateur(numero_rebond, majnum, identifiants_ce)
                    df.at[i, 'opr_arcep_rebond'] = operateur
                print(f" - rebond nettoyé = {numero_rebond} - typologie : {typologie_rebond}")

        # print(row['URL_REBOND_SIGNALE'])
        if pd.isna(row['URL_REBOND_SIGNALE']):
            df.at[i, 'categorie_no_cible'] = df.at[i, 'typologie_rebond']
        else:
            df.at[i, 'categorie_no_cible'] = 'URL'

        if calculer_phishing:
            # set_texte_clean = set(supprimer_accents_et_lower(row['DATE_SIGNALEMENT']))
            # intersection = set_texte_clean.intersection(set_mots_clefs_phising)
            #
            # score_phishing = len(intersection)
            # df.at[i, 'score_smishing'] = score_phishing
            # df.at[i, 'phishing'] = 1 if score_phishing else 0
            # if score_phishing:
            #     df.at[i, 'mots_clefs'] = next(iter(intersection))
            # if tous_mots_phishing:
            #     df.at[i, 'tous_les_mots_clefs'] = ", ".join(intersection)

            texte_clean = supprimer_accents_et_lower(texte_message)
            if not tous_mots_phishing:
                phishing = next((k for k in liste_mots_clefs_phising if k in texte_clean), None)

                if phishing:
                    df.at[i, 'phishing'] = 1
                    df.at[i, 'mots_clefs'] = phishing
                else:
                    df.at[i, 'phishing'] = 0
            else:
                liste_tous_les_mots_phishing = [k for k in liste_mots_clefs_phising if k in texte_clean]
                if liste_tous_les_mots_phishing:
                    df.at[i, 'phishing'] = 1
                    df.at[i, 'mots_clefs'] = liste_tous_les_mots_phishing[0]
                    df.at[i, 'tous_les_mots_clefs'] = ", ".join(liste_tous_les_mots_phishing)
                    df.at[i, 'score_smishing'] = len(liste_tous_les_mots_phishing)
                else:
                    df.at[i, 'phishing'] = 0
                    df.at[i, 'score_smishing'] = 0

    observateur(nombre_lignes, nombre_lignes)
    # Inform the user
    return df


def supprimer_accents_et_lower(texte):
    """
    Supprime les accents d'une chaîne de caractères.
    """
    if pd.isna(texte):
        return ""
    texte = str(texte)
    texte = unicodedata.normalize("NFD", texte)
    texte = "".join(c for c in texte if unicodedata.category(c) != "Mn")
    return texte.lower()


def exporter_df_vers_excel(df: DataFrame, filepath, outdir):
    # Save to Excel

    outfile = os.path.join(outdir, os.path.basename(filepath).split('.')[0] + '.xlsx')
    df.to_excel(outfile, index=False, engine='openpyxl')


# def reordonner_colonnes_df_pour_export(df: DataFrame) -> DataFrame:
#     # changer d'ordre des colonnes
#     column_order = ['DATE_SIGNALEMENT', 'MESSAGE', 'EMETTEUR', 'ALIAS_SIGNALANT', 'NUMERO_REBOND_SIGNAL',
#                     'OPERATEUR_SIGNALANT', 'URL_REBOND_SIGNALE', 'date_requalifiee', 'expediteur_nettoye',
#                     'typologie_expediteur', 'operateur_arcep', 'typologie_rebond', 'categorie_no_cible',
#                     'categorie_no_cible', 'mois', 'phishing', 'mots_clefs', 'score_smishing', 'tous_les_mots_clefs',
#                     'DATE_RECEPTION',
#                     'MOIS_RECEPTION',
#                     'ANALYSE_STOP', 'TYPE_EMETTEUR']
#
#     # Check if all columns in column_order are present in df.columns
#     if set(column_order).issubset(df.columns):
#         remaining_columns = [col for col in df.columns if col not in column_order]
#         new_order = column_order + remaining_columns
#         df = df[new_order]
#     else:
#         missing = set(column_order) - set(df.columns)
#         print(f"Impossible de préparer les colonnes pour l'export, certaines colonnes nécessaires sont manquantes \n"
#               f"{', '.join(missing)}")
#     return df
def get_column_order():
    return [
        'DATE_SIGNALEMENT',
        'MESSAGE',
        'EMETTEUR',
        'ALIAS_SIGNALANT',
        'NUMERO_REBOND_SIGNAL',
        'OPERATEUR_SIGNALANT',
        'URL_REBOND_SIGNALE',
        'date_requalifiee',
        'expediteur_nettoye',
        'typologie_expediteur',
        'operateur_arcep',
        'typologie_rebond',
        'categorie_no_cible',
        'mois',
        'phishing',
        'mots_clefs',
        'score_smishing',
        'tous_les_mots_clefs',
        'DATE_RECEPTION',
        'MOIS_RECEPTION',
        'ANALYSE_STOP',
        'TYPE_EMETTEUR'
    ]


def construire_ordre_colonnes_disponibles(colonnes_disponibles):
    """
    Retourne les colonnes dans l'ordre souhaité, en ajoutant à la fin
    les colonnes restantes non prévues dans l'ordre standard.
    """
    column_order = get_column_order()

    colonnes_disponibles = list(colonnes_disponibles)

    colonnes_ordonnees = [col for col in column_order if col in colonnes_disponibles]
    colonnes_restantes = [col for col in colonnes_disponibles if col not in colonnes_ordonnees]

    return colonnes_ordonnees + colonnes_restantes


def reordonner_colonnes_df_pour_export(df: DataFrame) -> DataFrame:
    new_order = construire_ordre_colonnes_disponibles(df.columns)
    return df[new_order]


def charger_source_dans_dataframe(filepath) -> DataFrame:
    if not filepath:
        raise ValueError("Veuillez choisir un fichier d'entrée CSV")

    # try:
    # Load and process the CSV
    # df = pd.read_csv(self.filepath, delimiter=';', encoding='ISO-8859-1', dtype={'EMETTEUR': str, 'ALIAS_SIGNALANT': str})

    df = pd.read_csv(filepath, delimiter=';', encoding='ISO-8859-1', dtype=str)

    # méthode qui marchait bien avant warning :
    # df.replace({re.compile(r'[\x01-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]'): ''}, regex=True, inplace=True)

    # remplacée par :
    pattern = r'[\x01-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]'

    df = df.apply(
        lambda col: col.str.replace(pattern, '', regex=True)
        if col.dtype == "object" else col
    )

    df['ALIAS_SIGNALANT'] = df['ALIAS_SIGNALANT'].astype(str).replace('nan', '')
    # df['ALIAS_SIGNALANT'] = df['ALIAS_SIGNALANT'].astype(str).replace('nan', '').str.rstrip('.0')

    # ajouter les colonnes pour avoir le bon format de dataframe
    df['expediteur_nettoye'] = ""
    df['typologie_expediteur'] = ""
    df['rebond_nettoye'] = ""
    df['typologie_rebond'] = ""
    df['date_requalifiee'] = ""
    df['categorie_no_cible'] = ""
    df['mois'] = ""
    df['type_protection'] = ''
    df['phishing'] = ''
    df['mots_clefs'] = ''
    df['score_smishing'] = ''
    df['tous_les_mots_clefs'] = ''
    df['opr_arcep_rebond'] = ''
    df['traitements'] = ''
    return df


def extraire_numero_de_texte(texte_source):
    # sourcery skip: use-named-expression
    # source_sans_espace = texte_source.replace(' ', '')  # Retire tous les espaces de la chaîne
    source_sans_espace = texte_source.replace(' ', '').replace('.', '').replace('-',
                                                                                '')  # Retire tous les séparateurs usuels de la chaine
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


def trouver_interdiction(oadc, base_oadc_interdits):
    print(f"\n\n oadc en cours de verification d'interdiction : {oadc} : ")
    try:
        oadc = oadc.lower()
    except ValueError:
        return ''

    try:
        ligne = base_oadc_interdits[base_oadc_interdits['OADC INTERDIT'] == oadc]
        print(ligne)
        return 'non protégé' if ligne.empty else ligne["TYPE D'INTERDICTION"].iloc[0].lower().strip()
    except Exception as e:
        print(f'erreur durant trouver_operateur : {e}')
        return ''
