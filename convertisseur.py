import os
import re
# from tkinter import messagebox

import chardet
import pandas as pd

MODE_DEFAUT = 'af2m'


def verifier_mode(mode):
    if mode.lower() in ['operateurs', 'af2m']:
        return mode
    else:
        return MODE_DEFAUT


def convert(filepath:str, outdir:str, mode=MODE_DEFAUT):  # sourcery skip: use-named-expression
    mode = verifier_mode(mode)
    if not filepath:
        return "Echec de l'opération : pas de chemin vers le fichier source fourni"

    # try:
    # Load and process the CSV
    # df = pd.read_csv(filepath, delimiter=';', encoding='ISO-8859-1', dtype={'EMETTEUR': str, 'ALIAS_SIGNALANT': str})
    df = pd.read_csv(filepath, delimiter=';', encoding='ISO-8859-1', dtype=str)
    df.replace({re.compile(r'[\x01-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]'): ''}, regex=True, inplace=True)
    df['ALIAS_SIGNALANT'] = df['ALIAS_SIGNALANT'].astype(str).replace('nan', '')
    # df['ALIAS_SIGNALANT'] = df['ALIAS_SIGNALANT'].astype(str).replace('nan', '').str.rstrip('.0')

    oadc_list = charger_liste_oadc_connus()
    print(oadc_list)

    base_oadc_interdits = charger_liste_oadc_sensibles()
    print(base_oadc_interdits)

    # Chercher dans la colonne EMETTEUR
    df['expediteur_nettoye'] = ""
    df['typologie_expediteur'] = ""
    df['rebond_nettoye'] = ""
    df['typologie_rebond'] = ""
    df['date_requalifiee'] = ""
    df['categorie_no_cible'] = ""
    df['mois'] = ""
    df['type_protection'] = ''
    df['opr arcep rebond'] = ''

    # identification des opérateurs de l'éxpéditeur
    majnum = pd.read_excel('MAJNUM.xls')

    identifiants_ce = charger_identifiants_ce()

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
            # message_nettoye = texte_message.replace(' ', '').replace('.', '').replace('-', '')
            # match = re.search(r'33\d{13}|0\d{13}|33\d{9}|0\d{9}|118\d{6}|\d{5}|\d{4}',
            #                   message_nettoye)
            # if match:
            if numero_rebond:
                typologie_rebond = typologie_numero(numero_rebond)
                # numero = normaliser_numero(match.group())
                df.at[i, 'rebond_nettoye'] = numero_rebond
                df.at[i, 'typologie_rebond'] = typologie_rebond
                # operateur = trouver_operateur(numero_rebond, majnum, identifiants_ce)
                # df.at[i, 'opr arcep rebond'] = operateur
                print(f" - rebond nettoyé = {numero_rebond} - typologie : {typologie_rebond}")

        # print(row['URL_REBOND_SIGNALE'])
        if pd.isna(row['URL_REBOND_SIGNALE']):
            # df.at[i, 'categorie_no_cible'] = row['typologie_rebond']
            df.at[i, 'categorie_no_cible'] = df.at[i, 'typologie_rebond']
        else:
            df.at[i, 'categorie_no_cible'] = 'URL'

        # Save to Excel
    outfile = os.path.join(outdir, os.path.basename(filepath).split('.')[0] + '.xlsx')
    df.to_excel(outfile, index=False, engine='openpyxl')

    # Inform the user
    return "Succès", "Conversion réussie!"


def charger_identifiants_ce():
    # Détecter l'encodage du fichier
    with open('identifiants_CE.csv', 'rb') as f:
        result = chardet.detect(f.read())
    # Lire le fichier avec l'encodage détecté
    encoding = result['encoding']
    identifiants_ce = pd.read_csv('identifiants_CE.csv', delimiter=';', encoding=encoding, dtype=str)
    return identifiants_ce


def charger_liste_oadc_sensibles():
    # Charger la liste OADC sensibles
    try:
        # Détecter l'encodage du fichier
        with open('oadc_sensibles.csv', 'rb') as f:
            result = chardet.detect(f.read())

        # Lire le fichier avec l'encodage détecté
        encoding = result['encoding']
        base_oadc_interdits = pd.read_csv('oadc_sensibles.csv', delimiter=';', encoding=encoding, dtype=str)
        print(base_oadc_interdits.columns)
        # Convert the 'OADC INTERDITS' column to lowercase
        base_oadc_interdits['OADC INTERDIT'] = base_oadc_interdits['OADC INTERDIT'].str.lower()
        return base_oadc_interdits

    except FileNotFoundError:
        return pd.DataFrame(columns=["OADC INTERDIT"])


def charger_liste_oadc_connus():
    # Charger la liste OADC
    try:
        # Détecter l'encodage du fichier
        with open('liste_oadc.csv', 'rb') as f:
            result = chardet.detect(f.read())

        # Lire le fichier avec l'encodage détecté
        encoding = result['encoding']
        df_oadc = pd.read_csv('liste_oadc.csv', delimiter=';', encoding=encoding, dtype=str)
        # df_oadc = pd.read_csv('liste_oadc.csv', delimiter=';', dtype=str)

        # oadc_list = df_oadc['OADC'].tolist()
        oadc_list = [str(value).lower() for value in df_oadc['OADC'].tolist()]
    except FileNotFoundError:
        oadc_list = []
    return oadc_list


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