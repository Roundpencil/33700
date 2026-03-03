import argparse
import json
from tkinter import messagebox
import os
import pickle

from GUI import print_gui
from convertisseur import convert

# todo :
#  séparer code traitement / lecture fichier / export
# mettre à jour le format d'export af2m
#  créer db
#  ajouter colonnes score phishing (utiliser code évaluation taille cert)
#  ajouter l'insersion dans la db en fin de traitement
#  ajouter l'ajout des pages données / chiffres automatiquement dans un onlget (voire les graphes si on peut faire cela...)
#  ajouter GUI pour réquisition

CONFIG_FILE = "config.json"

def charger_derniere_config(args):
    # est-ce que 'jai un fichier json?
    if os.path.exists(CONFIG_FILE):
        # si oui je le charge
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            config = json.load(f)
        if hasattr(args, "dossier_sortie"):
            args.dossier_sortie = config.get("dossier_sortie")
        if hasattr(args, "db_path"):
            args.db_path = config.get("db_path")
        if hasattr(args, "ajouter_operateurs"):
            args.ajouter_operateurs = config.get("ajouter_operateurs")
        if hasattr(args, "db_path"):
            args.score_phising = config.get("score_phising")
        if hasattr(args, "analyser_si_isa"):
            args.analyser_si_isa = config.get("analyser_si_isa")
        if hasattr(args, "format_etendu"):
            args.format_etendu = config.get("format_etendu")


    # sinon je regarde si j'ai un pickle
    else:
        try:
            with open("last_dir.pkl", "rb") as f:
                args.dossier_sortie = pickle.load(f)
        except (FileNotFoundError, EOFError):
            default_dir = os.path.join(os.path.dirname(__file__), "Fichiers sortie Excel")
            os.makedirs(default_dir, exist_ok=True)  # Creates the default directory if it doesn't exist
            args.dossier_sortie = default_dir
    return args

def sauver_config(args):
    config_out = {
        "dossier_sortie": args.dossier_sortie,
        "db_path": args.db_path,
        "ajouter_operateurs" : args.ajouter_operateurs,
        "score_phising" : args.score_phising,
        "analyser_si_isa": args.analyser_si_isa,
        "format_etendu": args.format_etendu
    }

    with (open(CONFIG_FILE, "w", encoding="utf-8") as f):
        json.dump(config_out, f, ensure_ascii=False, indent=2)

# def save_last_dir(path):
#     with open("last_dir.pkl", "wb") as f:
#         pickle.dump(path, f)

if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description="moulinette 33700 de l'afém avec base de données"
    )

    parser.add_argument(
        '--console',
        action='store_true',
        help="Utiliser uniquement l'interface en ligne de commande"
    )

    parser.add_argument(
        '--fichier_entree',
        type=str,
        help="Chemin vers le fichier csv à processer"
    )

    parser.add_argument(
        '--dossier_sortie',
        type=str,
        default='./',
        help="Chemin vers le dossier d'exportation, par défaut dossier d'exécution"
    )

    parser.add_argument(
        '--db_path',
        type=str,
        default='./test_db',
        help="Chemin vers la base de données locale (par défaut: './test_db')"
    )

    parser.add_argument(
        '--config_vierge',
        action='store_true',
        help="part d'une configuration vierge, sinon utilise la dernière configuration connue"
    )

    parser.add_argument(
        '--ajouter_operateurs',
        action='store_true',
        help="Pour ajouter les opérateurs des numéros de rebond"
    )

    parser.add_argument(
        '--score_phising',
        action='store_true',
        help="Pour ajouter le score phishing"
    )

    parser.add_argument(
        '--analyser_si_isa',
        action='store_true',
        help="Pour ajouter le type de protection des OADC"
    )
    parser.add_argument(
        '--format_etendu',
        action='store_true',
        help="Pour ajouter le type de protection des OADC"
    )

    # Parser les arguments
    args = parser.parse_args()
    print(args)

    if not args.config_vierge:
        charger_derniere_config(args)
    print(f"fichier de config chargé : {args}")

    if not args.console:
        args = print_gui(args)

    print(f"fichier de config à l'issue de la gui : {args}")

    sauver_config(args)
    retour = convert(args.fichier_entree, args.dossier_sortie)

    if not retour:
        messagebox.showinfo("Succès", "Conversion réussie!")
    else:
        messagebox.showerror("Erreur", retour)