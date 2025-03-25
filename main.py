import argparse

from GUI import print_GUI
from convertisseur import convert

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Script avec interface graphique ou CLI pour choisir des paramètres.")
    parser.add_argument('--console', action='store_true', help="Utiliser uniquement l'interface en ligne de commande")

    parser.add_argument('--csv_input', type=str, default='', help="Chemin vers le fichier CSV d'entrée")
    parser.add_argument('--output_dir', type=str, default='', help="Dossier de sortie")
    parser.add_argument('--mode', type=str, choices=['operateurs', 'af2m'], default='', help="Mode à utiliser")
    parser.add_argument('--db_path', type=str, default='',
                        help="Chemin vers le fichier base de données (obligatoire si mode af2m)")

    args = parser.parse_args()

    if not args.console:
        args = print_GUI(args)
    else:
        missing_fields = []
        if not args.csv_input:
            missing_fields.append("csv_input")
        if not args.output_dir:
            missing_fields.append("output_dir")
        if not args.mode:
            missing_fields.append("mode")
        if args.mode == "af2m" and not args.db_path:
            missing_fields.append("db_path (requis pour le mode af2m)")

        if missing_fields:
            print(f"Erreur : les champs suivants sont obligatoires : {', '.join(missing_fields)}")
            exit(1)

    # ✅ Utilisation des valeurs collectées
    print("Fichier CSV :", args.csv_input)
    print("Dossier de sortie :", args.output_dir)
    print("Mode sélectionné :", args.mode)

    convert(args.csv_input, args.output_dir, args.mode)
