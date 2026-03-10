import argparse
import json
import threading
import traceback
from datetime import datetime
from tkinter import messagebox
import os
import pickle

import database33700
from GUI_parametres import print_gui
from GUI_progression import FenetreProgression
from convertisseur import enrichir, charger_source_dans_dataframe, exporter_df_vers_excel, \
    reordonner_colonnes_df_pour_export

CONFIG_FILE = "config.json"

def charger_derniere_config(args_a_enrichir):
    # est-ce que 'jai un fichier json?
    if os.path.exists(CONFIG_FILE):
        print("fichier de configuration trouvé")
        # si oui je le charge
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            config = json.load(f)
        # if hasattr(config, "dossier_sortie"):
        #     args_a_enrichir.dossier_sortie = config.get("dossier_sortie")
        #     print("pouet")
        # if hasattr(config, "db_path"):
        #     args_a_enrichir.db_path = config.get("db_path")
        # if hasattr(config, "ajouter_operateurs"):
        #     args_a_enrichir.ajouter_operateurs = config.get("ajouter_operateurs")
        # if hasattr(config, "db_path"):
        #     args_a_enrichir.score_phising = config.get("score_phising")
        # if hasattr(config, "analyser_si_isa"):
        #     args_a_enrichir.analyser_si_isa = config.get("analyser_si_isa")
        # if hasattr(config, "contenu_xls"):
        #     args_a_enrichir.contenu_xls = config.get("contenu_xls")
        # if hasattr(config, "format_etendu"):
        #     args_a_enrichir.format_etendu = config.get("format_etendu")
        for k, v in config.items():
            if hasattr(args_a_enrichir, k):
                setattr(args_a_enrichir, k, v)

    # sinon je regarde si j'ai un pickle
    else:
        try:
            with open("last_dir.pkl", "rb") as f:
                args_a_enrichir.dossier_sortie = pickle.load(f)
        except (FileNotFoundError, EOFError):
            default_dir = os.path.join(os.path.dirname(__file__), "Fichiers sortie Excel")
            os.makedirs(default_dir, exist_ok=True)  # Creates the default directory if it doesn't exist
            args_a_enrichir.dossier_sortie = default_dir
    return args_a_enrichir

def sauver_config(args_to_sava):
    config_out = {
        "dossier_sortie": args_to_sava.dossier_sortie,
        "db_path": args_to_sava.db_path,
        "ajouter_operateurs" : args_to_sava.ajouter_operateurs,
        "score_phising" : args_to_sava.score_phising,
        "analyser_si_isa": args_to_sava.analyser_si_isa,
        "format_etendu": args_to_sava.format_etendu,
        "contenu_xls": args_to_sava.contenu_xls
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
        '--contenu_xls',
        type=str,
        default='complet',
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
    print(f"arguments au lancement du programme = {args}")

    if not args.config_vierge:
        charger_derniere_config(args)
    print(f"fichier de config chargé : {args}")

    if not args.console:
        args = print_gui(args)

    print(f"fichier de config à l'issue de la gui : {args}")

    sauver_config(args)
    try:
        df = charger_source_dans_dataframe(args.fichier_entree)
        print("Pré-traitement du fichier d'entrée réussi. \n Début de l'enrichissement")

        debut_conversion = datetime.now()
        # code avant GUI
        # df_enrichie = enrichir(df,
        #                        avec_arcep_rebond=args.ajouter_operateurs,
        #                        analyser_si_isa=args.analyser_si_isa,
        #                        format_etendu=args.format_etendu,
        #                        calculer_phishing=args.score_phising,
        #                        observateur=fenetre_progression.observateur)

        # code après GUI
        fenetre_progression = FenetreProgression()


        def travail():
            try:
                fenetre_progression.set_status("Enrichissement des données source", indeter=False)
                df_enrichie = enrichir(
                    df,
                    avec_arcep_rebond=args.ajouter_operateurs,
                    analyser_si_isa=args.analyser_si_isa,
                    format_etendu=args.format_etendu,
                    calculer_phishing=args.score_phising,
                    observateur=fenetre_progression.observateur,
                )
                # fenetre_progression.notifier_fin(resultat)
                fin_conversion = datetime.now()

                print("Fin de l'enrichissement des donnéees.")
                fenetre_progression.set_status("Fin de l'enrichissement des données.", indeter=False)

                if len(args.db_path) > 1:
                    print("Insertion dans la base de données.")
                    fenetre_progression.set_status("Début de l'insertion dans la base de données.")
                    database33700.exporter_vers_db(db_path=args.db_path,
                                                   noms_fichier=[args.fichier_entree],
                                                   dfs=[df_enrichie])
                    print("Données enregistrées dans la base de donnée.")
                else:
                    print("Pas de base de donnée spécifiée en entrée")

                print("Exportation vers Excel en cours.")
                fenetre_progression.set_status("Exportation vers Excel en cours...")

                if args.contenu_xls == 'nouveaux':
                    df_enrichie = reordonner_colonnes_df_pour_export(df_enrichie)
                    exporter_df_vers_excel(df_enrichie, args.fichier_entree, args.dossier_sortie)
                elif args.contenu_xls == 'complet':
                    # todo : nommer différemment les fichier issus de la db
                    # df_enrichie = reordonner_colonnes_df_pour_export(df_enrichie)
                    database33700.exporter_base_vers_excel(db_path=args.db_path,
                                                           filepath=args.fichier_entree,
                                                           outdir=args.dossier_sortie)
                fenetre_progression.done()

                messagebox.showinfo("Succès", f"Conversion réussie! \n"
                                              f"durée de la conversion : {fin_conversion - debut_conversion}")

            except Exception as exc:
                messagebox.showerror(
                    "Erreur inattendue",
                    f"Une erreur inattendue est survenue :\n{exc}"
                )
                traceback.print_exception(exc)

        thread = threading.Thread(target=travail, daemon=True)
        thread.start()

        fenetre_progression.run()

        # déporté dans le thread
        # fin_conversion = datetime.now()
        #
        # print("Fin de l'enrichissement des donnéees.")
        #
        # if len(args.db_path) > 1:
        #     print("Début de l'insertion dans la base de données.")
        #     database33700.exporter_vers_db(db_path=args.db_path,
        #                                    noms_fichier=[args.fichier_entree],
        #                                    dfs=[df_enrichie])
        #     print("Données enregistrées dans la base de donnée.")
        # else:
        #     print("Pas de base de donnée spécifiée en entrée")
        #
        # print("Exportation vers Excel en cours.")
        #
        # if args.contenu_xls == 'nouveaux':
        #     df_enrichie = reordonner_colonnes_df_pour_export(df_enrichie)
        #     exporter_df_vers_excel(df_enrichie, args.fichier_entree, args.dossier_sortie)
        # elif args.contenu_xls == 'complet':
        #     df_enrichie = reordonner_colonnes_df_pour_export(df_enrichie)
        #     database33700.exporter_base_vers_excel(db_path=args.db_path,
        #                                            filepath=args.fichier_entree,
        #                                            outdir = args.dossier_sortie)

        # messagebox.showinfo("Succès", f"Conversion réussie! \n"
        #                               f"durée de la conversion : {fin_conversion - debut_conversion}")

    except ValueError as e:
        messagebox.showerror("Erreur", str(e))
        traceback.print_exception(e)

    except Exception as e:
        messagebox.showerror(
            "Erreur inattendue",
            f"Une erreur inattendue est survenue :\n{e}"
        )
        traceback.print_exception(e)