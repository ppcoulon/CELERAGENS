#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generate_rapport.py
--------------------
Programme (ligne de commande) d'agregation des fichiers sources RH (DBFRAA,
DBCH, DBFRTW, DBIRLAA, DBITTW, DBMASER, DBMOTW, DBSPTW, DBUKAA, DBUSTW, et
futurs fichiers pays).

Toute la logique metier (schema de reference, synonymes de colonnes, filtre
salarie permanent, calcul ETP, calcul de Total_Salaires et
Total_Salaires_ETP, mediane hors max, ecriture du classeur) vit dans
rapport_core.py, partagee avec l'application web (app.py). Ce fichier ne
contient que l'orchestration ligne de commande : lecture de
config_sources.xlsx, boucle sur les fichiers du dossier "sources/", et
recalcul LibreOffice optionnel.

Ce script :
  1. Lit la configuration des sources (fichier, devise, taux de conversion vers EUR,
     colonnes d'avantages en nature specifiques, pays force si absent du fichier)
     depuis config_sources.xlsx (onglet "Sources")
  2. Charge chaque fichier source actif present dans le dossier "sources/"
  3. Harmonise les colonnes de chaque fichier sur un schema de reference commun
     par correspondance de noms de colonnes connus (dictionnaire SYNONYMS),
     avec repli positionnel si le fichier a exactement le meme nombre de
     colonnes que le schema de reference
  4. Convertit en EUR (montant / taux) les colonnes monetaires "coeur" (salaire
     contractuel, salaire recu, avantages en especes, valeur des actions) ainsi
     que les colonnes d'avantages en nature specifiques listees dans la config
     pour ce fichier, et la Severance (US) si presente
  5. Concatene l'ensemble des fichiers actifs (les colonnes d'avantages en
     nature propres a un pays restent visibles, vides pour les autres pays)
  6. Calcule la colonne "Total_Salaires" (somme, en formule Excel, de :
     Salary_Gross Received + Benefits_Cash Received + Benefits_Shares +
     les 3 buckets canoniques d'avantages en nature + toutes les colonnes
     d'avantages en nature specifiques actives). Severance est exclue (a la
     demande de l'utilisateur : paiement exceptionnel de fin de contrat).
  6bis. Ajoute pour chaque salarie : "Salarie_Permanent" (Oui/Non, d'apres
     Employee_Type et la liste de valeurs "permanent" definie par fichier
     dans PERMANENT_VALUES), "ETP" (Working Hours_Computed / heures temps
     plein annuel du pays, voir FULL_TIME_HOURS) et "Total_Salaires_ETP"
     (Total_Salaires / ETP, salaire ramene a un ETP plein). Ces colonnes
     restent visibles pour TOUS les salaries dans Feuil1 (aucune ligne
     supprimee) ; seul l'onglet Synthese_Salaires filtre sur les permanents.
  7. Reconstruit l'onglet "Synthese_Salaires" : effectifs (tous salaries et
     permanents) en formules Excel (vivantes) ; mediane globale et par pays
     calculees par ce script (valeurs figees au moment de la generation -
     voir note ci-dessous), sur le salaire ETP des salaries permanents
     uniquement, en excluant le salarie au salaire/ETP le plus eleve de
     chaque groupe (median_excluding_max).
  8. Enregistre le classeur

Note performance : au-dela de quelques dizaines de milliers de lignes, une
mediane par pays calculee via une formule tableau Excel (MEDIAN(SI(...)))
devient extremement lente a recalculer (des millions de cellules de calcul
intermediaire). Pour un volume de ~270 000 lignes et 10 pays, ces medianes
sont donc calculees directement par ce script (Python) au moment de la
generation et ecrites comme valeurs, avec une note dans le fichier. Relancer
ce script recalcule et remet a jour ces valeurs. Les effectifs restent des
formules Excel vivantes (peu couteuses).

Usage :
    python3 generate_rapport.py
        (utilise par defaut ./config_sources.xlsx, ./sources/, et ecrit
         ./DB_Combined_EUR.xlsx)

Pour ajouter un nouveau fichier source :
  1. Copier le fichier .xlsx dans le dossier "sources/"
  2. Ajouter une ligne dans l'onglet "Sources" de config_sources.xlsx :
       Fichier                    -> nom exact du fichier (dans sources/)
       Devise                     -> code devise (EUR, CHF, USD, GBP, MAD, ...)
       Taux_vers_EUR              -> taux tel que Montant_EUR = Montant_devise / Taux
                                      (laisser vide ou 1 si deja EUR)
       Actif                      -> Oui / Non
       Colonnes_Avantages_Nature  -> liste des colonnes d'avantages en nature
                                      specifiques a ce fichier, separees par ";"
                                      (uniquement celles qui NE correspondent PAS
                                      deja a un des 3 buckets canoniques : Health
                                      Insurance / Vehicle / Autre)
       Pays_Force                 -> optionnel : code pays a appliquer a toutes
                                      les lignes si le fichier n'a pas de colonne
                                      Country/Pays exploitable
  3. Si les libelles de colonnes du nouveau fichier ne sont reconnus par aucun
     synonyme, le script les signale en console (colonnes non reconnues) et
     les conserve telles quelles dans le fichier combine (pas de perte de
     donnees, juste pas harmonisees) -- completer SYNONYMS dans rapport_core.py
     si besoin. Pensez aussi a completer PERMANENT_VALUES et FULL_TIME_HOURS
     dans rapport_core.py pour que le filtre permanent et le calcul ETP
     fonctionnent pour ce nouveau fichier.

Pour l'usage web (upload de fichiers + saisie des taux dans un navigateur),
voir app.py (meme logique, via rapport_core.py).
"""

import argparse
import os
import shutil
import subprocess
import sys
import tempfile

import pandas as pd

import rapport_core as rc


def recalculate(path, timeout=120):
    recalc_script = os.path.join(os.path.dirname(os.path.abspath(__file__)), "scripts", "recalc.py")
    if not os.path.exists(recalc_script):
        candidates = ["/sessions/wonderful-awesome-wozniak/mnt/.claude/skills/xlsx/scripts/recalc.py"]
        recalc_script = next((c for c in candidates if os.path.exists(c)), None)

    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = os.path.join(tmpdir, os.path.basename(path))
        shutil.copy(path, tmp_path)
        if recalc_script:
            result = subprocess.run([sys.executable, recalc_script, tmp_path, str(timeout)],
                                     capture_output=True, text=True)
            print("  [recalc]", result.stdout.strip() or result.stderr.strip())
        else:
            print("  [recalc] script introuvable — l'ouverture dans Excel/LibreOffice recalculera.")
        shutil.copy(tmp_path, path)


def main():
    parser = argparse.ArgumentParser(description="Agrege les fichiers sources RH en un rapport EUR unique.")
    parser.add_argument("--config", default="config_sources.xlsx")
    parser.add_argument("--sources-dir", default="sources")
    parser.add_argument("--output", default="DB_Combined_EUR.xlsx")
    parser.add_argument("--skip-recalc", action="store_true",
                         help="Ne pas recalculer via LibreOffice (utile pour gros volumes)")
    args = parser.parse_args()

    print(f"Configuration : {args.config}")
    cfg = rc.load_config(args.config)

    frames = []
    all_kind_cols = []
    for _, row in cfg.iterrows():
        df, extra_cols = rc.load_and_prepare_source(args.sources_dir, row)
        if df is not None:
            frames.append(df)
            for c in extra_cols:
                if c not in all_kind_cols:
                    all_kind_cols.append(c)

    if not frames:
        print("Aucun fichier actif charge — rien a faire.")
        return

    combined = pd.concat(frames, ignore_index=True)
    print(f"Total lignes agregees : {len(combined)}")

    combined = rc.add_calculated_columns(combined, all_kind_cols)

    rc.build_workbook(combined, all_kind_cols, args.output)

    # Le recalcul LibreOffice n'a de sens que pour le moteur openpyxl (formules
    # vivantes) ; le moteur rapide (gros volumes) ecrit deja des valeurs figees.
    if not args.skip_recalc and len(combined) <= rc.LARGE_DATASET_THRESHOLD:
        print("Recalcul du classeur...")
        recalculate(args.output)
    print(f"Termine -> {args.output}")


if __name__ == "__main__":
    main()
