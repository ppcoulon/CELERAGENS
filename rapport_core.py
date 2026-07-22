#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
rapport_core.py
----------------
Logique metier partagee entre le script en ligne de commande
(generate_rapport.py) et l'application web (app.py) : schema de reference,
synonymes de colonnes, filtre "salarie permanent", calcul ETP, calcul de
Total_Salaires et Total_Salaires_ETP, mediane hors max, et generation du
classeur Excel de sortie (Feuil1 + Synthese_Salaires).

Ce module ne fait aucune hypothese sur l'origine des fichiers (chemin sur
disque ou fichier uploade en memoire) : toutes les fonctions de lecture
acceptent tout objet compatible avec pandas.read_excel (chemin, BytesIO,
fichier uploade Streamlit...).
"""

import os
import re

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# 1. SCHEMA DE REFERENCE + synonymes connus
# ---------------------------------------------------------------------------

REFERENCE_SCHEMA = [
    "Version", "Legal Entity", "BU", "BusUnit", "Country", "Legal_Entity",
    "Employee_ID Number", "Gender", "Date of Birth", "Employee_Type",
    "Employee_Category", "Employee_Top Management", "Contract_Type",
    "Contract_Duration Type", "Contract_Start Date", "Contract_End Date",
    "Bargaining Agreement", "Disability_Declaration", "Disability_Start Date",
    "Disability_End Date", "Working Hours_Contractual",
    "Working Hours_Periodicity", "Working Hours_Computed",
    "Absence_Hours Paid", "Absence_Hours Unpaid", "Salary_Gross Contractual",
    "Salary_Periodicity Contractual", "Salary_Gross Received",
    "Benefits_Cash Received", "Benefits_Shares",
    "Benefits_in Kind (Health Insurance)", "Benefits_in Kind (Vehicle)",
    "Benefits_in Kind (Autre)", "Data Entry/IN OUT",
]

EXTRA_COLUMN = "Comment"       # colonne additionnelle presente dans plusieurs fichiers
SEVERANCE_COLUMN = "Severance"  # specifique DBUSTW, exclue de Total_Salaires

# Synonymes connus par champ canonique (completes au fil de l'ajout de fichiers :
# DBFRAA, DBCH, DBFRTW, DBIRLAA, DBITTW, DBMASER, DBMOTW, DBSPTW, DBUKAA, DBUSTW)
SYNONYMS = {
    "Version": ["Version"],
    "Legal Entity": ["Legal Entity"],
    "BU": ["BU"],
    "BusUnit": ["BusUnit", "Business_Unit", "Business Unit", "BU.1", "Bus Unit", "BusinessUnit"],
    "Country": ["Country", "Pays"],
    "Legal_Entity": ["Legal_Entity", "Legal_entity", "Entité Légale", "LegalEntity"],
    "Employee_ID Number": ["Employee_ID Number", "Employee ID number", "Employee_ID",
                            "Numéro de matricule", "EmployeeIdNumber"],
    "Gender": ["Gender", "Genre"],
    "Date of Birth": ["Date of Birth", "Date of birth", "Date de naissance", "DateOfBirth,6", "DateOfBirth"],
    "Employee_Type": ["Employee_Type", "Type of employee", "Type de salarié", "TypeOfEmployee"],
    "Employee_Category": ["Employee_Category", "Employee category", "Catégorie de salarié", "EmployeeCategory"],
    "Employee_Top Management": ["Employee_Top Management", "Top management", "Top management ",
                                 "TopManagement"],
    "Contract_Type": ["Contract_Type", "Contract type", "Type de contrat", "ContractType"],
    "Contract_Duration Type": ["Contract_Duration Type", "Contract duration type",
                                "Type de Durée du contrat", "ContractDurationType"],
    "Contract_Start Date": ["Contract_Start Date", "Start date contract", "Date d'entrée contrat",
                             "StartDateContract,6", "StartDateContract"],
    "Contract_End Date": ["Contract_End Date", "End date contract", "Date de sortie contrat",
                           "EndDateContract,4", "EndDateContract"],
    "Bargaining Agreement": ["Bargaining Agreement", "Bargaining agreement", "Convention collective",
                              "BargainingAgreement"],
    "Disability_Declaration": ["Disability_Declaration", "Declared disability", "Handicap déclaré",
                                "DeclaredDisability"],
    "Disability_Start Date": ["Disability_Start Date", "Start date declaration disability",
                               "Date de début déclaration handicap new", "StartDateDeclarationDisability"],
    "Disability_End Date": ["Disability_End Date", "End date declaration disability",
                             "Date de fin déclaration handicap new", "EndDateDeclarationDisability"],
    "Working Hours_Contractual": ["Working Hours_Contractual", "Contracted theoretical working hours",
                                   "Nombre d'heures théoriques prévues au contrat", "ContractedWorkingHours"],
    "Working Hours_Periodicity": ["Working Hours_Periodicity",
                                   "Periodicity (contractual theoretical working hours)",
                                   "Périodicité (Nombre d'heures théoriques prévues au contrat)",
                                   "PeriodicityWorkingHours"],
    "Working Hours_Computed": ["Working Hours_Computed", "Actual hours worked", "Actual hours worked ",
                                "Nombres d'heures réellement travaillées durant la période",
                                "ActualHoursWorked"],
    "Absence_Hours Paid": ["Absence_Hours Paid", "Total number of hours of paid absence",
                            "Nombre total d'heures d'absence payées new", "PaidAbsenceHours"],
    "Absence_Hours Unpaid": ["Absence_Hours Unpaid", "Total number of hours of unpaid absence",
                              "Nombre total d'heures d'absence non-payée new", "UnpaidAbsenceHours"],
    "Salary_Gross Contractual": ["Salary_Gross Contractual", "Gross salary contractual",
                                  "Salaire brut contractuel (non-variable) ", "GrossSalaryContractual"],
    "Salary_Periodicity Contractual": ["Salary_Periodicity Contractual",
                                        "Periodicity (gross salary contractual)",
                                        "Periodicité (Salaire brut contractuel (non-variable))",
                                        "PeriodicityGrossSalary"],
    "Salary_Gross Received": ["Salary_Gross Received", "Gross salary received",
                               "Salaire brut reçu (non-variable) new", "GrossSalaryReceived"],
    "Benefits_Cash Received": ["Benefits_Cash Received", "Benefits in cash recieved",
                                "Benefits in cash recieved ", "Benefits in cash received",
                                "Avantages en espèces reçus (variable, court terme) new",
                                "BenefitsInCashRecieved"],
    "Benefits_Shares": ["Benefits_Shares", "Total fair value of annual long-term incentives received",
                         "Total fair value of annual long term incentives received",
                         "Valeur totale de toutes les primes annuelles à long terme reçues new",
                         "incentiveFairValue"],
    "Benefits_in Kind (Health Insurance)": ["Benefits_in Kind (Health Insurance)",
                                             "Benefits in kind (Eg. Health insurance)",
                                             "Benefits in kind (Health insurance)"],
    "Benefits_in Kind (Vehicle)": ["Benefits_in Kind (Vehicle)", "Benefits in kind (Eg. Company car)",
                                    "Benefits in kind (Company car)"],
    "Benefits_in Kind (Autre)": ["Benefits_in Kind (Autre)", "Benefits in kind (…)", "Other benefits kind"],
    "Data Entry/IN OUT": ["Data Entry/IN OUT", "In/Out", "Data Entry_IN/OUT", "Entrer/Sortir",
                           "Entrer/Sortir "],
}

# Colonnes "Comment"-like -> harmonisees vers EXTRA_COLUMN
COMMENT_SYNONYMS = ["Comment", "Comments", "Commentaire"]

# Cas ambigus ou seule la CASSE distingue deux champs canoniques differents
# (ex: 'Legal Entity' = code groupe/BU vs 'Legal entity' = entite juridique
# du pays -- les deux se normalisent au meme texte en minuscule). Pour ces
# libelles precis, on force une correspondance EXACTE (sensible a la casse)
# prioritaire sur la correspondance normalisee habituelle.
EXACT_SYNONYMS = {
    "Legal Entity": "Legal Entity",
    "Legal entity": "Legal_Entity",
    "Legal_Entity": "Legal_Entity",
    "Legal_entity": "Legal_Entity",
}

# Coeur monetaire (present dans REFERENCE_SCHEMA, converti pour tous les fichiers)
CORE_MONEY_FIELDS = [
    "Salary_Gross Contractual", "Salary_Gross Received", "Benefits_Cash Received",
    "Benefits_Shares", "Benefits_in Kind (Health Insurance)",
    "Benefits_in Kind (Vehicle)", "Benefits_in Kind (Autre)",
]

# Champs sommes dans Total_Salaires (exclut Salary_Gross Contractual : montant
# contractuel theorique, pas percu)
CORE_TOTAL_FIELDS = [
    "Salary_Gross Received", "Benefits_Cash Received", "Benefits_Shares",
    "Benefits_in Kind (Health Insurance)", "Benefits_in Kind (Vehicle)",
    "Benefits_in Kind (Autre)",
]

# ---------------------------------------------------------------------------
# 1bis. Filtre "salarie permanent" + calcul ETP (Equivalent Temps Plein)
# ---------------------------------------------------------------------------
# Champ canonique utilise : Employee_Type (colonne "J" dans la plupart des
# fichiers sources ; colonne "H" pour DBFRTW qui n'a pas de colonne
# Country/BusUnit). Les valeurs considerees comme "permanent" varient selon
# le fichier (langue, granularite intern/extern) -- liste EXACTE fournie par
# le client, correspondance stricte (pas de "contains") : ex. DBMOTW a une
# valeur "CDI EXTERNE" a ne pas confondre avec le champ Contract_Type, qui
# utilise aussi "CDI" pour un sens different.
PERMANENT_VALUES = {
    "DBCH": ["Permanent"],
    "DBFRAA": ["Permanent intern", "Permanent extern"],
    "DBFRTW": ["Salarié permanent"],
    "DBIRLAA": ["Permanent"],
    "DBITTW": ["Permanent intern"],
    "DBMASER": ["Permanent"],
    "DBMOTW": ["Permanent intern", "Permanent Externe", "CDI EXTERNE"],
    "DBSPTW": ["Permanent intern", "Permanent extern"],
    "DBUKAA": ["Permanent"],
    "DBUSTW": ["Permanent intern"],
}

# Nombre d'heures annuelles correspondant a un Equivalent Temps Plein (ETP=1)
# dans chaque pays/fichier. ETP = Working Hours_Computed / valeur ci-dessous.
FULL_TIME_HOURS = {
    "DBCH": 2184,
    "DBFRAA": 1820.04,
    "DBFRTW": 1820.04,
    "DBIRLAA": 1950,
    "DBITTW": 2016,
    "DBMASER": 1820.04,
    "DBMOTW": 2292,
    "DBSPTW": 1776,
    "DBUKAA": 1950,
    "DBUSTW": 2080,
}

# Colonnes d'avantages en nature specifiques a chaque fichier (celles qui ne
# correspondent PAS deja a un des 3 buckets canoniques Health Insurance /
# Vehicle / Autre) -- reprises de config_sources.xlsx. Utilise par l'app web
# pour pre-remplir automatiquement ces colonnes sans que l'utilisateur ait a
# les ressaisir a chaque generation (seul le taux de change est modifiable
# dans l'app).
KNOWN_EXTRA_KIND_COLUMNS = {
    "DBFRTW": ["Benefits_In Kind-Total"],
    "DBIRLAA": ["Benefits in kind (Work Life Balance)", "Benefits in kind (Clothing/Uniform)",
                "Benefits in kind (Broadband & Mobile Phone)", "Benefits in kind (Vouchers)"],
    "DBITTW": ["Avantages en nature reçus (...) Indemnités déplacement"],
    "DBMASER": ["Avantages en nature reçus Mutuelle", "Avantages en nature reçus prévoyance",
                "Avantages en nature reçus (per exemple : véhicule de fonction)",
                "Avantages en nature reçus (...) Ticket restaurant",
                "Avantages en nature reçus (...) Indemnités déplacement"],
    "DBMOTW": ["Benefits_In Kind-Total", "Avantages en nature reçus (...) Ticket restaurant",
               "Avantages en nature reçus (...) Indemnités déplacement"],
    "DBUSTW": ["Benefits in kind - Gym Membership", "Benefits in kind - Health Insurance",
               "Benefits in kind - 401k", "Benefits in Kind - Recognition Program"],
}

# Pays force (fichiers sans colonne Country/Pays exploitable)
KNOWN_PAYS_FORCE = {
    "DBFRTW": "FRA",
}

# Devise et taux par defaut (Montant_EUR = Montant_devise / Taux) -- pre-remplis
# dans l'app web mais modifiables par l'utilisateur (c'est le seul champ que
# l'app permet d'editer).
KNOWN_DEFAULT_DEVISE = {
    "DBCH": ("CHF", 0.93698),
    "DBMOTW": ("MAD", 10.54427),
    "DBUKAA": ("GBP", 0.85460),
    "DBUSTW": ("USD", 1.12427),
}

PERMANENT_FLAG_COLUMN = "Salarie_Permanent"       # "Oui" / "Non" / "Inconnu"
ETP_COLUMN = "ETP"                                # Working Hours_Computed / heures temps plein pays
TOTAL_SALAIRES_ETP_COLUMN = "Total_Salaires_ETP"  # Total_Salaires / ETP

_SOURCE_CODE_RE = re.compile(r"^([A-Za-z0-9]+)_source", re.IGNORECASE)


def source_code_from_filename(filename):
    """Extrait le code source (ex: 'DBFRAA') a partir du nom de fichier
    (ex: 'DBFRAA_source V2_260406.xlsx' -> 'DBFRAA'). Sert a retrouver les
    valeurs 'permanent' (PERMANENT_VALUES), les heures temps plein du pays
    (FULL_TIME_HOURS), les colonnes d'avantages en nature specifiques
    (KNOWN_EXTRA_KIND_COLUMNS) et la devise par defaut (KNOWN_DEFAULT_DEVISE)
    pour ce fichier. Retourne None si le nom de fichier ne suit pas le format
    attendu 'XXX_source...' (fichier non reconnu -> l'appelant doit alors
    gerer le cas 'code source inconnu' : filtre permanent/ETP desactives)."""
    m = _SOURCE_CODE_RE.match(str(filename))
    return m.group(1).upper() if m else None


def median_excluding_max(values):
    """Mediane d'une serie en excluant sa valeur maximale (une seule
    occurrence retiree). Utilise pour 'Salaire Total Median (EUR)' dans
    Synthese_Salaires : a la demande du client, le salarie au salaire/ETP le
    plus eleve du groupe (global ou pays) est exclu avant de calculer la
    mediane. Si le groupe a 0 ou 1 valeur (rien a exclure), retourne la
    mediane telle quelle."""
    s = pd.to_numeric(pd.Series(values), errors="coerce").dropna()
    if len(s) <= 1:
        return s.median() if len(s) else None
    return s.drop(s.idxmax()).median()


def _norm(s):
    return re.sub(r"\s+", " ", str(s)).strip().lower()


_REVERSE_LOOKUP = {}
for canonical, variants in SYNONYMS.items():
    for v in variants:
        _REVERSE_LOOKUP[_norm(v)] = canonical
for v in COMMENT_SYNONYMS:
    _REVERSE_LOOKUP[_norm(v)] = EXTRA_COLUMN


def map_columns(df, source_name):
    cols = list(df.columns)
    rename_map = {}
    used_targets = set()
    matched_by_name = 0
    collisions = []
    for c in cols:
        canonical = EXACT_SYNONYMS.get(c) or _REVERSE_LOOKUP.get(_norm(c))
        if canonical:
            if canonical in used_targets:
                # deux colonnes source pointent vers le meme champ canonique
                # (ex: 'In/Out' et 'Entrer/Sortir ' dans le meme fichier) ->
                # on garde la 1ere occurrence, l'autre reste sous son nom
                # d'origine pour eviter un doublon de colonne (bug de concat).
                collisions.append((c, canonical))
                continue
            rename_map[c] = canonical
            used_targets.add(canonical)
            matched_by_name += 1
    if collisions:
        print(f"  [info] '{source_name}': colonnes en double vers le meme champ, conservees "
              f"sous leur nom d'origine : {collisions}")

    n_ref = len(REFERENCE_SCHEMA)
    if matched_by_name < n_ref * 0.5 and len(cols) in (n_ref, n_ref + 1):
        print(f"  [info] '{source_name}': correspondance par nom incomplete "
              f"({matched_by_name}/{n_ref}) -> repli positionnel utilise.")
        rename_map = {cols[i]: REFERENCE_SCHEMA[i] for i in range(n_ref)}
        if len(cols) == n_ref + 1:
            rename_map[cols[n_ref]] = EXTRA_COLUMN
    else:
        found_canonicals = set(rename_map.values())
        missing = [c for c in REFERENCE_SCHEMA if c not in found_canonicals]
        unmatched = [c for c in cols if c not in rename_map]
        if missing:
            print(f"  [info] '{source_name}': champs du schema non trouves (resteront vides) : {missing}")
        if unmatched:
            print(f"  [info] '{source_name}': colonnes non reconnues (conservees telles quelles) : {unmatched}")

    return df.rename(columns=rename_map)


def load_config(config_path):
    """Lit config_sources.xlsx (onglet 'Sources') -- utilise par le CLI
    (generate_rapport.py). L'app web ne s'en sert pas : elle construit sa
    propre configuration a partir des fichiers uploades et des taux saisis."""
    try:
        cfg = pd.read_excel(config_path, sheet_name="Sources")
    except ValueError:
        cfg = pd.read_excel(config_path)
    cfg.columns = [str(c).strip() for c in cfg.columns]
    required = ["Fichier", "Devise", "Taux_vers_EUR", "Actif"]
    for r in required:
        if r not in cfg.columns:
            raise ValueError(f"Colonne '{r}' manquante dans {config_path}. Colonnes trouvees: {list(cfg.columns)}")
    for optional in ["Colonnes_Avantages_Nature", "Pays_Force"]:
        if optional not in cfg.columns:
            cfg[optional] = pd.NA
    return cfg


def load_and_prepare_source_obj(file_obj, filename, devise="EUR", taux=1.0,
                                 pays_force=None, extra_kind_cols=None):
    """Lit et prepare UN fichier source a partir d'un objet fichier (chemin,
    BytesIO, fichier uploade Streamlit...). C'est la fonction centrale
    utilisee aussi bien par le CLI (via load_and_prepare_source) que par
    l'app web (upload direct).

    Retourne (df_harmonise, extra_kind_cols) ou (None, []) en cas d'erreur de
    lecture."""
    devise = str(devise).strip().upper() if devise not in (None, "") else "EUR"
    try:
        taux = float(taux) if taux not in (None, "") else 1.0
    except (TypeError, ValueError):
        taux = 1.0
    extra_kind_cols = list(extra_kind_cols) if extra_kind_cols else []

    print(f"  [lecture] {filename}  (devise={devise}, taux vers EUR={taux})")
    df = pd.read_excel(file_obj, sheet_name=0)
    df = map_columns(df, filename)

    for c in REFERENCE_SCHEMA:
        if c not in df.columns:
            df[c] = pd.NA
    if EXTRA_COLUMN not in df.columns:
        df[EXTRA_COLUMN] = pd.NA
    if SEVERANCE_COLUMN not in df.columns:
        df[SEVERANCE_COLUMN] = pd.NA

    if pays_force:
        df["Country"] = pays_force

    source_code = source_code_from_filename(filename)

    # --- Filtre "permanent" : ajoute une colonne indicative Oui/Non (le
    # filtrage effectif pour la mediane se fait plus tard, dans
    # Synthese_Salaires -- Feuil1 conserve TOUS les salaries, permanents ou
    # non) ---
    if source_code in PERMANENT_VALUES:
        allowed = PERMANENT_VALUES[source_code]
        emp_type = df["Employee_Type"].apply(lambda x: str(x).strip() if pd.notna(x) else x)
        df[PERMANENT_FLAG_COLUMN] = emp_type.isin(allowed).map({True: "Oui", False: "Non"})
    else:
        print(f"  [ATTENTION] '{filename}' : code source '{source_code}' non reconnu pour le "
              f"filtre permanent -> colonne '{PERMANENT_FLAG_COLUMN}' = 'Inconnu' (a exclure "
              f"manuellement des medianes). Completer PERMANENT_VALUES dans rapport_core.py.")
        df[PERMANENT_FLAG_COLUMN] = "Inconnu"

    # --- ETP (Equivalent Temps Plein) = heures travaillees / heures temps
    # plein annuel du pays. Permet de ramener le salaire a un ETP plein. ---
    if source_code in FULL_TIME_HOURS:
        hours = df["Working Hours_Computed"].apply(lambda x: x.replace(",", ".") if isinstance(x, str) else x)
        hours = pd.to_numeric(hours, errors="coerce")
        df[ETP_COLUMN] = hours / FULL_TIME_HOURS[source_code]
    else:
        print(f"  [ATTENTION] '{filename}' : code source '{source_code}' non reconnu pour le "
              f"calcul ETP -> colonne '{ETP_COLUMN}' vide. Completer FULL_TIME_HOURS dans rapport_core.py.")
        df[ETP_COLUMN] = pd.NA

    money_cols_this_file = list(CORE_MONEY_FIELDS) + [c for c in extra_kind_cols if c in df.columns] \
        + ([SEVERANCE_COLUMN] if SEVERANCE_COLUMN in df.columns else [])

    # Nettoyage numerique systematique (independant de la devise) : certains
    # fichiers contiennent des espaces vides (" ") ou des virgules decimales
    # (ex: "131,51") a la place de nombres. On les corrige avant toute
    # conversion pour ne pas perdre ces valeurs ni faire planter l'ecriture.
    for c in money_cols_this_file:
        if c not in df.columns:
            continue
        col = df[c].apply(lambda x: x.replace(",", ".") if isinstance(x, str) else x)
        df[c] = pd.to_numeric(col, errors="coerce")

    if devise != "EUR" and taux and taux != 1.0:
        for c in money_cols_this_file:
            df[c] = df[c] / taux
        source_note = f"{os.path.splitext(str(filename))[0]} (converti {devise}->EUR, taux {taux})"
    else:
        source_note = f"{os.path.splitext(str(filename))[0]} (EUR)"

    df["Source_Fichier"] = source_note

    final_cols = REFERENCE_SCHEMA + [EXTRA_COLUMN, SEVERANCE_COLUMN, PERMANENT_FLAG_COLUMN,
                                      ETP_COLUMN, "Source_Fichier"] + extra_kind_cols
    # dedupe en gardant l'ordre
    seen = set()
    final_cols = [c for c in final_cols if not (c in seen or seen.add(c))]
    for c in final_cols:
        if c not in df.columns:
            df[c] = pd.NA
    return df[final_cols], extra_kind_cols


def load_and_prepare_source(sources_dir, row):
    """Wrapper CLI : lit une ligne de config_sources.xlsx (Fichier, Devise,
    Taux_vers_EUR, Actif, Colonnes_Avantages_Nature, Pays_Force), verifie le
    flag Actif et l'existence du fichier sur disque, puis delegue a
    load_and_prepare_source_obj. Conserve pour compatibilite avec
    generate_rapport.py (usage ligne de commande)."""
    filename = str(row["Fichier"]).strip()
    devise = str(row["Devise"]).strip().upper() if pd.notna(row["Devise"]) else "EUR"
    taux = row["Taux_vers_EUR"]
    taux = float(taux) if pd.notna(taux) and str(taux).strip() != "" else 1.0
    actif = str(row["Actif"]).strip().lower() if pd.notna(row["Actif"]) else "oui"
    pays_force = str(row["Pays_Force"]).strip() if pd.notna(row.get("Pays_Force")) else None
    extra_kind_cols = []
    if pd.notna(row.get("Colonnes_Avantages_Nature")):
        extra_kind_cols = [c.strip() for c in str(row["Colonnes_Avantages_Nature"]).split(";") if c.strip()]

    if actif not in ("oui", "yes", "true", "1", "y", "o"):
        print(f"  [ignore] '{filename}' marque inactif dans la config.")
        return None, []

    path = os.path.join(sources_dir, filename)
    if not os.path.exists(path):
        print(f"  [ERREUR] fichier introuvable : {path} -> ignore.")
        return None, []

    return load_and_prepare_source_obj(path, filename, devise=devise, taux=taux,
                                        pays_force=pays_force, extra_kind_cols=extra_kind_cols)


def add_calculated_columns(combined, all_kind_cols):
    """A appeler UNE fois sur le DataFrame combine (toutes les sources deja
    concatenees). Ajoute :
      - '_Total_Salaires_calc' : Total_Salaires (usage interne, non ecrit
        directement -- sert de base a la formule Excel ou a la valeur figee
        selon le moteur d'ecriture, et au calcul de Total_Salaires_ETP)
      - ETP_COLUMN convertie en numerique
      - TOTAL_SALAIRES_ETP_COLUMN = Total_Salaires / ETP (NaN si ETP manquant
        ou <= 0, pour eviter une division par zero / un ETP negatif absurde)
    Modifie et retourne le DataFrame passe en argument."""
    total_field_names = [c for c in (CORE_TOTAL_FIELDS + all_kind_cols) if c in combined.columns]
    combined["_Total_Salaires_calc"] = combined[total_field_names].apply(
        pd.to_numeric, errors="coerce").sum(axis=1, skipna=True)

    combined[ETP_COLUMN] = pd.to_numeric(combined[ETP_COLUMN], errors="coerce")
    valid_etp = combined[ETP_COLUMN].where(combined[ETP_COLUMN] > 0)
    combined[TOTAL_SALAIRES_ETP_COLUMN] = combined["_Total_Salaires_calc"] / valid_etp
    return combined


# Au-dela de ce nombre de lignes, on bascule sur le moteur d'ecriture rapide
# (xlsxwriter) : au-dela d'environ 60-70k lignes, l'ecriture cellule-par-cellule
# openpyxl (et plus encore une formule Excel par ligne) devient trop lente pour
# une generation interactive. Le moteur rapide ecrit Total_Salaires comme une
# VALEUR calculee par Python (au lieu d'une formule Excel vivante) : la donnee
# est correcte au moment de la generation, mais ne se recalculera pas toute
# seule si on modifie une cellule source a la main dans Excel -- il faut
# relancer generate_rapport.py / l'app.
LARGE_DATASET_THRESHOLD = 60000


def build_workbook_openpyxl(combined_df_with_calc, all_kind_cols, output_path, style_rows=True, max_styled_rows=60000):
    """Moteur d'ecriture openpyxl : formules Excel vivantes pour Total_Salaires
    et pour les effectifs. Utilise uniquement pour les petits volumes (voir
    LARGE_DATASET_THRESHOLD et la fonction build_workbook). output_path peut
    etre un chemin ou un objet fichier binaire (BytesIO) : les deux sont
    acceptes par openpyxl."""
    # combined_df_with_calc contient une colonne interne "_Total_Salaires_calc"
    # (mediane Python) qui ne doit PAS être ecrite dans la feuille de donnees.
    combined_df = combined_df_with_calc.drop(columns=["_Total_Salaires_calc"])
    final_cols = list(combined_df.columns)
    combined_df.to_excel(output_path, sheet_name="Feuil1", index=False)

    # output_path peut etre un objet fichier binaire (BytesIO, cas de l'app
    # web) : il faut revenir au debut avant de le relire.
    if hasattr(output_path, "seek"):
        output_path.seek(0)
    wb = openpyxl.load_workbook(output_path)
    ws = wb["Feuil1"]

    header_font = Font(name="Arial", bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    body_font = Font(name="Arial")

    ncols = ws.max_column
    nrows = ws.max_row

    for col in range(1, ncols + 1):
        cell = ws.cell(row=1, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    if PERMANENT_FLAG_COLUMN in final_cols:
        ws.cell(row=1, column=final_cols.index(PERMANENT_FLAG_COLUMN) + 1).comment = Comment(
            "Oui/Non selon Employee_Type et la liste de valeurs 'permanent' definie par fichier "
            "source (voir PERMANENT_VALUES dans rapport_core.py). 'Inconnu' si le fichier source "
            "n'est pas reconnu (a completer dans le script).", "rapport_core.py")
    if ETP_COLUMN in final_cols:
        ws.cell(row=1, column=final_cols.index(ETP_COLUMN) + 1).comment = Comment(
            "Equivalent Temps Plein = Working Hours_Computed / heures temps plein annuel du pays "
            "(voir FULL_TIME_HOURS dans rapport_core.py). ETP=1 correspond a un temps plein legal.",
            "rapport_core.py")
    if TOTAL_SALAIRES_ETP_COLUMN in final_cols:
        ws.cell(row=1, column=final_cols.index(TOTAL_SALAIRES_ETP_COLUMN) + 1).comment = Comment(
            "Total_Salaires / ETP : salaire ramene a un Equivalent Temps Plein plein (ETP=1), pour "
            "comparer des salaries a temps partiel et a temps plein. Vide si ETP manquant ou <= 0.",
            "rapport_core.py")

    money_fields_all = CORE_MONEY_FIELDS + all_kind_cols + [SEVERANCE_COLUMN, TOTAL_SALAIRES_ETP_COLUMN]
    money_col_idx = [final_cols.index(c) + 1 for c in money_fields_all if c in final_cols]
    date_cols = ["Date of Birth", "Contract_Start Date", "Contract_End Date",
                 "Disability_Start Date", "Disability_End Date"]
    date_col_idx = [final_cols.index(c) + 1 for c in date_cols if c in final_cols]
    etp_col_idx = [final_cols.index(c) + 1 for c in [ETP_COLUMN] if c in final_cols]

    # Nombre de formatage de cellules important au-dela d'un certain volume :
    # on applique toujours le format numerique des colonnes monetaires (rapide),
    # mais on saute la police par cellule sur les tres gros volumes (police par
    # defaut conservee) afin de rester dans des temps de generation raisonnables.
    apply_font = style_rows and nrows <= max_styled_rows
    for row in range(2, nrows + 1):
        for col in money_col_idx:
            ws.cell(row=row, column=col).number_format = '#,##0.00 €'
        for col in date_col_idx:
            ws.cell(row=row, column=col).number_format = 'yyyy-mm-dd'
        for col in etp_col_idx:
            ws.cell(row=row, column=col).number_format = '0.00'
        if apply_font:
            for col in range(1, ncols + 1):
                ws.cell(row=row, column=col).font = body_font

    ws.freeze_panes = "A2"
    for col in range(1, ncols + 1):
        letter = get_column_letter(col)
        header_len = len(str(final_cols[col - 1]))
        ws.column_dimensions[letter].width = min(max(header_len + 2, 12), 32)

    # --- Total_Salaires : Salary_Gross Received + Benefits_Cash Received +
    #     Benefits_Shares + 3 buckets canoniques + toutes colonnes d'avantages
    #     en nature specifiques (Severance exclue) ---
    headers = {ws.cell(row=1, column=c).value: c for c in range(1, ws.max_column + 1)}
    total_field_names = CORE_TOTAL_FIELDS + all_kind_cols
    total_col_idx = sorted(headers[c] for c in total_field_names if c in headers)

    new_col = ws.max_column + 1
    new_letter = get_column_letter(new_col)
    hcell = ws.cell(row=1, column=new_col, value="Total_Salaires")
    hcell.font = header_font
    hcell.fill = header_fill
    hcell.alignment = Alignment(horizontal="center", vertical="center")
    hcell.comment = Comment(
        "Somme (EUR) de : Salary_Gross Received + Benefits_Cash Received + Benefits_Shares "
        "+ toutes les colonnes d'avantages en nature (buckets canoniques + colonnes specifiques "
        "par pays). Exclut Salary_Gross Contractual (montant contractuel, pas percu) et Severance "
        "(indemnite de depart exceptionnelle, exclue sur demande).",
        "rapport_core.py"
    )
    # Formule par ligne : somme des cellules concernees (pas forcement contigues,
    # donc on construit une liste de references cellule par cellule)
    letters = [get_column_letter(c) for c in total_col_idx]
    for r in range(2, nrows + 1):
        formula = "=SUM(" + ",".join(f"{l}{r}" for l in letters) + ")"
        cell = ws.cell(row=r, column=new_col, value=formula)
        cell.number_format = '#,##0.00 €'
        if apply_font:
            cell.font = body_font
    ws.column_dimensions[new_letter].width = 18

    country_col = headers["Country"]
    country_letter = get_column_letter(country_col)
    permanent_col = headers.get(PERMANENT_FLAG_COLUMN)
    permanent_letter = get_column_letter(permanent_col) if permanent_col else None

    countries = sorted({ws.cell(row=r, column=country_col).value for r in range(2, nrows + 1)
                        if ws.cell(row=r, column=country_col).value not in (None, "")})
    print(f"  Pays detectes dans les donnees : {countries}")

    # --- Medianes par pays et globale : calculees en Python (voir note en
    # tete de fichier). Base = salaries "permanent" uniquement
    # (Salarie_Permanent = "Oui"), salaire ramene a un ETP plein
    # (Total_Salaires_ETP), mediane calculee en excluant le salarie au
    # salaire/ETP le plus eleve de chaque groupe (voir median_excluding_max). ---
    if "Synthese_Salaires" in wb.sheetnames:
        del wb["Synthese_Salaires"]
    if "Aide_Calculs" in wb.sheetnames:
        del wb["Aide_Calculs"]
    ws_syn = wb.create_sheet("Synthese_Salaires")
    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    headers_row = ["Pays", "Nombre de salaries (tous)", "Nombre de salaries Permanent (base mediane)",
                   "Salaire Total Median ETP (EUR)"]
    for i, h in enumerate(headers_row, start=1):
        c = ws_syn.cell(row=1, column=i, value=h)
        c.font = header_font
        c.fill = header_fill
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border

    note_cell = ws_syn.cell(row=1, column=6,
        value="Note : effectifs (tous salaries et permanents) = formules Excel vivantes. "
              "Mediane = salaire ramene a un ETP plein (Total_Salaires_ETP), calculee sur les "
              "salaries permanents uniquement (Salarie_Permanent = 'Oui'), en excluant le "
              "salarie au salaire/ETP le plus eleve du groupe (global ou pays). Valeur calculee "
              "au moment de la generation ; relancer le traitement pour la mettre a jour.")
    note_cell.font = Font(name="Arial", italic=True, size=9, color="808080")

    data_range_country = f"Feuil1!${country_letter}$2:${country_letter}${nrows}"
    data_range_permanent = f"Feuil1!${permanent_letter}$2:${permanent_letter}${nrows}" if permanent_letter else None

    permanent_mask_calc = (combined_df_with_calc[PERMANENT_FLAG_COLUMN] == "Oui") \
        if PERMANENT_FLAG_COLUMN in combined_df_with_calc.columns \
        else pd.Series(False, index=combined_df_with_calc.index)
    etp_salary_calc = combined_df_with_calc[TOTAL_SALAIRES_ETP_COLUMN] \
        if TOTAL_SALAIRES_ETP_COLUMN in combined_df_with_calc.columns \
        else combined_df_with_calc["_Total_Salaires_calc"]

    row = 2
    ws_syn.cell(row=row, column=1, value="GLOBAL").font = body_font
    ws_syn.cell(row=row, column=2, value=f"=COUNTA({data_range_country})").font = body_font
    if data_range_permanent:
        ws_syn.cell(row=row, column=3, value=f'=COUNTIF({data_range_permanent},"Oui")').font = body_font
    global_med = median_excluding_max(etp_salary_calc[permanent_mask_calc])
    gcell = ws_syn.cell(row=row, column=4, value=round(float(global_med), 2) if pd.notna(global_med) else None)
    gcell.font = body_font
    gcell.number_format = '#,##0.00 €'
    for col in range(1, 5):
        ws_syn.cell(row=row, column=col).border = border
    row += 1

    for country in countries:
        mask_all = combined_df_with_calc["Country"] == country
        mask_perm = mask_all & permanent_mask_calc
        med = median_excluding_max(etp_salary_calc[mask_perm])
        ws_syn.cell(row=row, column=1, value=country).font = body_font
        ws_syn.cell(row=row, column=2,
                    value=f'=COUNTIF({data_range_country},"{country}")').font = body_font
        if data_range_permanent:
            ws_syn.cell(row=row, column=3,
                        value=f'=COUNTIFS({data_range_country},"{country}",{data_range_permanent},"Oui")').font = body_font
        mcell = ws_syn.cell(row=row, column=4, value=round(float(med), 2) if pd.notna(med) else None)
        mcell.font = body_font
        mcell.number_format = '#,##0.00 €'
        for col in range(1, 5):
            ws_syn.cell(row=row, column=col).border = border
        row += 1

    ws_syn.column_dimensions["A"].width = 14
    ws_syn.column_dimensions["B"].width = 20
    ws_syn.column_dimensions["C"].width = 28
    ws_syn.column_dimensions["D"].width = 26
    ws_syn.column_dimensions["F"].width = 90
    ws_syn.freeze_panes = "A2"

    if hasattr(output_path, "seek"):
        output_path.seek(0)
        output_path.truncate()
    wb.save(output_path)
    return countries


def build_workbook_fast(combined_df_with_calc, all_kind_cols, output_path):
    """Moteur d'ecriture rapide (xlsxwriter) pour les gros volumes (au-dela de
    LARGE_DATASET_THRESHOLD lignes). Total_Salaires et les medianes par pays
    sont ecrits comme des VALEURS calculees par Python plutot que des formules
    Excel, afin de rester dans des temps de generation raisonnables. Accepte
    un chemin ou un objet fichier binaire (BytesIO) pour output_path."""
    import xlsxwriter

    data_cols = [c for c in combined_df_with_calc.columns if c != "_Total_Salaires_calc"]
    combined_df = combined_df_with_calc[data_cols]
    totals = combined_df_with_calc["_Total_Salaires_calc"]

    money_fields_all = set(CORE_MONEY_FIELDS + all_kind_cols + [SEVERANCE_COLUMN, TOTAL_SALAIRES_ETP_COLUMN])
    date_cols = ["Date of Birth", "Contract_Start Date", "Contract_End Date",
                 "Disability_Start Date", "Disability_End Date"]
    money_idx = {i for i, c in enumerate(data_cols) if c in money_fields_all}
    date_idx = {i for i, c in enumerate(data_cols) if c in date_cols}
    etp_idx = {i for i, c in enumerate(data_cols) if c == ETP_COLUMN}
    country_idx = data_cols.index("Country")
    permanent_idx = data_cols.index(PERMANENT_FLAG_COLUMN) if PERMANENT_FLAG_COLUMN in data_cols else None

    wb = xlsxwriter.Workbook(output_path, {"constant_memory": True})
    ws = wb.add_worksheet("Feuil1")

    header_fmt = wb.add_format({"bold": True, "font_color": "white", "bg_color": "#4472C4",
                                 "font_name": "Arial", "align": "center", "valign": "vcenter"})
    money_fmt = wb.add_format({"num_format": '#,##0.00 "€"'})
    date_fmt = wb.add_format({"num_format": "yyyy-mm-dd"})
    etp_fmt = wb.add_format({"num_format": "0.00"})

    for c, name in enumerate(data_cols):
        ws.write(0, c, name, header_fmt)
    if PERMANENT_FLAG_COLUMN in data_cols:
        ws.write_comment(0, data_cols.index(PERMANENT_FLAG_COLUMN),
            "Oui/Non selon Employee_Type et la liste de valeurs 'permanent' definie par fichier "
            "source (voir PERMANENT_VALUES dans rapport_core.py). 'Inconnu' si le fichier source "
            "n'est pas reconnu (a completer dans le script).")
    if ETP_COLUMN in data_cols:
        ws.write_comment(0, data_cols.index(ETP_COLUMN),
            "Equivalent Temps Plein = Working Hours_Computed / heures temps plein annuel du pays "
            "(voir FULL_TIME_HOURS dans rapport_core.py). ETP=1 correspond a un temps plein legal.")
    if TOTAL_SALAIRES_ETP_COLUMN in data_cols:
        ws.write_comment(0, data_cols.index(TOTAL_SALAIRES_ETP_COLUMN),
            "Total_Salaires / ETP : salaire ramene a un Equivalent Temps Plein plein (ETP=1), pour "
            "comparer des salaries a temps partiel et a temps plein. Vide si ETP manquant ou <= 0.")
    total_col = len(data_cols)
    ws.write(0, total_col, "Total_Salaires", header_fmt)
    ws.write_comment(
        0, total_col,
        "Somme (EUR) de : Salary_Gross Received + Benefits_Cash Received + Benefits_Shares "
        "+ toutes les colonnes d'avantages en nature (buckets canoniques + colonnes specifiques "
        "par pays). Exclut Salary_Gross Contractual et Severance. "
        "VALEUR calculee au moment de la generation (pas une formule vivante, voir volume de "
        "donnees) : relancer le traitement apres toute modification des donnees sources."
    )

    country_counts = {}
    permanent_counts = {}
    permanent_country_counts = {}
    nrows = 0
    for r, (row, tot) in enumerate(zip(combined_df.itertuples(index=False, name=None), totals), start=1):
        for c, v in enumerate(row):
            if v is None or (isinstance(v, float) and v != v) or v is pd.NaT:
                continue
            if isinstance(v, pd.Timestamp):
                ws.write_datetime(r, c, v.to_pydatetime(), date_fmt)
            elif c in money_idx:
                ws.write_number(r, c, float(v), money_fmt)
            elif c in etp_idx:
                ws.write_number(r, c, float(v), etp_fmt)
            else:
                ws.write(r, c, v)
        if pd.notna(tot):
            ws.write_number(r, total_col, float(tot), money_fmt)
        country_val = row[country_idx]
        if country_val not in (None, ""):
            country_counts[country_val] = country_counts.get(country_val, 0) + 1
        if permanent_idx is not None and row[permanent_idx] == "Oui":
            permanent_counts["GLOBAL"] = permanent_counts.get("GLOBAL", 0) + 1
            if country_val not in (None, ""):
                permanent_country_counts[country_val] = permanent_country_counts.get(country_val, 0) + 1
        nrows = r

    ws.freeze_panes(1, 0)
    for c, name in enumerate(data_cols):
        ws.set_column(c, c, min(max(len(str(name)) + 2, 12), 32))
    ws.set_column(total_col, total_col, 18)

    countries = sorted(country_counts.keys())
    print(f"  Pays detectes dans les donnees : {countries}")

    ws_syn = wb.add_worksheet("Synthese_Salaires")
    thin_border = wb.add_format({"border": 1, "font_name": "Arial"})
    header_border_fmt = wb.add_format({"bold": True, "font_color": "white", "bg_color": "#4472C4",
                                        "font_name": "Arial", "align": "center", "valign": "vcenter",
                                        "border": 1})
    note_fmt = wb.add_format({"italic": True, "font_size": 9, "font_color": "#808080", "font_name": "Arial"})
    money_border_fmt = wb.add_format({"num_format": '#,##0.00 "€"', "border": 1, "font_name": "Arial"})

    for i, h in enumerate(["Pays", "Nombre de salaries (tous)", "Nombre de salaries Permanent (base mediane)",
                            "Salaire Total Median ETP (EUR)"]):
        ws_syn.write(0, i, h, header_border_fmt)
    ws_syn.write(0, 5,
        "Note : effectifs (tous salaries et permanents) = valeurs calculees au moment de la "
        "generation. Mediane = salaire ramene a un ETP plein (Total_Salaires_ETP), calculee sur "
        "les salaries permanents uniquement (Salarie_Permanent = 'Oui'), en excluant le salarie "
        "au salaire/ETP le plus eleve du groupe (global ou pays). Relancer le traitement pour "
        "mettre a jour ces valeurs.",
        note_fmt)

    permanent_mask_calc = (combined_df_with_calc[PERMANENT_FLAG_COLUMN] == "Oui") \
        if PERMANENT_FLAG_COLUMN in combined_df_with_calc.columns \
        else pd.Series(False, index=combined_df_with_calc.index)
    etp_salary_calc = combined_df_with_calc[TOTAL_SALAIRES_ETP_COLUMN] \
        if TOTAL_SALAIRES_ETP_COLUMN in combined_df_with_calc.columns else totals

    global_median = median_excluding_max(etp_salary_calc[permanent_mask_calc])
    r = 1
    ws_syn.write(r, 0, "GLOBAL", thin_border)
    ws_syn.write_number(r, 1, int(sum(country_counts.values())), thin_border)
    ws_syn.write_number(r, 2, int(permanent_counts.get("GLOBAL", 0)), thin_border)
    ws_syn.write_number(r, 3, round(float(global_median), 2) if pd.notna(global_median) else 0, money_border_fmt)
    r += 1

    for country in countries:
        mask_all = combined_df_with_calc["Country"] == country
        mask_perm = mask_all & permanent_mask_calc
        med = median_excluding_max(etp_salary_calc[mask_perm])
        ws_syn.write(r, 0, country, thin_border)
        ws_syn.write_number(r, 1, int(country_counts[country]), thin_border)
        ws_syn.write_number(r, 2, int(permanent_country_counts.get(country, 0)), thin_border)
        ws_syn.write_number(r, 3, round(float(med), 2) if pd.notna(med) else 0, money_border_fmt)
        r += 1

    ws_syn.set_column(0, 0, 14)
    ws_syn.set_column(1, 1, 20)
    ws_syn.set_column(2, 2, 28)
    ws_syn.set_column(3, 3, 26)
    ws_syn.set_column(5, 5, 90)
    ws_syn.freeze_panes(1, 0)

    wb.close()
    print(f"  [xlsxwriter] {nrows} lignes ecrites")
    return countries


def build_workbook(combined_df_with_calc, all_kind_cols, output_path):
    """Choisit automatiquement le moteur d'ecriture selon le volume de
    donnees. output_path peut etre un chemin ou un objet fichier binaire
    (BytesIO). Retourne la liste des pays detectes."""
    n = len(combined_df_with_calc)
    if n > LARGE_DATASET_THRESHOLD:
        print(f"  Volume important ({n} lignes > {LARGE_DATASET_THRESHOLD}) -> moteur rapide (xlsxwriter).")
        return build_workbook_fast(combined_df_with_calc, all_kind_cols, output_path)
    else:
        return build_workbook_openpyxl(combined_df_with_calc, all_kind_cols, output_path)
