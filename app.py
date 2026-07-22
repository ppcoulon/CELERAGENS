#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
app.py (Celeragens)
--------------------
Application web (Streamlit) pour generer le rapport RH combine sans passer
par la ligne de commande : upload des fichiers sources, saisie/verification
de la devise et du taux de conversion vers EUR par fichier, generation du
classeur (Feuil1 + Synthese_Salaires) et telechargement.

Version brandee Celeragens de l'application (meme moteur que la version
Karbonpath) : seule l'identite visuelle change (logo, nom, couleurs). Toute
la logique metier vit dans rapport_core.py (identique dans les deux
versions) : les deux applications produisent des fichiers strictement
identiques pour les memes fichiers/taux.

Lancement local :
    streamlit run app.py
Deploiement en ligne : voir DEPLOIEMENT.md.
"""

import base64
import contextlib
import io
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st

import rapport_core as rc

# ---------------------------------------------------------------------------
# Charte graphique Celeragens
# ---------------------------------------------------------------------------
# Couleurs extraites directement du logo fourni (logo_celeragens.png) :
# anneau/texte "C" en teal, point et soulignement en turquoise vif, texte en
# bleu marine fonce, sous-titre en gris. Le theme des widgets natifs
# (boutons, sliders, focus...) est defini dans .streamlit/config.toml
# (mecanisme officiel Streamlit) ; le CSS ci-dessous ne fait que l'entete de
# marque et quelques finitions visuelles.
CG_TEAL = "#008080"
CG_TEAL_DARK = "#005A5A"
CG_TEAL_LIGHT = "#339999"
CG_ACCENT = "#00C8B4"
CG_BG = "#F7FBFB"
CG_SOFT = "#EBF5F5"
CG_BORDER = "#BFDFDF"
CG_NAVY = "#1E2832"
CG_GRAY = "#646E78"

LOGO_PATH = Path(__file__).parent / "celeragens_logo.png"


def _logo_data_uri():
    if LOGO_PATH.exists():
        b64 = base64.b64encode(LOGO_PATH.read_bytes()).decode()
        return f"data:image/png;base64,{b64}"
    return None


st.set_page_config(
    page_title="Celeragens — Rapport RH",
    page_icon=str(LOGO_PATH) if LOGO_PATH.exists() else "🌀",
    layout="wide",
)

_logo_uri = _logo_data_uri()

st.markdown(
    f"""
    <style>
    .cg-header {{
        display: flex;
        align-items: center;
        gap: 22px;
        padding: 18px 28px;
        margin: -1rem -1rem 1.5rem -1rem;
        background: #FFFFFF;
        border-bottom: 3px solid {CG_TEAL};
    }}
    .cg-header img {{ height: 44px; }}
    .cg-header .cg-header-text {{ display: flex; flex-direction: column; }}
    .cg-header .cg-header-title {{
        font-size: 1.05rem; font-weight: 600; color: {CG_NAVY};
        letter-spacing: 0.01em;
    }}
    .cg-header .cg-header-sub {{ font-size: 0.85rem; color: {CG_GRAY}; }}

    div[data-testid="stForm"] {{
        background: {CG_BG};
        border: 1px solid {CG_BORDER};
        border-radius: 10px;
        padding: 1.2rem 1.4rem;
    }}
    div[data-testid="stExpander"] {{
        border: 1px solid {CG_BORDER};
        border-radius: 8px;
    }}
    h1, h2, h3 {{ color: {CG_NAVY}; }}
    a {{ color: {CG_TEAL_LIGHT}; }}
    #MainMenu, footer {{ visibility: hidden; }}
    </style>

    <div class="cg-header">
        {f'<img src="{_logo_uri}" />' if _logo_uri else ''}
        <div class="cg-header-text">
            <span class="cg-header-title">Générateur de rapport RH</span>
            <span class="cg-header-sub">Agrégation, conversion EUR, filtre permanents,
                ETP et médiane salariale</span>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.expander("Comment ça marche", expanded=False):
    st.markdown(
        "- Chaque fichier chargé est harmonisé sur un schéma de colonnes commun, "
        "converti en EUR (`Montant_EUR = Montant_devise / Taux`), puis tous les "
        "fichiers actifs sont concaténés dans un seul rapport.\n"
        "- Un salarié est considéré **permanent** selon la colonne `Employee_Type` et "
        "une liste de valeurs propre à chaque fichier (fixée dans le programme, non "
        "modifiable ici).\n"
        "- `ETP` = heures travaillées / heures temps plein légal annuel du pays (fixé "
        "par fichier). `Total_Salaires_ETP` = salaire total / ETP : le salaire ramené à "
        "un temps plein.\n"
        "- La médiane (globale et par pays, onglet *Synthese_Salaires*) est calculée "
        "sur `Total_Salaires_ETP` des salariés permanents uniquement, en excluant le "
        "salarié au salaire/ETP le plus élevé de chaque groupe.\n"
        "- Seuls le nom de devise et le taux de conversion sont modifiables ici. Les "
        "valeurs \"permanent\", les heures temps plein et les colonnes d'avantages en "
        "nature spécifiques restent fixées dans le programme pour les 10 fichiers "
        "connus (DBFRAA, DBCH, DBFRTW, DBIRLAA, DBITTW, DBMASER, DBMOTW, DBSPTW, "
        "DBUKAA, DBUSTW)."
    )

uploaded_files = st.file_uploader(
    "Fichiers sources (.xlsx)", type=["xlsx"], accept_multiple_files=True,
)

if not uploaded_files:
    st.info("Chargez au moins un fichier source pour continuer.")
    st.stop()

st.subheader("Devise et taux de conversion par fichier")
st.caption("Montant_EUR = Montant_devise ÷ Taux. Laissez 1.0 si le fichier est déjà en EUR.")

with st.form("config_form"):
    rows_config = []
    for i, f in enumerate(uploaded_files):
        source_code = rc.source_code_from_filename(f.name)
        default_devise, default_taux = rc.KNOWN_DEFAULT_DEVISE.get(source_code, ("EUR", 1.0))

        cols = st.columns([3, 1.3, 1.3, 0.9, 2.2])
        cols[0].markdown(f"**{f.name}**")
        if source_code is None:
            cols[0].caption(
                "⚠️ nom de fichier non reconnu (format attendu : `XXX_source...`) — "
                "filtre permanent et ETP désactivés pour ce fichier."
            )
        elif source_code not in rc.PERMANENT_VALUES or source_code not in rc.FULL_TIME_HOURS:
            cols[0].caption(
                f"⚠️ code source « {source_code} » reconnu mais absent de la "
                "configuration permanent/ETP — filtre permanent et ETP désactivés "
                "(à compléter dans rapport_core.py)."
            )
        else:
            cols[0].caption(f"code source détecté : {source_code}")

        devise = cols[1].text_input("Devise", value=default_devise, key=f"devise_{i}")
        taux = cols[2].number_input(
            "Taux vers EUR", value=float(default_taux), min_value=0.0, format="%.5f", key=f"taux_{i}",
        )
        actif = cols[3].checkbox("Actif", value=True, key=f"actif_{i}")

        pays_force = rc.KNOWN_PAYS_FORCE.get(source_code)
        extra_kind_cols = rc.KNOWN_EXTRA_KIND_COLUMNS.get(source_code, [])
        note = []
        if pays_force:
            note.append(f"pays forcé : {pays_force}")
        if extra_kind_cols:
            note.append(f"{len(extra_kind_cols)} colonne(s) d'avantage en nature spécifique(s)")
        cols[4].caption(" · ".join(note) if note else "")

        rows_config.append({
            "file": f, "filename": f.name, "source_code": source_code,
            "devise": devise, "taux": taux, "actif": actif,
            "pays_force": pays_force, "extra_kind_cols": extra_kind_cols,
        })

    submitted = st.form_submit_button("Générer le rapport", type="primary")

if not submitted:
    st.stop()

log_buffer = io.StringIO()
frames = []
all_kind_cols = []
errors = []

with st.spinner("Lecture et harmonisation des fichiers sources..."):
    with contextlib.redirect_stdout(log_buffer):
        for row in rows_config:
            if not row["actif"]:
                print(f"  [ignore] '{row['filename']}' désactivé par l'utilisateur.")
                continue
            try:
                df, extra_cols = rc.load_and_prepare_source_obj(
                    row["file"], row["filename"], devise=row["devise"], taux=row["taux"],
                    pays_force=row["pays_force"], extra_kind_cols=row["extra_kind_cols"],
                )
            except Exception as exc:  # fichier illisible / format inattendu
                errors.append(f"{row['filename']} : {exc}")
                continue
            if df is not None:
                frames.append(df)
                for c in extra_cols:
                    if c not in all_kind_cols:
                        all_kind_cols.append(c)

for e in errors:
    st.error(f"Erreur de lecture — {e}")

if not frames:
    st.warning("Aucun fichier n'a pu être chargé (voir les erreurs ci-dessus).")
    st.stop()

with st.spinner("Agrégation et calculs (Total_Salaires, ETP)..."):
    with contextlib.redirect_stdout(log_buffer):
        combined = pd.concat(frames, ignore_index=True)
        combined = rc.add_calculated_columns(combined, all_kind_cols)

with st.spinner("Génération du classeur Excel..."):
    output_buffer = io.BytesIO()
    with contextlib.redirect_stdout(log_buffer):
        countries = rc.build_workbook(combined, all_kind_cols, output_buffer)
    output_buffer.seek(0)

st.success(f"Rapport généré : {len(combined):,} lignes, {len(countries)} pays détectés.".replace(",", " "))

# --- Synthèse des médianes (identique à l'onglet Synthese_Salaires du fichier) ---
permanent_mask = combined[rc.PERMANENT_FLAG_COLUMN] == "Oui"
etp_salary = combined[rc.TOTAL_SALAIRES_ETP_COLUMN]

summary_rows = [{
    "Pays": "GLOBAL",
    "Nombre de salariés (tous)": int(len(combined)),
    "Nombre de salariés permanents (base médiane)": int(permanent_mask.sum()),
    "Salaire Total Médian ETP (EUR)": rc.median_excluding_max(etp_salary[permanent_mask]),
}]
for country in countries:
    mask_all = combined["Country"] == country
    mask_perm = mask_all & permanent_mask
    summary_rows.append({
        "Pays": country,
        "Nombre de salariés (tous)": int(mask_all.sum()),
        "Nombre de salariés permanents (base médiane)": int(mask_perm.sum()),
        "Salaire Total Médian ETP (EUR)": rc.median_excluding_max(etp_salary[mask_perm]),
    })
summary_df = pd.DataFrame(summary_rows)
summary_df["Salaire Total Médian ETP (EUR)"] = summary_df["Salaire Total Médian ETP (EUR)"].astype(float)
summary_df["Salaire Total Médian ETP (EUR)"] = summary_df["Salaire Total Médian ETP (EUR)"].replace(
    {np.nan: None}
)

st.subheader("Synthèse des médianes")
st.dataframe(
    summary_df.style.format({
        "Nombre de salariés (tous)": "{:,}",
        "Nombre de salariés permanents (base médiane)": "{:,}",
        "Salaire Total Médian ETP (EUR)": lambda v: "—" if v is None else f"{v:,.2f} €",
    }),
    use_container_width=True,
    hide_index=True,
)

st.download_button(
    "⬇️ Télécharger le rapport (.xlsx)",
    data=output_buffer,
    file_name="DB_Combined_EUR.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)

with st.expander("Journal de traitement (détails, avertissements)"):
    st.text(log_buffer.getvalue() or "(aucun message)")
