# to run code try this: python -m streamlit run app.py
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import io

st.set_page_config(page_title="Nokia KPI Dashboard", layout="wide")

st.markdown("""
<style>
.stApp { background-color: #FFFFFF; }
[data-testid="stSidebar"] { background-color: #F5F6F8; }
.stButton > button {
    width: 100%;
    font-size: 11px;
    padding: 4px 8px;
    background: #F0F2F5;
    border: 1px solid #D0D5DD;
    color: #111111 !important;
    border-radius: 4px;
    margin-bottom: 3px;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
    transition: all 0.15s;
}
.stButton > button:hover {
    border-color: #005AFF;
    color: #005AFF;
    background: #FFFFFF;
}
div[data-testid="stButton"]:first-of-type > button {
    background: #E6F0FF;
    border-color: #005AFF;
    color: #005AFF !important;
    font-weight: 600;
}
[data-testid="stMetricLabel"] p { color: #111111 !important; }
[data-testid="stMetricValue"]   { color: #111111 !important; }
[data-testid="stMetricDelta"]   { color: #111111 !important; }
.stMarkdown p, .stMarkdown h1, .stMarkdown h2, .stMarkdown h3 {
    color: #111111 !important;
}
[data-testid="stCaptionContainer"] p { color: #444444 !important; }
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] .stMarkdown p,
[data-testid="stSidebar"] .stMarkdown h3,
[data-testid="stSidebar"] span,
[data-testid="stSidebar"] p {
    color: #111111 !important;
}
</style>
""", unsafe_allow_html=True)

PALETTE = [
    "#005AFF","#0EF7F7","#023D27","#FFB800","#FF4D6A",
    "#A855F7","#FF6B35","#00E5CC","#FFE033","#FF61D2",
    "#54D46C","#851B03","#4FC3F7","#CE93D8","#80DEEA",
    "#F97316","#5CB4C4","#8B5CF6","#EC4899","#14B8A6",
]

COULEUR_ROUGE = "#FF4D6A"
COULEUR_BLEU  = "#005AFF"

if "cellule_active" not in st.session_state:
    st.session_state.cellule_active = None


# ─────────────────────────────────────────────────────────────────────────────
# Sidebar : upload + feuille
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown(" Nokia KPI Dashboard")
    st.divider()

    fichiers = st.file_uploader(
        "📂 Fichiers Excel (.xlsx / .xls)",
        type=["xlsx", "xls"],
        accept_multiple_files=True
    )
    if not fichiers:
        st.info("Déposez votre export Nokia ici.")
        st.stop()

    noms = [f.name for f in fichiers]
    nom_actif = st.selectbox("Fichier", noms) if len(fichiers) > 1 else noms[0]
    fichier_actif = next(f for f in fichiers if f.name == nom_actif)

    xls = pd.ExcelFile(fichier_actif)
    feuilles_data = [
        s for s in xls.sheet_names
        if not any(x in s for x in ["Report Execution", "Documentation"])
    ] or xls.sheet_names
    feuille = st.selectbox("Feuille", feuilles_data)


# ─────────────────────────────────────────────────────────────────────────────
# Lecture + extraction des unités (feuille Documentation)
# ─────────────────────────────────────────────────────────────────────────────
df_raw = pd.read_excel(fichier_actif, sheet_name=feuille)
df_raw.columns = df_raw.columns.astype(str).str.strip()

kpi_units = {}
try:
    for sheet in xls.sheet_names:
        if "Documentation" in sheet:
            doc = pd.read_excel(fichier_actif, sheet_name=sheet, header=None)
            header_row = None
            for i, row in doc.iterrows():
                if "KPI Alias" in row.values and "Unit" in row.values:
                    header_row = i
                    break
            if header_row is not None:
                doc.columns = doc.iloc[header_row]
                doc = doc.iloc[header_row+1:].reset_index(drop=True)
                doc.columns = doc.columns.astype(str).str.strip()
                if "KPI Alias" in doc.columns and "Unit" in doc.columns:
                    for _, row in doc.iterrows():
                        alias = str(row["KPI Alias"]).strip()
                        unit  = str(row["Unit"]).strip()
                        if alias and unit and unit != "nan":
                            kpi_units[alias] = unit
except Exception:
    pass

DATE_COL = "Period start time"

# ─────────────────────────────────────────────────────────────────────────────
# Détection des colonnes d'identification via la LIGNE GRISE (ligne juste
# sous l'en-tête) : elle est toujours VIDE pour les colonnes d'identification
# (Period start time, MRBTS/NRBTS/LNBTS name, NRCEL/LNCEL name, WS_NAME) et
# toujours REMPLIE d'un code technique (ex: "NR_5150A") pour les colonnes KPI.
# C'est la frontière fiable entre identification et KPI, peu importe le nom
# des colonnes ou la techno (4G/5G/3G).
# ─────────────────────────────────────────────────────────────────────────────
ligne_marqueur = df_raw.iloc[0]

def _est_vide(v):
    return pd.isna(v) or (isinstance(v, str) and v.strip() == "")

colonnes_identification_brutes = [
    c for c in df_raw.columns if c != DATE_COL and _est_vide(ligne_marqueur[c])
]
# WS_NAME est géré séparément (superposition réseaux TDD/FDD), ce n'est pas
# un axe de regroupement de courbes par cellule/BTS.
colonnes_equipement = [c for c in colonnes_identification_brutes if c != "WS_NAME"]

# ── Liberté de choix : l'utilisateur décide par quelle colonne grouper les
#    courbes (cellule, BTS, site...), au lieu de laisser le code deviner. ──
with st.sidebar:
    if colonnes_equipement:
        st.divider()
        st.markdown("### 🧭 Axe de regroupement des courbes")
        col_entite_choisie = st.selectbox(
            "Tracer une courbe par :",
            colonnes_equipement,
            index=len(colonnes_equipement) - 1,
            help="Choisissez la colonne qui définit une courbe distincte sur les graphiques "
                 "(ex : nom de cellule, nom de BTS, ou nom de site). Par défaut, la colonne "
                 "la plus à droite (généralement la plus fine, la cellule) est sélectionnée."
        )
    else:
        col_entite_choisie = None

has_lncel = col_entite_choisie is not None

if has_lncel:
    # ── FORMAT B : une ligne par (date, cellule/BTS choisi) ─────────────────
    df = df_raw.iloc[1:].copy()
    df.columns = df_raw.columns

    df[DATE_COL] = pd.to_datetime(df[DATE_COL], errors="coerce")
    df = df.dropna(subset=[DATE_COL]).sort_values(DATE_COL)

    col_entite = col_entite_choisie
    df[col_entite] = df[col_entite].astype(str).str.strip()

    cols_texte = set(colonnes_identification_brutes) | {DATE_COL}
    kpi_cols_disponibles = []
    for c in df.columns:
        if c in cols_texte:
            continue
        df[c] = pd.to_numeric(df[c], errors="coerce")
        if df[c].notna().sum() > 0:
            kpi_cols_disponibles.append(c)

    cellules = sorted(df[col_entite].dropna().unique().tolist(), key=str)
    cell_of_col = {}
    FORMAT = "B"

else:
    # ── FORMAT A : une ligne par date, une cellule par colonne KPI ─────────
    colonnes_kpi_brutes = [
        c for c in df_raw.columns
        if c != DATE_COL and c not in colonnes_identification_brutes
    ]
    cell_of_col = {}
    for col in colonnes_kpi_brutes:
        cell_of_col[col] = str(df_raw.iloc[0][col]).strip()

    df = df_raw.iloc[1:].copy()
    df.columns = df_raw.columns
    df[DATE_COL] = pd.to_datetime(df[DATE_COL], errors="coerce")
    df = df.dropna(subset=[DATE_COL]).sort_values(DATE_COL)

    kpi_cols_disponibles = []
    for c in colonnes_kpi_brutes:
        df[c] = pd.to_numeric(df[c], errors="coerce")
        if df[c].notna().sum() > 0:
            kpi_cols_disponibles.append(c)

    col_entite = None
    cellules = []
    FORMAT = "A"


# ─────────────────────────────────────────────────────────────────────────────
# Détection des groupes WS_NAME (TDD / FDD)
# ─────────────────────────────────────────────────────────────────────────────
ws_groups = []
ws_labels = {}
ws_color_map = {}

if "WS_NAME" in df.columns:
    bruts = df["WS_NAME"].dropna().astype(str).str.strip().unique().tolist()
    ws_groups = sorted([w for w in bruts if w not in ["", "nan"]])

    def _simplifier_ws(w):
        u = w.upper()
        if "TDD" in u:
            return "TDD"
        if "FDD" in u:
            return "FDD"
        return w

    for w in ws_groups:
        ws_labels[w] = _simplifier_ws(w)

    for w in ws_groups:
        if ws_labels[w] == "TDD":
            ws_color_map[w] = COULEUR_ROUGE
        elif ws_labels[w] == "FDD":
            ws_color_map[w] = COULEUR_BLEU

    couleurs_dispo = [COULEUR_ROUGE, COULEUR_BLEU] + PALETTE
    idx_c = 0
    for w in ws_groups:
        if w not in ws_color_map:
            while couleurs_dispo[idx_c] in ws_color_map.values():
                idx_c += 1
            ws_color_map[w] = couleurs_dispo[idx_c]
            idx_c += 1


# ─────────────────────────────────────────────────────────────────────────────
# Sidebar suite : KPIs + réseaux + options + curseurs
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    if not kpi_cols_disponibles:
        st.error("Aucun KPI numérique trouvé dans cette feuille.")
        st.stop()

    kpis_choisis = kpi_cols_disponibles

    if ws_groups:
        st.divider()
        st.markdown("### 📶 Réseaux (TDD / FDD)")
        ws_selectionnes = st.multiselect(
            "Courbes superposées",
            ws_groups,
            default=ws_groups,
            format_func=lambda w: (
                f"{ws_labels.get(w, w)} "
                f"({'🔴 rouge' if ws_color_map.get(w) == COULEUR_ROUGE else '🔵 bleu' if ws_color_map.get(w) == COULEUR_BLEU else '⚪'})"
            )
        )
        st.caption("Chaque réseau est tracé en couche séparée sur le même graphique.")
    else:
        ws_selectionnes = []

    # FIX: build ws_axis_map here (was already correct, kept as-is)
    ws_axis_map = {}
    if ws_groups and ws_selectionnes:
        for i, w in enumerate(ws_selectionnes):
            ws_axis_map[w] = "y" if i == 0 else "y2"

    st.divider()
    st.markdown("### 🎛️ Options")
    hauteur     = st.slider("Hauteur graphique (px)", 200, 700, 625, 25)
    nb_cols     = st.radio("Colonnes", [1, 2, 3], index=0, horizontal=True)
    show_points = st.checkbox("Afficher les points", False)
    marge_y     = st.slider(
        "Marge auto-échelle axe Y (%)", 0, 50, 15, 5,
        help="Ajoute une marge au-dessus et en dessous des valeurs min/max."
    )
    afficher_topworst_lignes = st.checkbox(
        "📈 Repères meilleur / pire point (lignes verticales)", False,
        help="Trace une ligne verte sur la valeur la plus haute et une ligne rouge sur la valeur la plus basse."
    )

    st.divider()
    st.markdown("### 📅 Fenêtre temporelle")
    date_min = df[DATE_COL].min().date()
    date_max = df[DATE_COL].max().date()
    intervalles = df[DATE_COL].drop_duplicates().sort_values().diff().dropna()
    granularite_horaire = bool(len(intervalles) and intervalles.min() < pd.Timedelta(days=1))
    date_debut = st.date_input(
        "Date début", value=date_min, min_value=date_min, max_value=date_max, key="date_debut"
    )
    date_fin = st.date_input(
        "Date fin", value=date_max, min_value=date_min, max_value=date_max, key="date_fin"
    )

    st.divider()
    st.markdown("### 📍 Curseurs verticaux (repères)")

    curseurs = []
    cc1, cc2 = st.columns(2)

    with cc1:
        actif_c1 = st.checkbox("Activer curseur 1", False, key="actif_curseur1")
        if actif_c1:
            label_c1 = st.text_input("Label curseur 1", "Repère 1", key="label_curseur1")
            date_c1 = st.date_input(
                "Date curseur 1", value=date_min,
                min_value=date_min, max_value=date_max, key="curseur1_date"
            )
            if granularite_horaire:
                heure_c1 = st.time_input("Heure curseur 1", value=pd.Timestamp("00:00").time(), key="curseur1_heure")
                ts_c1 = pd.Timestamp(f"{date_c1} {heure_c1}")
            else:
                ts_c1 = pd.Timestamp(date_c1)
            curseurs.append((ts_c1, label_c1, "#22C55E"))

    with cc2:
        actif_c2 = st.checkbox("Activer curseur 2", False, key="actif_curseur2")
        if actif_c2:
            label_c2 = st.text_input("Label curseur 2", "Repère 2", key="label_curseur2")
            date_c2 = st.date_input(
                "Date curseur 2", value=date_max,
                min_value=date_min, max_value=date_max, key="curseur2_date"
            )
            if granularite_horaire:
                heure_c2 = st.time_input("Heure curseur 2", value=pd.Timestamp("00:00").time(), key="curseur2_heure")
                ts_c2 = pd.Timestamp(f"{date_c2} {heure_c2}")
            else:
                ts_c2 = pd.Timestamp(date_c2)
            curseurs.append((ts_c2, label_c2, "#720316"))

    afficher_curseurs = len(curseurs) > 0


# ─────────────────────────────────────────────────────────────────────────────
# Page principale
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("<h1 style='color:#111111'>NOKIA KPI DASHBOARD</h1>", unsafe_allow_html=True)
st.markdown("<p style='color:#444444; font-size:14px'>Radio Network Performance Analytics</p>", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# Boutons cellules (FORMAT B)
# ─────────────────────────────────────────────────────────────────────────────
if FORMAT == "B" and cellules:

    n_par_ligne = 7
    tous_les_boutons = ["✦ Tous"] + cellules
    groupes = [tous_les_boutons[i:i+n_par_ligne] for i in range(0, len(tous_les_boutons), n_par_ligne)]

    for groupe in groupes:
        cols_row = st.columns(len(groupe))
        for ci, label in enumerate(groupe):
            with cols_row[ci]:
                display = label if label == "✦ Tous" else (label[:12] + "…" if len(label) > 12 else label)
                if st.button(display, key=f"btn_{label}", use_container_width=True):
                    if label == "✦ Tous":
                        st.session_state.cellule_active = None
                    else:
                        st.session_state.cellule_active = label

    cellule_active = st.session_state.cellule_active
    if cellule_active and cellule_active not in cellules:
        st.session_state.cellule_active = None
        cellule_active = None

    if cellule_active:
        st.info(f"📌 **{cellule_active}** — courbe isolée · Cliquer **✦ Tous** pour revenir")
        entites_plot = [cellule_active]
    else:
        entites_plot = cellules

else:
    entites_plot   = []
    cellule_active = None

st.divider()

# ─────────────────────────────────────────────────────────────────────────────
# Métriques
# ─────────────────────────────────────────────────────────────────────────────
c1, c2, c3, c4 = st.columns(4)
c1.metric("📊 KPIs sélectionnés", len(kpis_choisis))
c2.metric("🏢 Cellules", len(cellules) if cellules else "—")
c3.metric("📅 Points de temps", df[DATE_COL].nunique())
c4.metric("📁 Fichier", nom_actif[:24] + "…" if len(nom_actif) > 24 else nom_actif)

st.divider()

# ─────────────────────────────────────────────────────────────────────────────
# Graphiques
# ─────────────────────────────────────────────────────────────────────────────
if not kpis_choisis:
    st.warning("Sélectionnez au moins un KPI dans la barre latérale.")
    st.stop()

mode   = "lines+markers" if show_points else "lines"
grille = st.columns(nb_cols)

df_plot = df[
    (df[DATE_COL].dt.date >= date_debut) &
    (df[DATE_COL].dt.date <= date_fin)
].copy()


def ajouter_padding_invisible(fig, sub, kpi, axe, legendgroup, marge_y):
    """Ajoute un point invisible pour forcer une marge sur l'autorange Plotly."""
    valeurs = sub[kpi].dropna()
    if valeurs.empty:
        return
    vmin, vmax = float(valeurs.min()), float(valeurs.max())
    span = vmax - vmin
    pad = span * (marge_y / 100) if span > 0 else (abs(vmax) * 0.1 if vmax != 0 else 1)
    y_lo, y_hi = vmin - pad, vmax + pad
    if vmin >= 0:
        y_lo = max(0, y_lo)
    x_ref = sub[DATE_COL].iloc[0]
    fig.add_trace(go.Scatter(
        x=[x_ref, x_ref],
        y=[y_lo, y_hi],
        mode="markers",
        marker=dict(opacity=0),
        showlegend=False,
        legendgroup=legendgroup,
        hoverinfo="skip",
        yaxis=axe,
    ))


for idx, kpi in enumerate(kpis_choisis):
    with grille[idx % nb_cols]:

        fig = go.Figure()
        traces_valeurs = []


        if FORMAT == "B":
            for i, ent in enumerate(entites_plot):
                base = df_plot[df_plot[col_entite] == ent].sort_values(DATE_COL)
                if base.empty:
                    continue

                if ws_groups and ws_selectionnes:
                    for ws in ws_selectionnes:
                        sub = base[base["WS_NAME"].astype(str).str.strip() == ws]
                        if sub.empty or sub[kpi].isna().all():
                            continue
                        nom = f"{ent} · {ws_labels.get(ws, ws)}" if len(entites_plot) > 1 else ws_labels.get(ws, ws)
                        couleur = ws_color_map.get(ws, PALETTE[i % len(PALETTE)])
                        axe = ws_axis_map.get(ws, "y")
                        fig.add_trace(go.Scatter(
                            x=sub[DATE_COL],
                            y=sub[kpi],
                            name=nom,
                            mode=mode,
                            line=dict(color=couleur, width=2, shape="spline", smoothing=1.3),
                            yaxis=axe,
                            legendgroup=nom,
                            hovertemplate=(
                                f"<b>{nom}</b><br>"
                                f"{kpi}: <b>%{{y:.4g}}</b><br>"
                                f"%{{x|%d.%m.%Y %H:%M:%S}}<extra></extra>"
                            )
                        ))
                        ajouter_padding_invisible(fig, sub, kpi, axe, nom, marge_y)
                        if axe == "y":
                            traces_valeurs.append(sub[[DATE_COL, kpi]].dropna())
                else:
                    sub = base
                    if sub.empty or sub[kpi].isna().all():
                        continue
                    fig.add_trace(go.Scatter(
                        x=sub[DATE_COL],
                        y=sub[kpi],
                        name=ent,
                        mode=mode,
                        line=dict(color=PALETTE[i % len(PALETTE)], width=2, shape="spline", smoothing=1.3),
                        legendgroup=ent,
                        hovertemplate=(
                            f"<b>{ent}</b><br>"
                            f"{kpi}: <b>%{{y:.4g}}</b><br>"
                            f"%{{x|%d.%m.%Y %H:%M:%S}}<extra></extra>"
                        )
                    ))
                    ajouter_padding_invisible(fig, sub, kpi, "y", ent, marge_y)
                    traces_valeurs.append(sub[[DATE_COL, kpi]].dropna())

            titre = kpi

        else:
            cell_label = cell_of_col.get(kpi, kpi)

            if ws_groups and ws_selectionnes:
                for ws in ws_selectionnes:
                    sub = df_plot[df_plot["WS_NAME"].astype(str).str.strip() == ws][[DATE_COL, kpi]].dropna()
                    if sub.empty:
                        continue
                    nom = ws_labels.get(ws, ws)
                    couleur = ws_color_map.get(ws, PALETTE[0])
                    axe = ws_axis_map.get(ws, "y")
                    fig.add_trace(go.Scatter(
                        x=sub[DATE_COL],
                        y=sub[kpi],
                        name=nom,
                        mode=mode,
                        line=dict(color=couleur, width=2, shape="spline", smoothing=0.3),
                        yaxis=axe,
                        legendgroup=nom,
                        hovertemplate=(
                            f"<b>{nom}</b><br>"
                            f"{kpi}: <b>%{{y:.4g}}</b><br>"
                            f"%{{x|%d.%m.%Y %H:%M:%S}}<extra></extra>"
                        )
                    ))
                    ajouter_padding_invisible(fig, sub, kpi, axe, nom, marge_y)
                    if axe == "y":
                        traces_valeurs.append(sub)
            else:
                sub = df_plot[[DATE_COL, kpi]].dropna()
                fig.add_trace(go.Scatter(
                    x=sub[DATE_COL],
                    y=sub[kpi],
                    name=cell_label,
                    mode=mode,
                    line=dict(color=PALETTE[0], width=2, shape="spline", smoothing=1.3),
                    legendgroup=cell_label,
                    hovertemplate=(
                        f"<b>{cell_label}</b><br>"
                        f"{kpi}: <b>%{{y:.4g}}</b><br>"
                        f"%{{x|%d.%m.%Y %H:%M:%S}}<extra></extra>"
                    )
                ))
                ajouter_padding_invisible(fig, sub, kpi, "y", cell_label, marge_y)
                traces_valeurs.append(sub)

            titre = f"{kpi}  <span style='font-size:11px;color:#8B949E'>({cell_label})</span>"

        # ── Top/Worst ────────────────────────────────────────────────────────
        points_topworst = []
        reseau_ambigu = bool(ws_groups) and len(ws_selectionnes) != 1
        label_reseau = (
            ws_labels.get(ws_selectionnes[0], ws_selectionnes[0])
            if (ws_groups and len(ws_selectionnes) == 1) else None
        )

        if afficher_topworst_lignes and not reseau_ambigu and traces_valeurs:
            df_vals_tw = pd.concat(traces_valeurs, ignore_index=True).dropna(subset=[kpi])
            if not df_vals_tw.empty:
                pmax = df_vals_tw.loc[df_vals_tw[kpi].idxmax()]
                pmin = df_vals_tw.loc[df_vals_tw[kpi].idxmin()]
                points_topworst.append({
                    "label": label_reseau,
                    "ts_max": pmax[DATE_COL], "val_max": pmax[kpi],
                    "ts_min": pmin[DATE_COL], "val_min": pmin[kpi],
                })

        unite = kpi_units.get(kpi, "")

        # FIX: decide dual-axis from ws_selectionnes, not from ranges_par_axe keys
        deux_reseaux = bool(ws_groups) and len(ws_selectionnes) >= 2
        label_axe_y  = ws_labels.get(ws_selectionnes[0], ws_selectionnes[0]) if deux_reseaux else ""
        label_axe_y2 = ws_labels.get(ws_selectionnes[1], ws_selectionnes[1]) if deux_reseaux else ""
        couleur_axe_y  = ws_color_map.get(ws_selectionnes[0], "#111111") if deux_reseaux else "#111111"
        couleur_axe_y2 = ws_color_map.get(ws_selectionnes[1], "#111111") if deux_reseaux else "#111111"

        # ── Auto-échelle TOUJOURS dynamique : Plotly recalcule l'échelle à
        # partir des seules courbes (+ leur marge invisible associée) qui
        # sont effectivement visibles, donc l'isolement via la légende
        # déclenche automatiquement le bon zoom, sans action manuelle.
        yaxis_cfg = dict(
            gridcolor="#E5E7EB",
            title=dict(
                text=f"{unite} {label_axe_y}".strip(),
                font=dict(size=13, color=couleur_axe_y),
                standoff=10,
            ),
            tickfont=dict(color="#111111"),
            linecolor="#111111",
            zeroline=False,
            autorange=True,
        )

        yaxis2_cfg = None
        if deux_reseaux:
            yaxis2_cfg = dict(
                overlaying="y",
                side="right",
                showgrid=False,
                title=dict(
                    text=f"{unite} {label_axe_y2}".strip(),
                    font=dict(size=13, color=couleur_axe_y2),
                    standoff=10,
                ),
                tickfont=dict(color="#111111"),
                linecolor="#111111",
                zeroline=False,
                autorange=True,
            )

        xaxis_cfg = dict(
            gridcolor="#E5E7EB",
            title="",
            type="date",
            tickangle=-45,
            showticklabels=True,
            tickfont=dict(color="#111111", size=10),
            linecolor="#111111",
            showspikes=False,
            showline=True,
            zeroline=False,
            tickformatstops=[
                dict(dtickrange=[None, 3600000],        value="%H:%M\n%d.%m"),
                dict(dtickrange=[3600000, 86400000],    value="%H:%M\n%d.%m"),
                dict(dtickrange=[86400000, None],       value="%d.%m.%Y"),
            ],
        )

        marge_droite = 260 if yaxis2_cfg else 200
        legend_x     = 1.20 if yaxis2_cfg else 1.01

        fig.update_layout(
            title=dict(text=titre, font=dict(size=20, color="#005AFF"), x=0.5, xanchor="center"),
            template="plotly_white",
            height=hauteur,
            margin=dict(l=55, r=marge_droite, t=50, b=60),
            hovermode="closest",
            paper_bgcolor="#FFFFFF",
            plot_bgcolor="#FFFFFF",
            xaxis=xaxis_cfg,
            yaxis=yaxis_cfg,
            legend=dict(
                orientation="v",
                yanchor="top", y=1,
                xanchor="left", x=legend_x,
                font=dict(size=9, color="#111111"),
                bgcolor="rgba(255,255,255,0.95)",
                bordercolor="#D0D5DD",
                borderwidth=1,
                itemclick="toggleothers",
                itemdoubleclick="toggle"
            )
        )
        if yaxis2_cfg:
            fig.update_layout(yaxis2=yaxis2_cfg)

        if afficher_topworst_lignes and reseau_ambigu:
            st.caption("⚠️ Sélectionnez **un seul** réseau (TDD ou FDD) dans la barre latérale "
                       "pour activer les repères Meilleur/Pire sur ce graphique.")

        if afficher_curseurs:
            for k_c, (c_ts, c_label, c_couleur) in enumerate(curseurs):
                fig.add_shape(
                    type="line",
                    xref="x", yref="paper",
                    x0=c_ts, x1=c_ts,
                    y0=0, y1=1,
                    line=dict(color=c_couleur, width=1.5, dash="dash"),
                )
                fig.add_annotation(
                    x=c_ts, y=1, xref="x", yref="paper",
                    yshift=-14 * k_c,
                    text=c_label,
                    showarrow=False,
                    yanchor="bottom",
                    font=dict(size=10, color=c_couleur),
                    bgcolor="rgba(255,255,255,0.85)",
                )

        if afficher_topworst_lignes and points_topworst:
            decalage_initial = 14 * len(curseurs) if afficher_curseurs else 0
            for j, pt in enumerate(points_topworst):
                prefix = f"{pt['label']} — " if pt["label"] else ""
                shift_max = -(decalage_initial + 14 * (2 * j))
                shift_min = -(decalage_initial + 14 * (2 * j + 1))

                fig.add_shape(
                    type="line", xref="x", yref="paper",
                    x0=pt["ts_max"], x1=pt["ts_max"], y0=0, y1=1,
                    line=dict(color="#16A34A", width=1.5, dash="dot"),
                )
                fig.add_annotation(
                    x=pt["ts_max"], y=1, xref="x", yref="paper",
                    yshift=shift_max,
                    text=f"🟢 {prefix}Meilleur: {pt['val_max']:.4g} ({pt['ts_max']:%d.%m %H:%M})",
                    showarrow=False,
                    yanchor="bottom",
                    font=dict(size=9, color="#16A34A"),
                    bgcolor="rgba(255,255,255,0.85)",
                )

                fig.add_shape(
                    type="line", xref="x", yref="paper",
                    x0=pt["ts_min"], x1=pt["ts_min"], y0=0, y1=1,
                    line=dict(color="#DC2626", width=1.5, dash="dot"),
                )
                fig.add_annotation(
                    x=pt["ts_min"], y=1, xref="x", yref="paper",
                    yshift=shift_min,
                    text=f"🔴 {prefix}Pire: {pt['val_min']:.4g} ({pt['ts_min']:%d.%m %H:%M})",
                    showarrow=False,
                    yanchor="bottom",
                    font=dict(size=9, color="#DC2626"),
                    bgcolor="rgba(255,255,255,0.85)",
                )

        st.plotly_chart(
            fig,
            use_container_width=True,
            config={
                "scrollZoom": True,
                "displaylogo": False,
                "modeBarButtonsToRemove": ["lasso2d", "select2d"]
            }
        )


# ─────────────────────────────────────────────────────────────────────────────
# Export
# ─────────────────────────────────────────────────────────────────────────────
st.divider()
e1, e2 = st.columns(2)

if FORMAT == "B":
    df_export = df[df[col_entite].isin(entites_plot)].copy() if entites_plot else df.copy()
else:
    df_export = df.copy()

if ws_groups and ws_selectionnes:
    df_export = df_export[df_export["WS_NAME"].astype(str).str.strip().isin(ws_selectionnes)]

with e1:
    st.download_button(
        "⬇️ Télécharger CSV",
        df_export.to_csv(index=False).encode("utf-8"),
        "nokia_export.csv", "text/csv"
    )
with e2:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df_export.to_excel(w, index=False)
    st.download_button(
        "⬇️ Télécharger Excel",
        buf.getvalue(), "nokia_export.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

with st.expander("🔎 Données brutes"):
    st.dataframe(df_export.head(500), use_container_width=True)