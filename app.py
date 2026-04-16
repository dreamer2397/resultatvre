import io
import unicodedata

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="Reporting financier M49", page_icon="📊", layout="wide")

st.markdown(
    """
    <style>
    .block-container {padding-top: 1rem;}
    .stMetric {background: #f7f9fc; border: 1px solid #e8edf5; border-radius: 10px; padding: 0.6rem;}
    </style>
    """,
    unsafe_allow_html=True,
)

# ── Référentiels (mettre à jour chaque année si nécessaire) ───────────────────
POPULATION: dict[str, int] = {
    "VRE":   112_899,
    "SIEPV":  27_385,
    "SIERS":  12_544,
    "DELMON":  1_209,
    "DELROM": 36_296,
}

ENCOURS_DETTE: dict[str, float] = {
    "VRE":   20_431_186,
    "SIEPV":  1_574_367,
    "SIERS":    109_461,
    "DELMON":         0,
    "DELROM": 2_301_447,
}

SECTION_LABELS = {"F": "Fonctionnement", "I": "Investissement"}

# M49 — chapitres d'ordre (non-cash, exclus du calcul de l'épargne)
ORDRE_DEP_F  = {"002", "022", "023", "040", "041"}
ORDRE_REC_F  = {"002", "021", "042"}
ORDRE_DEP_I  = {"001", "021", "040", "041"}
CHAP_CAPITAL = {"16"}   # remboursement du capital de la dette



def normalize(value: str) -> str:
    text = unicodedata.normalize("NFKD", str(value)).encode("ascii", "ignore").decode("ascii")
    return " ".join(text.lower().strip().split())


def resolve_columns(df: pd.DataFrame) -> tuple[dict[str, str], dict[str, str]]:
    norm = {normalize(c): c for c in df.columns}
    required = {
        "code":        ["code"],
        "groupe_sens": ["groupe sens", "sens"],
        "section":     ["section"],
        "ca_n":        ["ca 2025", "ca n", "realise", "realise ca n"],
        "budget":      ["mt vote cp", "montant vote cp", "budget vote"],
        "ca_n_1":      ["ca 2024", "ca n-1", "ca n- 1", "ca n 1"],
        "engage":      ["mt engage ht", "montant engage ht", "engage ht"],
        "chapitre":    ["chapitre", "chapter", "chap"],
    }
    optional = {
        "article":          ["article"],
        "libelle_chapitre": [
            "libelle chapitre", "libelle du chapitre", "lib chapitre",
            "chapitre nat. (code / libelle)", "chapitre nat (code libelle)",
        ],
        "libelle_article": [
            "libelle article", "article nat. (code / libelle)",
            "article nat (code libelle)",
        ],
        "tiers": ["tiers"],
    }

    req: dict[str, str] = {}
    for key, aliases in required.items():
        for alias in aliases:
            if alias in norm:
                req[key] = norm[alias]
                break

    missing = [k for k in required if k not in req]
    if missing:
        raise KeyError(
            "Colonnes manquantes ou non reconnues : "
            + ", ".join(missing)
            + ". Vérifiez le format du grand livre."
        )

    opt: dict[str, str] = {}
    for key, aliases in optional.items():
        for alias in aliases:
            if alias in norm:
                opt[key] = norm[alias]
                break

    return req, opt


@st.cache_data(show_spinner="Lecture du fichier Excel…")
def load_excel(file_bytes: bytes) -> pd.DataFrame:
    df = pd.read_excel(io.BytesIO(file_bytes))
    req, opt = resolve_columns(df)

    parsed = pd.DataFrame({
        "CODE":        df[req["code"]].astype(str).str.strip(),
        "Groupe Sens": df[req["groupe_sens"]].astype(str).str.strip(),
        "Section":     df[req["section"]].astype(str).str.strip(),
        "Chapitre":    df[req["chapitre"]].astype(str).str.strip(),
        "CA N":        pd.to_numeric(df[req["ca_n"]],    errors="coerce").fillna(0.0),
        "Budget voté": pd.to_numeric(df[req["budget"]],  errors="coerce").fillna(0.0),
        "CA N-1":      pd.to_numeric(df[req["ca_n_1"]], errors="coerce").fillna(0.0),
        "Engagé HT":   pd.to_numeric(df[req["engage"]], errors="coerce").fillna(0.0),
    })

    for key, col in [
        ("article",          "Article"),
        ("libelle_chapitre", "Libellé chapitre"),
        ("libelle_article",  "Libellé article"),
        ("tiers",            "Tiers"),
    ]:
        if key in opt:
            parsed[col] = df[opt[key]].astype(str).str.strip()

    parsed["Section"] = parsed["Section"].map(lambda s: SECTION_LABELS.get(s, s))
    parsed = parsed[parsed["Chapitre"].ne("")]
    return parsed


# ─────────────────────────────────────────────────────────────────────────────
# Utilitaires
# ─────────────────────────────────────────────────────────────────────────────

def fmt(value: float, decimals: int = 2) -> str:
    """Format en M€ si >= 1 M, sinon en euros entiers."""
    if abs(value) >= 1_000_000:
        return f"{value / 1_000_000:.{decimals}f} M€"
    return f"{value:_.0f} €".replace("_", "\u202f")


def fmt_hab(value: float | None) -> str:
    return f"{value:.2f} €" if value is not None else "N/A"


def delta_str(new: float, old: float) -> str | None:
    if old == 0:
        return None
    return f"{(new - old) / abs(old):.2%}"


def get_section(df: pd.DataFrame, sens: str, section: str) -> pd.DataFrame:
    return df[(df["Groupe Sens"] == sens) & (df["Section"] == section)]


def sum_hors(df: pd.DataFrame, col: str, exclus: set[str]) -> float:
    """Somme d'une colonne en excluant certains chapitres."""
    return df.loc[~df["Chapitre"].isin(exclus), col].sum()


def pop_total(codes: list[str]) -> int:
    return sum(POPULATION.get(c, 0) for c in codes) if codes else sum(POPULATION.values())


def encours_total(codes: list[str]) -> float:
    return sum(ENCOURS_DETTE.get(c, 0) for c in codes) if codes else sum(ENCOURS_DETTE.values())


# ─────────────────────────────────────────────────────────────────────────────
# Jauge Plotly
# ─────────────────────────────────────────────────────────────────────────────

def gauge_chart(
    value: float,
    title: str,
    max_val: float,
    suffix: str = "",
    steps: list[tuple[float, float, str]] | None = None,
) -> go.Figure:
    if steps is None:
        steps = [
            (0, max_val * 0.5, "#dcfce7"),
            (max_val * 0.5, max_val * 0.8, "#fef9c3"),
            (max_val * 0.8, max_val * 1.1, "#fee2e2"),
        ]
    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=value,
        number={"suffix": suffix, "valueformat": ".2f"},
        gauge={
            "axis": {"range": [0, max_val * 1.1], "ticksuffix": suffix},
            "bar":  {"color": "#2563eb"},
            "steps": [{"range": [lo, hi], "color": col} for lo, hi, col in steps],
        },
        title={"text": title, "font": {"size": 12}},
    ))
    fig.update_layout(height=220, margin=dict(t=60, b=10, l=20, r=20))
    return fig


# ─────────────────────────────────────────────────────────────────────────────
# Table par chapitre
# ─────────────────────────────────────────────────────────────────────────────

def render_chapter_table(df: pd.DataFrame, title: str) -> None:
    group_cols = ["Chapitre"]
    if "Libellé chapitre" in df.columns:
        group_cols.append("Libellé chapitre")
    tbl = (
        df.groupby(group_cols, as_index=False)[["CA N-1", "CA N"]]
        .sum()
        .sort_values("Chapitre")
    )
    tbl = tbl.rename(columns={"CA N-1": "CA 2024", "CA N": "CA 2025"})
    total = {"Chapitre": "Total", "CA 2024": tbl["CA 2024"].sum(), "CA 2025": tbl["CA 2025"].sum()}
    if "Libellé chapitre" in tbl.columns:
        total["Libellé chapitre"] = ""
    tbl = pd.concat([tbl, pd.DataFrame([total])], ignore_index=True)
    st.markdown(f"**{title}**")
    st.dataframe(
        tbl,
        use_container_width=True,
        hide_index=True,
        column_config={
            "CA 2024": st.column_config.NumberColumn("CA 2024 (€)", format="%.2f"),
            "CA 2025": st.column_config.NumberColumn("CA 2025 (€)", format="%.2f"),
        },
    )


# ─────────────────────────────────────────────────────────────────────────────
# Onglet Fonctionnement
# ─────────────────────────────────────────────────────────────────────────────

def tab_fonctionnement(df: pd.DataFrame, selected_codes: list[str]) -> None:
    dep_f = get_section(df, "Dépense", "Fonctionnement")
    rec_f = get_section(df, "Recette", "Fonctionnement")
    dep_i = get_section(df, "Dépense", "Investissement")
    pop = pop_total(selected_codes)

    def calc(col: str) -> dict:
        dep_reel = sum_hors(dep_f, col, ORDRE_DEP_F)
        rec_reel = sum_hors(rec_f, col, ORDRE_REC_F)
        eb = rec_reel - dep_reel
        capital = dep_i[dep_i["Chapitre"].isin(CHAP_CAPITAL)][col].sum()
        en = eb - capital
        return {
            "resultat":      rec_f[col].sum() - dep_f[col].sum(),
            "epargne_brute": eb,
            "epargne_nette": en,
            "eb_hab":        eb / pop if pop else None,
            "en_hab":        en / pop if pop else None,
        }

    kn  = calc("CA N")
    kn1 = calc("CA N-1")

    col_left, col_right = st.columns([1.3, 1])

    with col_left:
        render_chapter_table(dep_f, "Dépenses de Fonctionnement")
        render_chapter_table(rec_f, "Recettes de Fonctionnement")

    with col_right:
        st.markdown("#### Résultat de Fonctionnement")
        r1, r2, r3 = st.columns(3)
        r1.metric("2025", fmt(kn["resultat"]))
        r2.metric("2024", fmt(kn1["resultat"]))
        r3.metric("Évolution", delta_str(kn["resultat"], kn1["resultat"]) or "N/A",
                  delta=delta_str(kn["resultat"], kn1["resultat"]))

        st.divider()
        st.markdown("#### Épargne Brute")
        e1, e2, e3 = st.columns(3)
        e1.metric("2025", fmt(kn["epargne_brute"]))
        e2.metric("2024", fmt(kn1["epargne_brute"]))
        e3.metric("Évolution", "",
                  delta=delta_str(kn["epargne_brute"], kn1["epargne_brute"]))
        h1, h2 = st.columns(2)
        h1.metric("Par habitant 2025", fmt_hab(kn["eb_hab"]))
        h2.metric("Par habitant 2024", fmt_hab(kn1["eb_hab"]))

        st.divider()
        st.markdown("#### Épargne Nette")
        n1, n2, n3 = st.columns(3)
        n1.metric("2025", fmt(kn["epargne_nette"]))
        n2.metric("2024", fmt(kn1["epargne_nette"]))
        n3.metric("Évolution", "",
                  delta=delta_str(kn["epargne_nette"], kn1["epargne_nette"]))
        nh1, nh2 = st.columns(2)
        nh1.metric("Par habitant 2025", fmt_hab(kn["en_hab"]))
        nh2.metric("Par habitant 2024", fmt_hab(kn1["en_hab"]))


# ─────────────────────────────────────────────────────────────────────────────
# Onglet Investissement
# ─────────────────────────────────────────────────────────────────────────────

def tab_investissement(df: pd.DataFrame, selected_codes: list[str]) -> None:
    dep_i = get_section(df, "Dépense", "Investissement")
    rec_i = get_section(df, "Recette", "Investissement")
    dep_f = get_section(df, "Dépense", "Fonctionnement")
    rec_f = get_section(df, "Recette", "Fonctionnement")
    pop     = pop_total(selected_codes)
    encours = encours_total(selected_codes)

    def calc_inv(col: str) -> dict:
        effort = sum_hors(dep_i, col, ORDRE_DEP_I | CHAP_CAPITAL)
        return {
            "resultat": rec_i[col].sum() - dep_i[col].sum(),
            "effort":   effort,
            "inv_hab":  effort / pop if pop else None,
        }

    def calc_fct_eb_en(col: str) -> tuple[float, float]:
        dep_reel = sum_hors(dep_f, col, ORDRE_DEP_F)
        rec_reel = sum_hors(rec_f, col, ORDRE_REC_F)
        eb = rec_reel - dep_reel
        capital = dep_i[dep_i["Chapitre"].isin(CHAP_CAPITAL)][col].sum()
        return eb, eb - capital

    kn  = calc_inv("CA N")
    kn1 = calc_inv("CA N-1")
    eb_n, en_n = calc_fct_eb_en("CA N")

    delai = encours / eb_n if eb_n > 0 else None
    taf   = en_n / kn["effort"] if kn["effort"] > 0 else None

    col_left, col_right = st.columns([1.3, 1])

    with col_left:
        render_chapter_table(dep_i, "Dépenses d'Investissement")
        render_chapter_table(rec_i, "Recettes d'Investissement")

    with col_right:
        st.markdown("#### Résultat d'Investissement")
        ri1, ri2 = st.columns(2)
        ri1.metric("2025", fmt(kn["resultat"]))
        ri2.metric("2024", fmt(kn1["resultat"]))

        st.divider()
        st.markdown("#### Effort d'Investissement")
        ei1, ei2 = st.columns(2)
        ei1.metric("2025", fmt(kn["effort"]))
        ei2.metric("Ratio / Habitant 2025", fmt_hab(kn["inv_hab"]))

        st.divider()
        st.markdown("#### Encours de la dette")
        ed1, ed2 = st.columns(2)
        ed1.metric("2025", fmt(encours))
        ed2.metric("Ratio / Habitant", fmt_hab(encours / pop if pop else None))

        st.divider()
        g1, g2 = st.columns(2)
        with g1:
            if delai is not None:
                st.plotly_chart(
                    gauge_chart(
                        delai, "Délai Désendettement (années)", max_val=10, suffix=" ans",
                        steps=[(0, 5, "#dcfce7"), (5, 8, "#fef9c3"), (8, 11, "#fee2e2")],
                    ),
                    use_container_width=True,
                )
        with g2:
            if taf is not None:
                st.plotly_chart(
                    gauge_chart(
                        taf * 100, "Taux Autofinancement Invest.", max_val=100, suffix="%",
                        steps=[(0, 30, "#fee2e2"), (30, 60, "#fef9c3"), (60, 110, "#dcfce7")],
                    ),
                    use_container_width=True,
                )


# ─────────────────────────────────────────────────────────────────────────────
# Onglet Coût de Fonctionnement
# ─────────────────────────────────────────────────────────────────────────────

def tab_cout_fonctionnement(df: pd.DataFrame, selected_codes: list[str]) -> None:
    dep_f = get_section(df, "Dépense", "Fonctionnement")
    rec_f = get_section(df, "Recette", "Fonctionnement")
    pop   = pop_total(selected_codes)

    st.info(
        "Le **Coût de fonctionnement** est calculé depuis les dépenses réelles de fonctionnement "
        "(hors chapitres d'ordre M49). Configurez ci-dessous les chapitres propres à votre entité "
        "pour obtenir les sous-catégories *hors travaux*, *hors vente eau* et *hors frais de structure*."
    )

    with st.expander("⚙️ Configuration des catégories", expanded=True):
        c1, c2, c3 = st.columns(3)
        chap_dep_all = sorted(dep_f["Chapitre"].dropna().unique())
        chap_rec_all = sorted(rec_f["Chapitre"].dropna().unique())

        chap_travaux = c1.multiselect(
            "Chapitres 'Travaux' (exclus des dépenses)",
            chap_dep_all,
            help="Ex : articles 615x, chapitre 21 passé en fonctionnement…",
        )
        chap_vente_eau = c2.multiselect(
            "Chapitres 'Vente eau' (recettes à neutraliser)",
            chap_rec_all,
            default=[c for c in chap_rec_all if c.startswith("70")],
        )
        chap_frais_struct = c3.multiselect(
            "Chapitres 'Frais de structure' (exclus des dépenses)",
            chap_dep_all,
            help="Ex : charges de personnel 012, charges financières 66…",
        )

    def cout(col: str, excl_dep: set[str], excl_rec: set[str] | None = None) -> float:
        d = sum_hors(dep_f, col, ORDRE_DEP_F | excl_dep)
        if excl_rec:
            d -= rec_f[rec_f["Chapitre"].isin(excl_rec)][col].sum()
        return d

    categories = [
        ("Coût de fonctionnement",
         set(), set()),
        ("Hors travaux",
         set(chap_travaux), set()),
        ("Hors travaux et vente eau",
         set(chap_travaux), set(chap_vente_eau)),
        ("Hors travaux, frais struct. et vente eau",
         set(chap_travaux) | set(chap_frais_struct), set(chap_vente_eau)),
    ]

    rows = []
    for label, excl_dep, excl_rec in categories:
        vn  = cout("CA N",   excl_dep, excl_rec)
        vn1 = cout("CA N-1", excl_dep, excl_rec)
        rows.append({
            "label":  label,
            "vn":     vn,
            "vn1":    vn1,
            "hab_n":  vn  / pop if pop else None,
            "hab_n1": vn1 / pop if pop else None,
            "evol":   (vn - vn1) / abs(vn1) if vn1 != 0 else None,
        })

    for row in rows:
        c1, c2, c3, c4, c5 = st.columns([2.5, 1, 1, 1, 1])
        c1.markdown(f"**{row['label']}**")
        c2.metric("2025", fmt(row["vn"]))
        c3.metric("2024", fmt(row["vn1"]))
        c4.metric("Par hab. 2025", fmt_hab(row["hab_n"]))
        c5.metric("Évolution", "",
                  delta=f"{row['evol']:.2%}" if row["evol"] is not None else None)
        st.divider()

    # Graphique évolution
    df_evol = pd.DataFrame([{
        "Métrique":     r["label"],
        "Évolution (%)": (r["evol"] or 0) * 100,
    } for r in rows])

    fig = px.bar(
        df_evol,
        x="Évolution (%)",
        y="Métrique",
        orientation="h",
        text_auto=".2f",
        color="Évolution (%)",
        color_continuous_scale=["#5cb85c", "#f0ad4e", "#d9534f"],
        title="Évolution du Coût de Fonctionnement (%)",
    )
    fig.update_layout(template="plotly_white", coloraxis_showscale=False, yaxis_title="")
    fig.update_traces(texttemplate="%{x:.2f}%")
    st.plotly_chart(fig, use_container_width=True)


# ─────────────────────────────────────────────────────────────────────────────
# Onglet Grand Livre (vue exécution budgétaire)
# ─────────────────────────────────────────────────────────────────────────────

def tab_grand_livre(df: pd.DataFrame) -> None:
    c1, c2 = st.columns(2)
    sel_sens     = c1.multiselect("Sens",    sorted(df["Groupe Sens"].dropna().unique()))
    sel_sections = c2.multiselect("Section", sorted(df["Section"].dropna().unique()))

    filtered = df.copy()
    if sel_sens:
        filtered = filtered[filtered["Groupe Sens"].isin(sel_sens)]
    if sel_sections:
        filtered = filtered[filtered["Section"].isin(sel_sections)]

    if filtered.empty:
        st.warning("Aucune donnée après application des filtres.")
        return

    ca_n   = filtered["CA N"].sum()
    budget = filtered["Budget voté"].sum()
    ca_n_1 = filtered["CA N-1"].sum()
    engage = filtered["Engagé HT"].sum()
    exec_r = ca_n / budget if budget else None
    var    = ca_n - ca_n_1
    var_p  = var / ca_n_1 if ca_n_1 else None
    reste  = budget - ca_n

    st.subheader("Indicateurs clés")
    m1, m2, m3, m4, m5, m6 = st.columns(6)
    m1.metric("CA N réalisé",     fmt(ca_n))
    m2.metric("Budget voté",      fmt(budget))
    m3.metric("Taux d'exécution", f"{exec_r:.1%}" if exec_r is not None else "N/A")
    m4.metric("Variation vs N-1", fmt(var),
              delta=f"{var_p:.1%}" if var_p is not None else None)
    m5.metric("Engagé HT",        fmt(engage))
    m6.metric("Reste à réaliser", fmt(reste))

    if exec_r is not None:
        fig_g = go.Figure(go.Indicator(
            mode="gauge+number",
            value=exec_r * 100,
            number={"suffix": "%", "valueformat": ".1f"},
            gauge={
                "axis": {"range": [0, 120], "ticksuffix": "%"},
                "bar":  {"color": "#2563eb"},
                "steps": [
                    {"range": [0, 70],  "color": "#fee2e2"},
                    {"range": [70, 90], "color": "#fef9c3"},
                    {"range": [90, 120],"color": "#dcfce7"},
                ],
                "threshold": {
                    "line": {"color": "#dc2626", "width": 3},
                    "thickness": 0.75, "value": 100,
                },
            },
            title={"text": "Taux d'exécution global"},
        ))
        fig_g.update_layout(height=250, margin=dict(t=50, b=10, l=30, r=30))
        st.plotly_chart(fig_g, use_container_width=True)

    group_cols = ["Chapitre"]
    if "Libellé chapitre" in filtered.columns:
        group_cols.append("Libellé chapitre")

    by_chap = (
        filtered.groupby(group_cols, dropna=False, as_index=False)[
            ["CA N", "Budget voté", "CA N-1", "Engagé HT"]
        ]
        .sum()
        .sort_values("Chapitre")
    )
    by_chap["Taux d'exécution"] = (by_chap["CA N"] / by_chap["Budget voté"]).where(
        by_chap["Budget voté"] != 0
    )
    by_chap["Reste à réaliser"] = by_chap["Budget voté"] - by_chap["CA N"]

    st.subheader("Tableau récapitulatif par chapitre")
    tbl = by_chap.copy()
    tbl["Taux d'exécution"] = tbl["Taux d'exécution"].apply(
        lambda v: f"{v * 100:.1f}%" if pd.notna(v) else "N/A"
    )
    col_cfg = {
        "CA N":             st.column_config.NumberColumn("CA N (€)",             format="%.0f"),
        "Budget voté":      st.column_config.NumberColumn("Budget voté (€)",      format="%.0f"),
        "CA N-1":           st.column_config.NumberColumn("CA N-1 (€)",           format="%.0f"),
        "Engagé HT":        st.column_config.NumberColumn("Engagé HT (€)",        format="%.0f"),
        "Reste à réaliser": st.column_config.NumberColumn("Reste à réaliser (€)", format="%.0f"),
        "Taux d'exécution": st.column_config.TextColumn("Taux exec."),
    }
    st.dataframe(tbl, use_container_width=True, hide_index=True, column_config=col_cfg)

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        tbl.to_excel(writer, index=False, sheet_name="Chapitres")
        filtered.to_excel(writer, index=False, sheet_name="Détail")
    st.download_button(
        "⬇ Télécharger l'extrait filtré (Excel)",
        data=buf.getvalue(),
        file_name="reporting_filtre.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    if "Article" in filtered.columns:
        with st.expander("🔍 Détail par article"):
            art_cols = ["Chapitre", "Article"]
            if "Libellé article" in filtered.columns:
                art_cols.append("Libellé article")
            by_art = (
                filtered.groupby(art_cols, dropna=False, as_index=False)[
                    ["CA N", "Budget voté", "CA N-1", "Engagé HT"]
                ]
                .sum()
                .sort_values(["Chapitre", "Article"])
            )
            by_art["Taux d'exécution"] = (
                by_art["CA N"] / by_art["Budget voté"]
            ).where(by_art["Budget voté"] != 0).apply(
                lambda v: f"{v * 100:.1f}%" if pd.notna(v) else "N/A"
            )
            st.dataframe(by_art, use_container_width=True, hide_index=True)

    st.subheader("Comparatif CA N vs CA N-1 par chapitre")
    fig_cmp = px.bar(
        by_chap.melt("Chapitre", ["CA N", "CA N-1"], var_name="Série", value_name="Montant"),
        x="Chapitre", y="Montant", color="Série", barmode="group",
        color_discrete_map={"CA N": "#2563eb", "CA N-1": "#93c5fd"},
        labels={"Montant": "Montant (€)"}, text_auto=".3s",
    )
    fig_cmp.update_layout(legend_title_text="", template="plotly_white")
    st.plotly_chart(fig_cmp, use_container_width=True)

    exec_df = by_chap.copy()
    exec_df["Couleur"] = exec_df["Taux d'exécution"].apply(
        lambda v: "#d9534f" if pd.isna(v) or v < 0.7 else ("#f0ad4e" if v < 0.9 else "#5cb85c")
    )
    st.subheader("Taux d'exécution par chapitre")
    fig_exec = px.bar(
        exec_df, x="Chapitre", y="Taux d'exécution",
        color="Couleur", color_discrete_map="identity",
    )
    fig_exec.add_hline(y=1.0, line_dash="dash", line_color="#dc2626", annotation_text="100 %")
    fig_exec.update_traces(texttemplate="%{y:.1%}", textposition="outside", showlegend=False)
    fig_exec.update_layout(yaxis_tickformat=".0%", template="plotly_white", showlegend=False)
    st.plotly_chart(fig_exec, use_container_width=True)

    st.subheader("Budget, Réalisé et Engagé par chapitre")
    fig_tri = px.bar(
        by_chap.melt("Chapitre", ["Budget voté", "CA N", "Engagé HT"],
                     var_name="Série", value_name="Montant"),
        x="Chapitre", y="Montant", color="Série", barmode="group",
        color_discrete_map={"Budget voté": "#cbd5e1", "CA N": "#2563eb", "Engagé HT": "#f59e0b"},
        labels={"Montant": "Montant (€)"}, text_auto=".3s",
    )
    fig_tri.update_layout(legend_title_text="", template="plotly_white")
    st.plotly_chart(fig_tri, use_container_width=True)


# ─────────────────────────────────────────────────────────────────────────────
# Application principale
# ─────────────────────────────────────────────────────────────────────────────

st.title("Reporting financier M49")
st.write("Chargez votre grand livre annuel pour générer les indicateurs et graphiques de pilotage.")

uploaded_file = st.file_uploader("Uploader le fichier Excel du grand livre", type=["xlsx", "xls"])

if uploaded_file is None:
    st.info("Aucun fichier chargé. Importez un fichier Excel pour afficher le reporting.")
else:
    try:
        data = load_excel(uploaded_file.read())
        st.success(f"Fichier chargé : {len(data):,} lignes exploitables.".replace(",", "\u202f"))

        # Filtre entité dans la sidebar (commun à tous les onglets)
        st.sidebar.header("Filtre — Entité")
        all_codes = sorted(data["CODE"].dropna().unique())
        selected_codes = st.sidebar.multiselect("CODE", all_codes, default=all_codes)
        if not selected_codes:
            selected_codes = all_codes

        filtered = data[data["CODE"].isin(selected_codes)]
        label    = ", ".join(selected_codes)

        tab_gl, tab_fct, tab_inv, tab_cout = st.tabs([
            "📋 Grand Livre",
            "🏛️ Fonctionnement",
            "🏗️ Investissement",
            "💰 Coût de Fonctionnement",
        ])

        with tab_gl:
            st.subheader(f"Grand Livre — {label}")
            tab_grand_livre(filtered)

        with tab_fct:
            st.subheader(f"Fonctionnement — {label}")
            tab_fonctionnement(filtered, selected_codes)

        with tab_inv:
            st.subheader(f"Investissement — {label}")
            tab_investissement(filtered, selected_codes)

        with tab_cout:
            st.subheader(f"Coût de Fonctionnement — {label}")
            tab_cout_fonctionnement(filtered, selected_codes)

    except Exception as exc:
        st.error(f"Impossible de traiter le fichier : {exc}")

