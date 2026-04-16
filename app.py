import unicodedata

import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="Reporting financier", page_icon="📊", layout="wide")

st.markdown(
    """
    <style>
    .block-container {padding-top: 1.2rem;}
    .stMetric {background-color: #f7f9fc; border: 1px solid #e8edf5; border-radius: 10px; padding: 0.6rem;}
    </style>
    """,
    unsafe_allow_html=True,
)


def normalize(value: str) -> str:
    text = unicodedata.normalize("NFKD", str(value)).encode("ascii", "ignore").decode("ascii")
    return " ".join(text.lower().strip().split())


def resolve_columns(df: pd.DataFrame) -> dict[str, str]:
    normalized = {normalize(col): col for col in df.columns}
    expected = {
        "code": ["code"],
        "groupe_sens": ["groupe sens", "sens"],
        "section": ["section"],
        "ca_n": ["ca 2025", "ca n", "realise", "realise ca n"],
        "budget": ["mt vote cp", "montant vote cp", "budget vote"],
        "ca_n_1": ["ca 2024", "ca n-1", "ca n- 1", "ca n 1"],
        "engage": ["mt engage ht", "montant engage ht", "engage ht"],
        "chapitre": ["chapitre", "chapter", "chap"],
    }

    resolved = {}
    for key, aliases in expected.items():
        for alias in aliases:
            if alias in normalized:
                resolved[key] = normalized[alias]
                break

    missing = [k for k in ("code", "groupe_sens", "section", "ca_n", "budget", "ca_n_1", "engage", "chapitre") if k not in resolved]
    if missing:
        raise KeyError(
            "Colonnes manquantes ou non reconnues : "
            + ", ".join(missing)
            + ". Vérifiez le format du grand livre."
        )

    return resolved


@st.cache_data(show_spinner=False)
def load_excel(file) -> pd.DataFrame:
    df = pd.read_excel(file)
    cols = resolve_columns(df)

    parsed = pd.DataFrame(
        {
            "CODE": df[cols["code"]].astype(str).str.strip(),
            "Groupe Sens": df[cols["groupe_sens"]].astype(str).str.strip(),
            "Section": df[cols["section"]].astype(str).str.strip(),
            "Chapitre": df[cols["chapitre"]].astype(str).str.strip(),
            "CA N": pd.to_numeric(df[cols["ca_n"]], errors="coerce").fillna(0.0),
            "Budget voté": pd.to_numeric(df[cols["budget"]], errors="coerce").fillna(0.0),
            "CA N-1": pd.to_numeric(df[cols["ca_n_1"]], errors="coerce").fillna(0.0),
            "Engagé HT": pd.to_numeric(df[cols["engage"]], errors="coerce").fillna(0.0),
        }
    )

    parsed = parsed[parsed["Chapitre"].ne("")]
    return parsed


def format_currency(value: float) -> str:
    return f"{value:,.0f} €".replace(",", " ")


def build_dashboard(df: pd.DataFrame) -> None:
    st.subheader("Filtres")
    c1, c2, c3 = st.columns(3)

    selected_codes = c1.multiselect("Entité (CODE)", sorted(df["CODE"].dropna().unique()))
    selected_sens = c2.multiselect("Sens", sorted(df["Groupe Sens"].dropna().unique()))
    selected_sections = c3.multiselect("Section", sorted(df["Section"].dropna().unique()))

    filtered = df.copy()
    if selected_codes:
        filtered = filtered[filtered["CODE"].isin(selected_codes)]
    if selected_sens:
        filtered = filtered[filtered["Groupe Sens"].isin(selected_sens)]
    if selected_sections:
        filtered = filtered[filtered["Section"].isin(selected_sections)]

    if filtered.empty:
        st.warning("Aucune donnée après application des filtres.")
        return

    ca_n = filtered["CA N"].sum()
    budget = filtered["Budget voté"].sum()
    ca_n_1 = filtered["CA N-1"].sum()
    exec_rate = (ca_n / budget) if budget else 0.0
    variation = ca_n - ca_n_1
    variation_pct = (variation / ca_n_1) if ca_n_1 else 0.0

    st.subheader("Indicateurs clés")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total réalisé CA N", format_currency(ca_n))
    m2.metric("Total budget voté", format_currency(budget))
    m3.metric("Taux d'exécution", f"{exec_rate:.1%}")
    m4.metric("Variation vs CA N-1", format_currency(variation), delta=f"{variation_pct:.1%}")

    by_chapter = (
        filtered.groupby("Chapitre", dropna=False, as_index=False)[["CA N", "Budget voté", "CA N-1", "Engagé HT"]]
        .sum()
        .sort_values("Chapitre")
    )

    by_chapter["Taux d'exécution"] = by_chapter.apply(
        lambda row: (row["CA N"] / row["Budget voté"]) if row["Budget voté"] else 0.0,
        axis=1,
    )

    st.subheader("Tableau récapitulatif par chapitre")
    table_data = by_chapter.copy()
    table_data["Taux d'exécution"] = table_data["Taux d'exécution"] * 100

    st.dataframe(
        table_data,
        use_container_width=True,
        column_config={
            "CA N": st.column_config.NumberColumn(format="%.0f"),
            "Budget voté": st.column_config.NumberColumn(format="%.0f"),
            "CA N-1": st.column_config.NumberColumn(format="%.0f"),
            "Engagé HT": st.column_config.NumberColumn(format="%.0f"),
            "Taux d'exécution": st.column_config.NumberColumn(format="%.1f%%"),
        },
    )

    st.subheader("Comparatif CA N vs CA N-1 par chapitre")
    chart_data = by_chapter.melt(
        id_vars="Chapitre",
        value_vars=["CA N", "CA N-1"],
        var_name="Série",
        value_name="Montant",
    )
    fig_compare = px.bar(
        chart_data,
        x="Chapitre",
        y="Montant",
        color="Série",
        barmode="group",
        labels={"Montant": "Montant (€)"},
    )
    fig_compare.update_layout(legend_title_text="", template="plotly_white")
    st.plotly_chart(fig_compare, use_container_width=True)

    st.subheader("Taux d'exécution par chapitre")
    fig_exec = px.bar(
        by_chapter,
        x="Chapitre",
        y="Taux d'exécution",
        labels={"Taux d'exécution": "Taux d'exécution"},
    )
    fig_exec.update_traces(texttemplate="%{y:.1%}", textposition="outside")
    fig_exec.update_layout(yaxis_tickformat=".0%", template="plotly_white")
    st.plotly_chart(fig_exec, use_container_width=True)


st.title("Reporting financier — Phase 1")
st.write("Chargez votre grand livre annuel pour générer les indicateurs et graphiques de pilotage.")

uploaded_file = st.file_uploader("Uploader le fichier Excel du grand livre", type=["xlsx", "xls"])

if uploaded_file is None:
    st.info("Aucun fichier chargé. Importez un fichier Excel pour afficher le reporting.")
else:
    try:
        data = load_excel(uploaded_file)
        st.success(f"Fichier chargé : {len(data):,} lignes exploitables.".replace(",", " "))
        build_dashboard(data)
    except Exception as exc:
        st.error(f"Impossible de traiter le fichier : {exc}")
