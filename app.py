# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries
import unicodedata
import re
import altair as alt
from io import BytesIO

# -----------------------------------------------
# Estilos APC Colombia
# -----------------------------------------------
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600;700;800&family=Open+Sans:wght@400;600&display=swap');

:root {
    --apc-blue:   #003087;
    --apc-green:  #4CAF50;
    --apc-light:  #E8F0FE;
    --apc-gray:   #F5F7FA;
    --apc-border: #D0D9EA;
    --apc-text:   #1A1A2E;
    --apc-muted:  #6B7280;
    --apc-white:  #FFFFFF;
    --apc-accent: #1565C0;
}

html, body, [class*="css"] {
    font-family: 'Open Sans', sans-serif;
    color: var(--apc-text);
}

.apc-header {
    background: linear-gradient(135deg, #003087 0%, #1565C0 60%, #1976D2 100%);
    padding: 1.5rem 2rem 1.2rem 2rem;
    border-radius: 0 0 16px 16px;
    margin-bottom: 1.5rem;
    display: flex;
    align-items: center;
    box-shadow: 0 4px 20px rgba(0,48,135,0.25);
}
.apc-header-title {
    color: white;
    font-family: 'Montserrat', sans-serif;
    font-size: 1.7rem;
    font-weight: 800;
    letter-spacing: -0.5px;
    margin: 0;
    text-shadow: 0 2px 4px rgba(0,0,0,0.15);
}
.apc-header-subtitle {
    color: rgba(255,255,255,0.85);
    font-size: 0.85rem;
    font-weight: 400;
    margin-top: 2px;
    font-family: 'Open Sans', sans-serif;
}

.section-header {
    font-family: 'Montserrat', sans-serif;
    font-weight: 700;
    font-size: 1.05rem;
    color: var(--apc-blue);
    border-bottom: 2px solid var(--apc-blue);
    padding-bottom: 6px;
    margin: 1.5rem 0 1rem 0;
    letter-spacing: 0.2px;
}

.dept-title-banner {
    background: linear-gradient(90deg, var(--apc-blue) 0%, var(--apc-accent) 100%);
    color: white;
    font-family: 'Montserrat', sans-serif;
    font-size: 1.35rem;
    font-weight: 800;
    padding: 0.7rem 1.5rem;
    border-radius: 10px;
    margin-bottom: 1rem;
    letter-spacing: 0.3px;
    box-shadow: 0 2px 10px rgba(0,48,135,0.18);
}

div[data-testid="stMetric"] {
    background: var(--apc-white);
    border: 1.5px solid var(--apc-border);
    border-top: 4px solid var(--apc-blue);
    border-radius: 10px;
    padding: 1rem !important;
    box-shadow: 0 2px 8px rgba(0,48,135,0.07);
    transition: box-shadow 0.2s;
}
div[data-testid="stMetric"]:hover {
    box-shadow: 0 4px 16px rgba(0,48,135,0.14);
}
div[data-testid="stMetricLabel"] {
    font-family: 'Montserrat', sans-serif;
    font-weight: 600;
    font-size: 0.78rem;
    color: var(--apc-muted) !important;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}
div[data-testid="stMetricValue"] {
    font-family: 'Montserrat', sans-serif;
    font-weight: 700;
    color: var(--apc-blue) !important;
}

button[data-baseweb="tab"] {
    font-family: 'Montserrat', sans-serif !important;
    font-weight: 600 !important;
    font-size: 0.88rem !important;
    letter-spacing: 0.3px;
}
div[data-baseweb="tab-highlight"] {
    background-color: var(--apc-blue) !important;
}
div[data-baseweb="tab-border"] {
    background-color: var(--apc-border) !important;
}

div[data-testid="stDataFrame"] {
    border-radius: 8px;
    overflow: hidden;
    border: 1px solid var(--apc-border);
}

div[data-testid="stDownloadButton"] button {
    background: var(--apc-blue) !important;
    color: white !important;
    border: none !important;
    border-radius: 8px !important;
    font-family: 'Montserrat', sans-serif !important;
    font-weight: 600 !important;
    padding: 0.5rem 1.3rem !important;
    transition: background 0.2s !important;
}
div[data-testid="stDownloadButton"] button:hover {
    background: var(--apc-accent) !important;
}

.guia-card {
    background: var(--apc-white);
    border: 1px solid var(--apc-border);
    border-radius: 12px;
    padding: 1.8rem 2.2rem;
    margin-bottom: 1rem;
    line-height: 1.75;
    font-size: 0.96rem;
    box-shadow: 0 2px 8px rgba(0,48,135,0.06);
}
.guia-card p {
    color: var(--apc-text);
    margin-bottom: 1rem;
}
.guia-intro {
    font-family: 'Montserrat', sans-serif;
    font-size: 1rem;
    font-weight: 600;
    color: var(--apc-blue);
    background: var(--apc-light);
    border-left: 4px solid var(--apc-blue);
    border-radius: 0 8px 8px 0;
    padding: 0.8rem 1.2rem;
    margin-bottom: 1.2rem;
}

.apc-footer {
    text-align: center;
    color: var(--apc-muted);
    font-size: 0.78rem;
    margin-top: 2.5rem;
    padding-top: 1rem;
    border-top: 1px solid var(--apc-border);
}

/* Ocultar barra superior de Streamlit */
div[data-testid="stToolbar"],
div[data-testid="stDecoration"],
header[data-testid="stHeader"] {
    display: none !important;
}
</style>
""", unsafe_allow_html=True)

FILE = "Ficha territorial.xlsm"


def norm_text(x):
    if x is None:
        return ""
    s = str(x).strip().upper()
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
    s = re.sub(r"[^A-Z0-9\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def format_usd(n):
    try:
        n = float(n)
    except Exception:
        return ""
    return "USD " + f"{n:,.0f}".replace(",", ".")


def format_cop(n):
    try:
        n = float(n)
    except Exception:
        return ""
    return "$ " + f"{n:,.0f}".replace(",", ".")


def get_col(row, *names):
    """Busca una columna por varios nombres posibles (con y sin tilde)."""
    for name in names:
        val = row.get(name, None)
        if val is not None and str(val).strip() not in ("", "None", "nan"):
            return val
    return ""


@st.cache_data
def read_named_table(file_path: str, table_name: str) -> pd.DataFrame:
    wb = load_workbook(file_path, data_only=True, keep_vba=True)
    for ws in wb.worksheets:
        if table_name in ws.tables:
            tbl = ws.tables[table_name]
            min_col, min_row, max_col, max_row = range_boundaries(tbl.ref)
            data = []
            for row in ws.iter_rows(
                min_row=min_row, max_row=max_row,
                min_col=min_col, max_col=max_col,
                values_only=True
            ):
                data.append(list(row))
            header = data[0]
            rows = data[1:]
            return pd.DataFrame(rows, columns=header)
    raise KeyError(f"No encontre la tabla: {table_name}")


@st.cache_data
def load_data():
    infogeneral = read_named_table(FILE, "infogeneral")
    plan = read_named_table(FILE, "plan")
    ciclope = read_named_table(FILE, "ciclope")
    colcol = read_named_table(FILE, "colcol")
    contrapartidas = read_named_table(FILE, "contrapartidas")
    proyectos = read_named_table(FILE, "proyectos")

    for df in [infogeneral, plan, ciclope, colcol, contrapartidas, proyectos]:
        for c in df.columns:
            if df[c].dtype == "object":
                df[c] = df[c].astype(str).str.strip()

    if "VALOR APORTE (USD)" in ciclope.columns:
        ciclope["VALOR APORTE (USD)"] = pd.to_numeric(
            ciclope["VALOR APORTE (USD)"], errors="coerce"
        ).fillna(0)

    return infogeneral, plan, ciclope, colcol, contrapartidas, proyectos


def top_by_sum(df, group_col, value_col, n=5):
    if df.empty or group_col not in df.columns or value_col not in df.columns:
        return pd.DataFrame(columns=[group_col, value_col])
    return (
        df.groupby(group_col, dropna=False)[value_col]
        .sum()
        .sort_values(ascending=False)
        .head(n)
        .reset_index()
    )



ODS_NOMBRES = {
    "ODS 1":  "ODS 1 - Fin de la pobreza",
    "ODS 2":  "ODS 2 - Hambre cero",
    "ODS 3":  "ODS 3 - Salud y bienestar",
    "ODS 4":  "ODS 4 - Educaci\u00f3n de calidad",
    "ODS 5":  "ODS 5 - Igualdad de g\u00e9nero",
    "ODS 6":  "ODS 6 - Agua limpia y saneamiento",
    "ODS 7":  "ODS 7 - Energ\u00eda asequible y no contaminante",
    "ODS 8":  "ODS 8 - Trabajo decente y crecimiento econ\u00f3mico",
    "ODS 9":  "ODS 9 - Industria, innovaci\u00f3n e infraestructura",
    "ODS 10": "ODS 10 - Reducci\u00f3n de las desigualdades",
    "ODS 11": "ODS 11 - Ciudades y comunidades sostenibles",
    "ODS 12": "ODS 12 - Producci\u00f3n y consumo responsables",
    "ODS 13": "ODS 13 - Acci\u00f3n por el clima",
    "ODS 14": "ODS 14 - Vida submarina",
    "ODS 15": "ODS 15 - Vida de ecosistemas terrestres",
    "ODS 16": "ODS 16 - Paz, justicia e instituciones s\u00f3lidas",
    "ODS 17": "ODS 17 - Alianzas para lograr los objetivos",
}

def to_excel_ficha(info_row, cic_dept, colcol_dept, contr_dept):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        if not info_row.empty:
            df_info = info_row.T.reset_index()
            df_info.columns = ["Campo", "Valor"]
            df_info = df_info[
                ~df_info["Campo"].astype(str).str.lower().str.strip().isin(
                    ["porcentaje de avance", "dept_norm"]
                )
            ]
            df_info.to_excel(writer, sheet_name="Informacion General", index=False)
        cic_export = cic_dept.drop(columns=["DEPT_NORM"], errors="ignore")
        cic_export.to_excel(writer, sheet_name="AOD - Ciclope", index=False)
        colcol_dept.copy().to_excel(writer, sheet_name="ColCol", index=False)
        contr_dept.to_excel(writer, sheet_name="Contrapartidas", index=False)
    output.seek(0)
    return output.getvalue()


def to_excel_proyectos(df_proj):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_proj.to_excel(writer, sheet_name="Proyectos AOD", index=False)
    output.seek(0)
    return output.getvalue()


# -------------------------
# APP
# -------------------------
st.set_page_config(page_title="Ficha Territorial | APC Colombia", layout="wide")

# Header institucional
st.markdown("""
<div class="apc-header">
    <div>
        <div class="apc-header-title">Ficha Territorial</div>
        <div class="apc-header-subtitle">Caracterizaci\u00f3n por departamento</div>
    </div>
</div>
""", unsafe_allow_html=True)

infogeneral, plan, ciclope, colcol, contrapartidas, proyectos = load_data()

DEPT_COL_INFO = "Departamento"
depts = sorted(infogeneral[DEPT_COL_INFO].dropna().unique().tolist())

dept = st.selectbox("Selecciona un departamento", depts)

dept_norm = norm_text(dept)
infogeneral["DEPT_NORM"] = infogeneral[DEPT_COL_INFO].map(norm_text)
info = infogeneral[infogeneral["DEPT_NORM"] == dept_norm].head(1)

ciclope["DEPT_NORM"] = ciclope["DEPARTAMENTO"].map(norm_text)
proyectos["DEPT_NORM"] = proyectos["DEPARTAMENTO"].map(norm_text)

cic_dept = ciclope[ciclope["DEPT_NORM"] == dept_norm]
proj_dept = proyectos[proyectos["DEPT_NORM"] == dept_norm]

mask_colcol = pd.Series(False, index=colcol.index)
if "DEPARTAMENTOS PARTICIPANTES" in colcol.columns:
    mask_colcol = (
        colcol["DEPARTAMENTOS PARTICIPANTES"]
        .astype("string")
        .map(norm_text)
        .str.contains(dept_norm, na=False)
    )
else:
    st.warning("No encontre la columna 'DEPARTAMENTOS PARTICIPANTES' en ColCol.")

colcol_dept = colcol[mask_colcol]

if "Departamento" in contrapartidas.columns:
    contr_dept = contrapartidas[
        contrapartidas["Departamento"].astype("string").map(norm_text) == dept_norm
    ]
else:
    contr_dept = contrapartidas.iloc[0:0]


tab1, tab2, tab3, tab4 = st.tabs([
    "\U0001f4cb Ficha territorial",
    "\U0001f5c2\ufe0f Proyectos AOD",
    "\U0001f310 Panorama Nacional",
    "\U0001f4d6 Gu\u00eda de usuario"
])


# =========================================================
# TAB 1
# =========================================================
with tab1:

    st.markdown(
        f'<div class="dept-title-banner">\U0001f4cd {dept}</div>',
        unsafe_allow_html=True
    )

    st.markdown('<div class="section-header">Informaci\u00f3n General</div>', unsafe_allow_html=True)

    if info.empty:
        st.warning("No encontre el departamento en la tabla infogeneral.")
    else:
        c1, c2, c3 = st.columns(3)

        c1.metric("Capital", get_col(info.iloc[0], "Capital"))
        c2.metric(
            "N\u00famero de Municipios",
            get_col(info.iloc[0], "N\u00famero de Municipios", "Numero de Municipios", "Municipios")
        )

        pob_raw = get_col(info.iloc[0], "Poblaci\u00f3n", "Poblacion", "Poblaci\u00f3n")
        try:
            pob_fmt = f"{int(float(pob_raw)):,}".replace(",", ".")
        except Exception:
            pob_fmt = str(pob_raw)
        c3.metric("Poblaci\u00f3n", pob_fmt)

        with st.expander("Ver registro completo del departamento"):
            df_det = info.T.reset_index()
            df_det.columns = ["Campo", "Valor"]
            df_det = df_det[
                ~df_det["Campo"].astype(str).str.lower().str.strip().isin(
                    ["porcentaje de avance", "dept_norm"]
                )
            ]
            st.dataframe(df_det, use_container_width=True, hide_index=True)

    st.markdown(
        '<div class="section-header">Ayuda Oficial al Desarrollo (AOD)</div>',
        unsafe_allow_html=True
    )

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Intervenciones (unicas)",
              cic_dept["CODIGO INTERVENCION"].nunique()
              if "CODIGO INTERVENCION" in cic_dept.columns else 0)
    m2.metric("Cooperantes",
              cic_dept["NOMBRE ACTOR"].nunique()
              if "NOMBRE ACTOR" in cic_dept.columns else 0)
    municipios_count = (
        cic_dept["MUNICIPIO"].map(norm_text)
        .pipe(lambda s: s[~s.isin(["NO REPORTA", "SIN INFORMACION", "NO APLICA", ""])])
        .nunique()
        if "MUNICIPIO" in cic_dept.columns else 0
    )
    m3.metric("Municipios o \u00e1reas intervenidas", municipios_count)
    st.caption("\* Incluye municipios y \u00e1reas no municipalizadas")
    total_usd = cic_dept["VALOR APORTE (USD)"].sum() \
        if "VALOR APORTE (USD)" in cic_dept.columns else 0
    m4.metric("Total aporte estimado (USD)", format_usd(total_usd))

    c5, c6 = st.columns(2)

    with c5:
        st.markdown("**Top 5 cooperantes por USD**")
        top_act = top_by_sum(cic_dept, "NOMBRE ACTOR", "VALOR APORTE (USD)", 5)
        if not top_act.empty:
            chart_act = (
                alt.Chart(top_act)
                .mark_bar(color="#003087", cornerRadiusTopRight=4, cornerRadiusBottomRight=4)
                .encode(
                    y=alt.Y("NOMBRE ACTOR:N", sort="-x", title=""),
                    x=alt.X("VALOR APORTE (USD):Q", title="USD"),
                    tooltip=["NOMBRE ACTOR:N",
                             alt.Tooltip("VALOR APORTE (USD):Q", format=",.0f")]
                )
                .properties(height=200)
            )
            st.altair_chart(chart_act, use_container_width=True)
            top_act_disp = top_act.copy()
            top_act_disp["VALOR APORTE (USD)"] = top_act_disp["VALOR APORTE (USD)"].apply(format_usd)
            st.dataframe(top_act_disp, use_container_width=True, hide_index=True)
        else:
            st.info("Sin datos suficientes para cooperantes.")

    with c6:
        st.markdown("**Top 5 ODS por USD**")
        top_ods = top_by_sum(cic_dept, "ODS", "VALOR APORTE (USD)", 5)
        if not top_ods.empty:
            chart_ods = (
                alt.Chart(top_ods)
                .mark_bar(color="#1565C0", cornerRadiusTopRight=4, cornerRadiusBottomRight=4)
                .encode(
                    y=alt.Y("ODS:N", sort="-x", title=""),
                    x=alt.X("VALOR APORTE (USD):Q", title="USD"),
                    tooltip=["ODS:N",
                             alt.Tooltip("VALOR APORTE (USD):Q", format=",.0f")]
                )
                .properties(height=200)
            )
            st.altair_chart(chart_ods, use_container_width=True)
            top_ods_disp = top_ods.copy()
            top_ods_disp["VALOR APORTE (USD)"] = top_ods_disp["VALOR APORTE (USD)"].apply(format_usd)
            top_ods_disp["ODS"] = top_ods_disp["ODS"].map(lambda x: ODS_NOMBRES.get(x, x))
            st.dataframe(top_ods_disp, use_container_width=True, hide_index=True)
        else:
            st.info("Sin datos suficientes para ODS.")

    st.markdown(
        '<div class="section-header">Programas Internos APC-Colombia</div>',
        unsafe_allow_html=True
    )

    p1, p2 = st.columns(2)
    with p1:
        st.markdown("**ColCol - Colombia Ense\u00f1a Colombia 2025-2026**")
        st.metric("Registros encontrados", len(colcol_dept))
        colcol_view = colcol_dept.copy()
        if "PRESUPUESTO ESTIMADO APC COLOMBIA" in colcol_view.columns:
            colcol_view["PRESUPUESTO ESTIMADO APC COLOMBIA"] = (
                pd.to_numeric(colcol_view["PRESUPUESTO ESTIMADO APC COLOMBIA"], errors="coerce")
                .apply(format_cop)
            )
        st.dataframe(colcol_view.head(50), use_container_width=True, hide_index=True)

    with p2:
        st.markdown("**Contrapartidas 2025-2026**")
        st.metric("Registros encontrados", len(contr_dept))
        st.dataframe(contr_dept.head(50), use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("**Descargar ficha territorial completa**")
    excel_ficha = to_excel_ficha(info, cic_dept, colcol_dept, contr_dept)
    st.download_button(
        label="Descargar Ficha Territorial (Excel)",
        data=excel_ficha,
        file_name=f"Ficha_Territorial_{dept}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.markdown(
        '<div class="apc-footer">Agencia Presidencial de Cooperacion Internacional de Colombia - APC-Colombia</div>',
        unsafe_allow_html=True
    )


# =========================================================
# TAB 2
# =========================================================
with tab2:

    st.markdown(
        f'<div class="dept-title-banner">\U0001f4cd {dept} \u2014 Proyectos AOD activos</div>',
        unsafe_allow_html=True
    )
    st.caption("Fuente: Ciclope a corte de 31 de diciembre de 2025")

    search = st.text_input("Buscar en proyectos").strip()
    df = proj_dept.copy()

    if search and not df.empty:
        candidate_cols = [
            "NOMBRE INTERVENCION", "OBJETIVO GENERAL",
            "NOMBRE ACTOR", "MUNICIPIO", "ODS", "META ODS"
        ]
        cols = [c for c in candidate_cols if c in df.columns]
        if cols:
            mask = False
            for c in cols:
                mask = mask | df[c].astype(str).str.contains(search, case=False, na=False)
            df = df[mask]

    df = df.drop(columns=["DEPT_NORM"], errors="ignore")
    st.dataframe(df, use_container_width=True, hide_index=True)

    col_dl1, col_dl2 = st.columns(2)
    with col_dl1:
        st.download_button(
            "Descargar CSV filtrado",
            data=df.to_csv(index=False).encode("utf-8"),
            file_name=f"proyectos_aod_{dept}.csv",
            mime="text/csv"
        )
    with col_dl2:
        excel_proj = to_excel_proyectos(df)
        st.download_button(
            label="Descargar Excel",
            data=excel_proj,
            file_name=f"Proyectos_AOD_{dept}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    st.markdown(
        '<div class="apc-footer">Agencia Presidencial de Cooperacion Internacional de Colombia - APC-Colombia</div>',
        unsafe_allow_html=True
    )



# =========================================================
# TAB 3 - PANORAMA NACIONAL
# =========================================================
with tab3:

    st.markdown(
        '<div class="dept-title-banner">\U0001f310 Panorama Nacional de la Cooperaci\u00f3n Internacional</div>',
        unsafe_allow_html=True
    )
    st.caption("Fuente: C\u00edclope a corte de 31 de diciembre de 2025. Incluye \u00e1mbito nacional y territorial.")

    # Calcular datos nacionales
    cic_nacional = ciclope.copy()
    cic_nacional["VALOR APORTE (USD)"] = pd.to_numeric(cic_nacional["VALOR APORTE (USD)"], errors="coerce").fillna(0)

    n1, n2, n3, n4 = st.columns(4)
    n1.metric("Intervenciones (\u00fanicas)", cic_nacional["CODIGO INTERVENCION"].nunique()
              if "CODIGO INTERVENCION" in cic_nacional.columns else 0)
    n2.metric("Cooperantes", cic_nacional["NOMBRE ACTOR"].nunique()
              if "NOMBRE ACTOR" in cic_nacional.columns else 0)
    n3.metric("Departamentos con AOD",
              cic_nacional[cic_nacional["DEPARTAMENTO"] != "\u00c1mbito Nacional"]["DEPARTAMENTO"].nunique()
              if "DEPARTAMENTO" in cic_nacional.columns else 0)
    total_nac = cic_nacional["VALOR APORTE (USD)"].sum()
    total_nac_fmt = "USD " + f"{total_nac/1_000_000:,.0f} M".replace(",", ".")
    n4.metric("Total aporte estimado (USD)", total_nac_fmt)

    st.markdown('<div class="section-header">Cooperantes</div>', unsafe_allow_html=True)
    c_n1, c_n2 = st.columns(2)

    with c_n1:
        st.markdown("**Top 10 cooperantes por recursos (USD)**")
        top_coop_usd = (
            cic_nacional.groupby("NOMBRE ACTOR")["VALOR APORTE (USD)"]
            .sum().sort_values(ascending=False).head(10).reset_index()
        )
        if not top_coop_usd.empty:
            chart_coop_usd = (
                alt.Chart(top_coop_usd)
                .mark_bar(color="#003087", cornerRadiusTopRight=4, cornerRadiusBottomRight=4)
                .encode(
                    y=alt.Y("NOMBRE ACTOR:N", sort="-x", title=""),
                    x=alt.X("VALOR APORTE (USD):Q", title="USD"),
                    tooltip=["NOMBRE ACTOR:N", alt.Tooltip("VALOR APORTE (USD):Q", format=",.0f")]
                )
                .properties(height=280)
            )
            st.altair_chart(chart_coop_usd, use_container_width=True)
            top_coop_usd_disp = top_coop_usd.copy()
            top_coop_usd_disp["VALOR APORTE (USD)"] = top_coop_usd_disp["VALOR APORTE (USD)"].apply(format_usd)
            st.dataframe(top_coop_usd_disp, use_container_width=True, hide_index=True)

    with c_n2:
        st.markdown("**Top 10 cooperantes por n\u00famero de intervenciones**")
        top_coop_int = (
            cic_nacional.groupby("NOMBRE ACTOR")["CODIGO INTERVENCION"]
            .nunique().sort_values(ascending=False).head(10).reset_index()
        )
        top_coop_int.columns = ["NOMBRE ACTOR", "INTERVENCIONES"]
        if not top_coop_int.empty:
            chart_coop_int = (
                alt.Chart(top_coop_int)
                .mark_bar(color="#1565C0", cornerRadiusTopRight=4, cornerRadiusBottomRight=4)
                .encode(
                    y=alt.Y("NOMBRE ACTOR:N", sort="-x", title=""),
                    x=alt.X("INTERVENCIONES:Q", title="N\u00famero de intervenciones"),
                    tooltip=["NOMBRE ACTOR:N", "INTERVENCIONES:Q"]
                )
                .properties(height=280)
            )
            st.altair_chart(chart_coop_int, use_container_width=True)
            st.dataframe(top_coop_int, use_container_width=True, hide_index=True)

    st.markdown('<div class="section-header">ODS y Sectores</div>', unsafe_allow_html=True)
    c_n3, c_n4 = st.columns(2)

    with c_n3:
        st.markdown("**Top 10 ODS por recursos (USD)**")
        top_ods_nac = (
            cic_nacional.groupby("ODS")["VALOR APORTE (USD)"]
            .sum().sort_values(ascending=False).head(10).reset_index()
        )
        if not top_ods_nac.empty:
            chart_ods_nac = (
                alt.Chart(top_ods_nac)
                .mark_bar(color="#003087", cornerRadiusTopRight=4, cornerRadiusBottomRight=4)
                .encode(
                    y=alt.Y("ODS:N", sort="-x", title=""),
                    x=alt.X("VALOR APORTE (USD):Q", title="USD"),
                    tooltip=["ODS:N", alt.Tooltip("VALOR APORTE (USD):Q", format=",.0f")]
                )
                .properties(height=280)
            )
            st.altair_chart(chart_ods_nac, use_container_width=True)
            top_ods_nac_disp = top_ods_nac.copy()
            top_ods_nac_disp["VALOR APORTE (USD)"] = top_ods_nac_disp["VALOR APORTE (USD)"].apply(format_usd)
            top_ods_nac_disp["ODS"] = top_ods_nac_disp["ODS"].map(lambda x: ODS_NOMBRES.get(x, x))
            st.dataframe(top_ods_nac_disp, use_container_width=True, hide_index=True)

    with c_n4:
        st.markdown("**Top 10 sectores por recursos (USD)**")
        top_sect_nac = (
            cic_nacional.groupby("SECTORES GOB")["VALOR APORTE (USD)"]
            .sum().sort_values(ascending=False).head(10).reset_index()
        )
        if not top_sect_nac.empty:
            chart_sect_nac = (
                alt.Chart(top_sect_nac)
                .mark_bar(color="#1565C0", cornerRadiusTopRight=4, cornerRadiusBottomRight=4)
                .encode(
                    y=alt.Y("SECTORES GOB:N", sort="-x", title=""),
                    x=alt.X("VALOR APORTE (USD):Q", title="USD"),
                    tooltip=["SECTORES GOB:N", alt.Tooltip("VALOR APORTE (USD):Q", format=",.0f")]
                )
                .properties(height=280)
            )
            st.altair_chart(chart_sect_nac, use_container_width=True)
            top_sect_nac_disp = top_sect_nac.copy()
            top_sect_nac_disp["VALOR APORTE (USD)"] = top_sect_nac_disp["VALOR APORTE (USD)"].apply(format_usd)
            st.dataframe(top_sect_nac_disp, use_container_width=True, hide_index=True)

    st.markdown('<div class="section-header">Departamentos</div>', unsafe_allow_html=True)
    c_n5, c_n6 = st.columns(2)

    with c_n5:
        st.markdown("**Top 10 departamentos por recursos (USD)**")
        top_dept_usd = (
            cic_nacional[cic_nacional["DEPARTAMENTO"] != "\u00c1mbito Nacional"]
            .groupby("DEPARTAMENTO")["VALOR APORTE (USD)"]
            .sum().sort_values(ascending=False).head(10).reset_index()
        )
        if not top_dept_usd.empty:
            chart_dept_usd = (
                alt.Chart(top_dept_usd)
                .mark_bar(color="#003087", cornerRadiusTopRight=4, cornerRadiusBottomRight=4)
                .encode(
                    y=alt.Y("DEPARTAMENTO:N", sort="-x", title=""),
                    x=alt.X("VALOR APORTE (USD):Q", title="USD"),
                    tooltip=["DEPARTAMENTO:N", alt.Tooltip("VALOR APORTE (USD):Q", format=",.0f")]
                )
                .properties(height=280)
            )
            st.altair_chart(chart_dept_usd, use_container_width=True)
            top_dept_usd_disp = top_dept_usd.copy()
            top_dept_usd_disp["VALOR APORTE (USD)"] = top_dept_usd_disp["VALOR APORTE (USD)"].apply(format_usd)
            st.dataframe(top_dept_usd_disp, use_container_width=True, hide_index=True)

    with c_n6:
        st.markdown("**Top 10 departamentos por intervenciones**")
        top_dept_int = (
            cic_nacional[cic_nacional["DEPARTAMENTO"] != "\u00c1mbito Nacional"]
            .groupby("DEPARTAMENTO")["CODIGO INTERVENCION"]
            .nunique().sort_values(ascending=False).head(10).reset_index()
        )
        top_dept_int.columns = ["DEPARTAMENTO", "INTERVENCIONES"]
        if not top_dept_int.empty:
            chart_dept_int = (
                alt.Chart(top_dept_int)
                .mark_bar(color="#1565C0", cornerRadiusTopRight=4, cornerRadiusBottomRight=4)
                .encode(
                    y=alt.Y("DEPARTAMENTO:N", sort="-x", title=""),
                    x=alt.X("INTERVENCIONES:Q", title="N\u00famero de intervenciones"),
                    tooltip=["DEPARTAMENTO:N", "INTERVENCIONES:Q"]
                )
                .properties(height=280)
            )
            st.altair_chart(chart_dept_int, use_container_width=True)
            st.dataframe(top_dept_int, use_container_width=True, hide_index=True)

    st.markdown(
        '<div class="apc-footer">Agencia Presidencial de Cooperacion Internacional de Colombia - APC-Colombia</div>',
        unsafe_allow_html=True
    )

# =========================================================
# TAB 4
# =========================================================
with tab4:

    st.markdown(
        '<div class="dept-title-banner">Gu\u00eda de usuario</div>',
        unsafe_allow_html=True
    )

    guia_html = (
        '<div class="guia-card">'
        '<div class="guia-intro">\u00bfQu\u00e9 es la Ficha Territorial?</div>'
        '<p>La ficha territorial es una aplicaci\u00f3n que permite conocer c\u00f3mo se est\u00e1 moviendo '
        'la cooperaci\u00f3n internacional en Colombia.</p>'
        '<p>En la primera secci\u00f3n encontrar\u00e1 <strong>informaci\u00f3n general del territorio</strong>. '
        'Podr\u00e1 identificar si el departamento cuenta con una dependencia encargada de '
        'cooperaci\u00f3n internacional, los enlaces o personas clave que participan en la gobernanza '
        'de la cooperaci\u00f3n en el territorio y si dispone de un plan de trabajo dentro del '
        'Sistema Nacional de Cooperaci\u00f3n Internacional, entre otros elementos de contexto '
        'institucional.</p>'
        '<p>Posteriormente, encontrar\u00e1 una secci\u00f3n relacionada con la '
        '<strong>Ayuda Oficial al Desarrollo (AOD)</strong>. Esta informaci\u00f3n proviene del '
        'sistema de informaci\u00f3n C\u00edclope, administrado por la Agencia Presidencial de '
        'Cooperaci\u00f3n Internacional de Colombia (APC-Colombia). En este apartado se presentan, '
        'entre otros aspectos, los principales cooperantes presentes en el territorio, los '
        'municipios que est\u00e1n siendo intervenidos y una estimaci\u00f3n de los recursos provenientes '
        'de cooperaci\u00f3n internacional. Asimismo, podr\u00e1 identificar los Objetivos de Desarrollo '
        'Sostenible (ODS) que concentran mayor financiaci\u00f3n en cada departamento.</p>'
        '<p>La ficha tambi\u00e9n muestra si el departamento participa en algunos de los '
        '<strong>programas de la oferta institucional de APC-Colombia</strong>. Entre ellos se '
        'encuentran la estrategia <em>Colombia Ense\u00f1a Colombia</em>, orientada a promover '
        'intercambios de conocimiento en diversas tem\u00e1ticas, y el '
        '<em>Programa de Contrapartidas</em>, que busca facilitar recursos financieros para '
        'fortalecer iniciativas que ya cuentan con financiaci\u00f3n de cooperaci\u00f3n internacional.</p>'
        '<p>En la segunda pesta\u00f1a, denominada <strong>Proyectos AOD</strong>, encontrar\u00e1 el '
        'listado de proyectos que, de acuerdo con el sistema de informaci\u00f3n C\u00edclope, se est\u00e1n '
        'ejecutando en cada departamento.</p>'
        '<p>En la tercera pesta\u00f1a, denominada <strong>Panorama Nacional</strong>, se presentan '
        'los totales nacionales de intervenciones, cooperantes y recursos, e identifica los '
        'principales cooperantes por monto y por n\u00famero de intervenciones, los Objetivos de '
        'Desarrollo Sostenible (ODS) y los sectores de gobierno con mayor financiaci\u00f3n, as\u00ed '
        'como los departamentos con mayor presencia de cooperaci\u00f3n internacional. La informaci\u00f3n '
        'incluye tanto las intervenciones de \u00e1mbito territorial como las de \u00e1mbito nacional.</p>'
        '</div>'
    )
    st.markdown(guia_html, unsafe_allow_html=True)

    col_g1, col_g2 = st.columns(2)
    with col_g1:
        st.info("**Pesta\u00f1a 1 - Ficha territorial**\n\nInformaci\u00f3n general, AOD y programas internos APC-Colombia por departamento. Incluye descarga en Excel.")
    with col_g2:
        st.info("**Pesta\u00f1a 2 - Proyectos AOD**\n\nListado completo de proyectos activos seg\u00fan C\u00edclope. Descarga en CSV y Excel disponible.")

    st.markdown(
        '<div class="apc-footer">Agencia Presidencial de Cooperacion Internacional de Colombia - APC-Colombia</div>',
        unsafe_allow_html=True
    )
