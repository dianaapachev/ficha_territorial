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
 
/* Variables de color institucional APC */
:root {
    --apc-blue:      #003087;
    --apc-green:     #4CAF50;
    --apc-light:     #E8F0FE;
    --apc-gray:      #F5F7FA;
    --apc-border:    #D0D9EA;
    --apc-text:      #1A1A2E;
    --apc-muted:     #6B7280;
    --apc-white:     #FFFFFF;
    --apc-accent:    #1565C0;
}
 
html, body, [class*="css"] {
    font-family: 'Open Sans', sans-serif;
    color: var(--apc-text);
}
 
/* Header superior */
.apc-header {
    background: linear-gradient(135deg, #003087 0%, #1565C0 60%, #1976D2 100%);
    padding: 1.5rem 2rem 1.2rem 2rem;
    border-radius: 0 0 16px 16px;
    margin-bottom: 1.5rem;
    display: flex;
    align-items: center;
    justify-content: space-between;
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
.apc-logo-badge {
    background: rgba(255,255,255,0.15);
    border: 1.5px solid rgba(255,255,255,0.35);
    border-radius: 10px;
    padding: 6px 14px;
    color: white;
    font-family: 'Montserrat', sans-serif;
    font-weight: 700;
    font-size: 1.1rem;
    letter-spacing: 1px;
}
 
/* Selector departamento */
.dept-selector-card {
    background: var(--apc-gray);
    border: 1.5px solid var(--apc-border);
    border-left: 5px solid var(--apc-blue);
    border-radius: 10px;
    padding: 1rem 1.5rem;
    margin-bottom: 1.2rem;
}
 
/* Section headers */
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
 
/* Nombre del departamento destacado */
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
 
/* Métricas */
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
 
/* Tabs */
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
 
/* Dataframes */
div[data-testid="stDataFrame"] {
    border-radius: 8px;
    overflow: hidden;
    border: 1px solid var(--apc-border);
}
 
/* Botón descarga */
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
 
/* Guía de usuario */
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
.info-badge {
    display: inline-block;
    background: var(--apc-light);
    color: var(--apc-blue);
    font-family: 'Montserrat', sans-serif;
    font-weight: 700;
    font-size: 0.75rem;
    border-radius: 20px;
    padding: 2px 10px;
    margin-left: 6px;
    vertical-align: middle;
}
 
/* Footer */
.apc-footer {
    text-align: center;
    color: var(--apc-muted);
    font-size: 0.78rem;
    margin-top: 2.5rem;
    padding-top: 1rem;
    border-top: 1px solid var(--apc-border);
}
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
    raise KeyError(f"No encontré la tabla: {table_name}")
 
 
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
 
 
def to_excel_ficha(info_row, cic_dept, colcol_dept, contr_dept):
    """Genera Excel con múltiples hojas para la ficha territorial."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Hoja 1: Información general completa (registro completo)
        if not info_row.empty:
            df_info = info_row.T.reset_index()
            df_info.columns = ["Campo", "Valor"]
            df_info = df_info[
                ~df_info["Campo"].astype(str).str.lower().str.strip().isin(
                    ["porcentaje de avance", "dept_norm"]
                )
            ]
            df_info.to_excel(writer, sheet_name="Información General", index=False)
 
        # Hoja 2: AOD - Cíclope
        cic_export = cic_dept.drop(columns=["DEPT_NORM"], errors="ignore")
        cic_export.to_excel(writer, sheet_name="AOD - Cíclope", index=False)
 
        # Hoja 3: ColCol
        colcol_export = colcol_dept.copy()
        colcol_export.to_excel(writer, sheet_name="ColCol", index=False)
 
        # Hoja 4: Contrapartidas
        contr_dept.to_excel(writer, sheet_name="Contrapartidas", index=False)
 
    output.seek(0)
    return output.getvalue()
 
 
def to_excel_proyectos(df_proj):
    """Genera Excel para proyectos AOD."""
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
        <div class="apc-header-title">🌐 Ficha Territorial</div>
        <div class="apc-header-subtitle">Caracterización por departamento</div>
    </div>
<img src="https://raw.githubusercontent.com/dianaapachev/ficha_territorial/main/Logo-APC-Color.png" style="height:55px; filter: brightness(0) invert(1);">
</div>
""", unsafe_allow_html=True)
 
infogeneral, plan, ciclope, colcol, contrapartidas, proyectos = load_data()
 
DEPT_COL_INFO = "Departamento"
depts = sorted(infogeneral[DEPT_COL_INFO].dropna().unique().tolist())
 
# Selector con card
st.markdown('<div class="dept-selector-card">', unsafe_allow_html=True)
dept = st.selectbox("🗺️ Selecciona un departamento", depts, label_visibility="visible")
st.markdown('</div>', unsafe_allow_html=True)
 
dept_norm = norm_text(dept)
infogeneral["DEPT_NORM"] = infogeneral[DEPT_COL_INFO].map(norm_text)
 
info = infogeneral[infogeneral["DEPT_NORM"] == dept_norm].head(1)
 
ciclope["DEPT_NORM"] = ciclope["DEPARTAMENTO"].map(norm_text)
proyectos["DEPT_NORM"] = proyectos["DEPARTAMENTO"].map(norm_text)
 
cic_dept = ciclope[ciclope["DEPT_NORM"] == dept_norm]
proj_dept = proyectos[proyectos["DEPT_NORM"] == dept_norm]
 
# ColCol: SOLO por DEPARTAMENTOS PARTICIPANTES
mask_colcol = pd.Series(False, index=colcol.index)
if "DEPARTAMENTOS PARTICIPANTES" in colcol.columns:
    mask_colcol = (
        colcol["DEPARTAMENTOS PARTICIPANTES"]
        .astype("string")
        .map(norm_text)
        .str.contains(dept_norm, na=False)
    )
else:
    st.warning("No encontré la columna 'DEPARTAMENTOS PARTICIPANTES' en ColCol.")
 
colcol_dept = colcol[mask_colcol]
 
# Contrapartidas
if "Departamento" in contrapartidas.columns:
    contr_dept = contrapartidas[
        contrapartidas["Departamento"].astype("string").map(norm_text) == dept_norm
    ]
else:
    contr_dept = contrapartidas.iloc[0:0]
 
 
tab1, tab2, tab3 = st.tabs(["📋 Ficha territorial", "🗂️ Proyectos AOD", "📖 Guía de usuario"])
 
 
# =========================================================
# TAB 1 – FICHA TERRITORIAL
# =========================================================
with tab1:
 
    # Nombre del departamento como banner
    st.markdown(
        f'<div class="dept-title-banner">📍 {dept}</div>',
        unsafe_allow_html=True
    )
 
    # ---- Información general ----
    st.markdown('<div class="section-header">Información General</div>', unsafe_allow_html=True)
 
    if info.empty:
        st.warning("No encontré el departamento en la tabla infogeneral.")
    else:
        c1, c2, c3 = st.columns(3)
 
        c1.metric("🏙️ Capital", info.iloc[0].get("Capital", ""))
        c2.metric("🏘️ Número de Municipios",
                  info.iloc[0].get("Número de Municipios",
                                   info.iloc[0].get("Municipios", "")))
 
        pob = info.iloc[0].get("Población", None)
        pob_fmt = f"{int(float(pob)):,}".replace(",", ".") if pd.notna(pob) else ""
        c3.metric("👥 Población", pob_fmt)
 
        with st.expander("📄 Ver registro completo del departamento"):
            df_det = info.T.reset_index()
            df_det.columns = ["Campo", "Valor"]
            df_det = df_det[
                ~df_det["Campo"].astype(str).str.lower().str.strip().isin(
                    ["porcentaje de avance", "dept_norm"]
                )
            ]
            st.dataframe(df_det, use_container_width=True, hide_index=True)
 
    # ---- AOD ----
    st.markdown('<div class="section-header">Ayuda Oficial al Desarrollo (AOD)</div>',
                unsafe_allow_html=True)
 
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("📊 Intervenciones (únicas)",
              cic_dept["CODIGO INTERVENCION"].nunique()
              if "CODIGO INTERVENCION" in cic_dept.columns else 0)
    m2.metric("🤝 Cooperantes",
              cic_dept["NOMBRE ACTOR"].nunique()
              if "NOMBRE ACTOR" in cic_dept.columns else 0)
    m3.metric("📍 Municipios intervenidos",
              cic_dept["MUNICIPIO"].nunique()
              if "MUNICIPIO" in cic_dept.columns else 0)
 
    total_usd = cic_dept["VALOR APORTE (USD)"].sum() \
        if "VALOR APORTE (USD)" in cic_dept.columns else 0
    m4.metric("💰 Total aporte estimado (USD)", format_usd(total_usd))
 
    c5, c6 = st.columns(2)
 
    with c5:
        st.markdown("**🏆 Top 5 cooperantes por USD**")
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
        st.markdown("**🎯 Top 5 ODS por USD**")
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
            st.dataframe(top_ods_disp, use_container_width=True, hide_index=True)
        else:
            st.info("Sin datos suficientes para ODS.")
 
    # ---- Programas internos APC ----
    st.markdown('<div class="section-header">Programas Internos APC-Colombia</div>',
                unsafe_allow_html=True)
 
    p1, p2 = st.columns(2)
 
    with p1:
        st.markdown("**🇨🇴 ColCol – Colombia Enseña Colombia**")
        st.metric("Registros encontrados", len(colcol_dept))
        colcol_view = colcol_dept.copy()
        if "PRESUPUESTO ESTIMADO APC COLOMBIA" in colcol_view.columns:
            colcol_view["PRESUPUESTO ESTIMADO APC COLOMBIA"] = (
                pd.to_numeric(colcol_view["PRESUPUESTO ESTIMADO APC COLOMBIA"], errors="coerce")
                .apply(format_cop)
            )
        st.dataframe(colcol_view.head(50), use_container_width=True, hide_index=True)
 
    with p2:
        st.markdown("**💼 Contrapartidas**")
        st.metric("Registros encontrados", len(contr_dept))
        st.dataframe(contr_dept.head(50), use_container_width=True, hide_index=True)
 
    # ---- Descarga Excel Ficha ----
    st.markdown("---")
    st.markdown("**⬇️ Descargar ficha territorial completa**")
    excel_ficha = to_excel_ficha(info, cic_dept, colcol_dept, contr_dept)
    st.download_button(
        label="📥 Descargar Ficha Territorial (Excel)",
        data=excel_ficha,
        file_name=f"Ficha_Territorial_{dept}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
 
    st.markdown(
        '<div class="apc-footer">Agencia Presidencial de Cooperación Internacional de Colombia · APC-Colombia</div>',
        unsafe_allow_html=True
    )
 
 
# =========================================================
# TAB 2 – PROYECTOS AOD
# =========================================================
with tab2:
 
    st.markdown(
        f'<div class="dept-title-banner">📍 {dept} — Proyectos AOD activos</div>',
        unsafe_allow_html=True
    )
    st.caption("Fuente: Cíclope a corte de 31 de diciembre de 2025")
 
    search = st.text_input("🔍 Buscar en proyectos").strip()
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
            "📄 Descargar CSV filtrado",
            data=df.to_csv(index=False).encode("utf-8"),
            file_name=f"proyectos_aod_{dept}.csv",
            mime="text/csv"
        )
    with col_dl2:
        excel_proj = to_excel_proyectos(df)
        st.download_button(
            label="📥 Descargar Excel",
            data=excel_proj,
            file_name=f"Proyectos_AOD_{dept}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
 
    st.markdown(
        '<div class="apc-footer">Agencia Presidencial de Cooperación Internacional de Colombia · APC-Colombia</div>',
        unsafe_allow_html=True
    )
 
 
# =========================================================
# TAB 3 – GUÍA DE USUARIO
# =========================================================
with tab3:
 
    st.markdown(
        f'<div class="dept-title-banner">📖 Guía de usuario</div>',
        unsafe_allow_html=True
    )
 
    st.markdown("""
    <div class="guia-card">
        <div class="guia-intro">
            ¿Qué es la Ficha Territorial?
        </div>
        <p>
            La ficha territorial es una aplicación que permite conocer cómo se está moviendo la
            cooperación internacional en los diferentes departamentos de Colombia.
        </p>
        <p>
            En la primera sección encontrará <strong>información general del territorio</strong>,
            como la capital, el número de municipios, la población y otros datos relevantes para
            la gestión de la cooperación internacional. También podrá identificar si el departamento
            cuenta con una dependencia encargada de estos temas, los enlaces o personas clave que
            participan en la gobernanza de la cooperación en el territorio y si dispone de un plan
            de trabajo dentro del Sistema Nacional de Cooperación Internacional, entre otros
            elementos de contexto institucional.
        </p>
        <p>
            Posteriormente, encontrará una sección relacionada con la
            <strong>Ayuda Oficial al Desarrollo (AOD)</strong>. Esta información proviene del
            sistema de información Cíclope, administrado por la Agencia Presidencial de
            Cooperación Internacional de Colombia (APC-Colombia). En este apartado se presentan,
            entre otros aspectos, los principales cooperantes presentes en el territorio, los
            municipios que están siendo intervenidos y una estimación de los recursos provenientes
            de cooperación internacional. Asimismo, podrá identificar los Objetivos de Desarrollo
            Sostenible (ODS) que concentran mayor financiación en cada departamento.
        </p>
        <p>
            Finalmente, la ficha también muestra si el departamento participa en algunos de los
            <strong>programas de la oferta institucional de APC-Colombia</strong>. Entre ellos se
            encuentran la estrategia <em>Colombia Enseña Colombia</em>, orientada a promover
            intercambios de conocimiento en diversas temáticas, y el
            <em>Programa de Contrapartidas</em>, que busca facilitar recursos financieros para
            fortalecer iniciativas que ya cuentan con financiación de cooperación internacional.
        </p>
        <p>
            En la segunda pestaña, denominada <strong>Proyectos AOD</strong>, encontrará el
            listado de proyectos que, de acuerdo con el sistema de información Cíclope, se están
            ejecutando en cada departamento.
        </p>
    </div>
    """, unsafe_allow_html=True)
 
    # Resumen visual de las pestañas
    col_g1, col_g2 = st.columns(2)
    with col_g1:
        st.info("📋 **Pestaña 1 – Ficha territorial**\n\nInformación general, AOD y programas internos APC-Colombia por departamento. Incluye descarga en Excel.")
    with col_g2:
        st.info("🗂️ **Pestaña 2 – Proyectos AOD**\n\nListado completo de proyectos activos según Cíclope. Descarga en CSV y Excel disponible.")
 
    st.markdown(
        '<div class="apc-footer">Agencia Presidencial de Cooperación Internacional de Colombia · APC-Colombia</div>',
        unsafe_allow_html=True
    )
 