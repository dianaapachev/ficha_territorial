# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries
import unicodedata
import re
import altair as alt
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
from reportlab.lib.enums import TA_LEFT, TA_CENTER

# -----------------------------------------------
# Estilos APC Colombia
# -----------------------------------------------
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;600;700;800&family=Source+Sans+3:wght@400;600&display=swap');

:root {
    --apc-blue:      #003087;
    --apc-blue-mid:  #004BB4;
    --apc-red:       #C8102E;
    --apc-yellow:    #F5A623;
    --apc-light:     #EEF3FB;
    --apc-gray:      #F7F8FA;
    --apc-gray-mid:  #EAECF0;
    --apc-border:    #D1D9E6;
    --apc-text:      #1C2B4A;
    --apc-muted:     #5A6A85;
    --apc-white:     #FFFFFF;
}

html, body, [class*="css"] {
    font-family: 'Source Sans 3', sans-serif;
    color: var(--apc-text);
    background-color: #F7F8FA;
}

/* \u2500\u2500 Header \u2500\u2500 */
.apc-header {
    background: #003087;
    padding: 1.2rem 2.2rem 1rem 2.2rem;
    margin-bottom: 0;
    display: flex;
    align-items: center;
    justify-content: space-between;
    border-bottom: 4px solid var(--apc-red);
}
.apc-header-title {
    color: white;
    font-family: 'Montserrat', sans-serif;
    font-size: 1.5rem;
    font-weight: 700;
    letter-spacing: 0.2px;
    margin: 0;
}
.apc-header-subtitle {
    color: rgba(255,255,255,0.75);
    font-size: 0.82rem;
    font-weight: 400;
    margin-top: 3px;
    font-family: 'Source Sans 3', sans-serif;
    letter-spacing: 0.3px;
}
.apc-flag-bar {
    height: 5px;
    background: linear-gradient(90deg, #F5A623 33.3%, #003087 33.3% 66.6%, #C8102E 66.6%);
    margin-bottom: 1.5rem;
}

/* \u2500\u2500 Selector card \u2500\u2500 */
.dept-selector-card {
    background: var(--apc-white);
    border-left: 5px solid var(--apc-blue);
    border-radius: 0 8px 8px 0;
    padding: 0.9rem 1.4rem;
    margin-bottom: 1.4rem;
    box-shadow: 0 1px 4px rgba(0,48,135,0.07);
}

/* \u2500\u2500 Banner departamento \u2500\u2500 */
.dept-title-banner {
    background: var(--apc-blue);
    color: white;
    font-family: 'Montserrat', sans-serif;
    font-size: 1.2rem;
    font-weight: 700;
    padding: 0.75rem 1.5rem;
    border-radius: 6px;
    margin-bottom: 1.2rem;
    border-left: 6px solid var(--apc-red);
    letter-spacing: 0.2px;
}

/* \u2500\u2500 Section headers \u2500\u2500 */
.section-header {
    font-family: 'Montserrat', sans-serif;
    font-weight: 700;
    font-size: 0.95rem;
    color: var(--apc-blue);
    text-transform: uppercase;
    letter-spacing: 1px;
    border-bottom: 3px solid var(--apc-yellow);
    padding-bottom: 6px;
    margin: 2rem 0 1rem 0;
}

/* \u2500\u2500 M\u00e9tricas \u2500\u2500 */
div[data-testid="stMetric"] {
    background: var(--apc-white);
    border: 1px solid var(--apc-border);
    border-left: 5px solid var(--apc-blue);
    border-radius: 6px;
    padding: 1rem 1.1rem !important;
    box-shadow: 0 1px 6px rgba(0,48,135,0.06);
    transition: box-shadow 0.2s, border-left-color 0.2s;
}
div[data-testid="stMetric"]:hover {
    box-shadow: 0 4px 14px rgba(0,48,135,0.12);
    border-left-color: var(--apc-red);
}
div[data-testid="stMetricLabel"] {
    font-family: 'Montserrat', sans-serif;
    font-weight: 600;
    font-size: 0.72rem;
    color: var(--apc-muted) !important;
    text-transform: uppercase;
    letter-spacing: 0.7px;
}
div[data-testid="stMetricValue"] {
    font-family: 'Montserrat', sans-serif;
    font-weight: 700;
    font-size: 1.6rem !important;
    color: var(--apc-blue) !important;
}

/* \u2500\u2500 Tabs \u2500\u2500 */
button[data-baseweb="tab"] {
    font-family: 'Montserrat', sans-serif !important;
    font-weight: 600 !important;
    font-size: 0.85rem !important;
    letter-spacing: 0.4px;
    text-transform: uppercase;
    color: var(--apc-muted) !important;
}
button[data-baseweb="tab"][aria-selected="true"] {
    color: var(--apc-blue) !important;
}
div[data-baseweb="tab-highlight"] {
    background-color: var(--apc-red) !important;
    height: 3px !important;
}
div[data-baseweb="tab-border"] {
    background-color: var(--apc-border) !important;
}

/* \u2500\u2500 Dataframes \u2500\u2500 */
div[data-testid="stDataFrame"] {
    border-radius: 6px;
    overflow: hidden;
    border: 1px solid var(--apc-border);
    box-shadow: 0 1px 4px rgba(0,0,0,0.04);
}

/* \u2500\u2500 Botones descarga \u2500\u2500 */
div[data-testid="stDownloadButton"] button {
    background: var(--apc-blue) !important;
    color: white !important;
    border: none !important;
    border-radius: 4px !important;
    font-family: 'Montserrat', sans-serif !important;
    font-weight: 600 !important;
    font-size: 0.82rem !important;
    letter-spacing: 0.5px;
    text-transform: uppercase;
    padding: 0.5rem 1.4rem !important;
    transition: background 0.2s !important;
}
div[data-testid="stDownloadButton"] button:hover {
    background: var(--apc-red) !important;
}

/* \u2500\u2500 Guia card \u2500\u2500 */
.guia-card {
    background: var(--apc-white);
    border: 1px solid var(--apc-border);
    border-radius: 8px;
    padding: 2rem 2.4rem;
    margin-bottom: 1rem;
    line-height: 1.8;
    font-size: 0.97rem;
    box-shadow: 0 1px 6px rgba(0,48,135,0.05);
}
.guia-card p {
    color: var(--apc-text);
    margin-bottom: 1.1rem;
}
.guia-intro {
    font-family: 'Montserrat', sans-serif;
    font-size: 0.95rem;
    font-weight: 700;
    color: var(--apc-blue);
    background: var(--apc-light);
    border-left: 5px solid var(--apc-red);
    border-radius: 0 6px 6px 0;
    padding: 0.8rem 1.2rem;
    margin-bottom: 1.4rem;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}

/* \u2500\u2500 Footer \u2500\u2500 */
.apc-footer {
    text-align: center;
    color: var(--apc-muted);
    font-size: 0.76rem;
    margin-top: 3rem;
    padding-top: 1rem;
    border-top: 1px solid var(--apc-border);
    letter-spacing: 0.3px;
}

div[data-testid="stMetricLabel"] p {
    font-size: 0.65rem !important;
    line-height: 1.3 !important;
    white-space: normal !important;
}

/* Metrica personalizada con texto largo */
.metric-custom {
    background: var(--apc-white);
    border: 1px solid var(--apc-border);
    border-left: 5px solid var(--apc-blue);
    border-radius: 6px;
    padding: 1rem 1.1rem;
    box-shadow: 0 1px 6px rgba(0,48,135,0.06);
    transition: box-shadow 0.2s, border-left-color 0.2s;
    height: 100%;
}
.metric-custom:hover {
    box-shadow: 0 4px 14px rgba(0,48,135,0.12);
    border-left-color: #C8102E;
}
.metric-custom-label {
    font-family: 'Montserrat', sans-serif;
    font-weight: 600;
    font-size: 0.68rem;
    color: #5A6A85;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    line-height: 1.35;
    margin-bottom: 0.4rem;
}
.metric-custom-value {
    font-family: 'Montserrat', sans-serif;
    font-weight: 700;
    font-size: 1.6rem;
    color: #003087;
    line-height: 1.2;
}

/* Ocultar barra Streamlit */
div[data-testid="stToolbar"],
div[data-testid="stDecoration"],
header[data-testid="stHeader"] {
    display: none !important;
}
</style>
""", unsafe_allow_html=True)

FILE = "Ficha_territorial.xlsm"


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
    contrapartidas.columns = [str(c).strip().strip("'") for c in contrapartidas.columns]
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



def to_pdf_ficha(dept, info_row, cic_dept, colcol_dept, contr_dept):
    output = BytesIO()
    doc = SimpleDocTemplate(
        output, pagesize=A4,
        leftMargin=2*cm, rightMargin=2*cm,
        topMargin=2*cm, bottomMargin=2*cm
    )

    # Colores institucionales
    AZUL = colors.HexColor("#003087")
    AZUL_CLARO = colors.HexColor("#E8F0FE")
    GRIS = colors.HexColor("#F5F7FA")
    GRIS_BORDE = colors.HexColor("#D0D9EA")
    BLANCO = colors.white

    styles = getSampleStyleSheet()

    # Estilos personalizados
    estilo_titulo = ParagraphStyle("titulo",
        fontName="Helvetica-Bold", fontSize=20, textColor=BLANCO,
        spaceAfter=4, leading=24)
    estilo_subtitulo = ParagraphStyle("subtitulo",
        fontName="Helvetica", fontSize=9, textColor=colors.HexColor("#CBD5E1"),
        spaceAfter=0)
    estilo_dept = ParagraphStyle("dept",
        fontName="Helvetica-Bold", fontSize=15, textColor=BLANCO,
        spaceAfter=0, leading=20)
    estilo_seccion = ParagraphStyle("seccion",
        fontName="Helvetica-Bold", fontSize=11, textColor=AZUL,
        spaceBefore=14, spaceAfter=6)
    estilo_normal = ParagraphStyle("normal",
        fontName="Helvetica", fontSize=9, textColor=colors.HexColor("#1A1A2E"),
        spaceAfter=3, leading=13)
    estilo_label = ParagraphStyle("label",
        fontName="Helvetica-Bold", fontSize=8, textColor=colors.HexColor("#6B7280"),
        spaceAfter=1)
    estilo_valor = ParagraphStyle("valor",
        fontName="Helvetica-Bold", fontSize=14, textColor=AZUL,
        spaceAfter=2)
    estilo_caption = ParagraphStyle("caption",
        fontName="Helvetica", fontSize=7, textColor=colors.HexColor("#6B7280"),
        spaceAfter=4, alignment=TA_CENTER)

    story = []

    # --- HEADER ---
    header_data = [[
        Paragraph("Ficha Territorial", estilo_titulo),
        ""
    ],[
        Paragraph("Agencia Presidencial de Cooperaci\u00f3n Internacional de Colombia", estilo_subtitulo),
        Paragraph("APC-Colombia", ParagraphStyle("badge",
            fontName="Helvetica-Bold", fontSize=9, textColor=BLANCO,
            alignment=TA_CENTER))
    ]]
    header_table = Table(header_data, colWidths=["75%", "25%"])
    header_table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), AZUL),
        ("ROUNDEDCORNERS", [8,8,8,8]),
        ("TOPPADDING", (0,0), (-1,-1), 14),
        ("BOTTOMPADDING", (0,0), (-1,-1), 14),
        ("LEFTPADDING", (0,0), (0,-1), 16),
        ("RIGHTPADDING", (-1,0), (-1,-1), 16),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
    ]))
    story.append(header_table)
    story.append(Spacer(1, 10))

    # --- BANNER DEPARTAMENTO ---
    dept_table = Table([[Paragraph(f"\U0001f4cd  {dept}", estilo_dept)]], colWidths=["100%"])
    dept_table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), AZUL),
        ("ROUNDEDCORNERS", [8,8,8,8]),
        ("TOPPADDING", (0,0), (-1,-1), 10),
        ("BOTTOMPADDING", (0,0), (-1,-1), 10),
        ("LEFTPADDING", (0,0), (-1,-1), 14),
    ]))
    story.append(dept_table)
    story.append(Spacer(1, 10))

    # --- INFORMACION GENERAL ---
    story.append(HRFlowable(width="100%", thickness=2, color=AZUL, spaceAfter=4))
    story.append(Paragraph("Informaci\u00f3n General", estilo_seccion))

    if not info_row.empty:
        row = info_row.iloc[0]

        # Metricas principales
        capital = str(row.get("Capital", ""))
        municipios = str(row.get("N\u00famero de Municipios", row.get("Numero de Municipios", row.get("Municipios", ""))))
        pob_raw = row.get("Poblaci\u00f3n", row.get("Poblacion", ""))
        try:
            poblacion = f"{int(float(pob_raw)):,}".replace(",", ".")
        except:
            poblacion = str(pob_raw)

        metrics_data = [
            [Paragraph("CAPITAL", estilo_label),
             Paragraph("N\u00daMERO DE MUNICIPIOS", estilo_label),
             Paragraph("POBLACI\u00d3N", estilo_label)],
            [Paragraph(capital, estilo_valor),
             Paragraph(municipios, estilo_valor),
             Paragraph(poblacion, estilo_valor)],
        ]
        metrics_table = Table(metrics_data, colWidths=["33%", "33%", "34%"])
        metrics_table.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,-1), GRIS),
            ("BOX", (0,0), (0,-1), 1, GRIS_BORDE),
            ("BOX", (1,0), (1,-1), 1, GRIS_BORDE),
            ("BOX", (2,0), (2,-1), 1, GRIS_BORDE),
            ("TOPPADDING", (0,0), (-1,-1), 8),
            ("BOTTOMPADDING", (0,0), (-1,-1), 8),
            ("LEFTPADDING", (0,0), (-1,-1), 10),
            ("LINEABOVE", (0,0), (-1,0), 3, AZUL),
            ("ROUNDEDCORNERS", [6,6,6,6]),
        ]))
        story.append(metrics_table)
        story.append(Spacer(1, 8))

        # Registro completo
        story.append(Paragraph("Informaci\u00f3n detallada del departamento", estilo_seccion))
        df_det = info_row.T.reset_index()
        df_det.columns = ["Campo", "Valor"]
        df_det = df_det[~df_det["Campo"].astype(str).str.lower().str.strip().isin(
            ["porcentaje de avance", "dept_norm"])]
        df_det = df_det[df_det["Valor"].astype(str).str.strip().isin(["", "None", "nan"]) == False]

        tabla_info = []
        for _, fila in df_det.iterrows():
            tabla_info.append([
                Paragraph(str(fila["Campo"]), estilo_label),
                Paragraph(str(fila["Valor"]), estilo_normal)
            ])

        if tabla_info:
            t = Table(tabla_info, colWidths=["35%", "65%"])
            t.setStyle(TableStyle([
                ("BACKGROUND", (0,0), (-1,-1), BLANCO),
                ("ROWBACKGROUNDS", (0,0), (-1,-1), [BLANCO, GRIS]),
                ("GRID", (0,0), (-1,-1), 0.5, GRIS_BORDE),
                ("TOPPADDING", (0,0), (-1,-1), 5),
                ("BOTTOMPADDING", (0,0), (-1,-1), 5),
                ("LEFTPADDING", (0,0), (-1,-1), 8),
            ]))
            story.append(t)

    story.append(Spacer(1, 10))

    # --- AOD ---
    story.append(HRFlowable(width="100%", thickness=2, color=AZUL, spaceAfter=4))
    story.append(Paragraph("Ayuda Oficial al Desarrollo (AOD)", estilo_seccion))
    story.append(Paragraph(
        "Fuente: C\u00edclope a corte de 31 de diciembre de 2025",
        estilo_caption))

    cic = cic_dept.drop(columns=["DEPT_NORM"], errors="ignore")
    intervenciones = cic["CODIGO INTERVENCION"].nunique() if "CODIGO INTERVENCION" in cic.columns else 0
    cooperantes = cic["NOMBRE ACTOR"].nunique() if "NOMBRE ACTOR" in cic.columns else 0
    municipios_aod = (
        cic["MUNICIPIO"].map(norm_text)
        .pipe(lambda s: s[~s.isin(["NO REPORTA", "SIN INFORMACION", "NO APLICA", ""])])
        .nunique() if "MUNICIPIO" in cic.columns else 0
    )
    total_usd = cic["VALOR APORTE (USD)"].sum() if "VALOR APORTE (USD)" in cic.columns else 0
    total_fmt = "USD " + f"{total_usd:,.0f}".replace(",", ".")

    aod_metrics = [
        [Paragraph("INTERVENCIONES", estilo_label),
         Paragraph("COOPERANTES", estilo_label),
         Paragraph("MUNICIPIOS / \u00c1REAS", estilo_label),
         Paragraph("TOTAL APORTE (USD)", estilo_label)],
        [Paragraph(str(intervenciones), estilo_valor),
         Paragraph(str(cooperantes), estilo_valor),
         Paragraph(str(municipios_aod), estilo_valor),
         Paragraph(total_fmt, ParagraphStyle("valor_small",
             fontName="Helvetica-Bold", fontSize=10, textColor=AZUL))],
    ]
    t_aod = Table(aod_metrics, colWidths=["25%","25%","25%","25%"])
    t_aod.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), GRIS),
        ("LINEABOVE", (0,0), (-1,0), 3, AZUL),
        ("GRID", (0,0), (-1,-1), 0.5, GRIS_BORDE),
        ("TOPPADDING", (0,0), (-1,-1), 8),
        ("BOTTOMPADDING", (0,0), (-1,-1), 8),
        ("LEFTPADDING", (0,0), (-1,-1), 10),
    ]))
    story.append(t_aod)
    story.append(Spacer(1, 8))

    # Top cooperantes
    if "NOMBRE ACTOR" in cic.columns and "VALOR APORTE (USD)" in cic.columns:
        top_coop = (cic.groupby("NOMBRE ACTOR")["VALOR APORTE (USD)"]
                    .sum().sort_values(ascending=False).head(5).reset_index())
        story.append(Paragraph("Top 5 cooperantes por aporte estimado (USD)", estilo_seccion))
        coop_data = [[
            Paragraph("COOPERANTE", estilo_label),
            Paragraph("APORTE ESTIMADO", estilo_label)
        ]]
        for _, r in top_coop.iterrows():
            coop_data.append([
                Paragraph(str(r["NOMBRE ACTOR"]), estilo_normal),
                Paragraph("USD " + f"{r['VALOR APORTE (USD)']:,.0f}".replace(",","."), estilo_normal)
            ])
        t_coop = Table(coop_data, colWidths=["70%","30%"])
        t_coop.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), AZUL_CLARO),
            ("ROWBACKGROUNDS", (0,1), (-1,-1), [BLANCO, GRIS]),
            ("GRID", (0,0), (-1,-1), 0.5, GRIS_BORDE),
            ("TOPPADDING", (0,0), (-1,-1), 5),
            ("BOTTOMPADDING", (0,0), (-1,-1), 5),
            ("LEFTPADDING", (0,0), (-1,-1), 8),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ]))
        story.append(t_coop)
        story.append(Spacer(1, 8))

    # Top ODS
    if "ODS" in cic.columns and "VALOR APORTE (USD)" in cic.columns:
        top_ods = (cic.groupby("ODS")["VALOR APORTE (USD)"]
                   .sum().sort_values(ascending=False).head(5).reset_index())
        story.append(Paragraph("Top 5 ODS por aporte estimado (USD)", estilo_seccion))
        ods_data = [[
            Paragraph("ODS", estilo_label),
            Paragraph("APORTE ESTIMADO", estilo_label)
        ]]
        for _, r in top_ods.iterrows():
            nombre_ods = ODS_NOMBRES.get(r["ODS"], r["ODS"])
            ods_data.append([
                Paragraph(nombre_ods, estilo_normal),
                Paragraph("USD " + f"{r['VALOR APORTE (USD)']:,.0f}".replace(",","."), estilo_normal)
            ])
        t_ods = Table(ods_data, colWidths=["70%","30%"])
        t_ods.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), AZUL_CLARO),
            ("ROWBACKGROUNDS", (0,1), (-1,-1), [BLANCO, GRIS]),
            ("GRID", (0,0), (-1,-1), 0.5, GRIS_BORDE),
            ("TOPPADDING", (0,0), (-1,-1), 5),
            ("BOTTOMPADDING", (0,0), (-1,-1), 5),
            ("LEFTPADDING", (0,0), (-1,-1), 8),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ]))
        story.append(t_ods)

    story.append(Spacer(1, 10))

    # --- PROGRAMAS INTERNOS ---
    story.append(HRFlowable(width="100%", thickness=2, color=AZUL, spaceAfter=4))
    story.append(Paragraph("Programas Internos APC-Colombia", estilo_seccion))

    # ColCol
    story.append(Paragraph(f"ColCol - Colombia Ense\u00f1a Colombia ({len(colcol_dept)} registros)", estilo_label))
    if not colcol_dept.empty:
        COLS_COLCOL = [
            "NOMBRE DEL INTERCAMBIO",
            "A\u00d1O DE REALIZACI\u00d3N ",
            "DEPARTAMENTO EN EL QUE SE DESARROLL\u00d3",
            "PRESUPUESTO ESTIMADO APC COLOMBIA",
            "RUBRO ASUMIDO"
        ]
        cols_colcol = [c for c in COLS_COLCOL if c in colcol_dept.columns]
        colcol_show = colcol_dept[cols_colcol].head(30)
        HEADERS_COLCOL = {
            "NOMBRE DEL INTERCAMBIO": "Nombre del intercambio",
            "A\u00d1O DE REALIZACI\u00d3N ": "A\u00f1o",
            "DEPARTAMENTO EN EL QUE SE DESARROLL\u00d3": "Departamento",
            "PRESUPUESTO ESTIMADO APC COLOMBIA": "Presupuesto APC",
            "RUBRO ASUMIDO": "Rubro"
        }
        colcol_show = colcol_show.copy()
        if "PRESUPUESTO ESTIMADO APC COLOMBIA" in colcol_show.columns:
            colcol_show["PRESUPUESTO ESTIMADO APC COLOMBIA"] = (
                pd.to_numeric(colcol_show["PRESUPUESTO ESTIMADO APC COLOMBIA"], errors="coerce")
                .apply(format_cop)
            )
        cc_data = [[Paragraph(HEADERS_COLCOL.get(c, c), estilo_label) for c in colcol_show.columns]]
        for _, r in colcol_show.iterrows():
            cc_data.append([Paragraph(str(v)[:80], estilo_normal) for v in r.values])
        col_widths_cc = ["35%", "8%", "18%", "22%", "17%"][:len(colcol_show.columns)]
        t_cc = Table(cc_data, colWidths=col_widths_cc)
        t_cc.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), AZUL_CLARO),
            ("ROWBACKGROUNDS", (0,1), (-1,-1), [BLANCO, GRIS]),
            ("GRID", (0,0), (-1,-1), 0.5, GRIS_BORDE),
            ("TOPPADDING", (0,0), (-1,-1), 5),
            ("BOTTOMPADDING", (0,0), (-1,-1), 5),
            ("LEFTPADDING", (0,0), (-1,-1), 6),
            ("FONTSIZE", (0,0), (-1,-1), 8),
        ]))
        story.append(t_cc)
    else:
        story.append(Paragraph("Sin registros para este departamento.", estilo_normal))
    story.append(Spacer(1, 8))

    # Contrapartidas
    story.append(Paragraph(f"Contrapartidas ({len(contr_dept)} registros)", estilo_label))
    if not contr_dept.empty:
        cols_contr = [c for c in contr_dept.columns if c not in ["DEPT_NORM", "Departamento"]]
        contr_show = contr_dept[cols_contr].head(30).copy()
        for col in contr_show.columns:
            if str(col).strip().strip("\'") in ["Monto por APC", "Monto total", "Monto total "]:
                contr_show[col] = pd.to_numeric(contr_show[col], errors="coerce").apply(format_cop)
        ct_data = [[Paragraph(str(c), estilo_label) for c in contr_show.columns]]
        for _, r in contr_show.iterrows():
            ct_data.append([Paragraph(str(v)[:80], estilo_normal) for v in r.values])
        n_cols = len(contr_show.columns)
        col_w = 100 / n_cols
        t_ct = Table(ct_data, colWidths=[f"{col_w}%" for _ in range(n_cols)])
        t_ct.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), AZUL_CLARO),
            ("ROWBACKGROUNDS", (0,1), (-1,-1), [BLANCO, GRIS]),
            ("GRID", (0,0), (-1,-1), 0.5, GRIS_BORDE),
            ("TOPPADDING", (0,0), (-1,-1), 5),
            ("BOTTOMPADDING", (0,0), (-1,-1), 5),
            ("LEFTPADDING", (0,0), (-1,-1), 6),
            ("FONTSIZE", (0,0), (-1,-1), 8),
            ("WORDWRAP", (0,0), (-1,-1), True),
        ]))
        story.append(t_ct)
    else:
        story.append(Paragraph("Sin registros para este departamento.", estilo_normal))

    # Footer
    story.append(Spacer(1, 16))
    story.append(HRFlowable(width="100%", thickness=0.5, color=GRIS_BORDE, spaceAfter=4))
    story.append(Paragraph(
        "Agencia Presidencial de Cooperaci\u00f3n Internacional de Colombia - APC-Colombia",
        estilo_caption))

    doc.build(story)
    output.seek(0)
    return output.getvalue()


def to_pdf_proyectos(dept, df_proj):
    output = BytesIO()
    doc = SimpleDocTemplate(
        output, pagesize=A4,
        leftMargin=2*cm, rightMargin=2*cm,
        topMargin=2*cm, bottomMargin=2*cm
    )

    AZUL = colors.HexColor("#003087")
    AZUL_CLARO = colors.HexColor("#E8F0FE")
    GRIS = colors.HexColor("#F5F7FA")
    GRIS_BORDE = colors.HexColor("#D0D9EA")
    BLANCO = colors.white

    styles = getSampleStyleSheet()
    estilo_titulo = ParagraphStyle("titulo", fontName="Helvetica-Bold", fontSize=20,
        textColor=BLANCO, spaceAfter=4, leading=24)
    estilo_subtitulo = ParagraphStyle("subtitulo", fontName="Helvetica", fontSize=9,
        textColor=colors.HexColor("#CBD5E1"), spaceAfter=0)
    estilo_dept = ParagraphStyle("dept", fontName="Helvetica-Bold", fontSize=15,
        textColor=BLANCO, spaceAfter=0, leading=20)
    estilo_label = ParagraphStyle("label", fontName="Helvetica-Bold", fontSize=8,
        textColor=colors.HexColor("#6B7280"), spaceAfter=1)
    estilo_normal = ParagraphStyle("normal", fontName="Helvetica", fontSize=8,
        textColor=colors.HexColor("#1A1A2E"), spaceAfter=3, leading=11)
    estilo_caption = ParagraphStyle("caption", fontName="Helvetica", fontSize=7,
        textColor=colors.HexColor("#6B7280"), spaceAfter=4, alignment=TA_CENTER)

    story = []

    # Header
    header_data = [[
        Paragraph("Ficha Territorial", estilo_titulo), ""
    ],[
        Paragraph("Agencia Presidencial de Cooperaci\u00f3n Internacional de Colombia", estilo_subtitulo), ""
    ]]
    header_table = Table(header_data, colWidths=["100%", "0%"])
    header_table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), AZUL),
        ("TOPPADDING", (0,0), (-1,-1), 14),
        ("BOTTOMPADDING", (0,0), (-1,-1), 14),
        ("LEFTPADDING", (0,0), (-1,-1), 16),
    ]))
    story.append(header_table)
    story.append(Spacer(1, 10))

    # Banner departamento
    dept_table = Table([[Paragraph(f"\U0001f4cd  {dept} \u2014 Proyectos AOD activos", estilo_dept)]], colWidths=["100%"])
    dept_table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), AZUL),
        ("TOPPADDING", (0,0), (-1,-1), 10),
        ("BOTTOMPADDING", (0,0), (-1,-1), 10),
        ("LEFTPADDING", (0,0), (-1,-1), 14),
    ]))
    story.append(dept_table)
    story.append(Spacer(1, 6))
    story.append(Paragraph(
        f"Total proyectos: {len(df_proj)} | Fuente: C\u00edclope a corte de 31 de diciembre de 2025",
        estilo_caption))

    # Tabla de proyectos
    COLS = ["NOMBRE INTERVENCION", "FECHA INICIAL", "FECHA FINAL", "MUNICIPIO", "NOMBRE ACTOR"]
    HEADERS = {
        "NOMBRE INTERVENCION": "Nombre de la intervenci\u00f3n",
        "FECHA INICIAL": "Fecha inicial",
        "FECHA FINAL": "Fecha final",
        "MUNICIPIO": "Municipio",
        "NOMBRE ACTOR": "Cooperante"
    }
    COL_WIDTHS = ["32%", "10%", "10%", "16%", "32%"]

    cols_available = [c for c in COLS if c in df_proj.columns]
    widths_available = [COL_WIDTHS[COLS.index(c)] for c in cols_available]

    if not df_proj.empty and cols_available:
        proj_data = [[Paragraph(HEADERS.get(c, c), estilo_label) for c in cols_available]]
        for _, r in df_proj[cols_available].iterrows():
            proj_data.append([Paragraph(str(v)[:120] if v and str(v) != "nan" else "", estilo_normal) for v in r.values])

        t_proj = Table(proj_data, colWidths=widths_available, repeatRows=1)
        t_proj.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), AZUL_CLARO),
            ("ROWBACKGROUNDS", (0,1), (-1,-1), [BLANCO, GRIS]),
            ("GRID", (0,0), (-1,-1), 0.5, GRIS_BORDE),
            ("TOPPADDING", (0,0), (-1,-1), 5),
            ("BOTTOMPADDING", (0,0), (-1,-1), 5),
            ("LEFTPADDING", (0,0), (-1,-1), 6),
            ("FONTSIZE", (0,0), (-1,-1), 8),
            ("VALIGN", (0,0), (-1,-1), "TOP"),
        ]))
        story.append(t_proj)
    else:
        story.append(Paragraph("Sin proyectos para este departamento.", estilo_normal))

    story.append(Spacer(1, 16))
    story.append(HRFlowable(width="100%", thickness=0.5, color=GRIS_BORDE, spaceAfter=4))
    story.append(Paragraph(
        "Agencia Presidencial de Cooperaci\u00f3n Internacional de Colombia - APC-Colombia",
        estilo_caption))

    doc.build(story)
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
        <div class="apc-header-subtitle">Herramienta de caracterizaci\u00f3n territorial para la gesti\u00f3n de la cooperaci\u00f3n internacional</div>
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
    municipios_count_ant = (
        cic_dept_ant["MUNICIPIO"].map(norm_text)
        .pipe(lambda s: s[~s.isin(["NO REPORTA", "SIN INFORMACION", "NO APLICA", ""])])
        .nunique()
        if "MUNICIPIO" in cic_dept_ant.columns else 0
    )
    delta_mun = municipios_count - municipios_count_ant
    delta_mun_str = ("\u25b2 " if delta_mun >= 0 else "\u25bc ") + str(abs(delta_mun)) + " vs. 2025"
    delta_mun_color = "#2E7D32" if delta_mun >= 0 else "#C8102E"
    with m3:
        st.markdown(
            '<div class="metric-custom">'
            '<div class="metric-custom-label">Municipios o \u00e1reas no municipalizadas intervenidas</div>'
            f'<div class="metric-custom-value">{municipios_count}</div>'
            f'<div class="metric-custom-delta" style="color:{delta_mun_color};">{delta_mun_str}</div>'
            '</div>',
            unsafe_allow_html=True
        )
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
            # Tabla comparativa 2025 vs 2026
            top_act_ant = top_by_sum(cic_dept_ant, "NOMBRE ACTOR", "VALOR APORTE (USD)", 5)
            top_act_disp = top_act.copy()
            top_act_disp.columns = ["NOMBRE ACTOR", "USD 2026"]
            top_act_disp["USD 2026"] = top_act_disp["USD 2026"].apply(format_usd)
            if not top_act_ant.empty:
                top_act_ant.columns = ["NOMBRE ACTOR", "USD 2025"]
                top_act_ant["USD 2025"] = top_act_ant["USD 2025"].apply(format_usd)
                top_act_disp = top_act_disp.merge(top_act_ant, on="NOMBRE ACTOR", how="left").fillna("-")
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
            # Tabla comparativa 2025 vs 2026
            top_ods_ant = top_by_sum(cic_dept_ant, "ODS", "VALOR APORTE (USD)", 5)
            top_ods_disp = top_ods.copy()
            top_ods_disp.columns = ["ODS", "USD 2026"]
            top_ods_disp["ODS"] = top_ods_disp["ODS"].map(lambda x: ODS_NOMBRES.get(x, x))
            top_ods_disp["USD 2026"] = top_ods_disp["USD 2026"].apply(format_usd)
            if not top_ods_ant.empty:
                top_ods_ant_disp = top_ods_ant.copy()
                top_ods_ant_disp.columns = ["ODS", "USD 2025"]
                top_ods_ant_disp["ODS"] = top_ods_ant_disp["ODS"].map(lambda x: ODS_NOMBRES.get(x, x))
                top_ods_ant_disp["USD 2025"] = top_ods_ant_disp["USD 2025"].apply(format_usd)
                top_ods_disp = top_ods_disp.merge(top_ods_ant_disp, on="ODS", how="left").fillna("-")
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
        contr_view = contr_dept.copy()
        for col in contr_view.columns:
            if str(col).strip().strip("\'") in ["Monto por APC", "Monto total", "Monto total "]:
                contr_view[col] = pd.to_numeric(contr_view[col], errors="coerce").apply(format_cop)
        st.dataframe(contr_view.head(50), use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("**Descargar ficha territorial completa**")
    excel_ficha = to_excel_ficha(info, cic_dept, colcol_dept, contr_dept)
    col_pdf, col_xlsx = st.columns(2)
    with col_xlsx:
        st.download_button(
            label="Descargar Ficha Territorial (Excel)",
            data=excel_ficha,
            file_name=f"Ficha_Territorial_{dept}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with col_pdf:
        pdf_ficha = to_pdf_ficha(dept, info, cic_dept, colcol_dept, contr_dept)
        st.download_button(
            label="Descargar Ficha Territorial (PDF)",
            data=pdf_ficha,
            file_name=f"Ficha_Territorial_{dept}.pdf",
            mime="application/pdf",
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
    st.caption("Fuente: C\u00edclope a corte de 31 de diciembre de 2025")

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
        excel_proj = to_excel_proyectos(df)
        st.download_button(
            label="Descargar Excel",
            data=excel_proj,
            file_name=f"Proyectos_AOD_{dept}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with col_dl2:
        pdf_proj = to_pdf_proyectos(dept, df)
        st.download_button(
            label="Descargar PDF",
            data=pdf_proj,
            file_name=f"Proyectos_AOD_{dept}.pdf",
            mime="application/pdf",
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
            top_coop_usd_ant = (cic_ant_nac.groupby("NOMBRE ACTOR")["VALOR APORTE (USD)"]
                .sum().sort_values(ascending=False).head(10).reset_index())
            top_coop_usd_disp = top_coop_usd.copy()
            top_coop_usd_disp.columns = ["NOMBRE ACTOR", "USD 2026"]
            top_coop_usd_disp["USD 2026"] = top_coop_usd_disp["USD 2026"].apply(format_usd)
            top_coop_usd_ant.columns = ["NOMBRE ACTOR", "USD 2025"]
            top_coop_usd_ant["USD 2025"] = top_coop_usd_ant["USD 2025"].apply(format_usd)
            top_coop_usd_disp = top_coop_usd_disp.merge(top_coop_usd_ant, on="NOMBRE ACTOR", how="left").fillna("-")
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
            top_coop_int_ant = (cic_ant_nac.groupby("NOMBRE ACTOR")["CODIGO INTERVENCION"]
                .nunique().sort_values(ascending=False).head(10).reset_index())
            top_coop_int_ant.columns = ["NOMBRE ACTOR", "INT. 2025"]
            top_coop_int_disp = top_coop_int.copy()
            top_coop_int_disp.columns = ["NOMBRE ACTOR", "INT. 2026"]
            top_coop_int_disp = top_coop_int_disp.merge(top_coop_int_ant, on="NOMBRE ACTOR", how="left").fillna("-")
            st.dataframe(top_coop_int_disp, use_container_width=True, hide_index=True)

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
            top_ods_nac_ant = (cic_ant_nac.groupby("ODS")["VALOR APORTE (USD)"]
                .sum().sort_values(ascending=False).head(10).reset_index())
            top_ods_nac_disp = top_ods_nac.copy()
            top_ods_nac_disp.columns = ["ODS", "USD 2026"]
            top_ods_nac_disp["ODS"] = top_ods_nac_disp["ODS"].map(lambda x: ODS_NOMBRES.get(x, x))
            top_ods_nac_disp["USD 2026"] = top_ods_nac_disp["USD 2026"].apply(format_usd)
            top_ods_nac_ant.columns = ["ODS", "USD 2025"]
            top_ods_nac_ant["ODS"] = top_ods_nac_ant["ODS"].map(lambda x: ODS_NOMBRES.get(x, x))
            top_ods_nac_ant["USD 2025"] = top_ods_nac_ant["USD 2025"].apply(format_usd)
            top_ods_nac_disp = top_ods_nac_disp.merge(top_ods_nac_ant, on="ODS", how="left").fillna("-")
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
            top_sect_nac_ant = (cic_ant_nac.groupby("SECTORES GOB")["VALOR APORTE (USD)"]
                .sum().sort_values(ascending=False).head(10).reset_index())
            top_sect_nac_disp = top_sect_nac.copy()
            top_sect_nac_disp.columns = ["SECTORES GOB", "USD 2026"]
            top_sect_nac_disp["USD 2026"] = top_sect_nac_disp["USD 2026"].apply(format_usd)
            top_sect_nac_ant.columns = ["SECTORES GOB", "USD 2025"]
            top_sect_nac_ant["USD 2025"] = top_sect_nac_ant["USD 2025"].apply(format_usd)
            top_sect_nac_disp = top_sect_nac_disp.merge(top_sect_nac_ant, on="SECTORES GOB", how="left").fillna("-")
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
            top_dept_usd_ant = (
                cic_ant_nac[cic_ant_nac["DEPARTAMENTO"] != "\u00c1mbito Nacional"]
                .groupby("DEPARTAMENTO")["VALOR APORTE (USD)"]
                .sum().sort_values(ascending=False).head(10).reset_index())
            top_dept_usd_disp = top_dept_usd.copy()
            top_dept_usd_disp.columns = ["DEPARTAMENTO", "USD 2026"]
            top_dept_usd_disp["USD 2026"] = top_dept_usd_disp["USD 2026"].apply(format_usd)
            top_dept_usd_ant.columns = ["DEPARTAMENTO", "USD 2025"]
            top_dept_usd_ant["USD 2025"] = top_dept_usd_ant["USD 2025"].apply(format_usd)
            top_dept_usd_disp = top_dept_usd_disp.merge(top_dept_usd_ant, on="DEPARTAMENTO", how="left").fillna("-")
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
            top_dept_int_ant = (
                cic_ant_nac[cic_ant_nac["DEPARTAMENTO"] != "\u00c1mbito Nacional"]
                .groupby("DEPARTAMENTO")["CODIGO INTERVENCION"]
                .nunique().sort_values(ascending=False).head(10).reset_index())
            top_dept_int_ant.columns = ["DEPARTAMENTO", "INT. 2025"]
            top_dept_int_disp = top_dept_int.copy()
            top_dept_int_disp.columns = ["DEPARTAMENTO", "INT. 2026"]
            top_dept_int_disp = top_dept_int_disp.merge(top_dept_int_ant, on="DEPARTAMENTO", how="left").fillna("-")
            st.dataframe(top_dept_int_disp, use_container_width=True, hide_index=True)

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
        '<p>En algunos indicadores podr\u00e1 ver un comparativo con el trimestre 4 de 2025, '
        'lo que le permitir\u00e1 identificar cambios en la din\u00e1mica de la cooperaci\u00f3n internacional '
        'entre per\u00edodos. Las flechas \u25b2 (subi\u00f3) y \u25bc (baj\u00f3) indican la variaci\u00f3n '
        'respecto al per\u00edodo anterior.</p>'
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
