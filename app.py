import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries
import unicodedata
import re
import altair as alt

FILE = "Ficha territorial.xlsm"


def norm_text(x):
    if x is None:
        return ""
    s = str(x).strip().upper()
    # quita tildes
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
    # quita puntuación (comas, puntos, etc.) y deja solo letras/números/espacios
    s = re.sub(r"[^A-Z0-9\s]", " ", s)
    # colapsa espacios múltiples
    s = re.sub(r"\s+", " ", s).strip()
    return s


def format_usd(n):
    """Formatea números como USD con separador en puntos y sin decimales."""
    try:
        n = float(n)
    except Exception:
        return ""
    return "USD " + f"{n:,.0f}".replace(",", ".")


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

    # Normaliza strings
    for df in [infogeneral, plan, ciclope, colcol, contrapartidas, proyectos]:
        for c in df.columns:
            if df[c].dtype == "object":
                df[c] = df[c].astype(str).str.strip()

    # USD numérico
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


# --- APP ---
st.set_page_config(page_title="Ficha Territorial", layout="wide")
st.title("Ficha Territorial")

infogeneral, plan, ciclope, colcol, contrapartidas, proyectos = load_data()

DEPT_COL_INFO = "Departamento"
depts = sorted(infogeneral[DEPT_COL_INFO].dropna().unique().tolist())
dept = st.selectbox("Selecciona un departamento", depts)

dept_norm = norm_text(dept)
infogeneral["DEPT_NORM"] = infogeneral[DEPT_COL_INFO].map(norm_text)

# Info general por depto (con normalización)
info = infogeneral[infogeneral["DEPT_NORM"] == dept_norm].head(1)

# Normalización depto en ciclope y proyectos
ciclope["DEPT_NORM"] = ciclope["DEPARTAMENTO"].map(norm_text)
proyectos["DEPT_NORM"] = proyectos["DEPARTAMENTO"].map(norm_text)

cic_dept = ciclope[ciclope["DEPT_NORM"] == dept_norm]
proj_dept = proyectos[proyectos["DEPT_NORM"] == dept_norm]

# Programas internos
colcol_dept = colcol[
    (colcol.get("DEPARTAMENTO EN EL QUE SE DESARROLLÓ", "").astype(str) == dept)
    | (colcol.get("DEPARTAMENTOS PARTICIPANTES", "").astype(str).str.contains(dept, na=False))
    | (colcol.get("DEPARTAMENTOS MAY", "").astype(str).str.contains(dept, na=False))
]

contr_dept = (
    contrapartidas[contrapartidas["Departamento"] == dept]
    if "Departamento" in contrapartidas.columns
    else contrapartidas.iloc[0:0]
)

tab1, tab2 = st.tabs(["Ficha territorial", "Proyectos AOD"])


# -------------------------
# TAB 1
# -------------------------
with tab1:
    st.subheader("Información general")

    if info.empty:
        st.warning("No encontré el departamento en la tabla infogeneral.")
    else:
        c1, c2, c3 = st.columns(3)

        c1.metric("Capital", info.iloc[0].get("Capital", ""))
        c2.metric("# Municipios", info.iloc[0].get("Número de Municipios", info.iloc[0].get("Municipios", "")))

        pob = info.iloc[0].get("Población", None)
        if pd.notna(pob):
            pob_fmt = f"{int(float(pob)):,}".replace(",", ".")
        else:
            pob_fmt = ""
        c3.metric("Población", pob_fmt)

        with st.expander("Ver registro completo"):
            df_det = info.T.reset_index()
            df_det.columns = ["Campo", "Valor"]

            # eliminar filas técnicas
            df_det = df_det[
                ~df_det["Campo"].astype(str).str.lower().str.strip().isin(
                    ["porcentaje de avance", "dept_norm"]
                )
            ]

            st.dataframe(df_det, use_container_width=True, hide_index=True)

    st.subheader("Ayuda Oficial al Desarrollo (AOD)")

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Intervenciones (únicas)",
              cic_dept["CODIGO INTERVENCION"].nunique() if "CODIGO INTERVENCION" in cic_dept.columns else 0)
    m2.metric("Cooperantes",
              cic_dept["NOMBRE ACTOR"].nunique() if "NOMBRE ACTOR" in cic_dept.columns else 0)
    m3.metric("Municipios intervenidos",
              cic_dept["MUNICIPIO"].nunique() if "MUNICIPIO" in cic_dept.columns else 0)

    total_usd = cic_dept["VALOR APORTE (USD)"].sum() if "VALOR APORTE (USD)" in cic_dept.columns else 0
    m4.metric("Total aporte estimado (USD)", format_usd(total_usd))

    c5, c6 = st.columns(2)

    with c5:
        st.markdown("**Top 5 cooperantes por USD**")
        top_act = top_by_sum(cic_dept, "NOMBRE ACTOR", "VALOR APORTE (USD)", 5)

        if not top_act.empty:
            chart_act = (
                alt.Chart(top_act)
                .mark_bar()
                .encode(
                    y=alt.Y("NOMBRE ACTOR:N", sort="-x", title=""),
                    x=alt.X("VALOR APORTE (USD):Q", title="USD"),
                    tooltip=["NOMBRE ACTOR:N", alt.Tooltip("VALOR APORTE (USD):Q", format=",.0f")]
                )
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
                .mark_bar()
                .encode(
                    y=alt.Y("ODS:N", sort="-x", title=""),
                    x=alt.X("VALOR APORTE (USD):Q", title="USD"),
                    tooltip=["ODS:N", alt.Tooltip("VALOR APORTE (USD):Q", format=",.0f")]
                )
            )
            st.altair_chart(chart_ods, use_container_width=True)

            top_ods_disp = top_ods.copy()
            top_ods_disp["VALOR APORTE (USD)"] = top_ods_disp["VALOR APORTE (USD)"].apply(format_usd)
            st.dataframe(top_ods_disp, use_container_width=True, hide_index=True)
        else:
            st.info("Sin datos suficientes para ODS.")

    st.subheader("Programas internos APC")
    p1, p2 = st.columns(2)

    with p1:
        st.markdown("**ColCol**")
        st.metric("Registros encontrados", len(colcol_dept))
        st.dataframe(colcol_dept.head(50), use_container_width=True)

    with p2:
        st.markdown("**Contrapartidas**")
        st.metric("Registros encontrados", len(contr_dept))
        st.dataframe(contr_dept.head(50), use_container_width=True)


# -------------------------
# TAB 2
# -------------------------
with tab2:
    st.subheader(f"Listado AOD (tabla proyectos) — {dept}")
    st.caption("Fuente: tabla 'proyectos' (pestaña datosproyectos).")

    search = st.text_input("Buscar (opcional)").strip()
    df = proj_dept.copy()

    if search and not df.empty:
        candidate_cols = ["NOMBRE INTERVENCION", "OBJETIVO GENERAL", "NOMBRE ACTOR", "MUNICIPIO", "ODS", "META ODS"]
        cols = [c for c in candidate_cols if c in df.columns]
        if cols:
            mask = False
            for c in cols:
                mask = mask | df[c].astype(str).str.contains(search, case=False, na=False)
            df = df[mask]

    st.dataframe(df, use_container_width=True)

    st.download_button(
        "Descargar CSV filtrado",
        data=df.to_csv(index=False).encode("utf-8"),
        file_name=f"proyectos_aod_{dept}.csv",
        mime="text/csv"
    )
