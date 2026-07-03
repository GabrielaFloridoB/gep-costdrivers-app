import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="GEP – Cost Drivers Database Export",
    page_icon="📊",
    layout="wide",
)

# ── Styling ────────────────────────────────────────────────────────────────────
HEADER_FILL   = PatternFill("solid", start_color="1F4E79", end_color="1F4E79")
HEADER_FONT   = Font(name="Arial", bold=True, color="FFFFFF", size=10)
GLOSSARY_FILL = PatternFill("solid", start_color="2E75B6", end_color="2E75B6")
DATA_FONT     = Font(name="Arial", size=9)
ALT_FILL      = PatternFill("solid", start_color="DCE6F1", end_color="DCE6F1")
THIN_BORDER   = Border(
    left=Side(style="thin", color="B8CCE4"),
    right=Side(style="thin", color="B8CCE4"),
    top=Side(style="thin", color="B8CCE4"),
    bottom=Side(style="thin", color="B8CCE4"),
)

COLUMN_WIDTHS = {
    "ID": 10, "Indicator": 70, "Source": 25, "Country": 20,
    "HS Code": 18, "ImportExport": 20, "Update frequency": 22,
    "Group": 22, "LastDateUpdated": 20, "Last forecast update": 22,
    "Frequency of publication source": 35,
}

COLUMN_GLOSSARY = [
    ("ID",                              "Unique code that identifies each record in the table."),
    ("Indicator",                       "Name of the monitored indicator."),
    ("Source",                          "Data origin: the entity or platform responsible for publishing the indicator."),
    ("Country",                         "Country the indicator refers to."),
    ("HS Code",                         "Harmonized System code: classifies the product for international trade purposes."),
    ("ImportExport",                    "Indicates whether the trade flow is import or export."),
    ("Update frequency platform",       "How often the internal platform updates the data for this indicator."),
    ("Group",                           "Thematic category of the indicator."),
    ("LastDateUpdated",                 "Date of the last actual data update on the platform."),
    ("Last forecast update",            "Date of the last update to the forecast linked to the indicator."),
    ("Frequency of publication source", "The frequency with which the original source publishes its data."),
]

# Mapeamento: coluna Excel → coluna Databricks
DB_COLUMN_MAP = {
    "ID":                              "IndiceID",
    "Indicator":                       "Indice_1",
    "Source":                          "Fonte_1",
    "Country":                         "Pais_1",
    "HS Code":                         "NCM",
    "ImportExport":                    "ImportExport",
    "Update frequency":                "Periodicidade",
    "Group":                           "IndiceGrupo_1",
    "LastDateUpdated":                 "LastUpdateDate",
    "Last forecast update":            "LastUpdateDatePrediction",
    "Frequency of publication source": None,  # vem do Excel de referência
}

OUTPUT_COLUMNS = list(DB_COLUMN_MAP.keys())


# ── Excel builder ──────────────────────────────────────────────────────────────
def build_excel(df: pd.DataFrame) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "result"
    ws.freeze_panes = "A2"

    ws.append(OUTPUT_COLUMNS)
    for col_idx, col_name in enumerate(OUTPUT_COLUMNS, 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font      = HEADER_FONT
        cell.fill      = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border    = THIN_BORDER
        ws.column_dimensions[get_column_letter(col_idx)].width = COLUMN_WIDTHS.get(col_name, 18)
    ws.row_dimensions[1].height = 30

    for row_idx, row in enumerate(df.itertuples(index=False), 2):
        for col_idx, value in enumerate(row, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.font      = DATA_FONT
            cell.border    = THIN_BORDER
            cell.alignment = Alignment(vertical="center")
            if row_idx % 2 == 0:
                cell.fill = ALT_FILL

    ws.auto_filter.ref = f"A1:{get_column_letter(len(OUTPUT_COLUMNS))}1"

    wg = wb.create_sheet("Column glossary")
    wg.freeze_panes = "A2"
    wg.column_dimensions["A"].width = 35
    wg.column_dimensions["B"].width = 120
    wg.append(["Column", "Description"])
    for col_idx in range(1, 3):
        cell = wg.cell(row=1, column=col_idx)
        cell.font      = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        cell.fill      = GLOSSARY_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border    = THIN_BORDER
    for r_idx, (col_name, desc) in enumerate(COLUMN_GLOSSARY, 2):
        c1 = wg.cell(row=r_idx, column=1, value=col_name)
        c2 = wg.cell(row=r_idx, column=2, value=desc)
        for c in (c1, c2):
            c.font      = DATA_FONT
            c.border    = THIN_BORDER
            c.alignment = Alignment(vertical="center", wrap_text=True)
        if r_idx % 2 == 0:
            c1.fill = ALT_FILL
            c2.fill = ALT_FILL

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ── Databricks loader ──────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False, ttl=300)
def load_from_databricks(filters: dict) -> pd.DataFrame:
    from databricks import sql as dbsql

    host      = st.secrets["databricks"]["host"]
    token     = st.secrets["databricks"]["token"]
    http_path = st.secrets["databricks"]["http_path"]

    db_cols       = [v for v in DB_COLUMN_MAP.values() if v is not None]
    select_clause = ", ".join(f"`{c}`" for c in db_cols)

    where_clauses = []
    if filters.get("groups"):
        vals = ", ".join(f"'{g}'" for g in filters["groups"])
        where_clauses.append(f"`IndiceGrupo_1` IN ({vals})")
    if filters.get("countries"):
        vals = ", ".join(f"'{c}'" for c in filters["countries"])
        where_clauses.append(f"`Pais_1` IN ({vals})")
    if filters.get("sources"):
        vals = ", ".join(f"'{s}'" for s in filters["sources"])
        where_clauses.append(f"`Fonte_1` IN ({vals})")

    where_sql = ("WHERE " + " AND ".join(where_clauses)) if where_clauses else ""
    query = f"SELECT {select_clause} FROM costdrivers_dg_prod.`00_raw`.costdrivers_infos {where_sql}"

    with dbsql.connect(
        server_hostname=host,
        http_path=http_path,
        access_token=token,
    ) as conn:
        with conn.cursor() as cursor:
            cursor.execute(query)
            rows    = cursor.fetchall()
            columns = [d[0] for d in cursor.description]

    df_raw = pd.DataFrame(rows, columns=columns)
    reverse_map = {v: k for k, v in DB_COLUMN_MAP.items() if v is not None}
    df_raw.rename(columns=reverse_map, inplace=True)
    df_raw["Frequency of publication source"] = ""
    return df_raw[OUTPUT_COLUMNS]


# ── Merge com Excel de referência ──────────────────────────────────────────────
def merge_frequency(df: pd.DataFrame, ref_file) -> pd.DataFrame:
    """Cruza pelo ID e preenche Frequency of publication source do Excel antigo."""
    ref = pd.read_excel(ref_file, sheet_name="result", usecols=["ID", "Frequency of publication source"])
    ref["ID"] = ref["ID"].astype(str)
    df["ID"]  = df["ID"].astype(str)
    df = df.merge(ref, on="ID", how="left", suffixes=("", "_ref"))
    df["Frequency of publication source"] = df["Frequency of publication source_ref"].fillna("")
    df.drop(columns=["Frequency of publication source_ref"], inplace=True)
    return df[OUTPUT_COLUMNS]


# ── UI ─────────────────────────────────────────────────────────────────────────
st.title("📊 GEP – Cost Drivers Database Export")
st.caption("Busca os dados mais atualizados do Databricks e gera o Excel padronizado.")

with st.sidebar:
    st.header("📁 Arquivo de referência")
    st.markdown(
        "Faça upload do Excel anterior para preservar a coluna "
        "**Frequency of publication source** dos indicadores já existentes."
    )
    ref_file = st.file_uploader("Excel de referência (opcional)", type=["xlsx", "xls"])
    if ref_file:
        st.success("✅ Arquivo carregado!")

    st.divider()
    st.header("🔍 Filtros (opcionais)")
    group_filter   = st.text_area("Groups (um por linha)",    placeholder="Services\nFuels")
    country_filter = st.text_area("Countries (um por linha)", placeholder="Brazil\nUnited States")
    source_filter  = st.text_area("Sources (um por linha)",   placeholder="US BLS\nFRED")
    load_btn = st.button("🔄 Carregar dados do Databricks", use_container_width=True)

if load_btn:
    filters = {
        "groups":    [g.strip() for g in group_filter.splitlines() if g.strip()],
        "countries": [c.strip() for c in country_filter.splitlines() if c.strip()],
        "sources":   [s.strip() for s in source_filter.splitlines() if s.strip()],
    }
    with st.spinner("Conectando ao Databricks e carregando dados…"):
        try:
            df = load_from_databricks(filters)

            if ref_file:
                with st.spinner("Cruzando com o arquivo de referência…"):
                    df = merge_frequency(df, ref_file)
                filled = (df["Frequency of publication source"] != "").sum()
                empty  = (df["Frequency of publication source"] == "").sum()
                st.info(f"📋 {filled:,} indicadores com frequência preservada | {empty:,} novos (vazios)")

            st.session_state["df"] = df
            st.success(f"✅ {len(df):,} registros carregados.")
        except Exception as e:
            st.error(f"Erro ao conectar: {e}")

if "df" in st.session_state:
    df = st.session_state["df"]

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total de registros", f"{len(df):,}")
    col2.metric("Países",  df["Country"].nunique() if "Country" in df.columns else "–")
    col3.metric("Groups",  df["Group"].nunique()   if "Group"   in df.columns else "–")
    col4.metric("Fontes",  df["Source"].nunique()  if "Source"  in df.columns else "–")

    with st.expander("👀 Pré-visualização (primeiras 100 linhas)"):
        st.dataframe(df.head(100), use_container_width=True)

    st.divider()
    if st.button("📥 Gerar Excel padronizado", use_container_width=True):
        with st.spinner("Gerando Excel…"):
            xlsx_bytes = build_excel(df)
            date_str   = datetime.today().strftime("%d-%m-%Y")
            filename   = f"GEP_-_Data_Base_{date_str}.xlsx"
        st.download_button(
            label=f"⬇️ Baixar {filename}",
            data=xlsx_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
else:
    st.info("Use os filtros na barra lateral e clique em **Carregar dados do Databricks**.")
