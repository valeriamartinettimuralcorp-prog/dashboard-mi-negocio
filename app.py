import hashlib
import io
from datetime import datetime, timezone
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st
from supabase import create_client

st.set_page_config(page_title="Mi Negocio | Dashboard Comercial", page_icon="📊", layout="wide")

EXCLUDED_SELLER_CODES = {"11", "98", "1000"}
ROUTE_DAYS = ["Lunes", "Martes", "Miercoles", "Jueves", "Viernes", "Sabado", "Domingo"]

DATA_DIR = Path("data")
SAVED_EXCEL = DATA_DIR / "ultimo_reporte.xlsx"
SAVED_MASTER = DATA_DIR / "maestro_clientes.xlsx"

SUPABASE_BUCKET = "mi-negocio"
SUPABASE_FILE = "ultimo_reporte.xlsx"
SUPABASE_MASTER_FILE = "maestro_clientes.xlsx"
MASTER_TABLE = "maestro_clientes"

REQUIRED = [
    "Vendedor Codigo-Nombre", "Cliente Codigo Nombre", "Universo",
    "Clientes MN Registrados", "Clientes Mi Negocio Activos",
    "Sell Out ($)", "Sell Out ($) Mi Negocio", "Clientes MN Fully Digital",
]

NUMERIC = [
    "Universo", "Clientes MN", "Clientes MN Registrados", "Clientes Mi Negocio Activos",
    "Sell Out ($)", "Sell Out ($) Mi Negocio", "vs Total Sell Out ($)",
    "Sell Out (Kg)", "Sell Out (u)", "Drop Size (u)", "Visitas Exitosas",
    "Visitas Exitosas MN", "Clientes MN Fully Digital",
]

MASTER_COLUMNS = [
    "codigo", "estado", "razon_social", "direccion", "Barrio", "telefono",
    "localidad", "vendedor", "zona", "Lunes", "Martes", "Miercoles",
    "Jueves", "Viernes", "Sabado", "Domingo", "frecuencia",
]


def clean_code(value):
    if pd.isna(value):
        return None
    text = str(value).strip()
    if not text or text.lower() == "nan":
        return None
    if text.endswith(".0"):
        text = text[:-2]
    return text


@st.cache_data(show_spinner=False)
def load_excel(file_bytes: bytes) -> pd.DataFrame:
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=0)
    df.columns = [str(c).strip() for c in df.columns]
    missing = [c for c in REQUIRED if c not in df.columns]
    if missing:
        raise ValueError("Faltan columnas obligatorias: " + ", ".join(missing))

    vendedor = df["Vendedor Codigo-Nombre"].astype(str).str.strip().str.lower()
    cliente = df["Cliente Codigo Nombre"].astype(str).str.strip().str.lower()
    df = df[~vendedor.eq("total") & ~cliente.eq("total")].copy()

    seller_code = (
        df["Vendedor Codigo-Nombre"]
        .astype(str)
        .str.strip()
        .str.extract(r"^(\d+)", expand=False)
    )
    df = df[~seller_code.isin(EXCLUDED_SELLER_CODES)].copy()

    for col in NUMERIC:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    df["Vendedor Codigo-Nombre"] = df["Vendedor Codigo-Nombre"].fillna("Sin vendedor").astype(str).str.strip()
    df["Cliente Codigo Nombre"] = df["Cliente Codigo Nombre"].fillna("Sin nombre").astype(str).str.strip()
    df["ClienteCodigo"] = pd.to_numeric(
        df["Cliente Codigo Nombre"].astype(str).str.extract(r"^\s*(\d+)", expand=False),
        errors="coerce",
    ).astype("Int64")
    df["VendedorCodigo"] = (
        df["Vendedor Codigo-Nombre"]
        .astype(str)
        .str.extract(r"^\s*(\d+)", expand=False)
        .map(clean_code)
    )
    return df


@st.cache_data(show_spinner=False)
def parse_master_excel(file_bytes: bytes) -> pd.DataFrame:
    raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=0)
    raw.columns = [str(c).strip() for c in raw.columns]
    if "codigo" not in raw.columns:
        raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=0, header=1)
        raw.columns = [str(c).strip() for c in raw.columns]

    if "codigo" not in raw.columns:
        raise ValueError("No se encontró la columna 'codigo' en el maestro.")

    for col in MASTER_COLUMNS:
        if col not in raw.columns:
            raw[col] = None

    master = raw[MASTER_COLUMNS].copy()
    master["codigo"] = pd.to_numeric(master["codigo"], errors="coerce").astype("Int64")
    master = master[master["codigo"].notna()].copy()
    master = master.drop_duplicates(subset=["codigo"], keep="last")

    master["vendedor"] = master["vendedor"].map(clean_code)
    for day in ROUTE_DAYS:
        master[day] = master[day].map(clean_code)

    return master.reset_index(drop=True)


def get_supabase_client():
    try:
        url = st.secrets["SUPABASE_URL"]
        key = st.secrets.get("SUPABASE_SERVICE_ROLE_KEY", st.secrets.get("SUPABASE_KEY"))
        if not key:
            return None
        return create_client(url, key)
    except Exception:
        return None


def save_local(path: Path, file_bytes: bytes) -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    path.write_bytes(file_bytes)


def upload_storage(path: str, file_bytes: bytes, content_type: str) -> bool:
    client = get_supabase_client()
    if client is None:
        return False
    try:
        client.storage.from_(SUPABASE_BUCKET).upload(
            path,
            file_bytes,
            file_options={"content-type": content_type, "upsert": "true"},
        )
        return True
    except Exception as e:
        st.sidebar.warning(f"No se pudo guardar {path} en Supabase Storage: {e}")
        return False


def download_storage(path: str) -> bytes | None:
    client = get_supabase_client()
    if client is None:
        return None
    try:
        return client.storage.from_(SUPABASE_BUCKET).download(path)
    except Exception:
        return None


def save_daily_excel(file_bytes: bytes) -> str:
    saved_cloud = upload_storage(
        SUPABASE_FILE,
        file_bytes,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    save_local(SAVED_EXCEL, file_bytes)
    return "Supabase Storage" if saved_cloud else "almacenamiento local"


def get_active_excel(uploaded) -> tuple[bytes | None, str | None, str | None]:
    if uploaded is not None:
        file_bytes = uploaded.getvalue()
        digest = hashlib.sha256(file_bytes).hexdigest()
        if st.session_state.get("daily_upload_hash") != digest:
            source = save_daily_excel(file_bytes)
            st.session_state["daily_upload_hash"] = digest
            st.session_state["excel_bytes"] = file_bytes
            st.session_state["excel_name"] = uploaded.name
            st.session_state["excel_source"] = source
        return (
            st.session_state["excel_bytes"],
            st.session_state["excel_name"],
            st.session_state["excel_source"],
        )

    if "excel_bytes" in st.session_state:
        return (
            st.session_state["excel_bytes"],
            st.session_state.get("excel_name", SUPABASE_FILE),
            st.session_state.get("excel_source", "sesión actual"),
        )

    supabase_bytes = download_storage(SUPABASE_FILE)
    if supabase_bytes:
        save_local(SAVED_EXCEL, supabase_bytes)
        st.session_state["excel_bytes"] = supabase_bytes
        st.session_state["excel_name"] = SUPABASE_FILE
        st.session_state["excel_source"] = "Supabase Storage"
        return supabase_bytes, SUPABASE_FILE, "Supabase Storage"

    if SAVED_EXCEL.exists():
        file_bytes = SAVED_EXCEL.read_bytes()
        st.session_state["excel_bytes"] = file_bytes
        st.session_state["excel_name"] = SAVED_EXCEL.name
        st.session_state["excel_source"] = "almacenamiento local"
        return file_bytes, SAVED_EXCEL.name, "almacenamiento local"

    return None, None, None


def master_to_records(master: pd.DataFrame) -> list[dict]:
    data = master.copy()
    data = data.rename(columns={"Barrio": "barrio"})
    now = datetime.now(timezone.utc).isoformat()
    records = []
    for row in data.itertuples(index=False):
        item = row._asdict()
        item["codigo"] = int(item["codigo"])
        item["updated_at"] = now
        clean = {}
        for key, value in item.items():
            if pd.isna(value):
                clean[key] = None
            else:
                clean[key] = str(value) if key != "codigo" else value
        records.append(clean)
    return records


def save_master_db(master: pd.DataFrame) -> bool:
    client = get_supabase_client()
    if client is None:
        return False
    try:
        records = master_to_records(master)
        for start in range(0, len(records), 500):
            client.table(MASTER_TABLE).upsert(
                records[start:start + 500],
                on_conflict="codigo",
            ).execute()
        return True
    except Exception as e:
        st.sidebar.warning(f"No se pudo guardar el maestro en la base de datos: {e}")
        return False


def load_master_db() -> pd.DataFrame | None:
    client = get_supabase_client()
    if client is None:
        return None
    try:
        rows = []
        start = 0
        page_size = 1000
        while True:
            response = (
                client.table(MASTER_TABLE)
                .select("*")
                .range(start, start + page_size - 1)
                .execute()
            )
            page = response.data or []
            rows.extend(page)
            if len(page) < page_size:
                break
            start += page_size
        if not rows:
            return None
        master = pd.DataFrame(rows)
        master = master.rename(columns={"barrio": "Barrio"})
        for col in MASTER_COLUMNS:
            if col not in master.columns:
                master[col] = None
        master = master[MASTER_COLUMNS].copy()
        master["codigo"] = pd.to_numeric(master["codigo"], errors="coerce").astype("Int64")
        master["vendedor"] = master["vendedor"].map(clean_code)
        for day in ROUTE_DAYS:
            master[day] = master[day].map(clean_code)
        return master
    except Exception:
        return None


def save_master(master_bytes: bytes, master: pd.DataFrame) -> str:
    db_ok = save_master_db(master)
    storage_ok = upload_storage(
        SUPABASE_MASTER_FILE,
        master_bytes,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    save_local(SAVED_MASTER, master_bytes)

    if db_ok and storage_ok:
        return "Supabase DB + Storage"
    if db_ok:
        return "Supabase DB"
    if storage_ok:
        return "Supabase Storage (respaldo; DB pendiente)"
    return "almacenamiento local"


def get_active_master(uploaded) -> tuple[pd.DataFrame | None, str | None]:
    if uploaded is not None:
        master_bytes = uploaded.getvalue()
        digest = hashlib.sha256(master_bytes).hexdigest()
        if st.session_state.get("master_upload_hash") != digest:
            master = parse_master_excel(master_bytes)
            source = save_master(master_bytes, master)
            st.session_state["master_upload_hash"] = digest
            st.session_state["master_df"] = master
            st.session_state["master_source"] = source
        return st.session_state["master_df"], st.session_state["master_source"]

    if "master_df" in st.session_state:
        return st.session_state["master_df"], st.session_state.get("master_source", "sesión actual")

    master_db = load_master_db()
    if master_db is not None and not master_db.empty:
        st.session_state["master_df"] = master_db
        st.session_state["master_source"] = "Supabase DB"
        return master_db, "Supabase DB"

    master_bytes = download_storage(SUPABASE_MASTER_FILE)
    if master_bytes:
        try:
            master = parse_master_excel(master_bytes)
            save_local(SAVED_MASTER, master_bytes)
            st.session_state["master_df"] = master
            st.session_state["master_source"] = "Supabase Storage"
            return master, "Supabase Storage"
        except Exception:
            pass

    if SAVED_MASTER.exists():
        try:
            master = parse_master_excel(SAVED_MASTER.read_bytes())
            st.session_state["master_df"] = master
            st.session_state["master_source"] = "almacenamiento local"
            return master, "almacenamiento local"
        except Exception:
            pass

    return None, None


def master_to_routes(master: pd.DataFrame) -> pd.DataFrame:
    cols = ["codigo", "razon_social", "direccion", "localidad", "zona"] + ROUTE_DAYS
    base = master[cols].copy()
    routes = base.melt(
        id_vars=["codigo", "razon_social", "direccion", "localidad", "zona"],
        value_vars=ROUTE_DAYS,
        var_name="DiaRuta",
        value_name="VendedorCodigo",
    )
    routes["VendedorCodigo"] = routes["VendedorCodigo"].map(clean_code)
    routes = routes[routes["VendedorCodigo"].notna()].copy()
    routes = routes[~routes["VendedorCodigo"].isin(EXCLUDED_SELLER_CODES)].copy()
    routes["ClienteCodigo"] = pd.to_numeric(routes["codigo"], errors="coerce").astype("Int64")
    routes["DiaRuta"] = pd.Categorical(routes["DiaRuta"], categories=ROUTE_DAYS, ordered=True)
    return routes


def build_opportunities(report: pd.DataFrame, master: pd.DataFrame) -> pd.DataFrame:
    routes = master_to_routes(master)
    report_cols = [
        "ClienteCodigo", "Cliente Codigo Nombre", "Vendedor Codigo-Nombre",
        "Clientes MN Registrados", "Clientes Mi Negocio Activos",
        "Clientes MN Fully Digital", "Sell Out ($)", "Sell Out ($) Mi Negocio",
    ]
    merged = routes.merge(report[report_cols], on="ClienteCodigo", how="inner")

    merged["Oportunidad Registro"] = merged["Clientes MN Registrados"] <= 0
    merged["Oportunidad Activo"] = (
        (merged["Clientes MN Registrados"] > 0)
        & (merged["Clientes Mi Negocio Activos"] <= 0)
    )
    merged["Oportunidad Full Digital"] = (
        (merged["Clientes Mi Negocio Activos"] > 0)
        & (merged["Clientes MN Fully Digital"] <= 0)
    )
    merged["Oportunidad Sell Out Digital"] = (
        (merged["Clientes MN Fully Digital"] > 0)
        & (merged["Sell Out ($) Mi Negocio"] <= 0)
    )

    def action(row):
        if row["Oportunidad Registro"]:
            return "Registrar"
        if row["Oportunidad Activo"]:
            return "Activar"
        if row["Oportunidad Full Digital"]:
            return "Full Digital"
        if row["Oportunidad Sell Out Digital"]:
            return "Generar Sell Out Digital"
        return None

    merged["Accion"] = merged.apply(action, axis=1)
    merged = merged[merged["Accion"].notna()].copy()

    seller_names = (
        report[["VendedorCodigo", "Vendedor Codigo-Nombre"]]
        .dropna()
        .drop_duplicates(subset=["VendedorCodigo"])
        .set_index("VendedorCodigo")["Vendedor Codigo-Nombre"]
        .to_dict()
    )
    merged["Vendedor"] = merged["VendedorCodigo"].map(seller_names).fillna(merged["VendedorCodigo"])
    merged["Cliente"] = merged["Cliente Codigo Nombre"].fillna(
        merged["ClienteCodigo"].astype("string") + " - " + merged["razon_social"].fillna("")
    )
    merged["Registrado"] = merged["Clientes MN Registrados"].apply(lambda v: "Sí" if v > 0 else "No")
    merged["Activo"] = merged["Clientes Mi Negocio Activos"].apply(lambda v: "Sí" if v > 0 else "No")
    merged["Full Digital"] = merged["Clientes MN Fully Digital"].apply(lambda v: "Sí" if v > 0 else "No")

    out = merged[
        [
            "VendedorCodigo", "Vendedor", "DiaRuta", "ClienteCodigo", "Cliente",
            "razon_social", "direccion", "localidad", "zona",
            "Registrado", "Activo", "Full Digital",
            "Sell Out ($)", "Sell Out ($) Mi Negocio", "Accion",
        ]
    ].copy()
    out = out.rename(columns={
        "VendedorCodigo": "Código vendedor",
        "DiaRuta": "Día de ruta",
        "ClienteCodigo": "Código cliente",
        "razon_social": "Razón social",
        "direccion": "Dirección",
        "localidad": "Localidad",
        "zona": "Zona",
        "Sell Out ($)": "Venta acumulada",
        "Sell Out ($) Mi Negocio": "Sell Out Digital",
    })
    out = out.sort_values(["Código vendedor", "Día de ruta", "Accion", "Código cliente"])
    return out


def build_route_summary(opps: pd.DataFrame) -> pd.DataFrame:
    if opps.empty:
        return pd.DataFrame(columns=[
            "Código vendedor", "Vendedor", "Día de ruta",
            "Registrar", "Activar", "Full Digital", "Sell Out Digital", "Total",
        ])

    grouped = (
        opps.groupby(["Código vendedor", "Vendedor", "Día de ruta", "Accion"], observed=False)
        .size()
        .unstack(fill_value=0)
        .reset_index()
    )
    grouped = grouped.rename(columns={"Generar Sell Out Digital": "Sell Out Digital"})
    for col in ["Registrar", "Activar", "Full Digital", "Sell Out Digital"]:
        if col not in grouped.columns:
            grouped[col] = 0
    grouped["Total"] = grouped[["Registrar", "Activar", "Full Digital", "Sell Out Digital"]].sum(axis=1)
    return grouped[
        ["Código vendedor", "Vendedor", "Día de ruta", "Registrar", "Activar", "Full Digital", "Sell Out Digital", "Total"]
    ].sort_values(["Código vendedor", "Día de ruta"])


def flag_count(df: pd.DataFrame, col: str) -> int:
    return int((df[col] > 0).sum())


def calc_metrics(df: pd.DataFrame) -> dict:
    universo = int((df["Universo"] > 0).sum())
    registrados = flag_count(df, "Clientes MN Registrados")
    activos = flag_count(df, "Clientes Mi Negocio Activos")
    full = flag_count(df, "Clientes MN Fully Digital")
    venta_total = float(df["Sell Out ($)"].sum())
    venta_digital = float(df["Sell Out ($) Mi Negocio"].sum())
    pct_digital = (venta_digital / venta_total * 100) if venta_total else 0.0
    return {
        "universo": universo,
        "registrados": registrados,
        "activos": activos,
        "full": full,
        "venta_total": venta_total,
        "venta_digital": venta_digital,
        "pct_digital": pct_digital,
        "faltan_registrar": max(universo - registrados, 0),
        "pct_registrados": (registrados / universo * 100) if universo else 0,
        "pct_activos": (activos / registrados * 100) if registrados else 0,
        "pct_full": (full / activos * 100) if activos else 0,
    }


def vendedor_summary(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for vendedor, g in df.groupby("Vendedor Codigo-Nombre", dropna=False):
        m = calc_metrics(g)
        rows.append({
            "Vendedor": vendedor,
            "Universo": m["universo"],
            "Registrados": m["registrados"],
            "% Registrados": m["pct_registrados"],
            "Faltan registrar": m["faltan_registrar"],
            "Activos": m["activos"],
            "Full Digital": m["full"],
            "% Venta digital": m["pct_digital"],
        })
    if not rows:
        return pd.DataFrame()
    return pd.DataFrame(rows).sort_values(["% Registrados", "Registrados"], ascending=[False, False])


def fmt_money(v: float) -> str:
    return "$ " + f"{v:,.0f}".replace(",", ".")


def estado(valor: float) -> str:
    return "✅ Sí" if valor > 0 else "❌ No"


def to_excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, data in sheets.items():
            data.to_excel(writer, sheet_name=name[:31], index=False)
    return output.getvalue()


def opportunities_excel(opps: pd.DataFrame) -> bytes:
    summary = build_route_summary(opps)
    sheets = {
        "Resumen Ruta": summary,
        "Detalle oportunidades": opps,
    }

    for seller_code, seller_df in opps.groupby("Código vendedor", dropna=False):
        safe_code = str(seller_code)[:20]
        sheets[f"Vendedor {safe_code}"] = seller_df.sort_values(["Día de ruta", "Accion", "Código cliente"])

    return to_excel_bytes(sheets)


st.markdown("""
<style>
.block-container {padding-top: 1.5rem; padding-bottom: 2rem;}
[data-testid="stMetric"] {background: rgba(127,127,127,.08); border: 1px solid rgba(127,127,127,.18); padding: 14px; border-radius: 14px;}
.small-note {opacity:.72; font-size:.88rem;}
</style>
""", unsafe_allow_html=True)

st.title("📊 Mi Negocio · Dashboard Comercial")
st.caption("Mural Corp · Seguimiento de registrados, activos, Full Digital y venta digital")

with st.sidebar:
    st.header("Datos")
    uploaded = st.file_uploader(
        "Cargar o reemplazar reporte Mi Negocio",
        type=["xlsx", "xls"],
        help="El último archivo queda guardado hasta que cargues uno nuevo.",
    )
    master_uploaded = st.file_uploader(
        "Cargar o actualizar maestro de clientes",
        type=["xlsx", "xls"],
        key="master_uploader",
        help="Se usa para asignar cada oportunidad al vendedor y al día de ruta.",
    )

file_bytes, active_name, active_source = get_active_excel(uploaded)

try:
    master_df, master_source = get_active_master(master_uploaded)
except Exception as e:
    master_df, master_source = None, None
    st.sidebar.error(f"No pude procesar el maestro: {e}")

with st.sidebar:
    if file_bytes is not None:
        st.success(f"Reporte activo: {active_name}")
        st.caption(f"Origen: {active_source}")
    if master_df is not None:
        st.success(f"Maestro activo: {len(master_df):,} clientes".replace(",", "."))
        st.caption(f"Origen maestro: {master_source}")
    else:
        st.warning("Falta cargar el maestro de clientes.")

if file_bytes is None:
    st.info("👈 Cargá el reporte Mi Negocio desde el panel lateral.")
    st.stop()

try:
    df = load_excel(file_bytes)
except Exception as e:
    st.error(f"No pude procesar el reporte: {e}")
    st.stop()

opportunities = None
if master_df is not None and not master_df.empty:
    opportunities = build_opportunities(df, master_df)

vendedores = sorted([v for v in df["Vendedor Codigo-Nombre"].dropna().unique() if v and v.lower() != "nan"])
with st.sidebar:
    st.header("Filtros")
    vendedor_sel = st.selectbox("Vendedor", ["Todos"] + vendedores)
    day_sel = st.selectbox("Día de ruta", ["Todos"] + ROUTE_DAYS)
    buscar = st.text_input("Buscar cliente", placeholder="Código o nombre")

base_filtrada = df if vendedor_sel == "Todos" else df[df["Vendedor Codigo-Nombre"] == vendedor_sel]
filtered = base_filtrada.copy()

if buscar.strip():
    filtered = base_filtrada[
        base_filtrada["Cliente Codigo Nombre"].str.contains(buscar.strip(), case=False, na=False)
    ].copy()

    st.subheader("🔎 Resultado de búsqueda de cliente")
    if filtered.empty:
        st.warning("No se encontraron clientes que coincidan con la búsqueda.")
    else:
        estado_clientes = filtered[
            [
                "Cliente Codigo Nombre", "Vendedor Codigo-Nombre",
                "Clientes MN Registrados", "Clientes Mi Negocio Activos",
                "Clientes MN Fully Digital", "Sell Out ($)", "Sell Out ($) Mi Negocio",
            ]
        ].copy()
        estado_clientes["Registrado"] = estado_clientes["Clientes MN Registrados"].apply(estado)
        estado_clientes["Activo"] = estado_clientes["Clientes Mi Negocio Activos"].apply(estado)
        estado_clientes["Full Digital"] = estado_clientes["Clientes MN Fully Digital"].apply(estado)
        estado_clientes = estado_clientes.rename(columns={
            "Cliente Codigo Nombre": "Cliente",
            "Vendedor Codigo-Nombre": "Vendedor",
            "Sell Out ($)": "Venta acumulada",
            "Sell Out ($) Mi Negocio": "Venta digital",
        })

        if len(estado_clientes) == 1:
            row = estado_clientes.iloc[0]
            st.markdown(f"**{row['Cliente']}** · Vendedor: {row['Vendedor']}")
            e1, e2, e3, e4, e5 = st.columns(5)
            e1.metric("Registrado", row["Registrado"])
            e2.metric("Activo", row["Activo"])
            e3.metric("Full Digital", row["Full Digital"])
            e4.metric("Venta acumulada", fmt_money(row["Venta acumulada"]))
            e5.metric("Venta digital", fmt_money(row["Venta digital"]))
        else:
            st.caption(f"Se encontraron {len(estado_clientes)} coincidencias.")
            st.dataframe(
                estado_clientes[["Cliente", "Vendedor", "Registrado", "Activo", "Full Digital", "Venta acumulada", "Venta digital"]],
                use_container_width=True,
                hide_index=True,
            )
        st.divider()

m = calc_metrics(filtered)

st.subheader("Resumen general" if vendedor_sel == "Todos" else f"Resumen · {vendedor_sel}")
c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Universo", f"{m['universo']:,}".replace(",", "."))
c2.metric("Registrados", f"{m['registrados']:,}".replace(",", "."), f"{m['pct_registrados']:.1f}% del universo")
c3.metric("Activos", f"{m['activos']:,}".replace(",", "."), f"{m['pct_activos']:.1f}% de registrados")
c4.metric("Full Digital", f"{m['full']:,}".replace(",", "."), f"{m['pct_full']:.1f}% de activos")
c5.metric("Venta digital", f"{m['pct_digital']:.2f}%", f"{fmt_money(m['venta_digital'])}")

c6, c7, c8 = st.columns(3)
c6.metric("🔴 Faltan registrar", f"{m['faltan_registrar']:,}".replace(",", "."))
c7.metric("Venta acumulada", fmt_money(m["venta_total"]))
c8.metric("Venta Mi Negocio", fmt_money(m["venta_digital"]))

st.divider()

left, right = st.columns([1, 1])
with left:
    funnel = pd.DataFrame({
        "Etapa": ["Universo", "Registrados", "Activos", "Full Digital"],
        "Clientes": [m["universo"], m["registrados"], m["activos"], m["full"]],
    })
    fig = px.funnel(funnel, x="Clientes", y="Etapa", title="Embudo de adopción digital")
    fig.update_layout(margin=dict(l=20, r=20, t=55, b=20), height=360)
    st.plotly_chart(fig, use_container_width=True)

with right:
    ventas = pd.DataFrame({
        "Tipo": ["Mi Negocio", "Resto"],
        "Venta": [m["venta_digital"], max(m["venta_total"] - m["venta_digital"], 0)],
    })
    fig2 = px.pie(ventas, names="Tipo", values="Venta", hole=.62, title="Participación de venta digital")
    fig2.update_layout(margin=dict(l=20, r=20, t=55, b=20), height=360)
    st.plotly_chart(fig2, use_container_width=True)

if opportunities is not None:
    st.subheader("🗺️ Oportunidades por vendedor y día de ruta")
    opp_view = opportunities.copy()

    if vendedor_sel != "Todos":
        seller_code = clean_code(
            pd.Series([vendedor_sel]).astype(str).str.extract(r"^\s*(\d+)", expand=False).iloc[0]
        )
        if seller_code:
            opp_view = opp_view[opp_view["Código vendedor"] == seller_code]

    if day_sel != "Todos":
        opp_view = opp_view[opp_view["Día de ruta"].astype(str) == day_sel]

    r1, r2, r3, r4, r5 = st.columns(5)
    r1.metric("Registrar", int((opp_view["Accion"] == "Registrar").sum()))
    r2.metric("Activar", int((opp_view["Accion"] == "Activar").sum()))
    r3.metric("Full Digital", int((opp_view["Accion"] == "Full Digital").sum()))
    r4.metric("Sell Out Digital", int((opp_view["Accion"] == "Generar Sell Out Digital").sum()))
    r5.metric("Total oportunidades", len(opp_view))

    if opp_view.empty:
        st.info("No hay oportunidades para los filtros seleccionados.")
    else:
        st.dataframe(
            opp_view,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Venta acumulada": st.column_config.NumberColumn(format="$ %.0f"),
                "Sell Out Digital": st.column_config.NumberColumn(format="$ %.0f"),
            },
        )

    st.subheader("Resumen por ruta")
    st.dataframe(build_route_summary(opp_view), use_container_width=True, hide_index=True)

    export_route = opportunities_excel(opportunities)
    st.download_button(
        "Descargar oportunidades por vendedor y ruta",
        export_route,
        file_name="oportunidades_por_vendedor_y_ruta.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.warning("Cargá el maestro de clientes para generar oportunidades por vendedor y día de ruta.")

sin_reg = filtered[(filtered["Universo"] > 0) & (filtered["Clientes MN Registrados"] <= 0)].copy()
reg_no_act = filtered[(filtered["Clientes MN Registrados"] > 0) & (filtered["Clientes Mi Negocio Activos"] <= 0)].copy()
act_no_full = filtered[(filtered["Clientes Mi Negocio Activos"] > 0) & (filtered["Clientes MN Fully Digital"] <= 0)].copy()

show_cols = ["Vendedor Codigo-Nombre", "Cliente Codigo Nombre", "Sell Out ($)", "Sell Out ($) Mi Negocio"]
rename = {
    "Vendedor Codigo-Nombre": "Vendedor",
    "Cliente Codigo Nombre": "Cliente",
    "Sell Out ($)": "Venta acumulada",
    "Sell Out ($) Mi Negocio": "Venta digital",
}

st.subheader("🎯 Oportunidades por estado")
t1, t2, t3 = st.tabs([
    f"Sin registrar ({len(sin_reg)})",
    f"Registrados no activos ({len(reg_no_act)})",
    f"Activos no Full Digital ({len(act_no_full)})",
])
for tab, data in [(t1, sin_reg), (t2, reg_no_act), (t3, act_no_full)]:
    with tab:
        view = data[show_cols].rename(columns=rename).copy()
        st.dataframe(view, use_container_width=True, hide_index=True)

if vendedor_sel == "Todos":
    st.subheader("🏆 Ranking por vendedor")
    ranking = vendedor_summary(df)
    if not ranking.empty:
        st.dataframe(ranking, use_container_width=True, hide_index=True)
        chart_df = ranking.sort_values("% Registrados", ascending=True)
        fig3 = px.bar(
            chart_df,
            x="% Registrados",
            y="Vendedor",
            orientation="h",
            title="Cobertura de registro por vendedor",
            text_auto=".1f",
        )
        fig3.update_layout(
            height=max(380, len(chart_df) * 34),
            xaxis_range=[0, 100],
            margin=dict(l=20, r=20, t=55, b=20),
        )
        st.plotly_chart(fig3, use_container_width=True)

with st.expander("Ver datos procesados"):
    st.dataframe(filtered, use_container_width=True, hide_index=True)

st.caption(
    "El reporte diario se guarda en Supabase Storage. "
    "El maestro de clientes se sincroniza con Supabase DB y se usa para armar las oportunidades por vendedor y día de ruta."
)
