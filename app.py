import io
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st
from supabase import create_client

st.set_page_config(page_title="Mi Negocio | Dashboard Comercial", page_icon="📊", layout="wide")

EXCLUDED_SELLER_CODES = {"11", "98", "1000"}
DATA_DIR = Path("data")
SAVED_EXCEL = DATA_DIR / "ultimo_reporte.xlsx"
SUPABASE_BUCKET = "mi-negocio"
SUPABASE_FILE = "ultimo_reporte.xlsx"

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
    return df


def get_supabase_client():
    try:
        url = st.secrets["SUPABASE_URL"]
        key = st.secrets["SUPABASE_KEY"]
        return create_client(url, key)
    except Exception:
        return None


def save_local_excel(file_bytes: bytes) -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    SAVED_EXCEL.write_bytes(file_bytes)


def save_excel(file_bytes: bytes) -> str:
    client = get_supabase_client()
    if client is not None:
        try:
            client.storage.from_(SUPABASE_BUCKET).upload(
                SUPABASE_FILE,
                file_bytes,
                file_options={
                    "content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "upsert": "true",
                },
            )
            save_local_excel(file_bytes)
            return "Supabase Storage"
        except Exception as e:
            st.sidebar.warning(f"No se pudo guardar en Supabase: {e}")

    save_local_excel(file_bytes)
    return "almacenamiento local"


def load_from_supabase() -> bytes | None:
    client = get_supabase_client()
    if client is None:
        return None
    try:
        return client.storage.from_(SUPABASE_BUCKET).download(SUPABASE_FILE)
    except Exception:
        return None


def get_active_excel(uploaded) -> tuple[bytes | None, str | None, str | None]:
    if uploaded is not None:
        file_bytes = uploaded.getvalue()
        source = save_excel(file_bytes)
        st.session_state["excel_bytes"] = file_bytes
        st.session_state["excel_name"] = uploaded.name
        st.session_state["excel_source"] = source
        return file_bytes, uploaded.name, source

    if "excel_bytes" in st.session_state:
        return (
            st.session_state["excel_bytes"],
            st.session_state.get("excel_name", SUPABASE_FILE),
            st.session_state.get("excel_source", "sesión actual"),
        )

    supabase_bytes = load_from_supabase()
    if supabase_bytes:
        save_local_excel(supabase_bytes)
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
    return pd.DataFrame(rows).sort_values(["% Registrados", "Registrados"], ascending=[False, False])


def fmt_money(v: float) -> str:
    return "$ " + f"{v:,.0f}".replace(",", ".")


def to_excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, data in sheets.items():
            data.to_excel(writer, sheet_name=name[:31], index=False)
    return output.getvalue()


def estado(valor: float) -> str:
    return "✅ Sí" if valor > 0 else "❌ No"


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
    st.header("Carga y filtros")
    uploaded = st.file_uploader(
        "Cargar o reemplazar Excel",
        type=["xlsx", "xls"],
        help="El último archivo se guarda en Supabase y queda activo hasta que cargues uno nuevo.",
    )
    st.markdown("<div class='small-note'>El archivo debe conservar los nombres de columnas del reporte actual.</div>", unsafe_allow_html=True)

file_bytes, active_name, active_source = get_active_excel(uploaded)

with st.sidebar:
    if file_bytes is not None:
        st.success(f"Archivo activo: {active_name}")
        st.caption(f"Origen: {active_source}")
        st.caption("Se seguirá usando este archivo hasta que cargues uno nuevo.")
    elif get_supabase_client() is None:
        st.info("Para persistencia permanente configurá SUPABASE_URL y SUPABASE_KEY en Streamlit Secrets.")

if file_bytes is None:
    st.info("👈 Cargá un Excel desde el panel lateral. Luego quedará guardado hasta que sea reemplazado.")
    st.stop()

try:
    df = load_excel(file_bytes)
except Exception as e:
    st.error(f"No pude procesar el archivo: {e}")
    st.stop()

vendedores = sorted([v for v in df["Vendedor Codigo-Nombre"].dropna().unique() if v and v.lower() != "nan"])
with st.sidebar:
    vendedor_sel = st.selectbox("Vendedor", ["Todos"] + vendedores)
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
                "Cliente Codigo Nombre",
                "Vendedor Codigo-Nombre",
                "Clientes MN Registrados",
                "Clientes Mi Negocio Activos",
                "Clientes MN Fully Digital",
                "Sell Out ($)",
                "Sell Out ($) Mi Negocio",
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
                column_config={
                    "Venta acumulada": st.column_config.NumberColumn(format="$ %.0f"),
                    "Venta digital": st.column_config.NumberColumn(format="$ %.0f"),
                },
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
    ventas = pd.DataFrame({"Tipo": ["Mi Negocio", "Resto"], "Venta": [m["venta_digital"], max(m["venta_total"]-m["venta_digital"], 0)]})
    fig2 = px.pie(ventas, names="Tipo", values="Venta", hole=.62, title="Participación de venta digital")
    fig2.update_layout(margin=dict(l=20, r=20, t=55, b=20), height=360)
    st.plotly_chart(fig2, use_container_width=True)

sin_reg = filtered[(filtered["Universo"] > 0) & (filtered["Clientes MN Registrados"] <= 0)].copy()
reg_no_act = filtered[(filtered["Clientes MN Registrados"] > 0) & (filtered["Clientes Mi Negocio Activos"] <= 0)].copy()
act_no_full = filtered[(filtered["Clientes Mi Negocio Activos"] > 0) & (filtered["Clientes MN Fully Digital"] <= 0)].copy()

show_cols = ["Vendedor Codigo-Nombre", "Cliente Codigo Nombre", "Sell Out ($)", "Sell Out ($) Mi Negocio"]
rename = {"Vendedor Codigo-Nombre":"Vendedor", "Cliente Codigo Nombre":"Cliente", "Sell Out ($)":"Venta acumulada", "Sell Out ($) Mi Negocio":"Venta digital"}

st.subheader("🎯 Oportunidades por cliente")
t1, t2, t3 = st.tabs([
    f"Sin registrar ({len(sin_reg)})",
    f"Registrados no activos ({len(reg_no_act)})",
    f"Activos no Full Digital ({len(act_no_full)})",
])
for tab, data in [(t1, sin_reg), (t2, reg_no_act), (t3, act_no_full)]:
    with tab:
        view = data[show_cols].rename(columns=rename).copy()
        st.dataframe(view, use_container_width=True, hide_index=True, column_config={
            "Venta acumulada": st.column_config.NumberColumn(format="$ %.0f"),
            "Venta digital": st.column_config.NumberColumn(format="$ %.0f"),
        })

if vendedor_sel == "Todos":
    st.subheader("🏆 Ranking por vendedor")
    ranking = vendedor_summary(df)
    st.dataframe(ranking, use_container_width=True, hide_index=True, column_config={
        "% Registrados": st.column_config.ProgressColumn(min_value=0, max_value=100, format="%.1f%%"),
        "% Venta digital": st.column_config.ProgressColumn(min_value=0, max_value=100, format="%.1f%%"),
    })

    chart_df = ranking.sort_values("% Registrados", ascending=True)
    fig3 = px.bar(chart_df, x="% Registrados", y="Vendedor", orientation="h", title="Cobertura de registro por vendedor", text_auto=".1f")
    fig3.update_layout(height=max(380, len(chart_df)*34), xaxis_range=[0, 100], margin=dict(l=20,r=20,t=55,b=20))
    st.plotly_chart(fig3, use_container_width=True)
else:
    ranking = vendedor_summary(filtered)

st.divider()
st.subheader("⬇️ Exportar seguimiento")
export = to_excel_bytes({
    "Sin registrar": sin_reg[show_cols].rename(columns=rename),
    "Registrados no activos": reg_no_act[show_cols].rename(columns=rename),
    "Activos no Full Digital": act_no_full[show_cols].rename(columns=rename),
    "Ranking vendedores": vendedor_summary(df),
})
st.download_button("Descargar oportunidades en Excel", export, file_name="seguimiento_mi_negocio.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

with st.expander("Ver datos procesados"):
    st.dataframe(filtered, use_container_width=True, hide_index=True)

st.caption("El último Excel cargado se guarda en Supabase Storage y queda activo hasta ser reemplazado. Si Supabase no está configurado, la app usa almacenamiento local como respaldo.")
