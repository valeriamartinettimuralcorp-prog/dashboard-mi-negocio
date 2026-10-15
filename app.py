import io
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="Mi Negocio | Dashboard Comercial", page_icon="📊", layout="wide")

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

    # Excluir filas de totales/resúmenes para evitar doble conteo.
    vendedor = df["Vendedor Codigo-Nombre"].astype(str).str.strip().str.lower()
    cliente = df["Cliente Codigo Nombre"].astype(str).str.strip().str.lower()
    df = df[~vendedor.eq("total") & ~cliente.eq("total")].copy()

    for col in NUMERIC:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    df["Vendedor Codigo-Nombre"] = df["Vendedor Codigo-Nombre"].fillna("Sin vendedor").astype(str).str.strip()
    df["Cliente Codigo Nombre"] = df["Cliente Codigo Nombre"].fillna("Sin nombre").astype(str).str.strip()
    return df


def flag_count(df: pd.DataFrame, col: str) -> int:
    # Cada fila representa un cliente; cualquier valor > 0 cuenta como alcanzado.
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
    uploaded = st.file_uploader("Cargar Excel actualizado", type=["xlsx", "xls"], help="Podés cargar el archivo nuevo cada día sin cambiar el dashboard.")
    st.markdown("<div class='small-note'>El archivo debe conservar los nombres de columnas del reporte actual.</div>", unsafe_allow_html=True)

if uploaded is None:
    st.info("👈 Cargá el Excel actualizado desde el panel lateral para visualizar el dashboard.")
    st.stop()

try:
    df = load_excel(uploaded.getvalue())
except Exception as e:
    st.error(f"No pude procesar el archivo: {e}")
    st.stop()

vendedores = sorted([v for v in df["Vendedor Codigo-Nombre"].dropna().unique() if v and v.lower() != "nan"])
with st.sidebar:
    vendedor_sel = st.selectbox("Vendedor", ["Todos"] + vendedores)
    buscar = st.text_input("Buscar cliente", placeholder="Código o nombre")

filtered = df if vendedor_sel == "Todos" else df[df["Vendedor Codigo-Nombre"] == vendedor_sel]
if buscar.strip():
    filtered = filtered[filtered["Cliente Codigo Nombre"].str.contains(buscar.strip(), case=False, na=False)]

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

# Embudo y venta digital
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

# Estados de clientes
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

st.caption("Dashboard preparado para reemplazar el Excel diariamente. Próximo paso posible: guardar históricos diarios y automatizar la lectura desde Gescom/GescomWeb.")
