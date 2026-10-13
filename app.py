# ============================================================
# DASHBOARD FINANCIERO - CONFITERÍA ELÍS
# Modelo de Flujo de Caja Proyectado (formato BNB)
# ============================================================

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# ============================================================
# 1. CONFIGURACIÓN GENERAL DE LA PÁGINA
# ============================================================
st.set_page_config(
    page_title="Dashboard Financiero - Confitería Elís",
    page_icon="🍬",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .main-header {
        background: linear-gradient(90deg, #1f4e79 0%, #2e75b6 100%);
        padding: 20px;
        border-radius: 10px;
        color: white;
        text-align: center;
        margin-bottom: 20px;
    }
</style>
""", unsafe_allow_html=True)

# ============================================================
# 2. CARGA DEL ARCHIVO EXCEL
# ============================================================
EXCEL_FILE = "datos.xlsx"

@st.cache_data(show_spinner="Cargando datos del Excel...")
def cargar_datos_excel(ruta):
    df_datos = pd.read_excel(ruta, sheet_name="Datos", header=None)
    df_sup   = pd.read_excel(ruta, sheet_name="Supuestos", header=None)
    df_hist  = pd.read_excel(ruta, sheet_name="Flujo Historico 2024-2025", header=None)
    df_12m   = pd.read_excel(ruta, sheet_name="Flujo 12 Meses", header=None)
    df_anual = pd.read_excel(ruta, sheet_name="Flujo Anual", header=None)
    return df_datos, df_sup, df_hist, df_12m, df_anual

try:
    df_datos, df_sup, df_hist, df_12m, df_anual = cargar_datos_excel(EXCEL_FILE)
except FileNotFoundError:
    st.error(f"❌ No se encontró el archivo '{EXCEL_FILE}'. "
             f"Colócalo en la misma carpeta que app.py")
    st.stop()
except Exception as e:
    st.error(f"❌ Error al leer el Excel: {e}")
    st.stop()

# ============================================================
# 3. EXTRACCIÓN Y RECONSTRUCCIÓN DE DATOS
# ============================================================
def val(df, fila, col):
    try:
        v = df.iloc[fila, col]
        return float(v) if pd.notna(v) else 0.0
    except Exception:
        return 0.0

datos_hist = {
    "Año": ["2023", "2024", "2025"],
    "Caja y bancos":         [val(df_datos,4,1), val(df_datos,4,2), val(df_datos,4,3)],
    "Depósitos a plazo":     [val(df_datos,5,1), val(df_datos,5,2), val(df_datos,5,3)],
    "Documentos por cobrar": [val(df_datos,6,1), val(df_datos,6,2), val(df_datos,6,3)],
    "Existencias":           [val(df_datos,8,1), val(df_datos,8,2), val(df_datos,8,3)],
    "Activo fijo neto":      [val(df_datos,9,1), val(df_datos,9,2), val(df_datos,9,3)],
    "Deuda LP":              [val(df_datos,14,1),val(df_datos,14,2),val(df_datos,14,3)],
    "Patrimonio":            [val(df_datos,15,1),val(df_datos,15,2),val(df_datos,15,3)],
    "Ingresos explotación":  [val(df_datos,20,1),val(df_datos,20,2),val(df_datos,20,3)],
    "Costos explotación":    [val(df_datos,22,1),val(df_datos,22,2),val(df_datos,22,3)],
    "Gastos admin":          [val(df_datos,23,1),val(df_datos,23,2),val(df_datos,23,3)],
    "Gastos comerc.":        [val(df_datos,24,1),val(df_datos,24,2),val(df_datos,24,3)],
    "Depreciación":          [val(df_datos,25,1),val(df_datos,25,2),val(df_datos,25,3)],
    "Gastos financieros":    [val(df_datos,26,1),val(df_datos,26,2),val(df_datos,26,3)],
    "Corrección monetaria":  [val(df_datos,27,1),val(df_datos,27,2),val(df_datos,27,3)],
    "IUE":                   [val(df_datos,28,1),val(df_datos,28,2),val(df_datos,28,3)],
    "Utilidad neta":         [val(df_datos,30,1),val(df_datos,30,2),val(df_datos,30,3)],
}
df_hist_datos = pd.DataFrame(datos_hist)

ventas_2025   = val(df_sup, 3, 1)
pct_costo     = val(df_sup, 5, 1)
pct_admin     = val(df_sup, 6, 1)
pct_comercial = val(df_sup, 7, 1)
pct_corrmon   = val(df_sup, 8, 1)
dep_anual     = val(df_sup, 9, 1)
dep_mensual   = val(df_sup, 10, 1)
inv_anual     = val(df_sup, 11, 1)
inv_mensual   = val(df_sup, 12, 1)
tasa_interes  = val(df_sup, 13, 1)
deuda_inicial = val(df_sup, 14, 1)
tasa_iue      = val(df_sup, 15, 1)
caja_inicial  = val(df_sup, 16, 1)

# Valores por defecto si el Excel no trae los supuestos
if ventas_2025 == 0: ventas_2025 = 31707800
if pct_costo == 0: pct_costo = 0.35
if pct_admin == 0: pct_admin = 0.05
if pct_comercial == 0: pct_comercial = 0.03
if pct_corrmon == 0: pct_corrmon = 0.05
if dep_anual == 0: dep_anual = 496357
if inv_anual == 0: inv_anual = 496357
if tasa_interes == 0: tasa_interes = 0.12
if deuda_inicial == 0: deuda_inicial = 1034550
if tasa_iue == 0: tasa_iue = 0.25
if caja_inicial == 0: caja_inicial = 18172100

# --- Reconstrucción del FLUJO ANUAL 2026-2035 ---
tasa_crec_anual = 0.12
años = list(range(2026, 2036))
n = len(años)

ingresos = [ventas_2025 * (1 + tasa_crec_anual)]
for i in range(1, n):
    ingresos.append(ingresos[-1] * (1 + tasa_crec_anual))

flujo_anual = []
deuda_inicio = deuda_inicial
for i in range(n):
    ing = ingresos[i]
    costo = -ing * pct_costo
    admin = -ing * pct_admin
    comercial = -ing * pct_comercial
    dep = -dep_anual
    corrmon = -ing * pct_corrmon
    interes = -deuda_inicio * tasa_interes
    util_antes_imp = ing + costo + admin + comercial + dep + corrmon + interes
    iue = -max(0, util_antes_imp) * tasa_iue
    utilidad_neta = util_antes_imp + iue
    dep_suma = -dep
    corrmon_suma = -corrmon
    amort = -deuda_inicio if i == 0 else 0
    gasto_cap = -inv_anual
    flujo_caja = utilidad_neta + dep_suma + corrmon_suma + amort + gasto_cap
    deuda_fin = deuda_inicio + amort
    flujo_anual.append({
        "Año": años[i], "Ingresos": ing,
        "Costo explotación": costo, "Gasto admin": admin,
        "Gasto comercial": comercial, "Depreciación": dep,
        "Corrección monetaria": corrmon, "Costo financiero": interes,
        "IUE": iue, "Utilidad neta": utilidad_neta,
        "Amortización capital": amort, "Gastos de capital": gasto_cap,
        "Flujo de caja anual": flujo_caja,
        "Deuda inicio": deuda_inicio, "Deuda fin": deuda_fin,
    })
    deuda_inicio = deuda_fin

df_flujo_anual = pd.DataFrame(flujo_anual)
df_flujo_anual["Flujo acumulado"] = df_flujo_anual["Flujo de caja anual"].cumsum() + caja_inicial

# --- FLUJO MENSUAL 2026 ---
tasa_crec_mensual = (1 + tasa_crec_anual) ** (1/12) - 1
meses = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"]

ingresos_m = [ventas_2025 * (1 + tasa_crec_mensual)]
for i in range(1, 12):
    ingresos_m.append(ingresos_m[-1] * (1 + tasa_crec_mensual))

flujo_mensual = []
caja_acum = caja_inicial
for i in range(12):
    ing = ingresos_m[i]
    costo = -ing * pct_costo
    admin = -ing * pct_admin
    comercial = -ing * pct_comercial
    dep = -dep_mensual
    corrmon = -ing * pct_corrmon
    deuda_saldo = deuda_inicial * (1 - i/12)
    interes = -deuda_saldo * tasa_interes / 12
    util_antes = ing + costo + admin + comercial + dep + corrmon + interes
    iue = -max(0, util_antes) * tasa_iue
    util_neta = util_antes + iue
    dep_suma = -dep
    corrmon_suma = -corrmon
    amort = -deuda_inicial / 12
    gasto_cap = -inv_mensual
    flujo = util_neta + dep_suma + corrmon_suma + amort + gasto_cap
    caja_acum += flujo
    flujo_mensual.append({
        "Mes": meses[i], "Ingresos": ing,
        "Costo explotación": costo, "Gasto admin": admin,
        "Gasto comercial": comercial, "Depreciación": dep,
        "Corrección monetaria": corrmon, "Costo financiero": interes,
        "IUE": iue, "Utilidad neta": util_neta,
        "Amortización capital": amort, "Gastos de capital": gasto_cap,
        "Flujo de caja mensual": flujo, "Flujo acumulado": caja_acum,
    })

df_flujo_mensual = pd.DataFrame(flujo_mensual)

# ============================================================
# 4. ENCABEZADO PRINCIPAL
# ============================================================
st.markdown("""
<div class="main-header">
    <h1>🍬 Dashboard Financiero - Confitería Elís</h1>
    <p>Modelo de Flujo de Caja Proyectado | Formato BNB | Expresado en Bolivianos</p>
</div>
""", unsafe_allow_html=True)

# ============================================================
# 5. BARRA LATERAL - FILTROS
# ============================================================
st.sidebar.title("🎛️ Panel de Control")
st.sidebar.markdown("---")

seccion = st.sidebar.radio(
    "Selecciona una vista:",
    ["📊 Resumen Ejecutivo",
     "📈 Proyección Anual 2026-2035",
     "📅 Flujo Mensual 2026",
     "📉 Histórico 2023-2025",
     "🔍 Datos y Supuestos"]
)

st.sidebar.markdown("---")
st.sidebar.markdown("### 💡 Supuestos clave")
st.sidebar.metric("Crecimiento anual", f"{tasa_crec_anual:.1%}")
st.sidebar.metric("% Costo/Ventas", f"{pct_costo:.2%}")
st.sidebar.metric("% Admin/Ventas", f"{pct_admin:.2%}")
st.sidebar.metric("Tasa IUE", f"{tasa_iue:.0%}")

# ============================================================
# 6. VISTA 1: RESUMEN EJECUTIVO
# ============================================================
if seccion == "📊 Resumen Ejecutivo":
    st.subheader("📊 Resumen Ejecutivo")

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.metric("Ventas 2025 (real)", f"Bs {ventas_2025:,.0f}")
    with c2:
        st.metric("Ventas proyectadas 2026", f"Bs {ingresos[0]:,.0f}",
                  delta=f"{tasa_crec_anual:.0%}")
    with c3:
        total_10a = df_flujo_anual["Flujo de caja anual"].sum()
        st.metric("Flujo acumulado 10 años", f"Bs {total_10a:,.0f}")
    with c4:
        util_prom = df_flujo_anual["Utilidad neta"].mean()
        st.metric("Utilidad neta promedio", f"Bs {util_prom:,.0f}")

    st.markdown("---")

    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(go.Bar(x=df_flujo_anual["Año"], y=df_flujo_anual["Ingresos"],
                         name="Ingresos", marker_color="#2e75b6"), secondary_y=False)
    fig.add_trace(go.Bar(x=df_flujo_anual["Año"], y=df_flujo_anual["Costo explotación"],
                         name="Costo explotación", marker_color="#c00000"), secondary_y=False)
    fig.add_trace(go.Scatter(x=df_flujo_anual["Año"], y=df_flujo_anual["Utilidad neta"],
                             name="Utilidad neta", mode="lines+markers",
                             line=dict(color="#548235", width=3)), secondary_y=True)
    fig.update_layout(title="Evolución de Ingresos, Costos y Utilidad Neta (2026-2035)",
                      barmode="group", height=450,
                      legend=dict(orientation="h", y=1.1))
    fig.update_yaxes(title_text="Bs", secondary_y=False)
    fig.update_yaxes(title_text="Utilidad Neta", secondary_y=True)
    st.plotly_chart(fig, use_container_width=True)

    col1, col2 = st.columns(2)
    with col1:
        fig2 = px.pie(
            names=["Costo explotación","Gasto admin","Gasto comercial",
                   "Depreciación","Corrección monetaria","Costo financiero","IUE"],
            values=[abs(df_flujo_anual[c].sum()) for c in
                    ["Costo explotación","Gasto admin","Gasto comercial",
                     "Depreciación","Corrección monetaria","Costo financiero","IUE"]],
            title="Composición de Egresos (10 años)", hole=0.4)
        fig2.update_layout(height=400)
        st.plotly_chart(fig2, use_container_width=True)

    with col2:
        fig3 = px.area(df_flujo_anual, x="Año", y="Flujo acumulado",
                       title="Flujo de Caja Acumulado (2026-2035)", markers=True)
        fig3.update_traces(line_color="#548235", fillcolor="rgba(84,130,53,0.3)")
        fig3.update_layout(height=400)
        st.plotly_chart(fig3, use_container_width=True)

# ============================================================
# 7. VISTA 2: PROYECCIÓN ANUAL
# ============================================================
elif seccion == "📈 Proyección Anual 2026-2035":
    st.subheader("📈 Proyección Anual 2026-2035")

    años_sel = st.multiselect("Selecciona años a mostrar:",
        options=df_flujo_anual["Año"].tolist(),
        default=df_flujo_anual["Año"].tolist())
    df_fil = df_flujo_anual[df_flujo_anual["Año"].isin(años_sel)]

    st.markdown("### 📋 Estado de Flujo de Caja Proyectado")
    df_tabla = df_fil.set_index("Año").T
    st.dataframe(df_tabla.style.format("{:,.0f}"),
                 use_container_width=True, height=500)

    col1, col2 = st.columns(2)
    with col1:
        fig = px.bar(df_fil, x="Año", y="Flujo de caja anual",
                     title="Flujo de Caja Anual",
                     color="Flujo de caja anual",
                     color_continuous_scale="RdYlGn")
        st.plotly_chart(fig, use_container_width=True)
    with col2:
        fig = px.line(df_fil, x="Año", y=["Ingresos","Utilidad neta"],
                      title="Ingresos vs Utilidad Neta", markers=True)
        st.plotly_chart(fig, use_container_width=True)

    fig = go.Figure()
    for col in ["Costo explotación","Gasto admin","Gasto comercial",
                "Depreciación","Corrección monetaria","Costo financiero","IUE"]:
        fig.add_trace(go.Bar(x=df_fil["Año"], y=df_fil[col].abs(), name=col))
    fig.update_layout(barmode="stack", title="Composición de Egresos por Año", height=450)
    st.plotly_chart(fig, use_container_width=True)

# ============================================================
# 8. VISTA 3: FLUJO MENSUAL 2026
# ============================================================
elif seccion == "📅 Flujo Mensual 2026":
    st.subheader("📅 Flujo de Caja Mensual 2026")

    c1, c2, c3 = st.columns(3)
    with c1:
        st.metric("Ingresos totales 2026", f"Bs {df_flujo_mensual['Ingresos'].sum():,.0f}")
    with c2:
        st.metric("Flujo neto 2026", f"Bs {df_flujo_mensual['Flujo de caja mensual'].sum():,.0f}")
    with c3:
        st.metric("Caja final dic-2026", f"Bs {df_flujo_mensual['Flujo acumulado'].iloc[-1]:,.0f}")

    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(go.Bar(x=df_flujo_mensual["Mes"],
                         y=df_flujo_mensual["Flujo de caja mensual"],
                         name="Flujo mensual",
                         marker_color=np.where(df_flujo_mensual["Flujo de caja mensual"]>=0,
                                               "#548235","#c00000")), secondary_y=False)
    fig.add_trace(go.Scatter(x=df_flujo_mensual["Mes"],
                             y=df_flujo_mensual["Flujo acumulado"],
                             name="Flujo acumulado", mode="lines+markers",
                             line=dict(color="#2e75b6", width=3)), secondary_y=True)
    fig.update_layout(title="Flujo de Caja Mensual 2026", height=450,
                      legend=dict(orientation="h", y=1.1))
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("### 📋 Detalle mensual")
    df_m_tabla = df_flujo_mensual.set_index("Mes").T
    st.dataframe(df_m_tabla.style.format("{:,.0f}"),
                 use_container_width=True, height=500)

    fig2 = go.Figure()
    fig2.add_trace(go.Bar(x=df_flujo_mensual["Mes"],
                          y=df_flujo_mensual["Ingresos"],
                          name="Ingresos", marker_color="#2e75b6"))
    egresos_m = (df_flujo_mensual["Costo explotación"] +
                 df_flujo_mensual["Gasto admin"] +
                 df_flujo_mensual["Gasto comercial"] +
                 df_flujo_mensual["Costo financiero"] +
                 df_flujo_mensual["IUE"])
    fig2.add_trace(go.Bar(x=df_flujo_mensual["Mes"], y=egresos_m.abs(),
                          name="Egresos", marker_color="#c00000"))
    fig2.update_layout(barmode="group", title="Ingresos vs Egresos mensuales", height=400)
    st.plotly_chart(fig2, use_container_width=True)

# ============================================================
# 9. VISTA 4: HISTÓRICO
# ============================================================
elif seccion == "📉 Histórico 2023-2025":
    st.subheader("📉 Análisis Histórico 2023-2025 (datos auditados)")

    c1, c2, c3 = st.columns(3)
    with c1:
        st.metric("Ingresos 2025", f"Bs {df_hist_datos['Ingresos explotación'].iloc[-1]:,.0f}")
    with c2:
        st.metric("Utilidad 2025", f"Bs {df_hist_datos['Utilidad neta'].iloc[-1]:,.0f}")
    with c3:
        st.metric("Patrimonio 2025", f"Bs {df_hist_datos['Patrimonio'].iloc[-1]:,.0f}")

    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(go.Bar(x=df_hist_datos["Año"], y=df_hist_datos["Ingresos explotación"],
                         name="Ingresos", marker_color="#2e75b6"), secondary_y=False)
    fig.add_trace(go.Scatter(x=df_hist_datos["Año"], y=df_hist_datos["Utilidad neta"],
                             name="Utilidad neta", mode="lines+markers",
                             line=dict(color="#548235", width=3)), secondary_y=True)
    fig.update_layout(title="Ingresos vs Utilidad Neta (2023-2025)", height=400,
                      legend=dict(orientation="h", y=1.1))
    st.plotly_chart(fig, use_container_width=True)

    fig2 = go.Figure()
    fig2.add_trace(go.Bar(x=df_hist_datos["Año"], y=df_hist_datos["Caja y bancos"],
                          name="Caja y bancos", marker_color="#548235"))
    fig2.add_trace(go.Bar(x=df_hist_datos["Año"], y=df_hist_datos["Deuda LP"],
                          name="Deuda LP", marker_color="#c00000"))
    fig2.update_layout(barmode="group", title="Caja vs Deuda LP", height=400)
    st.plotly_chart(fig2, use_container_width=True)

    st.markdown("### 📋 Datos históricos completos")
    st.dataframe(df_hist_datos.set_index("Año").T.style.format("{:,.0f}"),
                 use_container_width=True)

# ============================================================
# 10. VISTA 5: DATOS Y SUPUESTOS
# ============================================================
elif seccion == "🔍 Datos y Supuestos":
    st.subheader("🔍 Datos y Supuestos del Modelo")

    st.markdown("### 🎯 Supuestos de la proyección")
    supuestos_df = pd.DataFrame({
        "Concepto": ["Ventas 2025 (base)","% Costo explotación / ventas",
                     "% Gastos admin / ventas","% Gastos comercialización / ventas",
                     "% Corrección monetaria / ventas","Depreciación anual",
                     "Inversiones anuales","Tasa de interés deuda",
                     "Deuda LP al 31/12/2025","Tasa IUE","Caja inicial 2026",
                     "Crecimiento anual de ingresos"],
        "Valor": [f"Bs {ventas_2025:,.0f}", f"{pct_costo:.2%}", f"{pct_admin:.2%}",
                  f"{pct_comercial:.2%}", f"{pct_corrmon:.2%}", f"Bs {dep_anual:,.0f}",
                  f"Bs {inv_anual:,.0f}", f"{tasa_interes:.2%}", f"Bs {deuda_inicial:,.0f}",
                  f"{tasa_iue:.0%}", f"Bs {caja_inicial:,.0f}", f"{tasa_crec_anual:.0%}"]
    })
    st.dataframe(supuestos_df, use_container_width=True, hide_index=True)

    st.markdown("### 📊 Estados financieros históricos (Datos)")
    st.dataframe(df_hist_datos, use_container_width=True, hide_index=True)

    st.markdown("### 📥 Descargar datos procesados")
    csv = df_flujo_anual.to_csv(index=False).encode("utf-8")
    st.download_button("📥 Descargar Flujo Anual (CSV)",
                       data=csv, file_name="flujo_anual_elis.csv", mime="text/csv")
    csv2 = df_flujo_mensual.to_csv(index=False).encode("utf-8")
    st.download_button("📥 Descargar Flujo Mensual (CSV)",
                       data=csv2, file_name="flujo_mensual_elis.csv", mime="text/csv")

# ============================================================
# 11. PIE DE PÁGINA
# ============================================================
st.markdown("---")
st.caption("📌 Dashboard generado a partir del modelo financiero Confitería Elís | "
           "Formato BNB | Valores expresados en Bolivianos (Bs)")
