# app.py
# =====================================================================
# Simulador Tarifario RV – Reconciliado con Excel
# Lee:
#   - "A.3 Negociación" (cabecera fila 9 aprox.) -> Real/Proyectado por cliente y por bolsa (BCS/BVL/BVC)
#   - "6. Customer Journey" -> totales ejecutivos por país/ítem
#   - "A.3 BBDD Neg" (header=6) -> base operativa (para #códigos y DMA)
#   - "1. Parametros" (bloque derecho, columna R+) -> tramos/por país (para Simulación)
#
# Muestra 3 columnas clave:
#   Real (Excel) | Proyectado (Excel) | Proyectado Sim (App)
# y una sección de Reconciliación para que SIEMPRE cuadre con el Excel.
#
# Requisitos:
#   pip install streamlit pandas numpy openpyxl plotly
# =====================================================================

import math
import numpy as np
import pandas as pd
import streamlit as st
import plotly.graph_objects as go
from io import BytesIO

st.set_page_config(page_title="Simulador Tarifario RV (Reconciliado)", page_icon="📊", layout="wide")

# -----------------------------
# Helpers
# -----------------------------
def is_num(x):
    return isinstance(x, (int, float)) and not (isinstance(x, float) and math.isnan(x))

def safe_float(x, default=0.0):
    try:
        if x is None or (isinstance(x, float) and math.isnan(x)):
            return default
        if isinstance(x, str):
            x = x.replace("$","").replace(",","").replace("\xa0","").strip()
        return float(x)
    except:
        return default

def calc_bps(ingreso, monto):
    if monto <= 0:
        return 0.0
    return (ingreso / monto) * 10000.0

def col_letter(idx):
    # 0->A
    letters = ""
    n = idx
    while True:
        n, r = divmod(n, 26)
        letters = chr(65 + r) + letters
        if n == 0:
            break
        n -= 1
    return letters

def to_grid(df):
    """Convierte DF en dict {fila: {letra_col: valor}} para indexar por A..AZ"""
    grid = {}
    nrows, ncols = df.shape
    cols = [col_letter(i) for i in range(ncols)]
    for i in range(nrows):
        row = {}
        for j, c in enumerate(cols):
            row[c] = df.iat[i, j]
        grid[i] = row
    return grid

def pick_tramo(value, tramos):
    """Devuelve el tramo cuyo rango incluye 'value'. Si no hay match, usa el último."""
    if not tramos:
        return None
    for t in tramos:
        if value >= t["min"] and value <= t["max"]:
            return t
    return tramos[-1]

def infer_rate(v):
    """
    Interpreta la 'variable' (por ejemplo bps) automáticamente:
    - Si v > 1.5 asumimos % (0.08% => 0.08)
    - Si 0 < v <= 1.5 asumimos fracción (0.0008)
    Devuelve fracción.
    """
    v = safe_float(v, 0.0)
    if v <= 0:
        return 0.0
    if v > 1.5:
        return v / 100.0
    return v

def normalize_country(x):
    if isinstance(x, str):
        s = x.strip().lower()
        if s in ("peru","pe"): return "Perú"
        if s in ("chile","cl"): return "Chile"
        if s in ("colombia","co"): return "Colombia"
    return x

# -----------------------------
# Lectura
# -----------------------------
@st.cache_data(show_spinner=False)
def read_excel_all(uploaded):
    # Cargamos todas las hojas necesarias
    bbdd_neg   = pd.read_excel(uploaded, sheet_name="A.3 BBDD Neg", header=6, engine="openpyxl")
    neg_raw    = pd.read_excel(uploaded, sheet_name="A.3 Negociación", header=None, engine="openpyxl")
    neg        = pd.read_excel(uploaded, sheet_name="A.3 Negociación", header=8, engine="openpyxl")
    cj         = pd.read_excel(uploaded, sheet_name="6. Customer Journey", header=None, engine="openpyxl")
    params_raw = pd.read_excel(uploaded, sheet_name="1. Parametros", header=None, engine="openpyxl")
    return bbdd_neg, neg_raw, neg, cj, params_raw

# -----------------------------
# Parse "A.3 Negociación" (Excel como verdad)
# -----------------------------
def parse_A3_negociacion(neg_df):
    """
    neg_df: leído con header=8.
    Estructura esperada:
      - 'Corredor' contiene filas agregadas 'BCS', 'BVL', 'BVC' y luego filas por cliente.
      - 'Monto neg', 'INGRESOS\\nReal', 'INGRESOS\\nProyectado'
    Salida:
      df_clientes  : detalle por cliente con país mapeado por bloque (BCS/BVL/BVC)
      df_bolsas    : totales por bolsa (las filas agregadas)
    """
    df = neg_df.copy()
    # Quedarnos con columnas clave, algunas vienen con saltos de línea
    cols_map = {
        'Corredor': 'Cliente',
        'PdL': 'PdL',
        'Monto neg': 'Monto USD',
        'INGRESOS\nReal': 'Real Excel',
        'INGRESOS\nProyectado': 'Proyectado Excel'
    }
    # Si nombres cambiaron levemente, intentar localizar
    for k in list(cols_map.keys()):
        if k not in df.columns:
            # buscar por startswith
            match = [c for c in df.columns if str(c).startswith(k.split("\n")[0])]
            if match:
                cols_map[match[0]] = cols_map.pop(k)
            else:
                # crear columna vacía
                df[cols_map[k]] = np.nan
                cols_map.pop(k)

    df2 = df[[c for c in cols_map.keys() if c in df.columns]].rename(columns=cols_map)
    # Identificar bloques por bolsa
    df2["Bolsa"] = None
    bolsa_current = None
    bolsas_order = []
    for i, row in df2.iterrows():
        corr = row["Cliente"]
        if isinstance(corr, str) and corr in ("BCS","BVL","BVC"):
            bolsa_current = corr
            bolsas_order.append((i, corr))
            df2.at[i, "Bolsa"] = corr
        else:
            df2.at[i, "Bolsa"] = bolsa_current

    # Filas agregadas por bolsa (= encabezados de bloque)
    df_bolsas = df2[df2["Cliente"].isin(["BCS","BVL","BVC"])].copy()
    # Mapa bolsa->país
    bolsa_to_pais = {"BCS":"Chile","BVL":"Perú","BVC":"Colombia"}
    df_bolsas["Pais"] = df_bolsas["Cliente"].map(bolsa_to_pais)

    # Filas de clientes (excluir None y las agregadas)
    df_cli = df2[~df2["Cliente"].isin(["BCS","BVL","BVC"])].copy()
    df_cli = df_cli[df_cli["Cliente"].notna()]
    df_cli["Pais"] = df_cli["Bolsa"].map(bolsa_to_pais)
    # Limpieza numérica
    for c in ["Monto USD","Real Excel","Proyectado Excel"]:
        df_cli[c] = pd.to_numeric(df_cli[c], errors="coerce").fillna(0.0)
    return df_cli, df_bolsas[["Pais","Cliente","Real Excel","Proyectado Excel"]].rename(columns={"Cliente":"Bolsa"})

# -----------------------------
# Parse "6. Customer Journey" (totales ejecutivos)
# -----------------------------
def parse_customer_journey(cj_df):
    """
    Extrae el bloque por país para 'Ingreso Actual' y 'Ingreso Propuesta'
    en Emisiones/Negociación/Cámara y TOTAL.
    """
    out = []
    # Buscar rótulos por país en col 2 o 1
    for i in range(cj_df.shape[0]):
        for j in range(1, cj_df.shape[1]):
            v = cj_df.iat[i,j]
            if isinstance(v, str) and ("BCS" in v or "BVL" in v or "BVC" in v or v.strip() in ("Chile","Perú","Colombia")):
                # intentar encontrar cabecera "Ingreso Actual / Propuesta" en fila siguiente
                # hallar columna donde dice "Ingreso Actual"
                for jj in range(j, min(j+6, cj_df.shape[1])):
                    if cj_df.iat[i+1, jj] == "Ingreso Actual":
                        col_actual = jj
                        col_prop   = jj+1
                        # recorrer las 10-15 filas siguientes con nombres de líneas
                        for ii in range(i+2, i+15):
                            name = cj_df.iat[ii, j-1] if j-1>=0 else None
                            if not isinstance(name, str):
                                continue
                            if name.strip().lower().startswith("total"):
                                total_name = name.strip()
                            out.append({
                                "Pais": v.split()[0] if " " in v else v,
                                "Concepto": name.strip(),
                                "Actual": safe_float(cj_df.iat[ii, col_actual], 0.0),
                                "Propuesta": safe_float(cj_df.iat[ii, col_prop], 0.0)
                            })
                        break
    df = pd.DataFrame(out)
    # Normalizar país
    df["Pais"] = df["Pais"].apply(lambda s: "Chile" if "Chile" in str(s) else ("Perú" if "Perú" in str(s) else ("Colombia" if "Colombia" in str(s) else s)))
    return df

# -----------------------------
# Parametría (1. Parametros) – bloque derecho (col R+)
# -----------------------------
def parse_descuento_bcs(grid):
    val = grid.get(114, {}).get("C")
    return safe_float(val, 0.15)

def parse_tramos_right(grid, rows, mapping):
    out = {k: [] for k in mapping}
    for country, (cmin, cmax, cvar, cfijo) in mapping.items():
        for r in rows:
            mn = grid.get(r, {}).get(cmin)
            mx = grid.get(r, {}).get(cmax)
            var = grid.get(r, {}).get(cvar)
            fijo = grid.get(r, {}).get(cfijo)
            if any(is_num(v) for v in [mn, mx, var, fijo]):
                out[country].append({
                    "min":  safe_float(mn, 0.0),
                    "max":  safe_float(mx, float("inf")),
                    "bps":  safe_float(var, 0.0),
                    "fijo": safe_float(fijo, 0.0)
                })
        out[country] = sorted(out[country], key=lambda d: d["min"])
    return out

def read_parametros_right(params_raw):
    grid = to_grid(params_raw)
    mapping = {
        "Colombia": ("T","U","V","W"),
        "Perú":     ("X","Y","Z","AA"),
        "Chile":    ("AB","AC","AD","AE"),
    }
    trans = parse_tramos_right(grid, rows=range(134, 145), mapping=mapping)   # ventana amplia
    dma   = parse_tramos_right(grid, rows=range(146, 160), mapping=mapping)
    # Pantallas / Códigos: usar var/fija por #códigos
    pant = {k: [] for k in mapping}
    for country, (cmin, cmax, cvar, cfija) in mapping.items():
        for r in range(117, 135):
            mn = grid.get(r, {}).get(cmin)
            mx = grid.get(r, {}).get(cmax)
            var = grid.get(r, {}).get(cvar)
            fija = grid.get(r, {}).get(cfija)
            if any(is_num(v) for v in [mn, mx, var, fija]):
                pant[country].append({
                    "min":  safe_float(mn, 0.0),
                    "max":  safe_float(mx, float("inf")),
                    "var":  safe_float(var, 0.0),
                    "fija": safe_float(fija, 0.0)
                })
        pant[country] = sorted(pant[country], key=lambda d: d["min"])
    return {
        "transaccion": trans,
        "dma": dma,
        "pantallas": pant,
        "desc_bcs": parse_descuento_bcs(grid)
    }

# -----------------------------
# Motor Simulación (What‑If)
# -----------------------------
def ingreso_transaccion(monto_usd, tramos):
    if not tramos:
        return 0.0
    t = pick_tramo(monto_usd, tramos)
    if t is None:
        return 0.0
    rate = infer_rate(t.get("bps", 0.0))
    return monto_usd * rate + safe_float(t.get("fijo", 0.0))

def ingreso_dma(monto_dma_usd, tramos_dma):
    if not tramos_dma:
        return 0.0
    t = pick_tramo(monto_dma_usd, tramos_dma)
    if t is None:
        return 0.0
    rate = infer_rate(t.get("bps", 0.0))
    return monto_dma_usd * rate + safe_float(t.get("fijo", 0.0))

def ingreso_acceso_codigos(codigos, tramos_categoria_pais, descuento=0.0):
    if not tramos_categoria_pais:
        return 0.0
    t = pick_tramo(codigos, tramos_categoria_pais)
    if t is None:
        return 0.0
    fija = safe_float(t.get("fija", 0.0))
    var  = safe_float(t.get("var", 0.0))
    bruto = fija + var * codigos
    return bruto * (1 - max(0.0, min(1.0, descuento)))

def simulate_from_params(df_clients, bbdd_neg, params):
    """
    df_clients: detalle por cliente desde A.3 Negociación (con Monto USD y País)
    bbdd_neg  : base para obtener #códigos y Monto DMA por cliente/país
    params    : tramos (transaccion/dma/pantallas) y descuento
    """
    # preparar base de #códigos y DMA por cliente/país
    base = bbdd_neg.copy()
    base["Pais"] = base["Pais"].apply(normalize_country)
    base["Codigos de Pantalla"] = pd.to_numeric(base["Codigos de Pantalla"], errors="coerce").fillna(0).astype(int)
    base["Monto DMA USD"] = pd.to_numeric(base["Monto DMA USD"], errors="coerce").fillna(0.0)
    agg = base.groupby(["Pais","Cliente estandar"]).agg({
        "Codigos de Pantalla":"max",
        "Monto DMA USD":"sum"
    }).reset_index().rename(columns={"Cliente estandar":"Cliente"})
    # merge
    df = df_clients.merge(agg, on=["Pais","Cliente"], how="left")
    df["Codigos de Pantalla"] = df["Codigos de Pantalla"].fillna(0).astype(int)
    df["Monto DMA USD"] = df["Monto DMA USD"].fillna(0.0)

    res = []
    for _, r in df.iterrows():
        pais = r["Pais"]
        monto = safe_float(r["Monto USD"], 0.0)
        cods = int(r["Codigos de Pantalla"])
        monto_dma = safe_float(r["Monto DMA USD"], 0.0)

        tr_trx = params.get("transaccion", {}).get(pais, [])
        tr_dma = params.get("dma", {}).get(pais, [])
        tr_pant = params.get("pantallas", {}).get(pais, [])
        desc = params.get("desc_bcs", 0.0)

        proj_trx  = ingreso_transaccion(monto, tr_trx)
        proj_dma  = ingreso_dma(monto_dma, tr_dma)
        proj_acc  = ingreso_acceso_codigos(cods, tr_pant, descuento=desc)

        total_proj = proj_trx + proj_dma + proj_acc
        res.append(total_proj)
    df["Proyectado Sim"] = res
    df["BPS Sim"] = df.apply(lambda r: calc_bps(r["Proyectado Sim"], r["Monto USD"]), axis=1)
    return df

# -----------------------------
# UI
# -----------------------------
st.title("📊 Simulador Tarifario RV (Reconciliado con Excel)")
st.caption("Los montos **Real** y **Proyectado** se leen **directo del Excel** (A.3 Negociación / 6. Customer Journey). "
           "La columna **Proyectado Sim** es el *what‑if* de esta app al mover tramos (columna R+) de *1. Parametros*.")

uploaded = st.file_uploader("📁 Sube tu Excel maestro", type=["xlsx"])
if not uploaded:
    st.info("Sube el archivo con las hojas: 'A.3 Negociación', '6. Customer Journey', 'A.3 BBDD Neg', '1. Parametros'.")
    st.stop()

with st.spinner("Leyendo y reconciliando…"):
    bbdd_neg, neg_raw, neg, cj, params_raw = read_excel_all(uploaded)
    df_cli, df_bolsas = parse_A3_negociacion(neg)
    df_cj = parse_customer_journey(cj)
    params = read_parametros_right(params_raw)

# Sidebar – edición de tramos
with st.sidebar:
    st.header("⚙️ Parámetros (Simulación)")
    desc_bcs = st.number_input("Descuento BCS (0–1)", min_value=0.0, max_value=1.0, step=0.01, value=float(params.get("desc_bcs",0.15)))

    st.subheader("Transacción – Tramos")
    trans_edit = {}
    for pais in ["Chile","Colombia","Perú"]:
        base = params.get("transaccion",{}).get(pais,[])
        df_e = pd.DataFrame(base) if base else pd.DataFrame(columns=["min","max","bps","fijo"])
        df_e_show = df_e.rename(columns={"min":"Mín (USD)","max":"Máx (USD)","bps":"Variable (fracción/%)","fijo":"Fijo (USD)"})
        df_e_show = st.data_editor(df_e_show, key=f"trx_{pais}", use_container_width=True, num_rows="dynamic")
        edited = []
        for _, row in df_e_show.iterrows():
            edited.append({
                "min": safe_float(row.get("Mín (USD)",0)),
                "max": safe_float(row.get("Máx (USD)", float("inf"))),
                "bps": safe_float(row.get("Variable (fracción/%)",0)),
                "fijo": safe_float(row.get("Fijo (USD)",0)),
            })
        trans_edit[pais] = edited

    st.subheader("Acceso – Códigos/Pantallas")
    pant_edit = {}
    for pais in ["Chile","Colombia","Perú"]:
        base = params.get("pantallas",{}).get(pais,[])
        df_e = pd.DataFrame(base) if base else pd.DataFrame(columns=["min","max","var","fija"])
        df_e_show = df_e.rename(columns={"min":"Mín #Códigos","max":"Máx #Códigos","var":"Variable por código (USD)","fija":"Fijo mensual tramo (USD)"})
        df_e_show = st.data_editor(df_e_show, key=f"pant_{pais}", use_container_width=True, num_rows="dynamic")
        edited = []
        for _, row in df_e_show.iterrows():
            edited.append({
                "min": safe_float(row.get("Mín #Códigos",0)),
                "max": safe_float(row.get("Máx #Códigos", float("inf"))),
                "var": safe_float(row.get("Variable por código (USD)",0)),
                "fija": safe_float(row.get("Fijo mensual tramo (USD)",0)),
            })
        pant_edit[pais] = edited

    st.subheader("DMA – Tramos")
    dma_edit = {}
    for pais in ["Chile","Colombia","Perú"]:
        base = params.get("dma",{}).get(pais,[])
        df_e = pd.DataFrame(base) if base else pd.DataFrame(columns=["min","max","bps","fijo"])
        df_e_show = df_e.rename(columns={"min":"Mín (USD)","max":"Máx (USD)","bps":"Variable (fracción/%)","fijo":"Fijo (USD)"})
        df_e_show = st.data_editor(df_e_show, key=f"dma_{pais}", use_container_width=True, num_rows="dynamic")
        edited = []
        for _, row in df_e_show.iterrows():
            edited.append({
                "min": safe_float(row.get("Mín (USD)",0)),
                "max": safe_float(row.get("Máx (USD)", float("inf"))),
                "bps": safe_float(row.get("Variable (fracción/%)",0)),
                "fijo": safe_float(row.get("Fijo (USD)",0)),
            })
        dma_edit[pais] = edited

params_live = {
    "desc_bcs": desc_bcs,
    "transaccion": trans_edit,
    "pantallas": pant_edit,
    "dma": dma_edit
}

# Filtro de país
paises = ["Todos"] + sorted(df_cli["Pais"].dropna().unique().tolist())
col_f1, col_f2 = st.columns(2)
with col_f1:
    pais_sel = st.selectbox("Filtrar por País", paises, index=0)
with col_f2:
    ver_detalle = st.toggle("Mostrar detalle por cliente", value=True)

df_cli_f = df_cli.copy() if pais_sel=="Todos" else df_cli[df_cli["Pais"]==pais_sel].copy()
df_cli_sim = simulate_from_params(df_cli_f, bbdd_neg, params_live)

# KPI – Excel (verdad) y Sim
total_real_xls = df_cli_f["Real Excel"].sum()
total_proj_xls = df_cli_f["Proyectado Excel"].sum()
total_proj_sim = df_cli_sim["Proyectado Sim"].sum()
monto_total   = df_cli_f["Monto USD"].sum()

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Ingreso Real (Excel)", f"${total_real_xls:,.0f}")
c2.metric("Ingreso Proyectado (Excel)", f"${total_proj_xls:,.0f}", delta=f"{(total_proj_xls/total_real_xls-1)*100:+.1f}%" if total_real_xls>0 else None)
c3.metric("Ingreso Proyectado (Sim)", f"${total_proj_sim:,.0f}", delta=f"{(total_proj_sim/total_real_xls-1)*100:+.1f}%" if total_real_xls>0 else None)
c4.metric("BPS Proyectado (Excel)", f"{calc_bps(total_proj_xls, monto_total):.2f} bps")
c5.metric("BPS Proyectado (Sim)", f"{calc_bps(total_proj_sim, monto_total):.2f} bps")

st.markdown("---")
col1, col2 = st.columns(2, gap="large")
with col1:
    st.subheader("Real vs Proyectado – Excel")
    fig = go.Figure(data=[
        go.Bar(name="Real (Excel)", x=["Negociación"], y=[total_real_xls]),
        go.Bar(name="Proyectado (Excel)", x=["Negociación"], y=[total_proj_xls])
    ])
    fig.update_layout(barmode="group", height=300)
    st.plotly_chart(fig, use_container_width=True)
with col2:
    st.subheader("Excel vs Sim (Proyectado)")
    fig2 = go.Figure(data=[
        go.Bar(name="Proyectado (Excel)", x=["Negociación"], y=[total_proj_xls]),
        go.Bar(name="Proyectado (Sim)", x=["Negociación"], y=[total_proj_sim])
    ])
    fig2.update_layout(barmode="group", height=300)
    st.plotly_chart(fig2, use_container_width=True)

# Tabla detalle
st.markdown("### 📋 Detalle por Cliente (Real/Proyectado de **Excel** + Proyectado **Sim**)")
tabla = df_cli_sim[["Pais","Cliente","Monto USD","Real Excel","Proyectado Excel","Proyectado Sim","BPS Sim"]].copy()
tabla["Δ Sim vs Excel (USD)"] = tabla["Proyectado Sim"] - tabla["Proyectado Excel"]
tabla["Δ Sim vs Excel (%)"]  = np.where(tabla["Proyectado Excel"]>0, 100*tabla["Δ Sim vs Excel (USD)"]/tabla["Proyectado Excel"], 0.0)
st.dataframe(tabla.sort_values(["Pais","Δ Sim vs Excel (USD)"], ascending=[True, False]), use_container_width=True, height=420)

# Resumen por bolsa (BCS/BVL/BVC) – Excel
st.markdown("### 🧮 Reconciliación por Bolsa (Excel)")
st.dataframe(df_bolsas.rename(columns={"Real Excel":"Real (Excel)","Proyectado Excel":"Proyectado (Excel)"}), use_container_width=True)

# Resumen Customer Journey (por país, Negociación)
st.markdown("### 🧭 Resumen por País desde '6. Customer Journey' (Excel)")
cj_neg = df_cj[df_cj["Concepto"].str.contains("Negociación", na=False)]
cj_pivot = cj_neg.pivot_table(index="Pais", values=["Actual","Propuesta"], aggfunc="sum").reset_index()
st.dataframe(cj_pivot.rename(columns={"Actual":"Ingreso Actual (Excel)","Propuesta":"Ingreso Propuesta (Excel)"}), use_container_width=True)

# Descargas
st.markdown("---")
col_d1, col_d2, col_d3 = st.columns(3)
with col_d1:
    csv_det = tabla.to_csv(index=False).encode("utf-8")
    st.download_button("📥 Descargar Detalle (CSV)", csv_det, "detalle_neg_proy_excel_sim.csv", "text/csv", use_container_width=True)
with col_d2:
    csv_bol = df_bolsas.to_csv(index=False).encode("utf-8")
    st.download_button("📥 Descargar Totales por Bolsa (CSV)", csv_bol, "totales_bolsa_excel.csv", "text/csv", use_container_width=True)
with col_d3:
    # Excel con hojas
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        tabla.to_excel(writer, index=False, sheet_name="Detalle_Clientes")
        df_bolsas.to_excel(writer, index=False, sheet_name="Totales_Bolsa_Excel")
        cj_pivot.to_excel(writer, index=False, sheet_name="CJ_Negociacion_Excel")
    st.download_button("📥 Descargar Excel (Detalle+Resúmenes)", output.getvalue(), "reconciliado_resultados.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
