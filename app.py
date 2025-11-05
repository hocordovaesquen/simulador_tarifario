# app.py
# ============================================================
# Simulador Tarifario RV (multi-país)
# Lee la parametría desde la hoja "1. Parametros" (bloque derecho, columna R+)
# y replica la lógica de Proyectado de "A.3 Negociación" para compararla con lo Real.
# ============================================================
# Requisitos:
#   pip install streamlit pandas numpy openpyxl plotly
#
# Ejecutar:
#   streamlit run app.py
# ============================================================

import math
import numpy as np
import pandas as pd
import streamlit as st
import plotly.graph_objects as go
from io import BytesIO

st.set_page_config(page_title="Simulador Tarifario RV", page_icon="📈", layout="wide")

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
    - Si v > 1.5 asumimos que está expresado en 'porcentaje' (%), p.ej. 0.08% => 0.08
    - Si 0 < v <= 1.5, asumimos que es fracción (0.0008)
    Devuelve la fracción (p.ej. 0.0008).
    """
    if v <= 0:
        return 0.0
    if v > 1.5:  # probablemente %
        return v / 100.0
    return v  # fracción

def calc_bps(ingreso, monto):
    if monto <= 0:
        return 0.0
    return (ingreso / monto) * 10000.0

def ensure_cols(df, cols):
    for c in cols:
        if c not in df.columns:
            df[c] = 0.0
    return df

def normalize_country(x):
    if isinstance(x, str):
        s = x.strip().lower()
        if s == "peru" or s == "pe":
            return "Perú"
        if s == "chile" or s == "cl":
            return "Chile"
        if s == "colombia" or s == "co":
            return "Colombia"
    return x

# -----------------------------
# Lectura del Excel
# -----------------------------
@st.cache_data(show_spinner=False)
def read_excel_book(uploaded):
    # Nota: no usamos pd.ExcelFile para evitar problemas de engine en algunos entornos de Streamlit Cloud
    param_raw = pd.read_excel(uploaded, sheet_name="1. Parametros", header=None, engine="openpyxl")
    bbdd_neg   = pd.read_excel(uploaded, sheet_name="A.3 BBDD Neg", header=6, engine="openpyxl")
    return param_raw, bbdd_neg

# -----------------------------
# Parsers de "1. Parametros" (bloque derecha, desde columna R)
# -----------------------------
def parse_descuento_bcs(grid):
    # Heurística: Fila 114, col C suele contener el descuento BCS (ej. 0.15). Si no, buscar "Descuento" cerca.
    val = grid.get(114, {}).get("C")
    if not is_num(val):
        # búsqueda heurística en la columna C
        for r in range(100, 130):
            cell = grid.get(r, {}).get("C")
            if isinstance(cell, str) and "descuento" in cell.replace("ó","o").replace("í","i").lower():
                # probar a la derecha (D)
                cand = safe_float(grid.get(r,{}).get("D"), 0.0)
                if cand > 0:
                    return cand
        return 0.15
    return safe_float(val, 0.15)

def parse_transaccion_tramos_right(grid):
    """
    Tramos de 'Transacción' en bloque derecho (ej. filas 138–141).
    Column mapping esperado por país:
      Colombia: T,U,V,W
      Perú:     X,Y,Z,AA
      Chile:    AB,AC,AD,AE
    Campos: min, max, bps, fijo
    """
    mapping = {
        "Colombia": ("T","U","V","W"),
        "Perú":     ("X","Y","Z","AA"),
        "Chile":    ("AB","AC","AD","AE"),
    }
    out = {k: [] for k in mapping}
    # buscamos 4 filas consecutivas con números en columna T/X/AB
    candidate_rows = list(range(130, 160))
    for country, (cmin, cmax, cbps, cfijo) in mapping.items():
        rows = []
        for r in candidate_rows:
            mn = grid.get(r, {}).get(cmin)
            mx = grid.get(r, {}).get(cmax)
            bps = grid.get(r, {}).get(cbps)
            fijo = grid.get(r, {}).get(cfijo)
            if any(is_num(v) for v in [mn, mx, bps, fijo]):
                rows.append(r)
        # Tomar los primeros 4 que tengan min/bps
        picked = 0
        for r in rows:
            if picked >= 4:
                break
            mn = grid.get(r, {}).get(cmin)
            bps = grid.get(r, {}).get(cbps)
            if is_num(mn) and is_num(bps):
                out[country].append({
                    "min": safe_float(mn),
                    "max": safe_float(grid.get(r, {}).get(cmax), float("inf")),
                    "bps": safe_float(bps, 0.0),
                    "fijo": safe_float(grid.get(r, {}).get(cfijo), 0.0)
                })
                picked += 1
        # Asegurar orden por min
        out[country] = sorted(out[country], key=lambda d: d["min"])
    return out

def parse_pantallas_categorias_right(grid):
    """
    Bloque "Códigos/Terminales" (aprox. filas 117–129) a la derecha.
    Detecta categorías escaneando la col C por etiquetas y levanta tramos en R+ por país.
    Campos por tramo: min, max, var, fija  (var para soportar modelos mixtos; fija es tarifa mensual por tramo).
    """
    # Detectar filas de categorías (columna C con texto no vacío)
    cats = []
    for r in range(110, 135):
        label = grid.get(r, {}).get("C")
        if isinstance(label, str) and len(label.strip()) > 0:
            cats.append((r, " ".join(label.split())))
    # Deduplicar conservando orden
    seen = set()
    cats = [(r, c) for (r, c) in cats if not (c in seen or seen.add(c))]
    cats.sort()

    mapping = {
        "Colombia": ("T","U","V","W"),
        "Perú":     ("X","Y","Z","AA"),
        "Chile":    ("AB","AC","AD","AE"),
    }
    out = {}
    for i, (start_row, cat_name) in enumerate(cats):
        end_row = cats[i+1][0] if i+1 < len(cats) else 135
        out[cat_name] = {k: [] for k in mapping}
        for r in range(start_row, end_row):
            for country, (cmin, cmax, cvar, cfija) in mapping.items():
                mn = grid.get(r, {}).get(cmin)
                mx = grid.get(r, {}).get(cmax)
                var = grid.get(r, {}).get(cvar)
                fija = grid.get(r, {}).get(cfija)
                if any(is_num(v) for v in [mn, mx, var, fija]):
                    out[cat_name][country].append({
                        "min":  safe_float(mn, 0.0),
                        "max":  safe_float(mx, float("inf")),
                        "var":  safe_float(var, 0.0),   # mensual por código si aplica
                        "fija": safe_float(fija, 0.0)   # mensual por tramo
                    })
        # ordenar por min
        for country in out[cat_name]:
            out[cat_name][country] = sorted(out[cat_name][country], key=lambda d: d["min"])
    return out

def parse_dma_tramos_right(grid):
    """
    DMA en bloque derecho (aprox. filas 149–151). Estructura similar a Transacción.
    Campos: min, max, bps, fijo
    """
    mapping = {
        "Colombia": ("T","U","V","W"),
        "Perú":     ("X","Y","Z","AA"),
        "Chile":    ("AB","AC","AD","AE"),
    }
    out = {k: [] for k in mapping}
    for country, (cmin, cmax, cbps, cfijo) in mapping.items():
        for r in range(145, 160):
            mn = grid.get(r, {}).get(cmin)
            mx = grid.get(r, {}).get(cmax)
            bps = grid.get(r, {}).get(cbps)
            fijo = grid.get(r, {}).get(cfijo)
            if any(is_num(v) for v in [mn, mx, bps, fijo]):
                out[country].append({
                    "min":  safe_float(mn, 0.0),
                    "max":  safe_float(mx, float("inf")),
                    "bps":  safe_float(bps, 0.0),
                    "fijo": safe_float(fijo, 0.0)
                })
        out[country] = sorted(out[country], key=lambda d: d["min"])
    return out

def read_parametros_right(param_raw):
    grid = to_grid(param_raw)
    params = {
        "transaccion": parse_transaccion_tramos_right(grid),
        "pantallas":   parse_pantallas_categorias_right(grid),
        "dma":         parse_dma_tramos_right(grid),
        "desc_bcs":    parse_descuento_bcs(grid)
    }
    return params

# -----------------------------
# Cálculos de simulación
# -----------------------------
def ingreso_transaccion(monto_usd, tramos):
    if not tramos:
        return 0.0
    t = pick_tramo(monto_usd, tramos)
    if t is None:
        return 0.0
    rate = infer_rate(safe_float(t.get("bps", 0.0)))
    return monto_usd * rate + safe_float(t.get("fijo", 0.0))

def ingreso_dma(monto_dma_usd, tramos_dma):
    if not tramos_dma:
        return 0.0
    t = pick_tramo(monto_dma_usd, tramos_dma)
    if t is None:
        return 0.0
    rate = infer_rate(safe_float(t.get("bps", 0.0)))
    return monto_dma_usd * rate + safe_float(t.get("fijo", 0.0))

def ingreso_acceso_codigos(codigos, tramos_categoria_pais, descuento=0.0):
    """
    Modelo flexible: ingreso = fija_por_tramo + var_por_codigo * #codigos
    Luego aplica descuento (p.ej., BCS).
    """
    if not tramos_categoria_pais:
        return 0.0
    t = pick_tramo(codigos, tramos_categoria_pais)
    if t is None:
        return 0.0
    fija = safe_float(t.get("fija", 0.0))
    var  = safe_float(t.get("var", 0.0))
    bruto = fija + var * codigos
    return bruto * (1 - max(0.0, min(1.0, descuento)))

def simular(bbdd, params, categoria_pantallas="Corredora Terminales"):
    # Normalización y columnas mínimas
    df = bbdd.copy()
    df.rename(columns=lambda c: str(c).strip(), inplace=True)
    # Columnas esperadas (si no están, se crean)
    df = ensure_cols(df, ["Monto USD","Monto DMA USD","Codigos de Pantalla","Cobro Acceso","Cobro Transacción","Cobro Perfiles","Pais","Cliente estandar"])
    # Limpieza
    df["Monto USD"] = pd.to_numeric(df["Monto USD"], errors="coerce").fillna(0.0)
    df["Monto DMA USD"] = pd.to_numeric(df["Monto DMA USD"], errors="coerce").fillna(0.0)
    df["Codigos de Pantalla"] = pd.to_numeric(df["Codigos de Pantalla"], errors="coerce").fillna(0).astype(int)
    df["Cobro Acceso"] = pd.to_numeric(df["Cobro Acceso"], errors="coerce").fillna(0.0)
    df["Cobro Transacción"] = pd.to_numeric(df["Cobro Transacción"], errors="coerce").fillna(0.0)
    df["Cobro Perfiles"] = pd.to_numeric(df["Cobro Perfiles"], errors="coerce").fillna(0.0)
    df["Pais"] = df["Pais"].apply(normalize_country)

    # Agregación por Cliente–País (Acceso se cobra a nivel mensual/cliente-país)
    agg = df.groupby(["Pais", "Cliente estandar"]).agg({
        "Monto USD": "sum",
        "Monto DMA USD": "sum",
        "Codigos de Pantalla": "max",  # dotación vigente
        "Cobro Acceso": "sum",
        "Cobro Transacción": "sum",
        "Cobro Perfiles": "sum"
    }).reset_index().rename(columns={"Cliente estandar":"Cliente"})

    resultados = []
    for _, r in agg.iterrows():
        pais = r["Pais"]
        monto = r["Monto USD"]
        monto_dma = r["Monto DMA USD"]
        cods = int(r["Codigos de Pantalla"])

        # Parámetros por país
        tr_trx = params.get("transaccion", {}).get(pais, [])
        tr_dma = params.get("dma", {}).get(pais, [])
        tr_pant = params.get("pantallas", {}).get(categoria_pantallas, {}).get(pais, [])
        desc = params.get("desc_bcs", 0.0)

        # Ingresos proyectados
        proj_trx  = ingreso_transaccion(monto, tr_trx)
        proj_dma  = ingreso_dma(monto_dma, tr_dma)
        proj_acc  = ingreso_acceso_codigos(cods, tr_pant, descuento=desc)

        bps_trx = calc_bps(proj_trx, monto)

        resultados.append({
            "Pais": pais,
            "Cliente": r["Cliente"],
            "Monto USD": monto,
            "Monto DMA USD": monto_dma,
            "Codigos Pantalla": cods,

            "Real Acceso": r["Cobro Acceso"],
            "Real Transacción": r["Cobro Transacción"],
            "Real Perfiles": r["Cobro Perfiles"],

            "Proj Acceso": proj_acc,
            "Proj Transacción": proj_trx,
            "Proj DMA": proj_dma,
            "Proj BPS Transacción": bps_trx
        })

    res = pd.DataFrame(resultados)
    res["Real Total Negociación"] = res["Real Acceso"] + res["Real Transacción"] + res["Real Perfiles"]
    res["Proj Total Negociación"] = res["Proj Acceso"] + res["Proj Transacción"] + res["Proj DMA"]
    res["Δ Total (USD)"] = res["Proj Total Negociación"] - res["Real Total Negociación"]
    res["Δ Total (%)"] = np.where(res["Real Total Negociación"]>0,
                                  100*(res["Δ Total (USD)"]/res["Real Total Negociación"]), 0.0)
    return res

# -----------------------------
# UI
# -----------------------------
st.title("📈 Simulador Tarifario RV")
st.caption("Edita tramos desde el bloque derecho de **“1. Parametros”** (equivale a mover desde **columna R** en tu Excel) y mira el impacto en Proyectado / BPS al estilo **A.3 Negociación**.")

uploaded = st.file_uploader("📁 Sube tu Excel", type=["xlsx"], help="Usa el archivo maestro con las hojas '1. Parametros' y 'A.3 BBDD Neg'")
if not uploaded:
    st.info("Sube el archivo de trabajo para comenzar.")
    st.stop()

with st.spinner("Cargando libro y parametría…"):
    param_raw, bbdd_neg = read_excel_book(uploaded)
    params = read_parametros_right(param_raw)

# -----------------------------
# Panel de edición de parámetros
# -----------------------------
with st.sidebar:
    st.header("⚙️ Parámetros (Propuesta)")
    st.markdown("Estos controles **replican** el bloque derecho de *1. Parametros*.")
    st.markdown("**Descuento BCS**")
    desc_bcs = st.number_input("Descuento BCS (0–1)", min_value=0.0, max_value=1.0, step=0.01, value=float(params.get("desc_bcs", 0.15)))

    # Guardrail opcional: objetivo de Δ consolidado
    objetivo = st.number_input("🎯 Objetivo Δ consolidado (USD)", value=0.0, step=50000.0, help="Para referencia visual; no optimiza aún.")

    st.markdown("---")
    st.subheader("Transacción – Tramos por país")
    trans_edit = {}
    for pais in ["Chile", "Colombia", "Perú"]:
        st.markdown(f"**{pais}**")
        base = params.get("transaccion", {}).get(pais, [])
        df_edit = pd.DataFrame(base) if base else pd.DataFrame(columns=["min","max","bps","fijo"])
        # Mostrar bps como columnas claras
        df_edit_show = df_edit.rename(columns={"bps":"Variable (fracción)", "fijo":"Fijo (USD)", "min":"Mín (USD)", "max":"Máx (USD)"})
        df_edit_show = st.data_editor(df_edit_show, key=f"trx_{pais}", use_container_width=True, num_rows="dynamic")
        # Volver al esquema interno
        edited = []
        for _, row in df_edit_show.iterrows():
            edited.append({
                "min": safe_float(row.get("Mín (USD)",0)),
                "max": safe_float(row.get("Máx (USD)", float("inf"))),
                "bps": safe_float(row.get("Variable (fracción)",0)),
                "fijo": safe_float(row.get("Fijo (USD)",0)),
            })
        trans_edit[pais] = edited

    st.markdown("---")
    # Descubrir categorías disponibles desde el Excel
    categorias = list(params.get("pantallas", {}).keys())
    if not categorias:
        categorias = ["Corredora Terminales","Institucional Terminales","Ruteador Terminales Full","Ruteador Terminales RV-RF","Institucional Códigos"]
    st.subheader("Acceso / Perfiles – Códigos/Pantallas")
    cat_pant = st.selectbox("Categoría a simular", categorias)
    pant_edit = {}
    for pais in ["Chile", "Colombia", "Perú"]:
        st.markdown(f"**{pais} – {cat_pant}**")
        base = params.get("pantallas", {}).get(cat_pant, {}).get(pais, [])
        df_edit = pd.DataFrame(base) if base else pd.DataFrame(columns=["min","max","var","fija"])
        df_edit_show = df_edit.rename(columns={"min":"Mín #Códigos","max":"Máx #Códigos","var":"Variable por código (USD)","fija":"Fijo mensual tramo (USD)"})
        df_edit_show = st.data_editor(df_edit_show, key=f"pant_{pais}", use_container_width=True, num_rows="dynamic")
        edited = []
        for _, row in df_edit_show.iterrows():
            edited.append({
                "min": safe_float(row.get("Mín #Códigos",0)),
                "max": safe_float(row.get("Máx #Códigos", float("inf"))),
                "var": safe_float(row.get("Variable por código (USD)",0)),
                "fija": safe_float(row.get("Fijo mensual tramo (USD)",0)),
            })
        pant_edit[pais] = edited

    st.markdown("---")
    st.subheader("DMA – Tramos por país")
    dma_edit = {}
    for pais in ["Chile", "Colombia", "Perú"]:
        st.markdown(f"**{pais}**")
        base = params.get("dma", {}).get(pais, [])
        df_edit = pd.DataFrame(base) if base else pd.DataFrame(columns=["min","max","bps","fijo"])
        df_edit_show = df_edit.rename(columns={"bps":"Variable (fracción)", "fijo":"Fijo (USD)", "min":"Mín (USD)", "max":"Máx (USD)"})
        df_edit_show = st.data_editor(df_edit_show, key=f"dma_{pais}", use_container_width=True, num_rows="dynamic")
        edited = []
        for _, row in df_edit_show.iterrows():
            edited.append({
                "min": safe_float(row.get("Mín (USD)",0)),
                "max": safe_float(row.get("Máx (USD)", float("inf"))),
                "bps": safe_float(row.get("Variable (fracción)",0)),
                "fijo": safe_float(row.get("Fijo (USD)",0)),
            })
        dma_edit[pais] = edited

# Parametría actualizada desde UI
params_live = {
    "desc_bcs": desc_bcs,
    "transaccion": trans_edit,
    "pantallas": {cat_pant: pant_edit},
    "dma": dma_edit
}

# -----------------------------
# Simulación
# -----------------------------
with st.spinner("Calculando Proyectado…"):
    # Filtros
    paises = ["Todos"] + sorted(bbdd_neg["Pais"].dropna().unique().tolist())
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        pais_sel = st.selectbox("Filtrar por País", paises, index=0)
    with col_f2:
        ver_detalle = st.toggle("Mostrar detalle por cliente", value=True)

    df_bbdd = bbdd_neg.copy()
    if pais_sel != "Todos":
        df_bbdd = df_bbdd[df_bbdd["Pais"] == pais_sel]

    resultados = simular(df_bbdd, params_live, categoria_pantallas=cat_pant)

# -----------------------------
# KPIs y visualizaciones
# -----------------------------
st.markdown("### 📊 KPIs de Negociación")
k1, k2, k3, k4 = st.columns(4)
total_real = resultados["Real Total Negociación"].sum()
total_proj = resultados["Proj Total Negociación"].sum()
delta = total_proj - total_real
delta_pct = (delta / total_real * 100) if total_real > 0 else 0.0
bps_prom_trx = calc_bps(resultados["Proj Transacción"].sum(), resultados["Monto USD"].sum())

k1.metric("Ingreso Real", f"${total_real:,.0f}")
k2.metric("Ingreso Proyectado", f"${total_proj:,.0f}", delta=f"{delta_pct:+.1f}%")
k3.metric("Δ Total", f"${delta:,.0f}")
k4.metric("BPS Transacción (prom.)", f"{bps_prom_trx:.2f} bps")

st.markdown("---")
c1, c2 = st.columns(2, gap="large")
with c1:
    st.subheader("Real vs Proyectado – Total")
    fig = go.Figure(data=[
        go.Bar(name="Real", x=["Negociación"], y=[total_real]),
        go.Bar(name="Proyectado", x=["Negociación"], y=[total_proj])
    ])
    fig.update_layout(barmode="group", height=320, showlegend=True)
    st.plotly_chart(fig, use_container_width=True)

with c2:
    st.subheader("Top 15 clientes por Δ (USD)")
    top = resultados.sort_values("Δ Total (USD)", ascending=False).head(15)
    fig2 = go.Figure(data=[
        go.Bar(x=top["Cliente"], y=top["Δ Total (USD)"])
    ])
    fig2.update_layout(height=320, xaxis_tickangle=-45)
    st.plotly_chart(fig2, use_container_width=True)

st.markdown("---")
st.subheader("Detalle por Cliente")
if ver_detalle:
    st.dataframe(
        resultados.sort_values(["Pais","Δ Total (USD)"], ascending=[True, False]),
        use_container_width=True, height=420
    )

# -----------------------------
# Resumen estilo "Customer Journey"
# -----------------------------
st.markdown("### 🧭 Resumen por País (Customer Journey – Negociación)")
resumen = resultados.groupby("Pais").agg(
    Monto_USD=("Monto USD","sum"),
    Real_Acceso=("Real Acceso","sum"),
    Real_Transaccion=("Real Transacción","sum"),
    Real_Perfiles=("Real Perfiles","sum"),
    Proj_Acceso=("Proj Acceso","sum"),
    Proj_Transaccion=("Proj Transacción","sum"),
    Proj_DMA=("Proj DMA","sum")
).reset_index()
resumen["Real Total"] = resumen["Real_Acceso"] + resumen["Real_Transaccion"] + resumen["Real_Perfiles"]
resumen["Proj Total"] = resumen["Proj_Acceso"] + resumen["Proj_Transaccion"] + resumen["Proj_DMA"]
resumen["Δ (USD)"] = resumen["Proj Total"] - resumen["Real Total"]
resumen["Δ (%)"] = np.where(resumen["Real Total"]>0, 100*resumen["Δ (USD)"]/resumen["Real Total"], 0.0)

st.dataframe(resumen, use_container_width=True)

# -----------------------------
# Descargas
# -----------------------------
st.markdown("---")
col_d1, col_d2, col_d3 = st.columns(3)
with col_d1:
    csv_detalle = resultados.to_csv(index=False).encode("utf-8")
    st.download_button("📥 Descargar Detalle (CSV)", csv_detalle, "detalle_negociacion.csv", mime="text/csv", use_container_width=True)
with col_d2:
    csv_res = resumen.to_csv(index=False).encode("utf-8")
    st.download_button("📥 Descargar Resumen País (CSV)", csv_res, "resumen_pais.csv", mime="text/csv", use_container_width=True)
with col_d3:
    # Guardar un Excel con dos hojas
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        resultados.to_excel(writer, index=False, sheet_name="Detalle")
        resumen.to_excel(writer, index=False, sheet_name="ResumenPais")
    st.download_button("📥 Descargar Excel (Detalle+Resumen)", output.getvalue(), "simulador_resultados.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
