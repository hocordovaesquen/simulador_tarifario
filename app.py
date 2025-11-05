# app.py
# =====================================================================
# Simulador Tarifario RV – Reconciliado (v5 robusto)
# - Real y Proyectado se leen EXACTOS desde "A.3 Negociación" con heurísticas
#   y múltiples rutas de fallback. Nunca lanza KeyError: garantiza columnas.
# - Customer Journey: referencia (opcional).
# - Simulación what‑if usando "1. Parametros" (columna R+) + BBDD Neg.
# - Panel de diagnóstico para ver qué detectó en tu Excel.
# Requisitos: streamlit pandas numpy openpyxl plotly
# =====================================================================

import math
import numpy as np
import pandas as pd
import streamlit as st
import plotly.graph_objects as go
from io import BytesIO

st.set_page_config(page_title="Simulador Tarifario RV – Reconciliado (v5)", page_icon="📊", layout="wide")

# -----------------------------
# Utils
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

def normalize_country(x):
    if isinstance(x, str):
        s = x.strip().lower()
        if s in ("peru","pe"): return "Perú"
        if s in ("chile","cl"): return "Chile"
        if s in ("colombia","co"): return "Colombia"
    return x

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
    grid = {}
    nrows, ncols = df.shape
    cols = [col_letter(i) for i in range(ncols)]
    for i in range(nrows):
        row = {}
        for j, c in enumerate(cols):
            row[c] = df.iat[i, j]
        grid[i] = row
    return grid

def infer_rate(v):
    v = safe_float(v, 0.0)
    if v <= 0:
        return 0.0
    return v/100.0 if v > 1.5 else v

def ensure_columns(df, cols_with_defaults):
    """Garantiza columnas con defaults. Devuelve df modificado y una lista de faltantes que fueron creadas."""
    missing = []
    for name, default in cols_with_defaults.items():
        if name not in df.columns:
            df[name] = default
            missing.append(name)
    return df, missing

# -----------------------------
# Lectura
# -----------------------------
@st.cache_data(show_spinner=False)
def read_excel_all(uploaded):
    bbdd_neg   = pd.read_excel(uploaded, sheet_name="A.3 BBDD Neg", header=6, engine="openpyxl")
    a3_raw     = pd.read_excel(uploaded, sheet_name="A.3 Negociación", header=None, engine="openpyxl")
    # Customer Journey es referencial; puede no tener un formato fijo
    try:
        cj_raw = pd.read_excel(uploaded, sheet_name="6. Customer Journey", header=None, engine="openpyxl")
    except Exception:
        cj_raw = pd.DataFrame()
    params_raw = pd.read_excel(uploaded, sheet_name="1. Parametros", header=None, engine="openpyxl")
    return bbdd_neg, a3_raw, cj_raw, params_raw

# -----------------------------
# Parse "A.3 Negociación" – super robusto
# -----------------------------
def find_header_row(a3_raw):
    # Busca una fila que contenga "Corredor" o "Cliente"
    for i in range(min(80, a3_raw.shape[0])):
        row_texts = a3_raw.iloc[i].astype(str).str.lower().tolist()
        if any(("corredor" in t) or ("cliente" in t) for t in row_texts):
            return i
    return 8  # fallback clásico

def build_headers(a3_raw, h):
    # Intento A: usar filas h y h+1 (multiheader)
    upper = a3_raw.iloc[h].fillna("").astype(str).tolist()
    lower = a3_raw.iloc[h+1].fillna("").astype(str).tolist() if h+1 < a3_raw.shape[0] else [""]*a3_raw.shape[1]
    headers = []
    for u, l in zip(upper, lower):
        u, l = u.strip(), l.strip()
        if u and l and u != l:
            headers.append(f"{u}|{l}")
        else:
            headers.append(u or l)
    # Desduplicar
    seen = {}
    final = []
    for n in headers:
        if n in seen:
            seen[n] += 1
            final.append(f"{n}_{seen[n]}")
        else:
            seen[n] = 0
            final.append(n)
    return final

def smart_find_col(df, must_include_words):
    words = [w.lower() for w in must_include_words]
    # 1) Búsqueda exacta por todas las palabras
    for c in df.columns:
        text = str(c).lower()
        if all(w in text for w in words):
            return c
    # 2) Búsqueda por variantes (singular/plural/acentos simples)
    repl = {"ó":"o","á":"a","é":"e","í":"i","ú":"u"}
    words2 = ["".join(repl.get(ch, ch) for ch in w) for w in words]
    for c in df.columns:
        text = "".join(repl.get(ch, ch) for ch in str(c).lower())
        if all(w in text for w in words2):
            return c
    return None

def parse_a3(a3_raw):
    h = find_header_row(a3_raw)
    names = build_headers(a3_raw, h)
    data = a3_raw.iloc[h+2:].copy()  # saltamos dos filas de header
    data.columns = names

    # Columnas clave
    col_cliente = smart_find_col(data, ["corredor"]) or smart_find_col(data, ["cliente"])
    col_monto   = smart_find_col(data, ["monto","neg"]) or smart_find_col(data, ["monto"])

    # Ingreso total Real/Proy
    col_real_tot = (smart_find_col(data, ["ingreso","total","real"]) or
                    smart_find_col(data, ["ingresos","total","real"]))
    col_proj_tot = (smart_find_col(data, ["ingreso","total","proy"]) or
                    smart_find_col(data, ["ingresos","total","proy"]) or
                    smart_find_col(data, ["ingreso","total","proyectado"]) or
                    smart_find_col(data, ["propuesta","ingreso"]))

    # Si no hay total, sumar componentes
    def sum_components(tag):
        # busca columnas tipo "...|Real" / "...|Proyectado"
        cols = [c for c in data.columns if ("ingreso" in str(c).lower()) and (tag in str(c).lower())]
        if cols:
            vals = data[cols].apply(pd.to_numeric, errors="coerce").fillna(0.0).sum(axis=1)
            key = f"__{tag}_total_calc__"
            data[key] = vals
            return key
        return None

    if col_real_tot is None:
        col_real_tot = sum_components("real")
    if col_proj_tot is None:
        col_proj_tot = sum_components("proy") or sum_components("proyectado")

    # Asegurar algunas columnas de trabajo
    if col_cliente is None:
        col_cliente = "__Cliente__"
        data[col_cliente] = data.iloc[:,0].astype(str)
    if col_monto is None:
        col_monto = "__MontoUSD__"
        data[col_monto] = 0.0
    if col_real_tot is None:
        col_real_tot = "__RealExcel__"
        data[col_real_tot] = 0.0
    if col_proj_tot is None:
        col_proj_tot = "__ProyExcel__"
        data[col_proj_tot] = 0.0

    # Limpieza numérica
    for c in [col_monto, col_real_tot, col_proj_tot]:
        data[c] = pd.to_numeric(data[c], errors="coerce").fillna(0.0)

    # Marcar bloques BCS/BVL/BVC
    bolsas_map = {"BCS":"Chile","BVL":"Perú","BVC":"Colombia"}
    data["Bolsa"] = None
    current = None
    for i, row in data.iterrows():
        val = str(row.get(col_cliente, "")).strip()
        if val in bolsas_map:
            current = val
            data.at[i, "Bolsa"] = val
        else:
            data.at[i, "Bolsa"] = current

    # Totales por bolsa (filas BCS/BVL/BVC)
    df_bolsas = data[data[col_cliente].isin(bolsas_map.keys())].copy()
    df_bolsas["Pais"] = df_bolsas[col_cliente].map(bolsas_map)
    df_bolsas = df_bolsas.rename(columns={col_real_tot:"Real Excel", col_proj_tot:"Proyectado Excel"})
    df_bolsas, _ = ensure_columns(df_bolsas, {"Pais":"", "Bolsa":"", "Real Excel":0.0, "Proyectado Excel":0.0})
    if "Bolsa" not in df_bolsas.columns:
        df_bolsas["Bolsa"] = df_bolsas[col_cliente]

    # Filas de clientes (excluye agregados)
    mask_cli = (~data[col_cliente].isin(bolsas_map.keys())) & data[col_cliente].notna() & (data[col_cliente].astype(str).str.strip()!="")
    clientes = data[mask_cli].copy()
    clientes["Pais"] = clientes["Bolsa"].map(bolsas_map)

    # Renombrar a estándar
    clientes = clientes.rename(columns={
        col_cliente: "Cliente",
        col_monto: "Monto USD",
        col_real_tot: "Real Excel",
        col_proj_tot: "Proyectado Excel"
    })

    # Garantizar columnas estándar (evita KeyError)
    clientes, created = ensure_columns(clientes, {
        "Pais":"",
        "Cliente":"",
        "Monto USD":0.0,
        "Real Excel":0.0,
        "Proyectado Excel":0.0
    })

    # Numéricos
    for c in ["Monto USD","Real Excel","Proyectado Excel"]:
        clientes[c] = pd.to_numeric(clientes[c], errors="coerce").fillna(0.0)

    meta = {
        "header_row": h,
        "col_cliente": col_cliente,
        "col_monto": col_monto,
        "col_real_total": col_real_tot,
        "col_proy_total": col_proj_tot,
        "created_fallback_cols": created
    }
    return clientes[["Pais","Cliente","Monto USD","Real Excel","Proyectado Excel"]], df_bolsas[["Pais","Bolsa","Real Excel","Proyectado Excel"]], meta

# -----------------------------
# Parametría – "1. Parametros"
# -----------------------------
def parse_tramos_right(grid, rows, mapping, fields=("min","max","bps","fijo")):
    out = {k: [] for k in mapping}
    for country, (cmin, cmax, cvar, cfijo) in mapping.items():
        for r in rows:
            mn = grid.get(r, {}).get(cmin); mx = grid.get(r, {}).get(cmax)
            var = grid.get(r, {}).get(cvar); fijo = grid.get(r, {}).get(cfijo)
            if any(is_num(v) for v in [mn, mx, var, fijo]):
                out[country].append({
                    "min":  safe_float(mn, 0.0),
                    "max":  safe_float(mx, float("inf")),
                    fields[2]: safe_float(var, 0.0),
                    fields[3]: safe_float(fijo, 0.0)
                })
        out[country] = sorted(out[country], key=lambda d: d["min"])
    return out

def read_params(params_raw):
    grid = to_grid(params_raw)
    mapping = {"Colombia":("T","U","V","W"), "Perú":("X","Y","Z","AA"), "Chile":("AB","AC","AD","AE")}
    trans = parse_tramos_right(grid, range(134,145), mapping, fields=("min","max","bps","fijo"))
    dma   = parse_tramos_right(grid, range(146,160), mapping, fields=("min","max","bps","fijo"))
    pant  = parse_tramos_right(grid, range(117,135), mapping, fields=("min","max","var","fija"))
    desc  = safe_float(grid.get(114,{}).get("C"), 0.15)
    return {"transaccion":trans, "dma":dma, "pantallas":pant, "desc_bcs":desc}

# -----------------------------
# Simulación
# -----------------------------
def ingreso_transaccion(monto_usd, tramos):
    if not tramos: return 0.0
    t = None
    for x in tramos:
        if monto_usd >= x["min"] and monto_usd <= x["max"]:
            t = x; break
    if t is None: t = tramos[-1]
    return monto_usd * infer_rate(t.get("bps",0.0)) + safe_float(t.get("fijo",0.0))

def ingreso_dma(monto_dma_usd, tramos_dma):
    if not tramos_dma: return 0.0
    t = None
    for x in tramos_dma:
        if monto_dma_usd >= x["min"] and monto_dma_usd <= x["max"]:
            t = x; break
    if t is None: t = tramos_dma[-1]
    return monto_dma_usd * infer_rate(t.get("bps",0.0)) + safe_float(t.get("fijo",0.0))

def ingreso_acceso_codigos(codigos, tramos, desc=0.0):
    if not tramos: return 0.0
    t = None
    for x in tramos:
        if codigos >= x["min"] and codigos <= x["max"]:
            t = x; break
    if t is None: t = tramos[-1]
    fijo = safe_float(t.get("fija",0.0)); var = safe_float(t.get("var",0.0))
    bruto = fijo + var * codigos
    return bruto * (1 - max(0.0, min(1.0, desc)))

def simulate(df_clients, bbdd_neg, params):
    base = bbdd_neg.copy()
    base["Pais"] = base["Pais"].apply(normalize_country)
    base["Codigos de Pantalla"] = pd.to_numeric(base["Codigos de Pantalla"], errors="coerce").fillna(0).astype(int)
    base["Monto DMA USD"] = pd.to_numeric(base["Monto DMA USD"], errors="coerce").fillna(0.0)
    agg = base.groupby(["Pais","Cliente estandar"]).agg({"Codigos de Pantalla":"max","Monto DMA USD":"sum"}).reset_index().rename(columns={"Cliente estandar":"Cliente"})

    df = df_clients.merge(agg, on=["Pais","Cliente"], how="left")
    df["Codigos de Pantalla"] = df["Codigos de Pantalla"].fillna(0).astype(int)
    df["Monto DMA USD"] = df["Monto DMA USD"].fillna(0.0)

    sims = []
    for _, r in df.iterrows():
        pais = r["Pais"]; monto = safe_float(r["Monto USD"],0.0)
        cods = int(r["Codigos de Pantalla"]); monto_dma = safe_float(r["Monto DMA USD"],0.0)
        tr_trx = params["transaccion"].get(pais, [])
        tr_dma = params["dma"].get(pais, [])
        tr_pant= params["pantallas"].get(pais, [])
        desc   = params.get("desc_bcs", 0.0)
        sims.append( ingreso_transaccion(monto, tr_trx) + ingreso_dma(monto_dma, tr_dma) + ingreso_acceso_codigos(cods, tr_pant, desc) )
    out = df.copy()
    out["Proyectado Sim"] = sims
    out["BPS Sim"] = out.apply(lambda r: calc_bps(r["Proyectado Sim"], r["Monto USD"]), axis=1)
    return out

# -----------------------------
# UI
# -----------------------------
st.title("📊 Simulador Tarifario RV – Reconciliado (v5)")

uploaded = st.file_uploader("📁 Sube tu Excel maestro", type=["xlsx"])
if not uploaded:
    st.info("Se requieren 'A.3 Negociación', '6. Customer Journey', 'A.3 BBDD Neg' y '1. Parametros'.")
    st.stop()

with st.spinner("Leyendo libro y reconciliando…"):
    bbdd_neg, a3_raw, cj_raw, params_raw = read_excel_all(uploaded)
    df_cli, df_bolsas, a3_meta = parse_a3(a3_raw)
    params = read_params(params_raw)

# Sidebar – edición de tramos
with st.sidebar:
    st.header("⚙️ Parámetros (Simulación)")
    desc_bcs = st.number_input("Descuento BCS (0–1)", min_value=0.0, max_value=1.0, step=0.01, value=float(params.get("desc_bcs",0.15)))

    st.subheader("Transacción – Tramos")
    trans_edit = {}
    for pais in ["Chile","Colombia","Perú"]:
        base = params["transaccion"].get(pais, [])
        df_e = pd.DataFrame(base) if base else pd.DataFrame(columns=["min","max","bps","fijo"])
        df_e_show = df_e.rename(columns={"min":"Mín (USD)","max":"Máx (USD)","bps":"Variable (fracción/%)","fijo":"Fijo (USD)"})
        df_e_show = st.data_editor(df_e_show, key=f"trx_{pais}", use_container_width=True, num_rows="dynamic")
        edited = []
        for _, row in df_e_show.iterrows():
            edited.append({"min":safe_float(row.get("Mín (USD)",0)),
                           "max":safe_float(row.get("Máx (USD)", float("inf"))),
                           "bps":safe_float(row.get("Variable (fracción/%)",0)),
                           "fijo":safe_float(row.get("Fijo (USD)",0))})
        trans_edit[pais] = edited

    st.subheader("Acceso – Códigos/Pantallas")
    pant_edit = {}
    for pais in ["Chile","Colombia","Perú"]:
        base = params["pantallas"].get(pais, [])
        df_e = pd.DataFrame(base) if base else pd.DataFrame(columns=["min","max","var","fija"])
        df_e_show = df_e.rename(columns={"min":"Mín #Códigos","max":"Máx #Códigos","var":"Variable por código (USD)","fija":"Fijo mensual tramo (USD)"})
        df_e_show = st.data_editor(df_e_show, key=f"pant_{pais}", use_container_width=True, num_rows="dynamic")
        edited = []
        for _, row in df_e_show.iterrows():
            edited.append({"min":safe_float(row.get("Mín #Códigos",0)),
                           "max":safe_float(row.get("Máx #Códigos", float("inf"))),
                           "var":safe_float(row.get("Variable por código (USD)",0)),
                           "fija":safe_float(row.get("Fijo mensual tramo (USD)",0))})
        pant_edit[pais] = edited

    st.subheader("DMA – Tramos")
    dma_edit = {}
    for pais in ["Chile","Colombia","Perú"]:
        base = params["dma"].get(pais, [])
        df_e = pd.DataFrame(base) if base else pd.DataFrame(columns=["min","max","bps","fijo"])
        df_e_show = df_e.rename(columns={"min":"Mín (USD)","max":"Máx (USD)","bps":"Variable (fracción/%)","fijo":"Fijo (USD)"})
        df_e_show = st.data_editor(df_e_show, key=f"dma_{pais}", use_container_width=True, num_rows="dynamic")
        edited = []
        for _, row in df_e_show.iterrows():
            edited.append({"min":safe_float(row.get("Mín (USD)",0)),
                           "max":safe_float(row.get("Máx (USD)", float("inf"))),
                           "bps":safe_float(row.get("Variable (fracción/%)",0)),
                           "fijo":safe_float(row.get("Fijo (USD)",0))})
        dma_edit[pais] = edited

params_live = {"desc_bcs":desc_bcs, "transaccion":trans_edit, "pantallas":pant_edit, "dma":dma_edit}

# Filtro
paises = ["Todos"] + sorted(df_cli["Pais"].dropna().unique().tolist())
col_f1, col_f2 = st.columns(2)
with col_f1:
    pais_sel = st.selectbox("Filtrar por País", paises, index=0)
with col_f2:
    ver_detalle = st.toggle("Mostrar detalle por cliente", value=False)

df_cli_f = df_cli if pais_sel=="Todos" else df_cli[df_cli["Pais"]==pais_sel]
df_sim = simulate(df_cli_f, bbdd_neg, params_live)

# KPIs
total_real_xls = df_cli_f["Real Excel"].sum()
total_proj_xls = df_cli_f["Proyectado Excel"].sum()
total_proj_sim = df_sim["Proyectado Sim"].sum()
monto_total    = df_cli_f["Monto USD"].sum()

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Ingreso Real (Excel)", f"${total_real_xls:,.0f}")
c2.metric("Ingreso Proyectado (Excel)", f"${total_proj_xls:,.0f}", delta=(f"{(total_proj_xls/total_real_xls-1)*100:+.1f}%" if total_real_xls>0 else None))
c3.metric("Ingreso Proyectado (Sim)", f"${total_proj_sim:,.0f}", delta=(f"{(total_proj_sim/total_real_xls-1)*100:+.1f}%" if total_real_xls>0 else None))
c4.metric("BPS Proyectado (Excel)", f"{calc_bps(total_proj_xls, monto_total):.2f} bps")
c5.metric("BPS Proyectado (Sim)", f"{calc_bps(total_proj_sim, monto_total):.2f} bps")

st.markdown("---")
cA, cB = st.columns(2, gap="large")
with cA:
    st.subheader("Real vs Proyectado – Excel")
    fig = go.Figure(data=[
        go.Bar(name="Real (Excel)", x=["Negociación"], y=[total_real_xls]),
        go.Bar(name="Proyectado (Excel)", x=["Negociación"], y=[total_proj_xls])
    ]); fig.update_layout(barmode="group", height=300)
    st.plotly_chart(fig, use_container_width=True)
with cB:
    st.subheader("Excel vs Sim (Proyectado)")
    fig2 = go.Figure(data=[
        go.Bar(name="Proyectado (Excel)", x=["Negociación"], y=[total_proj_xls]),
        go.Bar(name="Proyectado (Sim)", x=["Negociación"], y=[total_proj_sim])
    ]); fig2.update_layout(barmode="group", height=300)
    st.plotly_chart(fig2, use_container_width=True)

# Tabla detalle
st.markdown("### 📋 Detalle por Cliente (Excel + Sim)")
tabla = df_sim[["Pais","Cliente","Monto USD","Real Excel","Proyectado Excel","Proyectado Sim","BPS Sim"]].copy()
tabla["Δ Sim vs Excel (USD)"] = tabla["Proyectado Sim"] - tabla["Proyectado Excel"]
tabla["Δ Sim vs Excel (%)"] = np.where(tabla["Proyectado Excel"]>0, 100*tabla["Δ Sim vs Excel (USD)"]/tabla["Proyectado Excel"], 0.0)
if ver_detalle:
    st.dataframe(tabla.sort_values(["Pais","Δ Sim vs Excel (USD)"], ascending=[True, False]), use_container_width=True, height=420)

# Reconciliación por Bolsa
st.markdown("### 🧮 Reconciliación por Bolsa (Excel)")
bol = df_bolsas.copy()
st.dataframe(bol.rename(columns={"Real Excel":"Real (Excel)","Proyectado Excel":"Proyectado (Excel)"}), use_container_width=True)

# Alerta si descuadra suma clientes vs fila bolsa
toler = 1.0
for pais, bolsa in [("Chile","BCS"),("Perú","BVL"),("Colombia","BVC")]:
    suma_clientes = df_cli[df_cli["Pais"]==pais]["Proyectado Excel"].sum()
    fila_bolsa = bol[bol["Bolsa"]==bolsa]["Proyectado Excel"].sum() if "Bolsa" in bol.columns else 0.0
    if abs(suma_clientes - fila_bolsa) > toler:
        st.warning(f"{pais}: Suma clientes Proy (Excel) ${suma_clientes:,.0f} ≠ fila {bolsa} ${fila_bolsa:,.0f}")

# Diagnóstico
with st.expander("🔎 Diagnóstico de parsing (A.3 Negociación)"):
    st.json(a3_meta)
    st.caption("Primeras 10 columnas detectadas en A.3 Negociación:")
    st.write(pd.DataFrame({"col": list(a3_raw.columns[:10].astype(str))}))

# Descargas
st.markdown("---")
d1, d2, d3 = st.columns(3)
with d1:
    st.download_button("📥 Descargar Detalle (CSV)", tabla.to_csv(index=False).encode("utf-8"), "detalle_excel_vs_sim.csv", "text/csv", use_container_width=True)
with d2:
    st.download_button("📥 Descargar Totales por Bolsa (CSV)", bol.to_csv(index=False).encode("utf-8"), "totales_bolsa_excel.csv", "text/csv", use_container_width=True)
with d3:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        tabla.to_excel(writer, index=False, sheet_name="Detalle")
        bol.to_excel(writer, index=False, sheet_name="TotalesBolsa")
    st.download_button("📥 Descargar Excel (Detalle+Resúmenes)", buffer.getvalue(), "reconciliado_resultados.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
