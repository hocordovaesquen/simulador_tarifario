# app.py
# =====================================================================
# Simulador Tarifario RV – Reconciliado 100% con Excel
# - Toma "Real" y "Proyectado" EXACTOS desde la hoja "A.3 Negociación".
# - Muestra también los totales de "6. Customer Journey".
# - Permite SIMULAR (what‑if) moviendo tramos (columna R+) de "1. Parametros"
#   para Acceso (Códigos), Transacción y DMA.
# - Incluye chequeos de reconciliación (sumas por bolsa vs suma de filas).
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

st.set_page_config(page_title="Simulador Tarifario RV – Reconciliado", page_icon="📊", layout="wide")

# -----------------------------
# Utilidades
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

def infer_rate(v):
    v = safe_float(v, 0.0)
    if v <= 0:
        return 0.0
    # Si parece porcentaje (>1.5), convierto a fracción
    return v/100.0 if v > 1.5 else v

# -----------------------------
# Lectura del libro
# -----------------------------
@st.cache_data(show_spinner=False)
def read_excel_all(uploaded):
    bbdd_neg   = pd.read_excel(uploaded, sheet_name="A.3 BBDD Neg", header=6, engine="openpyxl")
    a3_raw     = pd.read_excel(uploaded, sheet_name="A.3 Negociación", header=None, engine="openpyxl")
    a3_guess   = pd.read_excel(uploaded, sheet_name="A.3 Negociación", header=8, engine="openpyxl")
    cj_raw     = pd.read_excel(uploaded, sheet_name="6. Customer Journey", header=None, engine="openpyxl")
    params_raw = pd.read_excel(uploaded, sheet_name="1. Parametros", header=None, engine="openpyxl")
    return bbdd_neg, a3_raw, a3_guess, cj_raw, params_raw

# -----------------------------
# Parser "A.3 Negociación" – Excel como verdad
# -----------------------------
def find_header_row(a3_raw):
    # Busco fila que contenga 'Corredor' y 'Ingreso' cerca
    for i in range(min(30, a3_raw.shape[0])):
        row_vals = [str(v).strip() if isinstance(v, str) else v for v in a3_raw.iloc[i].tolist()]
        if any(isinstance(v, str) and v.lower().startswith("corredor") for v in row_vals):
            return i
    # fallback
    return 8

def make_column_names(a3_raw, header_row):
    """
    Construye nombres compuestos: [super]|[sub]
    donde 'super' es la fila anterior (header_row-1) y 'sub' la fila de encabezado.
    """
    supers = a3_raw.iloc[header_row-1].fillna("").astype(str).tolist() if header_row-1 >= 0 else [""]*a3_raw.shape[1]
    subs   = a3_raw.iloc[header_row].fillna("").astype(str).tolist()
    names = []
    for s, sub in zip(supers, subs):
        s = s.strip()
        sub = sub.strip()
        if s and s != sub:
            names.append(f"{s}|{sub}")
        else:
            names.append(sub or s)
    # resolver duplicados
    seen = {}
    final = []
    for n in names:
        if n in seen:
            seen[n] += 1
            final.append(f"{n}_{seen[n]}")
        else:
            seen[n] = 0
            final.append(n)
    return final

def parse_a3_negociacion(a3_raw):
    h = find_header_row(a3_raw)
    cols = make_column_names(a3_raw, h)
    df = a3_raw.iloc[h+1:].copy()
    df.columns = cols
    # Identificar columnas clave por patrones (robusto a cambios menores)
    def find_col(patterns):
        pats = [p.lower() for p in patterns]
        for c in df.columns:
            lc = c.lower()
            if all(p in lc for p in pats):
                return c
        return None

    col_cliente = find_col(["corredor"]) or "Corredor"
    col_monto   = find_col(["monto","neg"]) or find_col(["monto"])
    # Ingreso Total
    col_real_tot = (find_col(["ingreso","total","real"]) or
                    find_col(["ingresos","total","real"]) or
                    find_col(["ingreso total","real"]))
    col_proj_tot = (find_col(["ingreso","total","proy"]) or
                    find_col(["ingresos","total","proy"]) or
                    find_col(["ingreso total","proyectado"]))

    # Si no encuentro Total, intentar sumar componentes (Acceso + Transacción + Perfiles) Real/Proy
    def try_components(prefix):
        acc = find_col([prefix,"acceso"]) or find_col([prefix,"accesos"])
        trx = find_col([prefix,"transacci"]) or find_col([prefix,"transacción"]) or find_col([prefix,"transaccion"])
        per = find_col([prefix,"perf"]) or find_col([prefix,"perfiles"])
        cols = [c for c in [acc,trx,per] if c]
        return cols

    if col_real_tot is None:
        comps = try_components("real")
        if comps:
            df["__real_total_calc__"] = df[comps].apply(pd.to_numeric, errors="coerce").fillna(0.0).sum(axis=1)
            col_real_tot = "__real_total_calc__"
    if col_proj_tot is None:
        comps = try_components("proy") or try_components("proyectado")
        if comps:
            df["__proj_total_calc__"] = df[comps].apply(pd.to_numeric, errors="coerce").fillna(0.0).sum(axis=1)
            col_proj_tot = "__proj_total_calc__"

    # Limpieza numérica
    for c in [col_monto, col_real_tot, col_proj_tot]:
        if c and c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)

    # Determinar bolsa por bloques (BCS/BVL/BVC)
    df["Bolsa"] = None
    current_bolsa = None
    bolsas = {"BCS":"Chile","BVL":"Perú","BVC":"Colombia"}
    for i, row in df.iterrows():
        val = str(row.get(col_cliente, "")).strip()
        if val in bolsas:
            current_bolsa = val
            df.at[i, "Bolsa"] = val
        else:
            df.at[i, "Bolsa"] = current_bolsa

    # Separar filas de totales por bolsa
    df_bolsas = df[df[col_cliente].isin(bolsas.keys())].copy()
    df_bolsas["Pais"] = df_bolsas[col_cliente].map(bolsas)

    # Filas de clientes (excluir agregados/blank)
    mask_clientes = (~df[col_cliente].isin(bolsas.keys())) & df[col_cliente].notna() & (df[col_cliente].astype(str).str.strip()!="")
    clientes = df[mask_clientes].copy()
    clientes["Pais"] = clientes["Bolsa"].map(bolsas)

    # Renombrar columnas estándar
    rename = {}
    if col_cliente: rename[col_cliente]="Cliente"
    if col_monto:   rename[col_monto]="Monto USD"
    if col_real_tot:rename[col_real_tot]="Real Excel"
    if col_proj_tot:rename[col_proj_tot]="Proyectado Excel"
    clientes = clientes.rename(columns=rename)
    df_bolsas = df_bolsas.rename(columns={col_real_tot:"Real Excel", col_proj_tot:"Proyectado Excel"})

    # Asegurar columnas
    for c in ["Monto USD","Real Excel","Proyectado Excel"]:
        if c not in clientes.columns: clientes[c]=0.0

    return clientes[["Pais","Cliente","Monto USD","Real Excel","Proyectado Excel"]], df_bolsas[["Pais","Bolsa","Real Excel","Proyectado Excel"]]

# -----------------------------
# Parser "6. Customer Journey" – solo para mostrar referencia
# -----------------------------
def parse_customer_journey(cj_raw):
    # Busco celdas con "Ingreso Actual" y "Ingreso Propuesta" en proximidad por país
    out = []
    grid = to_grid(cj_raw)
    # Buscar títulos de países
    def text_at(r,c): return grid.get(r,{}).get(c,"")
    for r in range(cj_raw.shape[0]-2):
        for c in ["B","C","D","E","F","G","H","I","J"]:
            v = text_at(r,c)
            if isinstance(v, str) and any(p in v for p in ["Chile","Perú","Colombia","BCS","BVL","BVC"]):
                # Buscar cabecera Ingreso Actual en la fila siguiente
                found = False
                for cc in ["B","C","D","E","F","G","H","I","J","K","L","M"]:
                    if text_at(r+1, cc) == "Ingreso Actual":
                        found = True
                        col_act = cc
                        col_prop = col_letter((ord(cc)-65)+1)
                        # Tomo 10 filas hacia abajo o hasta que se corte
                        for rr in range(r+2, min(r+14, cj_raw.shape[0])):
                            concepto = text_at(rr, col_letter(ord(cc)-66))  # una a la izquierda
                            if isinstance(concepto, str) and "Negociación" in concepto:
                                pais = "Chile" if "Chile" in v or "BCS" in v else ("Perú" if "Perú" in v or "BVL" in v else ("Colombia" if "Colombia" in v or "BVC" in v else v))
                                out.append({
                                    "Pais": pais,
                                    "Concepto": concepto,
                                    "Actual": safe_float(text_at(rr, col_act), 0.0),
                                    "Propuesta": safe_float(text_at(rr, col_prop), 0.0)
                                })
                                break
                if found:
                    break
    return pd.DataFrame(out)

# -----------------------------
# Parametría – "1. Parametros" (columna R+)
# -----------------------------
def parse_tramos_right(grid, rows, mapping, fields=("min","max","var","fijo")):
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
                    fields[2]: safe_float(var, 0.0),
                    fields[3]: safe_float(fijo, 0.0)
                })
        out[country] = sorted(out[country], key=lambda d: d["min"])
    return out

def read_parametros(params_raw):
    grid = to_grid(params_raw)
    mapping = {"Colombia":("T","U","V","W"), "Perú":("X","Y","Z","AA"), "Chile":("AB","AC","AD","AE")}
    # Transacción (asumo bps+fijo)
    trans = parse_tramos_right(grid, rows=range(134, 145), mapping=mapping, fields=("min","max","bps","fijo"))
    # DMA (bps+fijo)
    dma   = parse_tramos_right(grid, rows=range(146, 160), mapping=mapping, fields=("min","max","bps","fijo"))
    # Acceso/Códigos (var por código + fijo mensual por tramo)
    pant  = parse_tramos_right(grid, rows=range(117, 135), mapping=mapping, fields=("min","max","var","fija"))
    # Descuento BCS (si existe)
    desc = safe_float(grid.get(114,{}).get("C"), 0.15)
    return {"transaccion":trans, "dma":dma, "pantallas":pant, "desc_bcs":desc}

# -----------------------------
# Motor de Simulación
# -----------------------------
def ingreso_transaccion(monto_usd, tramos):
    if not tramos:
        return 0.0
    # tomo tramo cuyo rango contiene el monto
    t = None
    for x in tramos:
        if monto_usd >= x["min"] and monto_usd <= x["max"]:
            t = x; break
    if t is None:
        t = tramos[-1]
    rate = infer_rate(t.get("bps", 0.0))
    return monto_usd * rate + safe_float(t.get("fijo", 0.0))

def ingreso_dma(monto_dma_usd, tramos_dma):
    if not tramos_dma:
        return 0.0
    t = None
    for x in tramos_dma:
        if monto_dma_usd >= x["min"] and monto_dma_usd <= x["max"]:
            t = x; break
    if t is None:
        t = tramos_dma[-1]
    rate = infer_rate(t.get("bps", 0.0))
    return monto_dma_usd * rate + safe_float(t.get("fijo", 0.0))

def ingreso_acceso_codigos(codigos, tramos, desc=0.0):
    if not tramos:
        return 0.0
    t = None
    for x in tramos:
        if codigos >= x["min"] and codigos <= x["max"]:
            t = x; break
    if t is None:
        t = tramos[-1]
    fija = safe_float(t.get("fija", 0.0))
    var  = safe_float(t.get("var", 0.0))
    bruto = fija + var * codigos
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
        sims.append(
            ingreso_transaccion(monto, tr_trx) +
            ingreso_dma(monto_dma, tr_dma) +
            ingreso_acceso_codigos(cods, tr_pant, desc=desc)
        )
    out = df.copy()
    out["Proyectado Sim"] = sims
    out["BPS Sim"] = out.apply(lambda r: calc_bps(r["Proyectado Sim"], r["Monto USD"]), axis=1)
    return out

# -----------------------------
# UI
# -----------------------------
st.title("📊 Simulador Tarifario RV – Reconciliado con Excel")

uploaded = st.file_uploader("📁 Sube tu Excel maestro", type=["xlsx"])
if not uploaded:
    st.info("Se requieren las hojas: 'A.3 Negociación', '6. Customer Journey', 'A.3 BBDD Neg' y '1. Parametros'.")
    st.stop()

with st.spinner("Leyendo libro y construyendo vistas…"):
    bbdd_neg, a3_raw, a3_guess, cj_raw, params_raw = read_excel_all(uploaded)
    df_cli, df_bolsas = parse_a3_negociacion(a3_raw)
    df_cj = parse_customer_journey(cj_raw)
    params = read_parametros(params_raw)

# Sidebar – Edición de tramos
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

# Filtros
paises = ["Todos"] + sorted(df_cli["Pais"].dropna().unique().tolist())
col_f1, col_f2 = st.columns(2)
with col_f1:
    pais_sel = st.selectbox("Filtrar por País", paises, index=0)
with col_f2:
    ver_detalle = st.toggle("Mostrar detalle por cliente", value=False)

df_cli_f = df_cli if pais_sel=="Todos" else df_cli[df_cli["Pais"]==pais_sel]
df_sim = simulate(df_cli_f, bbdd_neg, params_live)

# KPIs (Excel vs Sim)
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

# Reconciliación por Bolsa (debe cuadrar con suma de clientes por país)
st.markdown("### 🧮 Reconciliación por Bolsa (Excel)")
bol = df_bolsas.copy()
st.dataframe(bol.rename(columns={"Real Excel":"Real (Excel)","Proyectado Excel":"Proyectado (Excel)"}), use_container_width=True)

# Alerta si hay descuadre entre 'suma de clientes' y 'fila bolsa' (Excel)
toler = 1.0  # USD de tolerancia
desc_msgs = []
for pais, bolsa in [("Chile","BCS"),("Perú","BVL"),("Colombia","BVC")]:
    suma_clientes = df_cli[df_cli["Pais"]==pais]["Proyectado Excel"].sum()
    fila_bolsa = bol[bol["Bolsa"]==bolsa]["Proyectado Excel"].sum()
    if abs(suma_clientes - fila_bolsa) > toler:
        desc_msgs.append(f"{pais}: suma clientes Proy (Excel) ${suma_clientes:,.0f} ≠ fila {bolsa} ${fila_bolsa:,.0f}")
if desc_msgs:
    st.warning("Posible descuadre detectado en A.3 Negociación → " + " | ".join(desc_msgs))

# Customer Journey – Negociación
st.markdown("### 🧭 '6. Customer Journey' – Negociación (Excel)")
if not df_cj.empty:
    cj_neg = df_cj[df_cj["Concepto"].str.contains("Negociación", na=False)]
    cj_pivot = cj_neg.pivot_table(index="Pais", values=["Actual","Propuesta"], aggfunc="sum").reset_index()
    st.dataframe(cj_pivot.rename(columns={"Actual":"Ingreso Actual (Excel)","Propuesta":"Ingreso Propuesta (Excel)"}), use_container_width=True)

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
        if not df_cj.empty:
            cj_pivot.to_excel(writer, index=False, sheet_name="CJ_Neg")
    st.download_button("📥 Descargar Excel (Detalle+Resúmenes)", buffer.getvalue(), "reconciliado_resultados.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
