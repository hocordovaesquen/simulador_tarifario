# app.py
# ==================================================================================
# Simulador Tarifario RV — v6 (Excel‑driven)
# (see previous cell for full explanation in comments)
# ==================================================================================

import io
import math
import numpy as np
import pandas as pd
import streamlit as st
import plotly.graph_objects as go
from io import BytesIO
from typing import Dict, List, Tuple, Optional
import openpyxl

try:
    from xlcalculator import ModelCompiler, Model
    XL_SUPPORTED = True
except Exception:
    XL_SUPPORTED = False

st.set_page_config(page_title="Simulador Tarifario RV – Excel‑driven (v6)", page_icon="📊", layout="wide")

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

PARAM_SHEET = "1. Parametros"
NEG_SHEET   = "A.3 Negociación"
MAP_RIGHT = {"Colombia": ("T","U","V","W"), "Perú": ("X","Y","Z","AA"), "Chile": ("AB","AC","AD","AE")}
ROWS_PANT   = list(range(117, 135))
ROWS_TRX    = list(range(134, 145))
ROWS_DMA    = list(range(146, 160))
DESC_BCS_POS= ("C", 114)

def load_workbook(bytes_data: bytes):
    bytes_io = io.BytesIO(bytes_data)
    wb = openpyxl.load_workbook(bytes_io, data_only=False)
    return wb

def read_block_from_wb(wb, rows, fields=("min","max","bps","fijo")):
    ws = wb[PARAM_SHEET]
    out = {k: [] for k in MAP_RIGHT}
    for pais, (cmin, cmax, cvar, cfijo) in MAP_RIGHT.items():
        for rr in rows:
            mn = ws[f"{cmin}{rr}"].value
            mx = ws[f"{cmax}{rr}"].value
            var = ws[f"{cvar}{rr}"].value
            fijo= ws[f"{cfijo}{rr}"].value
            if any(is_num(v) for v in [mn, mx, var, fijo]):
                out[pais].append({fields[0]: safe_float(mn, 0.0), fields[1]: safe_float(mx, float("inf")), fields[2]: safe_float(var, 0.0), fields[3]: safe_float(fijo, 0.0)})
    return out

def read_params_from_wb(wb):
    ws = wb[PARAM_SHEET]
    desc = safe_float(ws[f"{DESC_BCS_POS[0]}{DESC_BCS_POS[1]}"].value, 0.15)
    trans = read_block_from_wb(wb, ROWS_TRX, ("min","max","bps","fijo"))
    dma   = read_block_from_wb(wb, ROWS_DMA, ("min","max","bps","fijo"))
    pant  = read_block_from_wb(wb, ROWS_PANT, ("min","max","var","fija"))
    return {"desc_bcs":desc, "transaccion":trans, "dma":dma, "pantallas":pant}

def write_params_to_wb(wb, params: Dict):
    ws = wb[PARAM_SHEET]
    desc = params.get("desc_bcs", None)
    if desc is not None:
        c, r = DESC_BCS_POS
        ws[f"{c}{r}"] = float(desc)
    def write_block(rows: List[int], data_key: str, field_names: Tuple[str,str,str,str]):
        block = params.get(data_key, {})
        for pais, cols in MAP_RIGHT.items():
            cmin, cmax, cvar, cfijo = cols
            for rr in rows:
                for cc in cols:
                    ws[f"{cc}{rr}"] = None
            lst = block.get(pais, [])
            for i, item in enumerate(lst):
                if i >= len(rows): break
                rr = rows[i]
                ws[f"{cmin}{rr}"] = safe_float(item.get(field_names[0], item.get("min", 0)))
                ws[f"{cmax}{rr}"] = safe_float(item.get(field_names[1], item.get("max", 0)))
                ws[f"{cvar}{rr}"] = safe_float(item.get(field_names[2], item.get("bps", item.get("var", 0))))
                ws[f"{cfijo}{rr}"] = safe_float(item.get(field_names[3], item.get("fijo", item.get("fija", 0))))
    write_block(ROWS_TRX, "transaccion", ("min","max","bps","fijo"))
    write_block(ROWS_DMA, "dma", ("min","max","bps","fijo"))
    write_block(ROWS_PANT, "pantallas", ("min","max","var","fija"))
    return wb

def find_a3_headers(ws):
    max_row = min(100, ws.max_row)
    header_row = None
    for r in range(1, max_row+1):
        row_vals = [str(ws.cell(r, c).value).strip().lower() if ws.cell(r,c).value is not None else "" for c in range(1, ws.max_column+1)]
        if any(v.startswith("corredor") or v.startswith("cliente") for v in row_vals):
            header_row = r
            break
    if header_row is None:
        header_row = 9
    super_headers = {c: (ws.cell(header_row, c).value or "") for c in range(1, ws.max_column+1)}
    sub_headers   = {c: (ws.cell(header_row+1, c).value or "") for c in range(1, ws.max_column+1)}
    return header_row, super_headers, sub_headers

def locate_a3_cols(ws):
    h, super_h, sub_h = find_a3_headers(ws)
    def match(pats, text):
        t = str(text or "").lower()
        return all(p in t for p in pats)
    def find_col(pats_super, pats_sub=None):
        for c in range(1, ws.max_column+1):
            sh = super_h.get(c, ""); sub = sub_h.get(c, "")
            if pats_sub:
                if match(pats_super, sh) and match(pats_sub, sub):
                    return c
            else:
                if match(pats_super, sh) or match(pats_super, sub):
                    return c
        return None
    col_cliente = find_col(["corredor"]) or find_col(["cliente"])
    col_monto   = find_col(["monto","neg"]) or find_col(["monto"])
    col_real    = find_col(["ingreso","total"], ["real"]) or find_col(["real","ingreso"])
    col_proy    = find_col(["ingreso","total"], ["proy"]) or find_col(["ingreso","total"], ["proyectado"]) or find_col(["propuesta"])
    return {"header_row": h, "cliente": col_cliente, "monto": col_monto, "real": col_real, "proy": col_proy}

def get_a3_data_addresses(ws):
    cols = locate_a3_cols(ws)
    h = cols["header_row"]
    cliente_c = cols["cliente"]; monto_c = cols["monto"]; real_c = cols["real"]; proy_c = cols["proy"]
    if not cliente_c: raise ValueError("No pude ubicar 'Corredor/Cliente' en A.3 Negociación.")
    bolsas = {"BCS":"Chile", "BVL":"Perú", "BVC":"Colombia"}
    clientes = []; agregados = []
    r = h+2
    while r <= ws.max_row:
        name = ws.cell(r, cliente_c).value
        if name is None or str(name).strip()=="" or str(name).strip().lower().startswith("total"):
            r += 1; continue
        a = {"row": r, "cliente": str(name).strip(),
             "monto_addr": ws.cell(r, monto_c).coordinate if monto_c else None,
             "real_addr":  ws.cell(r, real_c).coordinate if real_c else None,
             "proy_addr":  ws.cell(r, proy_c).coordinate if proy_c else None}
        if str(name).strip() in bolsas:
            agregados.append({"pais": bolsas[str(name).strip()], "bolsa": str(name).strip(),
                              "real_addr": a["real_addr"], "proy_addr": a["proy_addr"]})
        else:
            clientes.append(a)
        r += 1
    return clientes, agregados

def recalc_and_extract(wb):
    ws_neg = wb[NEG_SHEET]
    meta = {"engine":"fallback"}
    model = None
    if XL_SUPPORTED:
        try:
            bio = BytesIO(); wb.save(bio); bio.seek(0)
            mc = ModelCompiler(); model = Model(mc.read_and_parse_archive(bio))
            meta["engine"] = "xlcalculator"
        except Exception as e:
            meta["engine_error"] = str(e); model = None
    clientes_addr, agregados_addr = get_a3_data_addresses(ws_neg)
    def eval_cell(address):
        if not address: return 0.0
        if model is None:
            val = ws_neg[address].value
            return safe_float(val, 0.0)
        try:
            return safe_float(model.evaluate(f"'{NEG_SHEET}'!{address}"), 0.0)
        except Exception:
            return safe_float(ws_neg[address].value, 0.0)
    out = []
    for a in clientes_addr:
        out.append({"Pais":"", "Cliente":a["cliente"],
                    "Monto USD": eval_cell(a["monto_addr"]),
                    "Real Excel": eval_cell(a["real_addr"]),
                    "Proyectado Excel": eval_cell(a["proy_addr"])})
    df_cli = pd.DataFrame(out)
    rows = []
    for ag in agregados_addr:
        rows.append({"Pais":ag["pais"], "Bolsa":ag["bolsa"],
                     "Real Excel": eval_cell(ag["real_addr"]), "Proyectado Excel": eval_cell(ag["proy_addr"])})
    df_bolsas = pd.DataFrame(rows)
    return df_cli, df_bolsas, meta

def approx_engine(df_clients, bbdd_neg):
    df = df_clients.copy()
    if "Pais" in bbdd_neg.columns and "Cliente estandar" in bbdd_neg.columns:
        m = bbdd_neg.groupby("Cliente estandar")["Pais"].agg(lambda s: s.dropna().iloc[0] if len(s.dropna()) else "").to_dict()
        df["Pais"] = df["Cliente"].map(m).fillna(df.get("Pais",""))
    return df

# ---------------- UI ----------------
st.title("📊 Simulador Tarifario RV – Excel‑driven (v6)")
st.caption("Edita tramos (columna R+ de **1. Parametros**) y recalcula el libro con **xlcalculator**.")

uploaded = st.file_uploader("📁 Sube tu Excel maestro", type=["xlsx"])
if not uploaded:
    st.stop()

raw_bytes = uploaded.read()
wb_prefill = load_workbook(raw_bytes)
params0 = read_params_from_wb(wb_prefill)

with st.sidebar:
    st.header("⚙️ Parámetros (columna R+)")
    desc_bcs = st.number_input("Descuento BCS (0–1)", min_value=0.0, max_value=1.0, step=0.01, value=float(params0.get("desc_bcs",0.15)))

    st.subheader("Transacción – Tramos")
    trans_edit = {}
    for pais in ["Chile","Colombia","Perú"]:
        base = params0["transaccion"].get(pais, [])
        df_e = pd.DataFrame(base) if base else pd.DataFrame(columns=["min","max","bps","fijo"])
        df_show = df_e.rename(columns={"min":"Mín (USD)","max":"Máx (USD)","bps":"Variable (fracción/%)","fijo":"Fijo (USD)"})
        df_show = st.data_editor(df_show, key=f"trx_{pais}", use_container_width=True, num_rows="dynamic")
        rows = []
        for _, r in df_show.iterrows():
            rows.append({"min":safe_float(r.get("Mín (USD)",0)), "max":safe_float(r.get("Máx (USD)",float("inf"))),
                         "bps":safe_float(r.get("Variable (fracción/%)",0)), "fijo":safe_float(r.get("Fijo (USD)",0))})
        trans_edit[pais] = rows

    st.subheader("DMA – Tramos")
    dma_edit = {}
    for pais in ["Chile","Colombia","Perú"]:
        base = params0["dma"].get(pais, [])
        df_e = pd.DataFrame(base) if base else pd.DataFrame(columns=["min","max","bps","fijo"])
        df_show = df_e.rename(columns={"min":"Mín (USD)","max":"Máx (USD)","bps":"Variable (fracción/%)","fijo":"Fijo (USD)"})
        df_show = st.data_editor(df_show, key=f"dma_{pais}", use_container_width=True, num_rows="dynamic")
        rows = []
        for _, r in df_show.iterrows():
            rows.append({"min":safe_float(r.get("Mín (USD)",0)), "max":safe_float(r.get("Máx (USD)",float("inf"))),
                         "bps":safe_float(r.get("Variable (fracción/%)",0)), "fijo":safe_float(r.get("Fijo (USD)",0))})
        dma_edit[pais] = rows

    st.subheader("Acceso – Códigos/Pantallas")
    pant_edit = {}
    for pais in ["Chile","Colombia","Perú"]:
        base = params0["pantallas"].get(pais, [])
        df_e = pd.DataFrame(base) if base else pd.DataFrame(columns=["min","max","var","fija"])
        df_show = df_e.rename(columns={"min":"Mín #Códigos","max":"Máx #Códigos","var":"Variable por código (USD)","fija":"Fijo mensual tramo (USD)"})
        df_show = st.data_editor(df_show, key=f"pant_{pais}", use_container_width=True, num_rows="dynamic")
        rows = []
        for _, r in df_show.iterrows():
            rows.append({"min":safe_float(r.get("Mín #Códigos",0)), "max":safe_float(r.get("Máx #Códigos",float("inf"))),
                         "var":safe_float(r.get("Variable por código (USD)",0)), "fija":safe_float(r.get("Fijo mensual tramo (USD)",0))})
        pant_edit[pais] = rows

params_live = {"desc_bcs":desc_bcs, "transaccion":trans_edit, "dma":dma_edit, "pantallas":pant_edit}

with st.spinner("Aplicando parámetros y recalculando…"):
    wb = load_workbook(raw_bytes)
    wb = write_params_to_wb(wb, params_live)
    df_cli, df_bol, meta = recalc_and_extract(wb)

try:
    bbdd = pd.read_excel(io.BytesIO(raw_bytes), sheet_name="A.3 BBDD Neg", header=6, engine="openpyxl")
except Exception:
    bbdd = pd.DataFrame()

df_cli2 = approx_engine(df_cli, bbdd)

paises = ["Todos"] + sorted([p for p in df_cli2.get("Pais","").unique() if isinstance(p,str) and p!=""])
c1, c2 = st.columns(2)
with c1:
    pais_sel = st.selectbox("Filtrar por País", paises, index=0)
with c2:
    ver_detalle = st.toggle("Mostrar detalle por cliente", value=True)

df_f = df_cli2.copy() if pais_sel=="Todos" else df_cli2[df_cli2["Pais"]==pais_sel].copy()
tot_real = safe_float(df_f["Real Excel"].sum(),0.0)
tot_proy = safe_float(df_f["Proyectado Excel"].sum(),0.0)
monto    = safe_float(df_f["Monto USD"].sum(),0.0)

a,b,c,d = st.columns(4)
a.metric("Ingreso Real (Excel)", f"${tot_real:,.0f}")
b.metric("Ingreso Proyectado (Excel)", f"${tot_proy:,.0f}", delta=(f"{(tot_proy/tot_real-1)*100:+.1f}%" if tot_real>0 else None))
b.caption(f"Motor: **{meta.get('engine','fallback')}**")
c.metric("BPS Proyectado (Excel)", f"{calc_bps(tot_proy, monto):.2f} bps")
d.metric("Filas cargadas", f"{len(df_f):,}")

st.markdown("---")
cA, cB = st.columns(2, gap="large")
with cA:
    st.subheader("Real vs Proyectado – Excel")
    fig = go.Figure(data=[
        go.Bar(name="Real (Excel)", x=["Negociación"], y=[tot_real]),
        go.Bar(name="Proyectado (Excel)", x=["Negociación"], y=[tot_proy])
    ]); fig.update_layout(barmode="group", height=300)
    st.plotly_chart(fig, use_container_width=True)

with cB:
    st.subheader("Totales por Bolsa (Excel)")
    st.dataframe(df_bol, use_container_width=True, height=300)

st.subheader("📋 Detalle por Cliente (Excel recalculado)")
if ver_detalle:
    st.dataframe(df_f.sort_values(["Pais","Proyectado Excel","Real Excel"], ascending=[True, False, False]), use_container_width=True, height=420)

st.markdown("---")
exp = st.expander("🔎 Diagnóstico")
exp.write(meta)

col1, col2 = st.columns(2)
with col1:
    st.download_button("📥 Descargar Detalle (CSV)", df_f.to_csv(index=False).encode("utf-8"), "detalle_recalculado.csv", "text/csv", use_container_width=True)
with col2:
    out = BytesIO(); wb.save(out)
    st.download_button("📥 Descargar Excel recalculado (.xlsx)", out.getvalue(), "excel_recalculado.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
