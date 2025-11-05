"""
DASHBOARD SIMULADOR TARIFARIO RV
=================================
INSTALACIÓN:
pip install streamlit pandas numpy openpyxl plotly

EJECUCIÓN:
streamlit run app.py
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go

st.set_page_config(page_title="Simulador Tarifario RV", page_icon="🚀", layout="wide")

# =============================================================================
# FUNCIONES AUXILIARES
# =============================================================================

def limpiar_numero(valor):
    """Convierte a número, manejando NaN y formatos raros"""
    if pd.isna(valor):
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    try:
        valor_str = str(valor).replace('$', '').replace(',', '').replace('\xa0', '')
        return float(valor_str)
    except:
        return 0.0

# =============================================================================
# LECTURA DE DATOS
# =============================================================================

@st.cache_data
def leer_datos_completos(archivo):
    """Lee ambas hojas y retorna datos limpios"""
    
    # 1. Leer negociación (header en fila 6)
    df_neg = pd.read_excel(archivo, sheet_name='A.3 BBDD Neg', header=6)
    
    # Limpiar y agrupar
    df_neg = df_neg[df_neg['Cliente estandar'].notna()].copy()
    df_neg['Monto USD'] = df_neg['Monto USD'].apply(limpiar_numero)
    df_neg['Cobro Acceso'] = df_neg['Cobro Acceso'].apply(limpiar_numero)
    df_neg['Cobro Transacción'] = df_neg['Cobro Transacción'].apply(limpiar_numero)
    
    # Agrupar por broker y país
    df_agrupado = df_neg.groupby(['Cliente estandar', 'Pais']).agg({
        'Monto USD': 'sum',
        'Cobro Acceso': 'sum',
        'Cobro Transacción': 'sum'
    }).reset_index()
    
    df_agrupado.columns = ['Broker', 'Pais', 'Monto_USD', 'Ingreso_Acceso', 'Ingreso_Transaccion']
    
    # 2. Parámetros por defecto (usuario puede editarlos)
    parametros = {
        'Colombia': {
            'Acceso': [
                {'min': 0, 'max': 10_000_000, 'var': 0.15, 'fija': 5000},
                {'min': 10_000_000, 'max': 50_000_000, 'var': 0.12, 'fija': 15000},
                {'min': 50_000_000, 'max': float('inf'), 'var': 0.10, 'fija': 30000}
            ],
            'Transacción': [
                {'min': 0, 'max': 10_000_000, 'var': 0.08, 'fija': 2000},
                {'min': 10_000_000, 'max': 50_000_000, 'var': 0.06, 'fija': 8000},
                {'min': 50_000_000, 'max': float('inf'), 'var': 0.05, 'fija': 15000}
            ]
        },
        'Peru': {
            'Acceso': [
                {'min': 0, 'max': 10_000_000, 'var': 0.15, 'fija': 5000},
                {'min': 10_000_000, 'max': 50_000_000, 'var': 0.12, 'fija': 15000},
                {'min': 50_000_000, 'max': float('inf'), 'var': 0.10, 'fija': 30000}
            ],
            'Transacción': [
                {'min': 0, 'max': 10_000_000, 'var': 0.08, 'fija': 2000},
                {'min': 10_000_000, 'max': 50_000_000, 'var': 0.06, 'fija': 8000},
                {'min': 50_000_000, 'max': float('inf'), 'var': 0.05, 'fija': 15000}
            ]
        },
        'Chile': {
            'Acceso': [
                {'min': 0, 'max': 193_368_604, 'var': 0.005, 'fija': 0},
                {'min': 193_368_604, 'max': 386_737_209, 'var': 0.003, 'fija': 0},
                {'min': 386_737_209, 'max': float('inf'), 'var': 0.000, 'fija': 0}
            ],
            'Transacción': [
                {'min': 0, 'max': 2500, 'var': 0.000, 'fija': 0},
                {'min': 2500, 'max': 125_000, 'var': 0.027, 'fija': 0},
                {'min': 125_000, 'max': float('inf'), 'var': 0.0214, 'fija': 0}
            ]
        }
    }
    
    return df_agrupado, parametros

# =============================================================================
# CÁLCULOS
# =============================================================================

def calcular_simulado(monto, tramos):
    """Calcula ingreso simulado según tramos"""
    for tramo in tramos:
        if tramo['min'] <= monto < tramo['max']:
            return monto * (tramo['var'] / 100) + tramo['fija']
    # Si no cae en ningún tramo, usar el último
    if len(tramos) > 0:
        ultimo = tramos[-1]
        return monto * (ultimo['var'] / 100) + ultimo['fija']
    return 0.0

def calcular_tarifa_bps(ingreso, monto):
    """Calcula tarifa implícita en basis points"""
    if monto == 0:
        return 0.0
    return (ingreso / monto) * 10_000

# =============================================================================
# INTERFAZ PRINCIPAL
# =============================================================================

def main():
    st.title("🚀 Simulador Tarifario RV")
    st.markdown("---")
    
    # Sidebar
    with st.sidebar:
        st.header("⚙️ Configuración")
        
        archivo = st.file_uploader("📁 Sube tu Excel", type=['xlsx'])
        
        if archivo is None:
            st.info("👆 Sube un archivo Excel para comenzar")
            st.stop()
        
        # Leer datos
        with st.spinner("Leyendo archivo..."):
            df_datos, parametros = leer_datos_completos(archivo)
        
        st.success(f"✅ {len(df_datos)} brokers cargados")
        
        # Filtros
        st.markdown("### 🔍 Filtros")
        
        pais_opts = ['Todos'] + sorted(df_datos['Pais'].unique().tolist())
        pais_filtro = st.selectbox("País", pais_opts)
        
        producto = st.selectbox("Producto", ['Negociación - Acceso', 'Negociación - Transacción'])
        producto_key = 'Acceso' if 'Acceso' in producto else 'Transacción'
        
        # Editor de tramos
        st.markdown("### 📊 Editar Tramos")
        
        if pais_filtro == 'Todos':
            paises_edit = ['Colombia', 'Peru', 'Chile']
        else:
            paises_edit = [pais_filtro]
        
        for pais in paises_edit:
            if pais not in parametros:
                continue
                
            with st.expander(f"🌎 {pais}"):
                tramos = parametros[pais][producto_key]
                
                # Crear DataFrame editable
                df_tramos = pd.DataFrame(tramos)
                df_tramos.columns = ['Mínimo (USD)', 'Máximo (USD)', 'Variable (%)', 'Fija (USD)']
                
                # Reemplazar inf con texto
                df_tramos['Máximo (USD)'] = df_tramos['Máximo (USD)'].replace(float('inf'), 999_999_999_999)
                
                df_editado = st.data_editor(
                    df_tramos,
                    num_rows="dynamic",
                    use_container_width=True,
                    key=f"tramos_{pais}_{producto_key}"
                )
                
                # Actualizar parámetros
                tramos_nuevos = []
                for _, row in df_editado.iterrows():
                    max_val = float('inf') if row['Máximo (USD)'] >= 999_999_999_999 else row['Máximo (USD)']
                    tramos_nuevos.append({
                        'min': row['Mínimo (USD)'],
                        'max': max_val,
                        'var': row['Variable (%)'],
                        'fija': row['Fija (USD)']
                    })
                parametros[pais][producto_key] = tramos_nuevos
        
        if st.button("🔄 Recalcular", type="primary", use_container_width=True):
            st.rerun()
    
    # Filtrar datos
    if pais_filtro == 'Todos':
        df_filtrado = df_datos.copy()
    else:
        df_filtrado = df_datos[df_datos['Pais'] == pais_filtro].copy()
    
    # Calcular simulado
    resultados = []
    for _, row in df_filtrado.iterrows():
        broker = row['Broker']
        pais = row['Pais']
        monto = row['Monto_USD']
        
        if producto_key == 'Acceso':
            ingreso_real = row['Ingreso_Acceso']
        else:
            ingreso_real = row['Ingreso_Transaccion']
        
        # Obtener tramos del país
        if pais in parametros and producto_key in parametros[pais]:
            tramos = parametros[pais][producto_key]
            ingreso_sim = calcular_simulado(monto, tramos)
        else:
            ingreso_sim = 0.0
        
        diferencia = ingreso_sim - ingreso_real
        var_pct = (diferencia / ingreso_real * 100) if ingreso_real != 0 else 0
        
        tarifa_real_bps = calcular_tarifa_bps(ingreso_real, monto)
        tarifa_sim_bps = calcular_tarifa_bps(ingreso_sim, monto)
        
        resultados.append({
            'Broker': broker,
            'Pais': pais,
            'Monto_Negociado': monto,
            'Ingreso_Real': ingreso_real,
            'Ingreso_Simulado': ingreso_sim,
            'Diferencia': diferencia,
            'Var_Pct': var_pct,
            'Tarifa_Real_bps': tarifa_real_bps,
            'Tarifa_Sim_bps': tarifa_sim_bps
        })
    
    df_resultados = pd.DataFrame(resultados)
    
    # =============================================================================
    # VISUALIZACIONES
    # =============================================================================
    
    # KPIs
    total_real = df_resultados['Ingreso_Real'].sum()
    total_sim = df_resultados['Ingreso_Simulado'].sum()
    diferencia = total_sim - total_real
    var_pct = (diferencia / total_real * 100) if total_real != 0 else 0
    total_monto = df_resultados['Monto_Negociado'].sum()
    tarifa_prom = calcular_tarifa_bps(total_sim, total_monto)
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("💰 Ingreso Real", f"${total_real:,.0f}")
    
    with col2:
        st.metric("🎯 Ingreso Simulado", f"${total_sim:,.0f}", 
                 delta=f"{var_pct:+.1f}%")
    
    with col3:
        st.metric("📊 Diferencia", f"${diferencia:,.0f}")
    
    with col4:
        st.metric("📈 Tarifa Implícita", f"{tarifa_prom:.2f} bps")
    
    st.markdown("---")
    
    # Gráficos
    col_g1, col_g2 = st.columns(2)
    
    with col_g1:
        st.subheader("📊 Comparación Ingresos")
        fig_comp = go.Figure(data=[
            go.Bar(name='Real', x=['Real'], y=[total_real], marker_color='#667eea'),
            go.Bar(name='Simulado', x=['Simulado'], y=[total_sim], marker_color='#27ae60')
        ])
        fig_comp.update_layout(height=300, showlegend=False)
        st.plotly_chart(fig_comp, use_container_width=True)
    
    with col_g2:
        st.subheader("🏆 Top 10 Brokers")
        top10 = df_resultados.nlargest(10, 'Monto_Negociado')
        fig_top = go.Figure(data=[
            go.Bar(name='Real', x=top10['Broker'], y=top10['Ingreso_Real'], marker_color='#667eea'),
            go.Bar(name='Simulado', x=top10['Broker'], y=top10['Ingreso_Simulado'], marker_color='#27ae60')
        ])
        fig_top.update_layout(height=300, xaxis_tickangle=-45)
        st.plotly_chart(fig_top, use_container_width=True)
    
    st.markdown("---")
    
    # Tabla
    st.subheader("📋 Detalle por Broker")
    
    df_display = df_resultados.sort_values('Monto_Negociado', ascending=False).copy()
    
    # Formatear
    df_display['Monto_Negociado'] = df_display['Monto_Negociado'].apply(lambda x: f"${x:,.0f}")
    df_display['Ingreso_Real'] = df_display['Ingreso_Real'].apply(lambda x: f"${x:,.2f}")
    df_display['Ingreso_Simulado'] = df_display['Ingreso_Simulado'].apply(lambda x: f"${x:,.2f}")
    df_display['Diferencia'] = df_display['Diferencia'].apply(lambda x: f"${x:,.2f}")
    df_display['Var_Pct'] = df_display['Var_Pct'].apply(lambda x: f"{x:.2f}%")
    df_display['Tarifa_Real_bps'] = df_display['Tarifa_Real_bps'].apply(lambda x: f"{x:.2f}")
    df_display['Tarifa_Sim_bps'] = df_display['Tarifa_Sim_bps'].apply(lambda x: f"{x:.2f}")
    
    st.dataframe(df_display, use_container_width=True, height=400)
    
    # Exportar
    st.markdown("---")
    col_exp1, col_exp2 = st.columns(2)
    
    with col_exp1:
        csv = df_resultados.to_csv(index=False).encode('utf-8')
        st.download_button("📥 Descargar Detalle CSV", csv, "detalle.csv", 
                          "text/csv", use_container_width=True)
    
    with col_exp2:
        resumen = df_resultados.groupby('Pais').agg({
            'Monto_Negociado': 'sum',
            'Ingreso_Real': 'sum',
            'Ingreso_Simulado': 'sum'
        }).reset_index()
        csv_resumen = resumen.to_csv(index=False).encode('utf-8')
        st.download_button("📥 Descargar Resumen por País", csv_resumen, 
                          "resumen_pais.csv", "text/csv", use_container_width=True)

if __name__ == "__main__":
    main()
