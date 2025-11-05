"""
🚀 SIMULADOR TARIFARIO RV - VERSIÓN SUPREMA
============================================
Versión ultra robusta, simple y funcional

pip install streamlit pandas numpy openpyxl plotly
streamlit run app_suprema.py
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io

# ==================== CONFIGURACIÓN ====================
st.set_page_config(
    page_title="Simulador Tarifario RV SUPREMO",
    page_icon="🚀",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==================== CSS PERSONALIZADO ====================
st.markdown("""
<style>
    .main-header {
        font-size: 3rem;
        font-weight: bold;
        background: linear-gradient(120deg, #667eea 0%, #764ba2 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        text-align: center;
        padding: 1rem 0;
        margin-bottom: 2rem;
    }
    .stMetric {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 1rem;
        border-radius: 10px;
        color: white;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .stMetric label {
        color: white !important;
        font-weight: 600;
    }
    .stMetric [data-testid="stMetricValue"] {
        color: white !important;
    }
</style>
""", unsafe_allow_html=True)

# ==================== FUNCIONES AUXILIARES ====================
def limpiar_numero(valor):
    """Limpia y convierte valores a float"""
    if pd.isna(valor):
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    try:
        valor_str = str(valor).replace('$', '').replace(',', '').strip()
        return float(valor_str) if valor_str else 0.0
    except:
        return 0.0

def calcular_bps(ingreso, monto):
    """Calcula basis points"""
    if monto == 0 or pd.isna(monto):
        return 0.0
    return (ingreso / monto) * 10_000

# ==================== LECTURA DE DATOS ====================
@st.cache_data
def cargar_datos_negociacion(archivo):
    """
    Carga datos de la hoja A.3 BBDD Neg
    Lee directamente las columnas importantes
    """
    try:
        # Leer con header en fila 6 (index 6 en pandas)
        df = pd.read_excel(archivo, sheet_name='A.3 BBDD Neg', header=6)
        
        # Verificar que tenemos las columnas necesarias
        columnas_necesarias = ['Cliente estandar', 'Pais', 'Monto USD']
        for col in columnas_necesarias:
            if col not in df.columns:
                st.error(f"❌ No se encontró la columna '{col}' en el Excel")
                return None
        
        # Filtrar filas válidas
        df = df[df['Cliente estandar'].notna()].copy()
        
        # Limpiar columnas numéricas
        columnas_numericas = ['Monto USD', 'Acceso actual', 'Transaccion actual', 
                             'Cobro Acceso', 'Cobro Transacción']
        
        for col in columnas_numericas:
            if col in df.columns:
                df[col] = df[col].apply(limpiar_numero)
        
        # Agrupar por cliente y país
        df_agrupado = df.groupby(['Cliente estandar', 'Pais'], dropna=True).agg({
            'Monto USD': 'sum',
            'Acceso actual': 'sum',
            'Transaccion actual': 'sum',
            'Cobro Acceso': 'sum',
            'Cobro Transacción': 'sum'
        }).reset_index()
        
        # Renombrar columnas para claridad
        df_agrupado.columns = [
            'Broker', 'Pais', 'Monto_USD',
            'Acceso_Real', 'Trans_Real',
            'Acceso_Propuesta', 'Trans_Propuesta'
        ]
        
        return df_agrupado
        
    except Exception as e:
        st.error(f"❌ Error al cargar datos: {str(e)}")
        st.error("Verifica que el Excel tenga la hoja 'A.3 BBDD Neg' con las columnas correctas")
        return None

@st.cache_data
def cargar_parametros_excel(archivo):
    """
    Lee los parámetros desde la hoja 1. Parametros
    Lee desde la columna R (Nuevo Tarifario)
    """
    parametros = {
        'Negociacion': {
            'Acceso': {
                'Colombia': [],
                'Peru': [],
                'Chile': []
            },
            'Transaccion': {
                'Colombia': [],
                'Peru': [],
                'Chile': []
            }
        }
    }
    
    try:
        # Leer hoja de parámetros sin header
        df_params = pd.read_excel(archivo, sheet_name='1. Parametros', header=None)
        
        # ACCESO - Leer filas 99-104 (índices 98-103 en pandas)
        # Colombia: cols 19-22 (T, U, V, W en Excel = índices 19-22)
        # Peru: cols 23-26 (X, Y, Z, AA)
        # Chile: cols 27-30 (AB, AC, AD, AE)
        
        # ACCESO COLOMBIA
        for i in range(99, 104):  # filas 99-103
            try:
                min_val = limpiar_numero(df_params.iloc[i, 19])  # Col T
                max_val = limpiar_numero(df_params.iloc[i, 20])  # Col U
                var_val = limpiar_numero(df_params.iloc[i, 21])  # Col V
                fija_val = limpiar_numero(df_params.iloc[i, 22]) # Col W
                
                if max_val > 1e15:
                    max_val = float('inf')
                
                if min_val > 0 or max_val > 0 or var_val > 0 or fija_val > 0:
                    parametros['Negociacion']['Acceso']['Colombia'].append({
                        'min': min_val,
                        'max': max_val,
                        'var': var_val,
                        'fija': fija_val
                    })
            except:
                pass
        
        # ACCESO PERU
        for i in range(99, 104):
            try:
                min_val = limpiar_numero(df_params.iloc[i, 23])  # Col X
                max_val = limpiar_numero(df_params.iloc[i, 24])  # Col Y
                var_val = limpiar_numero(df_params.iloc[i, 25])  # Col Z
                fija_val = limpiar_numero(df_params.iloc[i, 26]) # Col AA
                
                if max_val > 1e15:
                    max_val = float('inf')
                
                if min_val > 0 or max_val > 0 or var_val > 0 or fija_val > 0:
                    parametros['Negociacion']['Acceso']['Peru'].append({
                        'min': min_val,
                        'max': max_val,
                        'var': var_val,
                        'fija': fija_val
                    })
            except:
                pass
        
        # ACCESO CHILE
        for i in range(99, 104):
            try:
                min_val = limpiar_numero(df_params.iloc[i, 27])  # Col AB
                max_val = limpiar_numero(df_params.iloc[i, 28])  # Col AC
                var_val = limpiar_numero(df_params.iloc[i, 29])  # Col AD
                fija_val = limpiar_numero(df_params.iloc[i, 30]) # Col AE
                
                if max_val > 1e15:
                    max_val = float('inf')
                
                if min_val > 0 or max_val > 0 or var_val > 0 or fija_val > 0:
                    parametros['Negociacion']['Acceso']['Chile'].append({
                        'min': min_val,
                        'max': max_val,
                        'var': var_val,
                        'fija': fija_val
                    })
            except:
                pass
        
        # TRANSACCIÓN - Leer filas 139-144
        # TRANSACCIÓN COLOMBIA
        for i in range(139, 145):
            try:
                min_val = limpiar_numero(df_params.iloc[i, 19])
                max_val = limpiar_numero(df_params.iloc[i, 20])
                var_val = limpiar_numero(df_params.iloc[i, 21])
                fija_val = limpiar_numero(df_params.iloc[i, 22])
                
                if max_val > 1e15:
                    max_val = float('inf')
                
                if min_val > 0 or max_val > 0 or var_val > 0 or fija_val > 0:
                    parametros['Negociacion']['Transaccion']['Colombia'].append({
                        'min': min_val,
                        'max': max_val,
                        'var': var_val,
                        'fija': fija_val
                    })
            except:
                pass
        
        # TRANSACCIÓN PERU
        for i in range(139, 145):
            try:
                min_val = limpiar_numero(df_params.iloc[i, 23])
                max_val = limpiar_numero(df_params.iloc[i, 24])
                var_val = limpiar_numero(df_params.iloc[i, 25])
                fija_val = limpiar_numero(df_params.iloc[i, 26])
                
                if max_val > 1e15:
                    max_val = float('inf')
                
                if min_val > 0 or max_val > 0 or var_val > 0 or fija_val > 0:
                    parametros['Negociacion']['Transaccion']['Peru'].append({
                        'min': min_val,
                        'max': max_val,
                        'var': var_val,
                        'fija': fija_val
                    })
            except:
                pass
        
        # TRANSACCIÓN CHILE
        for i in range(139, 145):
            try:
                min_val = limpiar_numero(df_params.iloc[i, 27])
                max_val = limpiar_numero(df_params.iloc[i, 28])
                var_val = limpiar_numero(df_params.iloc[i, 29])
                fija_val = limpiar_numero(df_params.iloc[i, 30])
                
                if max_val > 1e15:
                    max_val = float('inf')
                
                if min_val > 0 or max_val > 0 or var_val > 0 or fija_val > 0:
                    parametros['Negociacion']['Transaccion']['Chile'].append({
                        'min': min_val,
                        'max': max_val,
                        'var': var_val,
                        'fija': fija_val
                    })
            except:
                pass
        
        # Si no hay parámetros, usar defaults
        for producto in ['Acceso', 'Transaccion']:
            for pais in ['Colombia', 'Peru', 'Chile']:
                if not parametros['Negociacion'][producto][pais]:
                    # Defaults básicos
                    parametros['Negociacion'][producto][pais] = [
                        {'min': 0, 'max': 5_000_000, 'var': 0, 'fija': 500},
                        {'min': 5_000_001, 'max': 15_000_000, 'var': 0, 'fija': 1500},
                        {'min': 15_000_001, 'max': float('inf'), 'var': 0, 'fija': 3000}
                    ]
        
        return parametros
        
    except Exception as e:
        st.error(f"⚠️ Error al leer parámetros: {str(e)}")
        st.info("Usando parámetros por defecto")
        
        # Retornar parámetros por defecto
        for producto in ['Acceso', 'Transaccion']:
            for pais in ['Colombia', 'Peru', 'Chile']:
                parametros['Negociacion'][producto][pais] = [
                    {'min': 0, 'max': 5_000_000, 'var': 0, 'fija': 500},
                    {'min': 5_000_001, 'max': 15_000_000, 'var': 0, 'fija': 1500},
                    {'min': 15_000_001, 'max': float('inf'), 'var': 0, 'fija': 3000}
                ]
        
        return parametros

# ==================== CÁLCULOS ====================
def calcular_ingreso(monto, tramos):
    """
    Calcula el ingreso según los tramos
    Formula: (Monto × Variable%) + Fija
    """
    if not tramos or len(tramos) == 0:
        return 0.0
    
    for tramo in tramos:
        min_val = tramo['min']
        max_val = tramo['max']
        
        if min_val <= monto < max_val or (max_val == float('inf') and monto >= min_val):
            return (monto * tramo['var'] / 100) + tramo['fija']
    
    # Si no encuentra tramo, usar el último
    ultimo = tramos[-1]
    return (monto * ultimo['var'] / 100) + ultimo['fija']

def simular_con_parametros(df_datos, parametros):
    """Simula los ingresos con los nuevos parámetros"""
    
    resultados = []
    
    for _, row in df_datos.iterrows():
        broker = row['Broker']
        pais = row['Pais']
        monto = row['Monto_USD']
        
        # Normalizar nombre del país
        pais_key = pais
        if pais == 'Perú':
            pais_key = 'Peru'
        
        # Valores reales y propuesta original
        acc_real = row['Acceso_Real']
        trans_real = row['Trans_Real']
        acc_prop = row['Acceso_Propuesta']
        trans_prop = row['Trans_Propuesta']
        
        # Calcular simulados
        tramos_acc = parametros['Negociacion']['Acceso'].get(pais_key, [])
        tramos_trans = parametros['Negociacion']['Transaccion'].get(pais_key, [])
        
        acc_sim = calcular_ingreso(monto, tramos_acc)
        trans_sim = calcular_ingreso(monto, tramos_trans)
        
        # Calcular totales
        total_real = acc_real + trans_real
        total_prop = acc_prop + trans_prop
        total_sim = acc_sim + trans_sim
        
        # Calcular diferencias
        diff_vs_real = total_sim - total_real
        diff_vs_prop = total_sim - total_prop
        
        # Calcular BPS
        bps_real = calcular_bps(total_real, monto)
        bps_prop = calcular_bps(total_prop, monto)
        bps_sim = calcular_bps(total_sim, monto)
        
        resultados.append({
            'Broker': broker,
            'Pais': pais,
            'Monto_USD': monto,
            # Acceso
            'Acc_Real': acc_real,
            'Acc_Propuesta': acc_prop,
            'Acc_Simulado': acc_sim,
            # Transacción
            'Trans_Real': trans_real,
            'Trans_Propuesta': trans_prop,
            'Trans_Simulado': trans_sim,
            # Totales
            'Total_Real': total_real,
            'Total_Propuesta': total_prop,
            'Total_Simulado': total_sim,
            # Diferencias
            'Diff_vs_Real': diff_vs_real,
            'Diff_vs_Propuesta': diff_vs_prop,
            # BPS
            'BPS_Real': bps_real,
            'BPS_Propuesta': bps_prop,
            'BPS_Simulado': bps_sim
        })
    
    return pd.DataFrame(resultados)

# ==================== UI PRINCIPAL ====================
def main():
    # Header
    st.markdown('<h1 class="main-header">🚀 SIMULADOR TARIFARIO RV SUPREMO</h1>', 
                unsafe_allow_html=True)
    st.caption("Versión ultra robusta y funcional | Edita parámetros y ve resultados en tiempo real")
    
    # Sidebar
    with st.sidebar:
        st.markdown("## ⚙️ Configuración")
        
        archivo = st.file_uploader(
            "📁 Cargar Excel",
            type=['xlsx'],
            help="Sube tu archivo de Modelamiento Estructura Tarifaria"
        )
        
        if archivo is None:
            st.info("👆 Por favor carga tu archivo Excel para comenzar")
            st.stop()
        
        # Cargar datos
        with st.spinner("🔄 Cargando datos..."):
            df_datos = cargar_datos_negociacion(archivo)
            if df_datos is None or len(df_datos) == 0:
                st.error("❌ No se pudieron cargar los datos del Excel")
                st.stop()
            
            parametros_base = cargar_parametros_excel(archivo)
        
        st.success(f"✅ {len(df_datos)} brokers cargados")
        
        st.markdown("---")
        st.markdown("### 🎯 Filtros")
        
        paises = ['Todos'] + sorted(df_datos['Pais'].unique().tolist())
        pais_filtro = st.selectbox("🌎 País", paises)
        
        producto = st.selectbox(
            "📊 Producto",
            ['Negociación - Acceso', 'Negociación - Transacción', 'Negociación - Total']
        )
        
        st.markdown("---")
        st.markdown("### ⚙️ Editar Parámetros")
        
        modo_edicion = st.checkbox("✏️ Activar modo edición", value=False)
    
    # Filtrar datos
    if pais_filtro == 'Todos':
        df_filtrado = df_datos.copy()
    else:
        df_filtrado = df_datos[df_datos['Pais'] == pais_filtro].copy()
    
    # Edición de parámetros (si está activada)
    parametros_editados = parametros_base.copy()
    
    if modo_edicion:
        with st.sidebar:
            st.markdown("#### 📝 Editar Tramos")
            
            pais_edit = st.selectbox(
                "País a editar",
                ['Colombia', 'Peru', 'Chile'],
                key='pais_edit'
            )
            
            producto_key = 'Acceso' if 'Acceso' in producto else 'Transaccion'
            
            tramos_actuales = parametros_base['Negociacion'][producto_key][pais_edit]
            
            st.write(f"**{producto_key} - {pais_edit}**")
            
            # Crear DataFrame para edición
            df_tramos = pd.DataFrame(tramos_actuales)
            df_tramos['max'] = df_tramos['max'].replace(float('inf'), 999999999999)
            df_tramos.columns = ['Mínimo USD', 'Máximo USD', 'Variable %', 'Fija USD']
            
            # Editor
            df_editado = st.data_editor(
                df_tramos,
                num_rows="dynamic",
                use_container_width=True,
                key=f"editor_{producto_key}_{pais_edit}"
            )
            
            # Convertir de vuelta
            tramos_nuevos = []
            for _, row in df_editado.iterrows():
                max_val = float('inf') if row['Máximo USD'] >= 999999999999 else row['Máximo USD']
                tramos_nuevos.append({
                    'min': row['Mínimo USD'],
                    'max': max_val,
                    'var': row['Variable %'],
                    'fija': row['Fija USD']
                })
            
            parametros_editados['Negociacion'][producto_key][pais_edit] = tramos_nuevos
            
            if st.button("🔄 Aplicar Cambios", type="primary", use_container_width=True):
                st.rerun()
    
    # Simular con parámetros
    with st.spinner("🔄 Calculando simulación..."):
        df_resultados = simular_con_parametros(df_filtrado, parametros_editados)
    
    # KPIs Principales
    st.markdown("### 💰 KPIs Principales")
    
    total_monto = df_resultados['Monto_USD'].sum()
    total_real = df_resultados['Total_Real'].sum()
    total_prop = df_resultados['Total_Propuesta'].sum()
    total_sim = df_resultados['Total_Simulado'].sum()
    
    var_vs_real = ((total_sim - total_real) / total_real * 100) if total_real > 0 else 0
    var_vs_prop = ((total_sim - total_prop) / total_prop * 100) if total_prop > 0 else 0
    
    bps_sim = calcular_bps(total_sim, total_monto)
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("💵 Monto Total", f"${total_monto/1e6:.2f}M")
    
    with col2:
        st.metric("📊 Ingreso Real", f"${total_real/1e6:.2f}M")
    
    with col3:
        st.metric(
            "🎯 Ingreso Simulado",
            f"${total_sim/1e6:.2f}M",
            delta=f"{var_vs_real:+.1f}% vs Real"
        )
    
    with col4:
        st.metric("📈 BPS Simulado", f"{bps_sim:.2f} bps")
    
    st.markdown("---")
    
    # Tabs principales
    tab1, tab2, tab3 = st.tabs(["📊 Comparativa", "🔍 Detalle", "📈 Análisis"])
    
    with tab1:
        col_g1, col_g2 = st.columns(2)
        
        with col_g1:
            st.markdown("#### 📊 Comparación de Ingresos")
            
            fig = go.Figure(data=[
                go.Bar(
                    name='Real',
                    x=['Total'],
                    y=[total_real],
                    marker_color='#e74c3c',
                    text=[f'${total_real/1e6:.2f}M'],
                    textposition='outside'
                ),
                go.Bar(
                    name='Propuesta',
                    x=['Total'],
                    y=[total_prop],
                    marker_color='#f39c12',
                    text=[f'${total_prop/1e6:.2f}M'],
                    textposition='outside'
                ),
                go.Bar(
                    name='Simulado',
                    x=['Total'],
                    y=[total_sim],
                    marker_color='#27ae60',
                    text=[f'${total_sim/1e6:.2f}M'],
                    textposition='outside'
                )
            ])
            
            fig.update_layout(
                height=400,
                barmode='group',
                showlegend=True,
                yaxis_title="Ingresos (USD)"
            )
            
            st.plotly_chart(fig, use_container_width=True)
        
        with col_g2:
            st.markdown("#### 🌎 Ingresos por País")
            
            df_por_pais = df_resultados.groupby('Pais').agg({
                'Total_Real': 'sum',
                'Total_Propuesta': 'sum',
                'Total_Simulado': 'sum'
            }).reset_index()
            
            fig2 = go.Figure(data=[
                go.Bar(
                    name='Real',
                    x=df_por_pais['Pais'],
                    y=df_por_pais['Total_Real'],
                    marker_color='#e74c3c'
                ),
                go.Bar(
                    name='Propuesta',
                    x=df_por_pais['Pais'],
                    y=df_por_pais['Total_Propuesta'],
                    marker_color='#f39c12'
                ),
                go.Bar(
                    name='Simulado',
                    x=df_por_pais['Pais'],
                    y=df_por_pais['Total_Simulado'],
                    marker_color='#27ae60'
                )
            ])
            
            fig2.update_layout(
                height=400,
                barmode='group',
                showlegend=True,
                yaxis_title="Ingresos (USD)"
            )
            
            st.plotly_chart(fig2, use_container_width=True)
    
    with tab2:
        st.markdown("#### 📋 Detalle por Broker")
        
        # Preparar tabla
        df_display = df_resultados[[
            'Broker', 'Pais', 'Monto_USD',
            'Total_Real', 'Total_Simulado', 'Diff_vs_Real',
            'BPS_Real', 'BPS_Simulado'
        ]].copy()
        
        df_display = df_display.sort_values('Diff_vs_Real', ascending=False)
        
        # Formatear
        df_display['Monto_USD'] = df_display['Monto_USD'].apply(lambda x: f"${x:,.0f}")
        df_display['Total_Real'] = df_display['Total_Real'].apply(lambda x: f"${x:,.2f}")
        df_display['Total_Simulado'] = df_display['Total_Simulado'].apply(lambda x: f"${x:,.2f}")
        df_display['Diff_vs_Real'] = df_display['Diff_vs_Real'].apply(lambda x: f"${x:,.2f}")
        df_display['BPS_Real'] = df_display['BPS_Real'].apply(lambda x: f"{x:.2f}")
        df_display['BPS_Simulado'] = df_display['BPS_Simulado'].apply(lambda x: f"{x:.2f}")
        
        df_display.columns = [
            'Broker', 'País', 'Monto',
            'Ing. Real', 'Ing. Simulado', 'Diferencia',
            'BPS Real', 'BPS Simulado'
        ]
        
        st.dataframe(df_display, use_container_width=True, height=500)
    
    with tab3:
        st.markdown("#### 🏆 Top 10 Mayores Cambios")
        
        top10 = df_resultados.nlargest(10, 'Diff_vs_Real')
        
        fig3 = go.Figure()
        
        fig3.add_trace(go.Bar(
            name='Real',
            x=top10['Broker'],
            y=top10['Total_Real'],
            marker_color='#e74c3c'
        ))
        
        fig3.add_trace(go.Bar(
            name='Simulado',
            x=top10['Broker'],
            y=top10['Total_Simulado'],
            marker_color='#27ae60'
        ))
        
        fig3.update_layout(
            height=500,
            barmode='group',
            xaxis_tickangle=-45,
            yaxis_title="Ingresos (USD)"
        )
        
        st.plotly_chart(fig3, use_container_width=True)
    
    # Exportar
    st.markdown("---")
    st.markdown("### 📥 Exportar Resultados")
    
    col_exp1, col_exp2 = st.columns(2)
    
    with col_exp1:
        csv = df_resultados.to_csv(index=False).encode('utf-8')
        st.download_button(
            "📥 Descargar Detalle Completo (CSV)",
            csv,
            "simulacion_completa.csv",
            "text/csv",
            use_container_width=True
        )
    
    with col_exp2:
        resumen = df_resultados.groupby('Pais').agg({
            'Monto_USD': 'sum',
            'Total_Real': 'sum',
            'Total_Simulado': 'sum',
            'Diff_vs_Real': 'sum'
        }).reset_index()
        
        csv_resumen = resumen.to_csv(index=False).encode('utf-8')
        st.download_button(
            "📥 Descargar Resumen por País (CSV)",
            csv_resumen,
            "resumen_por_pais.csv",
            "text/csv",
            use_container_width=True
        )

if __name__ == "__main__":
    main()
