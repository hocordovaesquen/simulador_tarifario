"""
DASHBOARD INTERACTIVO - SIMULADOR TARIFARIO RV
==============================================

INSTALACIÓN:
pip install streamlit pandas numpy openpyxl plotly

EJECUCIÓN:
streamlit run app.py

Luego sube tu archivo Excel en la interfaz.
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
import re

# Configuración de la página
st.set_page_config(
    page_title="Simulador Tarifario RV",
    page_icon="🚀",
    layout="wide"
)

# =============================================================================
# FUNCIONES AUXILIARES - LIMPIEZA Y PARSING
# =============================================================================

def limpiar_numero(valor):
    """Convierte cualquier valor a número, manejando formatos raros"""
    if pd.isna(valor):
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    
    # Convertir a string y limpiar
    valor_str = str(valor).strip()
    # Remover símbolos de moneda, espacios, etc
    valor_str = valor_str.replace('$', '').replace(',', '').replace('\xa0', '')
    valor_str = valor_str.replace('%', '')
    
    try:
        return float(valor_str)
    except:
        return 0.0

def normalizar_texto(texto):
    """Normaliza texto para comparaciones (quita acentos, minúsculas)"""
    if pd.isna(texto):
        return ""
    texto = str(texto).lower().strip()
    # Normalizar algunos casos comunes
    texto = texto.replace('perú', 'peru').replace('transacción', 'transaccion')
    return texto

def fuzzy_match_columna(columnas, palabras_clave):
    """Encuentra columna que contenga alguna de las palabras clave"""
    columnas_norm = [normalizar_texto(c) for c in columnas]
    for palabra in palabras_clave:
        palabra_norm = normalizar_texto(palabra)
        for i, col_norm in enumerate(columnas_norm):
            if palabra_norm in col_norm:
                return columnas[i]
    return None

# =============================================================================
# FUNCIONES DE LECTURA DEL EXCEL
# =============================================================================

@st.cache_data
def leer_negociacion(archivo):
    """Lee la hoja A.3 BBDD Neg"""
    df = pd.read_excel(archivo, sheet_name='A.3 BBDD Neg', header=0)
    
    # Identificar columnas clave con fuzzy matching
    col_pais = fuzzy_match_columna(df.columns, ['pais', 'país', 'country'])
    col_broker = fuzzy_match_columna(df.columns, ['cliente estandar', 'corredor', 'broker', 'cliente'])
    col_monto = fuzzy_match_columna(df.columns, ['monto usd', 'monto', 'valor usd'])
    
    # Buscar columnas de ingresos
    columnas_ingreso = {}
    for col in df.columns:
        col_lower = str(col).lower()
        
        # Acceso Real
        if 'acceso' in col_lower and 'real' in col_lower:
            columnas_ingreso['acceso_real'] = col
        # Acceso Proyectado
        elif 'acceso' in col_lower and 'proyectado' in col_lower:
            columnas_ingreso['acceso_proyectado'] = col
        # Transacción Real
        elif 'trans' in col_lower and 'real' in col_lower:
            columnas_ingreso['trans_real'] = col
        # Transacción Proyectado
        elif 'trans' in col_lower and 'proyectado' in col_lower:
            columnas_ingreso['trans_proyectado'] = col
    
    if not all([col_pais, col_broker, col_monto]):
        st.error(f"No se pudieron identificar columnas clave. Encontradas: País={col_pais}, Broker={col_broker}, Monto={col_monto}")
        return None, None, None, None
    
    # Crear dataframe limpio
    df_limpio = pd.DataFrame()
    df_limpio['Pais'] = df[col_pais].apply(normalizar_texto)
    df_limpio['Broker'] = df[col_broker]
    df_limpio['Monto_USD'] = df[col_monto].apply(limpiar_numero)
    
    # Agregar columnas de ingreso
    for key, col_name in columnas_ingreso.items():
        df_limpio[key] = df[col_name].apply(limpiar_numero)
    
    # Excluir filas de totales (índices 9, 34, 54 en base 0)
    filas_totales_idx = [9, 34, 54]
    df_brokers = df_limpio[~df_limpio.index.isin(filas_totales_idx)].copy()
    
    # Guardar filas de totales para validación
    totales_por_pais = {}
    paises = ['colombia', 'peru', 'chile']
    for idx, pais in zip(filas_totales_idx, paises):
        if idx < len(df_limpio):
            totales_por_pais[pais] = df_limpio.iloc[idx].to_dict()
    
    # Filtrar brokers válidos (que tengan nombre y monto > 0)
    df_brokers = df_brokers[df_brokers['Broker'].notna() & (df_brokers['Monto_USD'] > 0)].copy()
    
    return df_brokers, columnas_ingreso, totales_por_pais, col_pais

@st.cache_data
def leer_parametros(archivo):
    """Lee la hoja 1. Parametros desde columna R hacia adelante"""
    df = pd.read_excel(archivo, sheet_name='1. Parametros', header=None)
    
    # Buscar columna R (índice 17 en base 0)
    col_inicio = 17
    
    # Detectar dónde termina el bloque (primera columna completamente vacía después de R)
    col_fin = col_inicio
    for i in range(col_inicio, df.shape[1]):
        if df.iloc[:, i].isna().all():
            col_fin = i
            break
    else:
        col_fin = df.shape[1]
    
    # Extraer bloque
    bloque = df.iloc[:, col_inicio:col_fin]
    
    # Parsear estructura (asumiendo layout matricial)
    # Buscar filas con "Acceso" y "Transacción"
    parametros = {
        'colombia': {'acceso': [], 'transaccion': []},
        'peru': {'acceso': [], 'transaccion': []},
        'chile': {'acceso': [], 'transaccion': []}
    }
    
    # Lógica de parsing: detectar automáticamente estructura
    # Por simplicidad, usaremos valores por defecto basados en estructura típica
    # El usuario puede editarlos luego
    
    # Tramos por defecto (ajusta según tu Excel real)
    tramos_default = [
        {'rango_min': 0, 'rango_max': 10_000_000, 'prop_var': 0.15, 'prop_fija': 5000},
        {'rango_min': 10_000_000, 'rango_max': 50_000_000, 'prop_var': 0.12, 'prop_fija': 15000},
        {'rango_min': 50_000_000, 'rango_max': float('inf'), 'prop_var': 0.10, 'prop_fija': 30000}
    ]
    
    for pais in parametros.keys():
        parametros[pais]['acceso'] = tramos_default.copy()
        parametros[pais]['transaccion'] = tramos_default.copy()
    
    return parametros

# =============================================================================
# FUNCIONES DE CÁLCULO
# =============================================================================

def calcular_ingreso_simulado(monto, tramos):
    """Calcula ingreso simulado según tramos"""
    for tramo in tramos:
        if tramo['rango_min'] <= monto <= tramo['rango_max']:
            return monto * (tramo['prop_var'] / 100) + tramo['prop_fija']
    return 0.0

def calcular_tarifa_implicita_bps(ingreso, monto):
    """Calcula tarifa implícita en basis points"""
    if monto == 0:
        return 0.0
    return (ingreso / monto) * 10_000

def agregar_y_calcular(df_brokers, columnas_ingreso, parametros, pais_filtro, producto):
    """Agrega datos por broker y calcula simulaciones"""
    
    # Filtrar por país si no es "Todos"
    if pais_filtro != 'todos':
        df_filtrado = df_brokers[df_brokers['Pais'] == pais_filtro].copy()
    else:
        df_filtrado = df_brokers.copy()
    
    # Agrupar por broker y país
    agrupado = df_filtrado.groupby(['Broker', 'Pais']).agg({
        'Monto_USD': 'sum'
    }).reset_index()
    
    # Determinar columna de ingreso real según producto
    if producto == 'Negociación - Acceso':
        col_real = 'acceso_real' if 'acceso_real' in columnas_ingreso else 'acceso_proyectado'
        producto_key = 'acceso'
    else:  # Transacción
        col_real = 'trans_real' if 'trans_real' in columnas_ingreso else 'trans_proyectado'
        producto_key = 'transaccion'
    
    # Sumar ingresos reales
    if col_real in columnas_ingreso:
        ingresos_reales = df_filtrado.groupby(['Broker', 'Pais'])[col_real].sum().reset_index()
        agrupado = agrupado.merge(ingresos_reales, on=['Broker', 'Pais'], how='left')
        agrupado.rename(columns={col_real: 'Ingreso_Real'}, inplace=True)
    else:
        agrupado['Ingreso_Real'] = 0.0
    
    # Calcular ingreso simulado
    resultados = []
    for _, row in agrupado.iterrows():
        pais = row['Pais']
        monto = row['Monto_USD']
        ingreso_real = row.get('Ingreso_Real', 0)
        
        # Obtener tramos del país
        tramos = parametros.get(pais, {}).get(producto_key, [])
        
        # Calcular simulado
        ingreso_sim = calcular_ingreso_simulado(monto, tramos)
        diferencia = ingreso_sim - ingreso_real
        
        if ingreso_real != 0:
            var_pct = (diferencia / ingreso_real) * 100
        elif ingreso_sim > 0:
            var_pct = float('inf')
        else:
            var_pct = 0.0
        
        # Tarifa implícita
        tarifa_real_bps = calcular_tarifa_implicita_bps(ingreso_real, monto)
        tarifa_sim_bps = calcular_tarifa_implicita_bps(ingreso_sim, monto)
        
        resultados.append({
            'Broker': row['Broker'],
            'Pais': row['Pais'].title(),
            'Monto_Negociado': monto,
            'Ingreso_Real': ingreso_real,
            'Ingreso_Simulado': ingreso_sim,
            'Diferencia': diferencia,
            'Var_Pct': var_pct,
            'Tarifa_Real_bps': tarifa_real_bps,
            'Tarifa_Simulada_bps': tarifa_sim_bps
        })
    
    return pd.DataFrame(resultados)

# =============================================================================
# INTERFAZ PRINCIPAL
# =============================================================================

def main():
    st.title("🚀 Simulador Tarifario RV")
    st.markdown("---")
    
    # Sidebar - Upload y configuración
    with st.sidebar:
        st.header("⚙️ Configuración")
        
        # Upload
        archivo = st.file_uploader("📁 Sube tu Excel", type=['xlsx'])
        
        if archivo is None:
            st.info("👆 Por favor sube un archivo Excel para comenzar")
            st.stop()
        
        # Leer datos
        with st.spinner("Leyendo archivo..."):
            df_brokers, columnas_ingreso, totales_por_pais, col_pais = leer_negociacion(archivo)
            parametros = leer_parametros(archivo)
        
        if df_brokers is None:
            st.error("Error al leer el archivo")
            st.stop()
        
        st.success(f"✅ {len(df_brokers)} brokers cargados")
        
        # Filtros
        st.markdown("### 🔍 Filtros")
        pais_filtro = st.selectbox(
            "País",
            ['todos', 'colombia', 'peru', 'chile'],
            format_func=lambda x: x.title()
        )
        
        producto = st.selectbox(
            "Producto",
            ['Negociación - Acceso', 'Negociación - Transacción']
        )
        
        # Editor de tramos
        st.markdown("### 📊 Editar Tramos")
        
        producto_key = 'acceso' if 'Acceso' in producto else 'transaccion'
        
        if pais_filtro == 'todos':
            paises_editar = ['colombia', 'peru', 'chile']
        else:
            paises_editar = [pais_filtro]
        
        for pais in paises_editar:
            with st.expander(f"🌎 {pais.title()}"):
                tramos_df = pd.DataFrame(parametros[pais][producto_key])
                
                # Editor
                tramos_editados = st.data_editor(
                    tramos_df,
                    num_rows="dynamic",
                    use_container_width=True,
                    key=f"editor_{pais}_{producto_key}"
                )
                
                # Guardar cambios
                parametros[pais][producto_key] = tramos_editados.to_dict('records')
        
        # Toggles opcionales
        st.markdown("### 🎛️ Opciones")
        exoneracion = st.checkbox("Exoneración 100%", help="Fuerza prop_var=0 y prop_fija=0")
        
        if st.button("🔄 Recalcular", type="primary", use_container_width=True):
            st.rerun()
    
    # Aplicar exoneración si está activa
    if exoneracion:
        for pais in parametros.keys():
            for prod in parametros[pais].keys():
                for tramo in parametros[pais][prod]:
                    tramo['prop_var'] = 0
                    tramo['prop_fija'] = 0
    
    # Calcular resultados
    df_resultados = agregar_y_calcular(df_brokers, columnas_ingreso, parametros, pais_filtro, producto)
    
    # =============================================================================
    # VISUALIZACIONES
    # =============================================================================
    
    # KPIs
    total_real = df_resultados['Ingreso_Real'].sum()
    total_simulado = df_resultados['Ingreso_Simulado'].sum()
    diferencia_total = total_simulado - total_real
    var_pct_total = (diferencia_total / total_real * 100) if total_real != 0 else 0
    
    total_monto = df_resultados['Monto_Negociado'].sum()
    tarifa_promedio = calcular_tarifa_implicita_bps(total_simulado, total_monto)
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric(
            "💰 Ingreso Real",
            f"${total_real:,.0f}",
            help="Total de ingresos actuales/proyectados"
        )
    
    with col2:
        st.metric(
            "🎯 Ingreso Simulado",
            f"${total_simulado:,.0f}",
            delta=f"{var_pct_total:+.1f}%"
        )
    
    with col3:
        st.metric(
            "📊 Diferencia",
            f"${diferencia_total:,.0f}",
            help="Simulado - Real"
        )
    
    with col4:
        st.metric(
            "📈 Tarifa Implícita",
            f"{tarifa_promedio:.1f} bps",
            help="Basis points sobre monto negociado"
        )
    
    st.markdown("---")
    
    # Gráficos
    col_g1, col_g2 = st.columns(2)
    
    with col_g1:
        st.subheader("📊 Comparación Ingresos")
        fig_comp = go.Figure(data=[
            go.Bar(name='Real', x=['Real', 'Simulado'], y=[total_real, 0], marker_color='#667eea'),
            go.Bar(name='Simulado', x=['Real', 'Simulado'], y=[0, total_simulado], marker_color='#27ae60')
        ])
        fig_comp.update_layout(barmode='group', height=300)
        st.plotly_chart(fig_comp, use_container_width=True)
    
    with col_g2:
        st.subheader("🏆 Top 10 Brokers")
        top10 = df_resultados.nlargest(10, 'Monto_Negociado')
        fig_top10 = go.Figure(data=[
            go.Bar(name='Real', x=top10['Broker'], y=top10['Ingreso_Real'], marker_color='#667eea'),
            go.Bar(name='Simulado', x=top10['Broker'], y=top10['Ingreso_Simulado'], marker_color='#27ae60')
        ])
        fig_top10.update_layout(height=300)
        st.plotly_chart(fig_top10, use_container_width=True)
    
    st.markdown("---")
    
    # Tabla detalle
    st.subheader("📋 Detalle por Broker")
    
    # Formatear para mostrar
    df_display = df_resultados.copy()
    df_display['Monto_Negociado'] = df_display['Monto_Negociado'].apply(lambda x: f"${x:,.0f}")
    df_display['Ingreso_Real'] = df_display['Ingreso_Real'].apply(lambda x: f"${x:,.2f}")
    df_display['Ingreso_Simulado'] = df_display['Ingreso_Simulado'].apply(lambda x: f"${x:,.2f}")
    df_display['Diferencia'] = df_display['Diferencia'].apply(lambda x: f"${x:,.2f}")
    df_display['Var_Pct'] = df_display['Var_Pct'].apply(lambda x: f"{x:.2f}%" if x != float('inf') else "∞%")
    df_display['Tarifa_Real_bps'] = df_display['Tarifa_Real_bps'].apply(lambda x: f"{x:.1f}")
    df_display['Tarifa_Simulada_bps'] = df_display['Tarifa_Simulada_bps'].apply(lambda x: f"{x:.1f}")
    
    st.dataframe(df_display, use_container_width=True, height=400)
    
    # Exportar
    st.markdown("---")
    col_exp1, col_exp2 = st.columns(2)
    
    with col_exp1:
        csv_detalle = df_resultados.to_csv(index=False).encode('utf-8')
        st.download_button(
            "📥 Descargar Detalle CSV",
            csv_detalle,
            "detalle_brokers.csv",
            "text/csv",
            use_container_width=True
        )
    
    with col_exp2:
        # Resumen por país
        resumen_pais = df_resultados.groupby('Pais').agg({
            'Monto_Negociado': 'sum',
            'Ingreso_Real': 'sum',
            'Ingreso_Simulado': 'sum'
        }).reset_index()
        
        csv_resumen = resumen_pais.to_csv(index=False).encode('utf-8')
        st.download_button(
            "📥 Descargar Resumen por País",
            csv_resumen,
            "resumen_pais.csv",
            "text/csv",
            use_container_width=True
        )

if __name__ == "__main__":
    main()
