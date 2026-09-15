# -*- coding: utf-8 -*-
"""
Created on Tue Sep 15 12:24:15 2026

@author: Henry
"""

import pandas as pd
import streamlit as st
import os

# Configuración inicial de la página web del dashboard
st.set_page_config(
    page_title="Control de Izaje de Postes",
    page_icon="🏗️",
    layout="wide"
)

# Ruta de tu archivo Excel
ruta_archivo = r'C:\Users\Henry\Documents\HZT\QUANTUM\TUMBES\EJECUCION\PLANTADO DE POSTES\OBSERVACIONES DEL IZAJE DE POSTES.xlsx'

# Función para cargar datos de forma limpia
@st.cache_data
def cargar_datos(path):
    if not os.path.exists(path):
        return None
    df = pd.read_excel(path, header=2)
    df = df.dropna(subset=['N° DE POSTE/CAMARA', 'ESTADO'])
    df = df[df['ESTADO'] != 'ESTADO']
    return df

df = cargar_datos(ruta_archivo)

# Título principal
st.title("📊 Dashboard de Control: Izaje de Postes")
st.markdown("---")

if df is None:
    st.error(f"No se encontró el archivo en la ruta: {ruta_archivo}")
else:
    # --- BARRA LATERAL (FILTROS) ---
    st.sidebar.header("Filtros de Búsqueda")
    
    # Filtro por Estado
    estados_disponibles = ['TODOS'] + list(df['ESTADO'].unique())
    estado_seleccionado = st.sidebar.selectbox("Filtrar por Estado:", estados_disponibles)
    
    # Filtro por Tipo de Terreno
    terrenos_disponibles = ['TODOS'] + list(df['TERRENO'].unique())
    terreno_seleccionado = st.sidebar.selectbox("Filtrar por Tipo de Terreno:", terrenos_disponibles)
    
    # Aplicar filtros
    df_filtrado = df.copy()
    if estado_seleccionado != 'TODOS':
        df_filtrado = df_filtrado[df_filtrado['ESTADO'] == estado_seleccionado]
    if terreno_seleccionado != 'TODOS':
        df_filtrado = df_filtrado[df_filtrado['TERRENO'] == terreno_seleccionado]

    # --- MÉTRICAS PRINCIPALES (KPIs) ---
    total_postes = len(df)
    pendientes = len(df[df['ESTADO'] == 'PENDIENTE'])
    ok = len(df[df['ESTADO'] == 'OK'])
    porcentaje_avance = (ok / total_postes) * 100 if total_postes > 0 else 0

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total de Postes", total_postes)
    col2.metric("Pendientes", pendientes, delta_color="inverse")
    col3.metric("Aprobados (OK)", ok)
    col4.metric("Avance General", f"{porcentaje_avance:.1f}%")
    
    st.markdown("---")

    # --- GRÁFICAS VISUALES ---
    col_g1, col_g2 = st.columns(2)
    
    with col_g1:
        st.subheader("Estado de las Instalaciones")
        conteo_estados = df['ESTADO'].value_counts()
        st.bar_chart(conteo_estados)
        
    with col_g2:
        st.subheader("Distribución por Tipo de Terreno")
        conteo_terrenos = df['TERRENO'].value_counts()
        st.bar_chart(conteo_terrenos)

    st.markdown("---")

    # --- TABLA DE DATOS INTERACTIVA ---
    st.subheader(f"📋 Detalle de Registros ({len(df_filtrado)} postes encontrados)")
    
    # Mostrar la tabla interactiva (puedes ordenar columnas haciendo clic en ellas)
    st.dataframe(df_filtrado, use_container_width=True)
    
    # Botón para refrescar datos si modificaste el Excel recientemente
    if st.button("🔄 Actualizar Datos desde el Excel"):
        st.cache_data.clear()
        st.rerun()
        