# -*- coding: utf-8 -*-
"""
Created on Tue Sep 15 12:24:15 2026

@author: Henry Z
"""

import pandas as pd
import streamlit as st

# Configuración inicial de la página web del dashboard
st.set_page_config(
    page_title="Control de Izaje de Postes",
    page_icon="🏗️",
    layout="wide"
)

# PEGA AQUÍ EL ENLACE CSV OBTENIDO DE "PUBLICAR EN LA WEB"
sheet_url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTKr-q0vtewH2ya17kZjYnTNxW9JS-9kp1ZDtFjGqx7LQcXycDd_Gk_VrnPXS3NTA/pub?output=csv"

@st.cache_data(ttl=60)
def cargar_datos_gsheets(url):
    try:
        df = pd.read_csv(url, header=2)
        df = df.dropna(subset=['N° DE POSTE/CAMARA', 'ESTADO'])
        df = df[df['ESTADO'] != 'ESTADO']
        return df, None
    except Exception as e:
        return None, str(e)

df, error_detallado = cargar_datos_gsheets(sheet_url)

st.title("📊 Dashboard de Control: Izaje de Postes (En Vivo)")
st.markdown("---")

if df is None or len(df) == 0:
    st.error("No se pudo cargar la información desde Google Sheets.")
    if error_detallado:
        st.info(f"Detalle técnico del error: {error_detallado}")
else:
    # --- BARRA LATERAL (FILTROS) ---
    st.sidebar.header("Filtros de Búsqueda")
    
    estados_disponibles = ['TODOS'] + list(df['ESTADO'].unique())
    estado_seleccionado = st.sidebar.selectbox("Filtrar por Estado:", estados_disponibles)
    
    terrenos_disponibles = ['TODOS'] + list(df['TERRENO'].unique())
    terreno_seleccionado = st.sidebar.selectbox("Filtrar por Tipo de Terreno:", terrenos_disponibles)
    
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
    st.dataframe(df_filtrado, use_container_width=True)
    
    if st.button("🔄 Refrescar Datos de la Web"):
        st.cache_data.clear()
        st.rerun()
