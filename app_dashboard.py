# -*- coding: utf-8 -*-
"""
Created on Tue Sep 15 12:24:15 2026

@author: Henry Z
"""

import pandas as pd
import streamlit as st

# Configuración inicial de la página web del dashboard (modo ancho)
st.set_page_config(
    page_title="Control de Izaje de Postes",
    page_icon="🏗️",
    layout="wide"
)

# Estilos CSS personalizados para tarjetas y diseño ejecutivo
st.markdown("""
    <style>
        .metric-card {
            background-color: #f8f9fa;
            border: 1px solid #e9ecef;
            padding: 15px;
            border-radius: 10px;
            text-align: center;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        }
        .metric-title {
            font-size: 14px;
            color: #6c757d;
            font-weight: 600;
            text-transform: uppercase;
        }
        .metric-value {
            font-size: 28px;
            color: #212529;
            font-weight: bold;
        }
    </style>
""", unsafe_allow_html=True)

# Enlace directo de exportación CSV optimizado para tu Google Sheet
sheet_url = "https://docs.google.com/spreadsheets/d/1HTEq01G5xgyMCNrocYvOeKIXTV0xhsKg/export?format=csv"

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

# Título Principal con estilo
st.title("📊 Dashboard de Control: Izaje de Postes en Vivo")
st.markdown("Monitoreo en tiempo real de avance, sectores y levantamiento de observaciones en obra.")
st.markdown("---")

if df is None or len(df) == 0:
    st.error("No se pudo cargar la información desde Google Sheets.")
    if error_detallado:
        st.info(f"Detalle técnico del error: {error_detallado}")
    st.warning("⚠️ Asegúrate de que el Google Sheet esté compartido como 'Cualquier usuario que tenga el vínculo' (Lector o Editor).")
else:
    # --- BARRA LATERAL (FILTROS) ---
    st.sidebar.header("🔍 Filtros de Búsqueda")
    
    # Filtro por Zona (si la columna existe)
    if 'ZONA' in df.columns:
        zonas_disponibles = ['TODOS'] + list(df['ZONA'].dropna().unique())
        zona_seleccionada = st.sidebar.selectbox("Filtrar por Zona / Sector:", zonas_disponibles)
    else:
        zona_seleccionada = 'TODOS'
    
    estados_disponibles = ['TODOS'] + list(df['ESTADO'].unique())
    estado_seleccionado = st.sidebar.selectbox("Filtrar por Estado:", estados_disponibles)
    
    terrenos_disponibles = ['TODOS'] + list(df['TERRENO'].dropna().unique()) if 'TERRENO' in df.columns else ['TODOS']
    terreno_seleccionado = st.sidebar.selectbox("Filtrar por Tipo de Terreno:", terrenos_disponibles)
    
    # Aplicar filtros
    df_filtrado = df.copy()
    if zona_seleccionada != 'TODOS' and 'ZONA' in df.columns:
        df_filtrado = df_filtrado[df_filtrado['ZONA'] == zona_seleccionada]
    if estado_seleccionado != 'TODOS':
        df_filtrado = df_filtrado[df_filtrado['ESTADO'] == estado_seleccionado]
    if terreno_seleccionado != 'TODOS' and 'TERRENO' in df.columns:
        df_filtrado = df_filtrado[df_filtrado['TERRENO'] == terreno_seleccionado]

    # --- MÉTRICAS PRINCIPALES (KPIs) ---
    total_postes = len(df_filtrado)
    pendientes = len(df_filtrado[df_filtrado['ESTADO'] == 'PENDIENTE'])
    ok = len(df_filtrado[df_filtrado['ESTADO'] == 'OK'])
    porcentaje_avance = (ok / total_postes) * 100 if total_postes > 0 else 0

    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown(f"""
            <div class="metric-card">
                <div class="metric-title">Postes en Selección</div>
                <div class="metric-value">🏗️ {total_postes}</div>
            </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown(f"""
            <div class="metric-card">
                <div class="metric-title">Pendientes</div>
                <div class="metric-value" style="color: #d9534f;">⚠️ {pendientes}</div>
            </div>
        """, unsafe_allow_html=True)
    with col3:
        st.markdown(f"""
            <div class="metric-card">
                <div class="metric-title">Aprobados (OK)</div>
                <div class="metric-value" style="color: #5cb85c;">✅ {ok}</div>
            </div>
        """, unsafe_allow_html=True)
    with col4:
        st.markdown(f"""
            <div class="metric-card">
                <div class="metric-title">Avance Sectorial</div>
                <div class="metric-value" style="color: #0275d8;">📈 {porcentaje_avance:.1f}%</div>
            </div>
        """, unsafe_allow_html=True)
    
    st.markdown("<br>", unsafe_allow_html=True)

    # --- GRÁFICAS VISUALES MEJORADAS ---
    col_g1, col_g2 = st.columns(2)
    
    with col_g1:
        st.subheader("📌 Proporción de Estados")
        conteo_estados = df_filtrado['ESTADO'].value_counts()
        st.bar_chart(conteo_estados, color="#0275d8")
        
    with col_g2:
        if 'ZONA' in df.columns:
            st.subheader("🗺️ Distribución por Zonas")
            conteo_zonas = df['ZONA'].value_counts()
            st.bar_chart(conteo_zonas, color="#f0ad4e")
        elif 'TERRENO' in df.columns:
            st.subheader("🌍 Distribución por Tipo de Terreno")
            conteo_terrenos = df['TERRENO'].value_counts()
            st.bar_chart(conteo_terrenos, color="#5cb85c")

    st.markdown("---")

    # --- BUSCADOR RÁPIDO DE POSTE ESPECÍFICO ---
    st.subheader("🔍 Consulta Individual de Poste")
    lista_postes = list(df['N° DE POSTE/CAMARA'].astype(str).unique())
    poste_buscado = st.selectbox("Seleccione o busque el número de poste:", ["-- Seleccionar --"] + lista_postes)
    
    if poste_buscado != "-- Seleccionar --":
        datos_poste = df[df['N° DE POSTE/CAMARA'].astype(str) == str(poste_buscado)]
        st.dataframe(datos_poste, use_container_width=True)

    st.markdown("---")

    # --- TABLA DE DATOS INTERACTIVA COMPLETA ---
    st.subheader(f"📋 Detalle de Registros Filtrados ({len(df_filtrado)} elementos)")
    st.dataframe(df_filtrado, use_container_width=True)
    
    # Botón de refresco manual
    col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
    with col_btn2:
        if st.button("🔄 Refrescar Datos en Vivo desde Google Sheets", use_container_width=True):
            st.cache_data.clear()
            st.rerun()
