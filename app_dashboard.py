# -*- coding: utf-8 -*-
"""
Created on Tue Sep 15 12:24:15 2026

@author: Henry Z
"""

import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

# Configuración inicial de la página web del dashboard (modo ancho)
st.set_page_config(
    page_title="Control de Izaje de Postes",
    page_icon="🏗️",
    layout="wide"
)

# Estilos CSS personalizados para tarjetas ejecutivas y optimización de espacios
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
            font-size: 13px;
            color: #6c757d;
            font-weight: 600;
            text-transform: uppercase;
        }
        .metric-value {
            font-size: 26px;
            color: #212529;
            font-weight: bold;
        }
        /* Reduce espacio superior innecesario en Streamlit */
        .block-container {
            padding-top: 2rem;
            padding-bottom: 2rem;
        }
    </style>
""", unsafe_allow_html=True)

# Enlace directo de exportación CSV optimizado
sheet_url = "https://docs.google.com/spreadsheets/d/1HTEq01G5xgyMCNrocYvOeKIXTV0xhsKg/export?format=csv"

@st.cache_data(ttl=60)
def cargar_datos_gsheets(url):
    try:
        df = pd.read_csv(url, header=None)
        df = df.dropna(how='all')
        if len(df) > 2:
            columnas_genericas = [f"COL_{i}" for i in range(df.shape[1])]
            df.columns = columnas_genericas
        return df, None
    except Exception as e:
        return None, str(e)

df_raw, error_detallado = cargar_datos_gsheets(sheet_url)

# --- ENCABEZADO SUPERIOR OPTIMIZADO (TÍTULO + LOGOTIPO / IMAGEN) ---
col_head1, col_head2 = st.columns([3, 1])

with col_head1:
    st.title("📊 DASHBOARD CONTROL DE POSTES OBSERVADOS")
    st.markdown("Monitoreo en tiempo real de avance y levantamiento de observaciones en obra.")

with col_head2:
    try:
        st.image("logo.png", width=220)
    except:
        st.image("https://images.unsplash.com/photo-1541888946425-d0fbb18f86f7?q=80&w=300&auto=format&fit=crop", width=220, caption="Control de Infraestructura")

st.markdown("---")

if df_raw is None or len(df_raw) == 0:
    st.error("No se pudo cargar la información desde Google Sheets.")
    if error_detallado:
        st.info(f"Detalle técnico del error: {error_detallado}")
else:
    df = df_raw.copy()
    
    col_estado_idx = None
    for col in df.columns:
        valores_str = df[col].astype(str).str.upper()
        if valores_str.str.contains('PENDIENTE|ATENDIDO|CONFORME|OK').any() and col_estado_idx is None:
            col_estado_idx = col

    if col_estado_idx is not None:
        df[col_estado_idx] = df[col_estado_idx].astype(str).str.upper().str.strip()
        df[col_estado_idx] = df[col_estado_idx].replace('OK', 'ATENDIDO')
        df = df[df[col_estado_idx].isin(['PENDIENTE', 'ATENDIDO', 'CONFORME'])]
    
    df = df.reset_index(drop=True)
    df.index = df.index + 1  # Inicia estrictamente en 1

    if df.shape[1] >= 7:
        df.columns = ['ZONA', 'N° POSTE', 'TIPO / ALTURA', 'TERRENO', 'JUSTIFICACIÓN', 'OBSERVACIÓN / ACCIÓN', 'ESTADO'] + [f'EXTRA_{i}' for i in range(7, df.shape[1])]

    # --- BARRA LATERAL (FILTROS) ---
    st.sidebar.header("🔍 Filtros de Búsqueda")
    
    zona_seleccionada = 'TODOS'
    if 'ZONA' in df.columns:
        zonas_disp = ['TODOS'] + list(df['ZONA'].dropna().astype(str).unique())
        zona_seleccionada = st.sidebar.selectbox("Filtrar por Zona / Sector:", zonas_disp)
    
    estado_seleccionado = 'TODOS'
    if 'ESTADO' in df.columns:
        estados_disp = ['TODOS'] + list(df['ESTADO'].dropna().astype(str).unique())
        estado_seleccionado = st.sidebar.selectbox("Filtrar por Estado:", estados_disp)
        
    terreno_seleccionado = 'TODOS'
    if 'TERRENO' in df.columns:
        terrenos_disp = ['TODOS'] + list(df['TERRENO'].dropna().astype(str).unique())
        terreno_seleccionado = st.sidebar.selectbox("Filtrar por Tipo de Terreno:", terrenos_disp)
    
    # Aplicar filtros
    df_filtrado = df.copy()
    if zona_seleccionada != 'TODOS' and 'ZONA' in df.columns:
        df_filtrado = df_filtrado[df_filtrado['ZONA'].astype(str) == zona_seleccionada]
    if estado_seleccionado != 'TODOS' and 'ESTADO' in df.columns:
        df_filtrado = df_filtrado[df_filtrado['ESTADO'].astype(str) == estado_seleccionado]
    if terreno_seleccionado != 'TODOS' and 'TERRENO' in df.columns:
        df_filtrado = df_filtrado[df_filtrado['TERRENO'].astype(str) == terreno_seleccionado]

    # --- MÉTRICAS PRINCIPALES (KPIs) ---
    total_postes = len(df_filtrado)
    pendientes = len(df_filtrado[df_filtrado['ESTADO'] == 'PENDIENTE']) if 'ESTADO' in df.columns else 0
    atendidos = len(df_filtrado[df_filtrado['ESTADO'] == 'ATENDIDO']) if 'ESTADO' in df.columns else 0
    conformes = len(df_filtrado[df_filtrado['ESTADO'] == 'CONFORME']) if 'ESTADO' in df.columns else 0

    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown(f"""
            <div class="metric-card">
                <div class="metric-title">Observados</div>
                <div class="metric-value">🏛️ {total_postes}</div>
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
                <div class="metric-title">Atendidos</div>
                <div class="metric-value" style="color: #f0ad4e;">🔧 {atendidos}</div>
            </div>
        """, unsafe_allow_html=True)
    with col4:
        st.markdown(f"""
            <div class="metric-card">
                <div class="metric-title">Conformes</div>
                <div class="metric-value" style="color: #5cb85c;">🏆 {conformes}</div>
            </div>
        """, unsafe_allow_html=True)
    
    st.markdown("<br>", unsafe_allow_html=True)

    # --- GRÁFICAS VISUALES (BARRAS APILADAS CON ETIQUETAS NUMÉRICAS) ---
    col_g1, col_g2 = st.columns(2)
    
    color_estados = {
        'PENDIENTE': '#d9534f',
        'ATENDIDO': '#ffc107',
        'CONFORME': '#28a745'
    }

    with col_g1:
        if 'ZONA' in df.columns and 'ESTADO' in df.columns:
            st.subheader("🗺️ Observados por Zonas")
            
            df_zona_estado = df_filtrado.groupby(['ZONA', 'ESTADO']).size().unstack(fill_value=0).reset_index()
            
            fig_bar_zona = go.Figure()
            for estado in ['PENDIENTE', 'ATENDIDO', 'CONFORME']:
                if estado in df_zona_estado.columns:
                    fig_bar_zona.add_trace(go.Bar(
                        name=estado,
                        x=df_zona_estado['ZONA'],
                        y=df_zona_estado[estado],
                        text=df_zona_estado[estado],  # <--- Muestra la cantidad numérica
                        textposition='inside',        # <--- Ubica la etiqueta dentro de la sección
                        textfont=dict(color='white', size=13, family="Arial Black"), # Texto destacado
                        marker_color=color_estados.get(estado, '#333333'),
                        marker_line=dict(color='#111111', width=1.5)
                    ))

            fig_bar_zona.update_layout(
                barmode='stack',
                margin=dict(t=20, b=0, l=0, r=0),
                height=320,
                paper_bgcolor='rgba(0,0,0,0)',
                plot_bgcolor='rgba(0,0,0,0)',
                xaxis=dict(title='Zona', showgrid=False, linecolor='black', linewidth=2),
                yaxis=dict(title='Cantidad', showgrid=True, gridcolor='#dcdcdc', linecolor='black', linewidth=2),
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
            )
            st.plotly_chart(fig_bar_zona, use_container_width=True, config={'displayModeBar': False})
        
    with col_g2:
        if 'ESTADO' in df.columns:
            st.subheader("📌 Estado (% de Avance)")
            conteo_estados = df_filtrado['ESTADO'].value_counts().reset_index()
            conteo_estados.columns = ['ESTADO', 'CANTIDAD']
            
            fig_pie = px.pie(
                conteo_estados, 
                names='ESTADO', 
                values='CANTIDAD', 
                hole=0.35,
                color='ESTADO',
                color_discrete_map=color_estados
            )
            fig_pie.update_traces(
                textposition='inside', 
                textinfo='percent+label',
                pull=[0.05, 0, 0], 
                marker=dict(line=dict(color='#000000', width=1))
            )
            fig_pie.update_layout(
                margin=dict(t=0, b=0, l=0, r=0), 
                height=320,
                paper_bgcolor='rgba(0,0,0,0)',
                plot_bgcolor='rgba(0,0,0,0)',
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
            )
            st.plotly_chart(fig_pie, use_container_width=True, config={'displayModeBar': False})

    # Segunda fila de gráficos: Clasificación por Tipo de Terreno (Apilado con etiquetas)
    if 'TERRENO' in df.columns and 'ESTADO' in df.columns:
        st.markdown("<br>", unsafe_allow_html=True)
        st.subheader("🌍 Clasificación por Tipo de Terreno")
        
        df_terreno_estado = df_filtrado.groupby(['TERRENO', 'ESTADO']).size().unstack(fill_value=0).reset_index()
        
        fig_bar_terreno = go.Figure()
        for estado in ['PENDIENTE', 'ATENDIDO', 'CONFORME']:
            if estado in df_terreno_estado.columns:
                fig_bar_terreno.add_trace(go.Bar(
                    name=estado,
                    x=df_terreno_estado['TERRENO'],
                    y=df_terreno_estado[estado],
                    text=df_terreno_estado[estado],  # <--- Muestra la cantidad numérica
                    textposition='inside',        # <--- Ubica la etiqueta dentro de la sección
                    textfont=dict(color='white', size=13, family="Arial Black"),
                    marker_color=color_estados.get(estado, '#333333'),
                    marker_line=dict(color='#111111', width=1.5)
                ))

        fig_bar_terreno.update_layout(
            barmode='stack',
            margin=dict(t=20, b=0, l=0, r=0),
            height=320,
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
            xaxis=dict(title='Tipo de Terreno', showgrid=False, linecolor='black', linewidth=2),
            yaxis=dict(title='Cantidad', showgrid=True, gridcolor='#dcdcdc', linecolor='black', linewidth=2),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
        )
        st.plotly_chart(fig_bar_terreno, use_container_width=True, config={'displayModeBar': False})

    st.markdown("---")

    # --- BUSCADOR RÁPIDO DE POSTE ESPECÍFICO ---
    st.subheader("🔍 Consulta Individual de Poste")
    col_busqueda = 'N° POSTE' if 'N° POSTE' in df.columns else df.columns[1]
    lista_postes = list(df[col_busqueda].astype(str).unique())
    poste_buscado = st.selectbox("Seleccione o busque el número de poste:", ["-- Seleccionar --"] + lista_postes)
    
    if poste_buscado != "-- Seleccionar --":
        datos_poste = df[df[col_busqueda].astype(str) == str(poste_buscado)]
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
