# -*- coding: utf-8 -*-
"""
Created on Tue Sep 15 12:24:15 2026

@author: Henry Z
"""

import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
import io
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# Configuración inicial de la página web del dashboard (modo ancho)
st.set_page_config(
    page_title="Control de Izaje de Postes",
    page_icon="🏗️",
    layout="wide"
)

# Estilos CSS personalizados
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
        .block-container {
            padding-top: 2rem;
            padding-bottom: 2rem;
        }
        img {
            max-width: 150px !important;
            height: auto !important;
            display: block;
            margin-left: auto;
            margin-right: auto;
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

# --- ENCABEZADO SUPERIOR ---
col_logo1, col_title, col_logo2 = st.columns([1, 2.5, 1])

with col_logo1:
    try:
        st.image("quantum.png", width=160)
    except:
        st.write("Logo Quantum no encontrado")

with col_title:
    st.markdown("""
        <div style='text-align: center;'>
            <h2 style='color: #212529; margin-bottom: 0px;'>📊 DASHBOARD CONTROL DE POSTES OBSERVADOS</h2>
            <p style='color: #6c757d; margin-top: 5px; font-size: 16px;'>Monitoreo en tiempo real de avance y levantamiento de observaciones en obra.</p>
        </div>
    """, unsafe_allow_html=True)

with col_logo2:
    try:
        st.image("logo.png", width=130)
    except:
        st.write("Logo Municipalidad no encontrado")

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
                        text=df_zona_estado[estado],
                        textposition='inside',
                        textfont=dict(color='white', size=13, family="Arial Black"),
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
                    text=df_terreno_estado[estado],
                    textposition='inside',
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

    # Función para resaltar filas según su estado
    def resaltar_filas(row):
        if 'ESTADO' in row:
            estado = str(row['ESTADO']).upper()
            if estado == 'CONFORME':
                return ['background-color: #d9ead3; color: #274e13; font-weight: bold'] * len(row)
            elif estado == 'ATENDIDO':
                return ['background-color: #fff2cc; color: #7f6000; font-weight: bold'] * len(row)
        return [''] * len(row)

    # --- BUSCADOR RÁPIDO DE POSTE ESPECÍFICO ---
    st.subheader("🔍 Consulta Individual de Poste")
    col_busqueda = 'N° POSTE' if 'N° POSTE' in df.columns else df.columns[1]
    lista_postes = list(df[col_busqueda].astype(str).unique())
    poste_buscado = st.selectbox("Seleccione o busque el número de poste:", ["-- Seleccionar --"] + lista_postes)
    
    if poste_buscado != "-- Seleccionar --":
        datos_poste = df[df[col_busqueda].astype(str) == str(poste_buscado)]
        datos_poste_estilizado = datos_poste.style.apply(resaltar_filas, axis=1)
        st.dataframe(datos_poste_estilizado, use_container_width=True)

    st.markdown("---")

    # --- TABLA DE DATOS INTERACTIVA COMPLETA ---
    st.subheader(f"📋 Detalle de Registros Filtrados ({len(df_filtrado)} elementos)")
    
    df_filtrado_estilizado = df_filtrado.style.apply(resaltar_filas, axis=1)
    st.dataframe(df_filtrado_estilizado, use_container_width=True)

    # --- SECCIÓN DE EXPORTACIÓN DE REPORTES ---
    st.markdown("### 📥 Exportar Reportes de Obra")
    col_exp1, col_exp2 = st.columns(2)

    # 1. Botón para Exportar a EXCEL
    with col_exp1:
        output_excel = io.BytesIO()
        with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
            df_filtrado.to_excel(writer, index=False, sheet_name='Reporte Postes')
        excel_data = output_excel.getvalue()

        st.download_button(
            label="📊 Descargar Reporte en Excel (.xlsx)",
            data=excel_data,
            file_name="reporte_control_postes.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    # 2. Botón para Exportar a PDF
    with col_exp2:
        def generar_pdf(data_df, total, pend, aten, conf):
            buffer = io.BytesIO()
            doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
            elements = []
            
            styles = getSampleStyleSheet()
            titulo_estilo = ParagraphStyle('Titulo', parent=styles['Heading1'], fontSize=16, alignment=1, textColor=colors.HexColor('#212529'))
            sub_estilo = ParagraphStyle('Sub', parent=styles['Normal'], fontSize=10, alignment=1, textColor=colors.HexColor('#6c757d'))
            
            elements.append(Paragraph("<b>REPORTE DE CONTROL DE POSTES OBSERVADOS</b>", titulo_estilo))
            elements.append(Paragraph("Consorcio Quantum SC Tumbes", sub_estilo))
            elements.append(Spacer(1, 15))
            
            # Resumen de KPIs
            kpi_text = f"<b>Total Registros:</b> {total} &nbsp;&nbsp;|&nbsp;&nbsp; <font color='#d9534f'><b>Pendientes:</b> {pend}</font> &nbsp;&nbsp;|&nbsp;&nbsp; <font color='#7f6000'><b>Atendidos:</b> {aten}</font> &nbsp;&nbsp;|&nbsp;&nbsp; <font color='#274e13'><b>Conformes:</b> {conf}</font>"
            elements.append(Paragraph(kpi_text, ParagraphStyle('KPI', parent=styles['Normal'], fontSize=10, alignment=1)))
            elements.append(Spacer(1, 15))
            
            # Tabla de datos
            table_data = [list(data_df.columns)]
            for _, row in data_df.iterrows():
                table_data.append([str(val) for val in row])
                
            t = Table(table_data, repeatRows=1)
            t.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#343a40')),
                ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                ('FONTSIZE', (0,0), (-1,0), 8),
                ('BOTTOMPADDING', (0,0), (-1,0), 6),
                ('BACKGROUND', (0,1), (-1,-1), colors.HexColor('#f8f9fa')),
                ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#dee2e6')),
                ('FONTNAME', (0,1), (-1,-1), 'Helvetica'),
                ('FONTSIZE', (0,1), (-1,-1), 7),
            ]))
            elements.append(t)
            doc.build(elements)
            buffer.seek(0)
            return buffer.getvalue()

        try:
            pdf_data = generar_pdf(df_filtrado, total_postes, pendientes, atendidos, conformes)
            st.download_button(
                label="📄 Descargar Reporte Ejecutivo en PDF",
                data=pdf_data,
                file_name="reporte_control_postes.pdf",
                mime="application/pdf",
                use_container_width=True
            )
        except Exception as e:
            st.info("Para habilitar la descarga en PDF, asegúrate de incluir 'reportlab' en tu archivo requirements.txt de GitHub.")

    # Botón de refresco manual
    st.markdown("<br>", unsafe_allow_html=True)
    col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
    with col_btn2:
        if st.button("🔄 Refrescar Datos en Vivo desde Google Sheets", use_container_width=True):
            st.cache_data.clear()
            st.rerun()
