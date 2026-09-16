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
import requests
import matplotlib.pyplot as plt
import matplotlib
matplotlib.use('Agg')  # Configuración segura para servidor en la nube

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# Configuración inicial de la página web del dashboard (modo ancho)
st.set_page_config(
    page_title="Control de Izaje de Postes",
    page_icon="🏗️",
    layout="wide"
)

# Estilos CSS generales y clases para centrar títulos y leyendas de gráficos
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
        [data-testid="stImage"] img {
            padding: 4px 0px;
        }
        .centered-subheader {
            text-align: center;
            font-weight: bold;
            font-size: 20px;
            color: #212529;
            margin-bottom: 5px;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
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

# --- ENCABEZADO SUPERIOR CON QUANTUM.PNG Y TÍTULO CENTRADO ---
col_logo, col_title = st.columns([1, 4])

with col_logo:
    try:
        st.image("quantum.png", width=300)
    except:
        st.write("Logo Quantum no encontrado")

with col_title:
    st.markdown("""
        <div style='display: flex; flex-direction: column; justify-content: center; height: 100%; text-align: center; padding-top: 10px;'>
            <h2 style='color: #212529; margin-bottom: 0px; font-size: calc(1.3rem + 1vw);'>📊 DASHBOARD CONTROL DE POSTES OBSERVADOS</h2>
            <p style='color: #6c757d; margin-top: 5px; font-size: calc(0.9rem + 0.3vw);'>Monitoreo en tiempo real de avance y levantamiento de observaciones en obra.</p>
        </div>
    """, unsafe_allow_html=True)

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
        n_cols_base = 7
        n_extra = df.shape[1] - n_cols_base
        base_names = ['ZONA', 'N° POSTE', 'TIPO / ALTURA', 'TERRENO', 'JUSTIFICACIÓN', 'OBSERVACIÓN / ACCIÓN', 'ESTADO']
        if n_extra >= 2:
            base_names += ['FOTO ANTES', 'FOTO DESPUES'] + [f'EXTRA_{i}' for i in range(9, df.shape[1])]
        else:
            base_names += [f'EXTRA_{i}' for i in range(7, df.shape[1])]
        df.columns = base_names

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
                <div class="metric-value">🔍 {total_postes}</div>
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

    # --- GRÁFICAS VISUALES ---
    col_g1, col_g2 = st.columns(2)
    
    color_estados = {
        'PENDIENTE': '#d9534f',
        'ATENDIDO': '#ffc107',
        'CONFORME': '#28a745'
    }

    with col_g1:
        if 'ZONA' in df.columns and 'ESTADO' in df.columns:
            st.markdown('<div class="centered-subheader"><span>🗺️</span> <span>Observados por Zonas</span></div>', unsafe_allow_html=True)
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
                barmode='stack', margin=dict(t=20, b=0, l=0, r=0), height=320,
                paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                xaxis=dict(title='Zona', showgrid=False, linecolor='black', linewidth=2),
                yaxis=dict(title='Cantidad', showgrid=True, gridcolor='#dcdcdc', linecolor='black', linewidth=2),
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5)
            )
            st.plotly_chart(fig_bar_zona, use_container_width=True, config={'displayModeBar': False})
        
    with col_g2:
        if 'ESTADO' in df.columns:
            st.markdown('<div class="centered-subheader"><span>📌</span> <span>Estado (% Total de Avance)</span></div>', unsafe_allow_html=True)
            conteo_estados = df_filtrado['ESTADO'].value_counts().reset_index()
            conteo_estados.columns = ['ESTADO', 'CANTIDAD']
            
            fig_pie = px.pie(conteo_estados, names='ESTADO', values='CANTIDAD', hole=0.35, color='ESTADO', color_discrete_map=color_estados)
            fig_pie.update_traces(textposition='inside', textinfo='percent+label', pull=[0.05, 0, 0], marker=dict(line=dict(color='#000000', width=1)))
            fig_pie.update_layout(margin=dict(t=0, b=0, l=0, r=0), height=320, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5))
            st.plotly_chart(fig_pie, use_container_width=True, config={'displayModeBar': False})

    if 'TERRENO' in df.columns and 'ESTADO' in df.columns:
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown('<div class="centered-subheader"><span>🌍</span> <span>Clasificación por Tipo de Terreno</span></div>', unsafe_allow_html=True)
        df_terreno_estado = df_filtrado.groupby(['TERRENO', 'ESTADO']).size().unstack(fill_value=0).reset_index()
        
        fig_bar_terreno = go.Figure()
        for estado in ['PENDIENTE', 'ATENDIDO', 'CONFORME']:
            if estado in df_terreno_estado.columns:
                fig_bar_terreno.add_trace(go.Bar(
                    name=estado, x=df_terreno_estado['TERRENO'], y=df_terreno_estado[estado],
                    text=df_terreno_estado[estado], textposition='inside',
                    textfont=dict(color='white', size=13, family="Arial Black"),
                    marker_color=color_estados.get(estado, '#333333'), marker_line=dict(color='#111111', width=1.5)
                ))

        fig_bar_terreno.update_layout(
            barmode='stack', margin=dict(t=20, b=0, l=0, r=0), height=320,
            paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
            xaxis=dict(title='Tipo de Terreno', showgrid=False, linecolor='black', linewidth=2),
            yaxis=dict(title='Cantidad', showgrid=True, gridcolor='#dcdcdc', linecolor='black', linewidth=2),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5)
        )
        st.plotly_chart(fig_bar_terreno, use_container_width=True, config={'displayModeBar': False})

    st.markdown("---")

    def resaltar_filas(row):
        if 'ESTADO' in row:
            estado = str(row['ESTADO']).upper()
            if estado == 'CONFORME':
                return ['background-color: #d9ead3; color: #274e13; font-weight: bold'] * len(row)
            elif estado == 'ATENDIDO':
                return ['background-color: #fff2cc; color: #7f6000; font-weight: bold'] * len(row)
        return [''] * len(row)

    # --- BUSCADOR RÁPIDO DE POSTE Y VISOR GOOGLE DRIVE ---
    st.subheader("🔍 Consulta Individual de Poste y Fotografías de Obra")
    col_busqueda = 'N° POSTE' if 'N° POSTE' in df.columns else df.columns[1]
    lista_postes = list(df[col_busqueda].astype(str).unique())
    poste_buscado = st.selectbox("Seleccione o busque el número de poste para inspeccionar:", ["-- Seleccionar --"] + lista_postes)
    
    if poste_buscado != "-- Seleccionar --":
        datos_poste = df[df[col_busqueda].astype(str) == str(poste_buscado)]
        datos_poste_estilizado = datos_poste.style.apply(resaltar_filas, axis=1)
        st.dataframe(datos_poste_estilizado, use_container_width=True)
        
        # Visor fotográfico inteligente con soporte directo para Google Drive
        if 'FOTO ANTES' in df.columns and 'FOTO DESPUES' in df.columns:
            st.markdown("### 📸 Registro Fotográfico (Antes / Después)")
            
            val_antes = str(datos_poste['FOTO ANTES'].values[0]).strip()
            val_despues = str(datos_poste['FOTO DESPUES'].values[0]).strip()
            
            DRIVE_FOLDER_ID = "1rzca9ChAlo5_hsbKV4ZuPQq8_YDmaqcx"
            
            c_foto1, c_foto2 = st.columns(2)
            
            with c_foto1:
                st.markdown("**📸 Estado: ANTES**")
                if val_antes and val_antes.lower() != 'nan':
                    st.markdown(f"Código: **{val_antes}**")
                    # ID específico temporal de prueba para 116_A o mapeo automático
                    id_imagen = "1RfCUkxAShfPcxrTHWHBh7Byq-QNZHvJD" if val_antes == "116_A" else None
                    
                    if id_imagen:
                        try:
                            st.image(f"https://drive.google.com/uc?export=view&id={id_imagen}", caption=f"Poste {poste_buscado} - Antes ({val_antes})", use_column_width=True)
                        except:
                            st.info(f"Verificando enlace de Drive para {val_antes}...")
                    else:
                        # Enlace directo de respaldo a la carpeta compartida de Google Drive
                        url_drive = f"https://drive.google.com/drive/folders/{DRIVE_FOLDER_ID}"
                        st.info(f"Carpeta compartida para revisión: `{val_antes}.jpeg`")
                        st.markdown(f'<a href="{url_drive}" target="_blank" style="display: inline-block; background-color: #0d6efd; color: white; padding: 6px 14px; border-radius: 4px; text-decoration: none; font-weight: bold; font-size: 13px;">📂 Abrir Carpeta Google Drive</a>', unsafe_allow_html=True)
                else:
                    st.info("No hay código registrado para 'FOTO ANTES'.")
            
            with c_foto2:
                st.markdown("**📸 Estado: DESPUÉS**")
                if val_despues and val_despues.lower() != 'nan':
                    st.markdown(f"Código: **{val_despues}**")
                    url_drive = f"https://drive.google.com/drive/folders/{DRIVE_FOLDER_ID}"
                    st.info(f"Carpeta compartida para revisión: `{val_despues}.jpeg`")
                    st.markdown(f'<a href="{url_drive}" target="_blank" style="display: inline-block; background-color: #0d6efd; color: white; padding: 6px 14px; border-radius: 4px; text-decoration: none; font-weight: bold; font-size: 13px;">📂 Abrir Carpeta Google Drive</a>', unsafe_allow_html=True)
                else:
                    st.info("No hay código registrado para 'FOTO DESPUES'.")

    st.markdown("---")

    # --- TABLA DE DATOS INTERACTIVA COMPLETA ---
    st.subheader(f"📋 Detalle de Registros Filtrados ({len(df_filtrado)} elementos)")
    df_filtrado_estilizado = df_filtrado.style.apply(resaltar_filas, axis=1)
    st.dataframe(df_filtrado_estilizado, use_container_width=True)

    # --- SECCIÓN DE EXPORTACIÓN DE REPORTES ---
    st.markdown("### 📥 Exportar Reportes de Obra")
    col_exp1, col_exp2 = st.columns(2)

    with col_exp1:
        def generar_excel_estilizado(data_df):
            wb = Workbook()
            ws = wb.active
            ws.title = "Control de Postes"
            
            headers = list(data_df.columns)
            ws.append(headers)
            
            header_fill = PatternFill(start_color="343A40", end_color="343A40", fill_type="solid")
            header_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
            align_center = Alignment(horizontal="center", vertical="center")
            
            for col_num in range(1, len(headers) + 1):
                cell = ws.cell(row=1, column=col_num)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = align_center

            fill_conforme = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")
            font_conforme = Font(name="Arial", size=9, color="274E13", bold=True)
            
            fill_atendido = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
            font_atendico = Font(name="Arial", size=9, color="7F6000", bold=True)
            
            default_font = Font(name="Arial", size=9)
            thin_border = Border(
                left=Side(style='thin', color='DEE2E6'), right=Side(style='thin', color='DEE2E6'),
                top=Side(style='thin', color='DEE2E6'), bottom=Side(style='thin', color='DEE2E6')
            )

            estado_idx = headers.index('ESTADO') + 1 if 'ESTADO' in headers else None

            for row_idx, row in enumerate(data_df.values, start=2):
                ws.append(list(row))
                estado_val = str(row[estado_idx - 1]).upper() if estado_idx else ""
                
                for col_num in range(1, len(headers) + 1):
                    cell = ws.cell(row=row_idx, column=col_num)
                    cell.border = thin_border
                    cell.font = default_font
                    
                    if estado_val == 'CONFORME':
                        cell.fill = fill_conforme
                        cell.font = font_conforme
                    elif estado_val == 'ATENDIDO':
                        cell.fill = fill_atendido
                        cell.font = font_atendico

            for col in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = col[0].column_letter
                ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

            output = io.BytesIO()
            wb.save(output)
            output.seek(0)
            return output.getvalue()

        try:
            excel_data = generar_excel_estilizado(df_filtrado)
            st.download_button(
                label="📊 Descargar Reporte en Excel con Formato (.xlsx)",
                data=excel_data, file_name="reporte_control_postes.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        except Exception as e:
            st.error(f"Error al generar Excel: {e}")

    with col_exp2:
        def generar_pdf_con_matplot(data_df, total, pend, aten, conf):
            buffer = io.BytesIO()
            doc = SimpleDocTemplate(buffer, pagesize=landscape(letter), rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
            elements = []
            
            styles = getSampleStyleSheet()
            titulo_estilo = ParagraphStyle('Titulo', parent=styles['Heading1'], fontSize=15, alignment=1, textColor=colors.HexColor('#212529'))
            sub_estilo = ParagraphStyle('Sub', parent=styles['Normal'], fontSize=10, alignment=1, textColor=colors.HexColor('#6c757d'))
            seccion_estilo = ParagraphStyle('Sec', parent=styles['Heading2'], fontSize=12, textColor=colors.HexColor('#343a40'), spaceBefore=10, spaceAfter=5)
            
            elements.append(Paragraph("<b>REPORTE EJECUTIVO - CONTROL DE POSTES OBSERVADOS</b>", titulo_estilo))
            elements.append(Paragraph("Consorcio Quantum SC Tumbes", sub_estilo))
            elements.append(Spacer(1, 10))
            
            kpi_text = f"<b>Total Registros:</b> {total} &nbsp;&nbsp;|&nbsp;&nbsp; <font color='#d9534f'><b>Pendientes:</b> {pend}</font> &nbsp;&nbsp;|&nbsp;&nbsp; <font color='#7f6000'><b>Atendidos:</b> {aten}</font> &nbsp;&nbsp;|&nbsp;&nbsp; <font color='#274e13'><b>Conformes:</b> {conf}</font>"
            elements.append(Paragraph(kpi_text, ParagraphStyle('KPI', parent=styles['Normal'], fontSize=10, alignment=1)))
            elements.append(Spacer(1, 15))

            img_buf_1, img_buf_2 = None, None
            try:
                if 'ESTADO' in data_df.columns and len(data_df) > 0:
                    fig, ax = plt.subplots(figsize=(4.2, 2.2))
                    conteo = data_df['ESTADO'].value_counts()
                    colores_pie = ['#d9534f' if x=='PENDIENTE' else '#ffc107' if x=='ATENDIDO' else '#28a745' for x in conteo.index]
                    
                    wedges, texts, autotexts = ax.pie(conteo, labels=None, colors=colores_pie, autopct='%1.1f%%', startangle=90, pctdistance=0.55, textprops={'fontsize': 8, 'weight': 'bold', 'color': 'white'})
                    ax.legend(wedges, conteo.index, title="Estado", loc="center left", bbox_to_anchor=(0.95, 0.5), fontsize=8, title_fontsize=8)
                    ax.set_title("Distribución por Estado", fontsize=9, fontweight='bold')
                    plt.tight_layout()
                    
                    img_buf_1 = io.BytesIO()
                    plt.savefig(img_buf_1, format='png', dpi=150, bbox_inches='tight')
                    img_buf_1.seek(0)
                    plt.close(fig)

                if 'ZONA' in data_df.columns and 'ESTADO' in data_df.columns and len(data_df) > 0:
                    fig, ax = plt.subplots(figsize=(4.8, 2.2))
                    df_zona_est = data_df.groupby(['ZONA', 'ESTADO']).size().unstack(fill_value=0)
                    for est in ['PENDIENTE', 'ATENDIDO', 'CONFORME']:
                        if est not in df_zona_est.columns:
                            df_zona_est[est] = 0
                    df_zona_est = df_zona_est[['PENDIENTE', 'ATENDIDO', 'CONFORME']]
                    
                    bottom = None
                    colores_bar = ['#d9534f', '#ffc107', '#28a745']
                    for idx, estado in enumerate(df_zona_est.columns):
                        values = df_zona_est[estado].values
                        ax.bar(df_zona_est.index.astype(str), values, bottom=bottom, label=estado, color=colores_bar[idx], width=0.45)
                        if bottom is None:
                            bottom = values
                        else:
                            bottom = bottom + values

                    ax.legend(title="Estado", loc="upper right", fontsize=7, title_fontsize=7)
                    ax.set_title("Observados por Zonas", fontsize=9, fontweight='bold')
                    ax.tick_params(axis='x', rotation=10, labelsize=8)
                    ax.tick_params(axis='y', labelsize=8)
                    plt.tight_layout()
                    
                    img_buf_2 = io.BytesIO()
                    plt.savefig(img_buf_2, format='png', dpi=150, bbox_inches='tight')
                    img_buf_2.seek(0)
                    plt.close(fig)
            except Exception as e:
                pass

            elements.append(Paragraph("<b>📈 Resumen Estadístico de Avance</b>", seccion_estilo))
            img_elements = []
            if img_buf_1:
                img_elements.append(RLImage(img_buf_1, width=220, height=120))
            if img_buf_2:
                img_elements.append(RLImage(img_buf_2, width=240, height=120))
            
            if img_elements:
                while len(img_elements) < 2:
                    img_elements.append('')
                t_grafs = Table([img_elements], colWidths=[360, 360])
                t_grafs.setStyle(TableStyle([
                    ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                    ('BOTTOMPADDING', (0,0), (-1,-1), 10),
                ]))
                elements.append(t_grafs)

            elements.append(Spacer(1, 15))
            elements.append(Paragraph("<b>📋 Detalle de Registros en Obra</b>", seccion_estilo))

            estilo_celda = ParagraphStyle('Celda', parent=styles['Normal'], fontSize=8, textColor=colors.HexColor('#212529'))
            estilo_cabecera = ParagraphStyle('Cabecera', parent=styles['Normal'], fontSize=8, fontName='Helvetica-Bold', textColor=colors.whitesmoke, alignment=1)
            
            table_data = []
            header_row = [Paragraph(str(col), estilo_cabecera) for col in data_df.columns]
            table_data.append(header_row)
            
            for _, row in data_df.iterrows():
                row_cells = [Paragraph(str(val), estilo_celda) for val in row]
                table_data.append(row_cells)
                
            col_widths = None
                
            t = Table(table_data, colWidths=col_widths, repeatRows=1)
            t.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#343a40')),
                ('ALIGN', (0,0), (-1,-1), 'LEFT'),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('BOTTOMPADDING', (0,0), (-1,-1), 5),
                ('TOPPADDING', (0,0), (-1,-1), 5),
                ('BACKGROUND', (0,1), (-1,-1), colors.HexColor('#f8f9fa')),
                ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#dee2e6')),
            ]))
            elements.append(t)
            doc.build(elements)
            buffer.seek(0)
            return buffer.getvalue()

        try:
            pdf_data = generar_pdf_con_matplot(df_filtrado, total_postes, pendientes, atendidos, conformes)
            st.download_button(
                label="📄 Descargar Reporte Ejecutivo con Gráficos en PDF",
                data=pdf_data, file_name="reporte_ejecutivo_postes.pdf",
                mime="application/pdf", use_container_width=True
            )
        except Exception as e:
            st.info(f"Detalle del PDF: {e}")

    # Botón de refresco manual
    st.markdown("<br>", unsafe_allow_html=True)
    col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
    with col_btn2:
        if st.button("🔄 Refrescar Datos en Vivo desde Google Sheets", use_container_width=True):
            st.cache_data.clear()
            st.rerun()
