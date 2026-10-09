"""
Aplicativo de Registro de Verificación PTAR - PROLAN
Permite ingresar datos, generar Excel (mismo diseño del formato oficial)
y generar PDF profesional.
"""

import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
from copy import copy
from io import BytesIO
from datetime import datetime, date
import random

# Intentar importar reportlab para PDF
try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import mm, cm
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

# Configuración de la página
st.set_page_config(
    page_title="Registro PTAR - PROLAN",
    page_icon="💧",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS personalizados
st.markdown("""
<style>
    .main-header {
        font-size: 1.8rem;
        font-weight: bold;
        color: #1a5276;
        text-align: center;
        margin-bottom: 0.5rem;
    }
    .sub-header {
        font-size: 1.1rem;
        color: #2874a6;
        text-align: center;
        margin-bottom: 1.5rem;
    }
    .stDownloadButton button {
        width: 100%;
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">💧 REGISTRO DE VERIFICACIÓN PTAR</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">PROLAN · PI-R.CE-0031</div>', unsafe_allow_html=True)

# ==================== SIDEBAR ====================
with st.sidebar:
    st.header("⚙️ Configuración")
    
    st.subheader("Datos del documento")
    elaborado_por = st.text_input("Elaborado por:", value="")
    vb_responsable = st.text_input("V°B° Responsable de Saneamiento:", value="")
    
    st.markdown("---")
    st.subheader("Opciones de datos")
    
    num_filas = st.slider("Número de filas a mostrar", 5, 20, 20)
    
    if st.button("🧹 Limpiar tabla", use_container_width=True):
        if "df_data" in st.session_state:
            del st.session_state.df_data
        st.rerun()
    
    st.markdown("---")
    st.info("""
    **Instrucciones:**
    1. Complete los datos en la tabla
    2. Use formato de fecha: DD/MM/AA
    3. Horas: 01:00, 05:00, 08:00, 13:00, 17:00, 20:00
    4. Genere Excel o PDF
    """)

# ==================== DATOS DE EJEMPLO / INICIALES ====================
def crear_dataframe_vacio(n=20):
    return pd.DataFrame({
        "N°": list(range(1, n+1)),
        "FECHA": [""] * n,
        "HORA": [""] * n,
        "PH Cámara": [None] * n,
        "Turbidez (NTU)": [None] * n,
        "Conductividad": [None] * n,
        "PH Laguna": [None] * n,
        "Temperatura": [None] * n,
        "PH Agua Clarif.": [None] * n,
        "Turbidez Agua": [None] * n,
        "Conduct. Agua": [None] * n,
    })

if "df_data" not in st.session_state:
    st.session_state.df_data = crear_dataframe_vacio(20)

# Ajustar número de filas
if len(st.session_state.df_data) != num_filas:
    df_old = st.session_state.df_data
    df_new = crear_dataframe_vacio(num_filas)
    # Copiar datos existentes
    for col in df_new.columns:
        if col in df_old.columns:
            df_new[col] = df_old[col].reindex(range(num_filas)).values
    st.session_state.df_data = df_new

# ==================== EDITOR DE DATOS ====================
st.subheader("📋 Ingreso de datos de verificación")

column_config = {
    "N°": st.column_config.NumberColumn("N°", width="small", disabled=True),
    "FECHA": st.column_config.TextColumn("Fecha", width="medium", help="Formato: DD/MM/AA"),
    "HORA": st.column_config.TextColumn("Hora", width="small", help="Ej: 08:00"),
    "PH Cámara": st.column_config.NumberColumn("PH Cámara", min_value=0.0, max_value=14.0, format="%.1f", step=0.1),
    "Turbidez (NTU)": st.column_config.NumberColumn("Turbidez NTU", min_value=0, max_value=500, format="%d"),
    "Conductividad": st.column_config.NumberColumn("Conductividad", min_value=0.0, max_value=20.0, format="%.1f", step=0.1),
    "PH Laguna": st.column_config.NumberColumn("PH Laguna", min_value=0.0, max_value=14.0, format="%.1f", step=0.1),
    "Temperatura": st.column_config.NumberColumn("Temp. °C", min_value=0.0, max_value=50.0, format="%.1f", step=0.1),
    "PH Agua Clarif.": st.column_config.NumberColumn("PH Agua", min_value=0.0, max_value=14.0, format="%.1f", step=0.1),
    "Turbidez Agua": st.column_config.NumberColumn("Turb. Agua", min_value=0, max_value=500, format="%d"),
    "Conduct. Agua": st.column_config.NumberColumn("Cond. Agua", min_value=0.0, max_value=20.0, format="%.1f", step=0.1),
}

edited_df = st.data_editor(
    st.session_state.df_data,
    column_config=column_config,
    num_rows="fixed",
    use_container_width=True,
    hide_index=True,
    key="data_editor"
)

# Guardar cambios en session_state
st.session_state.df_data = edited_df

st.markdown("---")

# ==================== FUNCIONES DE GENERACIÓN ====================
def generar_excel(df, elaborado="", vb=""):
    """Genera el Excel usando el template original y rellena los datos."""
    # Cargar template
    template_path = "Registro_de_verificacion_PETAR.xlsx"
    try:
        wb = load_workbook(template_path)
    except FileNotFoundError:
        # Fallback: crear desde cero si no hay template
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Registro PETAR"
        # Headers básicos
        ws['A1'] = "◉ PROLAN"
        ws['D1'] = "REGISTRO DE VERIFICACIÓN PTAR"
        ws['I1'] = "PI-R.CE-0031"
        ws['I2'] = "Mar. 27, 2026 - V01"
        # ... (simplificado)
        return None
    
    ws = wb.active
    
    center = Alignment(horizontal='center', vertical='center')
    
    # Rellenar datos (filas 6 a 25)
    for i, row in df.iterrows():
        excel_row = 6 + i
        if excel_row > 25:
            break
        
        # Fecha solo si tiene valor
        fecha_val = row["FECHA"] if pd.notna(row["FECHA"]) and str(row["FECHA"]).strip() != "" else None
        hora_val = row["HORA"] if pd.notna(row["HORA"]) and str(row["HORA"]).strip() != "" else None
        
        ws.cell(row=excel_row, column=2).value = fecha_val
        ws.cell(row=excel_row, column=3).value = hora_val
        
        # Valores numéricos formateados
        def fmt(val, decimals=1):
            if pd.isna(val) or val is None or str(val).strip() == "":
                return None
            try:
                return f"{float(val):.{decimals}f}" if decimals > 0 else str(int(float(val)))
            except:
                return str(val)
        
        ws.cell(row=excel_row, column=4).value = fmt(row["PH Cámara"], 1)
        ws.cell(row=excel_row, column=5).value = fmt(row["Turbidez (NTU)"], 0)
        ws.cell(row=excel_row, column=6).value = fmt(row["Conductividad"], 1)
        ws.cell(row=excel_row, column=7).value = fmt(row["PH Laguna"], 1)
        ws.cell(row=excel_row, column=8).value = fmt(row["Temperatura"], 1)
        ws.cell(row=excel_row, column=9).value = fmt(row["PH Agua Clarif."], 1)
        ws.cell(row=excel_row, column=10).value = fmt(row["Turbidez Agua"], 0)
        ws.cell(row=excel_row, column=11).value = fmt(row["Conduct. Agua"], 1)
        
        for col in range(1, 12):
            ws.cell(row=excel_row, column=col).alignment = center
    
    # Firmas
    if elaborado:
        ws.cell(row=29, column=1).value = f"Elaborado por: {elaborado}"
    if vb:
        ws.cell(row=29, column=8).value = f"V°B° del Responsable de Saneamiento: {vb}"
    
    # Guardar en memoria
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


def generar_pdf(df, elaborado="", vb=""):
    """Genera un PDF profesional con el mismo contenido."""
    if not REPORTLAB_AVAILABLE:
        return None
    
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        rightMargin=10*mm,
        leftMargin=10*mm,
        topMargin=10*mm,
        bottomMargin=10*mm
    )
    
    elements = []
    styles = getSampleStyleSheet()
    
    # Estilos personalizados
    title_style = ParagraphStyle(
        'TitleCustom',
        parent=styles['Heading1'],
        fontSize=14,
        alignment=TA_CENTER,
        spaceAfter=2,
        textColor=colors.HexColor("#1a5276")
    )
    subtitle_style = ParagraphStyle(
        'SubtitleCustom',
        parent=styles['Normal'],
        fontSize=9,
        alignment=TA_CENTER,
        spaceAfter=8,
        textColor=colors.HexColor("#2874a6")
    )
    normal_center = ParagraphStyle(
        'NormalCenter',
        parent=styles['Normal'],
        fontSize=8,
        alignment=TA_CENTER
    )
    
    # Encabezado
    elements.append(Paragraph("◉ PROLAN", title_style))
    elements.append(Paragraph("<b>REGISTRO DE VERIFICACIÓN PTAR</b>", title_style))
    elements.append(Paragraph("PI-R.CE-0031 &nbsp;&nbsp;|&nbsp;&nbsp; Mar. 27, 2026 - V01", subtitle_style))
    elements.append(Spacer(1, 4*mm))
    
    # Preparar datos de la tabla
    header1 = ["N°", "FECHA", "HORA", "CÁMARA DE BOMBEO", "", "", "LAGUNA DE ECUALIZACIÓN", "", "AGUA CLARIFICADA", "", ""]
    header2 = ["", "", "", "PH", "TURBIDEZ\n(NTU)", "CONDUCTIVIDAD", "PH", "TEMPERATURA", "PH", "TURBIDEZ\n(NTU)", "CONDUCTIVIDAD"]
    
    data = [header1, header2]
    
    for i, row in df.iterrows():
        def val(v, dec=1):
            if pd.isna(v) or v is None or str(v).strip() == "":
                return ""
            try:
                if dec == 0:
                    return str(int(float(v)))
                return f"{float(v):.{dec}f}"
            except:
                return str(v)
        
        data.append([
            str(int(row["N°"])),
            str(row["FECHA"]) if pd.notna(row["FECHA"]) else "",
            str(row["HORA"]) if pd.notna(row["HORA"]) else "",
            val(row["PH Cámara"], 1),
            val(row["Turbidez (NTU)"], 0),
            val(row["Conductividad"], 1),
            val(row["PH Laguna"], 1),
            val(row["Temperatura"], 1),
            val(row["PH Agua Clarif."], 1),
            val(row["Turbidez Agua"], 0),
            val(row["Conduct. Agua"], 1),
        ])
    
    # Crear tabla
    col_widths = [12*mm, 22*mm, 18*mm, 18*mm, 22*mm, 25*mm, 18*mm, 22*mm, 18*mm, 22*mm, 25*mm]
    
    table = Table(data, colWidths=col_widths, repeatRows=2)
    
    style = TableStyle([
        # Header principal
        ('SPAN', (3, 0), (5, 0)),  # CÁMARA DE BOMBEO
        ('SPAN', (6, 0), (7, 0)),  # LAGUNA
        ('SPAN', (8, 0), (10, 0)), # AGUA CLARIFICADA
        ('BACKGROUND', (0, 0), (-1, 1), colors.HexColor("#1a5276")),
        ('TEXTCOLOR', (0, 0), (-1, 1), colors.white),
        ('FONTNAME', (0, 0), (-1, 1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 1), 7),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#2874a6")),
        ('BACKGROUND', (0, 2), (-1, -1), colors.HexColor("#f8f9f9")),
        ('FONTSIZE', (0, 2), (-1, -1), 8),
        ('ROWBACKGROUNDS', (0, 2), (-1, -1), [colors.white, colors.HexColor("#eaf2f8")]),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
    ])
    table.setStyle(style)
    elements.append(table)
    
    elements.append(Spacer(1, 10*mm))
    
    # Firmas
    firmas_data = [
        [f"Elaborado por: {elaborado or '________________________'}", 
         f"V°B° del Responsable de Saneamiento: {vb or '________________________'}"]
    ]
    firmas_table = Table(firmas_data, colWidths=[120*mm, 120*mm])
    firmas_table.setStyle(TableStyle([
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('ALIGN', (0, 0), (0, 0), 'LEFT'),
        ('ALIGN', (1, 0), (1, 0), 'RIGHT'),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
    ]))
    elements.append(firmas_table)
    
    # Pie
    elements.append(Spacer(1, 5*mm))
    elements.append(Paragraph(
        f"Generado el {datetime.now().strftime('%d/%m/%Y %H:%M')}",
        ParagraphStyle('Footer', parent=styles['Normal'], fontSize=7, alignment=TA_CENTER, textColor=colors.gray)
    ))
    
    doc.build(elements)
    buffer.seek(0)
    return buffer


# ==================== BOTONES DE GENERACIÓN ====================
col1, col2, col3 = st.columns([1, 1, 1])

with col1:
    if st.button("📊 Generar Excel", type="primary", use_container_width=True):
        with st.spinner("Generando archivo Excel..."):
            excel_buffer = generar_excel(edited_df, elaborado_por, vb_responsable)
            if excel_buffer:
                st.session_state.excel_file = excel_buffer
                st.session_state.excel_name = f"Registro_PTAR_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
                st.success("✅ Excel generado correctamente")
            else:
                st.error("Error al generar Excel. Verifique que el archivo template exista.")

with col2:
    if st.button("📄 Generar PDF", type="primary", use_container_width=True):
        if not REPORTLAB_AVAILABLE:
            st.error("reportlab no está instalado. Ejecute: pip install reportlab")
        else:
            with st.spinner("Generando archivo PDF..."):
                pdf_buffer = generar_pdf(edited_df, elaborado_por, vb_responsable)
                if pdf_buffer:
                    st.session_state.pdf_file = pdf_buffer
                    st.session_state.pdf_name = f"Registro_PTAR_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
                    st.success("✅ PDF generado correctamente")
                else:
                    st.error("Error al generar PDF")

with col3:
    if st.button("📥 Cargar datos de ejemplo", use_container_width=True):
        # Datos de ejemplo realistas
        ejemplo = crear_dataframe_vacio(num_filas)
        fechas = ["09/10/26", "", "", "", "", "", "10/10/26", "", "", "", "", "", "11/10/26", "", "", "", "", "", "", ""]
        horas = ["08:00", "13:00", "17:00", "20:00", "01:00", "05:00", "08:00", "13:00", "17:00", "20:00",
                 "01:00", "05:00", "08:00", "13:00", "17:00", "20:00", "01:00", "05:00", "08:00", "13:00"]
        random.seed(42)
        for i in range(num_filas):
            ejemplo.at[i, "FECHA"] = fechas[i] if i < len(fechas) else ""
            ejemplo.at[i, "HORA"] = horas[i] if i < len(horas) else ""
            ejemplo.at[i, "PH Cámara"] = round(random.uniform(3.2, 4.4), 1)
            ejemplo.at[i, "Turbidez (NTU)"] = random.randint(72, 110)
            ejemplo.at[i, "Conductividad"] = round(random.uniform(2.6, 3.6), 1)
            ejemplo.at[i, "PH Laguna"] = round(random.uniform(5.0, 6.4), 1)
        st.session_state.df_data = ejemplo
        st.rerun()

# ==================== DESCARGAS ====================
st.markdown("---")
st.subheader("📥 Descargar archivos generados")

dl_col1, dl_col2 = st.columns(2)

with dl_col1:
    if "excel_file" in st.session_state:
        st.download_button(
            label="⬇️ Descargar Excel",
            data=st.session_state.excel_file,
            file_name=st.session_state.excel_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    else:
        st.info("Genere primero el archivo Excel")

with dl_col2:
    if "pdf_file" in st.session_state:
        st.download_button(
            label="⬇️ Descargar PDF",
            data=st.session_state.pdf_file,
            file_name=st.session_state.pdf_name,
            mime="application/pdf",
            use_container_width=True
        )
    else:
        st.info("Genere primero el archivo PDF")

# Footer
st.markdown("---")
st.caption("Aplicativo desarrollado para PROLAN · Registro de Verificación PTAR · Formato PI-R.CE-0031")
