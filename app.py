import streamlit as st
import pandas as pd
from barcode import Code128
from barcode.writer import ImageWriter
from docx import Document
from docx.shared import Inches
from datetime import datetime
import io
import os

# Configuración de la página
st.set_page_config(page_title="UNT - UGRE", layout="centered")

# Título principal
st.title("📦 UGRE – GENERAR CÓDIGO DE BARRAS PARA LIBROS")

# Configuración personalizada del código de barras
BARCODE_CONFIG = {
    "module_height": 8.0,      # Altura reducida (por defecto 15.0)
    "font_size": 4,            # Tamaño de letra reducido (por defecto 12)
    "text_distance": 2.0,      # Distancia del texto al código
    "quiet_zone": 2.5,         # Margen silencioso
}

def generar_codigo_barras(valor, idx):
    """Genera imagen de código de barras con configuración personalizada"""
    writer = ImageWriter()
    
    # Aplicar configuración personalizada
    writer.set_options(BARCODE_CONFIG)
    
    code = Code128(valor, writer=writer)
    filename = f"temp_barcode_{idx}"
    fullpath = code.save(filename)
    return fullpath

# Subir archivo Excel
uploaded_file = st.file_uploader("📂 Subir archivo Excel", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)

    # Verificar que existan las columnas necesarias
    if "CODIGO BARRAS" not in df.columns or "SIGNATURA" not in df.columns:
        st.error("⚠️ El archivo debe contener las columnas 'CODIGO BARRAS' y 'SIGNATURA'")
        st.stop()

    st.success("✅ Archivo cargado correctamente")
    st.dataframe(df.head())

    # Botón para generar códigos de barras
    if st.button("🔲 GENERAR CÓDIGO DE BARRAS"):
        with st.spinner("Generando códigos de barras..."):
            # Crear documento Word
            doc = Document()
            doc.add_heading("Códigos de barras generados", level=1)

            # Crear tabla de 2 columnas
            table = doc.add_table(rows=1, cols=2)
            table.style = 'Table Grid'
            header_cells = table.rows[0].cells
            header_cells[0].text = "SIGNATURA"
            header_cells[1].text = "CÓDIGO DE BARRAS"

            temp_image_paths = []

            for idx, row in df.iterrows():
                signatura = str(row["SIGNATURA"])
                codigo_barras = str(row["CODIGO BARRAS"])

                # Generar imagen del código de barras
                try:
                    fullpath = generar_codigo_barras(codigo_barras, idx)
                    temp_image_paths.append(fullpath)

                    # Agregar fila a la tabla Word
                    row_cells = table.add_row().cells
                    row_cells[0].text = signatura

                    # Insertar imagen del código de barras
                    paragraph = row_cells[1].paragraphs[0]
                    run = paragraph.add_run()
                    run.add_picture(fullpath, width=Inches(1.8))  # Ancho ligeramente reducido

                except Exception as e:
                    st.warning(f"Error generando código para '{codigo_barras}': {e}")
                    row_cells = table.add_row().cells
                    row_cells[0].text = signatura
                    row_cells[1].text = "Error al generar"

            # Guardar documento Word en memoria
            doc_filename = f"COD_BAR_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
            doc_io = io.BytesIO()
            doc.save(doc_io)
            doc_io.seek(0)

            # Limpiar archivos temporales
            for path in temp_image_paths:
                try:
                    os.remove(path)
                except:
                    pass

            # Guardar en session_state para descarga
            st.session_state["doc_bytes"] = doc_io
            st.session_state["doc_filename"] = doc_filename

            st.success("✅ Códigos de barras generados correctamente")

    # Botón de descarga
    if "doc_bytes" in st.session_state:
        st.download_button(
            label="📥 DESCARGAR ARCHIVO WORD",
            data=st.session_state["doc_bytes"],
            file_name=st.session_state["doc_filename"],
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )