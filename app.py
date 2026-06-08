import streamlit as st
import pandas as pd
from barcode import Code128
from barcode.writer import ImageWriter
from docx import Document
from docx.shared import Inches
from datetime import datetime
import io
import os
from concurrent.futures import ThreadPoolExecutor, as_completed

# ─────────────────────────────────────────────
# Configuración de la página
# ─────────────────────────────────────────────
st.set_page_config(page_title="UNT - UGRE", layout="centered")
st.title("📦 UGRE – GENERAR CÓDIGO DE BARRAS PARA LIBROS")

# ─────────────────────────────────────────────
# Configuración del código de barras
# ─────────────────────────────────────────────
BARCODE_CONFIG = {
    "module_height": 8.0,
    "font_size": 4,
    "text_distance": 2.0,
    "quiet_zone": 2.5,
}

# ─────────────────────────────────────────────
# OPTIMIZACIÓN 1: Generación en memoria (sin I/O a disco)
# ─────────────────────────────────────────────
def generar_codigo_barras_en_memoria(valor: str) -> io.BytesIO | None:
    """
    Genera el código de barras directamente en un buffer en memoria.
    Evita escribir/leer archivos temporales del disco (I/O más lento).
    """
    try:
        writer = ImageWriter()
        writer.set_options(BARCODE_CONFIG)
        code = Code128(valor, writer=writer)
        buffer = io.BytesIO()
        code.write(buffer)
        buffer.seek(0)
        return buffer
    except Exception:
        return None

# ─────────────────────────────────────────────
# OPTIMIZACIÓN 2: Procesamiento paralelo con caché por sesión
# ─────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def generar_todos_los_codigos(codigos: tuple) -> dict:
    """
    Genera todos los códigos de barras en paralelo usando ThreadPoolExecutor.
    `st.cache_data` evita regenerarlos si el Excel no cambió.
    Recibe una tuple (hashable) para que el caché funcione correctamente.
    """
    resultados = {}

    def tarea(valor):
        return valor, generar_codigo_barras_en_memoria(valor)

    with ThreadPoolExecutor(max_workers=os.cpu_count() or 4) as executor:
        futuros = {executor.submit(tarea, v): v for v in codigos}
        for futuro in as_completed(futuros):
            valor, buffer = futuro.result()
            resultados[valor] = buffer

    return resultados

# ─────────────────────────────────────────────
# OPTIMIZACIÓN 3: Construcción del Word en un solo paso (sin re-opens)
# ─────────────────────────────────────────────
def construir_documento_word(df: pd.DataFrame, imagenes: dict) -> io.BytesIO:
    """
    Arma el documento Word usando las imágenes ya generadas en memoria.
    """
    doc = Document()
    doc.add_heading("Códigos de barras generados", level=1)

    table = doc.add_table(rows=1, cols=2)
    table.style = "Table Grid"
    header_cells = table.rows[0].cells
    header_cells[0].text = "SIGNATURA"
    header_cells[1].text = "CÓDIGO DE BARRAS"

    for _, row in df.iterrows():
        signatura   = str(row["SIGNATURA"])
        cod_barras  = str(row["CODIGO BARRAS"])
        buffer      = imagenes.get(cod_barras)

        row_cells = table.add_row().cells
        row_cells[0].text = signatura

        if buffer:
            buffer.seek(0)
            paragraph = row_cells[1].paragraphs[0]
            run = paragraph.add_run()
            run.add_picture(buffer, width=Inches(1.8))
        else:
            row_cells[1].text = "Error al generar"

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

# ─────────────────────────────────────────────
# UI principal
# ─────────────────────────────────────────────
uploaded_file = st.file_uploader("📂 Subir archivo Excel", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)

    if "CODIGO BARRAS" not in df.columns or "SIGNATURA" not in df.columns:
        st.error("⚠️ El archivo debe contener las columnas 'CODIGO BARRAS' y 'SIGNATURA'")
        st.stop()

    st.success(f"✅ Archivo cargado — {len(df)} registros")
    st.dataframe(df.head())

    if st.button("🔲 GENERAR CÓDIGO DE BARRAS"):
        codigos_unicos = tuple(df["CODIGO BARRAS"].astype(str).unique())
        total = len(df)

        # ── Fase 1: generar imágenes en paralelo ──────────────────────────
        with st.spinner(f"⚙️ Generando {len(codigos_unicos)} código(s) en paralelo…"):
            imagenes = generar_todos_los_codigos(codigos_unicos)

        # ── Fase 2: construir Word ────────────────────────────────────────
        progreso = st.progress(0, text="📄 Armando documento Word…")

        # Procesamos el df en chunks para actualizar la barra de progreso
        doc = Document()
        doc.add_heading("Códigos de barras generados", level=1)
        table = doc.add_table(rows=1, cols=2)
        table.style = "Table Grid"
        hc = table.rows[0].cells
        hc[0].text = "SIGNATURA"
        hc[1].text = "CÓDIGO DE BARRAS"

        for i, (_, row) in enumerate(df.iterrows()):
            signatura  = str(row["SIGNATURA"])
            cod_barras = str(row["CODIGO BARRAS"])
            buffer     = imagenes.get(cod_barras)

            row_cells = table.add_row().cells
            row_cells[0].text = signatura

            if buffer:
                buffer.seek(0)
                paragraph = row_cells[1].paragraphs[0]
                run = paragraph.add_run()
                run.add_picture(buffer, width=Inches(1.8))
            else:
                row_cells[1].text = "Error al generar"

            progreso.progress((i + 1) / total, text=f"📄 Fila {i+1} de {total}…")

        progreso.empty()

        # ── Guardar en memoria ───────────────────────────────────────────
        doc_filename = f"COD_BAR_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
        doc_io = io.BytesIO()
        doc.save(doc_io)
        doc_io.seek(0)

        st.session_state["doc_bytes"]    = doc_io
        st.session_state["doc_filename"] = doc_filename
        st.success("✅ ¡Códigos generados correctamente!")

    # ── Descarga ─────────────────────────────────────────────────────────
    if "doc_bytes" in st.session_state:
        st.download_button(
            label="📥 DESCARGAR ARCHIVO WORD",
            data=st.session_state["doc_bytes"],
            file_name=st.session_state["doc_filename"],
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )