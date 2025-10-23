import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO

st.title("📊 Consolidador de Archivos Excel (.xlsm)")

st.markdown("""
Esta aplicación consolida datos de varios archivos `.xlsm` en una **plantilla base**, 
manteniendo las fórmulas y formatos existentes en las hojas de la plantilla.
""")

# --- Carga de archivos ---
uploaded_files = st.file_uploader(
    "Sube los archivos Excel (.xlsm) a consolidar",
    type=["xlsm"],
    accept_multiple_files=True
)

template_file = st.file_uploader(
    "Sube el archivo plantilla (.xlsm)",
    type=["xlsm"]
)

if uploaded_files and template_file:
    st.success(f"Se cargaron {len(uploaded_files)} archivos y 1 plantilla.")
    
    if st.button("🔄 Consolidar"):
        # Cargar plantilla
        plantilla = load_workbook(template_file, keep_vba=True)
        
        # Recorremos cada hoja de la plantilla
        for hoja in plantilla.sheetnames:
            ws_plantilla = plantilla[hoja]
            
            # Crear un DataFrame temporal con los datos actuales de la hoja (sin encabezado)
            max_row = ws_plantilla.max_row
            start_row = 2  # evitar sobrescribir cabecera
            current_row = start_row
            
            for uploaded in uploaded_files:
                # Cargar libro fuente
                try:
                    libro = pd.ExcelFile(uploaded)
                    if hoja in libro.sheet_names:
                        df = pd.read_excel(uploaded, sheet_name=hoja, header=0)
                        
                        # Copiar datos al destino
                        for _, row in df.iterrows():
                            current_row += 1
                            for col_idx, value in enumerate(row, start=1):
                                ws_plantilla.cell(row=current_row, column=col_idx, value=value)
                except Exception as e:
                    st.warning(f"Error al procesar {uploaded.name}: {e}")
        
        # Guardar consolidado
        output = BytesIO()
        plantilla.save(output)
        output.seek(0)
        
        st.download_button(
            label="⬇️ Descargar archivo consolidado",
            data=output,
            file_name="consolidado.xlsm",
            mime="application/vnd.ms-excel.sheet.macroEnabled.12"
        )
        
        st.success("Consolidación completada con éxito ✅")

else:
    st.info("Por favor, carga los archivos a consolidar y la plantilla base para continuar.")
