import os
import pandas as pd
import streamlit as st

st.title("Unificador de Archivos XLSX (Nuevo Archivo)")

# Ruta de la carpeta con los archivos
folder_path = st.text_input("Ingrese la ruta de la carpeta con los archivos XLSX:")

if folder_path:
    try:
        st.write("🔄 Buscando archivos XLSX en la carpeta...")
        xlsx_files = [f for f in os.listdir(folder_path) if f.endswith(".xlsx")]

        if not xlsx_files:
            st.warning("⚠️ No se encontraron archivos XLSX en la carpeta.")
        else:
            st.write(f"📂 Se encontraron {len(xlsx_files)} archivos. Iniciando unificación...")

            all_data = []  # Lista para almacenar los datos de cada archivo

            for file in xlsx_files:
                file_path = os.path.join(folder_path, file)
                df = pd.read_excel(file_path, sheet_name="Datos")

                # Agregar una columna con el nombre del archivo para referencia
                df["Fuente"] = file
                all_data.append(df)

            # Concatenar todos los DataFrames en uno solo
            merged_df = pd.concat(all_data, ignore_index=True)

            # Generar un nuevo archivo Excel
            output_file = "C:/Users/Roger/Downloads/unificado.xlsx"
            merged_df.to_excel(output_file, index=False)

            st.success(f"✅ Se ha creado un nuevo archivo: **{output_file}**")

    except Exception as e:
        st.error(f"❌ Error al procesar archivos: {e}")
