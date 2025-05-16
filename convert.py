import streamlit as st
import pandas as pd
import os
from io import BytesIO

st.title("Conversión Masiva de CSV a XLSX desde Carpeta con Verificación")

# Selección de carpeta de origen
folder_path = st.text_input("Ingrese la ruta de la carpeta que contiene los CSV:")

if folder_path:
    try:
        csv_files = [f for f in os.listdir(folder_path) if f.endswith(".csv")]
        if not csv_files:
            st.warning("No se encontraron archivos CSV en la carpeta.")
        else:
            # Crear carpeta de destino si no existe
            output_folder = os.path.join(folder_path, "archivos_convertidos")
            os.makedirs(output_folder, exist_ok=True)

            for file in csv_files:
                file_path = os.path.join(folder_path, file)
                file_name = file.rsplit(".", 1)[0]

                df_original = pd.read_csv(file_path)  # Leer CSV

                # Función para convertir a Excel
                def convert_to_excel(df):
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                        df.to_excel(writer, index=False, sheet_name="Datos")
                    return output.getvalue()

                excel_data = convert_to_excel(df_original)

                # Guardar archivo en la carpeta de destino
                output_file_path = os.path.join(output_folder, f"{file_name}.xlsx")
                with open(output_file_path, "wb") as f:
                    f.write(excel_data)

                # Función de verificación
                def validate_conversion(csv_df, excel_bytes):
                    converted_df = pd.read_excel(BytesIO(excel_bytes), sheet_name="Datos")
                    same_shape = csv_df.shape == converted_df.shape
                    same_data = csv_df.equals(converted_df)
                    return same_shape, same_data

                # Comparar datos
                shape_match, data_match = validate_conversion(df_original, excel_data)

                # Mostrar mensaje según el resultado
                if shape_match and data_match:
                    st.success(f"✅ **{file_name}.xlsx** generado correctamente sin pérdida de datos.")
                else:
                    st.warning(f"⚠️ Diferencias detectadas en **{file_name}.xlsx**, revisa los datos.")

            st.success(f"✅ Todos los archivos han sido procesados. XLSX almacenados en: **{output_folder}**")

    except Exception as e:
        st.error(f"❌ Error al procesar archivos: {e}")

