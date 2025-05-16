import os
import pandas as pd
import streamlit as st

st.title("Comparador de Estructura de Archivos XLSX con Progreso")

folder_path = st.text_input("Ingrese la ruta de la carpeta con los archivos XLSX:")

if folder_path:
    try:
        st.write("🔄 Buscando archivos XLSX en la carpeta...")
        xlsx_files = [f for f in os.listdir(folder_path) if f.endswith(".xlsx")]

        if not xlsx_files:
            st.warning("⚠️ No se encontraron archivos XLSX en la carpeta.")
        else:
            st.write(f"📂 Se encontraron {len(xlsx_files)} archivos. Iniciando comparación...")

            progress_bar = st.progress(0)
            status_text = st.empty()  # Espacio para mostrar el mensaje dinámico

            def compare_structure(file1, file2):
                df1 = pd.read_excel(os.path.join(folder_path, file1), sheet_name="Datos")
                df2 = pd.read_excel(os.path.join(folder_path, file2), sheet_name="Datos")

                missing_in_df1 = set(df2.columns) - set(df1.columns)
                missing_in_df2 = set(df1.columns) - set(df2.columns)

                diff_report = []
                if missing_in_df1:
                    diff_report.append(f"⚠️ Columnas faltantes en **{file1}**: {missing_in_df1}")
                if missing_in_df2:
                    diff_report.append(f"⚠️ Columnas faltantes en **{file2}**: {missing_in_df2}")

                return diff_report if diff_report else "✅ Estructura idéntica."

            differences_report = []
            total_comparisons = len(xlsx_files) * (len(xlsx_files) - 1) // 2
            comparison_count = 0

            with st.status("Comparando estructura de archivos XLSX...", expanded=True) as status:
                for i in range(len(xlsx_files)):
                    for j in range(i + 1, len(xlsx_files)):
                        file1, file2 = xlsx_files[i], xlsx_files[j]

                        # Actualizar mensaje en pantalla
                        comparison_count += 1
                        progress_bar.progress(comparison_count / total_comparisons)
                        status_text.write(f"🔍 Comparando archivo {comparison_count} de {total_comparisons}: **{file1} vs {file2}**")

                        differences = compare_structure(file1, file2)
                        if isinstance(differences, list):
                            differences_report.append((file1, file2, differences))

                status.update(label="✅ Comparación completa.", state="complete", expanded=False)

            if differences_report:
                st.warning("⚠️ Detalles de diferencias estructurales encontradas:")
                for file1, file2, details in differences_report:
                    st.write(f"🔍 **{file1} vs {file2}**")
                    for detail in details:
                        st.write(f" - {detail}")
            else:
                st.success("✅ Todos los archivos tienen la misma estructura.")

    except Exception as e:
        st.error(f"❌ Error al procesar archivos: {e}")


