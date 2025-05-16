import streamlit as st
import pandas as pd
import sqlite3
import os

def cargar_todas_las_hojas(file, nombre_archivo):
    xl = pd.ExcelFile(file)
    hojas = xl.sheet_names
    frames = []

    for hoja in hojas:
        try:
            df = xl.parse(hoja)
            df.columns = df.columns.map(str)  # Asegura nombres de columnas como strings
            df['Fuente'] = f"{nombre_archivo} - {hoja}"
            frames.append(df)
        except Exception as e:
            st.warning(f"⚠️ No se pudo procesar la hoja '{hoja}': {e}")
    return frames

def unir_y_cargar_en_sqlite(lista_dfs):
    df_unificado = pd.concat(lista_dfs, ignore_index=True, sort=False).fillna("")
    conn = sqlite3.connect(":memory:")  # También puedes usar "data.db" si quieres persistencia
    df_unificado.to_sql("datos_unificados", conn, index=False, if_exists="replace")
    return conn

# Streamlit app
st.title("Unión de Archivos Excel para Consultas SQL")

archivo_v1 = st.file_uploader("📂 Cargar archivo V1", type=["xlsx"])
archivo_v2 = st.file_uploader("📂 Cargar archivo V2", type=["xlsx"])

if archivo_v1 and archivo_v2:
    with st.spinner("Procesando archivos..."):
        dfs_v1 = cargar_todas_las_hojas(archivo_v1, "V1")
        dfs_v2 = cargar_todas_las_hojas(archivo_v2, "V2")
        conn_sqlite = unir_y_cargar_en_sqlite(dfs_v1 + dfs_v2)
        st.success("✅ Archivos unidos y cargados en base SQLite")

        # Consulta libre
        consulta = st.text_area("🧠 Escribe tu consulta SQL sobre `datos_unificados`:", 
                                "SELECT * FROM datos_unificados LIMIT 100")
        if st.button("🔍 Ejecutar consulta"):
            try:
                resultado = pd.read_sql_query(consulta, conn_sqlite)
                st.dataframe(resultado)
            except Exception as e:
                st.error(f"❌ Error en consulta: {e}")

