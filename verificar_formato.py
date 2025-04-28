import pandas as pd

# Cambia a la ruta correcta del archivo .xlsx
data = pd.read_excel("Equipos y Herramientas 3Play _ Provision _ Mantencion_2023 (respuestas) (14).xlsx", engine='openpyxl')

# Verificar los nombres de las columnas
print("Nombres de las columnas:")
print([col for col in data.columns])  # Mostrar los nombres de las columnas en una lista

# Limpiar espacios en los nombres de las columnas
data.columns = data.columns.str.strip()

# Verificar si la columna 'Marca temporal' existe
if 'Marca temporal' in data.columns:
    print("\nLa columna 'Marca temporal' existe.")
    # Mostrar los primeros valores de la columna 'Marca temporal'
    print("\nPrimeros valores de 'Marca temporal':")
    print(data['Marca temporal'].head())
else:
    print("\nLa columna 'Marca temporal' NO existe en el DataFrame.")



    import pandas as pd
from collections import Counter
import streamlit as st
import re

# Palabras clave para ausencia de EPP
epp_ausencia_keywords = ["no tiene", "sin casco", "sin gafas", "sin guantes", "sin arnés", "sin zapatos", 
                         "no usa casco", "no usa gafas", "no usa guantes", "no usa arnés", "no usa zapatos", 
                         "falta de", "no cuenta con"]  # Ajustado para ausencia de EPP

# Palabras clave para cumplimiento de observaciones
cumple_keywords = ["sin observacion", "sin observación"]

# Función para procesar los datos
def process_data(uploaded_file):
    # Leer el archivo subido
    df = pd.read_excel(uploaded_file)

    # Nombre de la columna de observaciones
    observaciones_column_name = "Observaciones /  Separe con comas los temas"

    # Verificar si la columna existe
    if observaciones_column_name not in df.columns:
        st.error(f"No se encontró la columna '{observaciones_column_name}' en el archivo Excel. Por favor, asegúrese de que el archivo contiene una columna con este nombre.")
        return

    # Normalizar observaciones (convertir NaN a cadena vacía y poner en minúsculas)
    df['Observaciones'] = df[observaciones_column_name].fillna('').str.lower()

    # Función auxiliar para verificar problemas de EPP
    def check_epp_incompleto(observation):
        # Solo considerar registros que mencionen específicamente palabras clave relacionadas con EPP
        return any(keyword in observation for keyword in epp_ausencia_keywords)

    # Filtrar registros donde hay problemas de EPP
    epp_issues_df = df[df['Observaciones'].apply(check_epp_incompleto)]

    # Mostrar el porcentaje de técnicos con problemas de EPP
    total_registros = len(df)
    registros_epp_incompleto = len(epp_issues_df)
    percentage_epp_incompleto = (registros_epp_incompleto / total_registros) * 100

    st.write(f"**% de Técnicos que No Utilizan EPP Completo: **{percentage_epp_incompleto:.2f}%")

    # Mostrar los registros específicos de problemas de EPP
    st.write("### Registros para KPI de Problemas de EPP")
    if len(epp_issues_df) > 0:
        st.dataframe(epp_issues_df[[observaciones_column_name]])
    else:
        st.write("No se encontraron registros de problemas de EPP.")


