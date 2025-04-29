import streamlit as st
import pandas as pd
import io
import seaborn as sns
import plotly.express as px  
from pt import process_data
import unicodedata


# Función para normalizar los nombres (eliminar acentos y tildes)
def normalizar_texto(texto):
    if isinstance(texto, str):
        # Eliminar acentos y tildes
        return ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    return texto

# Configuración inicial de la app
st.set_page_config(page_title="Auditorías Técnicos", layout="wide")

st.title("📋 Análisis de Auditorías de Técnicos de Telecomunicaciones")

# Subir archivo
archivo = st.file_uploader("📁 Sube el archivo Excel con los datos de auditoría", type=["xlsx"])

tab1, tab2, = st.tabs(["📋 Información de Técnicos", "🛠️ Información de Auditores",])

if archivo:
    # Cargar todas las hojas del archivo Excel
    xls = pd.ExcelFile(archivo)
    hojas = xls.sheet_names
    data = pd.concat([xls.parse(hoja).astype(str) for hoja in hojas], ignore_index=True)

    # Normalizar columnas
    data.columns = data.columns.str.strip()

            # Definir una función para generar contexto compacto
    def generar_contexto(data):
            return f"""
        Estás analizando una base de datos con {data.shape[0]} registros de auditorías técnicas a técnicos de telecomunicaciones.
        Cada registro contiene:
        - Información del técnico, auditor, empresa y fecha.
        - Lista de herramientas y materiales verificados.
        - Observaciones y cumplimiento de estándares.
        - Elementos de protección personal (EPP) y condiciones del vehículo.
        Tu tarea es responder preguntas con base en los datos cargados.
        """

        # Consultar a Ollama con contexto y pregunta del usuario
    def consultar_ollama(contexto, pregunta):
            response = ollama.chat(
                model='mistral',
                messages=[
                    {"role": "system", "content": contexto},
                    {"role": "user", "content": pregunta}
                ]
            )
            return response['message']['content']

    with tab1:
        # Opciones de filtro
        tecnicos = sorted(data["Nombre de Técnico/Copiar el del Wfm"].dropna().apply(normalizar_texto).unique())
        tecnico = st.selectbox("👷‍♂️ Técnico", ["Todos"] + tecnicos)

        empresas = sorted(data["Empresa"].dropna().astype(str).unique())
        empresa = st.selectbox("🏢 Empresa", ["Todas"] + empresas)

        tipo_auditoria = sorted(data["Tipo de Auditoria"].dropna().astype(str).unique())
        tipo = st.selectbox("🔍 Tipo de Auditoría", ["Todas"] + tipo_auditoria)

        patente = st.text_input("🚗 Buscar por Patente").strip()

        orden_trabajo = st.text_input("📄 Buscar por Número de Orden de Trabajo / ID Externo").strip()

        # Filtros
        df_filtrado = data.copy()

        if tecnico != "Todos":
            df_filtrado = df_filtrado[df_filtrado["Nombre de Técnico/Copiar el del Wfm"].apply(normalizar_texto) == normalizar_texto(tecnico)]
        if empresa != "Todas":
            df_filtrado = df_filtrado[df_filtrado["Empresa"].astype(str) == empresa]
        if tipo != "Todas":
            df_filtrado = df_filtrado[df_filtrado["Tipo de Auditoria"].astype(str) == tipo]
        if patente:
            df_filtrado = df_filtrado[df_filtrado["Patente Camioneta"].astype(str).str.contains(patente, case=False)]
        if orden_trabajo:
            df_filtrado = df_filtrado[df_filtrado["Número de Orden de Trabajo/ ID externo"].astype(str).str.contains(orden_trabajo, case=False)]

        st.markdown("### 📊 Datos filtrados")
        st.dataframe(df_filtrado, use_container_width=True)

        # ----------------- Ranking Técnicos más Auditados -----------------
        st.markdown("### 🏆 Ranking Técnicos más Auditados")

        # Verificamos que existan las columnas necesarias
        columnas_necesarias = ['Nombre de Técnico/Copiar el del Wfm', 'Empresa', 'Fecha', 'Estado de Auditoria']

        if all(col in data.columns for col in columnas_necesarias):
            
            # 1. Filtramos solo las auditorías FINALIZADAS
            data['Estado de Auditoria'] = data['Estado de Auditoria'].str.strip().str.lower()
            data_finalizadas = data[data['Estado de Auditoria'] == 'finalizada'].copy()

            # Aseguramos que 'Fecha' esté en formato datetime
            data_finalizadas['Fecha'] = pd.to_datetime(data_finalizadas['Fecha'], errors='coerce')

            # Selección de rango de fechas
            fecha_min = data_finalizadas['Fecha'].min()
            fecha_max = data_finalizadas['Fecha'].max()

            # Mostramos el selector de fechas, pero no lo hacemos obligatorio
            fechas = st.date_input(
                "📅 Selecciona el rango de fechas (opcional)",
                value=[fecha_min, fecha_max],
                min_value=fecha_min,
                max_value=fecha_max
            )

            # Si se seleccionaron fechas, filtramos por el rango
            if isinstance(fechas, list) and len(fechas) == 2:
                fecha_inicio, fecha_fin = fechas
                # Filtro por rango de fecha
                mask = (data_finalizadas['Fecha'] >= pd.to_datetime(fecha_inicio)) & (data_finalizadas['Fecha'] <= pd.to_datetime(fecha_fin))
                data_finalizadas = data_finalizadas.loc[mask]

            # Si el dataframe no está vacío, mostramos el ranking
            if not data_finalizadas.empty:
                # Agrupamos por Técnico y Empresa
                ranking = (
                    data_finalizadas
                    .groupby(["Nombre de Técnico/Copiar el del Wfm", "Empresa"])
                    .agg(
                        Cantidad_de_Auditorias=('Fecha', 'count'),
                        Fecha_de_Auditorias=('Fecha', lambda x: ', '.join(sorted(pd.to_datetime(x).dt.strftime('%d/%m/%Y'))))
                    )
                    .reset_index()
                    .rename(columns={
                        "Nombre de Técnico/Copiar el del Wfm": "Técnico",
                        "Empresa": "Empresa",
                        "Cantidad_de_Auditorias": "Cantidad de Auditorías",
                        "Fecha_de_Auditorias": "Fechas de Auditoría"
                    })
                    .sort_values(by="Cantidad de Auditorías", ascending=False)
                )

                st.dataframe(ranking, use_container_width=True)
            else:
                st.warning("⚠️ No hay auditorías en el rango de fechas seleccionado. Mostrando el ranking con todas las auditorías finalizadas.")
                # Si no hay auditorías en el rango, mostramos el ranking con todas las auditorías finalizadas sin filtrar
                st.dataframe(
                    data_finalizadas.groupby(["Nombre de Técnico/Copiar el del Wfm", "Empresa"])
                    .agg(
                        Cantidad_de_Auditorias=('Fecha', 'count'),
                        Fecha_de_Auditorias=('Fecha', lambda x: ', '.join(sorted(pd.to_datetime(x).dt.strftime('%d/%m/%Y'))))
                    )
                    .reset_index()
                    .rename(columns={
                        "Nombre de Técnico/Copiar el del Wfm": "Técnico",
                        "Empresa": "Empresa",
                        "Cantidad_de_Auditorias": "Cantidad de Auditorías",
                        "Fecha_de_Auditorias": "Fechas de Auditoría"
                    })
                    .sort_values(by="Cantidad de Auditorías", ascending=False),
                    use_container_width=True
                )
        else:
            st.error("Faltan columnas necesarias para calcular el Ranking de Técnicos más Auditados ('Nombre de Técnico/Copiar el del Wfm', 'Empresa', 'Fecha' y 'Estado de Auditoria').")



        # ----------------- KPI Auditorías por Empresa -----------------
            st.markdown("### 🏢 Auditorías por Empresa")

                # ----------------- KPI Auditorías por Empresa -----------------
        st.markdown("### 🏢 Auditorías por Empresa")

        columnas_necesarias_empresa = ["Empresa", "Estado de Auditoria"]

        if all(col in data.columns for col in columnas_necesarias_empresa):
            # Normalizamos la columna Estado de Auditoria
            data['Estado de Auditoria'] = data['Estado de Auditoria'].str.strip().str.lower()

            # Filtramos solo las auditorías finalizadas
            data_finalizadas_empresa = data[data['Estado de Auditoria'] == 'finalizada'].copy()

            if not data_finalizadas_empresa.empty:
                auditorias_empresa = (
                    data_finalizadas_empresa["Empresa"]
                    .value_counts()
                    .rename_axis('Empresa')
                    .reset_index(name='Cantidad de Auditorías')
                )

                st.dataframe(auditorias_empresa, use_container_width=True)
            else:
                st.warning("⚠️ No hay auditorías finalizadas para mostrar auditorías por empresa.")
        else:
            st.error("⚠️ Faltan columnas necesarias para calcular auditorías por empresa ('Empresa' y 'Estado de Auditoria').")


        # Gráfico de barras interactivo con Plotly
        st.subheader("📈 Auditorías por Empresa")
        fig = px.bar(
            auditorias_empresa,
            x='Cantidad de Auditorías',
            y='Empresa',
            orientation='h',
            color='Empresa',
            text='Cantidad de Auditorías',
            color_discrete_sequence=px.colors.qualitative.Vivid
        )

        fig.update_layout(
            xaxis_title="Cantidad de Auditorías",
            yaxis_title="Empresa",
            yaxis=dict(autorange="reversed"),  # Para que la empresa con más auditorías quede arriba
            plot_bgcolor='white'
        )

        st.plotly_chart(fig, use_container_width=True)


                # ----------------- KPI Stock Crítico de Herramientas -----------------
                # ----------------- KPI Stock Crítico de Herramientas -----------------
        data['Nombre de Técnico/Copiar el del Wfm'] = data['Nombre de Técnico/Copiar el del Wfm'].apply(normalizar_texto)        
        st.markdown("### 🔧 Técnicos con Stock Crítico de Herramientas (con detalle de herramientas faltantes)")

        herramientas_criticas = [
            "Power meter GPON", "VFL Luz visible para localizar fallas", "Limpiador de conectores tipo “One Click”",
            "Deschaquetador de primera cubierta para DROP", "Deschaquetador de recubrimiento de FO 125micras Tipo Miller",
            "Cortadora de precisión 3 pasos", "Regla de corte", "Alcohol isopropilico 99%",
            "Paños secos para FO", "Crimper para cable UTP", "Deschaquetador para cables con cubierta redonda (UTP, RG6 )",
            "Tester para cable UTP"
        ]

        def obtener_herramientas_faltantes(tecnico_data):
            faltantes = []
            for herramienta in herramientas_criticas:
                if herramienta not in tecnico_data or pd.isna(tecnico_data[herramienta]) or tecnico_data[herramienta] in ["No", "Falta", "0"]:
                    faltantes.append(herramienta)
            return faltantes

        # Asegurarse que Fecha es de tipo datetime
        data['Fecha'] = pd.to_datetime(data['Fecha'], dayfirst=True, errors='coerce')  # Asegúrate de que está en formato de fecha correcto

        # Limpiar posibles espacios y caracteres invisibles en los nombres de técnicos
        data['Nombre de Técnico/Copiar el del Wfm'] = data['Nombre de Técnico/Copiar el del Wfm'].str.strip()

        # Filtrar solo auditorías finalizadas
        data_filtrada = data[data['Estado de Auditoria'].str.strip().str.lower() == 'finalizada'].copy()

        # Para cada técnico, buscar la última fecha de auditoría finalizada
        idx = data_filtrada.groupby('Nombre de Técnico/Copiar el del Wfm')['Fecha'].idxmax()
        data_ultima_auditoria = data_filtrada.loc[idx].reset_index(drop=True)

        # Procesamos stock crítico
        stock_critico_herramientas = data_ultima_auditoria[["Nombre de Técnico/Copiar el del Wfm", "Empresa", "Fecha"] + herramientas_criticas].copy()
        stock_critico_herramientas["Herramientas Faltantes"] = stock_critico_herramientas.apply(obtener_herramientas_faltantes, axis=1)
        stock_critico_herramientas = stock_critico_herramientas[stock_critico_herramientas["Herramientas Faltantes"].map(len) > 0]

        # Conteo
        stock_critico_herramientas["Cantidad Faltantes"] = stock_critico_herramientas["Herramientas Faltantes"].map(len)

        # Orden
        stock_critico_herramientas = stock_critico_herramientas.sort_values(by="Cantidad Faltantes", ascending=False)

        # Renombramos
        stock_critico_herramientas = stock_critico_herramientas.rename(columns={"Nombre de Técnico/Copiar el del Wfm": "Técnico"})

        # Agregar icono
        def agregar_icono_herramientas(row):
            if row["Cantidad Faltantes"] >= 2:
                return f"🔴 {row['Técnico']}"
            elif row["Cantidad Faltantes"] == 1:
                return f"🟡 {row['Técnico']}"
            else:
                return row['Técnico']

        stock_critico_herramientas["Técnico Con Icono"] = stock_critico_herramientas.apply(agregar_icono_herramientas, axis=1)

        stock_critico_herramientas["Herramientas Faltantes"] = stock_critico_herramientas["Herramientas Faltantes"].apply(lambda x: ", ".join(x))

        # KPI
        total_tecnicos_stock_critico_herramientas = stock_critico_herramientas.shape[0]
        st.markdown(f"**🔥 Total técnicos con stock crítico de herramientas: {total_tecnicos_stock_critico_herramientas}**")

        # Filtro empresa
        empresas_disponibles_herramientas = stock_critico_herramientas['Empresa'].unique()
        empresa_seleccionada_herramientas = st.selectbox("🔎 Filtrar por Empresa (Herramientas):", options=["Todas"] + list(empresas_disponibles_herramientas))

        stock_critico_herramientas_general = stock_critico_herramientas.copy()

        if empresa_seleccionada_herramientas != "Todas":
            stock_critico_herramientas = stock_critico_herramientas[stock_critico_herramientas["Empresa"] == empresa_seleccionada_herramientas]

        # Mostrar dataframe
        st.dataframe(
            stock_critico_herramientas[["Técnico Con Icono", "Empresa", "Fecha", "Herramientas Faltantes"]],
            use_container_width=True
        )

        # Botón descargar
        buffer_herramientas = io.BytesIO()
        with pd.ExcelWriter(buffer_herramientas, engine='xlsxwriter') as writer:
            stock_critico_herramientas[["Técnico Con Icono", "Empresa", "Fecha", "Herramientas Faltantes"]].rename(columns={"Técnico Con Icono": "Técnico"}).to_excel(writer, index=False, sheet_name='Stock_Critico_Herramientas')
        buffer_herramientas.seek(0)

        st.download_button(
            label="📥 Descargar Técnicos con Stock Crítico Herramientas",
            data=buffer_herramientas,
            file_name="tecnicos_stock_critico_herramientas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        data['Nombre de Técnico/Copiar el del Wfm'] = data['Nombre de Técnico/Copiar el del Wfm'].apply(normalizar_texto)
        st.subheader("📈 Técnicos con Stock Crítico de Herramientas por Empresa")

        empresas_stock_critico_herramientas = (
            stock_critico_herramientas_general.groupby('Empresa')
            .size()
            .reset_index(name='Cantidad de Técnicos con Stock Crítico Herramientas')
            .sort_values(by='Cantidad de Técnicos con Stock Crítico Herramientas', ascending=False)
        )

        st.dataframe(empresas_stock_critico_herramientas, use_container_width=True)

               

        fig_stock_herramientas = px.bar(
            empresas_stock_critico_herramientas,
            x='Cantidad de Técnicos con Stock Crítico Herramientas',
            y='Empresa',
            orientation='h',
            color='Empresa',
            text='Cantidad de Técnicos con Stock Crítico Herramientas',
            color_discrete_sequence=px.colors.qualitative.Vivid
        )

        fig_stock_herramientas.update_layout(
            xaxis_title="Cantidad de Técnicos con Stock Crítico de Herramientas",
            yaxis_title="Empresa",
            yaxis=dict(autorange="reversed"),
            plot_bgcolor='white'
        )

        st.plotly_chart(fig_stock_herramientas, use_container_width=True)

        
        # --- STOCK CRÍTICO EPP ---

        # ----------------- KPI Stock Crítico de EPP -----------------
        data['Nombre de Técnico/Copiar el del Wfm'] = data['Nombre de Técnico/Copiar el del Wfm'].apply(normalizar_texto)
        st.markdown("### 🦺 Técnicos con Stock Crítico de EPP (con detalle de elementos faltantes)")

        # Lista de EPP críticos a considerar
        epp_criticos = [
            "Conos de seguridad", "Refugio de PVC", "Casco de Altura", "Barbiquejo",
            "Legionario Para Casco", "Guantes Cabritilla", "Guantes Dielectricos",
            "Guantes trabajo Fino", "Zapatos de Seguridad Dielectricos",
            "LENTE DE SEGURIDAD (CLAROS Y OSCUROS)", "Arnes Dielectrico",
            "Estrobo Dielectrico", "Cuerda de vida /Dielectrico", "Chaleco reflectante",
            "DETECTOR DE TENSION TIPO LAPIZ CON LINTERNA", "Bloqueador Solar"
        ]

        # Definir EPP vitales
        epp_vitales = ["Casco de Altura", "Zapatos de Seguridad Dielectricos", "Arnes Dielectrico", "Estrobo Dielectrico"]

        # Función para verificar EPP faltantes
        def obtener_epp_faltantes(tecnico_data):
            faltantes = []
            for epp in epp_criticos:
                if epp not in tecnico_data or pd.isna(tecnico_data[epp]) or tecnico_data[epp] in ["No", "Falta", "0"]:
                    faltantes.append(epp)
            return faltantes

        # Asegurarse que Fecha es tipo datetime
        data['Fecha'] = pd.to_datetime(data['Fecha'], dayfirst=True, errors='coerce')

        # Limpiar posibles espacios en nombres de técnicos
        data['Nombre de Técnico/Copiar el del Wfm'] = data['Nombre de Técnico/Copiar el del Wfm'].str.strip()

        # Filtrar solo auditorías finalizadas
        data_filtrada_epp = data[data['Estado de Auditoria'].str.strip().str.lower() == 'finalizada'].copy()

        # Para cada técnico, tomar la última fecha de auditoría finalizada
        idx_epp = data_filtrada_epp.groupby('Nombre de Técnico/Copiar el del Wfm')['Fecha'].idxmax()
        data_ultima_auditoria_epp = data_filtrada_epp.loc[idx_epp].reset_index(drop=True)

        # Procesar stock crítico EPP
        stock_critico_epp = data_ultima_auditoria_epp[["Nombre de Técnico/Copiar el del Wfm", "Empresa", "Fecha"] + epp_criticos].copy()
        stock_critico_epp["EPP Faltantes"] = stock_critico_epp.apply(obtener_epp_faltantes, axis=1)
        stock_critico_epp = stock_critico_epp[stock_critico_epp["EPP Faltantes"].map(len) > 0]

        # Contar cantidad de faltantes
        stock_critico_epp["Cantidad Faltantes"] = stock_critico_epp["EPP Faltantes"].map(len)

        # Ordenar de más a menos
        stock_critico_epp = stock_critico_epp.sort_values(by="Cantidad Faltantes", ascending=False)

        # Renombrar técnico
        stock_critico_epp = stock_critico_epp.rename(columns={"Nombre de Técnico/Copiar el del Wfm": "Técnico"})

        # Agregar el circulito 🔴🟡
        def agregar_icono_epp(row):
            faltantes_vitales = [epp for epp in row["EPP Faltantes"] if epp in epp_vitales]
            if len(faltantes_vitales) >= 2:
                return f"🔴 {row['Técnico']}"
            elif len(faltantes_vitales) == 1:
                return f"🟡 {row['Técnico']}"
            else:
                return row['Técnico']

        stock_critico_epp["Técnico Con Icono"] = stock_critico_epp.apply(agregar_icono_epp, axis=1)

        # Convertir lista de EPP a texto
        stock_critico_epp["EPP Faltantes"] = stock_critico_epp["EPP Faltantes"].apply(lambda x: ", ".join(x))

        # KPI Total técnicos
        total_tecnicos_stock_critico_epp = stock_critico_epp.shape[0]
        st.markdown(f"**🔥 Total técnicos con stock crítico de EPP: {total_tecnicos_stock_critico_epp}**")

        # Filtro por empresa
        empresas_disponibles_epp = stock_critico_epp['Empresa'].unique()
        empresa_seleccionada_epp = st.selectbox("🔎 Filtrar por Empresa (EPP):", options=["Todas"] + list(empresas_disponibles_epp))

        stock_critico_epp_general = stock_critico_epp.copy()

        if empresa_seleccionada_epp != "Todas":
            stock_critico_epp = stock_critico_epp[stock_critico_epp["Empresa"] == empresa_seleccionada_epp]

        # Mostrar dataframe
        st.dataframe(
            stock_critico_epp[["Técnico Con Icono", "Empresa", "Fecha", "EPP Faltantes"]],
            use_container_width=True
        )

        # Botón descargar
        buffer_epp = io.BytesIO()
        with pd.ExcelWriter(buffer_epp, engine='xlsxwriter') as writer:
            stock_critico_epp[["Técnico Con Icono", "Empresa", "Fecha", "EPP Faltantes"]].rename(columns={"Técnico Con Icono": "Técnico"}).to_excel(writer, index=False, sheet_name='Stock_Critico_EPP')
        buffer_epp.seek(0)

        st.download_button(
            label="📥 Descargar Técnicos con Stock Crítico EPP",
            data=buffer_epp,
            file_name="tecnicos_stock_critico_epp.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Gráfico por empresa
        st.subheader("📈 Técnicos con Stock Crítico de EPP por Empresa")

        empresas_stock_critico_epp = (
            stock_critico_epp_general.groupby('Empresa')
            .size()
            .reset_index(name='Cantidad de Técnicos con Stock Crítico EPP')
            .sort_values(by='Cantidad de Técnicos con Stock Crítico EPP', ascending=False)
        )

        fig_stock_epp = px.bar(
            empresas_stock_critico_epp,
            x='Cantidad de Técnicos con Stock Crítico EPP',
            y='Empresa',
            orientation='h',
            color='Empresa',
            text='Cantidad de Técnicos con Stock Crítico EPP',
            color_discrete_sequence=px.colors.qualitative.Vivid
        )

        fig_stock_epp.update_layout(
            xaxis_title="Cantidad de Técnicos con Stock Crítico de EPP",
            yaxis_title="Empresa",
            yaxis=dict(autorange="reversed"),
            plot_bgcolor='white'
        )

        st.plotly_chart(fig_stock_epp, use_container_width=True)

        # --- Resumen General de KPIs ---
        st.markdown("---")
        st.subheader("📊 Resumen General de Stock Crítico")

        st.metric(label="🔥 Total Técnicos con EPP Crítico", value=total_tecnicos_stock_critico_epp)
        st.metric(label="🚀 Total Técnicos con Herramientas Críticas", value=total_tecnicos_stock_critico_herramientas)

        

        if archivo:
            # Llamamos a la función de KPIs
            kpis, empresa_kpis_df, total_auditorias, data = process_data(archivo)   

    with tab2:

        # # ----------------- 🎯 Dashboard de Cumplimiento -----------------
        # Ranking de auditores por trabajos realizados
        st.markdown("### Ranking de Auditores por Trabajos Realizados")

        # Filtrar auditorías finalizadas
        data_finalizadas = data[data['Estado de Auditoria'] == 'finalizada']

        # Agrupar por auditor y contar las auditorías finalizadas
        ranking_auditores = (
            data_finalizadas.groupby("Información del Auditor")  # Agrupar por Auditor
            .size()  # Contar las auditorías
            .reset_index(name="Cantidad de Auditorías")  # Resetear el índice y renombrar la columna
            .rename(columns={"Información del Auditor": "Auditor"})  # Renombrar la columna
            .sort_values(by="Cantidad de Auditorías", ascending=False)  # Ordenar por cantidad de auditorías
        )

        st.dataframe(ranking_auditores, use_container_width=True)

        # Verificar si las columnas necesarias existen para el KPI de distribución de auditorías
        if 'Información del Auditor' in data.columns and 'Empresa' in data.columns and 'Fecha' in data.columns:
            # Agrupar por auditor y empresa y concatenar las fechas de las auditorías (solo las finalizadas)
            distribucion_auditorias = data_finalizadas.groupby(['Información del Auditor', 'Empresa']).agg(
                Cantidad_de_Auditorias=('Fecha', 'size'),
                Fechas_de_Auditoria=('Fecha', lambda x: ', '.join(pd.to_datetime(x).dt.strftime('%d/%m/%Y')))
            ).reset_index()

            # Mostrar el KPI
            st.markdown("Distribución de Auditorías entre Empresas con Fechas")
            st.dataframe(distribucion_auditorias)
        else:
            st.error("Faltan columnas necesarias para calcular el KPI de distribución de auditorías ('Información del Auditor', 'Empresa' y 'Fecha').")

            # ----------------- KPI: Auditorías por Región -----------------
            
        st.subheader("🌎 Auditorías por Región")

        # Filtrar auditorías finalizadas
        data_finalizadas = data[data['Estado de Auditoria'] == 'finalizada']

        # Agrupar datos por Región y contar cantidad de auditorías finalizadas
        auditorias_por_region = (
            data_finalizadas.groupby('Region')  # Agrupar por Región
            .size()  # Contar las auditorías
            .reset_index(name='Cantidad de Auditorías')  # Resetear el índice y renombrar la columna
            .sort_values(by='Cantidad de Auditorías', ascending=False)  # Ordenar por cantidad de auditorías
        )

        # Crear gráfico de barras horizontal
        fig_auditorias_region = px.bar(
            auditorias_por_region,
            x='Cantidad de Auditorías',
            y='Region',
            orientation='h',
            color='Region',
            text='Cantidad de Auditorías',
            color_discrete_sequence=px.colors.qualitative.Set2
        )

        fig_auditorias_region.update_layout(
            xaxis_title="Cantidad de Auditorías",
            yaxis_title="Región",
            yaxis=dict(autorange="reversed"),
            plot_bgcolor='white'
        )

        st.plotly_chart(fig_auditorias_region, use_container_width=True)

        # Calcular el total de auditorías finalizadas
        total_auditorias_finalizadas = len(data[data['Estado de Auditoria'] == 'finalizada'])

        # Mostrar la etiqueta con el total de auditorías finalizadas
        st.markdown(f"""
            <div style="background-color: #f0f0f5; padding: 15px 25px; border-radius: 8px; font-size: 24px; font-weight: bold; color: #333;">
                <span style="color: #007bff;">Total de Auditorías Finalizadas: </span><span style="color: #28a745;">{total_auditorias_finalizadas}</span>
            </div>
        """, unsafe_allow_html=True)


        st.subheader("📋 Ranking de Auditores por información completa")

        # Filtrar solo auditorías finalizadas
        data_finalizadas = data[data["Estado de Auditoria"].str.lower() == "finalizada"]

        # Calcular el porcentaje de campos llenos por cada auditoría
        total_columnas = data_finalizadas.shape[1]
        data_finalizadas["% Completitud"] = data_finalizadas.notna().sum(axis=1) / total_columnas * 100

        # Agrupar por auditor y calcular el promedio
        ranking_completitud = data_finalizadas.groupby("Información del Auditor")["% Completitud"].mean().reset_index()
        ranking_completitud = ranking_completitud.sort_values(by="% Completitud", ascending=False)

        # Formatear porcentaje con coma y símbolo %, y aplicar color azul
        def formato_porcentaje(valor):
            return f"{valor:,.1f}%".replace('.', ',')  # Cambia . por , y añade %

        def estilo_azul(val):
            return 'color: blue; font-weight: bold;' if isinstance(val, float) else ''

        # Mostrar tabla con estilo
        st.dataframe(
            ranking_completitud.style
            .format({"% Completitud": formato_porcentaje})
            .applymap(estilo_azul, subset=["% Completitud"]),
            use_container_width=True
        )




else:
    st.warning("⚠️ Por favor, sube un archivo Excel con las auditorías.")



