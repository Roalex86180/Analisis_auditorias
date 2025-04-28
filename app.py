import streamlit as st
import pandas as pd
import io
import seaborn as sns
import plotly.express as px  # 
from kpi import process_data

# Configuración inicial de la app
st.set_page_config(page_title="Auditorías Técnicos", layout="wide")

st.title("📋 Análisis de Auditorías de Técnicos de Telecomunicaciones")

# Subir archivo
archivo = st.file_uploader("📁 Sube el archivo Excel con los datos de auditoría", type=["xlsx"])

tab1, tab2 = st.tabs(["📋 Información de Técnicos", "🛠️ Información de Auditores"])

if archivo:
    # Cargar todas las hojas del archivo Excel
    xls = pd.ExcelFile(archivo)
    hojas = xls.sheet_names
    data = pd.concat([xls.parse(hoja) for hoja in hojas], ignore_index=True)

    

    # Normalizar columnas
    data.columns = data.columns.str.strip()

    with tab1:
        # Opciones de filtro
        tecnicos = sorted(data["Nombre de Técnico/Copiar el del Wfm"].dropna().astype(str).unique())
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
            df_filtrado = df_filtrado[df_filtrado["Nombre de Técnico/Copiar el del Wfm"].astype(str) == tecnico]
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
        if 'Nombre de Técnico/Copiar el del Wfm' in data.columns and 'Empresa' in data.columns and 'Fecha' in data.columns:
            # Agrupar por Técnico y Empresa, y contar auditorías
            ranking = (
                data.groupby(["Nombre de Técnico/Copiar el del Wfm", "Empresa"])
                .size()
                .reset_index(name="Cantidad de Auditorías")
                .rename(columns={"Nombre de Técnico/Copiar el del Wfm": "Técnico"})
                .sort_values(by="Cantidad de Auditorías", ascending=False)
            )

            # Agregar la columna de Fechas de Auditoría
            ranking["Fechas de Auditoría"] = (
                data.groupby(["Nombre de Técnico/Copiar el del Wfm", "Empresa"])["Fecha"]
                .apply(lambda x: ', '.join(pd.to_datetime(x).dt.strftime('%d/%m/%Y')))
                .reset_index(name="Fechas de Auditoría")["Fechas de Auditoría"]
            )

            st.dataframe(ranking, use_container_width=True)
            # Mostrar el Ranking de Técnicos más Auditados
        else:
            st.error("Faltan columnas necesarias para calcular el Ranking de Técnicos más Auditados ('Nombre de Técnico/Copiar el del Wfm', 'Empresa' y 'Fecha').")

        # ----------------- KPI Auditorías por Empresa -----------------
        st.markdown("### 🏢 Auditorías por Empresa")
        auditorias_empresa = (
            data["Empresa"]
            .value_counts()
            .rename_axis('Empresa')
            .reset_index(name='Cantidad de Auditorías')
        )
        st.dataframe(auditorias_empresa, use_container_width=True)

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

        stock_critico_herramientas = (
            data.groupby(["Nombre de Técnico/Copiar el del Wfm", "Empresa"])
            .apply(lambda x: obtener_herramientas_faltantes(x.iloc[0]))
            .reset_index()
            .rename(columns={0: "Herramientas Faltantes", "Nombre de Técnico/Copiar el del Wfm": "Técnico"})
        )

        stock_critico_herramientas = stock_critico_herramientas[stock_critico_herramientas["Herramientas Faltantes"].map(len) > 0]

        stock_critico_herramientas["Cantidad Faltantes"] = stock_critico_herramientas["Herramientas Faltantes"].map(len)

        stock_critico_herramientas = stock_critico_herramientas.sort_values(by="Cantidad Faltantes", ascending=False)

        def agregar_icono_herramientas(row):
            if row["Cantidad Faltantes"] >= 2:
                return f"🔴 {row['Técnico']}"
            elif row["Cantidad Faltantes"] == 1:
                return f"🟡 {row['Técnico']}"
            else:
                return row['Técnico']

        stock_critico_herramientas["Técnico Con Icono"] = stock_critico_herramientas.apply(agregar_icono_herramientas, axis=1)

        stock_critico_herramientas["Herramientas Faltantes"] = stock_critico_herramientas["Herramientas Faltantes"].apply(lambda x: ", ".join(x))

        # KPI Herramientas
        total_tecnicos_stock_critico_herramientas = stock_critico_herramientas.shape[0]
        st.markdown(f"**🔥 Total técnicos con stock crítico de herramientas: {total_tecnicos_stock_critico_herramientas}**")

        # Filtro empresa
        empresas_disponibles_herramientas = stock_critico_herramientas['Empresa'].unique()
        empresa_seleccionada_herramientas = st.selectbox("🔎 Filtrar por Empresa (Herramientas):", options=["Todas"] + list(empresas_disponibles_herramientas))

        stock_critico_herramientas_general = stock_critico_herramientas.copy()

        if empresa_seleccionada_herramientas != "Todas":
            stock_critico_herramientas = stock_critico_herramientas[stock_critico_herramientas["Empresa"] == empresa_seleccionada_herramientas]

        st.dataframe(
            stock_critico_herramientas[["Técnico Con Icono", "Empresa", "Herramientas Faltantes"]],
            use_container_width=True
        )

        buffer_herramientas = io.BytesIO()
        with pd.ExcelWriter(buffer_herramientas, engine='xlsxwriter') as writer:
            stock_critico_herramientas[["Técnico Con Icono", "Empresa", "Herramientas Faltantes"]].rename(columns={"Técnico Con Icono": "Técnico"}).to_excel(writer, index=False, sheet_name='Stock_Critico_Herramientas')
        buffer_herramientas.seek(0)

        st.download_button(
            label="📥 Descargar Técnicos con Stock Crítico Herramientas",
            data=buffer_herramientas,
            file_name="tecnicos_stock_critico_herramientas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.subheader("📈 Técnicos con Stock Crítico de Herramientas por Empresa")

        empresas_stock_critico_herramientas = (
            stock_critico_herramientas_general.groupby('Empresa')
            .size()
            .reset_index(name='Cantidad de Técnicos con Stock Crítico Herramientas')
            .sort_values(by='Cantidad de Técnicos con Stock Crítico Herramientas', ascending=False)
        )

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

        # Crear tabla de técnicos con EPP faltantes
        stock_critico_epp = (
            data.groupby(["Nombre de Técnico/Copiar el del Wfm", "Empresa"])
            .apply(lambda x: obtener_epp_faltantes(x.iloc[0]))
            .reset_index()
            .rename(columns={0: "EPP Faltantes", "Nombre de Técnico/Copiar el del Wfm": "Técnico"})
        )

        # Filtrar solo técnicos con EPP faltantes
        stock_critico_epp = stock_critico_epp[stock_critico_epp["EPP Faltantes"].map(len) > 0]

        # Contar cantidad de faltantes
        stock_critico_epp["Cantidad Faltantes"] = stock_critico_epp["EPP Faltantes"].map(len)

        # Ordenar de más a menos faltantes
        stock_critico_epp = stock_critico_epp.sort_values(by="Cantidad Faltantes", ascending=False)

        # Agregar el circulito 🔴🟡
        def agregar_icono(row):
            faltantes_vitales = [epp for epp in row["EPP Faltantes"] if epp in epp_vitales]
            if len(faltantes_vitales) >= 2:
                return f"🔴 {row['Técnico']}"
            elif len(faltantes_vitales) == 1:
                return f"🟡 {row['Técnico']}"
            else:
                return row['Técnico']

        stock_critico_epp["Técnico Con Icono"] = stock_critico_epp.apply(agregar_icono, axis=1)

        # Convertir EPP faltantes a texto
        stock_critico_epp["EPP Faltantes"] = stock_critico_epp["EPP Faltantes"].apply(lambda x: ", ".join(x))

        # Guardar copia general para gráficos
        stock_critico_epp_general = stock_critico_epp.copy()

        # Filtro por empresa
        empresas_disponibles_epp = stock_critico_epp['Empresa'].unique()
        empresa_seleccionada_epp = st.selectbox("🔎 Filtrar por Empresa (EPP):", options=["Todas"] + list(empresas_disponibles_epp))

        if empresa_seleccionada_epp != "Todas":
            stock_critico_epp = stock_critico_epp[stock_critico_epp["Empresa"] == empresa_seleccionada_epp]

        # --- KPI EPP ---
        total_tecnicos_stock_critico_epp = stock_critico_epp_general.shape[0]
        st.markdown(f"**🔥 Total técnicos con stock crítico de EPP: {total_tecnicos_stock_critico_epp}**")

        # Mostrar tabla
        st.dataframe(
            stock_critico_epp[["Técnico Con Icono", "Empresa", "EPP Faltantes"]],
            use_container_width=True
        )

        # Botón de descarga
        buffer_epp = io.BytesIO()
        with pd.ExcelWriter(buffer_epp, engine='xlsxwriter') as writer:
            stock_critico_epp[["Técnico Con Icono", "Empresa", "EPP Faltantes"]].rename(columns={"Técnico Con Icono": "Técnico"}).to_excel(writer, index=False, sheet_name='Stock_Critico_EPP')
        buffer_epp.seek(0)

        st.download_button(
            label="📥 Descargar Técnicos con Stock Crítico EPP",
            data=buffer_epp,
            file_name="tecnicos_stock_critico_epp.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Gráfico por empresa
        st.subheader("📈 Técnicos con Stock Crítico de EPP por Empresa (Interactivo)")

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

        # ----------------- 🎯 Dashboard de Cumplimiento -----------------
            # Ranking de auditores por trabajos realizados
        st.markdown("### 🧑‍💼 Ranking de Auditores por Trabajos Realizados")
        ranking_auditores = (
            data.groupby("Información del Auditor")
            .size()
            .reset_index(name="Cantidad de Auditorías")
            .rename(columns={"Información del Auditor": "Auditor"})
            .sort_values(by="Cantidad de Auditorías", ascending=False)
        )
        st.dataframe(ranking_auditores, use_container_width=True)

        if 'Información del Auditor' in data.columns and 'Empresa' in data.columns and 'Fecha' in data.columns:
            # Agrupar por auditor y empresa y concatenar las fechas de las auditorías
            distribucion_auditorias = data.groupby(['Información del Auditor', 'Empresa']).agg(
                Cantidad_de_Auditorias=('Fecha', 'size'),
                Fechas_de_Auditoria=('Fecha', lambda x: ', '.join(pd.to_datetime(x).dt.strftime('%d/%m/%Y')))
            ).reset_index()

            # Mostrar el KPI
            st.write("KPI de Distribución de Auditorías entre Empresas con Fechas")
            st.dataframe(distribucion_auditorias)
        else:
            st.error("Faltan columnas necesarias para calcular el KPI de distribución de auditorías ('Información del Auditor', 'Empresa' y 'Fecha').")

            # ----------------- KPI: Auditorías por Región -----------------
        st.subheader("🌎 Auditorías por Región")

            # Agrupar datos por Región y contar cantidad de auditorías
        auditorias_por_region = (
                data.groupby('Region')
                .size()
                .reset_index(name='Cantidad de Auditorías')
                .sort_values(by='Cantidad de Auditorías', ascending=False)
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


else:
    st.warning("⚠️ Por favor, sube un archivo Excel con las auditorías.")


