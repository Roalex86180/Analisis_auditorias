import streamlit as st
import pandas as pd
import io
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.express as px  #
import re
from collections import Counter
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

    # --- Correcciones de tipos para evitar ArrowTypeError ---
    # Asegurarnos de que las columnas de nombre/técnico sean strings
    col_tecnico = "Nombre de Técnico/Copiar el del Wfm"
    data[col_tecnico] = data[col_tecnico].fillna("").astype(str)
    # Y también cualquier columna de iconos/técnico final
    # (Se volverá a convertir tras crearlas, pero las dejamos listas)
    
    with tab1:
        # Opciones de filtro
        tecnicos = sorted(data[col_tecnico].unique())
        tecnico = st.selectbox("👷‍♂️ Técnico", ["Todos"] + tecnicos)
        empresas = sorted(data["Empresa"].fillna("").astype(str).unique())
        empresa = st.selectbox("🏢 Empresa", ["Todas"] + empresas)
        tipo_auditoria = sorted(data["Tipo de Auditoria"].fillna("").astype(str).unique())
        tipo = st.selectbox("🔍 Tipo de Auditoría", ["Todas"] + tipo_auditoria)
        patente = st.text_input("🚗 Buscar por Patente").strip()
        orden_trabajo = st.text_input("📄 Buscar por Número de Orden de Trabajo / ID Externo").strip()

        # Aplicar filtros
        df_filtrado = data.copy()
        if tecnico != "Todos":
            df_filtrado = df_filtrado[df_filtrado[col_tecnico] == tecnico]
        if empresa != "Todas":
            df_filtrado = df_filtrado[df_filtrado["Empresa"] == empresa]
        if tipo != "Todas":
            df_filtrado = df_filtrado[df_filtrado["Tipo de Auditoria"] == tipo]
        if patente:
            df_filtrado = df_filtrado[
                df_filtrado["Patente Camioneta"].astype(str).str.contains(patente, case=False)
            ]
        if orden_trabajo:
            df_filtrado = df_filtrado[
                df_filtrado["Número de Orden de Trabajo/ ID externo"]
                .astype(str)
                .str.contains(orden_trabajo, case=False)
            ]
        st.markdown("### 📊 Datos filtrados")
        st.dataframe(df_filtrado, use_container_width=True)

        # ----------------- Ranking Técnicos más Auditados -----------------
        st.markdown("### 🏆 Ranking Técnicos más Auditados")
        if col_tecnico in data.columns and "Empresa" in data.columns and "Fecha" in data.columns:
            ranking = (
                data.groupby([col_tecnico, "Empresa"])
                .size()
                .reset_index(name="Cantidad de Auditorías")
                .rename(columns={col_tecnico: "Técnico"})
                .sort_values(by="Cantidad de Auditorías", ascending=False)
            )
            # Fechas de Auditoría
            fechas = (
                data.groupby([col_tecnico, "Empresa"])["Fecha"]
                .apply(lambda x: ", ".join(pd.to_datetime(x).dt.strftime("%d/%m/%Y")))
                .reset_index(name="Fechas de Auditoría")
            )
            ranking["Fechas de Auditoría"] = fechas["Fechas de Auditoría"]
            st.dataframe(ranking, use_container_width=True)
        else:
            st.error("Faltan columnas necesarias para el ranking de técnicos.")

        # ----------------- KPI Auditorías por Empresa -----------------
        st.markdown("### 🏢 Auditorías por Empresa")
        auditorias_empresa = (
            data["Empresa"]
            .value_counts()
            .rename_axis("Empresa")
            .reset_index(name="Cantidad de Auditorías")
        )
        st.dataframe(auditorias_empresa, use_container_width=True)
        st.subheader("📈 Auditorías por Empresa")
        fig = px.bar(
            auditorias_empresa,
            x="Cantidad de Auditorías",
            y="Empresa",
            orientation="h",
            color="Empresa",
            text="Cantidad de Auditorías",
            color_discrete_sequence=px.colors.qualitative.Vivid,
        )
        fig.update_layout(
            xaxis_title="Cantidad de Auditorías",
            yaxis_title="Empresa",
            yaxis=dict(autorange="reversed"),
            plot_bgcolor="white",
        )
        st.plotly_chart(fig, use_container_width=True)

        # ----------------- Técnicos con Stock Crítico de Herramientas -----------------
        st.markdown("### 🔧 Técnicos con Stock Crítico de Herramientas")

        herramientas_criticas = [
            "Power meter GPON",
            "VFL Luz visible para localizar fallas",
            "Limpiador de conectores tipo “One Click”",
            "Deschaquetador de primera cubierta para DROP",
            "Deschaquetador de recubrimiento de FO 125micras Tipo Miller",
            "Cortadora de precisión 3 pasos",
            "Regla de corte",
            "Alcohol isopropilico 99%",
            "Paños secos para FO",
            "Crimper para cable UTP",
            "Deschaquetador para cables con cubierta redonda (UTP, RG6 )",
            "Tester para cable UTP",
        ]

        def obtener_herramientas_faltantes(tecnico_data):
            faltantes = []
            for h in herramientas_criticas:
                if h not in tecnico_data or pd.isna(tecnico_data[h]) or tecnico_data[h] in ["No", "Falta", "0"]:
                    faltantes.append(h)
            return faltantes

        stock_h = (
            data.groupby([col_tecnico, "Empresa"], group_keys=False)
            .apply(lambda grp: pd.Series({
                "Herramientas Faltantes": obtener_herramientas_faltantes(grp.iloc[0])
            }))
            .reset_index()
            .rename(columns={col_tecnico: "Técnico"})
        )
        stock_h["Cantidad Faltantes"] = stock_h["Herramientas Faltantes"].map(len)
        stock_h = stock_h[stock_h["Cantidad Faltantes"] > 0].sort_values(by="Cantidad Faltantes", ascending=False)

        def agregar_icono_herr(row):
            if row["Cantidad Faltantes"] >= 2:
                return f"🔴 {row['Técnico']}"
            elif row["Cantidad Faltantes"] == 1:
                return f"🟡 {row['Técnico']}"
            else:
                return row["Técnico"]

        stock_h["Técnico Con Icono"] = stock_h.apply(agregar_icono_herr, axis=1)
        stock_h["Herramientas Faltantes"] = stock_h["Herramientas Faltantes"].apply(lambda lst: ", ".join(lst))

        # Asegurarnos de que "Técnico Con Icono" sea string
        stock_h["Técnico Con Icono"] = stock_h["Técnico Con Icono"].astype(str)

        total_tecs_h = stock_h.shape[0]
        st.markdown(f"**🔥 Total técnicos con stock crítico de herramientas: {total_tecs_h}**")

        empresas_h = stock_h["Empresa"].unique()
        empresa_sel_h = st.selectbox("🔎 Filtrar por Empresa (Herramientas):", options=["Todas"] + list(empresas_h))
        if empresa_sel_h != "Todas":
            stock_h = stock_h[stock_h["Empresa"] == empresa_sel_h]

        st.dataframe(stock_h[["Técnico Con Icono", "Empresa", "Herramientas Faltantes"]], use_container_width=True)

        # Descarga Excel
        buf_h = io.BytesIO()
        with pd.ExcelWriter(buf_h, engine="xlsxwriter") as writer:
            stock_h[["Técnico Con Icono", "Empresa", "Herramientas Faltantes"]].rename(
                columns={"Técnico Con Icono": "Técnico"}
            ).to_excel(writer, index=False, sheet_name="Stock_Critico_Herramientas")
        buf_h.seek(0)
        st.download_button(
            label="📥 Descargar Técnicos con Stock Crítico Herramientas",
            data=buf_h,
            file_name="tecnicos_stock_critico_herramientas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        # Gráfico de stock crítico herramientas
        df_emp_h = (
            stock_h.groupby("Empresa")
            .size()
            .reset_index(name="Cantidad de Técnicos con Stock Crítico Herramientas")
            .sort_values(by="Cantidad de Técnicos con Stock Crítico Herramientas", ascending=False)
        )
        fig_h = px.bar(
            df_emp_h,
            x="Cantidad de Técnicos con Stock Crítico Herramientas",
            y="Empresa",
            orientation="h",
            color="Empresa",
            text="Cantidad de Técnicos con Stock Crítico Herramientas",
            color_discrete_sequence=px.colors.qualitative.Vivid,
        )
        fig_h.update_layout(
            xaxis_title="Cantidad de Técnicos con Stock Crítico de Herramientas",
            yaxis_title="Empresa",
            yaxis=dict(autorange="reversed"),
            plot_bgcolor="white",
        )
        st.plotly_chart(fig_h, use_container_width=True)

        # --- STOCK CRÍTICO EPP ---
        st.markdown("### 🦺 Técnicos con Stock Crítico de EPP")
        epp_criticos = [
            "Conos de seguridad", "Refugio de PVC", "Casco de Altura", "Barbiquejo",
            "Legionario Para Casco", "Guantes Cabritilla", "Guantes Dielectricos",
            "Guantes trabajo Fino", "Zapatos de Seguridad Dielectricos",
            "LENTE DE SEGURIDAD (CLAROS Y OSCUROS)", "Arnes Dielectrico",
            "Estrobo Dielectrico", "Cuerda de vida /Dielectrico", "Chaleco reflectante",
            "DETECTOR DE TENSION TIPO LAPIZ CON LINTERNA", "Bloqueador Solar"
        ]
        def obtener_epp_faltantes(tecnico_data):
            falt = []
            for e in epp_criticos:
                if e not in tecnico_data or pd.isna(tecnico_data[e]) or tecnico_data[e] in ["No", "Falta", "0"]:
                    falt.append(e)
            return falt

        stock_e = (
            data.groupby([col_tecnico, "Empresa"], group_keys=False)
            .apply(lambda grp: pd.Series({
                "EPP Faltantes": obtener_epp_faltantes(grp.iloc[0])
            }))
            .reset_index()
            .rename(columns={col_tecnico: "Técnico"})
        )
        stock_e["Cantidad Faltantes"] = stock_e["EPP Faltantes"].map(len)
        stock_e = stock_e[stock_e["Cantidad Faltantes"] > 0].sort_values(by="Cantidad Faltantes", ascending=False)

        epp_vitales = ["Casco de Altura", "Zapatos de Seguridad Dielectricos", "Arnes Dielectrico", "Estrobo Dielectrico"]
        def agregar_icono_epp(row):
            falt_v = [e for e in row["EPP Faltantes"] if e in epp_vitales]
            if len(falt_v) >= 2:
                return f"🔴 {row['Técnico']}"
            elif len(falt_v) == 1:
                return f"🟡 {row['Técnico']}"
            else:
                return row["Técnico"]

        stock_e["Técnico Con Icono"] = stock_e.apply(agregar_icono_epp, axis=1).astype(str)
        stock_e["EPP Faltantes"] = stock_e["EPP Faltantes"].apply(lambda lst: ", ".join(lst))

        total_tecs_e = stock_e.shape[0]
        st.markdown(f"**🔥 Total técnicos con stock crítico de EPP: {total_tecs_e}**")

        empresas_e = stock_e["Empresa"].unique()
        empresa_sel_e = st.selectbox("🔎 Filtrar por Empresa (EPP):", options=["Todas"] + list(empresas_e))
        if empresa_sel_e != "Todas":
            stock_e = stock_e[stock_e["Empresa"] == empresa_sel_e]

        st.dataframe(stock_e[["Técnico Con Icono", "Empresa", "EPP Faltantes"]], use_container_width=True)

        buf_e = io.BytesIO()
        with pd.ExcelWriter(buf_e, engine="xlsxwriter") as writer:
            stock_e[["Técnico Con Icono", "Empresa", "EPP Faltantes"]].rename(
                columns={"Técnico Con Icono": "Técnico"}
            ).to_excel(writer, index=False, sheet_name="Stock_Critico_EPP")
        buf_e.seek(0)
        st.download_button(
            label="📥 Descargar Técnicos con Stock Crítico EPP",
            data=buf_e,
            file_name="tecnicos_stock_critico_epp.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        # Gráfico de stock crítico EPP
        df_emp_e = (
            stock_e.groupby("Empresa")
            .size()
            .reset_index(name="Cantidad de Técnicos con Stock Crítico EPP")
            .sort_values(by="Cantidad de Técnicos con Stock Crítico EPP", ascending=False)
        )
        fig_e = px.bar(
            df_emp_e,
            x="Cantidad de Técnicos con Stock Crítico EPP",
            y="Empresa",
            orientation="h",
            color="Empresa",
            text="Cantidad de Técnicos con Stock Crítico EPP",
            color_discrete_sequence=px.colors.qualitative.Vivid,
        )
        fig_e.update_layout(
            xaxis_title="Cantidad de Técnicos con Stock Crítico de EPP",
            yaxis_title="Empresa",
            yaxis=dict(autorange="reversed"),
            plot_bgcolor="white",
        )
        st.plotly_chart(fig_e, use_container_width=True)

        # --- Resumen General de KPIs ---
        st.markdown("---")
        st.subheader("📊 Resumen General de Stock Crítico")
        st.metric(label="🔥 Total Técnicos con EPP Crítico", value=total_tecs_e)
        st.metric(label="🚀 Total Técnicos con Herramientas Críticas", value=total_tecs_h)

        # Llamada a process_data si se desea
        kpis, empresa_kpis_df, total_auditorias, data = process_data(archivo)

    with tab2:
        # 🎯 Dashboard de Auditores
        st.markdown("### 🧑‍💼 Ranking de Auditores por Trabajos Realizados")
        ranking_auditores = (
            data.groupby("Información del Auditor")
            .size()
            .reset_index(name="Cantidad de Auditorías")
            .rename(columns={"Información del Auditor": "Auditor"})
            .sort_values(by="Cantidad de Auditorías", ascending=False)
        )
        st.dataframe(ranking_auditores, use_container_width=True)

        if "Información del Auditor" in data.columns and "Empresa" in data.columns and "Fecha" in data.columns:
            distribucion = data.groupby(
                ["Información del Auditor", "Empresa"], group_keys=False
            ).agg(
                Cantidad_de_Auditorias=("Fecha", "size"),
                Fechas_de_Auditoria=("Fecha", lambda x: ", ".join(pd.to_datetime(x).dt.strftime("%d/%m/%Y"))),
            ).reset_index()
            st.write("KPI de Distribución de Auditorías entre Empresas con Fechas")
            st.dataframe(distribucion, use_container_width=True)
        else:
            st.error("Faltan columnas para KPI de distribución de auditorías.")

        # KPI Auditarías por Región
        st.subheader("🌎 Auditorías por Región")
        aud_reg = (
            data.groupby("Region")
            .size()
            .reset_index(name="Cantidad de Auditorías")
            .sort_values(by="Cantidad de Auditorías", ascending=False)
        )
        fig_reg = px.bar(
            aud_reg,
            x="Cantidad de Auditorías",
            y="Region",
            orientation="h",
            color="Region",
            text="Cantidad de Auditorías",
            color_discrete_sequence=px.colors.qualitative.Set2,
        )
        fig_reg.update_layout(
            xaxis_title="Cantidad de Auditorías",
            yaxis_title="Región",
            yaxis=dict(autorange="reversed"),
            plot_bgcolor="white",
        )
        st.plotly_chart(fig_reg, use_container_width=True)

else:
    st.warning("⚠️ Por favor, sube un archivo Excel con las auditorías.")