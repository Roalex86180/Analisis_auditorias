import pandas as pd
import streamlit as st
import re
import plotly.express as px

# Palabras clave
tools_keywords = ["herramienta", "falta de herramienta", "herramientas"]
epp_keywords = ["epp", "equipos de protección", "protección"]
vehicle_order_keywords = ["camioneta", "orden de camioneta", "desorden en camioneta"]
not_completed_keywords = ["no realizada", "pendiente", "no ejecutada"]
agenda_keywords = ["no cumple agenda", "no llego", "retraso"]
gpon_keywords = ["one click", "cortadora 3 pasos", "vfl", "lápiz luz", "deschaquetadora de primera capa", "deschaquetadora miller", "power meter", "paños secos", "bolso kit gpon"]
malas_practicas_keywords = ["no utiliza", "no usa", "sin "]
epp_ausencia_keywords = ["no utiliza", "sin legnionario", "sin lentes", "sin casco", "sin gafas", "sin guantes", "sin arnés", "sin zapatos", "no usa casco", "no usa gafas", "no usa guantes", "no usa arnés", "no usa zapatos", "falta de", "no cuenta con"]
cumple_keywords = ["sin observacion"]
observaciones_a_excluir = ["sin obs", "sin comentarios", "sin observaciónes", "sin observaciones"]

def process_data(uploaded_file):
    df = pd.read_excel(uploaded_file)

    # Asegurarse de que las columnas relevantes son de tipo str
    df['Observaciones'] = df['Observaciones /  Separe con comas los temas'].fillna('').str.lower()  # Normaliza las observaciones
    df['Nombre de Técnico/Copiar el del Wfm'] = df['Nombre de Técnico/Copiar el del Wfm'].astype(str)  # Asegura que sea un string
    df['Empresa'] = df['Empresa'].astype(str)  # Asegura que 'Empresa' sea string
    df['Region'] = df['Region'].astype(str)  # Asegura que 'Region' sea string

    total_auditorias = len(df)

    # Funciones
    def match_keywords(x, keywords):
        return any(keyword in x for keyword in keywords)

    def match_malas_practicas(x):
        return match_keywords(x, malas_practicas_keywords) and not any(obs in x for obs in observaciones_a_excluir)

    def match_gpon(x):
        return match_keywords(x, gpon_keywords)

    def match_epp_incompleto(x):
        return match_keywords(x, epp_ausencia_keywords)

    def match_cumple(x):
        return any(phrase in x for phrase in ["sin observacion", "sin obs", "sin comentarios", "sin observaciónes", "sin observaciones"])

    # KPI
    kpis = {
        "Falta de Herramientas": df['Observaciones'].apply(lambda x: match_keywords(x, tools_keywords)),
        "Problemas de Orden en Camioneta": df['Observaciones'].apply(lambda x: match_keywords(x, vehicle_order_keywords)),
        "Auditorías No Realizadas": df['Observaciones'].apply(lambda x: match_keywords(x, not_completed_keywords)),
        "Técnicos con Malas Prácticas": df['Observaciones'].apply(match_malas_practicas),
        "Técnicos que No Cumplen Agenda": df['Observaciones'].apply(lambda x: match_keywords(x, agenda_keywords)),
        "Técnicos que No Utilizan Kit GPON Completo": df['Observaciones'].apply(match_gpon),
        "Técnicos que No Utilizan EPP Completo": df['Observaciones'].apply(match_epp_incompleto),
        "Técnicos que Cumplen": df['Observaciones'].apply(match_cumple),
    }

    # Agrupar por empresa y kpi
    empresa_kpis = {empresa: {k: 0 for k in kpis.keys()} for empresa in df['Empresa'].unique()}

    # Contar casos por empresa
    for empresa in empresa_kpis:
        for kpi, cases in kpis.items():
            empresa_kpis[empresa][kpi] = cases[df['Empresa'] == empresa].sum()

    # Crear DataFrame de casos por empresa
    empresa_kpis_df = pd.DataFrame(empresa_kpis).T

    # Reemplazar los NaN por 0
    empresa_kpis_df = empresa_kpis_df.fillna(0)

    # Ordenar por total de casos
    empresa_kpis_df['Total Casos'] = empresa_kpis_df.sum(axis=1)
    empresa_kpis_df = empresa_kpis_df.sort_values(by="Total Casos", ascending=False)

    # Layout principal
    st.title("📊 Reporte de Auditorías Técnicas")
    st.markdown("---")

    cols = st.columns(4)
    for idx, (label, value) in enumerate(kpis.items()):
        percentage = (value.sum() / total_auditorias) * 100 if total_auditorias else 0
        with cols[idx % 4]:
            st.metric(label=label, value=f"{percentage:.2f}%", delta=f"{value.sum()} casos", delta_color="inverse")

    st.markdown("---")

    # 📋 Observaciones Detalladas
    st.header("📋 Observaciones Detalladas")
    
    with st.expander("🔴 Técnicos que No Utilizan Kit GPON Completo"):
        st.write(f"- Técnicos que no usan el kit GPON o siguen los procedimientos correctos: {kpis['Técnicos que No Utilizan Kit GPON Completo'].sum()} casos")
        df_filtered = df[df['Observaciones'].apply(match_gpon)]
        st.dataframe(df_filtered[['Nombre de Técnico/Copiar el del Wfm', 'Observaciones /  Separe con comas los temas', 'Empresa', 'Region']].fillna(''))

    with st.expander("🔴 Técnicos que No Cumplen Agenda"):
        st.write(f"- Técnicos que no cumplen la agenda: {kpis['Técnicos que No Cumplen Agenda'].sum()} casos")
        df_filtered = df[df['Observaciones'].apply(lambda x: match_keywords(x, agenda_keywords))]
        st.dataframe(df_filtered[['Nombre de Técnico/Copiar el del Wfm', 'Observaciones /  Separe con comas los temas', 'Empresa', 'Region']].fillna(''))

    with st.expander("🔴 Técnicos que No Utilizan EPP Completo"):
        st.write(f"- Técnicos que no utilizan sus EPP completos: {kpis['Técnicos que No Utilizan EPP Completo'].sum()} casos")
        df_filtered = df[df['Observaciones'].apply(match_epp_incompleto)]
        st.dataframe(df_filtered[['Nombre de Técnico/Copiar el del Wfm', 'Observaciones /  Separe con comas los temas', 'Empresa', 'Region']].fillna(''))

    with st.expander("🔴 Falta de Herramientas"):
        st.write(f"- Falta de Herramientas: {kpis['Falta de Herramientas'].sum()} casos")
        df_filtered = df[df['Observaciones'].apply(lambda x: match_keywords(x, tools_keywords))]
        st.dataframe(df_filtered[['Nombre de Técnico/Copiar el del Wfm', 'Observaciones /  Separe con comas los temas', 'Empresa', 'Region']].fillna(''))

    with st.expander("🔴 Problemas de Orden en Camioneta"):
        st.write(f"- Problemas de Orden en Camioneta: {kpis['Problemas de Orden en Camioneta'].sum()} casos")
        df_filtered = df[df['Observaciones'].apply(lambda x: match_keywords(x, vehicle_order_keywords))]
        st.dataframe(df_filtered[['Nombre de Técnico/Copiar el del Wfm', 'Observaciones /  Separe con comas los temas', 'Empresa', 'Region']].fillna(''))

    with st.expander("🔴 Auditorías No Realizadas"):
        st.write(f"- Auditorías No Realizadas: {kpis['Auditorías No Realizadas'].sum()} casos")
        df_filtered = df[df['Observaciones'].apply(lambda x: match_keywords(x, not_completed_keywords))]
        st.dataframe(df_filtered[['Nombre de Técnico/Copiar el del Wfm', 'Observaciones /  Separe con comas los temas', 'Empresa', 'Region']].fillna(''))

    with st.expander("🔴 Técnicos con Malas Prácticas"):
        st.write(f"- Técnicos con Malas Prácticas: {kpis['Técnicos con Malas Prácticas'].sum()} casos")
        df_filtered = df[df['Observaciones'].apply(match_malas_practicas)]
        st.dataframe(df_filtered[['Nombre de Técnico/Copiar el del Wfm', 'Observaciones /  Separe con comas los temas', 'Empresa', 'Region']].fillna(''))

    with st.expander("🔴 Técnicos que Cumplen"):
        st.write(f"- Técnicos que Cumplen: {kpis['Técnicos que Cumplen'].sum()} casos")
        df_filtered = df[df['Observaciones'].apply(match_cumple)]
        st.dataframe(df_filtered[['Nombre de Técnico/Copiar el del Wfm', 'Observaciones /  Separe con comas los temas', 'Empresa', 'Region']].fillna(''))

    # Análisis de casos por empresa con gráfico
    st.markdown("---")
    st.header("📊 Ranking de Empresas por Casos")
    
    # Crear gráfico de barras apiladas por empresa
    fig = px.bar(
        empresa_kpis_df, 
        x=empresa_kpis_df.index, 
        y=list(kpis.keys()),  # Convertimos kpis.keys() a lista
        title="Ranking de Empresas por Casos",
        labels={'value': 'Número de Casos', 'Empresa': 'Empresa'},
        color_discrete_map={
            'Falta de Herramientas': 'red', 
            'Problemas de Orden en Camioneta': 'orange', 
            'Auditorías No Realizadas': 'yellow', 
            'Técnicos con Malas Prácticas': 'green', 
            'Técnicos que No Cumplen Agenda': 'blue', 
            'Técnicos que No Utilizan Kit GPON Completo': 'purple', 
            'Técnicos que No Utilizan EPP Completo': 'pink', 
            'Técnicos que Cumplen': 'gray'
        },
        height=600,
        barmode='stack'  # Aseguramos que las barras sean apiladas
    )

    # Mostrar gráfico
    st.plotly_chart(fig)

    return kpis, empresa_kpis_df, total_auditorias, df
















