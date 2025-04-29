import pandas as pd
import streamlit as st
import plotly.express as px
import unicodedata

# Palabras clave por categoría
tools_keywords = ["herramienta", "falta de herramienta", "herramientas"]
epp_keywords = ["epp", "equipos de protección", "protección"]
vehicle_order_keywords = ["camioneta", "orden de camioneta", "desorden en camioneta"]
agenda_keywords = ["no cumple agenda", "no llego", "retraso"]
gpon_keywords = ["one click", "cortadora 3 pasos", "vfl", "lápiz luz", "deschaquetadora de primera capa", "deschaquetadora miller", "power meter", "paños secos", "bolso kit gpon"]
malas_practicas_keywords = ["no utiliza", "no usa", "sin "]
epp_ausencia_keywords = ["no utiliza", "sin legnionario", "sin lentes", "sin casco", "sin gafas", "sin guantes", "sin arnés", "sin zapatos", "no usa casco", "no usa gafas", "no usa guantes", "no usa arnés", "no usa zapatos", "falta de", "no cuenta con"]
cumple_keywords = ["sin observacion", "sin obs", "sin comentarios", "sin observaciónes", "sin observaciones", "so", "s/o", "so,", "s/o,", "s/o.", "."]
observaciones_a_excluir = ["sin obs", "sin comentarios", "sin observaciónes", "sin observaciones", "so", "s/o", "so,", "s/o,", "s/o.", "."]

def normalize_text(text):
    """Normaliza texto: minúsculas, elimina acentos, convierte a string si es necesario."""
    if isinstance(text, str):
        text = text.lower()
        text = unicodedata.normalize('NFKD', text).encode('ascii', 'ignore').decode('utf-8')
    else:
        text = ""
    return text

def process_data(uploaded_file):
    df = pd.read_excel(uploaded_file)

    # Normalización de campos clave
    df['Observaciones'] = df['Observaciones /  Separe con comas los temas'].fillna('').apply(normalize_text)
    df['Nombre de Técnico/Copiar el del Wfm'] = df['Nombre de Técnico/Copiar el del Wfm'].astype(str).apply(normalize_text)
    df['Empresa'] = df['Empresa'].astype(str).apply(normalize_text)
    df['Region'] = df['Region'].astype(str).apply(normalize_text)
    df['Estado de Auditoria'] = df['Estado de Auditoria'].astype(str).apply(normalize_text)

    # Dividir por estado de auditoría
    df_finalizadas = df[df['Estado de Auditoria'] == "finalizada"]
    df_no_realizadas = df[df['Estado de Auditoria'] != "finalizada"]
    total_auditorias = len(df)

    # Funciones de evaluación
    def match_keywords(x, keywords):
        return any(keyword in x for keyword in keywords)

    def match_cumple(x):
        """Detecta si una observación es cumplimiento explícito."""
        x = normalize_text(x)
        return x.strip() in cumple_keywords

    def match_malas_practicas(x):
        """Detecta malas prácticas si no cumple y contiene malas prácticas."""
        x = normalize_text(x)
        if x.strip() in cumple_keywords or x == "":
            return False
        return any(keyword in x for keyword in malas_practicas_keywords)

    def match_gpon(x):
        """Detecta si la observación menciona elementos del kit GPON"""
        return match_keywords(x, gpon_keywords)

    def match_epp_incompleto(x):
        """Detecta si falta EPP (Equipo de Protección Personal)"""
        return match_keywords(x, epp_ausencia_keywords)

    # KPIs globales
    kpis = {
        "Falta de Herramientas": df_finalizadas['Observaciones'].apply(lambda x: match_keywords(x, tools_keywords)),
        "Problemas de Orden en Camioneta": df_finalizadas['Observaciones'].apply(lambda x: match_keywords(x, vehicle_order_keywords)),
        "Auditorías No Realizadas": df_no_realizadas['Estado de Auditoria'].apply(lambda x: True),
        "Técnicos con Malas Prácticas": df_finalizadas['Observaciones'].apply(match_malas_practicas),
        "Técnicos que No Cumplen Agenda": df_finalizadas['Observaciones'].apply(lambda x: match_keywords(x, agenda_keywords)),
        "Técnicos que No Utilizan Kit GPON Completo": df_finalizadas['Observaciones'].apply(match_gpon),
        "Técnicos que No Utilizan EPP Completo": df_finalizadas['Observaciones'].apply(match_epp_incompleto),
        "Técnicos que Cumplen": df_finalizadas['Observaciones'].apply(match_cumple),
    }

    # Inicialización de KPIs por empresa
    empresa_kpis = {empresa: {k: 0 for k in kpis} for empresa in df['Empresa'].unique()}

    # Recuento de casos por empresa
    for empresa in empresa_kpis:
        empresa_df = df[df['Empresa'] == empresa]
        empresa_finalizadas = empresa_df[empresa_df['Estado de Auditoria'] == "finalizada"]
        empresa_no_finalizadas = empresa_df[empresa_df['Estado de Auditoria'] != "finalizada"]

        empresa_kpis[empresa]["Falta de Herramientas"] = empresa_finalizadas['Observaciones'].apply(lambda x: match_keywords(x, tools_keywords)).sum()
        empresa_kpis[empresa]["Problemas de Orden en Camioneta"] = empresa_finalizadas['Observaciones'].apply(lambda x: match_keywords(x, vehicle_order_keywords)).sum()
        empresa_kpis[empresa]["Auditorías No Realizadas"] = len(empresa_no_finalizadas)
        empresa_kpis[empresa]["Técnicos con Malas Prácticas"] = empresa_finalizadas['Observaciones'].apply(match_malas_practicas).sum()
        empresa_kpis[empresa]["Técnicos que No Cumplen Agenda"] = empresa_finalizadas['Observaciones'].apply(lambda x: match_keywords(x, agenda_keywords)).sum()
        empresa_kpis[empresa]["Técnicos que No Utilizan Kit GPON Completo"] = empresa_finalizadas['Observaciones'].apply(match_gpon).sum()
        empresa_kpis[empresa]["Técnicos que No Utilizan EPP Completo"] = empresa_finalizadas['Observaciones'].apply(match_epp_incompleto).sum()
        empresa_kpis[empresa]["Técnicos que Cumplen"] = empresa_finalizadas['Observaciones'].apply(match_cumple).sum()

    # DataFrame de resumen por empresa
    empresa_kpis_df = pd.DataFrame(empresa_kpis).T.fillna(0)
    empresa_kpis_df['Total Casos'] = empresa_kpis_df.sum(axis=1)
    empresa_kpis_df = empresa_kpis_df.sort_values(by="Total Casos", ascending=False)

    # UI - Métricas generales
    st.title("📊 Reporte de Auditorías Técnicas")
    st.markdown("---")

    cols = st.columns(4)
    for idx, (label, values) in enumerate(kpis.items()):
        porcentaje = (values.sum() / total_auditorias) * 100 if total_auditorias else 0
        with cols[idx % 4]:
            st.metric(label=label, value=f"{porcentaje:.2f}%", delta=f"{values.sum()} casos", delta_color="inverse")

    st.markdown("---")
    st.header("📋 Observaciones Detalladas")

    # Expanders por KPI
    expander_info = [
        ("🔴 Técnicos que No Utilizan Kit GPON Completo", match_gpon),
        ("🔴 Técnicos que No Cumplen Agenda", lambda x: match_keywords(x, agenda_keywords)),
        ("🔴 Técnicos que No Utilizan EPP Completo", match_epp_incompleto),
        ("🔴 Falta de Herramientas", lambda x: match_keywords(x, tools_keywords)),
        ("🔴 Problemas de Orden en Camioneta", lambda x: match_keywords(x, vehicle_order_keywords)),
        ("🔴 Auditorías No Realizadas", lambda x: True),
        ("🔴 Técnicos con Malas Prácticas", match_malas_practicas),
        ("🔴 Técnicos que Cumplen", match_cumple),
    ]

    for title, func in expander_info:
        with st.expander(title):
            if title == "🔴 Auditorías No Realizadas":
                df_filtered = df_no_realizadas
            else:
                df_filtered = df_finalizadas[df_finalizadas['Observaciones'].apply(func)]
            st.write(f"- {title.split('🔴 ')[1]}: {len(df_filtered)} casos")
            st.dataframe(df_filtered[['Nombre de Técnico/Copiar el del Wfm', 'Observaciones /  Separe con comas los temas', 'Información del Auditor', 'Empresa', 'Region']].fillna(''))

    # Gráfico por empresa
    st.markdown("---")
    st.header("📊 Ranking de Empresas por Casos")

    fig = px.bar(
        empresa_kpis_df,
        x=empresa_kpis_df.index,
        y=list(kpis.keys()),
        title="Ranking de Empresas por Casos",
        labels={'value': 'Número de Casos', 'Empresa': 'Empresa'},
        height=600,
        color_discrete_sequence=px.colors.qualitative.Safe,
        barmode='stack'
    )
    st.plotly_chart(fig)

    return kpis, empresa_kpis_df, total_auditorias, df




