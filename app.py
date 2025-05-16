import pandas as pd
import streamlit as st
import plotly.express as px
import unicodedata
import io
from datetime import date
import traceback # Opcional: para debug de errores detallados
# Asegúrate de que xlsxwriter esté instalado: pip install xlsxwriter openpyxl
from pt import process_data


# --- Función de Normalización ---
def normalizar_texto(texto):
    """Normaliza texto: elimina espacios, acentos, tildes y convierte a minúsculas."""
    if isinstance(texto, str):
        texto = str(texto).strip().lower() # Convertir explícitamente a string por seguridad
        nfd_form = unicodedata.normalize('NFD', texto)
        return ''.join(c for c in nfd_form if unicodedata.category(c) != 'Mn')
    # Maneja casos donde el input no sea string, como NaN o None
    return '' # Devuelve string vacío si no es string para evitar errores en operaciones de string


# --- Configuración inicial de la app ---
st.set_page_config(page_title="Análisis Auditorías", layout="wide")
st.title("📊 Análisis de Auditorías de Técnicos")

# --- Carga de Datos con st.file_uploader y st.session_state ---
st.subheader("📁 Carga de Datos de Auditoría")
# Añadimos una key para que Streamlit gestione correctamente el estado del uploader
archivo = st.file_uploader("Sube el archivo Excel con los datos de auditoría", type=["xlsx"], key="excel_uploader")

# --- Lógica de Carga y Preprocesamiento del Archivo ---
# Solo cargamos y procesamos el archivo si se ha subido uno Y es diferente al que ya tenemos en session_state
if archivo is not None:
    # --- CORRECCIÓN AQUÍ: Usamos name y size para comparar si es un nuevo archivo ---
    # Verificamos si NO hay datos cargados, O si el nombre/tamaño del archivo subido
    # son diferentes a los que guardamos en session_state del archivo anterior.
    if 'data' not in st.session_state or \
       st.session_state.get('uploaded_file_name') != archivo.name or \
       st.session_state.get('uploaded_file_size') != archivo.size:

        st.info(f"Cargando y procesando archivo '{archivo.name}'...")
        try:
            # Leer el archivo Excel
            xls = pd.ExcelFile(archivo)
            hojas = xls.sheet_names
            # Cargar cada hoja y concatenar, permitiendo a pandas inferir tipos
            df_list = []
            for hoja in hojas:
                try:
                    df_list.append(xls.parse(hoja))
                except Exception as e:
                     st.warning(f"No se pudo leer la hoja '{hoja}': {e}")
                     continue # Saltar a la siguiente hoja si falla

            if df_list:
                 data = pd.concat(df_list, ignore_index=True)
                 
            else:
                 st.error("No se pudo cargar ninguna hoja del archivo Excel.")
                 data = pd.DataFrame() # Crear DataFrame vacío si no se cargó nada


            # --- Preparación General de Datos (Aplicar una sola vez al cargar) ---
            if not data.empty:
                # Normalizar nombres de columnas (Esto ya lo tenías)
                data.columns = data.columns.str.strip()

                # Normalizar columnas clave que se usarán en varios análisis
                cols_to_normalize_str = ['Nombre de Técnico/Copiar el del Wfm', 'Información del Auditor']
                for col in cols_to_normalize_str:
                    if col in data.columns:
                         # Rellenar posibles NaN antes de aplicar la normalización
                         # La función normalizar_texto ahora maneja no-strings y devuelve ''
                         data[col] = data[col].apply(normalizar_texto)
                    else:
                        # Añadir la columna si falta para evitar KeyErrors posteriores
                         data[col] = ''


                # Normalizar y limpiar estado de auditoría
                if 'Estado de Auditoria' in data.columns:
                    data['Estado de Auditoria'] = data['Estado de Auditoria'].astype(str).str.strip().str.lower()
                    data['Estado de Auditoria'] = data['Estado de Auditoria'].replace({'nan': 'desconocido', '': 'desconocido'})
                else:
                    data['Estado de Auditoria'] = 'desconocido' # Añadir si falta

                # Convertir la columna de fecha a datetime
                if 'Fecha' in data.columns:
                    data['Fecha'] = pd.to_datetime(data['Fecha'], errors='coerce') # NaT si falla
                # Notas: filas con Fecha=NaT o columnas no encontradas se manejarán en cada análisis que use fechas.
                    
                col_km = 'Kilometraje Camioneta'
                if col_km in data.columns:
                    # Intenta convertir la columna a tipo numérico.
                    # errors='coerce' convierte valores no válidos a NaN.
                    data[col_km] = pd.to_numeric(data[col_km], errors='coerce')
                    # Opcional: Si prefieres NaN a 0, comenta la línea de fillna(0)
                    # data[col_km] = data[col_km].fillna(0) # Si quieres 0 en lugar de NaN
                else:
                     # Añadir la columna si falta con valores nulos (para evitar errores)
                     data[col_km] = pd.NA # O np.nan

                     # --- Manejar la columna Número de Orden de Trabajo/ ID externo ---
                col_id_trabajo = 'Número de Orden de Trabajo/ ID externo' # Asegurarse de usar la variable definida antes
                if col_id_trabajo in data.columns:
                    # Convertir la columna a tipo string explícitamente
                    data[col_id_trabajo] = data[col_id_trabajo].astype(str)
                    # Opcional: Si después de convertir a string quedan valores 'nan' (por NaNs originales), puedes reemplazarlos por ''
                    data[col_id_trabajo] = data[col_id_trabajo].replace('nan', '')
                else:
                    # Si la columna falta, añadirla como vacía para evitar KeyErrors posteriores
                    data[col_id_trabajo] = ''

                    # --- NUEVA CORRECCIÓN: Manejar la columna Rut / tecnico ---
                col_rut_tecnico = 'Rut / tecnico' # Definir la variable para el nombre de la columna
                if col_rut_tecnico in data.columns:
                    # Convertir la columna ENTERA a tipo string
                    # Esto asegura que todos los valores (números, strings, NaNs) se conviertan a su representación de texto
                    data[col_rut_tecnico] = data[col_rut_tecnico].astype(str)
                    # Opcional: Reemplazar la representación de string de NaN ('nan') por una cadena vacía si lo prefieres
                    data[col_rut_tecnico] = data[col_rut_tecnico].replace('nan', '')
                else:
                    # Si la columna 'Rut / tecnico' falta, añadirla como columna de strings vacías
                    data[col_rut_tecnico] = ''

                # Opcional: Limpiar filas completamente vacías que podrían venir de hojas extra
                original_rows = len(data)
                data.dropna(how='all', inplace=True)
                if len(data) < original_rows:
                     st.info(f"Se eliminaron {original_rows - len(data)} filas completamente vacías.")


                # --- Almacenar el DataFrame procesado y la info del archivo en session_state ---
                st.session_state['data'] = data
                st.session_state['uploaded_file_name'] = archivo.name # Guardamos el nombre
                st.session_state['uploaded_file_size'] = archivo.size # Guardamos el tamaño
                # Ya NO guardamos archivo.id

                st.success("Archivo cargado y procesado correctamente.")

                # --- Re-ejecutar el script ---
                # Esto es crucial para que Streamlit actualice la interfaz y use los datos cargados
                st.rerun()

            else:
                 # Si data está vacía después de cargar, limpiar session_state
                 st.warning("⚠️ El archivo Excel cargado está vacío o no contiene datos procesables.")
                 if 'data' in st.session_state: del st.session_state['data']
                 if 'uploaded_file_name' in st.session_state: del st.session_state['uploaded_file_name']
                 if 'uploaded_file_size' in st.session_state: del st.session_state['uploaded_file_size']


        except Exception as e:
            st.error(f"Ocurrió un error al cargar o procesar el archivo: {e}")
            # Limpiar session_state en caso de error
            if 'data' in st.session_state: del st.session_state['data']
            if 'uploaded_file_name' in st.session_state: del st.session_state['uploaded_file_name']
            if 'uploaded_file_size' in st.session_state: del st.session_state['uploaded_file_size']
            # Mostrar el traceback completo para depuración si es necesario
            # st.code(traceback.format_exc())

    else:
        # Este mensaje se muestra si el mismo archivo ya está cargado y procesado
        st.info(f"El archivo '{archivo.name}' ya está cargado. Usa los filtros para explorar los datos.")


# --- Bloque Principal que se ejecuta SOLO si 'data' está en session_state ---
# Este bloque contiene todas las pestañas y su contenido
if 'data' in st.session_state:
    data = st.session_state['data'] # Recuperar el DataFrame de session_state

    # Verificar si el DataFrame no está vacío después de recuperarlo
    if not data.empty:

        # --- Realizar filtrados comunes aquí si aplican a ambas pestañas ---
        # Por ejemplo, filtrar por auditorías finalizadas, ya que se usa en varias secciones
        if 'Estado de Auditoria' in data.columns:
             # data['Estado de Auditoria'] ya está normalizada
             data_finalizadas = data[data['Estado de Auditoria'] == 'finalizada'].copy()
        else:
             # No hay columna de estado, data_finalizadas estará vacía
             data_finalizadas = pd.DataFrame()


        # --- Definición de Pestañas ---
        tab1, tab2 = st.tabs(["📋 Información de Técnicos", "🛠️ Información de Auditores"])

        # --- Contenido de la Pestaña 1 ---
        with tab1:
            st.header("📋 Información de Técnicos")

            # --- Filtros de la Pestaña 1 ---
            st.markdown("### 🔍 Filtros")
            # Nombres de columnas clave
            col_tec_nombre = 'Nombre de Técnico/Copiar el del Wfm'
            col_empresa = 'Empresa'
            col_tipo_auditoria = 'Tipo de Auditoria'
            col_patente = 'Patente Camioneta'
            col_orden_trabajo = 'Número de Orden de Trabajo/ ID externo'
            col_fecha = 'Fecha' # Necesaria para rango de fechas en ranking


            # Asegurarse de que las columnas existen antes de usar unique
            tecnicos = sorted(data[col_tec_nombre].unique().tolist()) if col_tec_nombre in data.columns else []
            tecnicos = [t for t in tecnicos if t.strip() != '' and t.lower() != 'nan'] # Limpiar vacíos/nan
            tecnico = st.selectbox("👷‍♂️ Técnico", ["Todos"] + tecnicos, key="filtro_tecnico_tab1") # Añadir key

            empresas = sorted(data[col_empresa].astype(str).unique().tolist()) if col_empresa in data.columns else []
            empresas = [e for e in empresas if e.strip() != '' and e.lower() != 'nan'] # Limpiar vacíos/nan
            empresa = st.selectbox("🏢 Empresa", ["Todas"] + empresas, key="filtro_empresa_tab1") # Añadir key

            tipos_auditoria = sorted(data[col_tipo_auditoria].astype(str).unique().tolist()) if col_tipo_auditoria in data.columns else []
            tipos_auditoria = [t for t in tipos_auditoria if t.strip() != '' and t.lower() != 'nan'] # Limpiar vacíos/nan
            tipo = st.selectbox("🔍 Tipo de Auditoría", ["Todas"] + tipos_auditoria, key="filtro_tipo_auditoria_tab1") # Añadir key

            patente = st.text_input("🚗 Buscar por Patente", key="filtro_patente_tab1").strip() if col_patente in data.columns else ""
            orden_trabajo = st.text_input("📄 Buscar por Número de Orden de Trabajo / ID Externo", key="filtro_orden_trabajo_tab1").strip() if col_orden_trabajo in data.columns else ""

            # Aplicar Filtros
            df_filtrado = data.copy() # Usar una copia para filtrar


            if tecnico != "Todos" and col_tec_nombre in df_filtrado.columns:
                 df_filtrado = df_filtrado[df_filtrado[col_tec_nombre] == tecnico]

            if empresa != "Todas" and col_empresa in df_filtrado.columns:
                 df_filtrado = df_filtrado[df_filtrado[col_empresa].astype(str) == empresa]

            if tipo != "Todas" and col_tipo_auditoria in df_filtrado.columns:
                 df_filtrado = df_filtrado[df_filtrado[col_tipo_auditoria].astype(str) == tipo]

            if patente and col_patente in df_filtrado.columns:
                 df_filtrado = df_filtrado[df_filtrado[col_patente].astype(str).str.contains(patente, case=False, na=False)]

            if orden_trabajo and col_orden_trabajo in df_filtrado.columns:
                 df_filtrado = df_filtrado[df_filtrado[col_orden_trabajo].astype(str).str.contains(orden_trabajo, case=False, na=False)]


            st.markdown("### 📊 Datos filtrados")
            st.dataframe(df_filtrado, use_container_width=True)


            # --- Ranking Técnicos más Auditados ---
            st.markdown("---")
            st.markdown("### 🏆 Ranking Técnicos más Auditados (Finalizadas)")

            columnas_ranking_tecnicos = [col_tec_nombre, col_empresa, col_fecha, 'Estado de Auditoria']
            if all(col in data_finalizadas.columns for col in columnas_ranking_tecnicos) and col_fecha in data_finalizadas.columns: # Usar data_finalizadas
                 # Aseguramos que data_finalizadas tiene fecha válida
                 data_finalizadas_ranking = data_finalizadas[data_finalizadas[col_fecha].notna()].copy()

                 if not data_finalizadas_ranking.empty:
                      # Selección de rango de fechas
                      fecha_min_ranking = data_finalizadas_ranking[col_fecha].min().date()
                      fecha_max_ranking = data_finalizadas_ranking[col_fecha].max().date()

                      fechas = st.date_input(
                          "📅 Selecciona el rango de fechas (opcional)",
                          value=[fecha_min_ranking, fecha_max_ranking],
                          min_value=fecha_min_ranking,
                          max_value=fecha_max_ranking,
                          key="rango_fechas_ranking_tecnicos"
                      )

                      # Si se seleccionaron fechas válidas (lista de 2 elementos)
                      if isinstance(fechas, list) and len(fechas) == 2:
                          fecha_inicio, fecha_fin = fechas
                          # Asegurarse que las fechas seleccionadas son date objects
                          if isinstance(fecha_inicio, date) and isinstance(fecha_fin, date):
                              # Filtro por rango de fecha
                              mask = (data_finalizadas_ranking[col_fecha] >= pd.to_datetime(fecha_inicio)) & (data_finalizadas_ranking[col_fecha] <= pd.to_datetime(fecha_fin))
                              data_finalizadas_ranking_filtrado = data_finalizadas_ranking.loc[mask].copy()
                          else:
                              data_finalizadas_ranking_filtrado = data_finalizadas_ranking.copy()
                              st.warning("Rango de fechas seleccionado inválido.")

                      else:
                          data_finalizadas_ranking_filtrado = data_finalizadas_ranking.copy()


                      if not data_finalizadas_ranking_filtrado.empty:
                           # Agrupar por Técnico y Empresa
                           ranking = (
                               data_finalizadas_ranking_filtrado
                               .groupby([col_tec_nombre, col_empresa])
                               .agg(
                                    Cantidad_de_Auditorias=(col_fecha, 'size'), # Count non-null dates in the group
                                    Fechas_de_Auditoria=(col_fecha, lambda x: ', '.join(sorted(x.dt.strftime('%d/%m/%Y').tolist())) if pd.api.types.is_datetime64_any_dtype(x) else 'Fechas no válidas')
                                )
                               .reset_index()
                               .rename(columns={
                                   col_tec_nombre: "Técnico",
                                   col_empresa: "Empresa",
                                   "Cantidad_de_Auditorias": "Cantidad de Auditorías",
                                   "Fechas_de_Auditoria": "Fechas de Auditoría"
                               })
                               .sort_values(by="Cantidad de Auditorías", ascending=False)
                           )
                           st.dataframe(ranking, use_container_width=True)
                      else:
                           st.info("⚠️ No hay auditorías finalizadas con fecha válida en el rango de fechas seleccionado.")

                 else:
                     st.info(f"⚠️ No hay auditorías marcadas como '{'finalizada'}' con fecha válida en el archivo para calcular el ranking de técnicos.")

            else:
                 st.error(f"Faltan una o más columnas necesarias para calcular el Ranking de Técnicos más Auditados: {', '.join(columnas_ranking_tecnicos)}")


            # --- KPI Auditorías por Empresa (Técnicos) ---
            st.markdown("---")
            st.subheader("🏢 Auditorías Finalizadas por Empresa (Técnicos)")

            columnas_necesarias_empresa = [col_empresa, 'Estado de Auditoria']
            if all(col in data.columns for col in columnas_necesarias_empresa):

                 if not data_finalizadas.empty:
                      auditorias_empresa = (
                          data_finalizadas[col_empresa]
                          .value_counts()
                          .rename_axis(col_empresa)
                          .reset_index(name='Cantidad de Auditorías Finalizadas')
                      )
                      auditorias_empresa = auditorias_empresa[auditorias_empresa[col_empresa].str.strip() != '']


                      st.dataframe(auditorias_empresa, use_container_width=True)

                      # Gráfico de barras interactivo con Plotly
                      st.markdown("### 📈 Gráfico Auditorías Finalizadas por Empresa")
                      if not auditorias_empresa.empty:
                           fig = px.bar(
                               auditorias_empresa,
                               x='Cantidad de Auditorías Finalizadas',
                               y=col_empresa,
                               orientation='h',
                               color=col_empresa,
                               text='Cantidad de Auditorías Finalizadas',
                               color_discrete_sequence=px.colors.qualitative.Vivid
                           )
                           fig.update_layout(
                               xaxis_title="Cantidad de Auditorías Finalizadas",
                               yaxis_title=col_empresa,
                               yaxis=dict(autorange="reversed"),
                               plot_bgcolor='white'
                           )
                           st.plotly_chart(fig, use_container_width=True)
                      else:
                           st.info("No hay datos de auditorías finalizadas por empresa para mostrar el gráfico.")

                 else:
                      st.warning(f"⚠️ No hay auditorías marcadas como '{'finalizada'}' en el archivo para mostrar auditorías por empresa.")
            else:
                 st.error(f"⚠️ Faltan columnas necesarias para calcular auditorías por empresa: {', '.join(columnas_necesarias_empresa)}")


            # --- KPI Stock Crítico de Herramientas ---
            st.markdown("---")
            st.markdown("### 🔧 Técnicos con Stock Crítico de Herramientas")

            herramientas_criticas = [ # Tu lista de herramientas críticas
                "Power meter GPON", "VFL Luz visible para localizar fallas", "Limpiador de conectores tipo “One Click”",
                "Deschaquetador de primera cubierta para DROP", "Deschaquetador de recubrimiento de FO 125micras Tipo Miller",
                "Cortadora de precisión 3 pasos", "Regla de corte", "Alcohol isopropilico 99%",
                "Paños secos para FO", "Crimper para cable UTP", "Deschaquetador para cables con cubierta redonda (UTP, RG6 )",
                "Tester para cable UTP"
            ]

            herramientas_criticas_existentes = [h for h in herramientas_criticas if h in data.columns]
            columnas_stock_herramientas = [col_tec_nombre, col_empresa, col_fecha, 'Estado de Auditoria'] + herramientas_criticas_existentes

            if all(col in data.columns for col in columnas_stock_herramientas[:4]) and herramientas_criticas_existentes:
                 data_finalizadas_stock_herr = data[(data['Estado de Auditoria'] == 'finalizada') & (data[col_fecha].notna())].copy()

                 if not data_finalizadas_stock_herr.empty:
                      idx_ultima_auditoria = data_finalizadas_stock_herr.groupby(col_tec_nombre)[col_fecha].idxmax()
                      data_ultima_auditoria_herr = data_finalizadas_stock_herr.loc[idx_ultima_auditoria].reset_index(drop=True)

                      def obtener_herramientas_faltantes(row):
                           faltantes = []
                           for herramienta in herramientas_criticas_existentes:
                               valor = row.get(herramienta)
                               if pd.isna(valor) or str(valor).strip().lower() in ["no", "falta", "0"]:
                                   faltantes.append(herramienta)
                           return faltantes

                      columnas_para_stock_herr_procesar = [col_tec_nombre, col_empresa, col_fecha] + herramientas_criticas_existentes
                      stock_critico_herramientas = data_ultima_auditoria_herr[columnas_para_stock_herr_procesar].copy()

                      stock_critico_herramientas["Herramientas Faltantes"] = stock_critico_herramientas.apply(obtener_herramientas_faltantes, axis=1)
                      stock_critico_herramientas = stock_critico_herramientas[stock_critico_herramientas["Herramientas Faltantes"].map(len) > 0]

                      stock_critico_herramientas["Cantidad Faltantes"] = stock_critico_herramientas["Herramientas Faltantes"].map(len)
                      stock_critico_herramientas = stock_critico_herramientas.sort_values(by="Cantidad Faltantes", ascending=False)
                      stock_critico_herramientas = stock_critico_herramientas.rename(columns={col_tec_nombre: "Técnico"})

                      def agregar_icono_herramientas(row):
                           if row["Cantidad Faltantes"] >= 2: return f"🔴 {row['Técnico']}"
                           elif row["Cantidad Faltantes"] == 1: return f"🟡 {row['Técnico']}"
                           else: return row['Técnico']

                      stock_critico_herramientas["Técnico Con Icono"] = stock_critico_herramientas.apply(agregar_icono_herramientas, axis=1)
                      stock_critico_herramientas["Herramientas Faltantes"] = stock_critico_herramientas["Herramientas Faltantes"].apply(lambda x: ", ".join(x))

                      total_tecnicos_stock_critico_herramientas = stock_critico_herramientas.shape[0]
                      st.markdown(f"**🔥 Total técnicos con stock crítico de herramientas: {total_tecnicos_stock_critico_herramientas}**")

                      empresas_disponibles_herr_tabla = stock_critico_herramientas[col_empresa].unique().tolist()
                      empresas_disponibles_herr_tabla = [e for e in empresas_disponibles_herr_tabla if e.strip() != '' and e.lower() != 'nan']
                      empresa_seleccionada_herr_tabla = st.selectbox("🔎 Filtrar por Empresa:", options=["Todas"] + empresas_disponibles_herr_tabla, key="filtro_empresa_stock_herr_tabla")

                      stock_critico_herramientas_general = stock_critico_herramientas.copy()

                      if empresa_seleccionada_herr_tabla != "Todas":
                           stock_critico_herramientas = stock_critico_herramientas[stock_critico_herramientas[col_empresa] == empresa_seleccionada_herr_tabla]

                      st.dataframe(
                          stock_critico_herramientas[["Técnico Con Icono", col_empresa, col_fecha, "Herramientas Faltantes"]],
                          use_container_width=True
                      )

                      buffer_herramientas = io.BytesIO()
                      with pd.ExcelWriter(buffer_herramientas, engine='xlsxwriter') as writer:
                           stock_critico_herramientas[["Técnico Con Icono", col_empresa, col_fecha, "Herramientas Faltantes"]].rename(columns={"Técnico Con Icono": "Técnico"}).to_excel(writer, index=False, sheet_name='Stock_Critico_Herramientas')
                      buffer_herramientas.seek(0)

                      st.download_button(
                          label="📥 Descargar Técnicos con Stock Crítico Herramientas (Tabla Filtrada)",
                          data=buffer_herramientas,
                          file_name="tecnicos_stock_critico_herramientas.xlsx",
                          mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                      )

                      st.markdown("---")
                      st.subheader("📈 Técnicos con Stock Crítico de Herramientas por Empresa")

                      if not stock_critico_herramientas_general.empty:
                           empresas_stock_critico_herramientas = (
                               stock_critico_herramientas_general.groupby(col_empresa)
                               .size()
                               .reset_index(name='Cantidad de Técnicos con Stock Crítico Herramientas')
                               .sort_values(by='Cantidad de Técnicos con Stock Crítico Herramientas', ascending=False)
                           )
                           empresas_stock_critico_herramientas = empresas_stock_critico_herramientas[empresas_stock_critico_herramientas[col_empresa].str.strip() != '']

                           if not empresas_stock_critico_herramientas.empty:
                                fig_stock_herramientas = px.bar(
                                    empresas_stock_critico_herramientas,
                                    x='Cantidad de Técnicos con Stock Crítico Herramientas',
                                    y=col_empresa,
                                    orientation='h',
                                    color=col_empresa,
                                    text='Cantidad de Técnicos con Stock Crítico Herramientas',
                                    color_discrete_sequence=px.colors.qualitative.Vivid
                                )
                                fig_stock_herramientas.update_layout(
                                    xaxis_title="Cantidad de Técnicos con Stock Crítico de Herramientas",
                                    yaxis_title=col_empresa,
                                    yaxis=dict(autorange="reversed"),
                                    plot_bgcolor='white'
                                )
                                st.plotly_chart(fig_stock_herramientas, use_container_width=True)
                           else:
                                st.info("No hay datos suficientes para el gráfico de stock crítico de herramientas por empresa.")
                      else:
                           st.info("No hay técnicos con stock crítico de herramientas para mostrar el gráfico.")


                 else:
                      st.info(f"⚠️ No hay auditorías finalizadas con fecha válida en el archivo para calcular el Stock Crítico de Herramientas.")

            else:
                 st.error(f"⚠️ Faltan columnas necesarias para calcular el Stock Crítico de Herramientas. Asegúrate de incluir {', '.join(columnas_stock_herramientas[:4])} y al menos una de las herramientas críticas: {', '.join(herramientas_criticas)}")


            # --- KPI Stock Crítico de EPP ---
            st.markdown("---")
            st.markdown("### 🦺 Técnicos con Stock Crítico de EPP")

            epp_criticos = [ # Tu lista de EPP críticos
                "Conos de seguridad", "Refugio de PVC", "Casco de Altura", "Barbiquejo",
                "Legionario Para Casco", "Guantes Cabritilla", "Guantes Dielectricos",
                "Guantes trabajo Fino", "Zapatos de Seguridad Dielectricos",
                "LENTE DE SEGURIDAD (CLAROS Y OSCUROS)", "Arnes Dielectrico",
                "Estrobo Dielectrico", "Cuerda de vida /Dielectrico", "Chaleco reflectante",
                "DETECTOR DE TENSION TIPO LAPIZ CON LINTERNA", "Bloqueador Solar"
            ]
            epp_vitales = ["Casco de Altura", "Zapatos de Seguridad Dielectricos", "Arnes Dielectrico", "Estrobo Dielectrico"] # Tu lista EPP vitales


            epp_criticos_existentes = [e for e in epp_criticos if e in data.columns]
            columnas_stock_epp = [col_tec_nombre, col_empresa, col_fecha, 'Estado de Auditoria'] + epp_criticos_existentes

            if all(col in data.columns for col in columnas_stock_epp[:4]) and epp_criticos_existentes:
                 data_finalizadas_stock_epp = data[(data['Estado de Auditoria'] == 'finalizada') & (data[col_fecha].notna())].copy()

                 if not data_finalizadas_stock_epp.empty:
                      idx_ultima_auditoria_epp = data_finalizadas_stock_epp.groupby(col_tec_nombre)[col_fecha].idxmax()
                      data_ultima_auditoria_epp = data_finalizadas_stock_epp.loc[idx_ultima_auditoria_epp].reset_index(drop=True)

                      def obtener_epp_faltantes(row):
                           faltantes = []
                           for epp in epp_criticos_existentes:
                               valor = row.get(epp)
                               if pd.isna(valor) or str(valor).strip().lower() in ["no", "falta", "0"]:
                                   faltantes.append(epp)
                           return faltantes

                      columnas_para_stock_epp_procesar = [col_tec_nombre, col_empresa, col_fecha] + epp_criticos_existentes
                      stock_critico_epp = data_ultima_auditoria_epp[columnas_para_stock_epp_procesar].copy()

                      stock_critico_epp["EPP Faltantes"] = stock_critico_epp.apply(obtener_epp_faltantes, axis=1)
                      stock_critico_epp = stock_critico_epp[stock_critico_epp["EPP Faltantes"].map(len) > 0]

                      stock_critico_epp["Cantidad Faltantes"] = stock_critico_epp["EPP Faltantes"].map(len)
                      stock_critico_epp = stock_critico_epp.sort_values(by="Cantidad Faltantes", ascending=False)
                      stock_critico_epp = stock_critico_epp.rename(columns={col_tec_nombre: "Técnico"})

                      def agregar_icono_epp(row):
                           faltantes_vitales = [epp for epp in row["EPP Faltantes"] if epp in epp_vitales]
                           if len(faltantes_vitales) >= 2: return f"🔴 {row['Técnico']}"
                           elif len(faltantes_vitales) == 1: return f"🟡 {row['Técnico']}"
                           else: return row['Técnico']

                      stock_critico_epp["Técnico Con Icono"] = stock_critico_epp.apply(agregar_icono_epp, axis=1)
                      stock_critico_epp["EPP Faltantes"] = stock_critico_epp["EPP Faltantes"].apply(lambda x: ", ".join(x))

                      total_tecnicos_stock_critico_epp = stock_critico_epp.shape[0]
                      st.markdown(f"**🔥 Total técnicos con stock crítico de EPP: {total_tecnicos_stock_critico_epp}**")

                      empresas_disponibles_epp_tabla = stock_critico_epp[col_empresa].unique().tolist()
                      empresas_disponibles_epp_tabla = [e for e in empresas_disponibles_epp_tabla if e.strip() != '' and e.lower() != 'nan']
                      empresa_seleccionada_epp_tabla = st.selectbox("🔎 Filtrar por Empresa:", options=["Todas"] + empresas_disponibles_epp_tabla, key="filtro_empresa_stock_epp_tabla")

                      stock_critico_epp_general = stock_critico_epp.copy()

                      if empresa_seleccionada_epp_tabla != "Todas":
                           stock_critico_epp = stock_critico_epp[stock_critico_epp[col_empresa] == empresa_seleccionada_epp_tabla]


                      st.dataframe(
                          stock_critico_epp[["Técnico Con Icono", col_empresa, col_fecha, "EPP Faltantes"]],
                          use_container_width=True
                      )

                      buffer_epp = io.BytesIO()
                      with pd.ExcelWriter(buffer_epp, engine='xlsxwriter') as writer:
                           stock_critico_epp[["Técnico Con Icono", col_empresa, col_fecha, "EPP Faltantes"]].rename(columns={"Técnico Con Icono": "Técnico"}).to_excel(writer, index=False, sheet_name='Stock_Critico_EPP')
                      buffer_epp.seek(0)

                      st.download_button(
                          label="📥 Descargar Técnicos con Stock Crítico EPP (Tabla Filtrada)",
                          data=buffer_epp,
                          file_name="tecnicos_stock_critico_epp.xlsx",
                          mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                      )

                      st.markdown("---")
                      st.subheader("📈 Técnicos con Stock Crítico de EPP por Empresa")

                      if not stock_critico_epp_general.empty:
                           empresas_stock_critico_epp = (
                               stock_critico_epp_general.groupby(col_empresa)
                               .size()
                               .reset_index(name='Cantidad de Técnicos con Stock Crítico EPP')
                               .sort_values(by='Cantidad de Técnicos con Stock Crítico EPP', ascending=False)
                           )
                           empresas_stock_critico_epp = empresas_stock_critico_epp[empresas_stock_critico_epp[col_empresa].str.strip() != '']

                           if not empresas_stock_critico_epp.empty:
                                fig_stock_epp = px.bar(
                                    empresas_stock_critico_epp,
                                    x='Cantidad de Técnicos con Stock Crítico EPP',
                                    y=col_empresa,
                                    orientation='h',
                                    color=col_empresa,
                                    text='Cantidad de Técnicos con Stock Crítico EPP',
                                    color_discrete_sequence=px.colors.qualitative.Vivid
                                )
                                fig_stock_epp.update_layout(
                                    xaxis_title="Cantidad de Técnicos con Stock Crítico de EPP",
                                    yaxis_title=col_empresa,
                                    yaxis=dict(autorange="reversed"),
                                    plot_bgcolor='white'
                                )
                                st.plotly_chart(fig_stock_epp, use_container_width=True)
                           else:
                                st.info("No hay datos suficientes para el gráfico de stock crítico de EPP por empresa.")
                      else:
                           st.info("No hay técnicos con stock crítico de EPP para mostrar el gráfico.")


                 else:
                     st.info(f"⚠️ No hay auditorías finalizadas con fecha válida en el archivo para calcular el Stock Crítico de EPP.")

            else:
                 st.error(f"⚠️ Faltan columnas necesarias para calcular el Stock Crítico de EPP. Asegúrate de incluir {', '.join(columnas_stock_epp[:4])} y al menos uno de los EPP críticos: {', '.join(epp_criticos)}")


            # --- Resumen General de Stock Crítico ---
            st.markdown("---")
            st.subheader("📊 Resumen General de Stock Crítico")
            # Calcular totales usando los DataFrames _general antes de filtrar por tabla
            if 'stock_critico_epp_general' in locals(): # Verifica si la variable fue creada
                 total_tecnicos_stock_critico_epp = stock_critico_epp_general.shape[0]
            else:
                 total_tecnicos_stock_critico_epp = 0

            if 'stock_critico_herramientas_general' in locals():
                 total_tecnicos_stock_critico_herramientas = stock_critico_herramientas_general.shape[0]
            else:
                 total_tecnicos_stock_critico_herramientas = 0

            st.metric(label="🔥 Total Técnicos con EPP Crítico", value=total_tecnicos_stock_critico_epp)
            st.metric(label="🔧 Total Técnicos con Herramientas Críticas", value=total_tecnicos_stock_critico_herramientas)

            if archivo:
                # Llamamos a la función de KPIs
                kpis, empresa_kpis_df, total_auditorias, data = process_data(archivo)


        # --- Contenido de la Pestaña 2 ---
        with tab2:
            st.header("🛠️ Información de Auditores")

            # Nombres de columnas clave para Auditores (ya definidos arriba)
            col_auditor = 'Información del Auditor'
            col_estado = 'Estado de Auditoria' # Ya normalizada
            col_fecha = 'Fecha' # Ya datetime
            col_id_trabajo = 'Número de Orden de Trabajo/ ID externo'
            col_empresa = 'Empresa' # Ya string
            col_region = 'Region' # Usado en esta pestaña


            # --- SECCIÓN: Ranking de Auditores por Trabajos Realizados (FINALIZADAS) ---
            st.markdown("### Ranking de Auditores por Trabajos Realizados (Finalizadas)") # Título ajustado

            # Verificar que las columnas necesarias existen en data_finalizadas
            columnas_ranking_auditores = [col_auditor, col_estado] # Estado ya usado para crear data_finalizadas
            if col_auditor in data_finalizadas.columns:

                 if not data_finalizadas.empty: # data_finalizadas ya filtrada y con Auditor normalizado
                      # Agrupar por auditor y contar las auditorías finalizadas
                      ranking_auditores = (
                          data_finalizadas.groupby(col_auditor) # Auditor ya normalizado
                          .size()
                          .reset_index(name="Cantidad de Auditorías Finalizadas") # Renombrado
                          .rename(columns={col_auditor: "Auditor"})
                          .sort_values(by="Cantidad de Auditorías Finalizadas", ascending=False)
                      )
                      st.dataframe(ranking_auditores, use_container_width=True)
                 else:
                      st.info(f"No hay auditorías marcadas como '{'finalizada'}' en el archivo para calcular el ranking de auditores.")

            else:
                 st.error(f"Falta la columna necesaria para el Ranking de Auditores por Trabajos Realizados: {col_auditor}")


            # --- NUEVA SECCIÓN: Conteo de Auditorías por Auditor por Día (Todas con ID válido) ---
            st.markdown("---") # Separador
            st.subheader("🗓️ Auditorías por Auditor por Día ") # Título ajustado

            # Validar y preparar datos para este cálculo específico
            # Necesitamos al menos Fecha válida, Auditor válido e ID de trabajo válido
            # Usamos el DataFrame 'data' que ya tiene la Fecha convertida (con NaT) y Auditor normalizado
            data_para_conteo_diario = data.dropna(subset=[col_fecha, col_auditor, col_id_trabajo]).copy()

            if not data_para_conteo_diario.empty:
                 # Asegurarnos que la columna Fecha es datetime
                 if pd.api.types.is_datetime64_any_dtype(data_para_conteo_diario[col_fecha]):

                     # --- 1. Calcular el conteo por día y auditor ---
                     conteo_auditorias_diario = data_para_conteo_diario.groupby([
                         data_para_conteo_diario[col_fecha].dt.date, # Agrupar solo por la fecha (el día)
                         data_para_conteo_diario[col_auditor]        # Agrupar por el auditor (ya normalizado)
                     ])[col_id_trabajo].nunique().reset_index() # nunique() cuenta valores únicos por grupo

                     # Renombrar las columnas resultantes
                     conteo_auditorias_diario.columns = ['Fecha', 'Auditor', 'Total_Auditorias']

                     # Ordenar el resultado por fecha y auditor (opcional)
                     conteo_auditorias_diario = conteo_auditorias_diario.sort_values(by=['Fecha', 'Auditor'])


                     # --- 2. Agregar el filtro por fecha específica ---
                     st.markdown("---")
                     st.subheader("🔍 Filtro por Día Específico para el Conteo Diario")

                     # Obtener el rango de fechas disponibles en los datos calculados del conteo diario
                     try:
                         min_date_diario = conteo_auditorias_diario['Fecha'].min()
                         max_date_diario = conteo_auditorias_diario['Fecha'].max()
                         # Manejar NaT si existen
                         if pd.isna(max_date_diario): max_date_diario = date.today()
                         if pd.isna(min_date_diario): min_date_diario = date.today()

                         # Valor por defecto para el selector: el último día con datos o el único día
                         default_date_diario = max_date_diario if min_date_diario != max_date_diario else min_date_diario
                         # Asegurarse de que default_date es un objeto date
                         if isinstance(default_date_diario, pd.Timestamp):
                              default_date_diario = default_date_diario.date()

                     except Exception:
                         # En caso de error o si conteo_auditorias_diario está vacío
                         min_date_diario = date.today()
                         max_date_diario = date.today()
                         default_date_diario = date.today()


                     # Widget st.date_input para seleccionar una fecha
                     fecha_seleccionada_filtro_diario = st.date_input(
                         "Selecciona una fecha para ver el conteo:",
                         value=default_date_diario, # Establece el valor inicial
                         min_value=min_date_diario, # Define la fecha mínima seleccionable
                         max_value=max_date_diario, # Define la fecha máxima seleccionable
                         key="filtro_conteo_fecha_input_auditor_diario" # Añadir una key única globalmente
                     )

                     # --- 3. Aplicar el filtro de fecha ---
                     resultados_filtrados_diario = pd.DataFrame() # Inicializar

                     if fecha_seleccionada_filtro_diario: # Si se seleccionó una fecha
                         # Convertir a objeto date
                         fecha_a_comparar_diario = fecha_seleccionada_filtro_diario

                         # Filtrar el DataFrame de conteo diario por la fecha seleccionada
                         resultados_filtrados_diario = conteo_auditorias_diario[
                             conteo_auditorias_diario['Fecha'] == fecha_a_comparar_diario
                         ].copy()


                         # --- 4. Mostrar los resultados filtrados en una tabla ---
                         st.markdown(f"### Resultados para la fecha: **{fecha_seleccionada_filtro_diario.strftime('%d/%m/%Y')}**")

                         if not resultados_filtrados_diario.empty:
                             # Mostrar la tabla con el conteo por auditor para el día seleccionado
                             st.dataframe(resultados_filtrados_diario, use_container_width=True)
                         else:
                             st.info(f"ℹ️ No se registraron auditorías (con ID válido) para ningún auditor en la fecha seleccionada (**{fecha_seleccionada_filtro_diario.strftime('%d/%m/%Y')}**).")

                     else:
                          st.warning("⚠️ Por favor, selecciona una fecha en el filtro para visualizar los resultados del conteo diario.")

                 else:
                      st.error("Error interno: La columna de fecha no es de tipo datetime después de la conversión inicial. Revisa el formato de fecha en tu archivo Excel.")


            else:
                 # Mensaje si no hay datos suficientes DESPUÉS de limpiar para este cálculo
                 st.warning(f"⚠️ El archivo Excel cargado no contiene suficientes filas con información válida ({col_fecha}, {col_auditor}, {col_id_trabajo}) para calcular el conteo de auditorías por auditor por día.")


            # --- KPI Distribución de Auditorías entre Empresas con Fechas (Usar data_finalizadas) ---
            st.markdown("---") # Separador
            st.markdown("### Distribución de Auditorías Finalizadas entre Empresas con Fechas")

            columnas_necesarias_distribucion = [col_auditor, col_empresa, col_fecha]
            if all(col in data_finalizadas.columns for col in columnas_necesarias_distribucion) and col_fecha in data_finalizadas.columns:
                 # Asegurarse que 'Fecha' en data_finalizadas es datetime
                 if pd.api.types.is_datetime64_any_dtype(data_finalizadas[col_fecha]):
                      distribucion_auditorias = data_finalizadas.groupby([col_auditor, col_empresa]).agg(
                          Cantidad_de_Auditorias=(col_fecha, 'size'),
                          Fechas_de_Auditoria=(col_fecha, lambda x: ', '.join(sorted(x.dt.strftime('%d/%m/%Y').tolist())) if pd.api.types.is_datetime64_any_dtype(x) else 'Fechas no válidas')
                      ).reset_index()

                      if not distribucion_auditorias.empty:
                           st.dataframe(distribucion_auditorias, use_container_width=True)
                      else:
                           st.info("No hay datos suficientes para la distribución de auditorías finalizadas por auditor y empresa.")
                 else:
                      st.error(f"Error: La columna '{col_fecha}' en las auditorías finalizadas no es de tipo datetime.")

            else:
                 st.error(f"Faltan columnas necesarias para calcular la distribución de auditorías: {', '.join(columnas_necesarias_distribucion)}")


            # --- KPI: Auditorías por Región (Usar data_finalizadas) ---
            st.markdown("---") # Separador
            st.subheader("🌎 Auditorías Finalizadas por Región")

            col_region = 'Region' # Asegurarse que este nombre es correcto

            # Verificar columna necesaria
            if col_region in data_finalizadas.columns:
                 # Asegurarse que la columna de región no tiene valores vacíos/NaN para agrupar
                 data_finalizadas_region = data_finalizadas.dropna(subset=[col_region]).copy()

                 if not data_finalizadas_region.empty:
                      # Agrupar datos por Región y contar cantidad de auditorías finalizadas
                      auditorias_por_region = (
                          data_finalizadas_region.groupby(col_region)
                          .size()
                          .reset_index(name='Cantidad de Auditorías Finalizadas')
                          .sort_values(by='Cantidad de Auditorías Finalizadas', ascending=False)
                      )
                      auditorias_por_region = auditorias_por_region[auditorias_por_region[col_region].str.strip() != '']


                      if not auditorias_por_region.empty:
                           fig_auditorias_region = px.bar(
                               auditorias_por_region,
                               x='Cantidad de Auditorías Finalizadas',
                               y=col_region,
                               orientation='h',
                               color=col_region,
                               text='Cantidad de Auditorías Finalizadas',
                               color_discrete_sequence=px.colors.qualitative.Set2
                           )
                           fig_auditorias_region.update_layout(
                               xaxis_title="Cantidad de Auditorías Finalizadas",
                               yaxis_title=col_region,
                               yaxis=dict(autorange="reversed"),
                               plot_bgcolor='white'
                           )
                           st.plotly_chart(fig_auditorias_region, use_container_width=True)
                      else:
                           st.info(f"No hay auditorías finalizadas con información de '{col_region}'.")

                 else:
                      st.info(f"No hay auditorías finalizadas con información de '{col_region}'.")
            else:
                 st.error(f"Falta la columna '{col_region}' para calcular las auditorías por región.")


            # Calcular el total de auditorías finalizadas (ya lo tenías)
            st.markdown("---")
            if 'Estado de Auditoria' in data.columns:
                 total_auditorias_finalizadas = len(data[data['Estado de Auditoria'] == 'finalizada'])
                 st.markdown(f"""
                      <div style="background-color: #f0f0f5; padding: 15px 25px; border-radius: 8px; font-size: 24px; font-weight: bold; color: #333;">
                          <span style="color: #007bff;">Total de Auditorías Finalizadas en el archivo: </span><span style="color: #28a745;">{total_auditorias_finalizadas}</span>
                      </div>
                  """, unsafe_allow_html=True)
            else:
                 st.error("Falta la columna 'Estado de Auditoria' para calcular el total de auditorías finalizadas.")


            # ----------------- KPI Ranking de Auditores por información completa (Usar data_finalizadas) -----------------
            st.markdown("---")
            st.subheader("📋 Ranking de Auditores por Información Completa")

            columnas_completitud = [col_auditor, col_estado] # Estado ya usado

            if col_auditor in data_finalizadas.columns: # Verificar si el auditor existe en data_finalizadas

                 if not data_finalizadas.empty:
                      data_finalizadas_completitud = data_finalizadas.copy()
                      total_columnas = data_finalizadas_completitud.shape[1]
                      if total_columnas > 0:
                           data_finalizadas_completitud["% Completitud"] = data_finalizadas_completitud.notna().sum(axis=1) / total_columnas * 100

                           ranking_completitud = data_finalizadas_completitud.groupby(col_auditor)["% Completitud"].mean().reset_index()
                           ranking_completitud = ranking_completitud.sort_values(by="% Completitud", ascending=False)

                           def formato_porcentaje(valor):
                                if pd.isna(valor): return ""
                                return f"{valor:,.1f}%".replace('.', ',')

                           def estilo_azul(val):
                               return 'color: blue; font-weight: bold;' if isinstance(val, (int, float)) and not pd.isna(val) else ''

                           st.dataframe(
                               ranking_completitud.style
                               .format({"% Completitud": formato_porcentaje})
                               .map(estilo_azul, subset=["% Completitud"]),
                               use_container_width=True
                           )
                      else:
                           st.warning("No hay columnas en los datos para calcular el porcentaje de completitud.")

                 else:
                      st.info(f"No hay auditorías marcadas como '{'finalizada'}' para calcular el Ranking de Auditores por Información Completa.")

            else:
                 st.error(f"Falta la columna '{col_auditor}' para calcular el Ranking de Auditores por Información Completa.")


    else:
        # Este mensaje se muestra si el DataFrame está vacío después de recuperarlo de session_state
        st.warning("⚠️ El archivo Excel cargado está vacío o no contiene datos procesables después de la limpieza inicial.")


else:
    # Este mensaje se muestra si 'data' NO está en session_state (es decir, nunca se ha cargado un archivo válido)
    st.warning("⚠️ Por favor, sube un archivo Excel con los datos de auditoría para comenzar el análisis.")

# --- Fin del script ---
    



