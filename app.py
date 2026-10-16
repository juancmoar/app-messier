import openpyxl
import pandas as pd
import streamlit as st

# Configuración de la página con estilo visual moderno
st.set_page_config(
    page_title="Catálogo Messier - Astronomía",
    page_icon="🌌",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Estilos CSS personalizados para mejorar el diseño (modo oscuro elegante)
st.markdown(
    """
    <style>
        .main {
            background-color: #0e1117;
        }
        .stMetric {
            background-color: #1a1c23;
            padding: 15px;
            border-radius: 10px;
            border: 1px solid #30363d;
        }
        h1 {
            color: #58a6ff;
            font-family: 'Helvetica Neue', sans-serif;
        }
        h3 {
            color: #c9d1d9;
        }
    </style>
""",
    unsafe_allow_html=True,
)

EXCEL_PATH = "Messier.xlsx"


# Cargar los datos del Excel
@st.cache_data
def load_data():
  df_raw = pd.read_excel(EXCEL_PATH, sheet_name="Messier Objects", header=None)
  df_objects = pd.read_excel(
      EXCEL_PATH, sheet_name="Messier Objects", skiprows=6, header=None
  )
  columns = [
      "M #",
      "NGC",
      "Con",
      "RA",
      "Deg",
      "Sec",
      "Mag",
      "Type",
      "Date",
      "Time",
      "Seeing",
      "Equip",
      "Power",
      "Initials",
      "Observing Notes",
      "Comments",
  ]
  df_objects.columns = columns
  df_objects = df_objects[
      df_objects["M #"].notna() & (df_objects["M #"] != "M #")
  ]
  df_types = pd.read_excel(EXCEL_PATH, sheet_name="Object Types")
  return df_raw, df_objects, df_types


df_raw, df, df_types = load_data()

# Encabezado principal de la aplicación
st.title("🌌 Rastreador del Catálogo Messier")
st.markdown(
    "Gestiona tus observaciones astronómicas de forma visual, limpia y"
    " organizada."
)

# Panel lateral para filtros rápidos
st.sidebar.header("🔍 Filtros de Búsqueda")
search_term = st.sidebar.text_input(
    "Buscar por M#, NGC o Constelación", ""
).upper()

selected_type = st.sidebar.selectbox(
    "Filtrar por Tipo de Objeto", ["Todos"] + df["Type"].dropna().unique().tolist()
)

# Aplicar filtros a los datos
filtered_df = df.copy()
if search_term:
  filtered_df = filtered_df[
      filtered_df["M #"].astype(str).str.contains(search_term)
      | filtered_df["NGC"].astype(str).str.contains(search_term)
      | filtered_df["Con"].astype(str).str.contains(search_term)
  ]
if selected_type != "Todos":
  filtered_df = filtered_df[filtered_df["Type"] == selected_type]

# Métricas del Dashboard con diseño mejorado
total_objects = len(df)
observed_objects = df["Date"].notna().sum()
percentage = (observed_objects / total_objects) * 100

col1, col2, col3 = st.columns(3)
with col1:
  st.metric(
      label="🔭 Total Objetos", value=total_objects, delta="Catálogo completo"
  )
with col2:
  st.metric(
      label="✨ Observados",
      value=observed_objects,
      delta=f"{(observed_objects/total_objects)*100:.0f}% completado",
  )
with col3:
  st.metric(
      label="🎯 Pendientes",
      value=total_objects - observed_objects,
      delta="Por descubrir",
  )

st.divider()

# Sección de la tabla interactiva
st.subheader("📋 Tabla Interactiva de Objetos")
st.markdown(
    "Haz doble clic en cualquier celda de las columnas de observación (**Fecha,"
    " Equip, Notas, etc.**) para editarla directamente:"
)

# Editor de datos optimizado
edited_df = st.data_editor(
    filtered_df,
    use_container_width=True,
    num_rows="fixed",
    key="messier_editor",
    height=450,
)

# Botón flotante/destacado para guardar cambios
st.markdown("<br>", unsafe_allow_html=True)
col_btn1, col_btn2 = st.columns([1, 4])
with col_btn1:
  save_button = st.button(
      "💾 Guardar Cambios", type="primary", use_container_width=True
  )

if save_button:
  try:
    book = openpyxl.load_workbook(EXCEL_PATH)
    with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl") as writer:
      df_raw.iloc[:6, :].to_excel(
          writer, sheet_name="Messier Objects", index=False, header=False
      )
      edited_df.to_excel(
          writer,
          sheet_name="Messier Objects",
          index=False,
          header=False,
          startrow=6,
      )
      for sheet in book.sheetnames:
        if sheet != "Messier Objects":
          temp_df = pd.read_excel(EXCEL_PATH, sheet_name=sheet)
          temp_df.to_excel(writer, sheet_name=sheet, index=False)
    st.success(
        "¡Cambios guardados correctamente en el archivo Excel y sincronizados!"
    )
  except Exception as e:
    st.error(f"Error al guardar los cambios: {e}")