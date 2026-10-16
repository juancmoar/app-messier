import openpyxl
import pandas as pd
import streamlit as st

# Configuración de la página
st.set_page_config(
    page_title="Catálogo Messier - Registro Interactivo",
    page_icon="✨",
    layout="wide",
)

EXCEL_PATH = "Messier.xlsx"


# Función para cargar los datos manteniendo la estructura del Excel
def load_data():
  # Cargamos el archivo completo para no perder las cabeceras y metadatos de las primeras filas
  df_raw = pd.read_excel(EXCEL_PATH, sheet_name="Messier Objects", header=None)
  # Los datos empiezan en la fila 6 (índice 6 en Python)
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
  # Filtrar filas válidas del catálogo Messier (1 al 110)
  df_objects = df_objects[
      df_objects["M #"].notna() & (df_objects["M #"] != "M #")
  ]

  df_types = pd.read_excel(EXCEL_PATH, sheet_name="Object Types")
  return df_raw, df_objects, df_types


df_raw, df, df_types = load_data()

# Título de la aplicación
st.title("🌌 Gestor y Rastreador del Catálogo Messier")
st.markdown(
    "Modifica tus fechas de observación, notas o equipos directamente en la"
    " tabla y haz clic en **Guardar en Excel**."
)

# Métricas principales (Dashboard)
total_objects = len(df)
observed_objects = df["Date"].notna().sum()
percentage = (observed_objects / total_objects) * 100

col1, col2, col3 = st.columns(3)
col1.metric("Total Objetos Messier", total_objects)
col2.metric("Observados", observed_objects)
col3.metric("Progreso Total", f"{percentage:.1f}%")

st.divider()

# Editor interactivo de datos
st.subheader("📝 Tabla de Observación Interactiva")
st.markdown(
    "Puedes editar celdas como *Fecha, Equipamiento, Notas u Observador* en"
    " la siguiente tabla:"
)

# Usamos st.data_editor para permitir la edición directa en pantalla
edited_df = st.data_editor(
    df,
    use_container_width=True,
    num_rows="fixed",
    key="messier_editor",
)

# Botón para guardar los cambios en el archivo Excel original
if st.button("💾 Guardar cambios en el Excel", type="primary"):
  try:
    # Cargamos el libro de trabajo con openpyxl para preservar formato y otras hojas
    book = openpyxl.load_workbook(EXCEL_PATH)
    writer = pd.ExcelWriter(
        EXCEL_PATH, engine="openpyxl", mode="a", if_sheet_exists="overlay"
    )

    # Actualizamos las filas de datos a partir de la fila 7 (índice de Excel 7, que corresponde a skiprows=6 + header)
    # Como pandas escribe sin índice, colocamos los datos limpios en la hoja
    with pd.ExcelWriter(
        EXCEL_PATH, engine="openpyxl", mode="a", if_sheet_exists="replace"
    ) as writer:
      # Para preservar las primeras 6 filas de metadatos, reconstruimos la hoja combinando raw y edited
      # Una forma más limpia y segura en Streamlit para este Excel específico:
      pass

    # Método directo alternativo y seguro para sobrescribir manteniendo metadatos:
    with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl") as writer:
      # Escribimos las 6 primeras filas originales de metadatos/cabecera
      df_raw.iloc[:6, :].to_excel(
          writer, sheet_name="Messier Objects", index=False, header=False
      )
      # Escribimos los datos editados justo debajo
      edited_df.to_excel(
          writer,
          sheet_name="Messier Objects",
          index=False,
          header=False,
          startrow=6,
      )

      # Copiamos el resto de hojas para no perderlas
      for sheet in book.sheetnames:
        if sheet != "Messier Objects":
          temp_df = pd.read_excel(EXCEL_PATH, sheet_name=sheet)
          temp_df.to_excel(writer, sheet_name=sheet, index=False)

    st.success(
        "¡Cambios guardados exitosamente en `Messier.xlsx`! El archivo en tu"
        " ordenador ha sido actualizado."
    )
  except Exception as e:
    st.error(
        f"Hubo un error al guardar el archivo. Asegúrate de cerrarlo si lo"
        f" tienes abierto en Excel. Detalle: {e}"
    )
