import streamlit as st
import pandas as pd
from dateutil.parser import parse

st.title("Generador de Reporte Excel con Descripción en línea")

uploaded_file = st.file_uploader("Sube tu archivo Excel", type=["xlsx"])

if uploaded_file is not None:
    # Leer Excel sin asumir encabezado aún
    df_raw = pd.read_excel(uploaded_file, sheet_name=0, header=None)

    # Detectar fila de encabezado automáticamente
    header_row = None
    for i, row in df_raw.iterrows():
        row_str = " ".join([str(c).strip().lower() for c in row if pd.notna(c)])
        if "nombres" in row_str and "cédula" in row_str:
            header_row = i
            break

    if header_row is None:
        st.error("No se pudo encontrar la fila de encabezado con 'Nombres' y 'Cédula'.")
        st.stop()

    # Leer nuevamente usando la fila detectada como encabezado
    df = pd.read_excel(uploaded_file, sheet_name=0, header=header_row)

    # Normalizar nombres de columnas
    def limpiar_col(col):
        col = str(col).strip().lower()
        col = col.replace("á","a").replace("é","e").replace("í","i").replace("ó","o").replace("ú","u").replace("ñ","n")
        return col

    df.columns = [limpiar_col(c) for c in df.columns]

    # Verificar columnas necesarias
    columnas_necesarias = ['nombres y apellidos', 'cedula', 'cargo', 'desde', 'hasta']
    faltantes = [col for col in columnas_necesarias if col not in df.columns]
    if faltantes:
        st.error(f"Faltan las columnas: {', '.join(faltantes)}")
        st.stop()

    # Rellenar celdas vacías de Nombres, Cédula y Cargo con el valor de la fila anterior
    df[['nombres y apellidos', 'cedula', 'cargo']] = df[['nombres y apellidos', 'cedula', 'cargo']].fillna(method='ffill')

    # Insertar la nueva columna después de "hasta"
    pos = df.columns.get_loc('hasta') + 1
    df.insert(pos, 'Descripción en línea', "")

    # Función para formatear fechas en español
    def format_fecha_es(fecha):
        if pd.isna(fecha):
            return ""
        try:
            if isinstance(fecha, (pd.Timestamp, pd.DatetimeIndex)):
                fecha_dt = pd.to_datetime(fecha)
            else:
                fecha_str = str(fecha).strip()
                fecha_dt = parse(fecha_str, dayfirst=True, fuzzy=True)
            meses = ["enero","febrero","marzo","abril","mayo","junio","julio",
                     "agosto","septiembre","octubre","noviembre","diciembre"]
            return f"{fecha_dt.day} de {meses[fecha_dt.month - 1]} de {fecha_dt.year}"
        except:
            return str(fecha)

    # Función para asegurar cédula de 10 dígitos
    def format_cedula(cedula):
        if pd.isna(cedula):
            return ""
        try:
            cedula_str = str(int(cedula))  # quitar decimales
        except:
            cedula_str = str(cedula).strip()
        if len(cedula_str) == 9:
            cedula_str = "0" + cedula_str
        return cedula_str

    # Función para generar la línea de texto para la columna
    def generar_linea(row):
        nombre = row['nombres y apellidos']
        cedula = format_cedula(row['cedula'])
        cargo = row['cargo']
        desde = format_fecha_es(row['desde'])
        hasta = format_fecha_es(row['hasta'])

        if pd.notna(nombre) and str(nombre).strip() != "":
            return f"El señor(a) {nombre} con cédula de ciudadanía {cedula}, en su calidad de {cargo} durante el período comprendido entre el {desde} y el {hasta}, incurrió en desviación administrativa por cuanto "
        else:
            return ""

    # Aplicar la función para generar la descripción
    df['Descripción en línea'] = df.apply(generar_linea, axis=1)

    st.success("✅ Reporte generado correctamente con columna 'Descripción en línea'.")
    st.dataframe(df)

    # Guardar archivo actualizado
    output_file = "reporte_actualizado.xlsx"
    df.to_excel(output_file, index=False)

    # Botón de descarga
    with open(output_file, "rb") as f:
        st.download_button(
            label="📥 Descargar Excel con descripción en línea",
            data=f,
            file_name="reporte_actualizado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
