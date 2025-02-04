import streamlit as st
import pandas as pd

# Agregar logo y título en la parte superior izquierda
col1, col2 = st.columns([1, 5])

with col1:
    st.image("logo.jpeg", width=120)  # Ajusta el tamaño del logo

with col2:
    st.title("📊 Análisis de Ausencias")

st.write("Sube los siguientes 5 archivos en formato Excel para realizar el análisis:")

# Cargar archivos
archivo_hcm = st.file_uploader("📂 Archivo HCM", type=["xlsx"])
archivo_fraccionadas_ps = st.file_uploader("📂 Archivo Fraccionadas PeopleSoft", type=["xlsx"])
archivo_total_ps = st.file_uploader("📂 Archivo Total PeopleSoft", type=["xlsx"])
archivo_dinero_seg = st.file_uploader("📂 Archivo Dinero Segovia", type=["xlsx"])
archivo_dinero_mar = st.file_uploader("📂 Archivo Dinero Marmato", type=["xlsx"])

if all([archivo_hcm, archivo_fraccionadas_ps, archivo_total_ps, archivo_dinero_seg, archivo_dinero_mar]):
    st.success("✅ ¡Archivos cargados con éxito! Procesando...")

    # Leer archivos
    df_hcm = pd.read_excel(archivo_hcm, skiprows=1, engine="openpyxl")
    df_fraccionadas_ps = pd.read_excel(archivo_fraccionadas_ps, engine="openpyxl")
    df_total_ps = pd.read_excel(archivo_total_ps, engine="openpyxl")
    df_dinero_seg = pd.read_excel(archivo_dinero_seg, skiprows=6, engine="openpyxl")
    df_dinero_mar = pd.read_excel(archivo_dinero_mar, skiprows=6, engine="openpyxl")

    # Ajustar tipos de datos
    df_hcm['START_DATE'] = pd.to_datetime(df_hcm['START_DATE'], format='%d/%m/%Y')
    df_hcm['END_DATE'] = pd.to_datetime(df_hcm['END_DATE'], format='%d/%m/%Y')
    df_hcm['PER_ABSENCE_ENTRY_ID'] = df_hcm['PER_ABSENCE_ENTRY_ID'].astype(str)
    df_hcm['DURATION'] = round(df_hcm['DURATION'], 2)

    df_dinero = pd.concat([df_dinero_seg, df_dinero_mar], ignore_index=True)
    df_dinero['Fecha Inicio Disfrute'] = pd.to_datetime(df_dinero['Fecha Inicio Disfrute'], format='%Y-%m-%d')

    # Identificar duplicados en HCM
    duplicados_hcm = df_hcm[df_hcm.duplicated(subset=["PER_ABSENCE_ENTRY_ID"], keep=False)]
    df_hcm = df_hcm.drop_duplicates(subset=["PER_ABSENCE_ENTRY_ID"], keep="first")

    # Ordenar para detectar solapamientos
    df_hcm = df_hcm.sort_values(by=["PERSON_NUMBER", "START_DATE"])
    solapadas = []
    for person, group in df_hcm.groupby("PERSON_NUMBER"):
        group = group.sort_values(by="START_DATE")
        prev_end = None
        first_overlap = None

        for index, row in group.iterrows():
            if prev_end and row["START_DATE"] <= prev_end:
                if first_overlap is None:
                    solapadas.append(prev_row)
                    first_overlap = prev_row
                solapadas.append(row)

            prev_end = max(prev_end, row["END_DATE"]) if prev_end else row["END_DATE"]
            prev_row = row

    df_solapadas = pd.DataFrame(solapadas)

    # Comparaciones entre bases de datos
    df_no_en_fraccionadas = df_hcm.merge(df_fraccionadas_ps, left_on=["PERSON_NUMBER", "START_DATE"],
                                         right_on=["ID", "Fecha Inicio"], how="left", indicator=True).query('_merge == "left_only"').loc[:, df_hcm.columns]

    df_no_en_fraccionadas_total = df_no_en_fraccionadas.merge(df_total_ps, left_on=["PERSON_NUMBER", "START_DATE"],
                                                               right_on=["ID", "Fecha Inicio Real"], how="left", indicator=True).query('_merge == "left_only"').loc[:, df_no_en_fraccionadas.columns]

    df_no_en_dinero = df_no_en_fraccionadas_total.merge(df_dinero, left_on=["PERSON_NUMBER", "START_DATE"],
                                                         right_on=["Id Empleado", "Fecha Inicio Disfrute"], how="left", indicator=True).query('_merge == "left_only"').loc[:, df_no_en_fraccionadas_total.columns]

    df_si_en_fraccionadas = df_hcm.merge(df_fraccionadas_ps, left_on=["PERSON_NUMBER", "START_DATE"],
                                         right_on=["ID", "Fecha Inicio"], how="inner")

    df_inconsistencias = df_si_en_fraccionadas[(df_si_en_fraccionadas["DURATION"] != df_si_en_fraccionadas["Horas"]) |
                                               (df_si_en_fraccionadas["UOM"] != "Horas")]

    columnas_finales = list(df_hcm.columns) + ["ID", "Nombre", "Fecha Inicio", "Horas", "Usuario", "Instancia", "Instancia.1", "Recepción", "Processed"]
    df_inconsistencias = df_inconsistencias[columnas_finales]

    # Mostrar resultados en la interfaz
    st.subheader("📊 Resultados del Análisis")

    st.write("### 🔍 Duplicados en HCM")
    st.dataframe(duplicados_hcm)

    st.write("### ❌ Ausencias No Integradas")
    st.dataframe(df_no_en_dinero)

    st.write("### ⚠️ Ausencias Solapadas")
    st.dataframe(df_solapadas)

    st.write("### ❗ Errores en Duración")
    st.dataframe(df_inconsistencias)

    # Generar archivo de salida
    st.write("### 📥 Descarga de Resultados")
    archivo_salida = "resultado.xlsx"

    with pd.ExcelWriter(archivo_salida, engine="openpyxl") as writer:
        duplicados_hcm.to_excel(writer, sheet_name="Duplicados HCM", index=False)
        df_no_en_dinero.to_excel(writer, sheet_name="No integradas", index=False)
        df_solapadas.to_excel(writer, sheet_name="Solapadas HCM", index=False)
        df_inconsistencias.to_excel(writer, sheet_name="Error Duracion", index=False)

    with open(archivo_salida, "rb") as file:
        st.download_button(label="📥 Descargar Resultados en Excel",
                           data=file,
                           file_name="resultado.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
