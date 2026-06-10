import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(
    page_title="Analizador de Horas Médicas",
    layout="wide"
)

st.title("Analizador de Horas Asignadas")

archivo = st.file_uploader("Sube archivo Excel", type=["xlsx"])

if archivo:

    # Leer hojas
    hoja1 = pd.read_excel(archivo, sheet_name=0)
    hoja2 = pd.read_excel(archivo, sheet_name=1)
    hoja3 = pd.read_excel(archivo, sheet_name=2)

    # Limpiar columnas
    hoja1.columns = hoja1.columns.str.strip().str.upper()
    hoja2.columns = hoja2.columns.str.strip().str.upper()
    hoja3.columns = hoja3.columns.str.strip().str.upper()

    # Columnas fijas según tu Excel
    col_h1_prof = "NOMBRE PROFESIONAL"
    col_h1_agr = "AGRUPACION"
    col_h1_estado = "ESTADO HORA"

    col_h2_prof = "PROFESIONAL"
    col_h2_esp = "ESPECIALIDAD"

    col_h3_prof = "PROFESIONAL LEY 18"

    # Validaciones básicas
    for col in [col_h1_prof, col_h1_agr, col_h1_estado]:
        if col not in hoja1.columns:
            st.error(f"Falta columna en Hoja 1: {col}")
            st.stop()

    if col_h2_prof not in hoja2.columns or col_h2_esp not in hoja2.columns:
        st.error("Hoja 2 no tiene columnas correctas")
        st.stop()

    if col_h3_prof not in hoja3.columns:
        st.error("Hoja 3 no tiene columna PROFESIONAL LEY 18")
        st.stop()

    # =========================
    # FILTRO PRINCIPAL
    # =========================
    df = hoja1[
        hoja1[col_h1_estado].astype(str).str.upper() == "ASIGNADA"
    ].copy()

    agrupaciones_validas = [
        "MEDICO APS",
        "MEDICO ESPECIALISTA",
        "ODONTOLOGIA APS",
        "ODONTOLOGIA ESPECIALIDADES",
        "QUIMICO FARMACEUTICO",
        "PROCEDIMIENTO"
    ]

    df = df[
        df[col_h1_agr].astype(str).str.upper().isin(agrupaciones_validas)
    ]

    # =========================
    # MAPA ESPECIALIDADES
    # =========================
    especialidades = dict(
        zip(
            hoja2[col_h2_prof].astype(str).str.strip(),
            hoja2[col_h2_esp].astype(str).str.strip()
        )
    )

    no_medicos = set(
        hoja3[col_h3_prof].astype(str).str.strip()
    )

    resultados = []
    desconocidos = []

    # =========================
    # PROCESAMIENTO
    # =========================
    for _, fila in df.iterrows():

        profesional = str(fila[col_h1_prof]).strip()

        if profesional in especialidades:
            resultados.append(especialidades[profesional])

        elif profesional in no_medicos:
            resultados.append(None)

        else:
            desconocidos.append(profesional)
            resultados.append(None)

    df["ESPECIALIDAD_FINAL"] = resultados

    # =========================
    # USUARIO (DESCONOCIDOS)
    # =========================
    st.subheader("Profesionales no encontrados")

    nuevos_medicos = []

    for profesional in sorted(set(desconocidos)):

        st.warning(f"{profesional} no encontrado en bases")

        es_medico = st.radio(
            f"¿{profesional} es médico?",
            ["No", "Sí"],
            key=profesional
        )

        if es_medico == "Sí":

            especialidad = st.text_input(
                f"Especialidad de {profesional}",
                key=f"esp_{profesional}"
            )

            if especialidad:

                nuevos_medicos.append({
                    "PROFESIONAL": profesional,
                    "ESPECIALIDAD": especialidad
                })

                df.loc[
                    df[col_h1_prof] == profesional,
                    "ESPECIALIDAD_FINAL"
                ] = especialidad

    # =========================
    # TABLA 1 (RESUMEN)
    # =========================
    tabla = (
        df.dropna(subset=["ESPECIALIDAD_FINAL"])
        .groupby("ESPECIALIDAD_FINAL")
        .size()
        .reset_index(name="TOTAL")
        .sort_values("TOTAL", ascending=False)
    )

    st.subheader("Resultado por Especialidad")
    st.dataframe(tabla, use_container_width=True)

    # =========================
    # TABLA 2 (CORREGIDA)
    # SOLO MÉDICOS (HOJA 2)
    # =========================

    df_medicos = df[df[col_h1_prof].isin(hoja2[col_h2_prof])].copy()

    df_medicos["OMISIONES"] = 1

    tabla2 = df_medicos.copy()

    # Renombrar columnas base
    tabla2.rename(columns={
        col_h1_prof: "NOMBRE PROFESIONAL",
        "ESPECIALIDAD_FINAL": "ESPECIALIDAD",
        "RUT PACIENTE": "RUT PACIENTE",
        "NOMBRE PACIENTE": "NOMBRE PACIENTE",
        "FECHA": "FECHA"
    }, inplace=True)

    # Asegurar columnas existentes (por si no vienen en Excel)
    for col in ["RUT PACIENTE", "NOMBRE PACIENTE", "FECHA"]:
        if col not in tabla2.columns:
            tabla2[col] = None
    
    # Orden exacto solicitado
    tabla2 = tabla2[[
        "ESPECIALIDAD",
        "NOMBRE PROFESIONAL",
        "RUT PACIENTE",
        "NOMBRE PACIENTE",
        "FECHA",
        "OMISIONES"
    ]]
    # =========================
    # TABLA 3 (RESUMEN POR MÉDICO)
    # =========================

    tabla3 = (
        df_medicos.groupby(["ESPECIALIDAD_FINAL", col_h1_prof])
        .size()
        .reset_index(name="OMISIONES")
    )

    tabla3.rename(columns={
        "ESPECIALIDAD_FINAL": "ESPECIALIDAD",
        col_h1_prof: "NOMBRE PROFESIONAL"
    }, inplace=True)

    tabla3 = tabla3[[
        "ESPECIALIDAD",
        "NOMBRE PROFESIONAL",
        "OMISIONES"
    ]]

    st.subheader("Tabla 3 - Resumen por Profesional")
    st.dataframe(tabla3, use_container_width=True)
    # =========================
    # EXCLUIDOS LEY 18
    # =========================
    if no_medicos:
        st.subheader("Excluidos (Ley 18)")
        st.dataframe(pd.DataFrame({
            "PROFESIONAL": sorted(set(no_medicos))
        }))

    # =========================
    # EXPORT EXCEL
    # =========================
    salida = BytesIO()

    with pd.ExcelWriter(salida, engine="xlsxwriter") as writer:

        tabla.to_excel(writer, sheet_name="Resumen", index=False)
        tabla2.to_excel(writer, sheet_name="Detalle Omisiones", index=False)
        tabla3.to_excel(writer, sheet_name="Resumen Medicos", index=False)
        if nuevos_medicos:
            pd.DataFrame(nuevos_medicos).to_excel(
                writer,
                sheet_name="Nuevos Medicos",
                index=False
            )

        if no_medicos:
            pd.DataFrame({
                "PROFESIONAL": sorted(set(no_medicos))
            }).to_excel(
                writer,
                sheet_name="Excluidos Ley18",
                index=False
            )

    st.download_button(
        "Descargar Excel",
        data=salida.getvalue(),
        file_name="resultado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
