import streamlit as st
import pandas as pd
from io import BytesIO
import plotly.express as px

st.set_page_config(
    page_title="Analizador de Horas Médicas",
    layout="wide"
)

st.title("Analizador de Horas Asignadas")

archivo = st.file_uploader("Sube archivo Excel", type=["xlsx"])

if archivo:

    # =========================
    # LECTURA DE HOJAS
    # =========================
    hoja1 = pd.read_excel(archivo, sheet_name=0)
    hoja2 = pd.read_excel(archivo, sheet_name=1)
    hoja3 = pd.read_excel(archivo, sheet_name=2)

    hoja1.columns = hoja1.columns.str.strip().str.upper()
    hoja2.columns = hoja2.columns.str.strip().str.upper()
    hoja3.columns = hoja3.columns.str.strip().str.upper()

    # =========================
    # COLUMNAS
    # =========================
    col_h1_prof = "NOMBRE PROFESIONAL"
    col_h1_agr = "AGRUPACION"
    col_h1_estado = "ESTADO HORA"

    col_h2_prof = "PROFESIONAL"
    col_h2_esp = "ESPECIALIDAD"

    col_h3_prof = "PROFESIONAL LEY 18"

    # =========================
    # VALIDACIONES
    # =========================
    for col in [col_h1_prof, col_h1_agr, col_h1_estado]:
        if col not in hoja1.columns:
            st.error(f"Falta columna en Hoja 1: {col}")
            st.stop()

    if col_h2_prof not in hoja2.columns or col_h2_esp not in hoja2.columns:
        st.error("Hoja 2 no válida")
        st.stop()

    if col_h3_prof not in hoja3.columns:
        st.error("Hoja 3 no válida")
        st.stop()

    # =========================
    # BASE ASIGNADAS
    # =========================
    df_asignadas = hoja1[
        hoja1[col_h1_estado].astype(str).str.upper().eq("ASIGNADA")
    ].copy()

    # =========================
    # PADRONES NORMALIZADOS
    # =========================
    medicos_hoja2 = set(
        hoja2[col_h2_prof].astype(str).str.strip().str.upper()
    )

    no_medicos_hoja3 = set(
        hoja3[col_h3_prof].astype(str).str.strip().str.upper()
    )

    especialidades = dict(
        zip(
            hoja2[col_h2_prof].astype(str).str.strip().str.upper(),
            hoja2[col_h2_esp].astype(str).str.strip()
        )
    )

    # =========================
    # CLASIFICACION
    # =========================
    tipos = []
    especialidad_final = []
    desconocidos = []

    for _, fila in df_asignadas.iterrows():

        profesional = str(fila[col_h1_prof]).strip().upper()

        if profesional in medicos_hoja2:

            tipos.append("MEDICO")
            especialidad_final.append(
                especialidades.get(profesional, "SIN ESPECIALIDAD")
            )

        elif profesional in no_medicos_hoja3:

            tipos.append("NO_MEDICO")
            especialidad_final.append(None)

        else:

            tipos.append("DESCONOCIDO")
            especialidad_final.append(None)
            desconocidos.append(profesional)

    df_asignadas["TIPO_PROFESIONAL"] = tipos
    df_asignadas["ESPECIALIDAD_FINAL"] = especialidad_final

    # =========================
    # PREGUNTAS USUARIO
    # =========================
    st.subheader("Profesionales no encontrados")

    nuevos_medicos = []
    nuevos_no_medicos = []

    for profesional in sorted(set(desconocidos)):

        st.warning(f"{profesional} no encontrado")

        tipo = st.radio(
            f"{profesional} es:",
            ["No Médico", "Médico"],
            key=profesional
        )

        if tipo == "Médico":

            esp = st.text_input(
                f"Especialidad {profesional}",
                key=f"esp_{profesional}"
            )

            if esp:
                nuevos_medicos.append({
                    "PROFESIONAL": profesional,
                    "ESPECIALIDAD": esp
                })
                medicos_hoja2.add(profesional.upper())
                especialidades[profesional.upper()] = esp

        else:
            nuevos_no_medicos.append(profesional)
            no_medicos_hoja3.add(profesional.upper())

    # =========================
    # RECLASIFICACION FINAL
    # =========================
    def clasificar(x):
        x = str(x).strip().upper()
        if x in medicos_hoja2:
            return "MEDICO"
        if x in no_medicos_hoja3:
            return "NO_MEDICO"
        return "SIN_CLASIFICAR"

    df_asignadas["TIPO_PROFESIONAL"] = df_asignadas[col_h1_prof].apply(clasificar)

    df_asignadas["ESPECIALIDAD_FINAL"] = df_asignadas[col_h1_prof].astype(str).str.upper().map(especialidades)

    # =========================
    # BASES FINALES
    # =========================
    df_medicos = df_asignadas[df_asignadas["TIPO_PROFESIONAL"] == "MEDICO"].copy()
    df_no_medicos = df_asignadas[df_asignadas["TIPO_PROFESIONAL"] == "NO_MEDICO"].copy()

    df_medicos["OMISIONES"] = 1
    df_no_medicos["OMISIONES"] = 1

    # =========================
    # CONTROL
    # =========================
    st.markdown("## Control de Consistencia")

    st.write("Asignadas:", len(df_asignadas))
    st.write("Médicos:", len(df_medicos))
    st.write("No Médicos:", len(df_no_medicos))
    st.write("Sin clasificar:", len(df_asignadas[df_asignadas["TIPO_PROFESIONAL"] == "SIN_CLASIFICAR"]))

    # =========================
    # TABLA RESUMEN MEDICOS
    # =========================
    tabla = (
        df_medicos.groupby("ESPECIALIDAD_FINAL")
        .size()
        .reset_index(name="TOTAL")
        .sort_values("TOTAL", ascending=False)
    )

    tabla2 = (
        df_medicos.groupby([col_h1_prof, "ESPECIALIDAD_FINAL"])
        .size()
        .reset_index(name="OMISIONES")
    )

    tabla3 = (
        df_medicos.groupby(col_h1_prof)
        .size()
        .reset_index(name="OMISIONES")
        .rename(columns={col_h1_prof: "NOMBRE PROFESIONAL"})
    )

    # =========================
    # TABLA NO MEDICOS
    # =========================
    tabla4 = (
        df_no_medicos.groupby(col_h1_prof)
        .size()
        .reset_index(name="OMISIONES")
    )

    tabla5 = df_no_medicos.copy()
    tabla5 = tabla5.rename(columns={col_h1_prof: "NOMBRE PROFESIONAL"})

    tabla6 = (
        df_no_medicos.groupby(col_h1_prof)
        .size()
        .reset_index(name="OMISIONES")
        .rename(columns={col_h1_prof: "NOMBRE PROFESIONAL"})
    )

    # =========================
    # DASHBOARD
    # =========================
    st.markdown("## Resumen")

    col1, col2, col3 = st.columns(3)

    col1.metric("Asignadas", len(df_asignadas))
    col2.metric("Médicos", len(df_medicos))
    col3.metric("No Médicos", len(df_no_medicos))

    st.dataframe(tabla)

    st.dataframe(tabla3)

    st.dataframe(tabla4)

    # =========================
    # EXPORT EXCEL
    # =========================
    salida = BytesIO()

    with pd.ExcelWriter(salida, engine="xlsxwriter") as writer:

        tabla.to_excel(writer, sheet_name="Resumen", index=False)
        tabla2.to_excel(writer, sheet_name="Detalle Medicos", index=False)
        tabla3.to_excel(writer, sheet_name="Ranking Medicos", index=False)
        tabla4.to_excel(writer, sheet_name="Ranking No Medicos", index=False)

        if nuevos_medicos:
            pd.DataFrame(nuevos_medicos).to_excel(
                writer,
                sheet_name="Nuevos Medicos",
                index=False
            )

    st.download_button(
        "Descargar Excel",
        data=salida.getvalue(),
        file_name="resultado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
