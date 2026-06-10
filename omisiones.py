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

    # Leer Excel
    hoja1 = pd.read_excel(archivo, sheet_name=0)
    hoja2 = pd.read_excel(archivo, sheet_name=1)
    hoja3 = pd.read_excel(archivo, sheet_name=2)

    # Limpiar nombres de columnas
    hoja1.columns = hoja1.columns.str.strip().str.upper()
    hoja2.columns = hoja2.columns.str.strip().str.upper()
    hoja3.columns = hoja3.columns.str.strip().str.upper()

    # Columna correcta según tu Excel
    col_prof = "NOMBRE PROFESIONAL"

    # Validación básica
    if col_prof not in hoja1.columns:
        st.error(f"No existe la columna {col_prof}. Columnas: {hoja1.columns.tolist()}")
        st.stop()

    # Filtrar ASIGNADAS
    df = hoja1[
        hoja1["ESTADO HORA"].astype(str).str.upper() == "ASIGNADA"
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
        df["AGRUPACION"].astype(str).str.upper().isin(agrupaciones_validas)
    ]

    # Diccionario Hoja 2 (médicos)
    especialidades = dict(
        zip(
            hoja2["PROFESIONAL"].astype(str).str.strip(),
            hoja2["ESPECIALIDAD"].astype(str).str.strip()
        )
    )

    # Hoja 3 (no médicos)
    no_medicos = set(hoja3["PROFESIONAL"].astype(str).str.strip())

    resultados = []
    excluidos = []
    desconocidos = []

    for _, fila in df.iterrows():

        profesional = str(fila[col_prof]).strip()

        if profesional in especialidades:
            resultados.append(especialidades[profesional])

        elif profesional in no_medicos:
            excluidos.append(profesional)
            resultados.append(None)

        else:
            desconocidos.append(profesional)
            resultados.append(None)

    df["ESPECIALIDAD_FINAL"] = resultados

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
                    df[col_prof] == profesional,
                    "ESPECIALIDAD_FINAL"
                ] = especialidad

    # Tabla final
    tabla = (
        df.dropna(subset=["ESPECIALIDAD_FINAL"])
        .groupby("ESPECIALIDAD_FINAL")
        .size()
        .reset_index(name="TOTAL")
        .sort_values("TOTAL", ascending=False)
    )

    st.subheader("Resultado por Especialidad")
    st.dataframe(tabla, use_container_width=True)

    if excluidos:
        st.subheader("Excluidos (Hoja 3)")
        st.dataframe(pd.DataFrame({"PROFESIONAL": list(set(excluidos))}))

    # Exportar Excel
    salida = BytesIO()

    with pd.ExcelWriter(salida, engine="xlsxwriter") as writer:

        tabla.to_excel(writer, sheet_name="Resumen", index=False)

        if nuevos_medicos:
            pd.DataFrame(nuevos_medicos).to_excel(
                writer,
                sheet_name="Nuevos Medicos",
                index=False
            )

        if excluidos:
            pd.DataFrame({"PROFESIONAL": list(set(excluidos))}).to_excel(
                writer,
                sheet_name="Excluidos",
                index=False
            )

    st.download_button(
        "Descargar Excel",
        data=salida.getvalue(),
        file_name="resultado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
