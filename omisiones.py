import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(
    page_title="Analizador de Horas Médicas",
    layout="wide"
)

st.title("Analizador de Horas Asignadas")

archivo = st.file_uploader(
    "Seleccione archivo Excel",
    type=["xlsx"]
)

if archivo:

    # Leer hojas
    hoja1 = pd.read_excel(archivo, sheet_name=0)
    hoja2 = pd.read_excel(archivo, sheet_name=1)
    hoja3 = pd.read_excel(archivo, sheet_name=2)

    # Normalizar
    hoja1.columns = hoja1.columns.str.strip().str.upper()
    hoja2.columns = hoja2.columns.str.strip().str.upper()
    hoja3.columns = hoja3.columns.str.strip().str.upper()

    hoja1["PROFESIONAL"] = hoja1["PROFESIONAL"].astype(str).str.strip()
    hoja2["PROFESIONAL"] = hoja2["PROFESIONAL"].astype(str).str.strip()
    hoja3["PROFESIONAL"] = hoja3["PROFESIONAL"].astype(str).str.strip()

    # Filtrar asignadas
    df = hoja1[
        hoja1["ESTADO HORA"]
        .astype(str)
        .str.upper()
        .eq("ASIGNADA")
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
        df["AGRUPACION"]
        .astype(str)
        .str.upper()
        .isin(agrupaciones_validas)
    ]

    # Diccionario médico -> especialidad
    especialidades = dict(
        zip(
            hoja2["PROFESIONAL"],
            hoja2["ESPECIALIDAD"]
        )
    )

    no_medicos = set(hoja3["PROFESIONAL"])

    especialidades_detectadas = []
    excluidos = []
    desconocidos = []

    for _, fila in df.iterrows():

        profesional = str(fila["PROFESIONAL"]).strip()

        if profesional in especialidades:
            especialidades_detectadas.append(
                especialidades[profesional]
            )

        elif profesional in no_medicos:
            excluidos.append(profesional)
            especialidades_detectadas.append(None)

        else:
            desconocidos.append(profesional)
            especialidades_detectadas.append(None)

    df["ESPECIALIDAD_FINAL"] = especialidades_detectadas

    st.subheader("Profesionales no encontrados")

    nuevos_medicos = []

    for profesional in sorted(set(desconocidos)):

        st.warning(
            f"{profesional} no existe en Hoja 2 ni Hoja 3"
        )

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
                    df["PROFESIONAL"] == profesional,
                    "ESPECIALIDAD_FINAL"
                ] = especialidad

    # Tabla principal
    tabla_especialidades = (
        df.dropna(subset=["ESPECIALIDAD_FINAL"])
        .groupby("ESPECIALIDAD_FINAL")
        .size()
        .reset_index(name="TOTAL")
        .sort_values("TOTAL", ascending=False)
    )

    st.subheader("Horas Asignadas por Especialidad")
    st.dataframe(
        tabla_especialidades,
        use_container_width=True
    )

    if excluidos:

        st.subheader(
            "Profesionales excluidos (Hoja 3)"
        )

        st.dataframe(
            pd.DataFrame(
                {"PROFESIONAL": sorted(set(excluidos))}
            ),
            use_container_width=True
        )

    # Actualizar hoja médicos
    hoja2_actualizada = hoja2.copy()

    if nuevos_medicos:

        hoja2_actualizada = pd.concat(
            [
                hoja2_actualizada,
                pd.DataFrame(nuevos_medicos)
            ],
            ignore_index=True
        )

    # Generar Excel de salida
    salida = BytesIO()

    with pd.ExcelWriter(
        salida,
        engine="xlsxwriter"
    ) as writer:

        tabla_especialidades.to_excel(
            writer,
            sheet_name="Resumen Especialidades",
            index=False
        )

        hoja2_actualizada.to_excel(
            writer,
            sheet_name="Medicos Actualizados",
            index=False
        )

        if excluidos:
            pd.DataFrame({
                "PROFESIONAL":
                sorted(set(excluidos))
            }).to_excel(
                writer,
                sheet_name="Excluidos",
                index=False
            )

    st.download_button(
        "Descargar Resultado",
        data=salida.getvalue(),
        file_name="resultado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
