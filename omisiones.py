import streamlit as st
import pandas as pd
from io import BytesIO
from docxtpl import DocxTemplate


# =========================================================
# CONFIGURACIÓN
# =========================================================

st.set_page_config(
    page_title="Analizador de Horas Médicas",
    layout="wide"
)

st.title("Analizador de Horas Asignadas")


# =========================================================
# FUNCIONES
# =========================================================

def generar_word(plantilla, contexto):

    documento = DocxTemplate(plantilla)

    documento.render(contexto)

    salida = BytesIO()

    documento.save(salida)

    salida.seek(0)

    return salida


def procesar_excel(archivo):

    # =====================================================
    # HOJAS
    # =====================================================

    hoja1 = pd.read_excel(
        archivo,
        sheet_name=0
    )

    hoja2 = pd.read_excel(
        archivo,
        sheet_name=1
    )

    hoja3 = pd.read_excel(
        archivo,
        sheet_name=2
    )

    hoja1.columns = (
        hoja1.columns
        .str.strip()
        .str.upper()
    )

    hoja2.columns = (
        hoja2.columns
        .str.strip()
        .str.upper()
    )

    hoja3.columns = (
        hoja3.columns
        .str.strip()
        .str.upper()
    )

    # =====================================================
    # COLUMNAS
    # =====================================================

    col_h1_prof = "NOMBRE PROFESIONAL"
    col_h1_agr = "AGRUPACION"
    col_h1_estado = "ESTADO HORA"

    col_h2_prof = "PROFESIONAL"
    col_h2_esp = "ESPECIALIDAD"

    col_h3_prof = "PROFESIONAL LEY 18"

    # =====================================================
    # VALIDACIÓN
    # =====================================================

    for col in [
        col_h1_prof,
        col_h1_agr,
        col_h1_estado
    ]:

        if col not in hoja1.columns:

            return None, f"Falta columna en Hoja 1: {col}"

    if (
        col_h2_prof not in hoja2.columns
        or col_h2_esp not in hoja2.columns
    ):

        return None, "Hoja 2 inválida"

    if col_h3_prof not in hoja3.columns:

        return None, "Hoja 3 inválida"

    # =====================================================
    # BASE
    # =====================================================

    df_asignadas = hoja1[
        hoja1[col_h1_estado]
        .astype(str)
        .str.upper()
        .eq("ASIGNADA")
    ].copy()

    # =====================================================
    # PADRONES
    # =====================================================

    medicos_hoja2 = set(
        hoja2[col_h2_prof]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    no_medicos_hoja3 = set(
        hoja3[col_h3_prof]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    especialidades = dict(
        zip(
            hoja2[col_h2_prof]
            .astype(str)
            .str.strip()
            .str.upper(),

            hoja2[col_h2_esp]
            .astype(str)
            .str.strip()
        )
    )

    # =====================================================
    # AGRUPACIONES
    # =====================================================

    agrup_medicos = {
        "MEDICO APS",
        "MEDICO ESPECIALISTA",
        "ODONTOLOGIA APS",
        "ODONTOLOGIA ESPECIALIDADES",
        "QUIMICO FARMACEUTICO"
    }

    agrup_no_medicos = {
        "TERAPEUTA OCUPACIONAL",
        "PSICOLOGIA",
        "ENFERMERA(O)",
        "ASISTENTE SOCIAL",
        "NUTRICIONISTA",
        "TECNOLOGO MEDICO",
        "FONOAUDIOLOGO",
        "MATRON(A)",
        "KINESIOLOGO"
    }

    # =====================================================
    # CLASIFICACIÓN INICIAL
    # =====================================================

    tipos = []
    especialidad_final = []
    desconocidos_proc = []

    for _, fila in df_asignadas.iterrows():

        prof = str(
            fila[col_h1_prof]
        ).strip().upper()

        agr = str(
            fila[col_h1_agr]
        ).strip().upper()

        if prof in no_medicos_hoja3:

            tipos.append("NO_MEDICO")
            especialidad_final.append(None)

        elif agr in agrup_medicos:

            tipos.append("MEDICO")

            especialidad_final.append(
                especialidades.get(
                    prof,
                    "SIN ESPECIALIDAD"
                )
            )

        elif agr in agrup_no_medicos:

            tipos.append("NO_MEDICO")
            especialidad_final.append(None)

        elif agr == "PROCEDIMIENTO":

            if prof in medicos_hoja2:

                tipos.append("MEDICO")

                especialidad_final.append(
                    especialidades.get(
                        prof,
                        "SIN ESPECIALIDAD"
                    )
                )

            elif prof in no_medicos_hoja3:

                tipos.append("NO_MEDICO")
                especialidad_final.append(None)

            else:

                tipos.append("PROC_DUDOSO")
                especialidad_final.append(None)

                desconocidos_proc.append(prof)

        else:

            tipos.append("PROC_DUDOSO")
            especialidad_final.append(None)

            desconocidos_proc.append(prof)

    df_asignadas["TIPO_PROFESIONAL"] = tipos

    df_asignadas["ESPECIALIDAD_FINAL"] = (
        especialidad_final
    )

    return {
        "hoja1": hoja1,
        "hoja2": hoja2,
        "hoja3": hoja3,
        "df_asignadas": df_asignadas,
        "medicos_hoja2": medicos_hoja2,
        "no_medicos_hoja3": no_medicos_hoja3,
        "especialidades": especialidades,
        "agrup_medicos": agrup_medicos,
        "agrup_no_medicos": agrup_no_medicos,
        "desconocidos_proc": sorted(
            set(desconocidos_proc)
        ),
        "col_h1_prof": col_h1_prof,
        "col_h1_agr": col_h1_agr,
    }, None


# =========================================================
# SUBIR EXCEL
# =========================================================

archivo = st.file_uploader(
    "Sube archivo Excel",
    type=["xlsx"]
)


# =========================================================
# PROCESAR EXCEL SOLO CUANDO CAMBIA EL ARCHIVO
# =========================================================

if archivo:

    # Identificador del archivo actual
    archivo_id = (
        archivo.name,
        archivo.size
    )

    # Si es un archivo nuevo, procesarlo
    if (
        "archivo_id" not in st.session_state
        or st.session_state["archivo_id"] != archivo_id
    ):

        with st.spinner(
            "Procesando archivo Excel..."
        ):

            resultado, error = procesar_excel(
                archivo
            )

        if error:

            st.error(error)

            st.stop()

        # Guardar resultado
        st.session_state["archivo_id"] = archivo_id

        st.session_state["resultado"] = resultado

        # Limpiar documentos Word anteriores
        st.session_state.pop(
            "documento_medico",
            None
        )

        st.session_state.pop(
            "documento_ley18",
            None
        )


    # =====================================================
    # RECUPERAR RESULTADO PROCESADO
    # =====================================================

    resultado = st.session_state["resultado"]

    hoja1 = resultado["hoja1"]
    df_asignadas = resultado["df_asignadas"]

    medicos_hoja2 = resultado["medicos_hoja2"]
    no_medicos_hoja3 = resultado["no_medicos_hoja3"]

    especialidades = resultado["especialidades"]

    agrup_medicos = resultado["agrup_medicos"]
    agrup_no_medicos = resultado["agrup_no_medicos"]

    desconocidos_proc = resultado["desconocidos_proc"]

    col_h1_prof = resultado["col_h1_prof"]
    col_h1_agr = resultado["col_h1_agr"]


    # =====================================================
    # REVISIÓN PROCEDIMIENTO
    # =====================================================

    st.subheader("🔎 Revisión PROCEDIMIENTO")

    nuevos_medicos = []
    nuevos_no_medicos = []

    for prof in desconocidos_proc:

        st.warning(
            f"{prof} no está en Hoja 2 ni Hoja 3"
        )

        tipo = st.radio(
            f"{prof} es:",
            ["No Médico", "Médico"],
            key=f"tipo_{prof}"
        )

        if tipo == "Médico":

            esp = st.text_input(
                f"Especialidad {prof}",
                key=f"esp_{prof}"
            )

            if esp:

                nuevos_medicos.append({
                    "PROFESIONAL": prof,
                    "ESPECIALIDAD": esp
                })

                medicos_hoja2.add(prof)

                especialidades[prof] = esp

        else:

            nuevos_no_medicos.append(prof)

            no_medicos_hoja3.add(prof)


    # =====================================================
    # RECLASIFICACIÓN FINAL
    # =====================================================

    def clasificar(prof, agr):

        prof = str(
            prof
        ).strip().upper()

        agr = str(
            agr
        ).strip().upper()

        if prof in no_medicos_hoja3:

            return "NO_MEDICO"

        if agr in agrup_medicos:

            return "MEDICO"

        if agr in agrup_no_medicos:

            return "NO_MEDICO"

        if agr == "PROCEDIMIENTO":

            if prof in medicos_hoja2:

                return "MEDICO"

            if prof in no_medicos_hoja3:

                return "NO_MEDICO"

            return "PROC_DUDOSO"

        return "PROC_DUDOSO"


    df_asignadas["TIPO_PROFESIONAL"] = (
        df_asignadas.apply(
            lambda r: clasificar(
                r[col_h1_prof],
                r[col_h1_agr]
            ),
            axis=1
        )
    )


    df_asignadas["ESPECIALIDAD_FINAL"] = (
        df_asignadas.apply(
            lambda r: (
                especialidades.get(
                    str(
                        r[col_h1_prof]
                    ).strip().upper(),
                    "SIN ESPECIALIDAD"
                )
                if r["TIPO_PROFESIONAL"] == "MEDICO"
                else None
            ),
            axis=1
        )
    )


    # =====================================================
    # BASES
    # =====================================================

    df_medicos = df_asignadas[
        df_asignadas["TIPO_PROFESIONAL"] == "MEDICO"
    ].copy()

    df_no_medicos = df_asignadas[
        df_asignadas["TIPO_PROFESIONAL"] == "NO_MEDICO"
    ].copy()

    df_proc = df_asignadas[
        df_asignadas["TIPO_PROFESIONAL"] == "PROC_DUDOSO"
    ].copy()

    df_medicos["OMISIONES"] = 1

    df_no_medicos["OMISIONES"] = 1


    # =====================================================
    # RESUMEN DE OMISIONES POR POLICLÍNICO
    # =====================================================
   
    tabla_omisiones_policlinico = (
        df_no_medicos
        .groupby(
            "POLICLINICO",
            dropna=False
        )
        .size()
        .reset_index(
            name="OMISIONES"
        )
    )
   
    # Reemplazar valores vacíos
    tabla_omisiones_policlinico["POLICLINICO"] = (
        tabla_omisiones_policlinico["POLICLINICO"]
        .fillna("SIN POLICLÍNICO")
        .astype(str)
        .str.strip()
    )
   
    # Ordenar de mayor a menor
    tabla_omisiones_policlinico = (
        tabla_omisiones_policlinico
        .sort_values(
            "OMISIONES",
            ascending=False
        )
        .reset_index(drop=True)
    )
   
    # =====================================================
    # NUMERACIÓN
    # =====================================================
   
    tabla_omisiones_policlinico["N"] = range(
        1,
        len(tabla_omisiones_policlinico) + 1
    )
   
       
    # =====================================================
    # CONVERTIR PARA WORD
    # =====================================================
   
    filas_omisiones_policlinico = (
        tabla_omisiones_policlinico[
            [
                "N",
                "POLICLINICO",
                "OMISIONES"
            ]
        ]
        .to_dict(
            orient="records"
        )
    )
    
    # =====================================================
    # TOTAL
    # =====================================================
    
    total_omisiones_policlinico = int(
        tabla_omisiones_policlinico["OMISIONES"].sum()
    )

    # =====================================================
    # TOTAL
    # =====================================================
    
    total_omisiones_policlinico = int(
        tabla_omisiones_policlinico["OMISIONES"].sum()
    )

    
    
    # =====================================================
    # RESUMEN DE OMISIONES POR ESPECIALIDAD MÉDICA
    # =====================================================
    
    tabla_omisiones_especialidad = (
        df_medicos
        .groupby(
            "ESPECIALIDAD_FINAL",
            dropna=False
        )
        .size()
        .reset_index(
            name="OMISIONES"
        )
    )
    
    # Limpiar especialidades vacías
    tabla_omisiones_especialidad["ESPECIALIDAD_FINAL"] = (
        tabla_omisiones_especialidad["ESPECIALIDAD_FINAL"]
        .fillna("SIN ESPECIALIDAD")
        .astype(str)
        .str.strip()
    )
    
    # Ordenar de mayor a menor
    tabla_omisiones_especialidad = (
        tabla_omisiones_especialidad
        .sort_values(
            "OMISIONES",
            ascending=False
        )
        .reset_index(drop=True)
    )
    
    # Convertir a lista para Word
    filas_omisiones_especialidad = (
        tabla_omisiones_especialidad
        .to_dict(
            orient="records"
        )
    )
    
    # Total
    total_omisiones_especialidad = int(
        tabla_omisiones_especialidad["OMISIONES"].sum()
)

    
    # =====================================================
    # RESUMEN DE OMISIONES POR FUNCIONARIO Y POLICLÍNICO
    # =====================================================
    
    tabla_funcionario_policlinico = (
        df_no_medicos
        .groupby(
            [
                col_h1_prof,
                "POLICLINICO"
            ],
            dropna=False
        )
        .size()
        .reset_index(name="OMISIONES")
    )
    
    # Renombrar funcionario
    tabla_funcionario_policlinico = (
        tabla_funcionario_policlinico
        .rename(
            columns={
                col_h1_prof: "FUNCIONARIO"
            }
        )
    )
    
    # Limpiar funcionario
    tabla_funcionario_policlinico["FUNCIONARIO"] = (
        tabla_funcionario_policlinico["FUNCIONARIO"]
        .fillna("SIN FUNCIONARIO")
        .astype(str)
        .str.strip()
    )
    
    # Limpiar policlínico
    tabla_funcionario_policlinico["POLICLINICO"] = (
        tabla_funcionario_policlinico["POLICLINICO"]
        .fillna("SIN POLICLÍNICO")
        .astype(str)
        .str.strip()
    )
    
    
    # =====================================================
    # ORDENAR SEGÚN LA TABLA DE POLICLÍNICOS
    # =====================================================
    
    orden_policlinicos = (
        tabla_omisiones_policlinico["POLICLINICO"]
        .tolist()
    )
    
    tabla_funcionario_policlinico["POLICLINICO"] = (
        pd.Categorical(
            tabla_funcionario_policlinico["POLICLINICO"],
            categories=orden_policlinicos,
            ordered=True
        )
    )
    
    
    # =====================================================
    # ORDEN FINAL
    # =====================================================
    
    tabla_funcionario_policlinico = (
        tabla_funcionario_policlinico
        .sort_values(
            [
                "POLICLINICO",
                "OMISIONES"
            ],
            ascending=[
                True,
                False
            ]
        )
        .reset_index(drop=True)
    )
    
    # Convertir a lista para Word
    filas_funcionario_policlinico = (
        tabla_funcionario_policlinico
        .to_dict(orient="records")
    )

    
    
    # =====================================================
    # CONVERTIR PARA WORD
    # =====================================================
    
    filas_funcionario_policlinico = (
        tabla_funcionario_policlinico[
            [
                "FUNCIONARIO",
                "POLICLINICO",
                "OMISIONES"
            ]
        ]
        .to_dict(
            orient="records"
        )
    )

    # =====================================================
    # RESUMEN DE OMISIONES POR PROFESIONAL Y ESPECIALIDAD
    # =====================================================
    
    tabla_omisiones_profesional = (
        df_medicos
        .groupby(
            [
                col_h1_prof,
                "ESPECIALIDAD_FINAL"
            ],
            dropna=False
        )
        .size()
        .reset_index(
            name="OMISIONES"
        )
    )
    
    # Limpiar nombre del profesional
    tabla_omisiones_profesional[col_h1_prof] = (
        tabla_omisiones_profesional[col_h1_prof]
        .fillna("SIN PROFESIONAL")
        .astype(str)
        .str.strip()
    )
    
    # Limpiar especialidad
    tabla_omisiones_profesional["ESPECIALIDAD_FINAL"] = (
        tabla_omisiones_profesional["ESPECIALIDAD_FINAL"]
        .fillna("SIN ESPECIALIDAD")
        .astype(str)
        .str.strip()
    )
    
    # =====================================================
    # CAMBIAR NOMBRES PARA WORD
    # =====================================================

    tabla_omisiones_profesional = (
        tabla_omisiones_profesional
        .rename(
            columns={
                col_h1_prof: "NOMBRE",
                "ESPECIALIDAD_FINAL": "ESPECIALIDAD"
            }
        )
    )
    
   
    # =====================================================
    # ORDENAR SEGÚN EL ORDEN DE LA TABLA DE ESPECIALIDADES
    # =====================================================

    orden_especialidades = (
        tabla_omisiones_especialidad[
            "ESPECIALIDAD_FINAL"
        ]
        .tolist()
    )


    tabla_omisiones_profesional["ESPECIALIDAD"] = (
        pd.Categorical(
            tabla_omisiones_profesional["ESPECIALIDAD"],
            categories=orden_especialidades,
            ordered=True
        )
    )

    # =====================================================
    # ORDEN FINAL
    # =====================================================

    tabla_omisiones_profesional = (
        tabla_omisiones_profesional
        .sort_values(
            [
                "ESPECIALIDAD",
                "OMISIONES"
            ],
            ascending=[
                True,
                False
            ]
        )
        .reset_index(drop=True)
    )


    # =====================================================
    # CONVERTIR PARA WORD
    # =====================================================

    filas_omisiones_profesional = (
        tabla_omisiones_profesional[
            [
                "NOMBRE",
                "ESPECIALIDAD",
                "OMISIONES"
            ]
        ]
        .to_dict(
            orient="records"
        )
    )




    
    # =====================================================
    # RESUMEN GENERAL
    # =====================================================

    st.markdown(
        "## 📊 Resumen General de Omisiones"
    )

    total_asignadas = len(
        df_asignadas
    )

    total_medicos = len(
        df_medicos
    )

    total_no_medicos = len(
        df_no_medicos
    )

    col1, col2, col3 = st.columns(3)

    col1.metric(
        "Total Omisiones (Asignadas)",
        total_asignadas
    )

    col2.metric(
        "Omisiones Médicos",
        total_medicos
    )

    col3.metric(
        "Omisiones No Médicos",
        total_no_medicos
    )


    # =====================================================
    # FECHAS
    # =====================================================

    st.markdown(
        "## 📄 Generación de documentos"
    )

    col_fecha1, col_fecha2 = st.columns(2)

    with col_fecha1:

        fecha_corte = st.date_input(
            "Fecha de corte",
            format="DD/MM/YYYY",
            key="fecha_corte"
        )

    with col_fecha2:

        fecha_envio_preliminar = st.date_input(
            "Fecha de envío preliminar",
            format="DD/MM/YYYY",
            key="fecha_envio_preliminar"
        )


    # =====================================================
    # TOTAL AGENDADAS - EJECUTADAS
    # =====================================================

    total_agendadas = len(
        hoja1
    )
        # Total de horas ejecutadas
    total_ejecutadas = (
        hoja1["ESTADO HORA"]
        .astype(str)
        .str.strip()
        .str.upper()
        .eq("EJECUTADA")
        .sum()
    )



    
    st.markdown("### 📊 Datos para documento")

    col1, col2, col3 = st.columns(3)
    
    with col1:
    
        st.metric(
            "Total asignadas",
            total_asignadas
        )
    
    with col2:
    
        st.metric(
            "Total horas agendadas",
            total_agendadas
        )
    
    with col3:
    
        st.metric(
            "Total horas ejecutadas",
            total_ejecutadas
        )


    # =====================================================
    # PLANTILLAS WORD
    # =====================================================

    st.markdown(
        "### 📄 Plantillas Word"
    )

    col_word1, col_word2 = st.columns(2)

    with col_word1:

        st.markdown(
            "#### 🩺 Plantilla Ley Médica"
        )

        plantilla_ley_medica = st.file_uploader(
            "Sube la plantilla Word de Ley Médica",
            type=["docx"],
            key="plantilla_ley_medica"
        )

    with col_word2:

        st.markdown(
            "#### 📋 Plantilla Ley 18"
        )

        plantilla_ley_18 = st.file_uploader(
            "Sube la plantilla Word de Ley 18",
            type=["docx"],
            key="plantilla_ley_18"
        )


    # =====================================================
    # TABLAS PARA EXCEL
    # =====================================================

    tabla_resumen_medicos = (
        df_medicos
        .groupby("ESPECIALIDAD_FINAL")
        .size()
        .reset_index(
            name="TOTAL ASIGNADAS"
        )
    )


    tabla_medicos_detalle = (
        df_medicos
        .groupby(
            [
                "ESPECIALIDAD_FINAL",
                col_h1_prof
            ]
        )
        .size()
        .reset_index(
            name="TOTAL ASIGNADAS"
        )
    )

    tabla_medicos_detalle = (
        tabla_medicos_detalle.rename(
            columns={
                "ESPECIALIDAD_FINAL": "ESPECIALIDAD",
                col_h1_prof: "NOMBRE PROFESIONAL"
            }
        )
    )


    tabla_medicos_pacientes = (
        df_medicos
        .groupby(
            [
                "ESPECIALIDAD_FINAL",
                "RUT PROFESIONAL",
                col_h1_prof,
                "RUT PACIENTE",
                "NOMBRE PACIENTE",
                "FECHA"
            ],
            dropna=False
        )
        .size()
        .reset_index(
            name="OMISIONES"
        )
    )

    tabla_medicos_pacientes = (
        tabla_medicos_pacientes.rename(
            columns={
                "ESPECIALIDAD_FINAL": "ESPECIALIDAD",
                col_h1_prof: "NOMBRE PROFESIONAL"
            }
        )
    )


    tabla_resumen_no_medicos = (
        df_no_medicos
        .groupby("POLICLINICO")
        .size()
        .reset_index(
            name="TOTAL ASIGNADAS"
        )
    )


    tabla_no_medicos_detalle = (
        df_no_medicos
        .groupby(
            [
                col_h1_prof,
                "POLICLINICO"
            ]
        )
        .size()
        .reset_index(
            name="TOTAL ASIGNADAS"
        )
    )

    tabla_no_medicos_detalle = (
        tabla_no_medicos_detalle.rename(
            columns={
                col_h1_prof: "NOMBRE PROFESIONAL"
            }
        )
    )


    tabla_no_medicos_pacientes = (
        df_no_medicos
        .groupby(
            [
                "POLICLINICO",
                "RUT PROFESIONAL",
                col_h1_prof,
                "RUT PACIENTE",
                "NOMBRE PACIENTE",
                "FECHA"
            ]
        )
        .size()
        .reset_index(
            name="TOTAL ASIGNADAS"
        )
    )

    tabla_no_medicos_pacientes = (
        tabla_no_medicos_pacientes.rename(
            columns={
                col_h1_prof: "NOMBRE PROFESIONAL"
            }
        )
    )


    # =====================================================
    # EXPORTAR EXCEL
    # =====================================================

    salida = BytesIO()

    with pd.ExcelWriter(
        salida,
        engine="xlsxwriter"
    ) as writer:

        tabla_resumen_medicos.to_excel(
            writer,
            sheet_name="Resumen Medicos",
            index=False
        )

        tabla_medicos_detalle.to_excel(
            writer,
            sheet_name="Detalle Medicos",
            index=False
        )

        tabla_medicos_pacientes.to_excel(
            writer,
            sheet_name="Pacientes Medicos",
            index=False
        )

        tabla_resumen_no_medicos.to_excel(
            writer,
            sheet_name="Resumen No Medicos",
            index=False
        )

        tabla_no_medicos_detalle.to_excel(
            writer,
            sheet_name="Detalle No Medicos",
            index=False
        )

        tabla_no_medicos_pacientes.to_excel(
            writer,
            sheet_name="Pacientes No Medicos",
            index=False
        )

        if nuevos_medicos:

            pd.DataFrame(
                nuevos_medicos
            ).to_excel(
                writer,
                sheet_name="Nuevos Medicos",
                index=False
            )


    # =====================================================
    # DESCARGAR EXCEL
    # =====================================================

    st.download_button(
        "📥 Descargar Excel",
        data=salida.getvalue(),
        file_name="resultado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="descargar_excel"
    )


    # =====================================================
    # GENERAR WORD
    # =====================================================

    st.markdown(
        "### 📄 Generación de documentos Word"
    )

    generar_documentos = st.button(
        "📄 Generar documentos Word",
        type="primary",
        key="generar_documentos_word"
    )


    if generar_documentos:

        # -------------------------------------------------
        # VALIDAR PLANTILLAS
        # -------------------------------------------------

        if (
            plantilla_ley_medica is None
            and plantilla_ley_18 is None
        ):

            st.warning(
                "Debes subir al menos una plantilla Word."
            )

        else:


            # =========================
            # CALCULO PORCENTAJE
            # =========================
    
            porcentaje_asignadas_agendadas = (
                round(
                    (total_asignadas / total_agendadas) * 100,
                    1
                )
                if total_agendadas > 0
                else 0
            )
            
            porcentaje_ley18_agendadas = (
                round(
                    (total_no_medicos / total_agendadas) * 100,
                    1
                )
                if total_agendadas > 0
                else 0
            )

            porcentaje_leymedica_agendadas = (
                round(
                    (total_medicos / total_agendadas) * 100,
                    1
                )
                if total_agendadas > 0
                else 0
            )

            # -------------------------------------------------
            # CONTEXTO
            # -------------------------------------------------

            contexto = {
            
                "fecha_corte":
                    fecha_corte.strftime("%d/%m/%Y"),
            
                "fecha_envio_preliminar":
                    fecha_envio_preliminar.strftime("%d/%m/%Y"),
            
                "mes_corte": [
                    "enero",
                    "febrero",
                    "marzo",
                    "abril",
                    "mayo",
                    "junio",
                    "julio",
                    "agosto",
                    "septiembre",
                    "octubre",
                    "noviembre",
                    "diciembre"
                ][fecha_corte.month - 1],
            
                "total_agendadas":
                    total_agendadas,
            
                "total_asignadas":
                    total_asignadas,
            
                "total_omisiones_medicos":
                    total_medicos,
            
                "total_omisiones_Ley18":
                    total_no_medicos,
            
                "porc_asignadas_agendadas":
                    porcentaje_asignadas_agendadas,
            
                "porc_ley18_agendadas":
                    porcentaje_ley18_agendadas,
            
                "porc_leymedica_agendadas":
                    porcentaje_leymedica_agendadas,
            
                "total_ejecutadas":
                    total_ejecutadas,
            
                # -----------------------------------------
                # POLICLÍNICOS
                # -----------------------------------------
            
                "omisiones_policlinico":
                    filas_omisiones_policlinico,
            
                "total_omisiones_policlinico":
                    total_omisiones_policlinico,
            
                # -----------------------------------------
                # FUNCIONARIOS
                # -----------------------------------------
            
                "funcionarios_policlinico":
                    filas_funcionario_policlinico,
            
                # -----------------------------------------
                # ESPECIALIDADES
                # -----------------------------------------
            
                "omisiones_especialidad":
                    filas_omisiones_especialidad,
            
                "total_omisiones_especialidad":
                    total_omisiones_especialidad,

                "omisiones_profesional":
                    filas_omisiones_profesional,

}


            # -------------------------------------------------
            # LEY MÉDICA
            # -------------------------------------------------

            if plantilla_ley_medica is not None:

                with st.spinner(
                    "Generando Word Ley Médica..."
                ):

                    documento_medico = generar_word(
                        plantilla_ley_medica,
                        contexto
                    )

                st.session_state[
                    "documento_medico"
                ] = documento_medico.getvalue()


            # -------------------------------------------------
            # LEY 18
            # -------------------------------------------------

            if plantilla_ley_18 is not None:

                with st.spinner(
                    "Generando Word Ley 18..."
                ):

                    documento_ley18 = generar_word(
                        plantilla_ley_18,
                        contexto
                    )

                st.session_state[
                    "documento_ley18"
                ] = documento_ley18.getvalue()


            st.success(
                "Documentos Word generados correctamente."
            )


    # =====================================================
    # DESCARGAS WORD
    # =====================================================

    if (
        "documento_medico"
        in st.session_state
        or
        "documento_ley18"
        in st.session_state
    ):

        st.markdown(
            "### 📥 Documentos completados"
        )

        col_word1, col_word2 = st.columns(2)


        # -------------------------------------------------
        # LEY MÉDICA
        # -------------------------------------------------

        with col_word1:

            if (
                "documento_medico"
                in st.session_state
            ):

                st.download_button(
                    label="📥 Descargar Word Ley Médica",
                    data=st.session_state[
                        "documento_medico"
                    ],
                    file_name="Informe_Ley_Medica.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="descargar_ley_medica"
                )


        # -------------------------------------------------
        # LEY 18
        # -------------------------------------------------

        with col_word2:

            if (
                "documento_ley18"
                in st.session_state
            ):

                st.download_button(
                    label="📥 Descargar Word Ley 18",
                    data=st.session_state[
                        "documento_ley18"
                    ],
                    file_name="Informe_Ley_18.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="descargar_ley_18"
                )
