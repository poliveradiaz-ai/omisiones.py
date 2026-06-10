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
    # FILTRO BASE
    # =========================
    df = hoja1[
        hoja1[col_h1_estado].astype(str).str.upper() == "ASIGNADA"
    ].copy()
   
    df_asignadas = df.copy()
    
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
    # BASE MÉDICOS (IMPORTANTE FIX)
    # =========================
    df_medicos = df[df[col_h1_prof].isin(hoja2[col_h2_prof])].copy()
    df_medicos["OMISIONES"] = 1

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
    # RANKING (BASE PARA ORDEN)
    # =========================
    ranking_esp = (
        df_medicos.groupby("ESPECIALIDAD_FINAL")["OMISIONES"]
        .sum()
        .sort_values(ascending=False)
        .reset_index()
    )

    orden_esp = ranking_esp["ESPECIALIDAD_FINAL"].tolist()

    # =========================
    # TABLA 1
    # =========================
    tabla = (
        df.dropna(subset=["ESPECIALIDAD_FINAL"])
        .groupby("ESPECIALIDAD_FINAL")
        .size()
        .reset_index(name="TOTAL")
        .sort_values("TOTAL", ascending=False)
    )

    # st.subheader("Resultado por Especialidad")
    # st.dataframe(tabla, use_container_width=True)

   
    # =========================
    # TABLA 2 (DETALLE CORREGIDO)
    # =========================

    tabla2 = df_medicos.copy()

    tabla2.rename(columns={
        col_h1_prof: "NOMBRE PROFESIONAL",
        "ESPECIALIDAD_FINAL": "ESPECIALIDAD"
    }, inplace=True)

    # =========================
    # 1. OMISIONES = 1 POR PACIENTE
    # =========================
    tabla2["OMISIONES"] = 1

    # =========================
    # 2. TOTAL OMISIONES POR MÉDICO
    # =========================
    ranking_prof = (
        tabla2.groupby(["ESPECIALIDAD", "NOMBRE PROFESIONAL"])["OMISIONES"]
        .sum()
        .reset_index(name="TOTAL OMISIONES")
    )
    
    # =========================
    # 3. RANKING DE ESPECIALIDAD
    # =========================
    ranking_esp = (
        ranking_prof.groupby("ESPECIALIDAD")["TOTAL OMISIONES"]
        .sum()
        .sort_values(ascending=False)
        .reset_index()
    )
    
    orden_esp = ranking_esp["ESPECIALIDAD"].tolist()
    
    # =========================
    # 4. UNIR TOTAL AL DETALLE
    # =========================
    tabla2 = tabla2.merge(
        ranking_prof,
        on=["ESPECIALIDAD", "NOMBRE PROFESIONAL"],
        how="left"
    )

    # =========================
    # 5. ORDEN FINAL (ESPECIALIDAD + SUBRANKING)
    # =========================
    tabla2["ORDEN_ESP"] = tabla2["ESPECIALIDAD"].apply(
        lambda x: orden_esp.index(x) if x in orden_esp else 999
    )
    
    tabla2 = tabla2.sort_values(
        by=["ORDEN_ESP", "TOTAL OMISIONES", "NOMBRE PROFESIONAL"],
        ascending=[True, False, True]
    )
    
    # =========================
    # 6. ASEGURAR COLUMNAS EXACTAS (6)
    # =========================
    for col in ["RUT PACIENTE", "NOMBRE PACIENTE", "FECHA"]:
        if col not in tabla2.columns:
            tabla2[col] = None
    
    tabla2 = tabla2[[
        "ESPECIALIDAD",
        "NOMBRE PROFESIONAL",
        "RUT PACIENTE",
        "NOMBRE PACIENTE",
        "FECHA",
        "OMISIONES"
    ]]

    # =========================
    # TABLA 3 (RESUMEN ORDENADO)
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

    tabla3["ORDEN"] = tabla3["ESPECIALIDAD"].apply(
        lambda x: orden_esp.index(x) if x in orden_esp else 999
    )

    tabla3 = tabla3.sort_values(["ORDEN", "OMISIONES"], ascending=[True, False])

    tabla3 = tabla3[[
        "ESPECIALIDAD",
        "NOMBRE PROFESIONAL",
        "OMISIONES"
    ]]
    # =========================
    # BASE LIMPIA NO MÉDICOS (SIN CONTAMINACIÓN)
    # =========================
    
    df_no_medicos_base = hoja1[
        hoja1[col_h1_estado].astype(str).str.upper() == "ASIGNADA"
    ].copy()
    
    agrupaciones_no_medicas = [
        "TERAPEUTA OCUPACIONAL",
        "PSICOLOGIA",
        "ENFERMERA(O)",
        "ASISTENTE SOCIAL",
        "NUTRICIONISTA",
        "TECNOLOGO MEDICO",
        "FONOAUDIOLOGO",
        "MATRON(A)",
        "KINESIOLOGO",
        "PROCEDIMIENTO"
    ]
    
    # =========================
    # FILTRAR AGRUPACIONES NO MÉDICAS
    # =========================
    
    df_no_medicos = df_no_medicos_base[
        df_no_medicos_base[col_h1_agr].astype(str).str.upper().isin(agrupaciones_no_medicas)
    ].copy()
    
    # =========================
    # RESTAR MÉDICOS SOLO EN PROCEDIMIENTO
    # =========================
    
    medicos_hoja2 = set(
        hoja2[col_h2_prof].astype(str).str.strip()
    )
    
    df_no_medicos = df_no_medicos[
        ~(
            (df_no_medicos[col_h1_agr].astype(str).str.upper() == "PROCEDIMIENTO") &
            (df_no_medicos[col_h1_prof].astype(str).str.strip().isin(medicos_hoja2))
        )
    ].copy()
    
    # =========================
    # OMISIONES
    # =========================
    
    df_no_medicos["OMISIONES"] = 1
    
    # =========================
    # TABLA 4 - RESUMEN POR POLICLINICO
    # =========================
    
    tabla4 = (
        df_no_medicos.groupby("POLICLINICO")
        .size()
        .reset_index(name="TOTAL")
        .sort_values("TOTAL", ascending=False)
    )
    
    # st.subheader("Tabla 4 - Resumen No Médicos por Policlínico")
    # st.dataframe(tabla4, use_container_width=True)
    
    orden_poli = tabla4["POLICLINICO"].tolist()
    
    # =========================
    # RANKING PROFESIONALES
    # =========================
    
    ranking_prof_nm = (
        df_no_medicos.groupby(["POLICLINICO", col_h1_prof])
        .size()
        .reset_index(name="TOTAL OMISIONES")
    )
    
    ranking_prof_nm["ORDEN_POLI"] = ranking_prof_nm["POLICLINICO"].apply(
        lambda x: orden_poli.index(x) if x in orden_poli else 999
    )
    
    ranking_prof_nm = ranking_prof_nm.sort_values(
        by=["ORDEN_POLI", "TOTAL OMISIONES"],
        ascending=[True, False]
    )
    
    # =========================
    # TABLA 5 - DETALLE
    # =========================
    
    tabla5 = df_no_medicos.copy()
    
    tabla5.rename(columns={
        col_h1_prof: "NOMBRE PROFESIONAL"
    }, inplace=True)
    
    tabla5 = tabla5.merge(
        ranking_prof_nm,
        left_on=["POLICLINICO", "NOMBRE PROFESIONAL"],
        right_on=["POLICLINICO", col_h1_prof],
        how="left"
    )
    
    tabla5["ORDEN_POLI"] = tabla5["POLICLINICO"].apply(
        lambda x: orden_poli.index(x) if x in orden_poli else 999
    )
    
    tabla5 = tabla5.sort_values(
        by=["ORDEN_POLI", "TOTAL OMISIONES", "NOMBRE PROFESIONAL"],
        ascending=[True, False, True]
    )
    
    for col in ["RUT PACIENTE", "NOMBRE PACIENTE", "FECHA"]:
        if col not in tabla5.columns:
            tabla5[col] = None
    
    tabla5 = tabla5[
        [
            "POLICLINICO",
            "NOMBRE PROFESIONAL",
            "RUT PACIENTE",
            "NOMBRE PACIENTE",
            "FECHA",
            "OMISIONES"
        ]
    ]
    
    # st.subheader("Tabla 5 - Detalle No Médicos")
    # st.dataframe(tabla5, use_container_width=True)
    
    # =========================
    # TABLA 6 - RANKING GLOBAL
    # =========================
    
    tabla6 = (
        df_no_medicos.groupby(col_h1_prof)
        .size()
        .reset_index(name="OMISIONES")
    )
    
    tabla6.rename(columns={
        col_h1_prof: "NOMBRE PROFESIONAL"
    }, inplace=True)
    
    policlinicos_prof = (
        df_no_medicos.groupby(col_h1_prof)["POLICLINICO"]
        .apply(lambda x: ", ".join(sorted(set(x.astype(str)))))
        .reset_index()
    )
    
    policlinicos_prof.rename(columns={
        col_h1_prof: "NOMBRE PROFESIONAL",
        "POLICLINICO": "POLICLINICOS"
    }, inplace=True)
    
    tabla6 = tabla6.merge(
        policlinicos_prof,
        on="NOMBRE PROFESIONAL",
        how="left"
    )
    
    tabla6 = tabla6[
        [
            "NOMBRE PROFESIONAL",
            "POLICLINICOS",
            "OMISIONES"
        ]
    ]
    
    tabla6 = tabla6.sort_values(
        "OMISIONES",
        ascending=False
    )
    
    # st.subheader("Tabla 6 - Ranking Profesionales No Médicos")
    # st.dataframe(tabla6, use_container_width=True)
       
    
        # =========================
        # MOSTRAR TABLAS
        # =========================
    # st.subheader("Tabla 2 - Detalle Ordenado")
    # st.dataframe(tabla2, use_container_width=True)
    # st.subheader("Tabla 3 - Resumen Médicos")
    # st.dataframe(tabla3, use_container_width=True)
    st.markdown("## 📊 Resumen General del Período")

    col1, col2, col3 = st.columns(3)
    
    col1.metric(
        "Total Asignadas",
        len(df_asignadas)
    )
    
    col2.metric(
        "Omisiones Ley Médica",
        len(df_medicos)
    )
    
    col3.metric(
        "Omisiones Ley 18",
        len(df_no_medicos)
    )
    st.markdown("## 🧑‍⚕️ Top Ley Médica - Mayores Omisiones")

    tabla_ley_medica = (
        df_medicos.groupby("ESPECIALIDAD_FINAL")
        .size()
        .reset_index(name="OMISIONES")
        .sort_values("OMISIONES", ascending=False)
        .head(5)
    )

    st.dataframe(tabla_ley_medica, use_container_width=True)

    st.markdown("## 🧑‍⚕️ Top 5 Médicos con más Omisiones")

    tabla_top_medicos = (
        df_medicos.groupby(col_h1_prof)
        .size()
        .reset_index(name="OMISIONES")
        .rename(columns={col_h1_prof: "NOMBRE PROFESIONAL"})
        .sort_values("OMISIONES", ascending=False)
        .head(5)
    )

    st.dataframe(tabla_top_medicos, use_container_width=True)

    st.markdown("## 🏥 Top 5 Policlínicos - Ley 18 (más omisiones)")

    # =========================
    # TABLA
    # =========================
    tabla_poli_18 = (
        df_no_medicos.groupby("POLICLINICO")
        .size()
        .reset_index(name="OMISIONES")
        .sort_values("OMISIONES", ascending=False)
        .head(5)
    )

    st.dataframe(tabla_poli_18, use_container_width=True)

    # =========================
    # GRÁFICO
    # =========================
    st.markdown("### 📊 Visualización")

    st.bar_chart(tabla_poli_18.set_index("POLICLINICO"))

    st.markdown("## 🏥 Top Ley 18 - Mayores Omisiones")

    tabla_ley18 = (
        df_no_medicos.groupby(col_h1_prof)
        .size()
        .reset_index(name="OMISIONES")
        .rename(columns={col_h1_prof: "NOMBRE PROFESIONAL"})
        .sort_values("OMISIONES", ascending=False)
        .head(5)
    )

    st.dataframe(tabla_ley18, use_container_width=True)
    
        # =========================
        # EXPORT EXCEL
        # =========================
    salida = BytesIO()
    
    with pd.ExcelWriter(salida, engine="xlsxwriter") as writer:

        tabla.to_excel(writer, sheet_name="Resumen", index=False)
        tabla2.to_excel(writer, sheet_name="Detalle Omisiones", index=False)
        tabla3.to_excel(writer, sheet_name="Resumen Medicos", index=False)
        tabla4.to_excel(writer, sheet_name="Resumen No Medicos", index=False)
        tabla5.to_excel(writer, sheet_name="Detalle No Medicos", index=False)
        tabla6.to_excel(writer, sheet_name="Ranking No Medicos", index=False)
        
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
