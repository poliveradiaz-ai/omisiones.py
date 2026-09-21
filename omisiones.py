import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import unicodedata
from io import BytesIO
import traceback

st.set_page_config(page_title="Reporte Omisiones", layout="wide")

st.title("📊 Generador de Reportes de Omisiones Médicas")

st.write("Sube los archivos base y opcionalmente las plantillas")

# =========================
# ARCHIVOS BASE
# =========================
st.markdown("## 📂 Archivos base")

lp_file = st.file_uploader("📄 Lista de Espera (.xlsx)", type=["xlsx"])
datos_file = st.file_uploader("📄 Datos RCE Especialidades (.xlsx)", type=["xlsx"])
medicos_file = st.file_uploader("📄 Nómina Médicos (.xlsx)", type=["xlsx"])

st.divider()

# =========================
# FUNCIONES DE PROCESAMIENTO
# =========================
def normalizar_texto(texto):
    if pd.isna(texto):
        return ""
    texto = str(texto).strip().upper()
    texto = unicodedata.normalize('NFD', texto)
    texto = ''.join(c for c in texto if unicodedata.category(c) != 'Mn')
    return texto

def limpiar_rut_definitivo(rut):
    rut = str(rut).strip()
    if "-" in rut:
        rut = rut.split("-")[0]
    return rut.replace(".", "")

def agregar_totales_por_especialidad(tabla):
    resultado = []
    for especialidad, grupo in tabla.groupby('Especialidad', sort=False):
        for _, fila in grupo.iterrows():
            resultado.append({
                'Especialidad': fila['Especialidad'],
                'Funcionario': fila['Funcionario'],
                'total': fila['total'],
                'es_total': False
            })
        
        total_especialidad = grupo['total'].sum()
        resultado.append({
            'Especialidad': f'TOTAL {especialidad}',
            'Funcionario': '',
            'total': total_especialidad,
            'es_total': True
        })

    total_general = tabla['total'].sum()
    resultado.append({
        'Especialidad': 'TOTAL GENERAL',
        'Funcionario': '',
        'total': total_general,
        'es_total': True
    })

    return resultado

def agregar_total_general(tabla, columna_total):
    resultado = tabla.to_dict('records')
    total_general = tabla[columna_total].sum()
    resultado.append({
        'Especialidad': 'TOTAL GENERAL',
        columna_total: total_general
    })
    return resultado

# =========================
# PLANTILLAS WORD
# =========================
col1, col2 = st.columns(2)

with col1:
    st.markdown("📄 Plantilla Informe 1")
    word_file = st.file_uploader("Plantilla 1 (.docx)", type=["docx"], key="w1")

with col2:
    st.markdown("📄 Plantilla Informe 2")
    preliminar2_word_file = st.file_uploader("Plantilla 2 (.docx)", type=["docx"], key="w2")

st.markdown("## 📅 Fechas del informe")

col_fecha1, col_fecha2, col_fecha3 = st.columns(3)

with col_fecha1:
    fecha_corte = st.date_input("Fecha de corte")

with col_fecha2:
    fecha_inf_preliminar = st.date_input("Fecha envío informe preliminar")

with col_fecha3:
    incluir_fecha_envio_final = st.checkbox("¿Ingresar fecha de envío del informe final?")
    if incluir_fecha_envio_final:
        fecha_envio_informe_final = st.date_input("Fecha de envío del informe final")
    else:
        fecha_envio_informe_final = None

# =========================
# INIT SESSION STATE
# =========================
if "informe1" not in st.session_state:
    st.session_state["informe1"] = None

if "informe2" not in st.session_state:
    st.session_state["informe2"] = None

if "reporte_excel" not in st.session_state:
    st.session_state["reporte_excel"] = None

# =========================
# BOTÓN GENERAR REPORTE
# =========================
if st.button("🚀 Generar Reporte"):

    if lp_file and datos_file and medicos_file:
        try:
            # 1. Cargar datos desde los Excel
            lp = pd.read_excel(lp_file, sheet_name="SIGTE_Salida")
            rce = pd.read_excel(datos_file, sheet_name="NOMINA CUADRATURA (REM7) SIN CO")
            medicos = pd.read_excel(medicos_file, sheet_name="Nomina Médico")

            # Renombrar columnas RCE
            if 'Rut' in rce.columns and 'Rut.1' in rce.columns:
                rce = rce.rename(columns={'Rut': 'Rut Paciente', 'Rut.1': 'Rut Funcionario'})

            # Construir nombre completo del funcionario
            rce['Funcionario'] = (
                rce['Nombres'].fillna('').astype(str) + ' ' +
                rce['Apellido Pat'].fillna('').astype(str) + ' ' +
                rce['Apellido Mat'].fillna('').astype(str)
            ).str.replace(r'\s+', ' ', regex=True).str.strip()

            # Normalizaciones para cruces
            lp['rut_puente'] = lp['RUN/RUT_PACIENTE'].apply(limpiar_rut_definitivo)
            rce['rut_puente'] = rce['Rut Paciente'].apply(limpiar_rut_definitivo)

            lp['esp_puente'] = lp['ESPECIALIDAD_DESTINO'].apply(normalizar_texto)
            rce['esp_puente'] = rce['Especialidad'].apply(normalizar_texto)

            rce['actividad_norm'] = rce['Actividad'].apply(normalizar_texto)

            # Filtrar RCE por CONSULTA NUEVA
            rce_cn = rce[rce['actividad_norm'] == 'CONSULTA NUEVA'].copy()

            # Merge Lista de Espera x RCE
            lp_merged = pd.merge(
                lp,
                rce_cn[['rut_puente', 'esp_puente', 'Rut Funcionario', 'Funcionario', 'Fecha Atencion']],
                on=['rut_puente', 'esp_puente'],
                how='left'
            )

            # Identificar Omisiones (donde no hubo atención en RCE)
            omisiones = lp_merged[lp_merged['Rut Funcionario'].isna()].copy()

            # Merge con Nómina de Médicos
            medicos['esp_puente'] = medicos['Especialidad Destino'].apply(normalizar_texto)
            medicos['rut_doc_puente'] = medicos['Rut'].apply(limpiar_rut_definitivo)

            omisiones = pd.merge(
                omisiones,
                medicos[['esp_puente', 'rut_doc_puente', 'Nombre Profesional']],
                on='esp_puente',
                how='left'
            )

            # 2. Resúmenes y métricas
            total_omisiones = len(omisiones)

            tabla_omisiones_esp = (
                omisiones.groupby('ESPECIALIDAD_DESTINO')
                .size()
                .reset_index(name='cantidad')
                .sort_values(by='cantidad', ascending=False)
                .rename(columns={'ESPECIALIDAD_DESTINO': 'Especialidad'})
            )

            tabla_omisiones_func = (
                omisiones.groupby(['ESPECIALIDAD_DESTINO', 'Nombre Profesional'])
                .size()
                .reset_index(name='total')
                .rename(columns={'ESPECIALIDAD_DESTINO': 'Especialidad', 'Nombre Profesional': 'Funcionario'})
                .sort_values(by=['Especialidad', 'total'], ascending=[True, False])
            )

            # Tablas estructuradas para Word
            tabla_omisiones_word = agregar_total_general(tabla_omisiones_esp, 'cantidad')
            tabla_funcionarios_omisiones_word = agregar_totales_por_especialidad(tabla_omisiones_func)

            # Formateo de fechas y meses
            fecha_corte_str = fecha_corte.strftime("%d/%m/%Y")
            fecha_inf_preliminar_str = fecha_inf_preliminar.strftime("%d/%m/%Y")
            fecha_envio_informe_final_str = (
                fecha_envio_informe_final.strftime("%d/%m/%Y")
                if fecha_envio_informe_final is not None else ""
            )

            meses = {
                1: "enero", 2: "febrero", 3: "marzo", 4: "abril", 5: "mayo", 6: "junio",
                7: "julio", 8: "agosto", 9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"
            }
            mes_corte = meses[fecha_corte.month]

            # =========================
            # 📊 GENERAR EXCEL CONSOLIDADO
            # =========================
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                tabla_omisiones_esp.to_excel(writer, sheet_name='OMISIONES_Especialidad', index=False)
                tabla_omisiones_func.to_excel(writer, sheet_name='OMISIONES_Funcionario', index=False)
                omisiones.to_excel(writer, sheet_name='OMISIONES_Detalle', index=False)

            output.seek(0)
            st.session_state["reporte_excel"] = output.read()

            # =========================
            # 📝 RENDERIZAR PLANTILLAS WORD (DocxTemplate)
            # =========================
            doc = DocxTemplate(word_file) if word_file else None
            doc2 = DocxTemplate(preliminar2_word_file) if preliminar2_word_file else None

            contexto = {
                'total_omisiones': total_omisiones,
                'tabla_omisiones': tabla_omisiones_word,
                'tabla_funcionarios_omisiones': tabla_funcionarios_omisiones_word,
                'filas_omisiones': omisiones.to_dict('records'),
                'fecha_corte': fecha_corte_str,
                'fecha_envio_preliminar': fecha_inf_preliminar_str,
                'fecha_envio_informe_final': fecha_envio_informe_final_str,
                'mes_corte': mes_corte,
            }

            if doc:
                doc.render(contexto)
                buffer1 = BytesIO()
                doc.save(buffer1)
                st.session_state["informe1"] = buffer1.getvalue()

            if doc2:
                doc2.render(contexto)
                buffer2 = BytesIO()
                doc2.save(buffer2)
                st.session_state["informe2"] = buffer2.getvalue()

            st.success("✅ Reporte de omisiones generado correctamente")

        except Exception as e:
            st.error(f"Error: {e}")
            st.code(traceback.format_exc())

    else:
        st.warning("Debes subir los tres archivos base (Lista de Espera, RCE y Nómina Médica)")

# =========================
# DESCARGAS
# =========================
st.divider()
st.subheader("📥 Descargar Informes")

colA, colB, colC = st.columns(3)

with colA:
    if st.session_state["informe1"]:
        st.download_button(
            "📥 Informe 1",
            data=st.session_state["informe1"],
            file_name="Informe_Omisiones_1.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key="dl_1"
        )

with colB:
    if st.session_state["informe2"]:
        st.download_button(
            "📥 Informe 2",
            data=st.session_state["informe2"],
            file_name="Informe_Omisiones_2.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key="dl_2"
        )

with colC:
    if st.session_state.get("reporte_excel"):
        st.download_button(
            "📥 Descargar Excel consolidado",
            data=st.session_state["reporte_excel"],
            file_name="Reporte_Omisiones.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_ex"
        )
