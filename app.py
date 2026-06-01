import os
from datetime import date, datetime
from io import BytesIO

import pandas as pd
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet


# =====================================================
# CONFIGURACIÓN
# =====================================================
st.set_page_config(
    page_title="Calculadora de Finiquito LFT México",
    page_icon="⚖️",
    layout="centered"
)

CARPETA_DATOS = "datos_finiquitos"
HISTORIAL_EXCEL = os.path.join(CARPETA_DATOS, "historial_finiquitos.xlsx")
os.makedirs(CARPETA_DATOS, exist_ok=True)


# =====================================================
# DISEÑO
# =====================================================
st.markdown("""
<style>
.stApp {
    background:
        radial-gradient(circle at top left, rgba(8,123,117,0.28), transparent 30%),
        radial-gradient(circle at bottom right, rgba(233,78,27,0.38), transparent 34%),
        linear-gradient(135deg, #EEF8F5 0%, #FFF7E7 50%, #F8C2A5 100%);
}

.block-container {
    padding-top: 30px;
    max-width: 1050px;
}

.title-card {
    background: linear-gradient(135deg, rgba(219,246,241,0.96), rgba(255,242,216,0.96));
    padding: 28px;
    border-radius: 22px;
    box-shadow: 0px 8px 24px rgba(0,0,0,0.13);
    text-align: center;
    margin-bottom: 25px;
    border: 1px solid rgba(255,255,255,0.85);
}

.title-card h1 {
    color: #087B75;
    font-weight: 900;
    margin-bottom: 5px;
}

.notice-card {
    background: white;
    padding: 22px;
    border-radius: 18px;
    box-shadow: 0px 5px 15px rgba(0,0,0,0.10);
    border-left: 8px solid #E87522;
    margin-bottom: 22px;
    color: #1f2937;
}

.form-card {
    background: rgba(255,255,255,0.62);
    padding: 24px;
    border-radius: 20px;
    border: 1px solid rgba(0,0,0,0.12);
    box-shadow: 0px 5px 15px rgba(0,0,0,0.07);
}

.result-card {
    background: linear-gradient(135deg, #E4F7F3, #FFF7EA);
    padding: 25px;
    border-radius: 20px;
    box-shadow: 0px 6px 18px rgba(0,0,0,0.12);
    margin-top: 20px;
    border-left: 8px solid #087B75;
}

.total-card {
    background: linear-gradient(90deg, #087B75, #14A39A);
    color: white;
    padding: 22px;
    border-radius: 18px;
    text-align: center;
    margin-top: 20px;
    box-shadow: 0px 6px 16px rgba(0,0,0,0.16);
}

.stButton > button {
    background: linear-gradient(90deg, #E94E1B, #F2B233);
    color: white;
    border-radius: 14px;
    border: none;
    padding: 14px;
    font-size: 18px;
    font-weight: 900;
    width: 100%;
}

.stDownloadButton > button {
    background: linear-gradient(90deg, #087B75, #14A39A);
    color: white;
    border-radius: 14px;
    border: none;
    padding: 12px;
    font-weight: 800;
    width: 100%;
}
</style>
""", unsafe_allow_html=True)


# =====================================================
# FUNCIONES
# =====================================================
def formato_moneda(valor):
    return f"${valor:,.2f}"


def calcular_antiguedad(fecha_ingreso, fecha_salida):
    dias = (fecha_salida - fecha_ingreso).days
    anios = dias / 365 if dias > 0 else 0
    return dias, anios


def dias_vacaciones_por_antiguedad(anios_trabajados):
    anios_enteros = max(1, int(anios_trabajados) if anios_trabajados >= 1 else 1)

    if anios_enteros == 1:
        return 12
    if anios_enteros == 2:
        return 14
    if anios_enteros == 3:
        return 16
    if anios_enteros == 4:
        return 18
    if anios_enteros == 5:
        return 20

    bloque = ((anios_enteros - 6) // 5) + 1
    return 20 + (bloque * 2)


def meses_entre_fechas(fecha_inicio, fecha_fin):
    if fecha_fin < fecha_inicio:
        return 0

    meses = (fecha_fin.year - fecha_inicio.year) * 12 + (fecha_fin.month - fecha_inicio.month)

    if fecha_fin.day >= fecha_inicio.day:
        meses += 1

    return max(0, meses)


def calcular_fondo_extra(
    tiene_fondo,
    aportacion_trabajador_mensual,
    aportacion_patron_mensual,
    fecha_inicio_fondo,
    fecha_fin_fondo,
    rendimiento_anual_pct
):
    if not tiene_fondo:
        return {
            "meses_fondo": 0,
            "aportacion_trabajador": 0.0,
            "aportacion_patron": 0.0,
            "rendimiento_estimado": 0.0,
            "total_fondo": 0.0
        }

    meses = meses_entre_fechas(fecha_inicio_fondo, fecha_fin_fondo)
    aport_trab = aportacion_trabajador_mensual * meses
    aport_patron = aportacion_patron_mensual * meses
    capital = aport_trab + aport_patron

    anios = meses / 12
    rendimiento = capital * (rendimiento_anual_pct / 100) * anios

    return {
        "meses_fondo": meses,
        "aportacion_trabajador": aport_trab,
        "aportacion_patron": aport_patron,
        "rendimiento_estimado": rendimiento,
        "total_fondo": capital + rendimiento
    }


def calcular_finiquito(
    fecha_ingreso,
    fecha_salida,
    salario_diario,
    salario_integrado,
    tipo_salida,
    dias_pendientes_pago,
    dias_aguinaldo_anual,
    prima_vacacional_pct,
    salario_minimo_diario,
    incluir_20_dias,
    salarios_caidos,
    otras_prestaciones
):
    dias_trabajados, anios_trabajados = calcular_antiguedad(fecha_ingreso, fecha_salida)
    dias_vacaciones = dias_vacaciones_por_antiguedad(anios_trabajados)

    dia_anio_salida = fecha_salida.timetuple().tm_yday

    salarios_pendientes = salario_diario * dias_pendientes_pago
    aguinaldo_proporcional = salario_diario * dias_aguinaldo_anual * (dia_anio_salida / 365)

    dias_ultimo_anio = dias_trabajados % 365
    if dias_trabajados > 0 and dias_ultimo_anio == 0:
        dias_ultimo_anio = 365

    vacaciones_proporcionales = salario_diario * dias_vacaciones * (dias_ultimo_anio / 365)
    prima_vacacional = vacaciones_proporcionales * (prima_vacacional_pct / 100)

    indemnizacion_3_meses = 0.0
    veinte_dias_por_anio = 0.0
    prima_antiguedad = 0.0

    aplica_prima_antiguedad = False

    if tipo_salida == "Despido injustificado":
        indemnizacion_3_meses = salario_integrado * 90
        aplica_prima_antiguedad = True

        if incluir_20_dias:
            veinte_dias_por_anio = salario_integrado * 20 * anios_trabajados

    elif tipo_salida == "Renuncia voluntaria":
        if anios_trabajados >= 15:
            aplica_prima_antiguedad = True

    elif tipo_salida == "Despido justificado":
        aplica_prima_antiguedad = True

    elif tipo_salida == "Jubilación":
        aplica_prima_antiguedad = True

    elif tipo_salida in ["Terminación de contrato", "Mutuo acuerdo"]:
        aplica_prima_antiguedad = False

    if aplica_prima_antiguedad:
        if salario_minimo_diario > 0:
            salario_base_prima = min(salario_diario, salario_minimo_diario * 2)
        else:
            salario_base_prima = salario_diario

        prima_antiguedad = salario_base_prima * 12 * anios_trabajados

    subtotal_lft = (
        salarios_pendientes +
        aguinaldo_proporcional +
        vacaciones_proporcionales +
        prima_vacacional +
        indemnizacion_3_meses +
        veinte_dias_por_anio +
        prima_antiguedad +
        salarios_caidos +
        otras_prestaciones
    )

    desglose = [
        ("Salarios pendientes", salarios_pendientes),
        ("Aguinaldo proporcional", aguinaldo_proporcional),
        ("Vacaciones proporcionales", vacaciones_proporcionales),
        ("Prima vacacional", prima_vacacional),
        ("Indemnización 3 meses", indemnizacion_3_meses),
        ("20 días por año trabajado", veinte_dias_por_anio),
        ("Prima de antigüedad", prima_antiguedad),
        ("Salarios caídos / vencidos capturados", salarios_caidos),
        ("Otras prestaciones capturadas", otras_prestaciones),
    ]

    return {
        "dias_trabajados": dias_trabajados,
        "anios_trabajados": anios_trabajados,
        "dias_vacaciones": dias_vacaciones,
        "desglose": desglose,
        "subtotal_lft": subtotal_lft
    }


def guardar_historial(fila):
    df_nuevo = pd.DataFrame([fila])

    if os.path.exists(HISTORIAL_EXCEL):
        try:
            df_actual = pd.read_excel(HISTORIAL_EXCEL)
            df_final = pd.concat([df_actual, df_nuevo], ignore_index=True)
        except Exception:
            df_final = df_nuevo
    else:
        df_final = df_nuevo

    df_final.to_excel(HISTORIAL_EXCEL, index=False)


def crear_excel_resultado(desglose_df, resumen):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        desglose_df.to_excel(writer, index=False, sheet_name="Desglose")
        pd.DataFrame([resumen]).to_excel(writer, index=False, sheet_name="Resumen")

    output.seek(0)
    return output


def crear_pdf(nombre, resumen, desglose_df):
    output = BytesIO()
    doc = SimpleDocTemplate(
        output,
        pagesize=letter,
        rightMargin=35,
        leftMargin=35,
        topMargin=35,
        bottomMargin=35
    )

    styles = getSampleStyleSheet()
    story = []

    story.append(Paragraph("Cálculo estimado de finiquito / liquidación", styles["Title"]))
    story.append(Spacer(1, 0.15 * inch))
    story.append(Paragraph("Aviso: Este documento es informativo y no sustituye asesoría legal.", styles["Normal"]))
    story.append(Spacer(1, 0.2 * inch))

    datos = [
        ["Trabajador", nombre or "No capturado"],
        ["Fecha de cálculo", datetime.now().strftime("%Y-%m-%d %H:%M")],
        ["Tipo de salida", resumen.get("Tipo de salida", "")],
        ["Antigüedad", resumen.get("Antigüedad", "")],
        ["Total bruto estimado", formato_moneda(resumen.get("Total bruto estimado", 0))],
        ["ISR / descuento estimado", formato_moneda(resumen.get("ISR estimado", 0))],
        ["Total neto estimado", formato_moneda(resumen.get("Total neto estimado", 0))],
    ]

    tabla_datos = Table(datos, colWidths=[2.2 * inch, 4.5 * inch])
    tabla_datos.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#E4F7F3")),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("PADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(tabla_datos)
    story.append(Spacer(1, 0.25 * inch))

    data = [["Concepto", "Importe"]]

    for _, row in desglose_df.iterrows():
        data.append([str(row["Concepto"]), str(row["Importe"])])

    tabla = Table(data, colWidths=[4.5 * inch, 2.2 * inch])
    tabla.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#087B75")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("PADDING", (0, 0), (-1, -1), 6),
    ]))

    story.append(tabla)
    doc.build(story)

    output.seek(0)
    return output


# =====================================================
# ENCABEZADO
# =====================================================
st.markdown("""
<div class="title-card">
<h1>⚖️ Calculadora de Finiquito y Liquidación</h1>
<p>Basada en conceptos generales de la Ley Federal del Trabajo de México</p>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="notice-card">
<b>Nota importante:</b><br>
Este sistema genera un cálculo estimado informativo. No sustituye asesoría legal, resolución de autoridad laboral, revisión de contrato, prestaciones superiores, sindicato, convenio colectivo ni criterios de autoridad.
</div>
""", unsafe_allow_html=True)


# =====================================================
# FORMULARIO
# =====================================================
st.markdown('<div class="form-card">', unsafe_allow_html=True)

with st.form("form_calculo"):
    st.subheader("Datos del trabajador")

    col1, col2 = st.columns(2)

    with col1:
        nombre = st.text_input("Nombre del trabajador")
        fecha_ingreso = st.date_input("Fecha de ingreso")
        salario_diario = st.number_input("Salario diario", min_value=0.0, step=10.0)
        salario_integrado = st.number_input(
            "Salario diario integrado",
            min_value=0.0,
            step=10.0,
            help="Si no lo conoces, puedes poner el mismo salario diario."
        )
        dias_pendientes_pago = st.number_input("Días pendientes de salario", min_value=0.0, step=1.0)

    with col2:
        fecha_salida = st.date_input("Fecha de salida", value=date.today())
        tipo_salida = st.selectbox(
            "Tipo de salida",
            [
                "Renuncia voluntaria",
                "Despido injustificado",
                "Despido justificado",
                "Terminación de contrato",
                "Mutuo acuerdo",
                "Jubilación"
            ]
        )
        dias_aguinaldo_anual = st.number_input("Días de aguinaldo al año", min_value=15.0, value=15.0, step=1.0)
        prima_vacacional_pct = st.number_input("Prima vacacional %", min_value=25.0, value=25.0, step=1.0)

    st.subheader("Opciones legales y adicionales")

    col3, col4 = st.columns(2)

    with col3:
        salario_minimo_diario = st.number_input(
            "Salario mínimo diario para tope de prima de antigüedad",
            min_value=0.0,
            value=0.0,
            step=10.0,
            help="Si lo dejas en 0, no se aplicará tope."
        )
        incluir_20_dias = st.checkbox(
            "Incluir 20 días por año trabajado",
            value=False,
            help="Actívalo solo cuando aplique según el caso concreto."
        )
        salarios_caidos = st.number_input("Salarios caídos / vencidos capturados", min_value=0.0, value=0.0, step=100.0)

    with col4:
        otras_prestaciones = st.number_input("Otras prestaciones o adeudos", min_value=0.0, value=0.0, step=100.0)
        oferta_empresa = st.number_input("¿Cuánto te ofrece la empresa?", min_value=0.0, value=0.0, step=100.0)
        isr_estimado = st.number_input("ISR / descuento estimado", min_value=0.0, value=0.0, step=100.0)

    st.subheader("Fondo adicional por retiro, jubilación o despido")

    tiene_fondo = st.checkbox("Tiene fondo adicional / aportación extra", value=False)

    col5, col6 = st.columns(2)

    with col5:
        fecha_inicio_fondo = st.date_input("Fecha inicio del fondo", value=fecha_ingreso)
        aportacion_trabajador_mensual = st.number_input(
            "Aportación mensual del trabajador",
            min_value=0.0,
            value=0.0,
            step=50.0
        )

    with col6:
        fecha_fin_fondo = st.date_input("Fecha fin del fondo", value=fecha_salida)
        aportacion_patron_mensual = st.number_input(
            "Aportación mensual del patrón / empresa",
            min_value=0.0,
            value=0.0,
            step=50.0
        )

    rendimiento_anual_pct = st.number_input(
        "Rendimiento anual estimado del fondo %",
        min_value=0.0,
        value=0.0,
        step=0.5
    )

    calcular = st.form_submit_button("🧮 Calcular finiquito / liquidación")

st.markdown("</div>", unsafe_allow_html=True)


# =====================================================
# RESULTADOS
# =====================================================
if calcular:
    if fecha_salida < fecha_ingreso:
        st.error("La fecha de salida no puede ser anterior a la fecha de ingreso.")
    elif salario_diario <= 0:
        st.error("Captura un salario diario válido.")
    else:
        if salario_integrado <= 0:
            salario_integrado = salario_diario

        resultado = calcular_finiquito(
            fecha_ingreso=fecha_ingreso,
            fecha_salida=fecha_salida,
            salario_diario=salario_diario,
            salario_integrado=salario_integrado,
            tipo_salida=tipo_salida,
            dias_pendientes_pago=dias_pendientes_pago,
            dias_aguinaldo_anual=dias_aguinaldo_anual,
            prima_vacacional_pct=prima_vacacional_pct,
            salario_minimo_diario=salario_minimo_diario,
            incluir_20_dias=incluir_20_dias,
            salarios_caidos=salarios_caidos,
            otras_prestaciones=otras_prestaciones
        )

        fondo = calcular_fondo_extra(
            tiene_fondo=tiene_fondo,
            aportacion_trabajador_mensual=aportacion_trabajador_mensual,
            aportacion_patron_mensual=aportacion_patron_mensual,
            fecha_inicio_fondo=fecha_inicio_fondo,
            fecha_fin_fondo=fecha_fin_fondo,
            rendimiento_anual_pct=rendimiento_anual_pct
        )

        subtotal_lft = resultado["subtotal_lft"]
        total_fondo = fondo["total_fondo"]
        total_bruto = subtotal_lft + total_fondo
        total_neto = max(0, total_bruto - isr_estimado)
        diferencia_empresa = total_bruto - oferta_empresa if oferta_empresa > 0 else 0

        filas = []

        for concepto, importe in resultado["desglose"]:
            filas.append({"Concepto": concepto, "Importe": importe})

        filas.extend([
            {"Concepto": "Fondo: aportación trabajador", "Importe": fondo["aportacion_trabajador"]},
            {"Concepto": "Fondo: aportación patrón / empresa", "Importe": fondo["aportacion_patron"]},
            {"Concepto": "Fondo: rendimiento estimado", "Importe": fondo["rendimiento_estimado"]},
            {"Concepto": "Total fondo adicional", "Importe": total_fondo},
            {"Concepto": "Subtotal prestaciones LFT", "Importe": subtotal_lft},
            {"Concepto": "TOTAL BRUTO ESTIMADO", "Importe": total_bruto},
            {"Concepto": "ISR / descuento estimado", "Importe": isr_estimado},
            {"Concepto": "TOTAL NETO ESTIMADO", "Importe": total_neto},
        ])

        if oferta_empresa > 0:
            filas.append({"Concepto": "Oferta de la empresa", "Importe": oferta_empresa})
            filas.append({"Concepto": "Diferencia contra cálculo bruto", "Importe": diferencia_empresa})

        df_desglose = pd.DataFrame(filas)
        df_mostrar = df_desglose.copy()
        df_mostrar["Importe"] = df_mostrar["Importe"].apply(formato_moneda)

        antiguedad_txt = f'{resultado["dias_trabajados"]} días / {resultado["anios_trabajados"]:.2f} años'

        resumen = {
            "Nombre": nombre,
            "Tipo de salida": tipo_salida,
            "Fecha ingreso": str(fecha_ingreso),
            "Fecha salida": str(fecha_salida),
            "Antigüedad": antiguedad_txt,
            "Salario diario": salario_diario,
            "Salario integrado": salario_integrado,
            "Días de vacaciones considerados": resultado["dias_vacaciones"],
            "Meses de fondo": fondo["meses_fondo"],
            "Total bruto estimado": total_bruto,
            "ISR estimado": isr_estimado,
            "Total neto estimado": total_neto,
            "Oferta empresa": oferta_empresa,
            "Diferencia": diferencia_empresa
        }

        guardar_historial({
            "Fecha cálculo": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            **resumen
        })

        st.markdown('<div class="result-card">', unsafe_allow_html=True)
        st.subheader("Resultado del cálculo")

        if nombre:
            st.write(f"**Trabajador:** {nombre.upper()}")

        st.write(f"**Tipo de salida:** {tipo_salida}")
        st.write(f"**Antigüedad:** {antiguedad_txt}")
        st.write(f"**Vacaciones anuales consideradas:** {resultado['dias_vacaciones']} días")
        st.write(f"**Meses considerados para fondo adicional:** {fondo['meses_fondo']}")

        colr1, colr2, colr3 = st.columns(3)
        colr1.metric("Subtotal LFT", formato_moneda(subtotal_lft))
        colr2.metric("Fondo adicional", formato_moneda(total_fondo))
        colr3.metric("Neto estimado", formato_moneda(total_neto))

        st.dataframe(df_mostrar, use_container_width=True)

        if oferta_empresa > 0:
            if diferencia_empresa > 0:
                st.warning(f"La oferta de la empresa está por debajo del cálculo bruto por {formato_moneda(diferencia_empresa)}.")
            elif diferencia_empresa < 0:
                st.success(f"La oferta de la empresa está por encima del cálculo bruto por {formato_moneda(abs(diferencia_empresa))}.")
            else:
                st.info("La oferta de la empresa coincide con el cálculo bruto estimado.")

        st.markdown(f"""
        <div class="total-card">
            <h2>Total bruto estimado: {formato_moneda(total_bruto)}</h2>
            <h2>Total neto estimado: {formato_moneda(total_neto)}</h2>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("</div>", unsafe_allow_html=True)

        excel = crear_excel_resultado(df_desglose, resumen)
        pdf = crear_pdf(nombre, resumen, df_mostrar)

        col_down1, col_down2 = st.columns(2)

        with col_down1:
            st.download_button(
                "📥 Descargar Excel",
                data=excel,
                file_name="calculo_finiquito_liquidacion.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        with col_down2:
            st.download_button(
                "📄 Descargar PDF",
                data=pdf,
                file_name="calculo_finiquito_liquidacion.pdf",
                mime="application/pdf"
            )

        if os.path.exists(HISTORIAL_EXCEL):
            with open(HISTORIAL_EXCEL, "rb") as f:
                st.download_button(
                    "📚 Descargar historial de cálculos",
                    data=f,
                    file_name="historial_finiquitos.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
