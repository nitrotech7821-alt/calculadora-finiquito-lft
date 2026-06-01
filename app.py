import streamlit as st
from datetime import date
from io import BytesIO
import pandas as pd

# =====================================================
# CONFIGURACIÓN
# =====================================================
st.set_page_config(
    page_title="Calculadora de Finiquito México",
    page_icon="⚖️",
    layout="centered"
)

# =====================================================
# DISEÑO
# =====================================================
st.markdown("""
<style>
.stApp {
    background: linear-gradient(135deg, #eef8f5 0%, #fff7e7 50%, #f8c2a5 100%);
}

.block-container {
    padding-top: 30px;
    max-width: 1000px;
}

.title-card {
    background: linear-gradient(135deg, #dff4f0, #fff1d6);
    padding: 28px;
    border-radius: 22px;
    box-shadow: 0px 8px 24px rgba(0,0,0,0.12);
    text-align: center;
    margin-bottom: 25px;
}

.title-card h1 {
    color: #087B75;
    font-weight: 900;
}

.info-card {
    background: white;
    padding: 22px;
    border-radius: 18px;
    box-shadow: 0px 5px 15px rgba(0,0,0,0.10);
    border-left: 8px solid #E87522;
    margin-bottom: 20px;
}

.result-card {
    background: linear-gradient(135deg, #e4f7f3, #fff7ea);
    padding: 25px;
    border-radius: 20px;
    box-shadow: 0px 6px 18px rgba(0,0,0,0.12);
    margin-top: 20px;
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
</style>
""", unsafe_allow_html=True)

# =====================================================
# FUNCIONES
# =====================================================
def calcular_anios(fecha_ingreso, fecha_salida):
    dias = (fecha_salida - fecha_ingreso).days
    anios = dias / 365
    return dias, anios


def dias_vacaciones_por_antiguedad(anios):
    anios_enteros = int(anios)

    if anios_enteros <= 0:
        return 12

    if anios_enteros == 1:
        return 12
    elif anios_enteros == 2:
        return 14
    elif anios_enteros == 3:
        return 16
    elif anios_enteros == 4:
        return 18
    elif anios_enteros == 5:
        return 20
    else:
        extra = ((anios_enteros - 6) // 5 + 1) * 2
        return 20 + extra


def formato_moneda(valor):
    return f"${valor:,.2f}"


def crear_excel(desglose):
    df = pd.DataFrame(desglose)
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Calculo")

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
<div class="info-card">
<b>Nota importante:</b><br>
Este sistema genera un cálculo estimado informativo. No sustituye asesoría legal, resolución de autoridad laboral ni revisión de contrato, prestaciones superiores o convenio colectivo.
</div>
""", unsafe_allow_html=True)

# =====================================================
# FORMULARIO
# =====================================================
with st.form("form_finiquito"):
    st.subheader("Datos del trabajador")

    col1, col2 = st.columns(2)

    with col1:
        nombre = st.text_input("Nombre del trabajador")
        fecha_ingreso = st.date_input("Fecha de ingreso")
        salario_diario = st.number_input("Salario diario", min_value=0.0, step=10.0)
        dias_pendientes_pago = st.number_input("Días pendientes de salario", min_value=0.0, step=1.0)

    with col2:
        fecha_salida = st.date_input("Fecha de salida", value=date.today())
        tipo_salida = st.selectbox(
            "Tipo de salida",
            [
                "Renuncia voluntaria",
                "Despido injustificado",
                "Terminación de contrato"
            ]
        )
        aguinaldo_anual = st.number_input("Días de aguinaldo al año", min_value=15.0, value=15.0, step=1.0)
        prima_vacacional_pct = st.number_input("Prima vacacional %", min_value=25.0, value=25.0, step=1.0)

    st.subheader("Opciones legales")

    salario_minimo_diario = st.number_input(
        "Salario mínimo diario vigente para tope de prima de antigüedad",
        min_value=0.0,
        value=0.0,
        step=10.0,
        help="Si lo dejas en 0, el sistema no aplicará tope."
    )

    incluir_20_dias = st.checkbox(
        "Incluir 20 días por año trabajado",
        value=False,
        help="Este concepto puede depender del caso concreto. Actívalo solo cuando aplique."
    )

    calcular = st.form_submit_button("🧮 Calcular")

# =====================================================
# CÁLCULO
# =====================================================
if calcular:
    if salario_diario <= 0:
        st.error("Captura un salario diario válido.")
    elif fecha_salida < fecha_ingreso:
        st.error("La fecha de salida no puede ser anterior a la fecha de ingreso.")
    else:
        dias_trabajados, anios_trabajados = calcular_anios(fecha_ingreso, fecha_salida)

        dias_del_anio = 365
        dias_vacaciones_anuales = dias_vacaciones_por_antiguedad(anios_trabajados)

        salario_pendiente = salario_diario * dias_pendientes_pago

        aguinaldo_proporcional = (
            salario_diario * aguinaldo_anual * ((fecha_salida.timetuple().tm_yday) / dias_del_anio)
        )

        vacaciones_proporcionales = (
            salario_diario * dias_vacaciones_anuales * (dias_trabajados % 365) / dias_del_anio
        )

        prima_vacacional = vacaciones_proporcionales * (prima_vacacional_pct / 100)

        indemnizacion_3_meses = 0
        veinte_dias_por_anio = 0
        prima_antiguedad = 0

        aplica_prima_antiguedad = False

        if tipo_salida == "Despido injustificado":
            indemnizacion_3_meses = salario_diario * 90
            aplica_prima_antiguedad = True

            if incluir_20_dias:
                veinte_dias_por_anio = salario_diario * 20 * anios_trabajados

        elif tipo_salida == "Renuncia voluntaria":
            if anios_trabajados >= 15:
                aplica_prima_antiguedad = True

        elif tipo_salida == "Terminación de contrato":
            aplica_prima_antiguedad = False

        if aplica_prima_antiguedad:
            if salario_minimo_diario > 0:
                salario_base_prima = min(salario_diario, salario_minimo_diario * 2)
            else:
                salario_base_prima = salario_diario

            prima_antiguedad = salario_base_prima * 12 * anios_trabajados

        total = (
            salario_pendiente +
            aguinaldo_proporcional +
            vacaciones_proporcionales +
            prima_vacacional +
            indemnizacion_3_meses +
            veinte_dias_por_anio +
            prima_antiguedad
        )

        desglose = [
            {"Concepto": "Días trabajados totales", "Importe": dias_trabajados},
            {"Concepto": "Años trabajados", "Importe": round(anios_trabajados, 2)},
            {"Concepto": "Salarios pendientes", "Importe": salario_pendiente},
            {"Concepto": "Aguinaldo proporcional", "Importe": aguinaldo_proporcional},
            {"Concepto": "Vacaciones proporcionales", "Importe": vacaciones_proporcionales},
            {"Concepto": "Prima vacacional", "Importe": prima_vacacional},
            {"Concepto": "Indemnización 3 meses", "Importe": indemnizacion_3_meses},
            {"Concepto": "20 días por año", "Importe": veinte_dias_por_anio},
            {"Concepto": "Prima de antigüedad", "Importe": prima_antiguedad},
            {"Concepto": "TOTAL ESTIMADO", "Importe": total},
        ]

        st.markdown('<div class="result-card">', unsafe_allow_html=True)

        st.subheader("Resultado del cálculo")

        if nombre:
            st.write(f"**Trabajador:** {nombre.upper()}")

        st.write(f"**Tipo de salida:** {tipo_salida}")
        st.write(f"**Tiempo trabajado:** {dias_trabajados} días / {anios_trabajados:.2f} años")
        st.write(f"**Vacaciones anuales consideradas:** {dias_vacaciones_anuales} días")

        st.metric("Total estimado a recibir", formato_moneda(total))

        df_resultado = pd.DataFrame(desglose)
        df_resultado["Importe"] = df_resultado["Importe"].apply(
            lambda x: formato_moneda(x) if isinstance(x, float) else x
        )

        st.dataframe(df_resultado, use_container_width=True)

        st.markdown("</div>", unsafe_allow_html=True)

        excel = crear_excel(desglose)

        st.download_button(
            "📥 Descargar cálculo en Excel",
            data=excel,
            file_name="calculo_finiquito_liquidacion.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )