"""Aplicación Streamlit: tres preguntas, un modelo mensual."""

from __future__ import annotations

import pandas as pd
import streamlit as st

from exportar_excel import crear_excel
from modelo import AHORRO, EDAD, INGRESO, Entradas, resolver


st.set_page_config(page_title="Plan de ahorro y retiro", page_icon="📈", layout="wide")


def dinero(monto: float) -> str:
    return f"${monto:,.0f}"


@st.cache_data(show_spinner=False)
def preparar_excel(entradas: Entradas) -> bytes:
    return crear_excel(entradas, resolver(entradas))


st.title("Plan de ahorro y retiro")
st.write("Responde una pregunta sobre tu retiro. Todos los montos que ingreses están en **pesos de hoy**.")

preguntas = {
    "¿Cuánto debo ahorrar al mes?": AHORRO,
    "¿Con cuánto podré vivir?": INGRESO,
    "¿A qué edad podría retirarme?": EDAD,
}
pregunta = st.radio("¿Qué quieres averiguar?", list(preguntas), horizontal=True)
modo = preguntas[pregunta]

st.subheader("Tu situación")
col1, col2, col3 = st.columns(3)
with col1:
    edad_actual = int(st.number_input("Edad actual", min_value=18, max_value=118,
                                      value=38, step=1))
with col2:
    edad_final = int(st.number_input("¿Hasta qué edad debe alcanzarte el dinero?",
                                     min_value=edad_actual+2, max_value=120,
                                     value=max(edad_actual+2, 75), step=1))
with col3:
    capital_inicial = float(st.number_input("¿Cuánto tienes ya ahorrado para el retiro?",
                                            min_value=0.0, value=0.0, step=10000.0,
                                            format="%.0f", help="Indica el saldo actual, en pesos de hoy."))

col1, col2 = st.columns(2)
edad_retiro = None
aporte = None
ingreso = None
with col1:
    if modo != EDAD:
        edad_minima = edad_actual + 1 if modo == AHORRO else edad_actual
        edad_retiro = int(st.number_input("¿A qué edad te retirarías?",
                                          min_value=edad_minima, max_value=edad_final-1,
                                          value=min(max(55, edad_minima), edad_final-1),
                                          step=1))
    else:
        aporte = float(st.number_input("¿Cuánto podrías ahorrar cada mes? (pesos de hoy)",
                                         min_value=0.0, value=20000.0, step=1000.0, format="%.0f"))
with col2:
    if modo == INGRESO:
        aporte = float(st.number_input("¿Cuánto podrías ahorrar cada mes? (pesos de hoy)",
                                         min_value=0.0, value=20000.0, step=1000.0, format="%.0f"))
    else:
        ingreso = float(st.number_input("¿Con cuánto quieres vivir al mes? (pesos de hoy)",
                                          min_value=1.0, value=25000.0, step=1000.0, format="%.0f"))

with st.expander("Rendimiento e inflación", expanded=True):
    col1, col2 = st.columns(2)
    with col1:
        nominal_pct = float(st.number_input("Rendimiento nominal anual (%)", min_value=-99.0,
                                             max_value=100.0, value=5.5, step=0.1, format="%.2f"))
    with col2:
        inflacion_pct = float(st.number_input("Inflación anual (%)", min_value=-99.0,
                                               max_value=100.0, value=4.5, step=0.1, format="%.2f"))
    st.caption("La app calcula la diferencia exacta entre ambas tasas y aplica una tasa mensual efectiva.")

entrada = Entradas(modo=modo, edad_actual=edad_actual, edad_final=edad_final,
                   capital_inicial=capital_inicial, edad_retiro=edad_retiro,
                   aporte_mensual=aporte, ingreso_mensual=ingreso,
                   rendimiento_nominal=nominal_pct/100, inflacion=inflacion_pct/100)

try:
    resultado = resolver(entrada)
except ValueError as error:
    st.error(str(error))
    st.stop()

st.divider()
st.subheader("Resultado")
if resultado.edad_retiro is None:
    st.warning("Con ese ahorro mensual y ese ingreso deseado, no existe una edad de retiro "
               "entre tu edad actual y la edad final que permita financiar el plan.")
    st.write("Puedes aumentar tu ahorro mensual, reducir el ingreso deseado o ampliar la edad final.")
else:
    if modo == AHORRO:
        st.metric("Ahorro mensual necesario, en pesos de hoy", dinero(resultado.aporte_mensual))
        if resultado.aporte_mensual < 0.005:
            st.write("Tu ahorro actual ya financia el ingreso que indicaste bajo estos supuestos.")
        else:
            st.write(f"Aporta **{dinero(resultado.aporte_mensual)} al mes**, ajustando ese importe "
                     f"con la inflación, para retirarte a los **{resultado.edad_retiro} años**.")
    elif modo == INGRESO:
        st.metric("Ingreso mensual posible, en pesos de hoy", dinero(resultado.ingreso_mensual))
        st.write(f"Podrías retirar **{dinero(resultado.ingreso_mensual)} al mes en poder adquisitivo "
                 f"actual** desde los **{resultado.edad_retiro}** hasta un mes antes de los **{edad_final} años**.")
    else:
        st.metric("Primera edad de retiro que cumple tu meta", f"{resultado.edad_retiro} años")
        st.write(f"Con aportaciones de **{dinero(resultado.aporte_mensual)} al mes**, "
                 f"podrías retirar **{dinero(resultado.ingreso_mensual)} al mes** desde esa edad.")

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Capital al retirarte, pesos de hoy", dinero(resultado.capital_al_retiro))
    m2.metric(
        "Saldo nominal al retirarte",
        dinero(
            resultado.capital_al_retiro
            * (1 + inflacion_pct / 100) ** (resultado.edad_retiro - edad_actual)
        ),
    )
    m3.metric("Años de aportaciones", resultado.edad_retiro - edad_actual)
    m4.metric("Años financiados", edad_final - resultado.edad_retiro)

    st.subheader("Cómo evolucionaría tu saldo")
    tabla = pd.DataFrame(resultado.proyeccion)
    grafico = tabla.iloc[::3, :].copy()
    if grafico.iloc[-1]["mes"] != tabla.iloc[-1]["mes"]:
        grafico = pd.concat([grafico, tabla.tail(1)], ignore_index=True)
    grafico = grafico.rename(columns={"edad": "Edad", "saldo_real": "Saldo en pesos de hoy"})
    st.line_chart(grafico, x="Edad", y="Saldo en pesos de hoy", height=325)
    st.caption("Las aportaciones se realizan al final de cada mes. El primer retiro ocurre al "
               "cumplir la edad elegida, después de la última aportación. El dinero restante sigue invertido.")

if modo == EDAD:
    st.subheader("Ingreso posible según la edad de retiro")
    puntos = pd.DataFrame(resultado.alternativas,
                          columns=["Edad de retiro", "Ingreso mensual posible, pesos de hoy"])
    st.line_chart(puntos, x="Edad de retiro", y="Ingreso mensual posible, pesos de hoy", height=275)

st.download_button(
    "Descargar análisis en Excel",
    data=preparar_excel(entrada),
    file_name="plan_ahorro_retiro.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    help="Incluye supuestos editables, proyección mensual y gráficos nativos de Excel.",
)
st.caption("Modelo con rendimiento e inflación constantes; no incorpora impuestos, comisiones, pensiones ni otros ingresos.")
