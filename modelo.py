"""Modelo mensual de ahorro y retiro, expresado en pesos de hoy.

Convención: las aportaciones se hacen al final de cada mes hasta cumplir la
edad de retiro. El primer retiro se hace en ese mismo momento. Los siguientes
retiros ocurren al principio de cada mes hasta un mes antes de la edad final.
"""

from __future__ import annotations

from dataclasses import dataclass
from math import expm1, log1p


AHORRO = "Ahorro necesario"
INGRESO = "Ingreso posible"
EDAD = "Edad de retiro"
MODOS = (AHORRO, INGRESO, EDAD)


@dataclass(frozen=True)
class Entradas:
    modo: str
    edad_actual: int
    edad_final: int
    capital_inicial: float
    rendimiento_nominal: float  # fracción anual, p. ej. 0.055
    inflacion: float             # fracción anual, p. ej. 0.045
    edad_retiro: int | None = None
    aporte_mensual: float | None = None  # pesos de hoy
    ingreso_mensual: float | None = None  # pesos de hoy


@dataclass(frozen=True)
class Resultado:
    edad_retiro: int | None
    aporte_mensual: float | None
    ingreso_mensual: float | None
    capital_al_retiro: float | None
    excedente_al_retiro: float | None
    rendimiento_real_mensual: float
    proyeccion: tuple[dict[str, float], ...]
    alternativas: tuple[tuple[int, float], ...]


def validar(e: Entradas) -> None:
    if e.modo not in MODOS:
        raise ValueError("Selecciona una de las tres preguntas.")
    if not (18 <= e.edad_actual < e.edad_final <= 120):
        raise ValueError("La edad final debe ser mayor que la actual (máximo: 120).")
    if e.capital_inicial < 0:
        raise ValueError("El ahorro actual no puede ser negativo.")
    if e.rendimiento_nominal <= -1 or e.inflacion <= -1:
        raise ValueError("Rendimiento e inflación deben ser mayores que −100 %.")
    if e.modo != EDAD and (e.edad_retiro is None or not e.edad_actual <= e.edad_retiro < e.edad_final):
        raise ValueError("La edad de retiro debe estar entre la edad actual y la edad final.")
    if e.modo != AHORRO and (e.aporte_mensual is None or e.aporte_mensual < 0):
        raise ValueError("La aportación mensual no puede ser negativa.")
    if e.modo != INGRESO and (e.ingreso_mensual is None or e.ingreso_mensual <= 0):
        raise ValueError("El ingreso deseado debe ser mayor que cero.")


def tasa_real_mensual(e: Entradas) -> float:
    return ((1 + e.rendimiento_nominal) / (1 + e.inflacion)) ** (1 / 12) - 1


def factor_aportes(meses: int, tasa: float) -> float:
    """Valor al retiro de un peso aportado al final de cada mes."""
    if meses == 0:
        return 0.0
    if abs(tasa) < 1e-12:
        return float(meses)
    return expm1(meses * log1p(tasa)) / tasa


def factor_retiros(meses: int, tasa: float) -> float:
    """Capital al retiro necesario para retiros al inicio de cada mes."""
    if meses <= 0:
        raise ValueError("Debe existir al menos un mes de retiro.")
    if abs(tasa) < 1e-12:
        return float(meses)
    return -expm1(-meses * log1p(tasa)) * (1 + tasa) / tasa


def capital_retiro(capital: float, aporte: float, meses: int, tasa: float) -> float:
    return capital * (1 + tasa) ** meses + aporte * factor_aportes(meses, tasa)


def ingresos_para_edad(e: Entradas, edad: int, tasa: float) -> float:
    ahorro = 12 * (edad - e.edad_actual)
    retiro = 12 * (e.edad_final - edad)
    capital = capital_retiro(e.capital_inicial, e.aporte_mensual or 0, ahorro, tasa)
    return capital / factor_retiros(retiro, tasa)


def proyectar(e: Entradas, edad: int, aporte: float, ingreso: float, tasa: float) -> tuple[dict[str, float], ...]:
    meses_ahorro = 12 * (edad - e.edad_actual)
    meses_retiro = 12 * (e.edad_final - edad)
    meses_totales = 12 * (e.edad_final - e.edad_actual)
    saldo = e.capital_inicial
    filas = []
    for mes in range(meses_totales + 1):
        saldo_inicial = saldo
        interes = saldo * tasa if mes else 0.0
        contribucion = aporte if 0 < mes <= meses_ahorro else 0.0
        retiro = ingreso if meses_ahorro <= mes < meses_ahorro + meses_retiro else 0.0
        saldo = saldo + interes + contribucion - retiro
        # Los residuos de coma flotante no son dinero adeudado.
        if abs(saldo) < 1e-7:
            saldo = 0.0
        indice_precios = (1 + e.inflacion) ** (mes / 12)
        filas.append({
            "mes": float(mes),
            "edad": e.edad_actual + mes / 12,
            "saldo_inicial_real": saldo_inicial,
            "interes_real": interes,
            "aporte_real": contribucion,
            "retiro_real": retiro,
            "saldo_real": saldo,
            "saldo_nominal": saldo * indice_precios,
            "aporte_nominal": contribucion * indice_precios,
            "retiro_nominal": retiro * indice_precios,
        })
    return tuple(filas)


def resolver(e: Entradas) -> Resultado:
    validar(e)
    tasa = tasa_real_mensual(e)
    alternativas: tuple[tuple[int, float], ...] = ()

    if e.modo == EDAD:
        alternativas = tuple((edad, ingresos_para_edad(e, edad, tasa))
                             for edad in range(e.edad_actual, e.edad_final))
        edad = next((edad for edad, ingreso_posible in alternativas
                     if ingreso_posible >= (e.ingreso_mensual or 0) - 1e-8), None)
        if edad is None:
            return Resultado(None, e.aporte_mensual, e.ingreso_mensual, None,
                             None, tasa, (), alternativas)
    else:
        edad = e.edad_retiro

    assert edad is not None
    meses_ahorro = 12 * (edad - e.edad_actual)
    meses_retiro = 12 * (e.edad_final - edad)
    factor_ahorro = factor_aportes(meses_ahorro, tasa)
    factor_ingreso = factor_retiros(meses_retiro, tasa)
    capital_sin_aportes = e.capital_inicial * (1 + tasa) ** meses_ahorro

    if e.modo == AHORRO:
        if meses_ahorro == 0:
            raise ValueError("Para calcular cuánto ahorrar, el retiro debe ser posterior a la edad actual.")
        ingreso = float(e.ingreso_mensual)
        aporte = max(0.0, (ingreso * factor_ingreso - capital_sin_aportes) / factor_ahorro)
    else:
        aporte = float(e.aporte_mensual)
        ingreso = (capital_sin_aportes + aporte * factor_ahorro) / factor_ingreso if e.modo == INGRESO else float(e.ingreso_mensual)

    capital = capital_sin_aportes + aporte * factor_ahorro
    excedente = capital - ingreso * factor_ingreso
    return Resultado(edad, aporte, ingreso, capital, excedente, tasa,
                     proyectar(e, edad, aporte, ingreso, tasa), alternativas)
