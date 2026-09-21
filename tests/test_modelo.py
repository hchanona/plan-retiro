"""Ejecutar con: python -m unittest discover -s tests -v"""

from __future__ import annotations

from io import BytesIO
import unittest
from xml.etree import ElementTree
from zipfile import ZipFile

from exportar_excel import crear_excel
from modelo import AHORRO, EDAD, INGRESO, Entradas, resolver


def ejemplo(modo: str, **cambios) -> Entradas:
    datos = dict(modo=modo, edad_actual=38, edad_final=75, capital_inicial=0.0,
                 rendimiento_nominal=0.055, inflacion=0.045,
                 edad_retiro=55 if modo != EDAD else None,
                 aporte_mensual=20000.0 if modo != AHORRO else None,
                 ingreso_mensual=25000.0 if modo != INGRESO else None)
    datos.update(cambios)
    return Entradas(**datos)


def valor_excel(xlsx: bytes, referencia: str):
    with ZipFile(BytesIO(xlsx)) as archivo:
        root = ElementTree.fromstring(archivo.read("xl/worksheets/sheet1.xml"))
    ns = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    celda = root.find(f".//x:c[@r='{referencia}']", ns)
    valor = celda.find("x:v", ns) if celda is not None else None
    return float(valor.text) if valor is not None and valor.text else None


class ModeloTest(unittest.TestCase):
    def test_ahorro_ingreso_y_saldo_se_reconcilian(self):
        meta = resolver(ejemplo(AHORRO))
        inverso = resolver(ejemplo(INGRESO, aporte_mensual=meta.aporte_mensual))
        self.assertAlmostEqual(inverso.ingreso_mensual, 25000, places=6)
        self.assertAlmostEqual(meta.capital_al_retiro, inverso.capital_al_retiro, places=6)
        self.assertAlmostEqual(meta.proyeccion[-1]["saldo_real"], 0, places=5)
        primer_retiro = (55-38)*12
        self.assertEqual(meta.proyeccion[primer_retiro]["retiro_real"], 25000)
        self.assertEqual(meta.proyeccion[primer_retiro]["aporte_real"], meta.aporte_mensual)
        self.assertEqual(meta.proyeccion[primer_retiro-1]["retiro_real"], 0)
        self.assertEqual(meta.proyeccion[-1]["retiro_real"], 0)

    def test_edad_es_la_primera_que_cumple_la_meta(self):
        r = resolver(ejemplo(EDAD))
        self.assertEqual(r.edad_retiro, 57)
        self.assertLess(dict(r.alternativas)[56], 25000)
        self.assertGreaterEqual(dict(r.alternativas)[57], 25000)
        self.assertGreaterEqual(r.proyeccion[-1]["saldo_real"], 0)

    def test_casos_limite_y_tasa_real_negativa(self):
        sin_ahorro = resolver(ejemplo(EDAD, aporte_mensual=0, capital_inicial=0))
        self.assertIsNone(sin_ahorro.edad_retiro)
        inmediato = resolver(ejemplo(INGRESO, edad_retiro=38, capital_inicial=5000000,
                                     aporte_mensual=0, rendimiento_nominal=0, inflacion=0))
        self.assertAlmostEqual(inmediato.ingreso_mensual, 5000000 / (37*12), places=6)
        self.assertAlmostEqual(inmediato.proyeccion[0]["retiro_real"], inmediato.ingreso_mensual)
        tasa_negativa = resolver(ejemplo(INGRESO, rendimiento_nominal=0.02, inflacion=0.06))
        self.assertLess(tasa_negativa.rendimiento_real_mensual, 0)
        self.assertAlmostEqual(tasa_negativa.proyeccion[-1]["saldo_real"], 0, places=5)
        ya_alcanza = resolver(ejemplo(AHORRO, capital_inicial=10000000))
        self.assertEqual(ya_alcanza.aporte_mensual, 0)
        self.assertGreater(ya_alcanza.excedente_al_retiro, 0)

    def test_excel_tiene_resultados_y_graficos_nativos(self):
        for modo in (AHORRO, INGRESO, EDAD):
            with self.subTest(modo=modo):
                entrada = ejemplo(modo)
                r = resolver(entrada)
                datos = crear_excel(entrada, r)
                with ZipFile(BytesIO(datos)) as archivo:
                    graficos = [n for n in archivo.namelist() if n.startswith("xl/charts/chart") and n.endswith(".xml")]
                    self.assertEqual(len(graficos), 3 if modo == EDAD else 2)
                    self.assertIn(b"<f>", archivo.read("xl/worksheets/sheet1.xml"))
                    self.assertIn(b"Graficos!$S$5:$S$125", archivo.read("xl/charts/chart1.xml"))
                    self.assertGreater(archivo.read("xl/charts/chart1.xml").count(b"<c:pt "), 200)
                self.assertAlmostEqual(valor_excel(datos, "B16"), r.rendimiento_real_mensual, places=10)
                self.assertAlmostEqual(valor_excel(datos, "B17"), r.edad_retiro, places=8)
                self.assertAlmostEqual(valor_excel(datos, "B23"), r.aporte_mensual, places=6)
                self.assertAlmostEqual(valor_excel(datos, "B25"), r.ingreso_mensual, places=6)

    def test_excel_sin_edad_factible(self):
        entrada = ejemplo(EDAD, aporte_mensual=0, capital_inicial=0)
        resultado = resolver(entrada)
        datos = crear_excel(entrada, resultado)
        self.assertIsNone(valor_excel(datos, "B17"))
        with ZipFile(BytesIO(datos)) as archivo:
            self.assertIn("xl/charts/chart3.xml", archivo.namelist())


if __name__ == "__main__":
    unittest.main()
