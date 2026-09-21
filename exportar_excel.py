"""Libro Excel con fórmulas y gráficos vinculados al modelo de la app."""

from __future__ import annotations

from io import BytesIO

import xlsxwriter

from modelo import AHORRO, EDAD, INGRESO, Entradas, Resultado


MAX_MESES = 1224  # Edades editables de 18 a 120 años: máximo 102 * 12 meses.
MAX_EDADES = 102
MONEDA = '"$"#,##0;[Red]("$"#,##0);–'


def crear_excel(e: Entradas, r: Resultado) -> bytes:
    """Genera un xlsx editable; las fórmulas se recalculan al abrir en Excel."""
    salida = BytesIO()
    libro = xlsxwriter.Workbook(salida, {"in_memory": True})
    libro.set_calc_mode("auto")
    hoja = libro.add_worksheet("Resumen")
    edades = libro.add_worksheet("Edades")
    detalle = libro.add_worksheet("Proyeccion")
    graficas = libro.add_worksheet("Graficos")

    fondo = "#17324D"
    acento = "#087E79"
    azul = "#145A9D"
    tinta = "#19324A"
    titulo = libro.add_format({"bold": True, "font_size": 17, "font_color": "white", "bg_color": fondo,
                               "valign": "vcenter"})
    seccion = libro.add_format({"bold": True, "font_color": "white", "bg_color": acento})
    etiqueta = libro.add_format({"font_color": tinta})
    nota = libro.add_format({"font_color": "#5D6F80", "text_wrap": True, "valign": "top"})
    cabecera = libro.add_format({"bold": True, "font_color": "white", "bg_color": fondo,
                                 "text_wrap": True, "valign": "vcenter"})
    dinero = libro.add_format({"num_format": MONEDA})
    dinero_entrada = libro.add_format({"num_format": MONEDA, "font_color": azul, "bg_color": "#EAF3FF"})
    porcentaje_entrada = libro.add_format({"num_format": "0.00%", "font_color": azul, "bg_color": "#EAF3FF"})
    entero_entrada = libro.add_format({"num_format": "0", "font_color": azul, "bg_color": "#EAF3FF"})
    texto_entrada = libro.add_format({"font_color": azul, "bg_color": "#EAF3FF"})
    resultado_dinero = libro.add_format({"num_format": MONEDA, "bold": True,
                                         "font_size": 13, "font_color": acento})
    resultado_edad = libro.add_format({"num_format": '0" años"', "bold": True,
                                       "font_size": 13, "font_color": acento})
    porcentaje = libro.add_format({"num_format": "0.0000%"})
    numero = libro.add_format({"num_format": "0"})
    edad_decimal = libro.add_format({"num_format": "0.0"})

    hoja.set_column("A:A", 38)
    hoja.set_column("B:B", 23)
    hoja.set_column("C:C", 3)
    hoja.set_column("D:D", 48)
    hoja.merge_range("A1:D2", "Plan de ahorro y retiro", titulo)
    hoja.set_row(0, 23)
    hoja.merge_range("A4:D4", "Supuestos editables", seccion)
    hoja.write("A3", "Pregunta", etiqueta)
    hoja.write("B3", e.modo, texto_entrada)
    datos = [
        (6, "Edad actual", e.edad_actual, entero_entrada),
        (7, "Edad de retiro indicada", e.edad_retiro if e.modo != EDAD else None, entero_entrada),
        (8, "El dinero debe alcanzar hasta", e.edad_final, entero_entrada),
        (9, "Ahorro actual (pesos de hoy)", e.capital_inicial, dinero_entrada),
        (10, "Aportación mensual (pesos de hoy)", e.aporte_mensual if e.modo != AHORRO else None, dinero_entrada),
        (11, "Ingreso deseado (pesos de hoy)", e.ingreso_mensual if e.modo != INGRESO else None, dinero_entrada),
        (12, "Rendimiento nominal anual", e.rendimiento_nominal, porcentaje_entrada),
        (13, "Inflación anual", e.inflacion, porcentaje_entrada),
    ]
    for fila, nombre, valor, formato in datos:
        hoja.write(f"A{fila}", nombre, etiqueta)
        if valor is None:
            hoja.write_blank(f"B{fila}", None, formato)
        else:
            hoja.write(f"B{fila}", valor, formato)
    hoja.write("D6", "Modifica las celdas azules. La pregunta elegida determina qué variable calcula el libro.", nota)
    hoja.write("D8", "Los ingresos y aportaciones reales conservan su poder adquisitivo: sus importes nominales crecen con la inflación.", nota)
    hoja.write("D11", "Aportación al final de cada mes. Primer retiro al cumplir la edad de retiro, después de la última aportación.", nota)
    hoja.write("D13", "Último retiro: un mes antes de la edad final. El capital restante sigue invertido.", nota)
    for fila, altura in ((6, 32), (8, 47), (11, 47), (13, 32)):
        hoja.set_row(fila - 1, altura)
    hoja.data_validation("B6", {"validate": "integer", "criteria": "between", "minimum": 18, "maximum": 118})
    hoja.data_validation("B7", {"validate": "custom", "value":
                                     '=OR(B7="",AND(B7>=B6+(B3="Ahorro necesario"),B7<B8,MOD(B7,1)=0))'})
    hoja.data_validation("B8", {"validate": "custom", "value": '=AND(B8>=B6+2,B8<=120,MOD(B8,1)=0)'})
    hoja.data_validation("B9:B11", {"validate": "decimal", "criteria": ">=", "value": 0})
    hoja.data_validation("B12:B13", {"validate": "decimal", "criteria": "between", "minimum": -0.99, "maximum": 1})
    hoja.merge_range("A15:B15", "Resultado y cálculos", seccion)

    def formula(ref: str, expresion: str, formato, valor=0):
        hoja.write_formula(ref, expresion, formato, valor)

    labels = {
        16: "Rendimiento real mensual", 17: "Edad de retiro usada",
        18: "Meses de aportaciones", 19: "Meses de retiros",
        20: "Factor de acumulación", 21: "Factor de retiros",
        22: "Capital inicial acumulado al retiro", 23: "Aportación mensual calculada",
        24: "Capital al retiro (pesos de hoy)", 25: "Ingreso mensual calculado",
        26: "Excedente frente al ingreso deseado",
    }
    for fila, label in labels.items():
        hoja.write(f"A{fila}", label, etiqueta)

    n_ahorro = 12 * (r.edad_retiro - e.edad_actual) if r.edad_retiro is not None else 0
    n_retiro = 12 * (e.edad_final - r.edad_retiro) if r.edad_retiro is not None else 0
    q = 1 + r.rendimiento_real_mensual
    fa = n_ahorro if abs(q - 1) < 1e-12 else ((q ** n_ahorro) - 1) / (q - 1)
    fr = n_retiro if abs(q - 1) < 1e-12 else (1 - q ** (-n_retiro)) * q / (q - 1) if n_retiro else 0
    capital_sin = e.capital_inicial * q ** n_ahorro
    formula("B16", "=((1+B12)/(1+B13))^(1/12)-1", porcentaje, r.rendimiento_real_mensual)
    formula("B17", '=IF(B3="Edad de retiro",IFERROR(INDEX(Edades!$A$5:$A$106,MATCH(1,Edades!$E$5:$E$106,0)),""),B7)',
            resultado_edad, r.edad_retiro if r.edad_retiro is not None else "")
    formula("B18", '=IF(B17="","",12*(B17-B6))', numero, n_ahorro if r.edad_retiro is not None else "")
    formula("B19", '=IF(B17="","",12*(B8-B17))', numero, n_retiro if r.edad_retiro is not None else "")
    formula("B20", '=IF(B18="","",IF(ABS(B16)<1E-10,B18,((1+B16)^B18-1)/B16))',
            numero, fa if r.edad_retiro is not None else "")
    formula("B21", '=IF(B19="","",IF(ABS(B16)<1E-10,B19,(1-(1+B16)^(-B19))*(1+B16)/B16))',
            numero, fr if r.edad_retiro is not None else "")
    formula("B22", '=IF(B18="","",B9*(1+B16)^B18)', dinero,
            capital_sin if r.edad_retiro is not None else "")
    formula("B23", '=IF(B17="","",IF(B3="Ahorro necesario",MAX(0,(B11*B21-B22)/B20),B10))',
            resultado_dinero, r.aporte_mensual if r.edad_retiro is not None else "")
    formula("B24", '=IF(B17="","",B22+B23*B20)', dinero,
            r.capital_al_retiro if r.edad_retiro is not None else "")
    formula("B25", '=IF(B17="","",IF(B3="Ingreso posible",B24/B21,B11))',
            resultado_dinero, r.ingreso_mensual if r.edad_retiro is not None else "")
    formula("B26", '=IF(OR(B17="",B3="Ingreso posible"),"",B24-B25*B21)', dinero,
            r.excedente_al_retiro if r.edad_retiro is not None and e.modo != INGRESO else "")
    hoja.write("A28", "Tasa anual real equivalente", etiqueta)
    formula("B28", "=(1+B12)/(1+B13)-1", porcentaje,
            (1 + e.rendimiento_nominal) / (1 + e.inflacion) - 1)
    hoja.write("A30", "La proyección mensual y los gráficos se actualizan al cambiar las celdas azules.", nota)
    hoja.set_row(29, 30)
    hoja.freeze_panes(5, 0)

    edades.merge_range("A1:E2", "Ingreso posible por edad de retiro", titulo)
    edades.write("A3", "Se usa para encontrar la primera edad que cubre el ingreso deseado.", nota)
    edades.set_column("A:C", 20)
    edades.set_column("D:D", 24)
    edades.set_column("E:E", 17)
    edades.write_row("A4", ["Edad de retiro", "Meses de ahorro", "Meses de retiro",
                             "Ingreso mensual posible", "Cumple la meta"], cabecera)
    edades.set_row(3, 34)
    for i in range(MAX_EDADES):
        fila = i + 5
        candidata = e.edad_actual + i
        ingreso_posible = 0.0
        if candidata < e.edad_final and e.aporte_mensual is not None:
            ingreso_posible = next((v for a, v in r.alternativas if a == candidata), 0.0)
        activa = e.modo == EDAD and candidata < e.edad_final
        edades.write_formula(f"A{fila}", f'=IF(Resumen!$B$6+{i}<Resumen!$B$8,Resumen!$B$6+{i},"")',
                            numero, candidata if candidata < e.edad_final else "")
        edades.write_formula(f"B{fila}", f'=IF(A{fila}="","",12*(A{fila}-Resumen!$B$6))',
                            numero, (candidata-e.edad_actual)*12 if candidata < e.edad_final else "")
        edades.write_formula(f"C{fila}", f'=IF(A{fila}="","",12*(Resumen!$B$8-A{fila}))',
                            numero, (e.edad_final-candidata)*12 if candidata < e.edad_final else "")
        # Con una sola fórmula mensual se comparan todas las edades elegibles.
        ahorro_expr = (f'(Resumen!$B$9*(1+Resumen!$B$16)^B{fila}'
                       f'+Resumen!$B$10*IF(ABS(Resumen!$B$16)<1E-10,B{fila},'
                       f'((1+Resumen!$B$16)^B{fila}-1)/Resumen!$B$16))')
        retiro_expr = (f'IF(ABS(Resumen!$B$16)<1E-10,C{fila},'
                       f'(1-(1+Resumen!$B$16)^(-C{fila}))*'
                       f'(1+Resumen!$B$16)/Resumen!$B$16)')
        edades.write_formula(f"D{fila}",
                            f'=IF(OR(A{fila}="",Resumen!$B$3<>"Edad de retiro"),"",{ahorro_expr}/{retiro_expr})',
                            dinero, ingreso_posible if activa else "")
        edades.write_formula(f"E{fila}", f'=IF(D{fila}="",0,IF(D{fila}+0.00000001>=Resumen!$B$11,1,0))',
                            numero, int(activa and ingreso_posible >= (e.ingreso_mensual or 0)-1e-8))
    edades.freeze_panes(4, 1)
    edades.autofilter(3, 0, 3 + MAX_EDADES, 4)

    detalle.merge_range("A1:J2", "Proyección mes a mes", titulo)
    detalle.write("A3", "Valores reales en pesos de hoy; valores nominales en los pesos de cada mes.", nota)
    detalle.set_column("A:A", 9)
    detalle.set_column("B:B", 12)
    detalle.set_column("C:J", 19)
    detalle.set_row(4, 36)
    encabezados = ["Mes", "Edad", "Saldo inicial real", "Interés real", "Aportación real",
                   "Retiro real", "Saldo final real", "Saldo final nominal",
                   "Aportación nominal", "Retiro nominal"]
    detalle.write_row("A5", encabezados, cabecera)
    for mes in range(MAX_MESES+1):
        fila = 6 + mes
        ejemplo = r.proyeccion[mes] if mes < len(r.proyeccion) else None
        detalle.write_number(f"A{fila}", mes, numero)
        def v(k: str):
            return ejemplo[k] if ejemplo is not None else ""
        detalle.write_formula(f"B{fila}",
                             f'=IF(OR(Resumen!$B$17="",A{fila}>12*(Resumen!$B$8-Resumen!$B$6)),"",Resumen!$B$6+A{fila}/12)',
                             edad_decimal, v("edad"))
        ref_apertura = "Resumen!$B$9" if mes == 0 else f"G{fila-1}"
        detalle.write_formula(f"C{fila}", f'=IF(B{fila}="","",{ref_apertura})', dinero, v("saldo_inicial_real"))
        detalle.write_formula(f"D{fila}", f'=IF(B{fila}="","",'+("0" if mes == 0 else f"C{fila}*Resumen!$B$16")+")",
                             dinero, v("interes_real"))
        detalle.write_formula(f"E{fila}", f'=IF(B{fila}="","",IF(AND(A{fila}>0,A{fila}<=Resumen!$B$18),Resumen!$B$23,0))',
                             dinero, v("aporte_real"))
        detalle.write_formula(f"F{fila}",
                             f'=IF(B{fila}="","",IF(AND(A{fila}>=Resumen!$B$18,A{fila}<Resumen!$B$18+Resumen!$B$19),Resumen!$B$25,0))',
                             dinero, v("retiro_real"))
        detalle.write_formula(f"G{fila}", f'=IF(B{fila}="","",C{fila}+D{fila}+E{fila}-F{fila})',
                             dinero, v("saldo_real"))
        for col, base, llave in (("H", "G", "saldo_nominal"), ("I", "E", "aporte_nominal"),
                                 ("J", "F", "retiro_nominal")):
            detalle.write_formula(f"{col}{fila}",
                                 f'=IF(B{fila}="","",{base}{fila}*(1+Resumen!$B$13)^(A{fila}/12))',
                                 dinero, v(llave))
    detalle.freeze_panes(5, 2)
    detalle.autofilter(4, 0, MAX_MESES+5, 9)

    # 121 puntos distribuidos por todo el horizonte. Así el gráfico conserva
    # siempre la escala correcta si el usuario modifica cualquiera de las
    # edades. El detalle completo sigue disponible mes a mes en Proyeccion.
    graficas.set_column("R:V", 21)
    graficas.write_row("R4", ["Edad", "Saldo real", "Saldo nominal",
                             "Aportación nominal", "Retiro nominal"], cabecera)
    for i in range(121):
        fila = i + 5
        indice_muestra = int(i / 120 * (len(r.proyeccion)-1) + 0.5) if r.proyeccion else 0
        fila_real = r.proyeccion[indice_muestra] if r.proyeccion else None
        mes_excel = f'ROUND({i}/120*12*(Resumen!$B$8-Resumen!$B$6),0)+1'
        for columna, origen, campo, formato in (
            ("R", "B", "edad", edad_decimal),
            ("S", "G", "saldo_real", dinero),
            ("T", "H", "saldo_nominal", dinero),
            ("U", "I", "aporte_nominal", dinero),
            ("V", "J", "retiro_nominal", dinero),
        ):
            cache = fila_real[campo] if fila_real is not None else ""
            graficas.write_formula(f"{columna}{fila}",
                                  f'=IF(Resumen!$B$17="","",INDEX(Proyeccion!${origen}$6:${origen}$1230,{mes_excel}))',
                                  formato, cache)
    categorias = ["Graficos", 4, 17, 124, 17]
    graficas.merge_range("A1:P2", "Gráficos del retiro", titulo)
    graficas.set_column("A:P", 13)
    graficas.write("A3", "Los datos se actualizan al editar los supuestos en Resumen.", nota)
    chart1 = libro.add_chart({"type": "line"})
    chart1.add_series({"name": "Saldo real", "categories": categorias,
                       "values": ["Graficos", 4, 18, 124, 18],
                       "line": {"color": acento, "width": 2.5}})
    chart1.set_title({"name": "Saldo de la cuenta (pesos de hoy)"})
    chart1.set_x_axis({"name": "Edad", "num_format": "0"})
    chart1.set_y_axis({"name": "Pesos de hoy", "num_format": '#,##0'})
    chart1.set_legend({"none": True})
    chart1.set_size({"width": 780, "height": 360})
    graficas.insert_chart("A5", chart1)

    chart2 = libro.add_chart({"type": "line"})
    for nombre, columna, color in (("Aportación nominal", 20, "#2876BA"),
                                   ("Retiro nominal", 21, "#E38939")):
        chart2.add_series({"name": nombre, "categories": categorias,
                           "values": ["Graficos", 4, columna, 124, columna],
                           "line": {"color": color, "width": 2.0}})
    chart2.set_title({"name": "Movimientos mensuales (pesos de cada mes)"})
    chart2.set_x_axis({"name": "Edad", "num_format": "0"})
    chart2.set_y_axis({"name": "Pesos nominales", "num_format": '#,##0'})
    chart2.set_legend({"position": "bottom"})
    chart2.set_size({"width": 780, "height": 360})
    graficas.insert_chart("A25", chart2)

    if e.modo == EDAD:
        chart3 = libro.add_chart({"type": "line"})
        chart3.add_series({"name": "Ingreso posible", "categories": "=Edades!$A$5:$A$106",
                           "values": "=Edades!$D$5:$D$106", "line": {"color": acento, "width": 2.0}})
        chart3.set_title({"name": "Ingreso según la edad de retiro (pesos de hoy)"})
        chart3.set_x_axis({"name": "Edad"})
        chart3.set_y_axis({"name": "Ingreso mensual", "num_format": '#,##0'})
        chart3.set_legend({"none": True})
        chart3.set_size({"width": 780, "height": 360})
        graficas.insert_chart("A45", chart3)

    libro.close()
    salida.seek(0)
    return salida.getvalue()
