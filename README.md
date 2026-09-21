# Plan de ahorro y retiro

Aplicación de Streamlit para resolver tres preguntas:

1. ¿Cuánto debo ahorrar al mes para alcanzar un ingreso deseado?
2. Si ahorro cierta cantidad, ¿cuánto podré retirar al mes?
3. ¿Cuál es la primera edad a la que podría retirarme con una meta de ingreso?

Todos los montos introducidos se expresan en **pesos de hoy**. Puedes descargar un
Excel con las entradas editables, el cálculo mensual y gráficos nativos.

## Ejecutar localmente

Con Python 3.10 o posterior, desde la carpeta del proyecto:

```bash
python -m venv .venv
# Linux/macOS:
source .venv/bin/activate
# Windows (PowerShell):
# .venv\Scripts\Activate.ps1
pip install -r requirements.txt
streamlit run app.py
```

## Publicar en GitHub y Streamlit Community Cloud

1. Descomprime el ZIP y sube **los archivos de esta carpeta a la raíz** de un repositorio de GitHub.
2. En [Streamlit Community Cloud](https://share.streamlit.io/), crea una app y selecciona ese repositorio.
3. Indica `app.py` como archivo principal. La plataforma instalará `requirements.txt`.

El proyecto no requiere claves, cuentas ni archivos de datos.

Para verificar los cálculos antes de publicar:

```bash
python -m unittest discover -s tests -v
```

## Convenciones de cálculo

- El rendimiento anual nominal y la inflación anual se convierten a una tasa
  mensual real efectiva: `((1 + nominal) / (1 + inflacion)) ** (1/12) - 1`.
- Las aportaciones mantienen constante su poder adquisitivo, se realizan al
  final de cada mes y terminan al cumplir la edad de retiro.
- El primer retiro se realiza al cumplir la edad de retiro, después de la
  última aportación. Los retiros siguientes se hacen cada mes. El último se
  hace un mes antes de la edad final.
- El modelo supone que el rendimiento y la inflación son constantes. No incluye
  impuestos, comisiones, pensiones u otras fuentes de ingreso.
- Para la tercera pregunta se prueban edades **enteras** desde la actual hasta
  un año antes de la edad final. Se devuelve la primera que cumple la meta.
- Cuando el capital ya basta para financiar la meta, el ahorro necesario se
  muestra como cero y se informa el excedente.

## Archivos

- `app.py`: interfaz y gráficos en Streamlit.
- `modelo.py`: único motor de cálculo de los tres problemas.
- `exportar_excel.py`: genera el libro editable con fórmulas y gráficos.
- `requirements.txt`: dependencias para ejecutar y publicar.
- `tests/test_modelo.py`: pruebas de coherencia y del libro descargable.

Las celdas azules del Excel son editables. Abre el libro en una aplicación que
recalcule fórmulas para que los resultados y los gráficos respondan a cambios.
La selección de pregunta se hace en la app: el libro exportado conserva esa
pregunta y deja editables sus supuestos.
