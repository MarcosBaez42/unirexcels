# Unir Exceles en un solo archivo

Este repositorio incluye un script sencillo para combinar varias hojas de
cálculo de Excel en un único libro. Cada archivo de entrada genera una
hoja en el libro consolidado utilizando el mismo nombre del archivo (sin
la extensión).

## Requisitos

- Python 3.10 o superior.
- [openpyxl](https://openpyxl.readthedocs.io/) para manipular los
  archivos de Excel.
- [xlrd](https://xlrd.readthedocs.io/) para leer libros en formato `.xls`.

Puedes instalar la dependencia principal con:

```bash
pip install -r requirements.txt
```

## Uso

1. Coloca todos tus archivos `.xlsx` o `.xls` en una misma carpeta.
2. Ejecuta el script indicando la carpeta de origen y la ruta del archivo
   combinado. Por ejemplo:

```bash
python merge_excel_files.py carpeta/de/origen --output combinado.xlsx
```

Opciones disponibles:

- `--pattern`: patrón *glob* para filtrar los archivos (por defecto
  `*.xls*`).
- `--recursive`: busca archivos en subcarpetas.
- `--values-only`: copia únicamente los valores calculados en lugar de
  las fórmulas.

El script creará un libro de Excel con cada archivo en una hoja distinta,
utilizando el nombre del archivo como nombre de la hoja (se aplican las
restricciones de Excel respecto a longitud y caracteres válidos). Las
fórmulas presentes en archivos `.xls` se copian como sus últimos valores
calculados porque ese formato no expone las expresiones originales.

### Ejecución directa (doble clic)

Si ejecutas `merge_excel_files.py` sin parámetros, el programa tomará como
origen la carpeta donde está guardado el propio script y guardará el
resultado en `combined.xlsx` dentro de esa misma carpeta. Esta modalidad es
útil para usuarios sin experiencia técnica: basta con copiar todos los
Excel junto al script y abrir el archivo de Python con doble clic.

## Pruebas

Para ejecutar los tests automatizados se necesita instalar las
dependencias de desarrollo:

```bash
pip install -r requirements-dev.txt
pytest
```