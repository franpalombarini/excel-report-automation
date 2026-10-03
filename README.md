# Excel Report Automation

Generador automático de **reportes Excel con resumen estadístico** a partir de archivos CSV.

Servicio típico para PyMEs: convertir ventas/gastos sueltos en un reporte profesional de una hoja, sin tocar Excel a mano.

## Uso

```bash
pip install -r requirements.txt
python reporte.py ventas.csv --salida reporte.xlsx
```

Genera `reporte.xlsx` con:
- **Hoja "Datos"**: el CSV completo, con formato de tabla.
- **Hoja "Resumen"**: totales y promedios por columna numérica + top 5 filas.

## Ejemplo

```bash
python reporte.py ejemplo_ventas.csv --salida salida.xlsx
```

## Tests

```bash
pytest -v
```

3 tests: generación del archivo, hojas correctas, cálculo de totales.
