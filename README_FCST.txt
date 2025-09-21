# README — FCST con Excel + Python (xlwings + Prophet)

Este proyecto permite generar pronósticos de ventas directamente desde **Excel** usando **Python** con la librería **xlwings** y el modelo **Prophet**.

## Requisitos
- Windows + Excel de escritorio (Office 365 o 2019+).
- Python 3.10 – 3.13 (recomendado 3.11/3.12).
- Paquetes: xlwings, pandas, numpy, matplotlib, prophet, cmdstanpy.

## Instalación rápida
```bash
pip install xlwings pandas numpy matplotlib prophet cmdstanpy
python -c "import cmdstanpy; cmdstanpy.install_cmdstan()"
```

## Archivos del proyecto
- `forecast_demo.py` → lógica del modelo y generación de hojas.
- `fcst_file.xlsx.xlsm` → libro Excel con botón y macro para llamar a Python.

## Configuración en Excel
1. Habilitar complemento **xlwings** en Excel.
2. En la cinta de xlwings configurar la ruta a `python.exe`.
3. Marcar “Add workbook to PYTHONPATH”.
4. Crear un botón y asignarle el macro:
```vb
Sub Generar_FCST()
    RunPython "import forecast_demo; forecast_demo.run_from_selection()"
End Sub
```

## Datos de entrada en Excel
Parámetros (A:B):
```
weekly seasonality | False
yearly seasonality | True
periods            | 12
freq               | W
sensibility        | 0.8
output_mode        | ambos
...
```
Tabla de ventas:
```
Producto | Fecha       | Ventas
---------|-------------|-------
Prod A   | 2024-01-07  | 96
Prod A   | 2024-01-14  | 112.67
...
```

## Salida
- Hojas por producto con histórico, forecast y gráficas.
- Hoja de **Resumen** con:
  - Forecast total sumado de todos los productos.
  - Promedios (histórico, últimas N semanas, mismo periodo año anterior).
  - Crecimientos calculados.
  - Gráficas de histórico + forecast.

## Errores comunes
- “Could not find Interpreter!” → revisar ruta de Python en cinta xlwings.
- `ModuleNotFoundError: forecast_demo` → colocar `.py` y `.xlsm` en misma carpeta y activar PYTHONPATH.
- Primer corrida lenta → Prophet instala `cmdstan`.

