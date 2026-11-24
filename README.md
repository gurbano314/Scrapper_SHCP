# Conciliador de Cotizaciones v6.6 (EFIDEPORTE · SHCP-UPIT)

Aplicación de escritorio nativa para Windows (PyQt6) que automatiza la segmentación, extracción y conciliación de montos económicos en expedientes PDF de cotizaciones de proveedores y presupuestos globales.

## Características principales

- **Auto-detección de secciones (PDF único):** Identifica el *Presupuesto Global* y cada cotización individual (por marcadores PDF y heurísticas visuales/OCR ligero), sugiriendo desmarcar resúmenes financieros y cartas manifiesto EFIDEPORTE (Anexo 29 y 34–40).
- **Motor multinivel y OCR avanzado de doble pasada:** 5 estrategias de análisis (tablas nativas, triangulación, OCR inteligente PyTesseract + RapidOCR con preprocesamiento para fondos oscuros, búsqueda textual y pares adyacentes) y casilla opcional *Forzar OCR de alta precisión*.
- **Validación normativa EFIDEPORTE en tiempo real:** Audita automáticamente 7 reglas (vigencia ≤ 90 días, tope 5% administrativo, rubros prohibidos, duplicados, seguro obligatorio, *PII solo compra de equipo* y recordatorio CFDI), actualizándose en vivo al editar la tabla o parámetros del proyecto.
- **API Banxico integrada:** Conversión automática de USD, EUR y CAD según fecha de emisión y actualización en caliente con el botón *Aplicar*.
- **Exportación a Excel y proyectos `.cshcp`:** Reportes `.xlsx` en formato Carta Horizontal con fórmulas vivas (`ROUND`, `SUMIF`) y guardado integral de sesiones en archivos `.cshcp`.

## Estructura del proyecto

- `src/engine.py` — Motor puro de extracción, OCR, validación normativa, API Banxico y exportación Excel.
- `src/main_qt.py` — Interfaz gráfica nativa PyQt6 (panel lateral fijo de 310 px, visor PDF, editor de datos y diálogo de auto-detección).
- `tests/test_conciliador.py` — Pruebas automatizadas de regresión (motor, validación normativa, Excel e interfaz PyQt6).
- `docs/manual.tex` / `docs/manual.pdf` — Manual de usuario v6.6.
- `docs/presentacion/` — Presentación ejecutiva en Beamer (`presentacion.tex` / `presentacion.pdf`).
- `conciliador_qt.spec` y `build_qt.bat` — Configuración para empaquetar el ejecutable único `dist/Conciliador.exe`.

## Requisitos y ejecución desde código fuente

```powershell
python -m venv build_venv_qt
.\build_venv_qt\Scripts\Activate.ps1
pip install -r requirements_qt.txt
python src\main_qt.py
```

## Pruebas automatizadas

```powershell
python -m unittest discover -s tests -v
```

## Compilación de ejecutable único (`.exe`)

Coloca Tesseract portable en `bin/tesseract/` si deseas incluirlo dentro del paquete y ejecuta:

```cmd
build_qt.bat
```

Para automatización sin abrir el Explorador de Windows, utiliza `build_qt.bat --ci`. El archivo resultante se genera en `dist\Conciliador.exe`.
