# IA Resumen Bancario – Banco Macro

App Streamlit independiente para procesar PDFs de **Últimos Movimientos** de Banco Macro.

## Funciones

- Procesa uno o varios PDFs.
- Detecta cuenta, empresa, período y movimientos.
- Muestra importes completos sin recortes visuales.
- Genera resumen operativo.
- Exporta:
  - PDF de resumen operativo.
  - Excel general con resumen y movimientos.
  - Excel orientado a importación Holistor.

## Instalación local

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Estructura

```text
app.py
assets/
  logo_aie.png
  favicon-aie.ico
modules/
  extraction.py
  parsing.py
  reports.py
requirements.txt
README.md
```

## Notas

El parser está preparado para PDFs con columnas:

- Fecha
- Referencia
- Causal
- Concepto
- Importe
- Saldo

Para gastos e impuestos, la app respeta el signo del PDF: los débitos negativos suman como gasto y los créditos o ajustes positivos restan.
