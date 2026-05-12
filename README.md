# IA Resumen Bancario – Banco Macro

App Streamlit para procesar PDFs de **Banco Macro - Últimos Movimientos**.

## Funcionalidades

- Carga de uno o varios PDFs.
- Extracción de movimientos con columnas: fecha, referencia, causal, concepto, importe y saldo.
- Resumen operativo por categoría.
- Tratamiento firmado de débitos y créditos.
- Exportación a Excel con hojas:
  - `Resumen operativo`
  - `Movimientos`

## Criterio de cálculo

Para impuestos y gastos se usa el signo real del PDF:

```text
Costo neto = -SUMA(importes firmados)
```

Por lo tanto:

- Débito negativo: suma como gasto.
- Crédito positivo / ajuste: resta del gasto.

Esto evita inflar conceptos como Ley 25.413 o SIRCREB cuando Banco Macro trae ajustes o devoluciones positivas.

## Instalación local

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy en Streamlit Cloud

1. Subir este repositorio a GitHub.
2. En Streamlit Cloud seleccionar el repositorio.
3. Main file path: `app.py`.
4. Deploy.

## Estructura

```text
ia-macro-ultimos-movimientos/
├── app.py
├── requirements.txt
├── README.md
├── assets/
│   ├── logo_aie.png
│   └── favicon-aie.ico
└── modules/
    ├── __init__.py
    ├── extraction.py
    ├── parsing.py
    └── reports.py
```
