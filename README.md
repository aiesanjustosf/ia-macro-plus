# IA Resumen Bancario – Banco Macro

App independiente en Streamlit para procesar PDFs de **Banco Macro - Últimos Movimientos**.

## Funciones

- Carga de uno o varios PDFs.
- Detección de cuenta, empresa y movimientos.
- Resumen operativo por conceptos.
- Conciliación bancaria:
  - reconstruye el saldo anterior porque el PDF no lo informa;
  - calcula `saldo anterior = saldo del primer movimiento cronológico - importe del primer movimiento`;
  - concilia contra el saldo final informado por el último movimiento cronológico.
- Descarga de PDF de resumen operativo.
- Descarga de Excel general con hojas:
  - Resumen operativo;
  - Conciliación bancaria;
  - Movimientos.
- Descarga de Excel Holistor.

## Holistor

Ajustes incluidos:

- Columna **Cód** agregada antes de **Exento / No Gravado**.
- La columna **Cód** se completa con `584`.
- Código neto actualizado de `524` a `506`.
- Código de SIRCREB actualizado de `P006` a `SIRC`.

## Instalación local

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy en Streamlit Cloud

1. Subir el repositorio a GitHub.
2. Crear una app nueva en Streamlit Cloud.
3. Seleccionar `app.py` como archivo principal.

---

Herramienta para uso interno AIE San Justo.
