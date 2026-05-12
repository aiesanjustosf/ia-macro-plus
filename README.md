# IA Resumen Bancario – Banco Macro | Últimos Movimientos

App independiente en Streamlit para procesar PDFs de **Banco Macro – Últimos Movimientos**.

## Funciones

- Carga de uno o varios PDFs.
- Detección de cuenta, empresa, período y movimientos.
- Resumen operativo.
- Conciliación bancaria.
- Reconstrucción de saldo anterior aunque el PDF no lo informe.
- Cálculo correcto tomando como inicio el movimiento con **fecha más antigua** y como cierre el movimiento con **fecha más reciente**.
- Descarga de PDF de resumen operativo.
- Descarga de Excel general.
- Descarga de Excel Holistor.
- Separación automática de gastos bancarios al 21% y al 10,5%.
- Descarga de Excel con detalle de créditos, préstamos y pago de cuotas.

## Conciliación bancaria

Banco Macro no informa saldo anterior en este formato. La app lo calcula así:

```text
Saldo anterior calculado = saldo del primer movimiento cronológico - importe del primer movimiento cronológico
```

Para evitar errores cuando el PDF viene listado de más nuevo a más viejo, la app reconstruye el orden por cadena de saldos y, como respaldo, toma la fecha más antigua como inicio y la fecha más reciente como cierre.

## Excel Holistor

Ajustes incluidos:

- Columna `Cód` antes de `Exento / No Gravado`.
- `Cód = 584`.
- Código de neto gasto al 21%: `506`.
- Código de neto gasto al 10,5%: `604`.
- SIRCREB: `SIRC`.

## Instalación

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Uso interno

Herramienta para uso interno AIE San Justo.
