import io
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


def fmt_money(value: float) -> str:
    try:
        return f"{float(value):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "0,00"


def _neto(df: pd.DataFrame, mask: pd.Series) -> float:
    # Para gastos/impuestos: débitos negativos suman como costo; créditos positivos restan.
    return -df.loc[mask, "importe"].sum()


def build_operational_summary(df: pd.DataFrame) -> pd.DataFrame:
    concepto = df["concepto_norm"].fillna("")

    m_25413 = concepto.str.contains("25413", na=False)
    m_25413_db = m_25413 & concepto.str.contains("S/DB|S DB|DB TASA", regex=True, na=False)
    m_25413_cr = m_25413 & concepto.str.contains("S/CR|S CR|CR TASA", regex=True, na=False)
    m_sircreb = concepto.str.contains("SIRCREB|IIBB SANTA FE|AJ IIBB|IIBB STAFE", regex=True, na=False)
    m_iva_basico = concepto.str.contains("DEBITO FISCAL IVA BASICO", na=False)
    m_iva_percepcion = concepto.str.contains("DEBITO FISCAL IVA PERCEPCION", na=False)
    m_comisiones = concepto.str.contains("COMISION|MANTENIMIENTO|INTER.ADEL|INTER ADEL", regex=True, na=False) & ~m_iva_basico & ~m_iva_percepcion
    m_afip_arca = concepto.str.contains("AFIP|ARCA", regex=True, na=False)
    m_cheques = concepto.str.contains("CHEQUE", na=False)
    m_tarjetas = concepto.str.contains("TARJETA", na=False)
    m_prestamos = concepto.str.contains("PRESTAMOS|PRESTAMO", regex=True, na=False)
    m_sueldos = concepto.str.contains("PAGO REMUNERACIONES", na=False)

    rows = [
        {"Concepto": "Ley 25.413 / Débitos y Créditos - Neto", "Importe": _neto(df, m_25413), "incluye_total": True},
        {"Concepto": "Ley 25.413 S/DB", "Importe": _neto(df, m_25413_db), "incluye_total": False},
        {"Concepto": "Ley 25.413 S/CR", "Importe": _neto(df, m_25413_cr), "incluye_total": False},
        {"Concepto": "SIRCREB / IIBB Santa Fe - Neto", "Importe": _neto(df, m_sircreb), "incluye_total": True},
        {"Concepto": "Gastos / Comisiones bancarias", "Importe": _neto(df, m_comisiones), "incluye_total": True},
        {"Concepto": "IVA básico sobre gastos bancarios", "Importe": _neto(df, m_iva_basico), "incluye_total": True},
        {"Concepto": "IVA percepción", "Importe": _neto(df, m_iva_percepcion), "incluye_total": True},
        {"Concepto": "AFIP / ARCA", "Importe": _neto(df, m_afip_arca), "incluye_total": True},
        {"Concepto": "Cheques", "Importe": _neto(df, m_cheques), "incluye_total": True},
        {"Concepto": "Tarjetas", "Importe": _neto(df, m_tarjetas), "incluye_total": True},
        {"Concepto": "Préstamos", "Importe": _neto(df, m_prestamos), "incluye_total": True},
        {"Concepto": "Sueldos / Remuneraciones", "Importe": _neto(df, m_sueldos), "incluye_total": True},
    ]

    resumen = pd.DataFrame(rows)
    resumen = resumen[resumen["Importe"].round(2) != 0].copy()
    total = resumen.loc[resumen["incluye_total"], "Importe"].sum()
    resumen = resumen.drop(columns=["incluye_total"])
    resumen.loc[len(resumen)] = {
        "Concepto": "TOTAL RESUMEN OPERATIVO",
        "Importe": total,
    }
    resumen["Importe formateado"] = resumen["Importe"].apply(fmt_money)
    return resumen


def _style_sheet(ws):
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    total_fill = PatternFill("solid", fgColor="D9EAF7")
    thin = Side(style="thin", color="D9D9D9")

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")
        cell.border = Border(bottom=thin)

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.border = Border(bottom=thin)
            if isinstance(cell.value, (int, float)):
                cell.number_format = '#,##0.00'
        if row[0].value and "TOTAL" in str(row[0].value).upper():
            for cell in row:
                cell.fill = total_fill
                cell.font = Font(bold=True)

    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            max_len = max(max_len, len(str(cell.value)) if cell.value is not None else 0)
        ws.column_dimensions[col_letter].width = min(max(max_len + 2, 12), 55)

    ws.freeze_panes = "A2"


def make_excel(df_mov: pd.DataFrame, df_resumen: pd.DataFrame) -> bytes:
    output = io.BytesIO()

    df_mov_export = df_mov.copy()
    if "fecha" in df_mov_export.columns:
        df_mov_export["fecha"] = df_mov_export["fecha"].dt.strftime("%d/%m/%Y")

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_resumen.to_excel(writer, sheet_name="Resumen operativo", index=False)
        df_mov_export.to_excel(writer, sheet_name="Movimientos", index=False)

        for ws in writer.book.worksheets:
            _style_sheet(ws)

    return output.getvalue()
