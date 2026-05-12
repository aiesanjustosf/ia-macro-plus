import re
import pandas as pd


def _macro_parse_amount(value: str) -> float:
    if value is None:
        return 0.0
    value = str(value).replace("$", "").replace(" ", "").strip()
    value = value.replace(".", "").replace(",", ".")
    try:
        return float(value)
    except ValueError:
        return 0.0


def _macro_classify_movement(concepto: str) -> str:
    c = concepto.upper()

    if "25413" in c:
        return "Ley 25.413"
    if "SIRCREB" in c or "IIBB SANTA FE" in c or "AJ IIBB" in c or "IIBB STAFE" in c:
        return "SIRCREB / IIBB"
    if "DEBITO FISCAL IVA BASICO" in c:
        return "IVA básico"
    if "DEBITO FISCAL IVA PERCEPCION" in c:
        return "IVA percepción"
    if "COMISION" in c or "MANTENIMIENTO" in c or "INTER.ADEL" in c or "INTER ADEL" in c:
        return "Gasto bancario"
    if "AFIP" in c or "ARCA" in c:
        return "AFIP / ARCA"
    if "CHEQUE" in c:
        return "Cheques"
    if "TARJETA" in c:
        return "Tarjetas"
    if "PRESTAMOS" in c or "PRESTAMO" in c:
        return "Préstamos"
    if "PAGO REMUNERACIONES" in c:
        return "Sueldos / Remuneraciones"
    if "CREDIN" in c:
        return "Créditos / CREDIN"
    if "ING TRANSF" in c:
        return "Transferencias recibidas"
    if "TRANSF" in c or "TRF" in c:
        return "Transferencias"
    return "Otros"


def _extract_account(lines: list[str]) -> str:
    # En Últimos Movimientos suele aparecer una línea con el número largo de cuenta.
    for ln in lines:
        clean = ln.strip()
        if re.fullmatch(r"\d{10,}", clean):
            return clean
    return "s/n"


def _extract_company(lines: list[str]) -> str:
    for ln in lines:
        if ln.upper().startswith("EMPRESA:"):
            return ln.split(":", 1)[1].strip()
    return ""


def macro_ultimos_movimientos_extract(text: str) -> dict:
    """
    Parser para Banco Macro - Últimos Movimientos.

    Formato de línea esperado:
    30/04/2026 1959071229 1685 IMPDBCR 25413 S/CR TASA GRAL $ 12.000,00 $ -162.635,57
    """
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]

    pattern = re.compile(
        r"^"
        r"(?P<fecha>\d{2}/\d{2}/\d{4})\s+"
        r"(?P<referencia>\d+)\s+"
        r"(?P<causal>\d+)\s+"
        r"(?P<concepto>.+?)\s+"
        r"\$\s*(?P<importe>-?\d{1,3}(?:\.\d{3})*,\d{2})\s+"
        r"\$\s*(?P<saldo>-?\d{1,3}(?:\.\d{3})*,\d{2})"
        r"$"
    )

    rows = []
    orden_pdf = 0
    for ln in lines:
        match = pattern.match(ln)
        if not match:
            continue

        concepto = match.group("concepto").strip()
        importe = _macro_parse_amount(match.group("importe"))
        saldo = _macro_parse_amount(match.group("saldo"))

        rows.append(
            {
                "orden_pdf": orden_pdf,
                "fecha": pd.to_datetime(match.group("fecha"), format="%d/%m/%Y", errors="coerce"),
                "referencia": match.group("referencia"),
                "causal": match.group("causal"),
                "concepto": concepto,
                "concepto_norm": concepto.upper(),
                "importe": importe,
                "saldo": saldo,
                "categoria": _macro_classify_movement(concepto),
            }
        )
        orden_pdf += 1

    df = pd.DataFrame(rows)
    if not df.empty:
        # El PDF de Últimos Movimientos viene normalmente en orden descendente.
        # Conservamos orden_pdf para la conciliación y dejamos la vista en orden cronológico.
        df = df.sort_values(["orden_pdf"], ascending=[False]).reset_index(drop=True)
        df["orden_cronologico"] = range(len(df))

    return {
        "cuenta": _extract_account(lines),
        "empresa": _extract_company(lines),
        "movimientos": df,
    }
