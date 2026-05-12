import io
from pathlib import Path

import pandas as pd
import streamlit as st

from modules.extraction import text_from_pdf
from modules.parsing import macro_ultimos_movimientos_extract
from modules.reports import (
    build_operational_summary,
    build_bank_reconciliation,
    build_credit_detail,
    fmt_money,
    make_excel,
    make_holistor_excel,
    make_credit_detail_excel,
    make_operational_summary_pdf,
)

HERE = Path(__file__).parent
LOGO = HERE / "assets" / "logo_aie.png"
FAVICON = HERE / "assets" / "favicon-aie.ico"

st.set_page_config(
    page_title="IA Resumen Bancario – Banco Macro",
    page_icon=str(FAVICON) if FAVICON.exists() else None,
    layout="wide",
)

st.markdown(
    """
    <style>
    .block-container {max-width: 1180px; padding-top: 2rem;}
    .aie-metric-card {
        border: 1px solid rgba(128,128,128,.25);
        border-radius: 12px;
        padding: 18px 20px;
        min-height: 112px;
        background: rgba(128,128,128,.06);
    }
    .aie-metric-label {
        font-size: 0.92rem;
        font-weight: 700;
        margin-bottom: 10px;
        opacity: .9;
    }
    .aie-metric-value {
        font-size: clamp(1.45rem, 2.4vw, 2.15rem);
        line-height: 1.15;
        font-weight: 650;
        white-space: nowrap;
        overflow: visible;
    }
    div[data-testid="stDataFrame"] {width: 100%;}
    </style>
    """,
    unsafe_allow_html=True,
)


def metric_card(label: str, value: float) -> str:
    return f"""
    <div class=\"aie-metric-card\">
        <div class=\"aie-metric-label\">{label}</div>
        <div class=\"aie-metric-value\">$ {fmt_money(value)}</div>
    </div>
    """


if LOGO.exists():
    st.image(str(LOGO), width=210)

st.title("IA Resumen Bancario – Banco Macro")
st.caption("Versión para PDF de Últimos Movimientos | Uso interno AIE San Justo")

uploaded_files = st.file_uploader(
    "Subí uno o varios PDF de Últimos Movimientos Banco Macro",
    type=["pdf"],
    accept_multiple_files=True,
)

if not uploaded_files:
    st.info("La app no almacena datos. Toda la información se procesa en la sesión actual.")
    st.stop()

all_movs = []
file_results = []

for uploaded in uploaded_files:
    data = uploaded.read()
    text = text_from_pdf(io.BytesIO(data)).strip()

    if not text:
        file_results.append({"archivo": uploaded.name, "estado": "Sin texto detectable", "cuenta": "", "empresa": "", "movimientos": 0})
        continue

    result = macro_ultimos_movimientos_extract(text)
    df_file = result["movimientos"]

    if not df_file.empty:
        df_file["archivo"] = uploaded.name
        df_file["cuenta"] = result.get("cuenta", "s/n")
        df_file["empresa"] = result.get("empresa", "")
        all_movs.append(df_file)

    file_results.append(
        {
            "archivo": uploaded.name,
            "estado": "Procesado" if not df_file.empty else "Sin movimientos detectados",
            "cuenta": result.get("cuenta", "s/n"),
            "empresa": result.get("empresa", ""),
            "movimientos": len(df_file),
        }
    )

st.subheader("Archivos procesados")
st.dataframe(pd.DataFrame(file_results), use_container_width=True, hide_index=True)

if not all_movs:
    st.error(
        "No se detectaron movimientos. Revisá si el PDF mantiene el formato con columnas "
        "Fecha, Referencia, Causal, Concepto, Importe y Saldo."
    )
    st.stop()

df = pd.concat(all_movs, ignore_index=True)
df = df.drop_duplicates(subset=["fecha", "referencia", "causal", "concepto", "importe", "saldo", "cuenta"])
sort_cols = [c for c in ["cuenta", "archivo", "orden_cronologico"] if c in df.columns]
if sort_cols:
    df = df.sort_values(sort_cols, ascending=True).reset_index(drop=True)

st.subheader("Datos detectados")
st.write(f"**Cuentas detectadas:** {df['cuenta'].nunique()}")
st.write(f"**Movimientos detectados:** {len(df)}")

fecha_desde = df["fecha"].min()
fecha_hasta = df["fecha"].max()
if pd.notna(fecha_desde) and pd.notna(fecha_hasta):
    st.write(f"**Período detectado:** {fecha_desde.strftime('%d/%m/%Y')} al {fecha_hasta.strftime('%d/%m/%Y')}")

total_creditos = df.loc[df["importe"] > 0, "importe"].sum()
total_debitos = -df.loc[df["importe"] < 0, "importe"].sum()
neto_movimientos = df["importe"].sum()

col1, col2, col3 = st.columns(3)
col1.markdown(metric_card("Créditos", total_creditos), unsafe_allow_html=True)
col2.markdown(metric_card("Débitos", total_debitos), unsafe_allow_html=True)
col3.markdown(metric_card("Neto", neto_movimientos), unsafe_allow_html=True)

st.subheader("Resumen operativo")
resumen = build_operational_summary(df)
st.dataframe(resumen[["Concepto", "Importe formateado"]], use_container_width=True, hide_index=True)

st.subheader("Conciliación bancaria")
conciliacion = build_bank_reconciliation(df)
conc_cols = [
    "Archivo", "Cuenta", "Desde", "Hasta", "Movimientos",
    "Saldo anterior calculado formateado", "Créditos formateado", "Débitos formateado",
    "Neto movimientos formateado", "Saldo final calculado formateado",
    "Saldo final informado formateado", "Diferencia formateado", "Estado",
]
conc_cols = [c for c in conc_cols if c in conciliacion.columns]
st.dataframe(conciliacion[conc_cols], use_container_width=True, hide_index=True)

st.subheader("Detalle de créditos / cuotas")
detalle_creditos = build_credit_detail(df)
if detalle_creditos.empty:
    st.info("No se detectaron movimientos vinculados a créditos, préstamos o pago de cuotas.")
else:
    dc_show = detalle_creditos.copy()
    if "fecha" in dc_show.columns:
        dc_show["fecha"] = dc_show["fecha"].dt.strftime("%d/%m/%Y")
    for col in ["importe", "saldo"]:
        if col in dc_show.columns:
            dc_show[col] = dc_show[col].apply(fmt_money)
    st.dataframe(dc_show, use_container_width=True, hide_index=True)

with st.expander("Ver movimientos detectados", expanded=False):
    df_show = df.copy()
    df_show["fecha"] = df_show["fecha"].dt.strftime("%d/%m/%Y")
    df_show["importe"] = df_show["importe"].apply(fmt_money)
    df_show["saldo"] = df_show["saldo"].apply(fmt_money)

    cols = [
        "archivo",
        "cuenta",
        "fecha",
        "referencia",
        "causal",
        "concepto",
        "importe",
        "saldo",
        "categoria",
    ]
    st.dataframe(df_show[cols], use_container_width=True, hide_index=True)

suffix = "macro_ultimos_movimientos"
if df["cuenta"].nunique() == 1:
    suffix = f"macro_ultimos_movimientos_{str(df['cuenta'].iloc[0]).replace('/', '-') }"

excel_bytes = make_excel(df, resumen, conciliacion)
holistor_bytes = make_holistor_excel(df)
credit_detail_bytes = make_credit_detail_excel(df)
pdf_bytes = make_operational_summary_pdf(df, resumen, conciliacion, logo_path=LOGO if LOGO.exists() else None)

st.subheader("Descargas")
d1, d2, d3, d4 = st.columns(4)

d1.download_button(
    "Descargar PDF resumen operativo",
    data=pdf_bytes,
    file_name=f"{suffix}_resumen_operativo.pdf",
    mime="application/pdf",
    use_container_width=True,
)

d2.download_button(
    "Descargar Excel general",
    data=excel_bytes,
    file_name=f"{suffix}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)

d3.download_button(
    "Descargar Excel Holistor",
    data=holistor_bytes,
    file_name=f"{suffix}_holistor.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)

d4.download_button(
    "Descargar detalle créditos",
    data=credit_detail_bytes,
    file_name=f"{suffix}_detalle_creditos.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)

st.caption("Herramienta para uso interno AIE San Justo | Developer Alfonso Alderete")
