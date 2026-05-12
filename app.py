import io
from pathlib import Path

import pandas as pd
import streamlit as st

from modules.extraction import text_from_pdf
from modules.parsing import macro_ultimos_movimientos_extract
from modules.reports import build_operational_summary, fmt_money, make_excel

HERE = Path(__file__).parent
LOGO = HERE / "assets" / "logo_aie.png"
FAVICON = HERE / "assets" / "favicon-aie.ico"

st.set_page_config(
    page_title="IA Resumen Bancario – Banco Macro",
    page_icon=str(FAVICON) if FAVICON.exists() else None,
    layout="centered",
)

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
        file_results.append({"archivo": uploaded.name, "estado": "Sin texto detectable", "movimientos": 0})
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
df = df.sort_values(["cuenta", "fecha", "referencia"], ascending=[True, True, True]).reset_index(drop=True)

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
col1.metric("Créditos", f"$ {fmt_money(total_creditos)}")
col2.metric("Débitos", f"$ {fmt_money(total_debitos)}")
col3.metric("Neto", f"$ {fmt_money(neto_movimientos)}")

st.subheader("Resumen operativo")
resumen = build_operational_summary(df)
st.dataframe(resumen[["Concepto", "Importe formateado"]], use_container_width=True, hide_index=True)

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

excel_bytes = make_excel(df, resumen)

suffix = "macro_ultimos_movimientos"
if df["cuenta"].nunique() == 1:
    suffix = f"macro_ultimos_movimientos_{str(df['cuenta'].iloc[0]).replace('/', '-') }"

st.download_button(
    "Descargar Excel",
    data=excel_bytes,
    file_name=f"{suffix}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.caption("Herramienta para uso interno AIE San Justo | Developer Alfonso Alderete")
