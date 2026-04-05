from pathlib import Path
import io
import pandas as pd
import streamlit as st

from transform_nomina import leer_excel, transformar_bloque, ajustar_hoja_excel, OUTPUT_SHEETS

st.set_page_config(page_title="Transformador de nómina", layout="wide")
st.title("Transformador de nómina")
st.write(
    "Sube un Excel con la hoja **NOM MAR**. "
    "El proceso usa la columna **N (CONCEPTO)** para separar en **PERCEPCIONES** y **DEDUCCIONES**, "
    "toma **S** como nombre del concepto, usa **U** como EXENTO y **V** como GRAVADO, "
    "y agrega al final **TOTAL_EXENTO** y **TOTAL_GRAVADO**."
)

archivo = st.file_uploader("Sube tu archivo Excel", type=["xlsx", "xlsm", "xls"])

if archivo is not None:
    try:
        # guardar temporalmente el upload para relectura segura
        temp_path = Path("temp_nomina_upload.xlsx")
        temp_path.write_bytes(archivo.getvalue())

        df_origen = leer_excel(str(temp_path))
        df_per = transformar_bloque(df_origen, "PERCEPCION")
        df_ded = transformar_bloque(df_origen, "DEDUCCION")

        st.subheader("Vista rápida del origen")
        st.dataframe(df_origen.head(20), use_container_width=True)

        c1, c2 = st.columns(2)
        with c1:
            st.subheader("PERCEPCIONES")
            st.write(f"Filas: {len(df_per):,} | Columnas: {len(df_per.columns):,}")
            st.dataframe(df_per.head(20), use_container_width=True)

        with c2:
            st.subheader("DEDUCCIONES")
            st.write(f"Filas: {len(df_ded):,} | Columnas: {len(df_ded.columns):,}")
            st.dataframe(df_ded.head(20), use_container_width=True)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df_per.to_excel(writer, sheet_name=OUTPUT_SHEETS["PERCEPCION"], index=False)
            ajustar_hoja_excel(writer, OUTPUT_SHEETS["PERCEPCION"], df_per)

            df_ded.to_excel(writer, sheet_name=OUTPUT_SHEETS["DEDUCCION"], index=False)
            ajustar_hoja_excel(writer, OUTPUT_SHEETS["DEDUCCION"], df_ded)

        output.seek(0)

        st.download_button(
            "Descargar archivo transformado",
            data=output.getvalue(),
            file_name=f"{Path(archivo.name).stem}_transformado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.error(f"Ocurrió un error al procesar el archivo: {e}")
