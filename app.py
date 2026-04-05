from io import BytesIO

import streamlit as st

from transform_nomina import build_output_workbook, load_nomina_dataframe, transform_nomina_dataframe


st.set_page_config(page_title="Transformador de Nómina", layout="wide")
st.title("Transformador de nómina")
st.caption("Convierte conceptos de filas a columnas usando 'Texto expl.CC-nómina' como clave.")

with st.sidebar:
    st.header("Configuración")
    fixed_cols_count = st.selectbox(
        "Columnas fijas al inicio",
        options=[13, 18],
        index=0,
        help=(
            "13 es lo recomendado porque A:M suelen ser los datos estables del recibo. "
            "18 solo úsalo si quieres forzarlo."
        ),
    )
    source_sheet = st.text_input("Hoja origen", value="NOM MAR")
    output_sheet = st.text_input("Hoja salida", value="NOM MAR TRANSFORMADA")

uploaded_file = st.file_uploader("Sube tu archivo Excel", type=["xlsx"])

if uploaded_file is not None:
    try:
        df_raw = load_nomina_dataframe(BytesIO(uploaded_file.getvalue()), source_sheet=source_sheet)
        df_preview = transform_nomina_dataframe(df_raw, fixed_cols_count=fixed_cols_count)

        c1, c2, c3 = st.columns(3)
        c1.metric("Filas originales", f"{len(df_raw):,}")
        c2.metric("Filas transformadas", f"{len(df_preview):,}")
        c3.metric("Columnas transformadas", f"{df_preview.shape[1]:,}")

        st.subheader("Vista previa")
        st.dataframe(df_preview.head(50), use_container_width=True)

        excel_bytes = build_output_workbook(
            input_file_path_or_buffer=BytesIO(uploaded_file.getvalue()),
            source_sheet=source_sheet,
            output_sheet=output_sheet,
            fixed_cols_count=fixed_cols_count,
        ).getvalue()

        st.download_button(
            label="Descargar archivo transformado",
            data=excel_bytes,
            file_name="nomina_transformada.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.error(f"Ocurrió un error: {e}")
else:
    st.info("Sube un archivo .xlsx para comenzar.")
