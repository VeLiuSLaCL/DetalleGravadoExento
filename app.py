import streamlit as st
import pandas as pd
from transform_nomina import transformar_nomina
from io import BytesIO

st.set_page_config(page_title="Transformador Nómina", layout="wide")

st.title("Transformador de Nómina (Exento / Gravado)")

archivo = st.file_uploader("Sube tu archivo Excel", type=["xlsx"])

if archivo:
    try:
        percepciones, deducciones = transformar_nomina(archivo)

        st.success("Archivo procesado correctamente")

        st.subheader("Vista previa - Percepciones")
        st.dataframe(percepciones.head())

        st.subheader("Vista previa - Deducciones")
        st.dataframe(deducciones.head())

        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            percepciones.to_excel(writer, index=False, sheet_name='PERCEPCIONES')
            deducciones.to_excel(writer, index=False, sheet_name='DEDUCCIONES')

        st.download_button(
            label="Descargar archivo transformado",
            data=output.getvalue(),
            file_name="nomina_transformada.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error al procesar: {e}")
