import streamlit as st

st.set_page_config(page_title="Prueba de carga", layout="centered")

st.title("Prueba básica de archivo")

archivo = st.file_uploader("Sube un archivo Excel", type=["xlsx", "xls"])

if archivo is not None:
    st.success(f"Archivo recibido: {archivo.name}")
    st.write(f"Tamaño: {archivo.size} bytes")
