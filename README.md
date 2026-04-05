# Transformador de nómina

## Qué hace
- Lee la hoja `NOM MAR`
- Usa la columna **N = CONCEPTO** para separar en:
  - `PERCEPCIONES`
  - `DEDUCCIONES`
- Usa la columna **S = Texto expl.CC-nómina** como nombre del concepto
- Convierte conceptos de filas a columnas
- Para cada concepto crea 2 columnas, en este orden:
  - `CONCEPTO EXENTO`
  - `CONCEPTO GRAVADO`
- Usa:
  - **U** como exento
  - **V** como gravado
- Agrega al final:
  - `TOTAL_EXENTO`
  - `TOTAL_GRAVADO`

## Archivos
- `app.py`: interfaz en Streamlit
- `transform_nomina.py`: proceso por script
- `requirements.txt`

## Ejecutar en local
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Ejecutar por script
```bash
python transform_nomina.py archivo.xlsx
```
