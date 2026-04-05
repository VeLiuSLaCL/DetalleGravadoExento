# Transformador de nómina

## Qué hace
- Lee la hoja **NOM MAR**
- Usa **Texto expl.CC-nómina** como concepto
- Convierte cada concepto en 2 columnas:
  - `Concepto Exento` usando la columna **U**
  - `Concepto Gravado` usando la columna **V**
- Devuelve 1 fila consolidada por registro de nómina

## Archivos
- `transform_nomina.py`: lógica de transformación
- `app.py`: interfaz en Streamlit
- `requirements.txt`: dependencias

## Uso local
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Uso por script
Coloca tu archivo como `1.xlsx` en la misma carpeta y ejecuta:
```bash
python transform_nomina.py
```

## Nota
La configuración recomendada es **13 columnas fijas (A:M)**.
Si quieres forzar otra estructura, cambia `fixed_cols_count`.
