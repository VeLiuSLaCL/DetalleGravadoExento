import pandas as pd

def transformar_nomina(archivo):
    df = pd.read_excel(archivo, sheet_name=0)

    df.columns = [str(col).strip() for col in df.columns]

    col_concepto_tipo = df.columns[13]
    col_concepto = df.columns[18]
    col_exento = df.columns[20]
    col_gravado = df.columns[21]

    columnas_base = df.columns[:13].tolist()

    def procesar(df_filtrado):
        exento = df_filtrado.pivot_table(
            index=columnas_base,
            columns=col_concepto,
            values=col_exento,
            aggfunc='sum',
            fill_value=0
        )

        gravado = df_filtrado.pivot_table(
            index=columnas_base,
            columns=col_concepto,
            values=col_gravado,
            aggfunc='sum',
            fill_value=0
        )

        exento.columns = [f"{col} EXENTO" for col in exento.columns]
        gravado.columns = [f"{col} GRAVADO" for col in gravado.columns]

        resultado = pd.concat([exento, gravado], axis=1).reset_index()

        cols_exento = [c for c in resultado.columns if "EXENTO" in c]
        cols_gravado = [c for c in resultado.columns if "GRAVADO" in c]

        resultado["TOTAL_EXENTO"] = resultado[cols_exento].sum(axis=1)
        resultado["TOTAL_GRAVADO"] = resultado[cols_gravado].sum(axis=1)

        return resultado

    percepciones = df[df[col_concepto_tipo].str.contains("PERCEPC", case=False, na=False)]
    deducciones = df[df[col_concepto_tipo].str.contains("DEDUC", case=False, na=False)]

    return procesar(percepciones), procesar(deducciones)
