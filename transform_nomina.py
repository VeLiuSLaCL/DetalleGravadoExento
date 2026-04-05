from pathlib import Path
import sys
import pandas as pd

SHEET_SOURCE = "NOM MAR"
BASE_COLUMNS = [
    "Período cál.nómina",
    "Año de nómina",
    "Mes",
    "Nº de secuencia",
    "Número de personal",
    "Sociedad",
    "Área de nómina",
    "Tipo de nómina",
    "Identificador de nómina",
    "Motivo nóm.especial",
    "Nº ejecución contabil.",
    "Estado Impuesto Estatal",
    "Folio CFDi",
]
SPLIT_COLUMN = "CONCEPTO"
CONCEPT_NAME_COLUMN = "Texto expl.CC-nómina"
EXENTO_COLUMN = "U"
GRAVADO_COLUMN = "V"
OUTPUT_SHEETS = {"PERCEPCION": "PERCEPCIONES", "DEDUCCION": "DEDUCCIONES"}


def excel_col_letter(n: int) -> str:
    result = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


def normalizar_texto(valor) -> str:
    if pd.isna(valor):
        return ""
    texto = str(valor).strip()
    texto = " ".join(texto.split())
    return texto.upper()


def safe_numeric(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0.0)


def leer_excel(path_archivo: str) -> pd.DataFrame:
    df = pd.read_excel(path_archivo, sheet_name=SHEET_SOURCE, dtype=object)

    # Limpiar encabezados
    nuevas_cols = []
    for idx, col in enumerate(df.columns, start=1):
        nuevas_cols.append(f"COL_{idx}" if pd.isna(col) else str(col).strip())
    df.columns = nuevas_cols

    # Forzar que U y V sean accesibles por letra de Excel
    renombres = {}
    for idx, col in enumerate(df.columns, start=1):
        letra = excel_col_letter(idx)
        if letra == "U":
            renombres[col] = "U"
        elif letra == "V":
            renombres[col] = "V"
    df = df.rename(columns=renombres)

    requeridas = BASE_COLUMNS + [SPLIT_COLUMN, CONCEPT_NAME_COLUMN, EXENTO_COLUMN, GRAVADO_COLUMN]
    faltantes = [c for c in requeridas if c not in df.columns]
    if faltantes:
        raise ValueError(f"Faltan columnas requeridas: {', '.join(faltantes)}")

    return df


def transformar_bloque(df: pd.DataFrame, tipo_concepto: str) -> pd.DataFrame:
    bloque = df[df[SPLIT_COLUMN].astype(str).str.upper().str.strip() == tipo_concepto].copy()

    if bloque.empty:
        return pd.DataFrame(columns=BASE_COLUMNS + ["TOTAL_EXENTO", "TOTAL_GRAVADO"])

    bloque[CONCEPT_NAME_COLUMN] = bloque[CONCEPT_NAME_COLUMN].apply(normalizar_texto)
    bloque[EXENTO_COLUMN] = safe_numeric(bloque[EXENTO_COLUMN])
    bloque[GRAVADO_COLUMN] = safe_numeric(bloque[GRAVADO_COLUMN])

    for c in BASE_COLUMNS:
        bloque[c] = bloque[c].fillna("").astype(str).str.strip()

    pivot_exento = pd.pivot_table(
        bloque,
        index=BASE_COLUMNS,
        columns=CONCEPT_NAME_COLUMN,
        values=EXENTO_COLUMN,
        aggfunc="sum",
        fill_value=0,
    )
    pivot_gravado = pd.pivot_table(
        bloque,
        index=BASE_COLUMNS,
        columns=CONCEPT_NAME_COLUMN,
        values=GRAVADO_COLUMN,
        aggfunc="sum",
        fill_value=0,
    )

    conceptos = sorted(set(pivot_exento.columns.tolist()) | set(pivot_gravado.columns.tolist()))
    base_index = pivot_exento.index if len(pivot_exento.index) else pivot_gravado.index

    piezas = []
    for concepto in conceptos:
        if concepto in pivot_exento.columns:
            ex_df = pivot_exento[[concepto]].rename(columns={concepto: f"{concepto} EXENTO"})
        else:
            ex_df = pd.DataFrame({f"{concepto} EXENTO": 0}, index=base_index)

        if concepto in pivot_gravado.columns:
            gr_df = pivot_gravado[[concepto]].rename(columns={concepto: f"{concepto} GRAVADO"})
        else:
            gr_df = pd.DataFrame({f"{concepto} GRAVADO": 0}, index=base_index)

        piezas.append(ex_df)
        piezas.append(gr_df)

    resultado = pd.concat(piezas, axis=1).reset_index() if piezas else pd.DataFrame(columns=BASE_COLUMNS)

    dynamic_cols = []
    for concepto in conceptos:
        dynamic_cols.append(f"{concepto} EXENTO")
        dynamic_cols.append(f"{concepto} GRAVADO")

    exento_cols = [c for c in dynamic_cols if c.endswith(" EXENTO") and c in resultado.columns]
    gravado_cols = [c for c in dynamic_cols if c.endswith(" GRAVADO") and c in resultado.columns]

    resultado["TOTAL_EXENTO"] = resultado[exento_cols].sum(axis=1) if exento_cols else 0.0
    resultado["TOTAL_GRAVADO"] = resultado[gravado_cols].sum(axis=1) if gravado_cols else 0.0

    ordered_cols = BASE_COLUMNS + [c for c in dynamic_cols if c in resultado.columns] + ["TOTAL_EXENTO", "TOTAL_GRAVADO"]
    return resultado[ordered_cols]


def ajustar_hoja_excel(writer, sheet_name: str, df: pd.DataFrame):
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]

    header_fmt = workbook.add_format({
        "bold": True,
        "text_wrap": True,
        "valign": "vcenter",
        "fg_color": "#D9EAF7",
        "border": 1,
    })
    money_fmt = workbook.add_format({"num_format": "#,##0.00"})
    text_fmt = workbook.add_format({"num_format": "@"})

    worksheet.freeze_panes(1, 0)
    worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)

    for col_idx, col_name in enumerate(df.columns):
        worksheet.write(0, col_idx, col_name, header_fmt)
        sample = df[col_name].head(500).fillna("").astype(str).tolist() if len(df) else []
        max_len = max([len(str(col_name))] + [len(x) for x in sample]) if sample else len(str(col_name))
        width = min(max(max_len + 2, 12), 35)

        if col_name.endswith(" EXENTO") or col_name.endswith(" GRAVADO") or col_name.startswith("TOTAL_"):
            worksheet.set_column(col_idx, col_idx, max(width, 14), money_fmt)
        elif col_name in ["Período cál.nómina", "Año de nómina", "Mes", "Nº de secuencia", "Número de personal"]:
            worksheet.set_column(col_idx, col_idx, max(width, 12), text_fmt)
        else:
            worksheet.set_column(col_idx, col_idx, width)


def transformar_archivo(path_entrada: str, path_salida: str):
    df = leer_excel(path_entrada)
    percepciones = transformar_bloque(df, "PERCEPCION")
    deducciones = transformar_bloque(df, "DEDUCCION")

    with pd.ExcelWriter(path_salida, engine="xlsxwriter") as writer:
        percepciones.to_excel(writer, sheet_name=OUTPUT_SHEETS["PERCEPCION"], index=False)
        ajustar_hoja_excel(writer, OUTPUT_SHEETS["PERCEPCION"], percepciones)

        deducciones.to_excel(writer, sheet_name=OUTPUT_SHEETS["DEDUCCION"], index=False)
        ajustar_hoja_excel(writer, OUTPUT_SHEETS["DEDUCCION"], deducciones)

    return path_salida


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python transform_nomina.py archivo_entrada.xlsx [archivo_salida.xlsx]")
        sys.exit(1)

    entrada = sys.argv[1]
    salida = sys.argv[2] if len(sys.argv) >= 3 else str(Path(entrada).with_name(f"{Path(entrada).stem}_transformado.xlsx"))
    transformar_archivo(entrada, salida)
    print(f"Archivo generado: {salida}")
