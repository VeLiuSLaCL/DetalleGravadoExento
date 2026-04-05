from __future__ import annotations

from io import BytesIO
from pathlib import Path

import pandas as pd


SOURCE_SHEET = "NOM MAR"
OUTPUT_SHEET = "NOM MAR TRANSFORMADA"


def load_nomina_dataframe(file_path_or_buffer, source_sheet: str = SOURCE_SHEET) -> pd.DataFrame:
    df = pd.read_excel(file_path_or_buffer, sheet_name=source_sheet, dtype=object)
    df = df.dropna(how="all").copy()
    return df


def transform_nomina_dataframe(
    df: pd.DataFrame,
    fixed_cols_count: int = 13,
    concept_col: str | None = None,
    exento_col: str | None = None,
    gravado_col: str | None = None,
) -> pd.DataFrame:
    cols = list(df.columns)

    if len(cols) < 22:
        raise ValueError("La hoja no tiene la estructura esperada. Se esperaban al menos 22 columnas.")

    # Recomendado: A:M
    fixed_cols = cols[:fixed_cols_count]

    # Según tu archivo:
    concept_col = concept_col or cols[18]   # S = Texto expl.CC-nómina
    exento_col = exento_col or cols[20]     # U = Exento
    gravado_col = gravado_col or cols[21]   # V = Gravado

    required = fixed_cols + [concept_col, exento_col, gravado_col]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Faltan columnas requeridas: {missing}")

    work = df.copy()
    work[concept_col] = work[concept_col].fillna("").astype(str).str.strip()
    work = work[work[concept_col] != ""].copy()

    work[exento_col] = pd.to_numeric(work[exento_col], errors="coerce").fillna(0)
    work[gravado_col] = pd.to_numeric(work[gravado_col], errors="coerce").fillna(0)

    concept_order = list(dict.fromkeys(work[concept_col].tolist()))
    work[concept_col] = pd.Categorical(work[concept_col], categories=concept_order, ordered=True)

    exento_wide = (
        work.groupby(fixed_cols + [concept_col], dropna=False, observed=True)[exento_col]
        .sum()
        .unstack(concept_col, fill_value=0)
        .reindex(columns=concept_order, fill_value=0)
    )

    gravado_wide = (
        work.groupby(fixed_cols + [concept_col], dropna=False, observed=True)[gravado_col]
        .sum()
        .unstack(concept_col, fill_value=0)
        .reindex(columns=concept_order, fill_value=0)
    )

    base = exento_wide.index.to_frame(index=False).reset_index(drop=True)
    blocks = [base]

    for concept in concept_order:
        blocks.append(
            exento_wide[[concept]]
            .rename(columns={concept: f"{concept} Exento"})
            .reset_index(drop=True)
        )
        blocks.append(
            gravado_wide[[concept]]
            .rename(columns={concept: f"{concept} Gravado"})
            .reset_index(drop=True)
        )

    result = pd.concat(blocks, axis=1)
    result.iloc[:, len(fixed_cols):] = result.iloc[:, len(fixed_cols):].round(2)
    return result


def build_output_workbook(
    input_file_path_or_buffer,
    output_path: str | Path | None = None,
    source_sheet: str = SOURCE_SHEET,
    output_sheet: str = OUTPUT_SHEET,
    fixed_cols_count: int = 13,
) -> BytesIO:
    df_raw = load_nomina_dataframe(input_file_path_or_buffer, source_sheet=source_sheet)
    df_out = transform_nomina_dataframe(df_raw, fixed_cols_count=fixed_cols_count)

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df_out.to_excel(writer, sheet_name=output_sheet, index=False)

        workbook = writer.book
        worksheet = writer.sheets[output_sheet]

        header_fmt = workbook.add_format({"bold": True, "bg_color": "#D9EAF7"})
        money_fmt = workbook.add_format({"num_format": "#,##0.00"})
        text_fmt = workbook.add_format({"num_format": "@"})

        rows, cols = df_out.shape
        worksheet.freeze_panes(1, 0)
        worksheet.autofilter(0, 0, rows, cols - 1)

        for col_idx, col_name in enumerate(df_out.columns):
            sample = df_out.iloc[:200, col_idx].fillna("").astype(str).tolist()
            max_len = max([len(str(col_name))] + [len(v) for v in sample])
            width = min(max(max_len + 2, 12), 28)

            fmt = money_fmt if col_idx >= fixed_cols_count else None
            worksheet.set_column(col_idx, col_idx, width, fmt)
            worksheet.write(0, col_idx, col_name, header_fmt)

        for name in ["Número de personal", "Folio CFDi", "UUID"]:
            if name in df_out.columns:
                idx = df_out.columns.get_loc(name)
                worksheet.set_column(idx, idx, 18, text_fmt)

    buffer.seek(0)

    if output_path:
        Path(output_path).write_bytes(buffer.getvalue())

    return buffer


def transform_file(
    input_path: str | Path,
    output_path: str | Path,
    source_sheet: str = SOURCE_SHEET,
    output_sheet: str = OUTPUT_SHEET,
    fixed_cols_count: int = 13,
) -> Path:
    build_output_workbook(
        input_file_path_or_buffer=input_path,
        output_path=output_path,
        source_sheet=source_sheet,
        output_sheet=output_sheet,
        fixed_cols_count=fixed_cols_count,
    )
    return Path(output_path)


if __name__ == "__main__":
    input_file = Path("1.xlsx")
    output_file = Path("1_transformado.xlsx")

    if not input_file.exists():
        raise SystemExit("No encontré 1.xlsx en la carpeta actual.")

    transform_file(input_file, output_file, fixed_cols_count=13)
    print(f"Archivo generado: {output_file}")
