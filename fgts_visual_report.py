#!/usr/bin/env python3
"""Gera um relatório visual em HTML a partir de dados tabulares (CSV/XLSX).

Uso:
  python fgts_visual_report.py --input dados.csv --output relatorio.html
"""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Iterable

import pandas as pd
import plotly.express as px


SUPPORTED_EXTENSIONS = {".csv", ".xlsx", ".xls"}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Trata dados e gera relatório visual em HTML para análise rápida."
    )
    parser.add_argument("--input", required=True, help="Arquivo de entrada (.csv/.xlsx/.xls)")
    parser.add_argument("--output", required=True, help="Arquivo HTML de saída")
    parser.add_argument(
        "--separator",
        default=",",
        help="Separador do CSV (padrão: ','). Ignorado para XLSX/XLS.",
    )
    parser.add_argument(
        "--encoding",
        default="utf-8",
        help="Encoding do CSV (padrão: utf-8). Ignorado para XLSX/XLS.",
    )
    parser.add_argument(
        "--date-columns",
        nargs="*",
        default=[],
        help="Colunas para tentativa explícita de conversão para data.",
    )
    return parser.parse_args()


def load_dataframe(path: Path, separator: str, encoding: str) -> pd.DataFrame:
    suffix = path.suffix.lower()
    if suffix not in SUPPORTED_EXTENSIONS:
        raise ValueError(
            f"Formato não suportado: {suffix}. Use um destes: {sorted(SUPPORTED_EXTENSIONS)}"
        )

    if suffix == ".csv":
        return pd.read_csv(path, sep=separator, encoding=encoding)

    return pd.read_excel(path)


def normalize_dataframe(df: pd.DataFrame, date_columns: Iterable[str]) -> pd.DataFrame:
    clean_df = df.copy()

    clean_df.columns = [str(c).strip() for c in clean_df.columns]

    for col in clean_df.select_dtypes(include="object").columns:
        clean_df[col] = clean_df[col].astype(str).str.strip()

    for col in date_columns:
        if col in clean_df.columns:
            clean_df[col] = pd.to_datetime(clean_df[col], errors="coerce")

    for col in clean_df.columns:
        if clean_df[col].dtype == "object":
            maybe_date = pd.to_datetime(clean_df[col], errors="coerce")
            if maybe_date.notna().mean() >= 0.85:
                clean_df[col] = maybe_date

    return clean_df


def dataframe_overview(df: pd.DataFrame) -> str:
    rows, cols = df.shape
    missing = int(df.isna().sum().sum())
    dup = int(df.duplicated().sum())
    return f"""
    <ul>
      <li><strong>Linhas:</strong> {rows}</li>
      <li><strong>Colunas:</strong> {cols}</li>
      <li><strong>Células faltantes:</strong> {missing}</li>
      <li><strong>Linhas duplicadas:</strong> {dup}</li>
    </ul>
    """


def generate_charts(df: pd.DataFrame) -> list[str]:
    charts_html: list[str] = []

    numeric_cols = list(df.select_dtypes(include="number").columns)
    datetime_cols = list(df.select_dtypes(include="datetime").columns)
    categorical_cols = [
        c
        for c in df.select_dtypes(include=["object", "category"]).columns
        if df[c].nunique(dropna=True) > 1
    ]

    for col in numeric_cols[:4]:
        fig = px.histogram(
            df,
            x=col,
            nbins=30,
            title=f"Distribuição de {col}",
            template="plotly_white",
        )
        charts_html.append(fig.to_html(full_html=False, include_plotlyjs="cdn"))

    for col in categorical_cols[:3]:
        top_values = df[col].fillna("(vazio)").value_counts().head(15).reset_index()
        top_values.columns = [col, "quantidade"]
        fig = px.bar(
            top_values,
            x=col,
            y="quantidade",
            title=f"Top categorias - {col}",
            template="plotly_white",
        )
        charts_html.append(fig.to_html(full_html=False, include_plotlyjs=False))

    if datetime_cols and numeric_cols:
        date_col = datetime_cols[0]
        num_col = numeric_cols[0]
        timeline = (
            df[[date_col, num_col]]
            .dropna()
            .sort_values(date_col)
            .groupby(date_col, as_index=False)[num_col]
            .sum()
        )
        if not timeline.empty:
            fig = px.line(
                timeline,
                x=date_col,
                y=num_col,
                title=f"Evolução temporal de {num_col}",
                template="plotly_white",
            )
            charts_html.append(fig.to_html(full_html=False, include_plotlyjs=False))

    return charts_html


def build_html_report(df: pd.DataFrame) -> str:
    head_rows = df.head(20).to_html(index=False, border=0)
    describe_table = df.describe(include="all", datetime_is_numeric=True).transpose().fillna("")
    describe_html = describe_table.to_html(border=0)
    charts = "\n".join(generate_charts(df))

    return f"""
<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Relatório Visual de Dados</title>
  <style>
    body {{ font-family: Arial, sans-serif; max-width: 1200px; margin: 24px auto; padding: 0 16px; }}
    h1, h2 {{ color: #123; }}
    table {{ width: 100%; border-collapse: collapse; margin-bottom: 16px; }}
    th, td {{ border: 1px solid #ddd; padding: 8px; font-size: 13px; text-align: left; }}
    th {{ background: #f4f6f8; }}
    .section {{ margin-bottom: 32px; }}
  </style>
</head>
<body>
  <h1>Relatório Visual de Dados</h1>

  <div class="section">
    <h2>1) Visão geral</h2>
    {dataframe_overview(df)}
  </div>

  <div class="section">
    <h2>2) Prévia dos dados</h2>
    {head_rows}
  </div>

  <div class="section">
    <h2>3) Estatísticas descritivas</h2>
    {describe_html}
  </div>

  <div class="section">
    <h2>4) Gráficos automáticos</h2>
    {charts if charts else '<p>Sem colunas adequadas para geração automática de gráficos.</p>'}
  </div>
</body>
</html>
"""


def main() -> None:
    args = parse_args()

    input_path = Path(args.input)
    output_path = Path(args.output)

    if not input_path.exists():
        raise FileNotFoundError(f"Arquivo de entrada não encontrado: {input_path}")

    df = load_dataframe(input_path, args.separator, args.encoding)
    df = normalize_dataframe(df, args.date_columns)

    report = build_html_report(df)
    output_path.write_text(report, encoding="utf-8")
    print(f"Relatório gerado com sucesso: {output_path}")


if __name__ == "__main__":
    main()
