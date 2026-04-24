# Extrato_FGTS

Script para tratamento de dados e geração de relatório visual em HTML.

## Requisitos

- Python 3.10+
- Dependências:
  - pandas
  - plotly
  - openpyxl (para XLSX/XLS)

Instalação rápida:

```bash
pip install pandas plotly openpyxl
```

## Uso

```bash
python fgts_visual_report.py --input dados.csv --output relatorio.html
```

### Parâmetros úteis

- `--separator ";"`: separador de CSV.
- `--encoding "latin-1"`: encoding para CSV.
- `--date-columns data_competencia data_pagamento`: força colunas como data.

## O que o relatório inclui

1. Visão geral (linhas, colunas, faltantes, duplicadas)
2. Prévia das primeiras linhas
3. Estatísticas descritivas
4. Gráficos automáticos:
   - Histograma para colunas numéricas
   - Barras para principais categorias
   - Série temporal (quando houver data + valor numérico)

## Exemplo

```bash
python fgts_visual_report.py \
  --input extrato_fgts.csv \
  --output relatorio_fgts.html \
  --separator ";" \
  --encoding "latin-1" \
  --date-columns data_movimento
```
