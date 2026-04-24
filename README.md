# Extrato_FGTS

Conversão de extrato analítico do FGTS (TXT) para Excel, com classificação mensal de recolhimento:

- **No Prazo**
- **Em Atraso**
- **Não recolhido**

## Arquivo principal

- `fgts_extrato_to_excel.py`

## Requisitos

- Python 3.10+
- Dependências:
  - pandas
  - openpyxl
  - python-dateutil

Instalação:

```bash
pip install pandas openpyxl python-dateutil
```

## Uso

```bash
python fgts_extrato_to_excel.py <extrato.txt> [saida.xlsx]
```

Exemplo:

```bash
python fgts_extrato_to_excel.py extrato_analitico.txt Extrato_FGTS_Analitico_Processado.xlsx
```

## O que o script gera

- Uma aba por trabalhador encontrado no TXT.
- Cabeçalho com nome do trabalhador, empregador, inscrição, admissão e afastamento.
- Tabela mensal por competência com:
  - Competência
  - Data(s) do depósito
  - Situação FGTS
  - Valor (R$)
  - Prazo (20 do mês seguinte)

## Observações

- O parser considera o layout padrão de extrato analítico da Caixa.
- Se uma competência não tiver recolhimento identificado, a linha é marcada como **Não recolhido** com valor `0,00`.
