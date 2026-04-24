# Extrato_FGTS

Conversão de extrato analítico do FGTS (TXT) para Excel, com classificação mensal de recolhimento:

- **No Prazo**
- **Em Atraso**
- **Não recolhido**

## Arquivos principais

- `fgts_extrato_to_excel.py` (motor de processamento)
- `streamlit_app.py` (interface web)

## Requisitos

- Python 3.10+
- Dependências em `requirements.txt`:
  - pandas
  - openpyxl
  - python-dateutil
  - streamlit

Instalação:

```bash
pip install -r requirements.txt
```

## Uso local (CLI)

```bash
python fgts_extrato_to_excel.py <extrato.txt> [saida.xlsx]
```

Exemplo:

```bash
python fgts_extrato_to_excel.py extrato_analitico.txt Extrato_FGTS_Analitico_Processado.xlsx
```

## Uso local (Streamlit)

```bash
streamlit run streamlit_app.py
```

Depois abra o endereço local mostrado no terminal (normalmente `http://localhost:8501`).

## Deploy no Streamlit Community Cloud

1. Suba este projeto para um repositório no GitHub.
2. Acesse Streamlit Community Cloud e clique em **New app**.
3. Selecione o repositório e configure:
   - **Main file path**: `streamlit_app.py`
   - **Python version**: 3.10+
4. Faça deploy.

## O que o script gera

- Uma aba por trabalhador encontrado no TXT.
- Blocos repetidos do mesmo trabalhador/vínculo (mesmo nome + inscrição do empregador) são consolidados na mesma aba.
- Cabeçalho com nome do trabalhador, empregador, inscrição, admissão e afastamento.
- Tabela mensal por competência com:
  - Competência
  - Data(s) do depósito
  - Situação FGTS
  - Valor (R$)
  - Prazo (20 do mês seguinte)
- Aba **Resumo** com consolidação por trabalhador (recolhidas, em atraso, não recolhidas e total).
- Aba **Lançamentos** para auditoria linha a linha dos depósitos capturados no TXT.
- Realce visual no Excel:
  - verde para **No Prazo**
  - laranja para **Em Atraso**
  - vermelho claro para **Não recolhido**
- Bloco de métricas no topo de cada aba de trabalhador (competências, recolhidas, atraso, total).

## Observações

- O parser considera o layout padrão de extrato analítico da Caixa.
- Se uma competência não tiver recolhimento identificado, a linha é marcada como **Não recolhido** com valor `0,00`.
- A leitura do TXT tenta automaticamente `utf-8-sig`, `utf-8`, `cp1252` e `latin1`.
