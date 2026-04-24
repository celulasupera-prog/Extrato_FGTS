# -*- coding: utf-8 -*-
"""Converte extrato analítico FGTS (TXT) em planilha Excel por trabalhador."""

import re
from datetime import date, datetime
from pathlib import Path

import pandas as pd
from dateutil.relativedelta import relativedelta
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

MESES_MAP = {
    "JANEIRO": "01",
    "FEVEREIRO": "02",
    "MARCO": "03",
    "MARÇO": "03",
    "ABRIL": "04",
    "MAIO": "05",
    "JUNHO": "06",
    "JULHO": "07",
    "AGOSTO": "08",
    "SETEMBRO": "09",
    "OUTUBRO": "10",
    "NOVEMBRO": "11",
    "DEZEMBRO": "12",
}

TABLE_COLUMNS = [
    "Competência",
    "Data do Depósito",
    "Situação FGTS",
    "Valor (R$)",
    "Prazo (20 do mês seguinte)",
]

DEP_RE = re.compile(
    r"(\d{2}/\d{2}/\d{4})\s+DEPOSITO(?:\s+EM\s+ATRASO)?\s+([A-ZÇÃÁÉÍÓÚÊÔÕÂÜ]+)\/(\d{4})\s+([\d\.,-]+)",
    re.UNICODE,
)


def _comp_to_date(comp: str) -> date:
    """Converte 'MM/AAAA' em date(AAAA, MM, 1)."""
    mm, yy = comp.split("/")
    return date(int(yy), int(mm), 1)


def _prazo_fgts(comp: str) -> date:
    """Prazo padrão: dia 20 do mês seguinte à competência."""
    base = _comp_to_date(comp)
    prox = base + relativedelta(months=1)
    return date(prox.year, prox.month, 20)


def _parse_header(block_text: str) -> dict:
    m_nome = re.search(r"NOME DO TRABALHADOR.*\n\s*([A-ZÀ-ÜÇ ]+?)\s+\d{3,}", block_text, re.M)
    m_emp = re.search(r"NOME DO EMPREGADOR.*\n\s*([^\n]+?)\s+\d{14}\s*$", block_text, re.M)
    m_insc = re.search(r"INSCRICAO EMPREGADOR\s*\n\s*[^\n]+?\s+(\d{14})\s*$", block_text, re.M)
    m_adm = re.search(r"DTA\.ADM\.\s+SITUACAO CTA\s*\n[^\n]*?(\d{2}/\d{2}/\d{4})", block_text, re.M)
    m_line = re.search(r"DATA DE OPCAO.*\n\s*([0-9/]{10})\s+([0-9/]{10})\s+([0-9/]{10})", block_text, re.M)

    return {
        "nome_trabalhador": m_nome.group(1).strip() if m_nome else "",
        "empregador": m_emp.group(1).strip() if m_emp else "",
        "inscricao_empregador": m_insc.group(1) if m_insc else "",
        "data_admissao": m_adm.group(1) if m_adm else "",
        "data_afastamento": m_line.group(3) if m_line else "",
    }


def _parse_date_br(value: str) -> date | None:
    if not value:
        return None
    try:
        return datetime.strptime(value, "%d/%m/%Y").date()
    except ValueError:
        return None


def _merge_headers(current: dict, incoming: dict) -> dict:
    merged = current.copy()

    for field in ("nome_trabalhador", "empregador", "inscricao_empregador"):
        if not merged.get(field) and incoming.get(field):
            merged[field] = incoming[field]

    curr_adm = _parse_date_br(merged.get("data_admissao", ""))
    inc_adm = _parse_date_br(incoming.get("data_admissao", ""))
    if curr_adm is None or (inc_adm is not None and inc_adm < curr_adm):
        merged["data_admissao"] = incoming.get("data_admissao", merged.get("data_admissao", ""))

    curr_af = _parse_date_br(merged.get("data_afastamento", ""))
    inc_af = _parse_date_br(incoming.get("data_afastamento", ""))
    if curr_af is None or (inc_af is not None and inc_af > curr_af):
        merged["data_afastamento"] = incoming.get("data_afastamento", merged.get("data_afastamento", ""))

    return merged


def _extract_deposits(block_text: str) -> pd.DataFrame:
    deposits = []
    for m in DEP_RE.finditer(block_text):
        data_dep = datetime.strptime(m.group(1), "%d/%m/%Y").date()
        mes_nome = m.group(2).strip()
        ano = int(m.group(3))
        valor = float(m.group(4).strip().replace(".", "").replace(",", "."))

        mes_num = MESES_MAP.get(mes_nome)
        if not mes_num:
            continue

        deposits.append(
            {
                "competencia": f"{mes_num}/{ano}",
                "data": data_dep,
                "valor": valor,
            }
        )

    return pd.DataFrame(deposits)


def _build_competence_table(header: dict, dep_df: pd.DataFrame) -> pd.DataFrame:
    if dep_df.empty:
        raise ValueError("Nenhum recolhimento identificado neste extrato.")

    agg = dep_df.groupby("competencia").agg(
        valor_total=("valor", "sum"),
        datas=("data", lambda s: sorted(set(s))),
    ).reset_index()

    if header.get("data_admissao"):
        adm_date = datetime.strptime(header["data_admissao"], "%d/%m/%Y").date()
        start_comp_date = date(adm_date.year, adm_date.month, 1)
    else:
        first_comp = min(agg["competencia"], key=_comp_to_date)
        start_comp_date = _comp_to_date(first_comp)

    last_comp = max(agg["competencia"], key=_comp_to_date)
    end_comp_date = _comp_to_date(last_comp)

    competencias = []
    cur = start_comp_date
    while cur <= end_comp_date:
        competencias.append(f"{cur.month:02d}/{cur.year}")
        cur += relativedelta(months=1)

    agg_map = {row["competencia"]: row for _, row in agg.iterrows()}

    rows = []
    for comp in competencias:
        prazo = _prazo_fgts(comp)
        if comp in agg_map:
            datas = agg_map[comp]["datas"]
            valor_total = float(agg_map[comp]["valor_total"])
            em_atraso = any(d > prazo for d in datas)
            situacao = "Em Atraso" if em_atraso else "No Prazo"
            data_dep_str = ", ".join(d.strftime("%d/%m/%Y") for d in datas)
            rows.append([comp, data_dep_str, situacao, valor_total, prazo.strftime("%d/%m/%Y")])
        else:
            rows.append([comp, "—", "Não recolhido", 0.0, prazo.strftime("%d/%m/%Y")])

    return pd.DataFrame(rows, columns=TABLE_COLUMNS)


def _split_workers(full_text: str) -> list[str]:
    marker = "FGTS - EXTRATO ANALITICO DO TRABALHADOR"
    parts = full_text.split(marker)
    if len(parts) <= 1:
        return []
    return [marker + p for p in parts[1:]]


def _read_txt_with_fallback(path: Path) -> str:
    raw = path.read_bytes()
    for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin1"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    raise ValueError("Não foi possível decodificar o arquivo TXT com os encodings suportados.")


def _safe_sheet_name(name: str, idx: int, used: set[str]) -> str:
    base = (name or f"Trabalhador_{idx}").strip()[:31]
    candidate = base
    suffix = 1
    while candidate in used:
        tail = f"_{suffix}"
        candidate = f"{base[:31-len(tail)]}{tail}"
        suffix += 1
    used.add(candidate)
    return candidate


def _apply_status_fill(ws, start_row: int, end_row: int) -> None:
    fill_ok = PatternFill("solid", fgColor="E2F0D9")
    fill_late = PatternFill("solid", fgColor="FCE4D6")
    fill_not_paid = PatternFill("solid", fgColor="F8CBAD")

    for row in range(start_row, end_row + 1):
        status = ws.cell(row=row, column=3).value
        if status == "No Prazo":
            ws.cell(row=row, column=3).fill = fill_ok
        elif status == "Em Atraso":
            ws.cell(row=row, column=3).fill = fill_late
        elif status == "Não recolhido":
            for col in range(1, 6):
                ws.cell(row=row, column=col).fill = fill_not_paid


def _add_summary_sheet(wb: Workbook, summaries: list[dict]) -> None:
    ws = wb.create_sheet(title="Resumo", index=0)
    ws["A1"] = "Resumo FGTS"
    ws["A1"].font = Font(bold=True, size=14)
    ws.merge_cells("A1:F1")

    headers = ["Trabalhador", "Empregador", "Recolhidas", "Em Atraso", "Não recolhidas", "Total (R$)"]
    for idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=3, column=idx, value=header)
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="D9E1F2")
        cell.alignment = Alignment(horizontal="center")

    for row_idx, item in enumerate(summaries, start=4):
        ws.cell(row=row_idx, column=1, value=item["trabalhador"])
        ws.cell(row=row_idx, column=2, value=item["empregador"])
        ws.cell(row=row_idx, column=3, value=item["recolhidas"])
        ws.cell(row=row_idx, column=4, value=item["em_atraso"])
        ws.cell(row=row_idx, column=5, value=item["nao_recolhidas"])
        total_cell = ws.cell(row=row_idx, column=6, value=item["total"])
        total_cell.number_format = "#,##0.00"

    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:F{3 + len(summaries)}"
    for col in range(1, 7):
        ws.column_dimensions[get_column_letter(col)].width = 20


def _add_lancamentos_sheet(wb: Workbook, launch_rows: list[dict]) -> None:
    ws = wb.create_sheet(title="Lançamentos")
    ws["A1"] = "Lançamentos capturados no TXT"
    ws["A1"].font = Font(bold=True, size=14)
    ws.merge_cells("A1:H1")

    headers = [
        "Trabalhador",
        "Empregador",
        "Inscrição Empregador",
        "Competência",
        "Data do Depósito",
        "Valor (R$)",
        "Prazo",
        "Situação FGTS",
    ]
    for idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=3, column=idx, value=header)
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="D9E1F2")
        cell.alignment = Alignment(horizontal="center")

    for row_idx, item in enumerate(launch_rows, start=4):
        ws.cell(row=row_idx, column=1, value=item["trabalhador"])
        ws.cell(row=row_idx, column=2, value=item["empregador"])
        ws.cell(row=row_idx, column=3, value=item["inscricao_empregador"])
        ws.cell(row=row_idx, column=4, value=item["competencia"])
        ws.cell(row=row_idx, column=5, value=item["data_deposito"])
        ws.cell(row=row_idx, column=6, value=item["valor"]).number_format = "#,##0.00"
        ws.cell(row=row_idx, column=7, value=item["prazo"])
        ws.cell(row=row_idx, column=8, value=item["situacao"])

    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:H{3 + len(launch_rows)}"
    for col in range(1, 9):
        ws.column_dimensions[get_column_letter(col)].width = 20


def extrato_fgts_txt_para_excel(txt_path: str, xlsx_path: str = "Extrato_FGTS_Analitico_Processado.xlsx") -> str:
    content = _read_txt_with_fallback(Path(txt_path))

    marker = "FGTS - EXTRATO ANALITICO DO TRABALHADOR"
    if marker not in content:
        raise ValueError("O arquivo enviado não segue o formato do extrato analítico do FGTS.")

    blocks = _split_workers(content)
    if not blocks:
        raise ValueError("O texto não segue o formato de extrato analítico padrão da Caixa.")

    grouped: dict[tuple[str, str], dict] = {}
    for idx, block in enumerate(blocks, start=1):
        header = _parse_header(block)
        dep_df = _extract_deposits(block)

        worker_name = header.get("nome_trabalhador", "").strip()
        employer_id = header.get("inscricao_empregador", "").strip()
        group_key = (worker_name or f"TRABALHADOR_{idx}", employer_id or "SEM_INSCRICAO")

        if group_key not in grouped:
            grouped[group_key] = {
                "idx": idx,
                "header": header,
                "deposits": [],
            }
        else:
            grouped[group_key]["header"] = _merge_headers(grouped[group_key]["header"], header)

        if not dep_df.empty:
            grouped[group_key]["deposits"].append(dep_df)

    wb = Workbook()
    wb.remove(wb.active)
    used_sheet_names: set[str] = set()
    summaries: list[dict] = []
    launch_rows: list[dict] = []

    for _, group in sorted(grouped.items(), key=lambda item: item[1]["idx"]):
        header = group["header"]
        if group["deposits"]:
            merged_dep_df = pd.concat(group["deposits"], ignore_index=True)
            df = _build_competence_table(header, merged_dep_df)
        else:
            df = pd.DataFrame(columns=TABLE_COLUMNS)

        sheet_name = _safe_sheet_name(header.get("nome_trabalhador", ""), group["idx"], used_sheet_names)
        ws = wb.create_sheet(title=sheet_name)

        bold = Font(bold=True)
        ws["A1"], ws["B1"] = "Nome do Trabalhador:", header.get("nome_trabalhador", "")
        ws["A2"] = "Empregador:"
        insc = header.get("inscricao_empregador", "")
        emp = header.get("empregador", "")
        ws["B2"] = f"{emp} ({insc})" if insc else emp
        ws["A3"], ws["B3"] = "Data de Admissão:", header.get("data_admissao", "")
        ws["A4"], ws["B4"] = "Data de Afastamento:", header.get("data_afastamento", "")
        for c in ("A1", "A2", "A3", "A4"):
            ws[c].font = bold

        recolhidas = int((df["Situação FGTS"] == "No Prazo").sum() + (df["Situação FGTS"] == "Em Atraso").sum())
        em_atraso = int((df["Situação FGTS"] == "Em Atraso").sum())
        nao_recolhidas = int((df["Situação FGTS"] == "Não recolhido").sum())
        total_valor = float(df["Valor (R$)"].sum()) if not df.empty else 0.0

        metric_labels = ["Competências", "Recolhidas", "Em Atraso", "Não recolhidas", "Total (R$)"]
        metric_values = [len(df), recolhidas, em_atraso, nao_recolhidas, total_valor]
        for idx, label in enumerate(metric_labels, start=1):
            cell = ws.cell(row=6, column=idx, value=label)
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor="D9E1F2")
            ws.cell(row=7, column=idx, value=metric_values[idx - 1])
        ws.cell(row=7, column=5).number_format = "#,##0.00"

        ws.cell(row=8, column=1, value="Detalhamento por competência").font = Font(bold=True)

        start_row = ws.max_row + 1
        for row in dataframe_to_rows(df, index=False, header=True):
            ws.append(row)

        header_row = start_row
        fill = PatternFill("solid", fgColor="D9E1F2")
        for col in range(1, df.shape[1] + 1):
            cell = ws.cell(row=header_row, column=col)
            cell.font = Font(bold=True)
            cell.fill = fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        ws.freeze_panes = ws[f"A{header_row + 1}"]
        ws.auto_filter.ref = f"A{header_row}:{get_column_letter(df.shape[1])}{header_row + len(df)}"

        currency_format = "#,##0.00"
        for r in range(header_row + 1, header_row + 1 + len(df)):
            ws.cell(row=r, column=4).number_format = currency_format
            ws.cell(row=r, column=1).alignment = Alignment(horizontal="center")
            ws.cell(row=r, column=3).alignment = Alignment(horizontal="center")
            ws.cell(row=r, column=4).alignment = Alignment(horizontal="right")
            ws.cell(row=r, column=5).alignment = Alignment(horizontal="center")

        for col in range(1, df.shape[1] + 1):
            max_len = 0
            for r in range(1, header_row + 1 + len(df)):
                val = ws.cell(row=r, column=col).value
                if val is not None:
                    max_len = max(max_len, len(str(val)))
            ws.column_dimensions[get_column_letter(col)].width = min(max(max_len + 2, 12), 45)

        ws.column_dimensions["A"].width = 14
        _apply_status_fill(ws, header_row + 1, header_row + len(df))

        summaries.append(
            {
                "trabalhador": header.get("nome_trabalhador", ""),
                "empregador": emp,
                "recolhidas": recolhidas,
                "em_atraso": em_atraso,
                "nao_recolhidas": nao_recolhidas,
                "total": total_valor,
            }
        )

        if group["deposits"]:
            for _, dep in merged_dep_df.iterrows():
                competencia = dep["competencia"]
                prazo = _prazo_fgts(competencia).strftime("%d/%m/%Y")
                data_dep = dep["data"].strftime("%d/%m/%Y")
                situacao = "Em Atraso" if dep["data"] > _prazo_fgts(competencia) else "No Prazo"
                launch_rows.append(
                    {
                        "trabalhador": header.get("nome_trabalhador", ""),
                        "empregador": emp,
                        "inscricao_empregador": insc,
                        "competencia": competencia,
                        "data_deposito": data_dep,
                        "valor": float(dep["valor"]),
                        "prazo": prazo,
                        "situacao": situacao,
                    }
                )

    _add_summary_sheet(wb, summaries)
    _add_lancamentos_sheet(wb, launch_rows)
    wb.save(xlsx_path)
    return xlsx_path


if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("Uso: python fgts_extrato_to_excel.py <extrato.txt> [saida.xlsx]")
        raise SystemExit(1)

    txt = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) >= 3 else "Extrato_FGTS_Analitico_Processado.xlsx"
    generated = extrato_fgts_txt_para_excel(txt, out)
    print(f"Planilha gerada em: {generated}")
