"""Monta o Excel no formato da planilha de caixas: abas Caixas, Pendências, Mudanças e Resumo."""
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

FONTE = Font(name="Arial", size=10)
NEGRITO = Font(name="Arial", size=10, bold=True)
ALERTA = PatternFill("solid", fgColor="FCE4D6")


def _aba(wb: Workbook, titulo: str, df: pd.DataFrame, nome_tabela: str, larguras: dict) -> None:
    ws = wb.create_sheet(titulo)
    ws.append(list(df.columns))
    for linha in df.itertuples(index=False):
        ws.append(list(linha))
    for row in ws.iter_rows():
        for cel in row:
            cel.font = FONTE
    if len(df):
        ref = f"A1:{get_column_letter(len(df.columns))}{len(df) + 1}"
        tabela = Table(displayName=nome_tabela, ref=ref)
        tabela.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        ws.add_table(tabela)
    for i, col in enumerate(df.columns, start=1):
        ws.column_dimensions[get_column_letter(i)].width = larguras.get(col, 22)
        ws.cell(1, i).font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    ws.freeze_panes = "A2"
    return ws


def gerar(resultado, pasta_saida: Path) -> Path:
    pasta_saida.mkdir(exist_ok=True)
    agora = datetime.now()
    wb = Workbook()
    resumo = wb.active
    resumo.title = "Resumo"

    ws = _aba(wb, "Caixas", resultado.caixas, "tbCaixas", {
        "Display name": 38, "User principal name": 44, "Dominio": 28, "Cargo": 34, "Setor": 22,
        "Empreendimento": 24, "Licenses": 44, "Licença 2026": 44, "Tenant Novo": 36, "Bloqueado": 11})
    # Destaca o que não casou com as tabelas de referência
    cols = list(resultado.caixas.columns)
    for r in range(2, ws.max_row + 1):
        for nome in ("Licença 2026", "Tenant Novo"):
            cel = ws.cell(r, cols.index(nome) + 1)
            if cel.value == "não encontrado":
                cel.fill = ALERTA

    _aba(wb, "Pendências", resultado.pendencias, "tbPendencias",
         {"User principal name": 44, "Display name": 38, "Pendência": 36, "Detalhe": 70})
    _aba(wb, "Mudanças", resultado.mudancas, "tbMudancas",
         {"User principal name": 44, "Display name": 38, "Mudança": 20, "Campo": 18, "Antes": 40, "Depois": 40})

    resumo.append(["Relatório de caixas M365"])
    resumo["A1"].font = Font(name="Arial", size=14, bold=True)
    resumo.append([f"Gerado em {agora:%d/%m/%Y %H:%M}"])
    resumo.append([])
    for chave, valor in resultado.totais.items():
        resumo.append([chave.capitalize(), valor])
    resumo.append([])
    resumo.append(["Pendências por tipo"])
    resumo.cell(resumo.max_row, 1).font = NEGRITO
    for tipo, qtd in resultado.pendencias["Pendência"].value_counts().items():
        resumo.append([tipo, int(qtd)])
    for row in resumo.iter_rows(min_row=2):
        for cel in row:
            if not cel.font.bold:
                cel.font = FONTE
    saldo = resultado.saldo_licencas
    if saldo is not None and len(saldo):
        resumo.append([])
        resumo.append(["Licenças pagas", "Compradas", "Em uso", "Livres"])
        for cel in resumo[resumo.max_row]:
            cel.font = NEGRITO
        for linha in saldo.itertuples(index=False):
            resumo.append(list(linha))
            if linha.Livres < 0:
                resumo.cell(resumo.max_row, 4).fill = ALERTA
    resumo.column_dimensions["C"].width = 12
    resumo.column_dimensions["D"].width = 12
    resumo.column_dimensions["A"].width = 42
    resumo.column_dimensions["B"].width = 14
    for row in resumo.iter_rows(min_row=4, min_col=2, max_col=4):
        for cel in row:
            cel.alignment = Alignment(horizontal="right")

    caminho = pasta_saida / f"Relatorio_Caixas_{agora:%Y%m%d_%H%M}.xlsx"
    wb.save(caminho)
    return caminho
