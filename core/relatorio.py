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
MOEDA = 'R$ #,##0.00;[Red]-R$ #,##0.00'
CAB_FILL = PatternFill("solid", fgColor="1F4E78")


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


def _custos(wb: Workbook, resultado) -> None:
    """Aba Custos: totais e tabelas por licença, Licença 2026, tenant, empreendimento, setor e cargo."""
    from core.analise import CUSTO_2026, CUSTO_ATUAL, custos, projetar

    ws = wb.create_sheet("Custos")
    cx = resultado.caixas
    atual = round(float(cx[CUSTO_ATUAL].sum()), 2)
    futuro = round(float(projetar(cx)[CUSTO_2026].sum()), 2)
    sem_perfil = int(cx[CUSTO_2026].isna().sum())
    saldo = resultado.saldo_licencas
    livres = round(float(saldo["Custo das livres (mês)"].fillna(0).sum()), 2) if saldo is not None and len(saldo) else 0.0

    ws.append(["Custos mensais das licenças pagas"])
    ws["A1"].font = Font(name="Arial", size=14, bold=True)
    ws.append(["Contas do relatório (licença paga e login ativo). Valores de config/licencas.csv."])
    ws.append([])
    linhas_total = [
        ("Contas no relatório", len(cx), None),
        ("Custo atual (mês)", atual, MOEDA), ("Custo atual (ano)", round(atual * 12, 2), MOEDA),
        ("Custo 2026 (mês)", futuro, MOEDA), ("Custo 2026 (ano)", round(futuro * 12, 2), MOEDA),
        ("Diferença 2026 - atual (mês)", round(futuro - atual, 2), MOEDA),
        ("Licenças compradas e sem uso (mês)", livres, MOEDA),
        ("Licenças pagas em contas bloqueadas (mês)", resultado.custo_bloqueadas, MOEDA),
    ]
    if sem_perfil:
        linhas_total.append(("Contas sem Licença 2026 (na projeção, mantêm o custo atual)", sem_perfil, None))
    if resultado.licencas_sem_preco:
        linhas_total.append(("Licenças sem preço no licencas.csv (contam 0)",
                             ", ".join(sorted(resultado.licencas_sem_preco)), None))
    for rotulo, valor, fmt in linhas_total:
        ws.append([rotulo, valor])
        ws.cell(ws.max_row, 1).font = FONTE
        cel = ws.cell(ws.max_row, 2)
        cel.font = NEGRITO
        if fmt:
            cel.number_format = fmt

    for n, (titulo, df) in enumerate(custos(cx), start=1):
        ws.append([])
        ws.append([titulo])
        ws.cell(ws.max_row, 1).font = Font(name="Arial", size=12, bold=True)
        inicio = ws.max_row + 1
        ws.append(list(df.columns))
        for i in range(1, len(df.columns) + 1):
            c = ws.cell(inicio, i)
            c.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
            c.fill = CAB_FILL
        for linha in df.itertuples(index=False):
            ws.append(list(linha))
            for i in range(1, len(df.columns) + 1):
                c = ws.cell(ws.max_row, i)
                c.font = FONTE
                if i >= 3:
                    c.number_format = MOEDA
        if len(df):
            ref = f"A{inicio}:{get_column_letter(len(df.columns))}{ws.max_row}"
            tabela = Table(displayName=f"tbCusto{n}", ref=ref)
            tabela.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
            ws.add_table(tabela)
    for col, w in zip("ABCDE", (62, 10, 18, 18, 18)):
        ws.column_dimensions[col].width = w


def gerar(resultado, pasta_saida: Path) -> Path:
    pasta_saida.mkdir(exist_ok=True)
    agora = datetime.now()
    wb = Workbook()
    resumo = wb.active
    resumo.title = "Resumo"

    ws = _aba(wb, "Caixas", resultado.caixas, "tbCaixas", {
        "Display name": 38, "User principal name": 44, "Dominio": 28, "Cargo": 34, "Setor": 22,
        "Empreendimento": 24, "Licenses": 44, "Licença 2026": 44, "Tenant Novo": 36, "Bloqueado": 11,
        "Custo atual (mês)": 16, "Custo 2026 (mês)": 16})
    # Destaca o que não casou com as tabelas de referência
    cols = list(resultado.caixas.columns)
    for r in range(2, ws.max_row + 1):
        for nome in ("Licença 2026", "Tenant Novo"):
            cel = ws.cell(r, cols.index(nome) + 1)
            if cel.value == "não encontrado":
                cel.fill = ALERTA

    for nome in ("Custo atual (mês)", "Custo 2026 (mês)"):
        col = cols.index(nome) + 1
        for r in range(2, ws.max_row + 1):
            ws.cell(r, col).number_format = MOEDA

    _custos(wb, resultado)

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
        resumo.append(["Licenças pagas", "Compradas", "Em uso", "Livres", "Preço (mês)", "Custo das livres (mês)"])
        for cel in resumo[resumo.max_row]:
            cel.font = NEGRITO
        for linha in saldo.itertuples(index=False):
            resumo.append(list(linha))
            r = resumo.max_row
            resumo.cell(r, 5).number_format = MOEDA
            resumo.cell(r, 6).number_format = MOEDA
            for c in range(1, 7):
                resumo.cell(r, c).font = FONTE
            if linha.Livres < 0:
                resumo.cell(r, 4).fill = ALERTA
    resumo.column_dimensions["C"].width = 12
    resumo.column_dimensions["D"].width = 12
    resumo.column_dimensions["E"].width = 14
    resumo.column_dimensions["F"].width = 22
    resumo.column_dimensions["A"].width = 42
    resumo.column_dimensions["B"].width = 14
    for row in resumo.iter_rows(min_row=4, min_col=2, max_col=6):
        for cel in row:
            cel.alignment = Alignment(horizontal="right")

    caminho = pasta_saida / f"Relatorio_Caixas_{agora:%Y%m%d_%H%M}.xlsx"
    wb.save(caminho)
    return caminho