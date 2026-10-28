"""Ciclo do relatório de caixas M365.

1. Push: aplica as linhas do Alteracoes.xlsx (simulação enquanto DRY_RUN=true)
2. Pull: lê as contas do Entra ID
3. Análise: Licença 2026, tenant novo, pendências, mudanças
4. Relatório Excel local e no SharePoint

Uso:
  python main.py                       ciclo completo
  python main.py --somente-alteracoes  só o push: simula, mostra e, se DRY_RUN=false, pede confirmação
"""
import argparse
import os
import sys
from pathlib import Path

from dotenv import load_dotenv

RAIZ = Path(__file__).resolve().parent
NOME_RELATORIO_SHAREPOINT = "Relatorio_Caixas.xlsx"
load_dotenv(RAIZ / ".env", override=True)
sys.stdout.reconfigure(encoding="utf-8")


def dry_run() -> bool:
    return os.getenv("DRY_RUN", "true").strip().lower() not in ("false", "0", "nao", "não")


def imprimir_resumo_alteracoes(r: dict) -> None:
    print(f"  pendentes: {r['pendentes']} | aplicadas: {r['aplicadas']} | sem mudança: {r['sem mudança']} | "
          f"simuladas: {r['simuladas']} | erros: {r['erros']} | inválidas: {r['inválidas']} | "
          f"aguardando: {r['aguardando']}")


def somente_alteracoes(refs) -> None:
    from core import alteracoes
    print("Simulando as alterações...")
    r = alteracoes.processar(refs, simular=True)
    imprimir_resumo_alteracoes(r)
    for email, res, det in r["detalhes"][:40]:
        print(f"    {res:<12} {email}: {det}")
    if len(r["detalhes"]) > 40:
        print(f"    ... e mais {len(r['detalhes']) - 40} (veja a coluna Status do Alteracoes.xlsx)")
    if dry_run():
        print("DRY_RUN=true: nada foi gravado. Para aplicar, mude DRY_RUN=false no .env e rode de novo.")
        return
    if not r["simuladas"]:
        print("Nenhuma conta para alterar.")
        return
    resposta = input(f"\nDigite APLICAR para gravar {r['simuladas']} contas no Exchange: ").strip()
    if resposta != "APLICAR":
        print("Cancelado; nada foi gravado.")
        return
    print("Gravando...")
    imprimir_resumo_alteracoes(alteracoes.processar(refs, simular=False))


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--somente-alteracoes", action="store_true")
    args = ap.parse_args()

    from core import alteracoes, analise, leitura_tenant, referencias, relatorio, sharepoint

    print("Lendo tabelas de referência...")
    refs = referencias.carregar(RAIZ / "config")
    print(f"  {len(refs.cargos)} cargos, {len(refs.setores)} setores, {len(refs.empreendimentos)} empreendimentos, "
          f"{len(refs.tenants)} domínios, {len(refs.licencas)} licenças cadastradas")

    if args.somente_alteracoes:
        somente_alteracoes(refs)
        return

    resumo_alt = None
    print(f"Processando alterações ({'simulação' if dry_run() else 'gravando'})...")
    try:
        resumo_alt = alteracoes.processar(refs, simular=dry_run())
        imprimir_resumo_alteracoes(resumo_alt)
    except Exception as e:  # o relatório sai mesmo se o push falhar
        print(f"  ERRO nas alterações (o relatório continua): {e}")

    print("Lendo contas do Entra ID...")
    skus = leitura_tenant.ler_skus()
    contas = leitura_tenant.ler_contas(skus)
    print(f"  {len(contas)} contas membro")

    resultado = analise.analisar(contas, refs)
    resultado.saldo_licencas = analise.saldo_licencas(skus, refs)
    if resumo_alt:
        for k in ("aplicadas", "simuladas", "erros", "inválidas", "aguardando"):
            resultado.totais[f"alterações: {k}"] = resumo_alt[k]
    analise.comparar_com_retrato(resultado, RAIZ / "retratos")
    caminho = relatorio.gerar(resultado, RAIZ / "relatorios")

    for chave, valor in resultado.totais.items():
        print(f"  {chave}: {valor}")
    print(f"Relatório gerado: {caminho}")

    if sharepoint.configurado():
        print("Enviando ao SharePoint...")
        envio = sharepoint.enviar_arquivo(caminho, NOME_RELATORIO_SHAREPOINT)
        if envio["copia"]:
            print(f"  ATENÇÃO: {NOME_RELATORIO_SHAREPOINT} estava aberto; gravei como {envio['nome']}")
        print(f"  {envio['url']}")
    else:
        print("SharePoint não configurado (SHAREPOINT_SITE_ID vazio): relatório só local.")


if __name__ == "__main__":
    main()