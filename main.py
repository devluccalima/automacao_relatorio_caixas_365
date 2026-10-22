"""Ciclo do relatório de caixas M365.

Fase 1 (atual): pull do tenant -> análise -> relatório Excel local.
Fases seguintes: SharePoint e e-mail (2), push pela aba Alterações (3).
"""
from pathlib import Path

from dotenv import load_dotenv

RAIZ = Path(__file__).resolve().parent
load_dotenv(RAIZ / ".env", override=True)


def main() -> None:
    from core import analise, leitura_tenant, referencias, relatorio

    refs = referencias.carregar(RAIZ / "config")
    contas = leitura_tenant.ler_contas()
    resultado = analise.analisar(contas, refs)
    caminho = relatorio.gerar(resultado, RAIZ / "relatorios")
    print(f"Relatório gerado: {caminho}")


if __name__ == "__main__":
    main()
