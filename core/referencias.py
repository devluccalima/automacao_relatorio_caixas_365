"""Tabelas de referência lidas de config/.

- cargos_licencas.csv      Cargo;Perfil
- dominios_tenants.csv     Dominio;Tenant novo
- licencas.csv             SkuPartNumber;Nome;Paga   (a ordem das linhas é a ordem de exibição)
- setores.csv              Setor            (opcional; vazio = não valida)
- empreendimentos.csv      Empreendimento   (opcional; vazio = não valida)
"""
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd


def chave(texto) -> str:
    """Comparação igual à do PROCX: ignora maiúsculas e espaços nas pontas e repetidos."""
    return " ".join(str(texto or "").split()).casefold()


@dataclass
class Referencias:
    cargos: dict = field(default_factory=dict)          # chave(cargo) -> perfil
    nomes_cargos: dict = field(default_factory=dict)    # chave(cargo) -> grafia oficial
    tenants: dict = field(default_factory=dict)         # dominio -> tenant novo
    licencas: dict = field(default_factory=dict)        # SKU em maiúsculas -> (nome, paga, ordem)
    setores: dict = field(default_factory=dict)         # chave(setor) -> grafia oficial (vazio = não valida)
    empreendimentos: dict = field(default_factory=dict) # chave(empreendimento) -> grafia oficial


def _ler_csv(caminho: Path) -> pd.DataFrame:
    if not caminho.exists():
        return pd.DataFrame()
    for codificacao in ("utf-8-sig", "cp1252"):
        try:
            df = pd.read_csv(caminho, sep=";", dtype=str, encoding=codificacao).fillna("")
            return df.apply(lambda col: col.str.strip())
        except UnicodeDecodeError:
            continue
    raise ValueError(f"Não consegui ler {caminho} (codificação desconhecida)")


def carregar(pasta: Path) -> Referencias:
    refs = Referencias()
    df = _ler_csv(pasta / "cargos_licencas.csv")
    if df.empty:
        raise ValueError("config/cargos_licencas.csv está vazio ou não existe")
    refs.cargos = {chave(c): p for c, p in zip(df["Cargo"], df["Perfil"]) if c}
    refs.nomes_cargos = {chave(c): c for c in df["Cargo"] if c}

    df = _ler_csv(pasta / "dominios_tenants.csv")
    refs.tenants = {d.lower(): t for d, t in zip(df["Dominio"], df["Tenant novo"]) if d}

    df = _ler_csv(pasta / "licencas.csv")
    refs.licencas = {
        s.upper(): (n, p.lower() in ("sim", "s", "true", "1"), i)
        for i, (s, n, p) in enumerate(zip(df["SkuPartNumber"], df["Nome"], df["Paga"]))
    }

    df = _ler_csv(pasta / "setores.csv")
    refs.setores = {chave(v): v for v in df.get("Setor", []) if v}
    df = _ler_csv(pasta / "empreendimentos.csv")
    refs.empreendimentos = {chave(v): v for v in df.get("Empreendimento", []) if v}
    return refs