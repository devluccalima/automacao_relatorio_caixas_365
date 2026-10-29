"""Licença 2026, tenant novo, pendências e mudanças desde o último retrato."""
import os
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd

from core.referencias import Referencias, chave

COLUNAS = ["Display name", "User principal name", "Dominio", "Cargo", "Setor", "Empreendimento",
           "Licenses", "Licença 2026", "Tenant Novo", "Bloqueado", "Custo atual (mês)", "Custo 2026 (mês)"]
CUSTO_ATUAL, CUSTO_2026 = "Custo atual (mês)", "Custo 2026 (mês)"
CAMPOS_MUDANCA = ["Display name", "Cargo", "Setor", "Empreendimento", "Licenses", "Bloqueado"]
SEM_PERFIL = "não encontrado"
PERFIL_SEM_LICENCA = "desinstalar"


@dataclass
class Resultado:
    caixas: pd.DataFrame
    pendencias: pd.DataFrame
    mudancas: pd.DataFrame
    totais: dict
    saldo_licencas: pd.DataFrame = None
    custo_bloqueadas: float = 0.0          # licenças pagas em contas bloqueadas (fora do relatório)
    licencas_sem_preco: set = None


def _flag(nome: str, padrao: bool) -> bool:
    return os.getenv(nome, str(padrao)).strip().lower() in ("true", "1", "sim", "s")


def _licencas(skus: list[str], refs: Referencias) -> tuple[list[str], list[str], list[str]]:
    """Devolve (nomes pagos em ordem fixa, nomes gratuitos, skus sem cadastro em licencas.csv)."""
    pagas, gratis, sem_cadastro = [], [], []
    for sku in skus:
        if sku.upper() in refs.licencas:
            nome, paga, ordem, _preco = refs.licencas[sku.upper()]
            (pagas if paga else gratis).append((ordem, nome))
        else:
            sem_cadastro.append(sku)
            pagas.append((999, sku))  # desconhecida conta como paga, para não sumir do relatório
    return [n for _, n in sorted(pagas)], [n for _, n in sorted(gratis)], sem_cadastro


def custo(nomes: list[str], refs: Referencias, sem_preco: set) -> float:
    """Soma o preço mensal das licenças (pelo nome). Licença sem preço conta 0 e é anotada."""
    total = 0.0
    for n in nomes:
        n = n.strip()
        if not n:
            continue
        if chave(n) in refs.precos:
            total += refs.precos[chave(n)]
        else:
            sem_preco.add(n)
    return round(total, 2)


def custo_perfil(perfil: str, refs: Referencias, sem_preco: set) -> float | None:
    if perfil == SEM_PERFIL:
        return None
    if chave(perfil) == PERFIL_SEM_LICENCA:
        return 0.0
    return custo(perfil.split("+"), refs, sem_preco)


def _conjunto(texto: str) -> set[str]:
    return {chave(p) for p in str(texto).split("+") if p.strip()}


def saldo_licencas(skus: list[dict], refs: Referencias) -> pd.DataFrame:
    """Licenças pagas do tenant: compradas, em uso e livres (na ordem de licencas.csv)."""
    linhas = []
    for s in skus:
        cad = refs.licencas.get(s["sku"].upper())
        if cad and not cad[1]:
            continue
        preco = cad[3] if cad else None
        livres = s["total"] - s["em_uso"]
        linhas.append({"Licença": cad[0] if cad else s["sku"], "Compradas": s["total"],
                       "Em uso": s["em_uso"], "Livres": livres, "Preço (mês)": preco,
                       "Custo das livres (mês)": round(max(livres, 0) * preco, 2) if preco is not None else None,
                       "_ordem": cad[2] if cad else 999})
    df = pd.DataFrame(linhas, columns=["Licença", "Compradas", "Em uso", "Livres", "Preço (mês)",
                                       "Custo das livres (mês)", "_ordem"])
    return df.sort_values("_ordem").drop(columns="_ordem").reset_index(drop=True)


def analisar(contas: list[dict], refs: Referencias) -> Resultado:
    incluir_sem_paga = _flag("INCLUIR_SEM_LICENCA_PAGA", False)
    incluir_bloqueadas = _flag("INCLUIR_BLOQUEADAS", False)

    linhas, pend = [], []
    sem_preco: set = set()
    custo_bloqueadas = 0.0

    def pendencia(c, tipo, detalhe):
        pend.append({"User principal name": c["upn"], "Display name": c["nome"],
                     "Pendência": tipo, "Detalhe": detalhe})

    fora = {"sem licença paga": 0, "bloqueadas": 0}
    for c in contas:
        pagas, _gratis, sem_cadastro = _licencas(c["skus"], refs)

        # Verificada em todas as contas, mesmo as que ficam fora do relatório
        if c["bloqueada"] and pagas:
            pendencia(c, "Bloqueada com licença paga", " + ".join(pagas))
            custo_bloqueadas += custo(pagas, refs, sem_preco)

        if not pagas and not incluir_sem_paga:
            fora["sem licença paga"] += 1
            continue
        if c["bloqueada"] and not incluir_bloqueadas:
            fora["bloqueadas"] += 1
            continue

        dominio = c["upn"].split("@")[-1].lower()
        perfil = refs.cargos.get(chave(c["cargo"]), SEM_PERFIL) if c["cargo"].strip() else SEM_PERFIL
        tenant = refs.tenants.get(dominio, SEM_PERFIL)
        licenses = "+".join(pagas)
        linhas.append({
            "Display name": c["nome"], "User principal name": c["upn"], "Dominio": dominio,
            "Cargo": c["cargo"].strip(), "Setor": c["setor"].strip(),
            "Empreendimento": c["empreendimento"].strip(), "Licenses": licenses,
            "Licença 2026": perfil, "Tenant Novo": tenant,
            "Bloqueado": "Sim" if c["bloqueada"] else "Não",
            CUSTO_ATUAL: custo(pagas, refs, sem_preco),
            CUSTO_2026: custo_perfil(perfil, refs, sem_preco),
        })

        # Pendências de cadastro
        for campo, valor in (("Cargo", c["cargo"]), ("Setor", c["setor"]), ("Empreendimento", c["empreendimento"])):
            if not valor.strip():
                pendencia(c, f"{campo} vazio", "")
            elif valor != valor.strip() or "  " in valor:
                pendencia(c, f"{campo} com espaço sobrando", repr(valor))
        if c["cargo"].strip() and perfil == SEM_PERFIL:
            pendencia(c, "Cargo fora da aba Cargos", c["cargo"].strip())
        if refs.setores and c["setor"].strip() and chave(c["setor"]) not in refs.setores:
            pendencia(c, "Setor fora da lista", c["setor"].strip())
        if refs.empreendimentos and c["empreendimento"].strip() and chave(c["empreendimento"]) not in refs.empreendimentos:
            pendencia(c, "Empreendimento fora da lista", c["empreendimento"].strip())
        if tenant == SEM_PERFIL:
            pendencia(c, "Domínio sem tenant definido", dominio)
        for sku in sem_cadastro:
            pendencia(c, "Licença sem cadastro em licencas.csv", sku)
        if perfil != SEM_PERFIL:
            if chave(perfil) == PERFIL_SEM_LICENCA:
                if pagas:
                    pendencia(c, "Licença diferente da prevista", f"tem {licenses}; prevista: {perfil}")
            elif _conjunto(perfil) != _conjunto(licenses):
                pendencia(c, "Licença diferente da prevista", f"tem {licenses or '(nenhuma)'}; prevista: {perfil}")

    caixas = pd.DataFrame(linhas, columns=COLUNAS).sort_values("Display name", key=lambda s: s.str.casefold())
    pendencias = pd.DataFrame(pend, columns=["User principal name", "Display name", "Pendência", "Detalhe"])
    pendencias = pendencias.sort_values(["Pendência", "Display name"]).reset_index(drop=True)

    totais = {"contas no tenant": len(contas), "no relatório": len(caixas),
              "fora: sem licença paga": fora["sem licença paga"], "fora: bloqueadas": fora["bloqueadas"],
              "pendências": len(pendencias)}
    return Resultado(caixas.reset_index(drop=True), pendencias, pd.DataFrame(), totais,
                     custo_bloqueadas=round(custo_bloqueadas, 2), licencas_sem_preco=sem_preco)


# ---------------------------------------------------------------- custos

AGRUPAMENTOS = [("Por licença atual", "Licenses"), ("Por Licença 2026", "Licença 2026"),
                ("Por tenant novo", "Tenant Novo"), ("Por empreendimento", "Empreendimento"),
                ("Por setor", "Setor"), ("Por cargo", "Cargo")]


def projetar(caixas: pd.DataFrame) -> pd.DataFrame:
    """Contas sem Licença 2026 (cargo vazio ou fora da lista) mantêm o custo atual na projeção."""
    df = caixas.copy()
    df[CUSTO_2026] = df[CUSTO_2026].fillna(df[CUSTO_ATUAL])
    return df


def custos(caixas: pd.DataFrame) -> list[tuple[str, pd.DataFrame]]:
    """Tabelas de custo mensal agrupadas; ordenadas pelo custo atual, do maior para o menor."""
    tabelas = []
    base = projetar(caixas)
    for titulo, coluna in AGRUPAMENTOS:
        df = base.copy()
        df[coluna] = df[coluna].replace("", "(vazio)")
        g = df.groupby(coluna).agg(**{
            "Contas": ("User principal name", "count"),
            CUSTO_ATUAL: (CUSTO_ATUAL, "sum"),
            CUSTO_2026: (CUSTO_2026, "sum"),
        }).reset_index().rename(columns={coluna: titulo.replace("Por ", "").capitalize()})
        g["Diferença (mês)"] = (g[CUSTO_2026] - g[CUSTO_ATUAL]).round(2)
        g = g.sort_values(CUSTO_ATUAL, ascending=False).reset_index(drop=True)
        tabelas.append((titulo, g))
    return tabelas


# ---------------------------------------------------------------- retratos

def comparar_com_retrato(resultado: Resultado, pasta: Path) -> None:
    """Compara com o retrato mais recente, grava o novo e apaga os antigos."""
    pasta.mkdir(exist_ok=True)
    anteriores = sorted(pasta.glob("retrato_*.csv"))
    atual = resultado.caixas
    mudancas = []
    if anteriores:
        antes = pd.read_csv(anteriores[-1], sep=";", dtype=str, encoding="utf-8-sig").fillna("")
        base = anteriores[-1].stem.replace("retrato_", "")
        a = antes.set_index(antes["User principal name"].str.lower())
        d = atual.set_index(atual["User principal name"].str.lower())
        for upn in d.index.difference(a.index):
            mudancas.append({"User principal name": d.at[upn, "User principal name"],
                             "Display name": d.at[upn, "Display name"], "Mudança": "Entrou no relatório",
                             "Campo": "", "Antes": "", "Depois": ""})
        for upn in a.index.difference(d.index):
            mudancas.append({"User principal name": a.at[upn, "User principal name"],
                             "Display name": a.at[upn, "Display name"], "Mudança": "Saiu do relatório",
                             "Campo": "", "Antes": "", "Depois": ""})
        for upn in d.index.intersection(a.index):
            for campo in CAMPOS_MUDANCA:
                v0, v1 = str(a.at[upn, campo]), str(d.at[upn, campo])
                if v0 != v1:
                    mudancas.append({"User principal name": d.at[upn, "User principal name"],
                                     "Display name": d.at[upn, "Display name"], "Mudança": "Alterado",
                                     "Campo": campo, "Antes": v0, "Depois": v1})
        resultado.totais["comparado com"] = base
    resultado.mudancas = pd.DataFrame(
        mudancas, columns=["User principal name", "Display name", "Mudança", "Campo", "Antes", "Depois"])
    resultado.totais["mudanças"] = len(resultado.mudancas)

    atual.to_csv(pasta / f"retrato_{datetime.now():%Y%m%d_%H%M}.csv", sep=";", index=False, encoding="utf-8-sig")
    limite = datetime.now() - timedelta(days=int(os.getenv("RETENCAO_RETRATOS_DIAS", "90")))
    for arq in pasta.glob("retrato_*.csv"):
        if datetime.fromtimestamp(arq.stat().st_mtime) < limite:
            arq.unlink()