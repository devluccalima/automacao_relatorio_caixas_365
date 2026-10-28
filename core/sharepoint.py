"""Leitura e gravação de arquivos na pasta do SharePoint (permissão Sites.Selected)."""
import os
import time
from pathlib import Path
from urllib.parse import quote

from core import graph
from core.graph import GraphErro

LIMITE_UPLOAD_SIMPLES = 4 * 1024 * 1024
TENTATIVAS_BLOQUEIO = 4
ESPERA_BLOQUEIO_S = 60


def configurado() -> bool:
    return bool(os.getenv("SHAREPOINT_SITE_ID", "").strip())


def _site() -> str:
    return os.environ["SHAREPOINT_SITE_ID"].strip()


def _pasta() -> str:
    return os.getenv("SHAREPOINT_PASTA", "").strip().strip("/")


def _caminho(nome: str) -> str:
    partes = [p for p in (_pasta(), nome) if p]
    return quote("/".join(partes))


def garantir_pasta() -> None:
    """Cria a pasta (e as subpastas) se ainda não existirem."""
    atual = ""
    for parte in [p for p in _pasta().split("/") if p]:
        caminho = f"{atual}/{parte}" if atual else parte
        try:
            graph.get(f"/sites/{_site()}/drive/root:/{quote(caminho)}")
        except GraphErro as e:
            if e.status != 404:
                raise
            pai = f"/sites/{_site()}/drive/root" + (f":/{quote(atual)}:" if atual else "")
            graph.requisicao("POST", f"{pai}/children",
                             json={"name": parte, "folder": {}, "@microsoft.graph.conflictBehavior": "fail"})
        atual = caminho


def _enviar(nome: str, conteudo: bytes) -> dict:
    alvo = f"/sites/{_site()}/drive/root:/{_caminho(nome)}"
    if len(conteudo) <= LIMITE_UPLOAD_SIMPLES:
        return graph.requisicao("PUT", f"{alvo}:/content", dados=conteudo,
                                cabecalhos={"Content-Type": "application/octet-stream"})
    sessao = graph.requisicao("POST", f"{alvo}:/createUploadSession",
                              json={"item": {"@microsoft.graph.conflictBehavior": "replace"}})
    url, total, bloco, inicio, resposta = sessao["uploadUrl"], len(conteudo), 5 * 1024 * 1024, 0, {}
    while inicio < total:
        fim = min(inicio + bloco, total) - 1
        import requests  # a URL da sessão já é autenticada; não leva o token
        r = requests.put(url, data=conteudo[inicio:fim + 1], timeout=120,
                         headers={"Content-Range": f"bytes {inicio}-{fim}/{total}"})
        r.raise_for_status()
        resposta = r.json() if r.content else {}
        inicio = fim + 1
    return resposta


def enviar_arquivo(local: Path, nome_destino: str) -> dict:
    """Sobe o arquivo sobrescrevendo o existente. Se ele estiver bloqueado (aberto no Excel desktop),
    tenta de novo algumas vezes e, se não conseguir, grava uma cópia com a data no nome.

    Devolve {"url": ..., "nome": ..., "copia": bool}.
    """
    garantir_pasta()
    conteudo = local.read_bytes()
    for tentativa in range(1, TENTATIVAS_BLOQUEIO + 1):
        try:
            item = _enviar(nome_destino, conteudo)
            return {"url": item.get("webUrl", ""), "nome": nome_destino, "copia": False}
        except GraphErro as e:
            if e.status not in (409, 423):
                raise
            print(f"  {nome_destino} está bloqueado (aberto por alguém); tentativa {tentativa}/{TENTATIVAS_BLOQUEIO}")
            if tentativa < TENTATIVAS_BLOQUEIO:
                time.sleep(ESPERA_BLOQUEIO_S)
    alternativo = f"{Path(nome_destino).stem} ({time.strftime('%Y%m%d_%H%M')}){Path(nome_destino).suffix}"
    item = _enviar(alternativo, conteudo)
    return {"url": item.get("webUrl", ""), "nome": alternativo, "copia": True}


def baixar_arquivo(nome: str, destino: Path) -> Path | None:
    """Baixa um arquivo da pasta. Devolve None se ele não existir (usado na fase 3)."""
    try:
        item = graph.get(f"/sites/{_site()}/drive/root:/{_caminho(nome)}")
    except GraphErro as e:
        if e.status == 404:
            return None
        raise
    import requests
    r = requests.get(item["@microsoft.graph.downloadUrl"], timeout=120)
    r.raise_for_status()
    destino.write_bytes(r.content)
    return destino


def listar_pasta() -> list[dict]:
    caminho = _pasta()
    base = f"/sites/{_site()}/drive/root" + (f":/{quote(caminho)}:" if caminho else "")
    return graph.get_paginado(f"{base}/children", {"$select": "name,size,lastModifiedDateTime,webUrl"})