"""Cliente mínimo do Microsoft Graph (credenciais de aplicativo)."""
import os
import time

import msal
import requests

GRAPH = "https://graph.microsoft.com/v1.0"
_token = {"valor": None, "expira": 0}


class GraphErro(RuntimeError):
    def __init__(self, status: int, mensagem: str):
        super().__init__(mensagem)
        self.status = status


def _obter_token() -> str:
    if _token["valor"] and time.time() < _token["expira"] - 60:
        return _token["valor"]
    app = msal.ConfidentialClientApplication(
        client_id=os.environ["AZURE_CLIENT_ID"],
        client_credential=os.environ["AZURE_CLIENT_SECRET"],
        authority=f"https://login.microsoftonline.com/{os.environ['AZURE_TENANT_ID']}",
    )
    resultado = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in resultado:
        raise RuntimeError(f"Falha ao obter token do Graph: {resultado.get('error_description', resultado)}")
    _token["valor"] = resultado["access_token"]
    _token["expira"] = time.time() + int(resultado.get("expires_in", 3600))
    return _token["valor"]


def requisicao(metodo: str, url: str, *, params=None, json=None, dados: bytes | None = None,
               cabecalhos: dict | None = None) -> dict:
    """Chamada ao Graph com nova tentativa em 429/5xx (respeita o Retry-After)."""
    if not url.startswith("http"):
        url = GRAPH + url
    for tentativa in range(5):
        headers = {"Authorization": f"Bearer {_obter_token()}", **(cabecalhos or {})}
        resp = requests.request(metodo, url, params=params, json=json, data=dados,
                                headers=headers, timeout=120)
        if resp.status_code == 429 or resp.status_code >= 500:
            time.sleep(int(resp.headers.get("Retry-After", 2 ** (tentativa + 1))))
            continue
        if not resp.ok:
            raise GraphErro(resp.status_code, f"Graph {resp.status_code} em {metodo} {url}: {resp.text[:300]}")
        return resp.json() if resp.content else {}
    raise GraphErro(0, f"Graph não respondeu após várias tentativas: {metodo} {url}")


def get(url: str, params: dict | None = None) -> dict:
    return requisicao("GET", url, params=params)


def get_paginado(url: str, params: dict | None = None) -> list[dict]:
    itens, pagina = [], get(url, params)
    itens.extend(pagina.get("value", []))
    while "@odata.nextLink" in pagina:
        pagina = get(pagina["@odata.nextLink"])
        itens.extend(pagina.get("value", []))
    return itens