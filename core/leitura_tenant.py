"""Pull: usuários do Entra ID com atributos de contato, estado do login e licenças."""
from core import graph

CAMPOS = ",".join([
    "id", "displayName", "userPrincipalName", "givenName", "surname", "jobTitle",
    "department", "officeLocation", "accountEnabled", "assignedLicenses", "userType",
    "createdDateTime",
])


def ler_skus() -> list[dict]:
    """Licenças assinadas pelo tenant, com quantidade comprada e em uso."""
    return [{"id": s["skuId"], "sku": s["skuPartNumber"],
             "total": s.get("prepaidUnits", {}).get("enabled", 0), "em_uso": s.get("consumedUnits", 0)}
            for s in graph.get_paginado("/subscribedSkus")]


def ler_contas(lista_skus: list[dict] | None = None) -> list[dict]:
    """Todas as contas membro (convidados ficam de fora), com as licenças como skuPartNumber."""
    skus = {s["id"]: s["sku"] for s in (lista_skus if lista_skus is not None else ler_skus())}
    usuarios = graph.get_paginado("/users", {"$select": CAMPOS, "$top": "999"})
    contas = []
    for u in usuarios:
        if (u.get("userType") or "Member") != "Member":
            continue
        contas.append({
            "nome": u.get("displayName") or "",
            "upn": u.get("userPrincipalName") or "",
            "nome_proprio": u.get("givenName") or "",
            "apelido": u.get("surname") or "",
            "cargo": u.get("jobTitle") or "",
            "setor": u.get("department") or "",
            "empreendimento": u.get("officeLocation") or "",
            "bloqueada": u.get("accountEnabled") is False,
            "skus": [skus.get(l["skuId"], l["skuId"]) for l in u.get("assignedLicenses", [])],
            "criada_em": u.get("createdDateTime") or "",
        })
    return contas
