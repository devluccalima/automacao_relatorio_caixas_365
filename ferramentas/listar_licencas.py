"""Lista as licenças do tenant e mostra quais ainda não estão em config/licencas.csv.

Uso: python ferramentas/listar_licencas.py
"""
import sys
from pathlib import Path

from dotenv import load_dotenv

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))
load_dotenv(RAIZ / ".env", override=True)

from core import graph, referencias  # noqa: E402

refs = referencias.carregar(RAIZ / "config")
print(f"{'SkuPartNumber':<40}{'Em uso':>8}{'Total':>8}   Cadastro")
for s in sorted(graph.get_paginado("/subscribedSkus"), key=lambda x: x["skuPartNumber"]):
    sku = s["skuPartNumber"]
    total = s.get("prepaidUnits", {}).get("enabled", 0)
    cad = refs.licencas.get(sku)
    situacao = f"{cad[0]} ({'paga' if cad[1] else 'gratuita'})" if cad else "FALTA CADASTRAR"
    print(f"{sku:<40}{s.get('consumedUnits', 0):>8}{total:>8}   {situacao}")
