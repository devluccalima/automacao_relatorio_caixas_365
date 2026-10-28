"""Testa o acesso ao SharePoint: cria a pasta se preciso, sobe um arquivo de teste e lista a pasta.

Uso: python ferramentas/testar_sharepoint.py
"""
import sys
import tempfile
from datetime import datetime
from pathlib import Path

from dotenv import load_dotenv

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))
load_dotenv(RAIZ / ".env", override=True)

from core import sharepoint  # noqa: E402

if not sharepoint.configurado():
    sys.exit("SHAREPOINT_SITE_ID não está no .env")

with tempfile.TemporaryDirectory() as tmp:
    arq = Path(tmp) / "teste.txt"
    arq.write_text(f"Teste de acesso em {datetime.now():%d/%m/%Y %H:%M:%S}\n", encoding="utf-8")
    r = sharepoint.enviar_arquivo(arq, "_teste_acesso.txt")
print(f"Upload OK: {r['url']}")
print("\nConteúdo da pasta:")
for item in sharepoint.listar_pasta():
    print(f"  {item['name']:<45} {item.get('size', 0):>10} bytes   {item['lastModifiedDateTime']}")
print("\nPode apagar o _teste_acesso.txt pelo navegador.")