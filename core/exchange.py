"""Chama o scripts/aplicar_contatos.ps1 e devolve o resultado de cada conta."""
import os
import subprocess
import tempfile
from pathlib import Path

import pandas as pd

RAIZ = Path(__file__).resolve().parents[1]
SCRIPT = RAIZ / "scripts" / "aplicar_contatos.ps1"
COLUNAS = ["Email", "Cargo", "Setor", "Empreendimento", "NomeProprio", "Apelido", "NomeApresentar"]


def aplicar(linhas: list[dict], simular: bool) -> pd.DataFrame:
    """linhas: dicts com as chaves de COLUNAS. Devolve DataFrame Email;Resultado;Detalhe."""
    for var in ("EXO_APP_ID", "EXO_THUMBPRINT", "EXO_ORGANIZATION"):
        if not os.getenv(var, "").strip():
            raise RuntimeError(f"{var} não está no .env (necessário para gravar no Exchange)")
    exe = os.getenv("POWERSHELL_EXE", "").strip() or "powershell.exe"
    with tempfile.TemporaryDirectory() as tmp:
        entrada, saida = Path(tmp) / "entrada.csv", Path(tmp) / "saida.csv"
        pd.DataFrame(linhas, columns=COLUNAS).fillna("").to_csv(
            entrada, sep=";", index=False, encoding="utf-8-sig")
        cmd = [exe, "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(SCRIPT),
               "-Entrada", str(entrada), "-Saida", str(saida)]
        if simular:
            cmd.append("-Simular")
        limite = 300 + 10 * len(linhas)
        proc = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                              errors="replace", timeout=limite)
        for linha in proc.stdout.splitlines():
            if linha.startswith("[exchange]"):
                print(f"  {linha}")
        if proc.returncode != 0 or not saida.exists():
            raise RuntimeError(f"O script do Exchange falhou (código {proc.returncode}):\n"
                               f"{(proc.stderr or proc.stdout)[-1500:]}")
        return pd.read_csv(saida, sep=";", dtype=str, encoding="utf-8-sig").fillna("")