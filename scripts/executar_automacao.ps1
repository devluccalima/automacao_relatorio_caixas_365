<#
.SYNOPSIS
  Usado pelo Agendador de Tarefas: roda o main.py e grava a saída em logs\execucao_AAAAMMDD_HHMM.log.
  Apaga logs com mais de 90 dias. Devolve o código de saída do Python (0 = sucesso).
#>
# 'Continue' de propósito: no Windows PowerShell 5.1, avisos que o Python escreve no stderr
# viram erro e interromperiam o script com 'Stop'.
$ErrorActionPreference = 'Continue'

$raiz = Split-Path -Parent $PSScriptRoot
Set-Location $raiz
$logs = Join-Path $raiz 'logs'
New-Item -ItemType Directory -Force -Path $logs | Out-Null
$log = Join-Path $logs ('execucao_{0:yyyyMMdd_HHmm}.log' -f (Get-Date))

$env:PYTHONUTF8 = '1'
$env:PYTHONIOENCODING = 'utf-8'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$python = Join-Path $raiz '.venv\Scripts\python.exe'
"=== Início: $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss') como $env:USERDOMAIN\$env:USERNAME ===" |
    Out-File -FilePath $log -Encoding utf8

if (-not (Test-Path $python)) {
    "ERRO: não encontrei $python. Crie o .venv (veja o README)." | Out-File -FilePath $log -Append -Encoding utf8
    exit 1
}

& $python (Join-Path $raiz 'main.py') *>&1 | ForEach-Object { "$_" } | Out-File -FilePath $log -Append -Encoding utf8
$codigo = $LASTEXITCODE

"=== Fim: $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss'), código $codigo ===" | Out-File -FilePath $log -Append -Encoding utf8

Get-ChildItem -Path $logs -Filter 'execucao_*.log' |
    Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-90) } |
    Remove-Item -Force

exit $codigo