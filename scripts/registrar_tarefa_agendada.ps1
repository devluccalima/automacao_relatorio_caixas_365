<#
.SYNOPSIS
  Cria (ou recria) a tarefa diária do relatório de caixas, rodando como a conta de serviço.

.EXAMPLE
  .\scripts\registrar_tarefa_agendada.ps1 -Usuario "SHOPPING\svc.autm365"
.EXAMPLE
  .\scripts\registrar_tarefa_agendada.ps1 -Usuario "SHOPPING\svc.autm365" -Horario 06:30 -PrimeiraExecucao 2026-10-29
#>
param(
    [Parameter(Mandatory = $true)][string]$Usuario,
    [string]$Horario = '06:00',
    [datetime]$PrimeiraExecucao = (Get-Date).Date.AddDays(1),
    [string]$NomeTarefa = 'Relatorio Caixas M365'
)

$ErrorActionPreference = 'Stop'
$raiz = Split-Path -Parent $PSScriptRoot
$script = Join-Path $raiz 'scripts\executar_automacao.ps1'
if (-not (Test-Path $script)) { throw "Não encontrei $script" }

$credencial = Get-Credential -UserName $Usuario -Message "Senha da conta de serviço $Usuario"
$inicio = $PrimeiraExecucao.Date + [timespan]$Horario

$acao = New-ScheduledTaskAction -Execute 'powershell.exe' `
    -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$script`"" -WorkingDirectory $raiz
$gatilho = New-ScheduledTaskTrigger -Daily -At $inicio
# StartWhenAvailable: se o servidor estiver desligado no horário, roda assim que ligar.
# IgnoreNew: nunca duas execuções ao mesmo tempo.
$config = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit (New-TimeSpan -Hours 2)

Register-ScheduledTask -TaskName $NomeTarefa -Action $acao -Trigger $gatilho -Settings $config `
    -User $Usuario -Password $credencial.GetNetworkCredential().Password -RunLevel Limited -Force | Out-Null

Write-Host "Tarefa '$NomeTarefa' registrada: todo dia às $Horario, a partir de $($inicio.ToString('dd/MM/yyyy HH:mm')), como $Usuario."
Get-ScheduledTask -TaskName $NomeTarefa | Get-ScheduledTaskInfo | Select-Object LastRunTime, LastTaskResult, NextRunTime