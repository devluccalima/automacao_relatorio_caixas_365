<#
.SYNOPSIS
  Aplica cargo, departamento, escritório e nomes nas contas (Set-User), conta por conta.

.PARAMETER Entrada
  CSV (;) com as colunas Email;Cargo;Setor;Empreendimento;NomeProprio;Apelido;NomeApresentar.
  Campo vazio = não mexe.
.PARAMETER Saida
  CSV (;) gerado com Email;Resultado;Detalhe. Resultado: OK, SEM_MUDANCA, SIMULADO ou ERRO.
.PARAMETER Simular
  Só lê e diz o que mudaria; nada é gravado.

  Usa EXO_APP_ID, EXO_THUMBPRINT e EXO_ORGANIZATION do ambiente (o main.py os repassa do .env).
#>
param(
    [Parameter(Mandatory = $true)][string]$Entrada,
    [Parameter(Mandatory = $true)][string]$Saida,
    [switch]$Simular
)

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Coluna do CSV -> propriedade do Get-User / parâmetro do Set-User
$mapa = [ordered]@{
    Cargo          = 'Title'
    Setor          = 'Department'
    Empreendimento = 'Office'
    NomeProprio    = 'FirstName'
    Apelido        = 'LastName'
    NomeApresentar = 'DisplayName'
}

function Resultado($email, $resultado, $detalhe) {
    [pscustomobject]@{ Email = $email; Resultado = $resultado; Detalhe = $detalhe }
}

Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -AppId $env:EXO_APP_ID -CertificateThumbprint $env:EXO_THUMBPRINT `
    -Organization $env:EXO_ORGANIZATION -ShowBanner:$false

$linhas = Import-Csv -Path $Entrada -Delimiter ';' -Encoding UTF8
$total = @($linhas).Count
$i = 0
$resultados = foreach ($l in $linhas) {
    $i++
    $email = $l.Email.Trim()
    try {
        $u = Get-User -Identity $email -ErrorAction Stop
        # Identity (o nome curto) pode se repetir entre domínios; o GUID é único
        $id = $u.Guid.ToString()
        $params = @{}
        $mudancas = @()
        foreach ($coluna in $mapa.Keys) {
            $valor = [string]$l.$coluna
            if ([string]::IsNullOrWhiteSpace($valor)) { continue }
            $valor = $valor.Trim()
            $propriedade = $mapa[$coluna]
            $atual = [string]$u.$propriedade
            # -cne: diferencia maiúsculas, para corrigir "arquiteta" -> "Arquiteto(a)" e espaços sobrando
            if ($atual -cne $valor) {
                $params[$propriedade] = $valor
                $mudancas += "${coluna}: '$atual' -> '$valor'"
            }
        }

        if ($params.Count -eq 0) {
            Resultado $email 'SEM_MUDANCA' 'Já estava com esses valores'
        }
        elseif ($Simular) {
            Resultado $email 'SIMULADO' ($mudancas -join ' | ')
        }
        else {
            Set-User -Identity $id @params -Confirm:$false -ErrorAction Stop
            $depois = Get-User -Identity $id -ErrorAction Stop
            $falhas = @($params.Keys | Where-Object { [string]$depois.$_ -cne $params[$_] })
            if ($falhas.Count) {
                Resultado $email 'ERRO' ("Gravou, mas não confirmou: " + ($falhas -join ', '))
            }
            else {
                Resultado $email 'OK' ($mudancas -join ' | ')
            }
        }
    }
    catch {
        Resultado $email 'ERRO' $_.Exception.Message
    }
    if ($i % 25 -eq 0) { Write-Host "[exchange] $i de $total contas" }
}

$resultados | Export-Csv -Path $Saida -Delimiter ';' -NoTypeInformation -Encoding UTF8
Disconnect-ExchangeOnline -Confirm:$false
Write-Host "[exchange] concluído: $total contas"