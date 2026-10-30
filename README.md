# Relatório de Caixas — Microsoft 365

Gera automaticamente a planilha de caixas de e-mail (nome, cargo, setor, empreendimento, licença atual,
licença 2026 e tenant novo) a partir do Entra ID e do Exchange Online, aponta o que está fora do padrão e,
a partir da fase 3, aplica alterações em massa listadas numa aba de alterações.

## Ciclo (diário, de manhã)

1. Ler o SharePoint: tabelas de referência e aba Alterações
2. Push: validar e aplicar as alterações (`Set-User`), com simulação e limite
3. Pull: ler Entra e Exchange
4. Analisar: licença 2026, tenant novo, pendências, mudanças desde o último retrato
5. Gravar o relatório no SharePoint e o retrato no servidor
6. E-mail: só quando houver novidade, mais um resumo semanal

Só o passo 2 escreve no Exchange, e só o que estiver na aba Alterações.

## Fases

| Fase | Entrega | Status |
|---|---|---|
| 1 | Pull e relatório Excel local | Pronto |
| 2 | Relatório no SharePoint e e-mail de resumo | SharePoint pronto; e-mail pendente |
| 3 | Push pelo Alteracoes.xlsx (em simulação primeiro) | Pronto, em teste |
| 4 | Carga inicial: planilha atual → aba Alterações | — |
| 5 | Agendamento diário no servidor | Em implantação |

## Estrutura

```
relatorio_caixas_m365/
├── main.py                  Orquestra o ciclo
├── requirements.txt
├── .env / .env.example      Segredos e controles (.env não é versionado)
├── config/
│   ├── cargos_licencas.csv  Cargo;Perfil (fase 1; depois vem do SharePoint)
│   ├── dominios_tenants.csv Dominio;Tenant novo
│   ├── licencas.csv         SkuPartNumber;Nome;Paga;PrecoMensal (a ordem das linhas é a ordem de exibição)
│   ├── setores.csv          Lista padrão de setores (vazia = não valida)
│   └── empreendimentos.csv  Lista padrão de empreendimentos (vazia = não valida)
├── core/
│   ├── graph.py             Autenticação e chamadas ao Graph (com nova tentativa em 429/5xx)
│   ├── sharepoint.py        Sobe e baixa arquivos da pasta do site (Sites.Selected)
│   ├── alteracoes.py        Push: lê, valida e aplica o Alteracoes.xlsx
│   ├── autocorrecao.py      Classifica valores: oficial, autocorreção, sugestão ou sem correspondência
│   ├── exchange.py          Chama o scripts\aplicar_contatos.ps1
│   ├── referencias.py
│   ├── leitura_tenant.py
│   ├── analise.py
│   └── relatorio.py
├── scripts/                 aplicar_contatos.ps1 (Set-User com verificação),
│                            executar_automacao.ps1 (usado pelo Agendador), registrar_tarefa_agendada.ps1
├── carga_inicial/           padronizacao_final.xlsx (entrada) e Alteracoes_carga_inicial.xlsx (saída)
├── ferramentas/             listar_licencas.py, testar_sharepoint.py, carga_inicial.py
├── relatorios/              Saída (não versionado)
├── retratos/                Um CSV por execução (não versionado)
└── logs/
```

## Preparar o ambiente

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
Copy-Item .env.example .env
```

## Rodar (fase 1)

```powershell
.\.venv\Scripts\python.exe ferramentas\listar_licencas.py   # confere se todas as licenças do tenant estão em licencas.csv
.\.venv\Scripts\python.exe main.py                           # gera relatorios\Relatorio_Caixas_AAAAMMDD_HHMM.xlsx
```

O Excel tem quatro abas:

| Aba | Conteúdo |
|---|---|
| Resumo | Totais da execução, pendências por tipo e saldo das licenças pagas (compradas, em uso, livres) |
| Caixas | O formato da planilha de caixas, com Licença 2026, Tenant Novo, Bloqueado e o custo mensal atual e 2026 de cada conta |
| Custos | Totais (mês e ano, atual e 2026, licenças sem uso, licenças em contas bloqueadas) e tabelas por licença atual, Licença 2026, tenant novo, empreendimento, setor e cargo |
| Pendências | Cargo/setor/empreendimento vazio ou fora do padrão, espaço sobrando, licença diferente da prevista, domínio sem tenant, licença sem cadastro, bloqueada com licença paga |
| Mudanças | O que mudou desde a execução anterior (entrou, saiu, campo alterado) |

Regras:

- Entram no relatório as contas membro (convidados ficam de fora) com pelo menos uma licença paga e o login ativo.
  `INCLUIR_SEM_LICENCA_PAGA` e `INCLUIR_BLOQUEADAS` no `.env` mudam isso.
- "Bloqueada com licença paga" é verificada em todas as contas, mesmo as que ficam fora do relatório.
- Cargo, domínio e licenças são comparados como o PROCX: sem diferenciar maiúsculas e ignorando espaços nas pontas.
- A coluna Licenses mostra só licenças pagas, na ordem de `licencas.csv`. Licença que não está no arquivo aparece pelo
  código (SkuPartNumber), conta como paga e gera pendência.
- Perfil "Desinstalar" na aba Cargos significa que a conta deveria estar sem licença paga (custo 2026 = 0).
- Custos: preço mensal por licença em `licencas.csv` (aceita `70,73`). Uma combinação custa a soma das licenças.
  Contas sem Licença 2026 entram na projeção com o custo atual. Licença paga sem preço conta 0 e aparece na aba Custos.
- Cada execução grava um retrato em `retratos\`; os mais velhos que `RETENCAO_RETRATOS_DIAS` são apagados.

## Alterações (push)

O `Alteracoes.xlsx` fica na mesma pasta do SharePoint do relatório (é criado vazio na primeira execução).
Na aba **Pendentes**, uma linha por conta: E-mail e só os campos que devem mudar (vazio = não mexe).
Cargo, Setor e Empreendimento têm lista suspensa com os valores de `config\`.

A cada execução, antes do pull:

| Situação da linha | O que acontece |
|---|---|
| Valor fora da lista, e-mail repetido ou nada para alterar | Fica em Pendentes com Status "Inválida" e o motivo |
| Além de `LIMITE_ALTERACOES` | Fica com Status "Aguardando" para a próxima execução |
| `DRY_RUN=true` | O Exchange só é lido; Status "Simulado" mostra o que mudaria |
| Gravada e confirmada (OK) ou já estava igual (SEM_MUDANCA) | Sai de Pendentes e vai para o Histórico |
| Erro no Exchange | Fica em Pendentes com o erro no Status e vai também para o Histórico |

O Exchange grava no Entra, mas o Graph pode levar alguns minutos para refletir: a mudança às vezes só aparece no
relatório da execução seguinte.

Rodar só as alterações, na hora: `python main.py --somente-alteracoes`. Sempre simula primeiro e mostra o resultado;
com `DRY_RUN=false`, pede para digitar APLICAR antes de gravar.

### Autocorreção e sugestões

A cada execução, cada cargo, setor e empreendimento das contas do relatório é comparado com as listas oficiais
(`core\autocorrecao.py`):

| Caso | Exemplo | O que acontece |
|---|---|---|
| Só caixa, acento, espaço, "(a)", feminino/masculino ou abreviação | "arquiteta", "Coord. de TI", "brasilia shopping" | **Autocorreção**: entra em Pendentes com Observação "Autocorreção" e é aplicada na mesma execução |
| Está em `config\apelidos.csv` | "Tec. Segurança do Trabalho" | **Autocorreção** (é uma decisão sua já registrada) |
| Parecido com um valor oficial (≥ 80%) | "Gerente de Contrato", "Finaceiro" | **Sugestão**: vai para a aba Sugestões; só é aplicada se você escrever S em "Aprovar (S)" |
| Nada parecido | "Cargo Maluco" | Continua só como pendência no relatório |

A autocorreção nunca passa por cima de uma linha que alguém já colocou em Pendentes para a mesma conta.
`AUTOCORRIGIR=false` no `.env` desliga a autocorreção (as sugestões continuam aparecendo).

`config\apelidos.csv` (Campo;Valor errado;Valor correto;Observação) começou com as decisões da padronização.
Para ensinar um erro novo, acrescente uma linha no repositório. As linhas com "Conferir" na Observação podem ter
sido decisões para uma pessoa específica; se não valem como regra geral, apague-as.

### Carga inicial

1. `python ferramentas\carga_inicial.py` lê `carga_inicial\padronizacao_final.xlsx` (de-para de cargo, setor e
   empreendimento) e gera `carga_inicial\Alteracoes_carga_inicial.xlsx`, com a aba **Conferência** (atual e novo,
   conta por conta) e **Sem destino** (valores que não estão no de-para nem na lista).
2. Conferido, `python ferramentas\carga_inicial.py --publicar` envia como `Alteracoes.xlsx` ao SharePoint (recusa se
   lá já houver linhas pendentes).
3. Simule com `python main.py --somente-alteracoes` e, depois, aplique em lotes (`LIMITE_ALTERACOES`).

## SharePoint

Com `SHAREPOINT_SITE_ID` preenchido, o `main.py` sobe o relatório para `SHAREPOINT_PASTA` sempre com o nome
`Relatorio_Caixas.xlsx`, sobrescrevendo o anterior: o link compartilhado não muda, e as versões antigas ficam no
histórico de versões da biblioteca. A cópia local com data continua em `relatorios\`.

Se o arquivo estiver aberto no Excel desktop e o SharePoint recusar a gravação, o script tenta de novo 4 vezes, com 1 minuto
entre elas, e, se não conseguir, grava `Relatorio_Caixas (AAAAMMDD_HHMM).xlsx` e avisa.

Para testar o acesso: `python ferramentas\testar_sharepoint.py` (cria a pasta, se preciso, e sobe um arquivo de teste).

O app recebe acesso só ao site, com o PowerShell de administrador:

```powershell
Connect-MgGraph -Scopes "Sites.FullControl.All" -UseDeviceCode
$site = Get-MgSite -SiteId "<tenant>.sharepoint.com:/sites/<site>"
$params = @{ roles = @("write"); grantedToIdentities = @(@{ application = @{ id = "<AZURE_CLIENT_ID>"; displayName = "Automacao-Caixas-M365" } }) }
New-MgSitePermission -SiteId $site.Id -BodyParameter $params
```

## Servidor e agendamento

Roda no mesmo servidor e com a mesma conta de serviço da automação de contas inativas, que já tem o Python para todos os
usuários, o módulo ExchangeOnlineManagement e o certificado do Exchange no perfil.

1. Copie o projeto para `C:\Automacoes\automacao_relatorio_caixas` (sem `.venv`, `relatorios` e `logs`). Leve junto o
   retrato mais recente de `retratos\`, para a aba Mudanças continuar comparando.
2. No PowerShell elevado, na pasta do projeto:
   ```powershell
   & "C:\Program Files\Python313\python.exe" -m venv .venv
   .\.venv\Scripts\python.exe -m pip install --upgrade pip
   .\.venv\Scripts\python.exe -m pip install -r requirements.txt
   icacls C:\Automacoes\automacao_relatorio_caixas /grant "SHOPPING\svc.autm365:(OI)(CI)M"
   ```
3. Copie o `.env` e trave as permissões dele:
   `icacls .env /inheritance:r /grant "Administrators:(F)" /grant "SYSTEM:(F)" /grant "SHOPPING\svc.autm365:(R)"`
4. Rede: além do que a automação de inativos já usa, saída HTTPS para `cscshoppingspauloctavio.sharepoint.com`.
5. Teste como a conta de serviço: `.\.venv\Scripts\python.exe main.py`.
6. Registre a tarefa: `.\scripts\registrar_tarefa_agendada.ps1 -Usuario "SHOPPING\svc.autm365"` (padrão: todo dia às 06:00).

Acompanhar e operar (PowerShell elevado):

```powershell
Start-ScheduledTask -TaskName "Relatorio Caixas M365"
Get-ScheduledTaskInfo -TaskName "Relatorio Caixas M365" | Select LastRunTime, LastTaskResult, NextRunTime
Get-Content (Get-ChildItem logs\*.log | Sort LastWriteTime | Select -Last 1) -Tail 40
Disable-ScheduledTask -TaskName "Relatorio Caixas M365"   # pausar (Enable- para voltar)
```

`LastTaskResult` 0 = sucesso; 267009 = em execução; 267011 = ainda não rodou.

Depois da implantação, o servidor é o único lugar que roda o ciclo de verdade. No seu PC, para desenvolver, deixe
`SHAREPOINT_SITE_ID` vazio no `.env`: assim o relatório e o `Alteracoes.xlsx` ficam só locais e nada concorre com o
servidor.

## Permissões

Registro de aplicativo próprio (separado da automação de inativos):
`User.Read.All` e `Organization.Read.All`; na fase 2, `Sites.Selected` e `Mail.Send`.
Exchange: certificado com `Exchange.ManageAsApp` (a escrita da fase 3 usa `Set-User`).