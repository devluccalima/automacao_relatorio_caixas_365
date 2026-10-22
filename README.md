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
| 1 | Pull e relatório Excel local | Em andamento |
| 2 | Relatório no SharePoint e e-mail de resumo | — |
| 3 | Push pela aba Alterações (em simulação primeiro) | — |
| 4 | Carga inicial: planilha atual → aba Alterações | — |
| 5 | Agendamento diário no servidor | — |

## Estrutura

```
relatorio_caixas_m365/
├── main.py                  Orquestra o ciclo
├── requirements.txt
├── .env / .env.example      Segredos e controles (.env não é versionado)
├── config/
│   ├── cargos_licencas.csv  Cargo;Perfil (fase 1; depois vem do SharePoint)
│   └── dominios_tenants.csv Dominio;Tenant novo
├── core/
│   ├── graph_auth.py        Copiar da automação de contas inativas
│   ├── referencias.py
│   ├── leitura_tenant.py
│   ├── analise.py
│   └── relatorio.py
├── scripts/                 extrai_tipos_caixa.ps1 (copiar da automação de inativos)
├── ferramentas/             carga_inicial.py (fase 4)
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

## Permissões

Registro de aplicativo próprio (separado da automação de inativos):
`User.Read.All` e `Organization.Read.All`; na fase 2, `Sites.Selected` e `Mail.Send`.
Exchange: certificado com `Exchange.ManageAsApp` (a escrita da fase 3 usa `Set-User`).
