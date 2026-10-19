# Spec: Homologação e diagnóstico mínimo (ciclo 005)

> Concluída em 18/10/2026 - feature/homologacao-e-diagnostico

## Objetivo

Fechar o resto de DEV-01 (testes e ferramentas não podem confundir banco de teste com banco de verdade),
DEV-25 (saber qual estação e conta chamou o conector) e DEV-26 (diagnóstico do ambiente por estação).
Depende do ciclo 004.

## Critérios de aceite

- [x] Teste que grava só roda em banco de HOMOLOGAÇÃO: `config.ambiente = TESTE` **e** `config.homologacao = 1`
  (`Assert-BancoDeHomologacao`); `criar-banco.ps1 -Ambiente TESTE` declara a marca; o `MONTAR-FRONTEND -Teste`
  a declara na cópia; `execucao\homologacao\` existe (`Get-PastaHomologacao`). Prova: `testar-preparacao-rede`.
- [x] Caminho canônico: UNC x letra x relativo x `\<máquina>\X$` x `\?\` x caixa contam como o mesmo arquivo
  (`Test-MesmoCaminho`), usado na recusa de `-CopiaDeTrabalho` sobre a produção. Prova: `testar-build-lib`.
- [x] `modConfig.CaminhoBanco` não tem mais fallback: sem `Config!B2` devolve vazio, `CaminhoConfigurado` é falso e
  `BancoOK` explica a causa. Prova: `Confere` no `modAutoteste` (160 testes) no build `-Teste`.
- [x] Diagnóstico do conector: um arquivo por estação (`IA\logs\mcp-<ESTACAO>-AAAAMMDD.log`), cada linha com
  `[ESTACAO\usuário]` e a ferramenta; guarda 90 dias. Prova: `testar-preparacao-rede`.
- [x] `checar-ambiente` grava o relatório e confere conector (Claude Desktop), pasta IA e OneDrive. Prova: `testar-preparacao-rede`.

## Decisões assumidas novas

- Ambiente TESTE sozinho não basta para gravar: exige a marca `homologacao` (declarada, não deduzida). Bancos de teste
  antigos ganham a marca pelo `MONTAR-FRONTEND -Teste` ou recriando com `criar-banco-teste`.
- Sem fallback de caminho no front: planilha sem `Config!B2` não abre banco nenhum (falha visível em vez de apontar
  para um caminho de reserva).
- Logs do conector: 90 dias, apagados na abertura do servidor (só os do próprio prefixo).
- OneDrive em Documentos é só aviso (a cópia local fica em Documentos por decisão de projeto).

## Itens 🖥

- Em cada estação: `SISTEMA\DADOS\testes\CHECAR-AMBIENTE.bat` e ler o relatório (conector e OneDrive).
- Abrir o front pelo `INICIAR-CRM` e conferir que conecta (Config!B2 vem gravado pelo build).

## Limites

Nunca gravar em banco sem `homologacao`; nunca apontar o front para um caminho que o build não gravou.
