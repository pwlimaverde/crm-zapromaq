# Spec: Texto literal e consumidores do conector (ciclo 013)

> Situação: aprovada em 22/10/2026 (plano revisado de 19/10/2026: DEV-24 e matriz de consumidores do DEV-07; sem fila
> de exceções). Sem migração; o dado gravado não muda.

## Objetivo

- **DEV-24 — texto literal em grade e CSV.** Texto do cadastro que começa com `=`, `+`, `-`, `@` (ou tabulação) nunca
  vira fórmula:
  - grade (`modGrade.Montar`): colunas de texto (e as ocultas `_b`/`_a`) recebem formato Texto **antes** de preencher;
    moeda, data e número continuam numéricos;
  - CSV para planilha (`modRelatorio.CampoCsv`, os CSVs do relatório): campo de texto iniciado por `=`, `+`, `-`, `@` ou
    tabulação ganha um apóstrofo na frente; número e data não mudam. A exportação fiel para integração continua sendo
    o `exportar_base` do conector (sem neutralização).
- **DEV-07 — matriz de consumidores.** `IA\teste\consumidores.json`: cada skill do comercial (seção 6 do documento-base)
  com as ferramentas e os parâmetros de que depende, o que a muda e o aviso da versão atual. O `testar-mcp` confere que
  toda ferramenta e todo parâmetro citados existem no catálogo (falha se uma mudança de contrato quebrar um consumidor
  sem a matriz ser atualizada junto). `IA\LEIA-ME.md` aponta a matriz.

## Critérios de aceite

- [ ] `Confere` em `modRelatorio.CampoCsv`: `=1+1`, `+5541`, `-X`, `@X` e tabulação no início saem com apóstrofo;
  texto comum, número negativo e data não mudam.
- [ ] testar-uso: empresas `=1+1`, `+5541`, `-X` e `@X` aparecem como texto na grade de Clientes (sem fórmula); o
  valor do atendimento continua número.
- [ ] testar-mcp: matriz de consumidores contra o catálogo (ferramenta e parâmetro existentes).
- 🖥 CSV do relatório aberto no Excel 2019 da estação: `'=1+1` aparece como texto; moeda e data numéricas.

## Decisões assumidas

| # | Decisão | Premissa | O que mudaria |
|---|---|---|---|
| 013-1 | Neutralização por apóstrofo só nos CSVs do relatório (para planilha); roteiro e `exportar_base` seguem fiéis. | DEV-24 cita `modGrade` e `modRelatorio`; roteiro/extrato são entrada de scripts e de rota. | Neutralizar também o roteiro, se for aberto direto no Excel. |
| 013-2 | Sem fila de exceções (plano revisado). | Decisão do plano. | — |
| 013-3 | Matriz em JSON versionado no repositório (não há acesso às skills do comercial daqui); parâmetros listados pelo uso descrito no documento-base. | Skills ficam fora de BASE. | Atualizar a matriz quando o comercial mandar as skills. |

## Limites

- Sempre: teste primeiro; teste que grava só em banco TESTE; nada de dado real.
- Nunca: alterar o dado gravado para "consertar" a exibição.
