# Plano: Texto literal e consumidores do conector (ciclo 013)

Spec: [SPEC.md](SPEC.md)

## Fatias (teste que falha primeiro, depois prova)

1. **DEV-24 CSV** — `Confere` de `CampoCsv` (falha: sai `=1+1`), neutralização em `modRelatorio`.
2. **DEV-24 grade** — testar-uso com empresas de prefixo perigoso (falha: célula vira fórmula/valor), formato Texto em
   `modGrade.Montar`.
3. **DEV-07 matriz** — `consumidores.json` + caso no testar-mcp; LEIA-ME.
4. Docs (stack, PLANO D32) e `verificar-tudo -Completo`.

## Checkpoints 🖥

- CSV do relatório aberto no Excel 2019 da estação.
