# Plano: Tipos e resultado da gravação (rascunho)

Spec: [SPEC.md](SPEC.md)

## Fatias (teste que falha primeiro, depois prova)

- parse que distingue vazio, válido e inválido (matriz 31/02, abc, 12.50, 1,5)
- estados não enviada / falhou / confirmada / indeterminada, sem "nada foi gravado" falso
- erro nativo e repetição só antes do commit (testar-mcp com falha injetada)

## Depende de

005