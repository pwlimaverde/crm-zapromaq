# Plano: Tipos, resultado da gravação e falha transitória (ciclo 008)

Spec: [SPEC.md](SPEC.md)

## Ordem e dependências

1. DEV-09 conversor do front (`modValidacao.Converter`) + `MontarDados` sem perda + lote — depende de: —
2. DEV-09 conector (`ConvertTo-Moeda`, `ConvertTo-Inteiro`) + casos C17–C21 nos dois lados — depende de: 1
3. DEV-13 estados da gravação no front (modDB + ficha) — depende de: —
4. DEV-14 repetição de erro transitório (conector e front) + mensagem — depende de: 3
5. Docs (contrato, nota de mudança, stacks), `verificar-tudo -Completo` — depende de: 1-4

## Checkpoints

- 🖥 digitar `12.50`, `31/02/2026`, `1,5` (tentativas) na ficha e ver a recusa sem perder o texto.
- 🖥 duas estações gravando o mesmo atendimento ao mesmo tempo: uma espera e grava, sem erro.

## Riscos

- `Like "#"` e `CCur` dependem da localidade só para texto com separador: o conversor monta o número a partir dos dígitos.
