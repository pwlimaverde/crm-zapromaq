# Plano: Esquema 1.2 (ciclo 010)

Spec: [SPEC.md](SPEC.md)

## Ordem e dependências

1. Migração 004 (inventário + estrutura) e teste em banco fictício com casos legados — depende de: —
2. Esquema 1.2 em todo lugar (front, conector, montagem, testes) — depende de: 1
3. DEV-16 anterior sem ciclo (C33) e `codigo_atendimento_anterior` no `atualizar_atendimento` — depende de: 2
4. DEV-06B autor técnico no log (front e conector) — depende de: 2
5. DEV-17 `ctx_versao` na gravação de contexto — depende de: 2
6. DEV-13 identificador de operação (conector, ficha, lote) — depende de: 2
7. Docs, nota de mudança (conector 1.8.0), `verificar-tudo -Completo` — depende de: 1-6

## Checkpoints

- 🖥 `MONTAR-FRONTEND.bat -Ensaio` na rede: inventário da 004 sem achados e esquema 1.1 → 1.2 na cópia.

## Riscos

- Banco da rede com `controle` repetido ou anterior inconsistente: a migração para (correção operacional antes da 1.6).
