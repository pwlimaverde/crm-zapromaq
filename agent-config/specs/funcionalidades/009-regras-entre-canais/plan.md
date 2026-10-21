# Plano: Regras iguais nos três canais (ciclo 009)

Spec: [SPEC.md](SPEC.md)

## Ordem e dependências

1. DEV-10 normalização da seção 5.1 (front, conector, lote) — C22, C23, C24 — depende de: —
2. DEV-10 transições de etapa, retomada e datas (front e conector) — C25 a C28 — depende de: 1
3. DEV-11 contato fixo e validação na transação — C29; `codigo_cliente_destino` removido — depende de: —
4. DEV-12 lote com as regras do cadastro — C30, C31 — depende de: 1
5. DEV-15 versão nas ações, reativar só a empresa, conflito por campo — C32 — depende de: —
6. Matriz de campos versionada + conferência no `verificar.py`; migração 005 de normalização — depende de: 1
7. Docs, nota de mudança (conector 1.7.0), `verificar-tudo -Completo` — depende de: 1-6

## Checkpoints

- 🖥 ficha: encerrar em Perdido e ver a próxima ação limpa; reabrir e ver a orientação; retomada em Negociação recusada.
- 🖥 lote: colar linha com 12 colunas, DV errado, UF "XX", nota 1,5: todas ERRO.
- 🖥 ficha antiga: desativar uma empresa alterada em outra estação é recusado.

## Riscos

- Maiúsculas em nomes e observações mudam a aparência do cadastro existente (migração 005 com prévia no relatório).
