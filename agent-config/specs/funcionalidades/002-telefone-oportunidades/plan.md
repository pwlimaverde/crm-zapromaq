# Plano: coluna Telefone na grade de Oportunidades

Spec: [SPEC.md](SPEC.md)

## Ordem e dependências

1. Teste que falha (RED): `Confere` no modAutoteste + caso no `testar-uso.ps1` + banco demo com
   contato sem telefone e empresa sem telefone — depende de: —
2. Regra + coluna (GREEN): SELECT com `telefone_empresa`, `TelefoneDoAtendimento`, coluna
   calculada em `ColunasPlanilha`/`CelulaDaColuna`, `verificar.py` — depende de: 1
   (fatia única: grade e feed passam pelo mesmo `CelulaDaColuna`)
3. Spec de stack + PLANO (D20) — depende de: 2

## Checkpoints

- Depois da fatia 2: build `-Teste` com `-Completo` verde; 🖥 na estação (largura, AutoFiltro, feed).

## Riscos

- Largura total da grade mínima: +15 de coluna; Empresa cai de 42 para 36 para seguir cabendo
  em 1366×768 (🖥 conferir).
- Banco demo muda: contato/empresa sem telefone não podem quebrar outros casos do `testar-uso`.
