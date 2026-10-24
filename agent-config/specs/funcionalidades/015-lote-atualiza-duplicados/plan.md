# Plano: Lote atualiza duplicados por CNPJ (ciclo 015)

Spec: [SPEC.md](SPEC.md)

## Fatias (teste que falha primeiro, depois prova)

1. **Regra pura** — `Confere` de `modLote.MudancasDoLote` (falha: não existe).
2. **Gravação** — testar-uso com empresa ativa, inativa e linha repetida (falha: duplicado não atualiza);
   `modCRM.AtualizarCampos` (só os campos dados, versão, log com a ação) e `modLote.AtualizarDuplicado(s)`.
3. **Tela** — `cmdAtualizar` no `frmLote.json`/`.txt`, duplo clique com a diferença; prévia regerada.
4. Docs (stack front, PLANO D34) e `verificar-tudo -Completo`.

## Checkpoints 🖥

- frmLote na estação (botão, confirmação da diferença).
