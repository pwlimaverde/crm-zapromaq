# Plano: Feed e tela (ciclo 011)

Spec: [SPEC.md](SPEC.md)

## Fatias (teste que falha primeiro, depois prova)

1. **DEV-18 cursor e lacunas** — `Confere` de `NovasLacunas`/`ReferenciaInvalida` (falha: funções não existem);
   `modFeed.Processar` com estado só confirmado no fim; geração da base; testar-uso: falha injetada, commit fora de
   ordem, geração trocada.
2. **DEV-19 dependências e dia** — `Confere` de `Dependentes`; ids dependentes por consulta agrupada; Painel marcado
   velho; virada do dia; `UltimaCarga`/`DiaDoCalculo`; testar-uso: telefone do contato na linha do atendimento.
3. **DEV-20 filtro e ordem** — `modGrade.ReaplicarCriterios` + seleção por id; linha nova/excluída remonta
   preservando; testar-uso: filtro por responsável + ordem por valor.
4. Docs (stack front-vba, PLANO D30), `verificar-tudo -Completo`.

## Checkpoints 🖥

- Duas estações com filtro e ordem; virada da meia-noite com o CRM aberto.

## Depende de

006 (geração da base), 010 (develop).
