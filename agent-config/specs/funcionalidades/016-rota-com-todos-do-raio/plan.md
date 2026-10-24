# Plano: Rota com todos os clientes do raio (ciclo 016)

Spec: [SPEC.md](SPEC.md)

## Fatias (teste que falha primeiro, depois prova)

1. **Conector** — testar-mcp `sem_atendimento` (falha: campo não existe) → `Get-CandidatosPlanejamento`; 1.10.0, nota,
   matriz, skill.
2. **Front, regra** — testar-uso `EmpresasSemAtendimento`, `GerarVisitas` com id 0, `AnteriorParaRetomada`,
   `ContatoParaAtalho` (falha: não existem) → `modPlanejamento`.
3. **Front, tela** — linhas "(sem)" na página Candidatos, marcar recusado, botão `cmdCriarAtendimento`
   (layout + `frmCRM.DefinirAnteriorInicial`), volta marcado.
4. Docs (stacks, PLANO D35) e `verificar-tudo -Completo`.

## Checkpoints 🖥

- frmPlanejamento na estação: "(sem)", atalho, volta marcado.
