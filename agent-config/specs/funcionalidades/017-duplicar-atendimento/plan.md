# Plano: Duplicar atendimento (ciclo 017)

Spec: [SPEC.md](SPEC.md)

## Fatias (teste que falha primeiro, depois prova)

1. **Conector** — testar-mcp `duplicar_atendimento` (falha: ferramenta não existe) → `Get-DadosParaDuplicar`,
   `Invoke-DuplicarAtendimento`; 1.11.0, contrato, nota, matriz, skill.
2. **Front, regra** — testar-uso `DadosParaDuplicar` (falha: não existe) → `modCRM`.
3. **Front, tela** — menu da linha + `frmCRM.PrepararDuplicacao`.
4. **Regra igual** — verificar.py compara as listas; docs (stacks, PLANO D36); `verificar-tudo -Completo`.
