# Tarefas: coluna Telefone na grade de Oportunidades

- [x] Tarefa: testes que falham (posição da coluna, regra do telefone, grade montada x banco)
  - Aceite: falham pelo motivo esperado — `testar-uso`: "Telefone logo depois de Contato" (coluna
    inexistente); `verificar.py`: `modCRM.TelefoneDoAtendimento nao encontrado`
  - Teste primeiro: Confere em modAutoteste + caso em testar-uso.ps1
  - Arquivos: modAutoteste.bas, testar-uso.ps1, criar-banco-demo.ps1
- [x] Tarefa: coluna Telefone (contato › empresa › vazio) na grade e no feed
  - Aceite: critérios da SPEC; autoteste OK 108; testar-uso: contato (18), empresa (1), vazio (1)
  - Verificar: `verificar-tudo -Rapido`, depois `-Completo`
  - Arquivos: modCRM.bas, modSchema.bas, modGrade.bas, verificar.py, IA/mcp-dev/lib/Conferencia.ps1
- [x] Tarefa: spec de stack (front-vba) e PLANO.md (D20)
- [ ] 🖥 Na estação: largura da coluna em 1366×768, AutoFiltro na coluna, atualização pelo feed
