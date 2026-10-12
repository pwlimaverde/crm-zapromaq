> Concluída em 12/10/2026 — feature/telefone-contato-oportunidades

# Spec: coluna Telefone na grade de Oportunidades

> Situação: aprovada em 12/10/2026 (decisão do usuário, fechada)

## Objetivo

Quem trabalha a aba Oportunidades liga para o contato do atendimento e hoje precisa abrir a
ficha (ou a aba Contatos) para ver o número. A grade passa a mostrar o telefone ao lado do
contato; quando o contato não tem telefone, mostra o telefone geral da empresa do atendimento.

## Critérios de aceite

- [ ] Aba Oportunidades (mínima e "Ver tudo"): coluna **Telefone** logo à direita de **Contato**.
- [ ] Contato com telefone: a célula mostra o telefone do contato, formatado por
      `modValidacao.FormatarTelefone`.
- [ ] Contato sem telefone: mostra o telefone geral da empresa (`clientes.telefone` do cliente
      do atendimento), formatado.
- [ ] Nenhum dos dois: célula vazia.
- [ ] O feed (`modGrade.AtualizarRegistros`) reescreve a coluna nova; `_b`, `_id`, `_a` seguem
      depois da última coluna visível.
- [ ] Ficha, abas Clientes/Contatos e `frmVinculo` sem mudança.

## Stacks tocadas

front-vba (grade e consulta da aba). banco-access: nenhuma mudança de esquema — a coluna
`clientes.telefone` já existe; só entra no SELECT de `SQLOportunidades` (alias
`telefone_empresa`). Nulo tratado no código (sem `NZ`/`IIf` no SQL).

## Estrutura

- `fonte/modulos/modCRM.bas`: `cl.telefone AS telefone_empresa` no SELECT; função pura
  `TelefoneDoAtendimento(doContato, daEmpresa)`.
- `fonte/modulos/modSchema.bas`: `ColunasPlanilha("oportunidades")` ganha a coluna calculada
  `telefone_atendimento` depois de `contato`.
- `fonte/modulos/modGrade.bas`: `CelulaDaColuna` calcula `telefone_atendimento`.
- `ferramentas/verificacao/verificar.py` e `IA/mcp-dev/lib/Conferencia.ps1` (conferência do
  MCP de desenvolvimento): `telefone_atendimento` na lista de calculados.
- `fonte/modulos/modAutoteste.bas`, `testes/testar-uso.ps1`, `testes/criar-banco-demo.ps1`
  (contatos/empresa sem telefone para cobrir os casos).

## Estratégia de teste

- `Confere` (modAutoteste): posição da coluna; regra do telefone (contato, empresa, nenhum).
- `testar-uso.ps1`: a grade montada tem Telefone depois de Contato e cada linha bate com o
  banco (contato › empresa › vazio), com ao menos uma linha de cada caso no banco demo.
- 🖥 na estação: largura da coluna em 1366×768, AutoFiltro, atualização pelo feed.

## Limites

- Nunca: migração, mudança de esquema, nomes de controle, B7, ferramentas MCP.
- Perguntar antes: mexer na ficha ou nas abas Clientes/Contatos.
