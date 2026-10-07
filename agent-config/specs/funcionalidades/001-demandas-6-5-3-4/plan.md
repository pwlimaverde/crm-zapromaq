# Plano: demandas 6, 5, 3 e 4

Spec: [SPEC.md](SPEC.md) · Branch: `feature/demandas-6-5-3-4`

Cada fatia: teste que falha (RED) → código mínimo (GREEN) → `verificar-tudo -Rapido` (+ o teste
da stack) → commit. Teste de `modAutoteste`/`testar-uso` só roda no build: o RED é visto com
`build\montar-frontend.ps1 -Teste`.

## Ordem e dependências

| # | Fatia | Depende de | Teste que falha primeiro | 🖥 |
|---|---|---|---|---|
| 6.1 | Ficha guarda o vínculo **por campo** (dicionário campo → id; `DefinirValor` "K" mantém em sincronia) no lugar de `mVinculoID`; `modSchema.CamposK` | — | `Confere CamposK("oportunidades") = "id_contato"` (e contatos/clientes) | — |
| 6.2 | `modCRM.DadosDoContato` (uma consulta, Dictionary) e `CriticarTrocaContato` (sem banco); `ResumoContato` reaproveita | 6.1 | `Confere` da troca entre empresas; caso no `testar-uso` | — |
| 6.3 | `frmVinculo.EscolherContatoDe(idCliente)`: passo 2 direto, título "Trocar contato", sem Voltar | 6.2 | contrato de compilação no autoteste | 🖥 |
| 6.4 | Botões pequenos por campo K criados em execução (`clsUI.LigarBotao`, clique no texto repassado a `UI_Acao`); "Procurar" na edição chama `EscolherContatoDe`; cargo/telefone/e-mail atualizados; checagem ao salvar; regra `BotaoVinculoAtivo` | 6.3 | `Confere BotaoVinculoAtivo` + contrato `clsUI.LigarBotao` | 🖥 |
| — | **ponto de verificação:** `-Completo` | 6.4 | | |
| 5.1 | `modLote`: `ComecaIgnorada`, `PodeAlternar`, `RotuloLinha`; `Conferir` e `AlternarIgnorar`; cabeçalho com a reversão | — | `Confere` dos estados e rótulos | — |
| 5.2 | Aceite de ponta a ponta com lote fictício (OK, SUSPEITO, DUPLICADO, ERRO) | 5.1 | `testar-uso` (RED provado rodando-o contra o `modLote` de `develop`) | — |
| 5.3 | `frmLote`: rótulo efetivo, ajuda e mensagem novas (textos em `modLote`); cor do selo pelo prefixo antes de " - " (`clsGrade.ChaveSelo`) | 5.1 | `Confere` dos textos e de `ChaveSelo` | 🖥 |
| 3.1 | `SQLClientes`/`SQLContatos` com parâmetro `situacao` ("ativos", "inativos", "todos") | — | `Confere InStr(SQL, "cl.ativo=True")` etc. | — |
| 3.2 | `criar-banco-demo.ps1`: 2 clientes e seus 2 contatos inativos | — | parte do RED do 3.3 | — |
| 3.3 | `modGrade`: coluna oculta `_a` depois de `_id` (também em `AtualizarRegistros`); célula de situação com lista; AutoFiltro; regra vermelha; Limpar filtros → Ativos; `GuardarFiltros` ignora colunas de controle; `ThisWorkbook` encaminha a célula | 3.1, 3.2 | `testar-uso`: visíveis < total ao abrir; Todos = total; regra condicional; Limpar volta a Ativos | 🖥 |
| 3.4 | Ficha: `cboSituacao` em clientes e contatos (posição em `Geometria`), `modCRM.OpcoesSituacao`/`ChaveSituacao` | 3.1 | `Confere ChaveSituacao`/`OpcoesSituacao` | 🖥 |
| — | **ponto de verificação:** `-Completo` | 3.4 | | |
| 4.1 | Campo `id_atendimento_anterior|Atendimento anterior|K||1. Cliente e contato|0|1`; `SQLOportunidades` com `op.id_atendimento_anterior` e `ant.codigo AS codigo_anterior` (LEFT JOIN entre parênteses) | 6.1 | `verificar.py` (campo da ficha fora da consulta) + `Confere CamposK = "id_contato,id_atendimento_anterior"`; `conferir-sql-gerado.py` | — |
| 4.2 | `modCRM.CriticarAnterior` (sem banco) e `ListarAtendimentosDoCliente(idCli, exceto, texto)` | 4.1 | `Confere` (próprio, outra empresa, vazio); caso no `testar-uso` | — |
| 4.3 | `frmVinculo.EscolherAtendimentoDe`; botão do campo; resumo com o código AT-; botão "abrir" abre a ficha em leitura; remover vínculo; checagem ao salvar | 4.2, 6.4 | contrato de compilação + `Confere BotaoVinculoAtivo` do anterior | 🖥 |
| 4.4 | IA: `abrir_atendimento` recusa anterior de outra empresa; skill da IA atualizada | — | caso no `testar-mcp.ps1` | — |
| F | Fechamento: specs de stack (`front-vba`, `ia-mcp`, `banco-access`, `build-operacao-powershell`), `PLANO.md` (D17–D19, 4c), `docs/contrato-banco.md`, carimbo | todas | `verificar-tudo -Completo` | — |

## Checkpoints

- Depois do 6.4 e do 3.4: `verificar-tudo -Completo`.
- Estágio 4: `-Completo` uma vez, personas `code-reviewer`, `security-auditor`, `test-engineer`
  e `code-simplification`.

## Itens 🖥 (Excel 2019, na estação)

1. 6.3/6.4: na edição de um atendimento, o botão ao lado de Contato abre a lista da empresa sem
   Voltar; a escolha atualiza cargo/telefone/e-mail; salvar e reabrir mantém; o botão responde
   ao clique no ícone e no fundo.
2. 6.4: "Procurar" na edição não mostra outras empresas; no atendimento novo continua com 2 passos.
3. 5.3: lote com as 4 situações; total inicial só OK; duplo clique em OK tira, em SUSPEITO
   inclui; DUPLICADO/ERRO mostram a mensagem nova; cores dos rótulos.
4. 3.3: aba abre em Ativos; Todos mostra inativos em vermelho; Limpar filtros volta a Ativos; a
   lista da célula funciona com a aba protegida; largura da célula em Clientes e Contatos,
   inclusive no "Ver tudo".
5. 3.3 + feed: desativar um cliente em outra estação com a grade aberta; a linha fica vermelha
   sem quebrar a grade.
6. 3.4: filtro de situação na ficha de clientes e contatos sem corte em 1366×768 (100% e 125%).
7. 4.3: anterior num atendimento novo para um perdido da mesma empresa, salvar e reabrir; botão
   abre a ficha do anterior em leitura; remover funciona; o próprio não aparece na lista.
8. Antes do `/ship`: avisar a equipe que o SUSPEITO agora começa fora.

## Riscos

| Risco | Impacto | Como se percebe / mitiga |
|---|---|---|
| Segundo campo K sem o 6.1 gravaria o id do contato no anterior | Alto | 6.1 antes de tudo; `Confere CamposK` |
| Botão criado em execução: o clique no texto não chega ao formulário | Médio | `clsBotao` repassa por `UI_Acao`; 🖥 |
| Coluna `_a` nova e filtro reaplicado por índice no "Ver tudo" | Médio | `_a` depois de `_id`; `GuardarFiltros` ignora colunas de controle; teste "Ver tudo" existente |
| Regra condicional depende do idioma do Excel | Médio | fórmula simples sem função (`=$X8="INATIVO"`); 🖥 |
| Equipe acostumada ao SUSPEITO entrando | Médio | aviso no relatório e antes do `/ship` |
