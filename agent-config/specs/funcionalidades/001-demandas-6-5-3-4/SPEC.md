# Spec: demandas 6, 5, 3 e 4 do comercial

> Concluída em 07/10/2026 — feature/demandas-6-5-3-4 (aprovada em 07/10/2026, GATE 1)

Origem: documento de demandas de 29/09/2026 (relatos do comercial de 18 a 29/09/2026), itens
6, 5, 3 e 4, nesta ordem. Itens 1, 2 e 7 ficam fora deste ciclo.

## Objetivo

- **6 — Trocar o contato do atendimento sem refazer a busca da empresa.** Na edição de um
  atendimento, o vendedor troca o contato por um botão ao lado do campo, que abre direto a lista
  de contatos da mesma empresa. Hoje é preciso procurar a empresa de novo.
- **5 — Qualificação em lote.** Duplo clique alterna OK e SUSPEITO entre entrar e ficar fora;
  SUSPEITO começa fora. Reverte a decisão do comercial de 17/09/2026.
- **3 — Inativos em vermelho e filtro Ativos/Inativos/Todos** nas abas Clientes e Contatos e na
  ficha; abre em Ativos.
- **4 — Atendimento anterior referenciado** no atendimento (retomada de perdido/descartado).

## Decisões (Q1–Q9 do plano entregue)

| # | Decisão |
|---|---|
| Q1 | BASE da rede em dia (versão 1.5, sem ajuste na empresa desde 04/10). |
| Q2 | O item 4 usa a coluna **existente** `oportunidades.id_atendimento_anterior` (esquema 1.0, já gravada pela IA). Sem migração 004; esquema continua 1.0. |
| Q3 | O atendimento **não troca de empresa** (código AT- congelado; a IA já recusa). A troca abre o passo 2 do `frmVinculo` **sem Voltar**. Defeito corrigido junto: o "Procurar" na edição deixava escolher contato de outra empresa e o `id_cliente` ficava inconsistente; passa a haver checagem ao salvar (`modCRM.CriticarTrocaContato`). |
| Q4 | O botão "Procurar" (`cmdVinculo`) continua: na edição de atendimento gravado faz o mesmo que o botão novo; no atendimento novo mantém os dois passos. |
| Q5 | Filtro sem controle novo no layout: na aba, **célula com lista** (validação de dados) na linha da pesquisa, à direita da caixa; coluna oculta `_a`; na ficha, `lblSituacao`/`cboSituacao` reposicionados em `Geometria`. |
| Q6 | Vermelho = cor `perigo` que já existe no tema (`modTema.COR_PERIGO`). Painel e relatórios fora. |
| Q7 | Reversão do SUSPEITO em 06/10/2026: "Relatos do comercial de 18 a 29/09/2026: o suspeito que entrava sozinho virava cadastro repetido que ninguém revia. Passa a começar fora e só entra por escolha consciente (duplo clique)." Rótulos: `OK`, `OK - IGNORAR`, `SUSPEITO - FORA`, `SUSPEITO - ENTRA`. |
| Q8 | Lista do anterior: todos os atendimentos da mesma empresa, menos o próprio, mais recentes primeiro (entrada desc., nulo no fim, id desc.); vale no novo e na edição; dá para remover o vínculo. |
| Q9 | A IA (`abrir_atendimento`) passa a recusar anterior de outra empresa. Nome, parâmetros e esquema da ferramenta não mudam. |

## Critérios de aceite

Conferidos aqui (autoteste, teste de uso, testes do MCP, `verificar-tudo -Completo`); o que é
tela fica nos itens 🖥 do [plan.md](plan.md), **pendentes na estação com Excel 2019**.

- [x] **6:** na edição de um atendimento, o botão ao lado do campo Contato abre direto a lista de
      contatos da empresa (sem Voltar); a escolha atualiza o vínculo e cargo, telefone e e-mail
      somente leitura; salvar e reabrir mantém o contato. Contato de outra empresa é recusado ao
      salvar. Atendimento novo mantém os dois passos.
- [x] **5:** num lote com linhas OK, SUSPEITO, DUPLICADO e ERRO, o total inicial "vai cadastrar"
      conta só as OK; duplo clique em OK tira a linha e em SUSPEITO inclui, com rótulo e total
      atualizados; DUPLICADO e ERRO mostram "Só OK e SUSPEITO alternam".
- [x] **3:** as abas Clientes e Contatos abrem só com ativos; em Todos os inativos aparecem em
      vermelho; "Limpar filtros" volta para Ativos; o filtro sobrevive ao Atualizar e à
      atualização automática; a ficha de clientes e contatos tem o mesmo filtro (abre em Ativos).
- [x] **4:** um atendimento (novo ou em edição) aponta para outro da mesma empresa; o vínculo
      aparece ao reabrir (código AT-); um botão abre a ficha do anterior em leitura; dá para
      remover; apontar para si mesmo ou para outra empresa é recusado (front e IA).
- [x] Nada de dado real; esquema 1.0; nenhum contrato tocado (B7, `VERSAO-FRONT`, nomes de
      controle do layout, ferramentas MCP, códigos `CT-`/`AT-`).

## Stacks tocadas

- **front-vba:** `frmCRM`, `frmVinculo`, `frmLote`, `modSchema`, `modCRM`, `modLote`, `modGrade`,
  `clsUI`, `clsBotao`, `clsGrade`, `ThisWorkbook`, `modAutoteste`.
- **banco-access:** só SQL (parâmetros, `TOP` com desempate, JOIN aninhado). Sem DDL nova.
- **build-operacao-powershell:** casos novos em `testes/testar-uso.ps1`; `criar-banco-demo.ps1`
  com 2 clientes e 2 contatos inativos.
- **ia-mcp:** regra do anterior em `abrir_atendimento` + caso em `IA/teste/testar-mcp.ps1`.
- **visual-layout:** nenhum JSON de layout ou tema muda (botões do vínculo e posição do filtro
  são de tempo de execução).

## Estrutura

Ver a tabela de fatias em [plan.md](plan.md).

## Estratégia de teste

Tabela "TDD por stack" de `specs/PROJETO.md`. Para os formulários (modal, não testáveis sem
interação) o teste que falha primeiro é um **contrato de compilação** no `modAutoteste`
(`If False Then f.Metodo ...` com variável tipada: método inexistente = erro de compilação no
build) mais uma regra pura em `Confere`; o comportamento visual vira item 🖥.

## Limites

- Sempre: SQL parametrizado; data como parâmetro; versão otimista e log na mesma transação
  (`modCRM.Atualizar`/`InserirEm` já fazem); uma consulta por ação do usuário.
- Perguntar antes: qualquer mudança de esquema, de nome de controle do layout ou de ferramenta MCP.
- Nunca: migração 004; tocar o banco de produção; `MONTAR-FRONTEND.bat` sem `-Teste`.

## Questões em aberto

Nenhuma. Aviso para o `/ship`: a equipe precisa saber que o SUSPEITO agora começa fora.
