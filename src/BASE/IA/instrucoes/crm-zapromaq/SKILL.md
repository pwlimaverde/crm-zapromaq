---
name: crm-zapromaq
description: Consulta e mantém todo o CRM Comercial Zapromaq (clientes, contatos, atendimentos AT-0000-0000, visitas, planejamentos por região PL-AAAA-NNN, listas e metas) pelas ferramentas do conector crm-zapromaq. Use quando o usuário pedir para ver, procurar, conferir ou alterar dados do CRM, atualizar etapa, próxima ação ou valor de um atendimento, corrigir cadastro de cliente, cadastrar, alterar, transferir ou desativar contato, cadastrar, promover, desativar ou excluir empresa, abrir atendimento, agendar, realizar ou cancelar visita, montar um planejamento por região e o roteiro, manter listas suspensas e metas, ou registrar o resumo de contexto de um cliente ou atendimento.
---

# CRM Comercial Zapromaq

O banco do CRM é um arquivo Access na rede, usado ao mesmo tempo pelos vendedores
pela planilha do front. **Todo acesso é feito pelas ferramentas do conector
`crm-zapromaq`.** Nunca abra, copie ou grave o `.accdb` por outro caminho
(Python, mdbtools, cópia de arquivo): só o conector usa o motor que respeita o
bloqueio do arquivo, e gravar por fora pode corromper o banco de todos.

## Claude e Codex

Esta skill usa o servidor MCP local `crm-zapromaq` no Claude Desktop e no Codex.
Os nomes das ferramentas podem aparecer com um prefixo do aplicativo; use a
ferramenta correspondente do servidor `crm-zapromaq`. Se a conexão não estiver
disponível, informe o problema e solicite a reconexão; não improvise acesso direto
ao banco. O MCP de desenvolvimento `crm-zapromaq-dev` não é necessário para usar o CRM.

## Antes de começar

Chame `crm_status`. Se `gravacao_liberada` vier `false`, o banco está numa versão
de esquema que este conector não conhece: faça só leituras e avise o usuário.
`capacidades` lista as ferramentas desta versão (`versao_conector`): não use ferramenta que não
esteja ali. `alerta_backup` = `true` com `motivos_backup`: avise o usuário (backup atrasado).

Se uma gravação voltar **"Sistema em manutenção"** ou **"Cópia de segurança (backup) do banco
em andamento"**, nada foi gravado: avise o usuário e tente de novo só quando ele pedir (o
backup leva segundos; a manutenção, o tempo da atualização). Não repita sozinho em laço.

## Ler

| Para | Ferramenta |
|---|---|
| Achar uma empresa | `buscar_clientes` (nome, cidade, segmento, CNPJ ou código) |
| Ficha da empresa, contatos e atendimentos | `obter_cliente` (`codigo_cliente`; pré-cliente não tem código: use `id_cliente`) |
| Achar uma pessoa | `buscar_contatos` (nome, cargo, e-mail, telefone, código CT ou empresa) |
| Ficha do contato e atendimentos dele | `obter_contato` com `CT-0000-0000` ou `id_contato` |
| Atendimentos de um vendedor, etapa ou situação | `buscar_atendimentos` |
| Atendimentos com última interação num período (relatório semanal) | `atendimentos_por_interacao` |
| Atendimentos em aberto com próxima ação de hoje em diante (planejamento) | `atendimentos_por_proxima_acao` |
| Ficha de um atendimento | `obter_atendimento` com `AT-0000-0000` |
| Valores aceitos em listas | `listar_opcoes` |
| Metas mensais | `listar_metas` |
| Quem alterou o quê | `ultimas_alteracoes` |
| Extrato da base inteira para análise por script | `exportar_base` |
| Visitas (de um atendimento, de um planejamento ou por situação) | `listar_visitas` |
| Planejamentos por região | `listar_planejamentos`, `obter_planejamento` (id ou `PL-AAAA-NNN`) |
| Atendimentos candidatos a visita perto de uma cidade | `listar_candidatos` |
| Código IBGE de uma cidade | `buscar_municipios`, `resolver_cidade` |

A **situação** (Ação atrasada, Retomar hoje, Ação nesta semana, Em dia, Sem próxima
ação, Encerrado) é calculada com a mesma regra da planilha.

## Relatório semanal — `atendimentos_por_interacao`

Use para montar o relatório semanal da direção ou qualquer pergunta do tipo
"quem foi contatado na semana". Uma única chamada devolve os atendimentos
**abertos e encerrados** cuja `ultima_interacao` cai no período, ordenados por
responsável e data, com um `resumo` por responsável (total, contagem por etapa e
por situação, calculados no conector). Somente leitura.

| Parâmetro | Uso |
|---|---|
| `semana` + `ano` | Semana ISO 8601 (segunda a domingo). Vão juntos. |
| `data_inicio` + `data_fim` | AAAA-MM-DD, inclusive as pontas. Alternativa à semana; os dois grupos juntos são recusados. |
| (nenhum) | Semana ISO imediatamente anterior à data atual (em 28/09/2026 → semana 39/2026, 21/09 a 27/09). |
| `responsavel` | Opcional: filtra um responsável. |
| `completo` | `true` devolve todos os campos, inclusive os técnicos (id, versao, controle…). Padrão `false`. |

Exemplos: `{}` (semana passada) · `{"semana": 39, "ano": 2026, "responsavel": "Paulo"}` ·
`{"data_inicio": "2026-09-01", "data_fim": "2026-09-15"}`.

Saída: JSON compacto, sem campos vazios — `periodo`, `quantidade`, `resumo`,
`atendimentos` (observações e `ctx_resumo` com texto integral). A `situacao` é a
mesma de `buscar_atendimentos`. Se o volume passar do limite da ferramenta, vem
`truncado: true`, `devolvidos` e `quantidade` com o total real: o `resumo` continua
completo, e os atendimentos faltantes se obtêm repetindo a chamada por
`responsavel`. Semana fora de 1 a 53, data inválida ou `data_fim` anterior a
`data_inicio` voltam como erro, sem consulta.

## Planejamento — `atendimentos_por_proxima_acao`

Use para o relatório semanal de planejamento ou qualquer pergunta do tipo
"quem vamos contatar esta semana / neste período". Uma única chamada devolve os
atendimentos **em aberto** (etapa diferente de Pedido Fechado, Perdido e
Descartado; etapa vazia conta como aberta) cuja `dt_prox_acao` cai no período,
com `resumo` por responsável e `equipe` com o total (os dois só na primeira página). Somente leitura.

O período vale **só de hoje em diante**: o CRM sobrescreve a próxima ação quando
ela é executada, então período passado não tem dado. Para o que já aconteceu use
`atendimentos_por_interacao`.

| Parâmetro | Uso |
|---|---|
| `data_inicio` + `data_fim` | AAAA-MM-DD, inclusive as pontas. Vão juntas. |
| `semana` + `ano` | Semana ISO 8601. Alternativa às datas; os dois grupos juntos são recusados. |
| (nenhum) | De hoje até o domingo da semana ISO corrente. |
| `responsavel` | Opcional: filtra um responsável. Sem registro, volta `aviso` para conferir a grafia em `listar_opcoes`. |
| `incluir_vencidas` | Padrão `true`: inclui os abertos com `dt_prox_acao` anterior a hoje, com `vencida: true`. |
| `completo` | `true` devolve todos os campos, inclusive os técnicos. Padrão `false`. |
| `a_partir_de` | Paginação: índice do primeiro registro. Padrão `0`. |

Regras do período, conferidas antes de consultar o banco:

- `data_fim` anterior a hoje, ou semana ISO já encerrada: **recusado**.
- `data_inicio` anterior a hoje com `data_fim` de hoje em diante (ou a semana corrente): começa hoje e volta `periodo.ajustado: true` com `data_inicio_pedida`.
- Mais de 92 dias, data inválida ou `data_fim` anterior a `data_inicio`: recusado.
- Período por datas que coincide com uma semana ISO inteira volta com `semana` e `ano`.

Exemplos: `{}` (resto da semana) · `{"semana": 41, "ano": 2026}` ·
`{"data_inicio": "2026-10-01", "data_fim": "2026-10-31", "responsavel": "Paulo"}`.

Saída: JSON compacto, sem campos vazios — `periodo` (com `hoje`), `quantidade`
(total real, vencidas incluídas), `devolvidos`, `resumo` e `equipe` (só com
`a_partir_de` = 0) e `atendimentos`, ordenados por vencidas primeiro, depois responsável,
`dt_prox_acao`, prioridade (Alta, Média, Baixa, vazio) e empresa. Campos: os de
`atendimentos_por_interacao` mais `quadro`, `categoria`, `codigo_contato` e
`vencida`. Atendimento sem contato vinculado também aparece.

Resumo (por responsável e da equipe): `total` e `por_dia`, `por_etapa`,
`por_quadro` e `valor_funil` (soma de `valor` do quadro "2. Funil comercial")
cobrem só o período; `vencidas` conta à parte; `sem_proxima_acao` conta os
abertos sem `dt_prox_acao` na carteira toda, fora do filtro de período. O resumo
cobre sempre todos os registros e vem só na primeira página; as seguintes trazem
só a lista.

**Paginação:** o limite desta ferramenta é 35.000 caracteres por página (uns 40
atendimentos). Se passar, vem `truncado: true` e `proximo`.
Repita a mesma chamada com `a_partir_de = proximo` até não vir mais `proximo`; a
soma de `devolvidos` fecha com `quantidade`. Não refaça por responsável — um
mês de um único responsável já passa do limite.

## Extrato para análise — `exportar_base`

Use quando a pergunta exige a base inteira (contar, somar, cruzar, filtrar por
data): as consultas devolvem no máximo 200 registros e não agregam. A ferramenta
grava `crm-<tabela>.csv` e `crm-extrato.json` numa pasta e devolve **só o
envelope** (data do extrato, linhas e colunas por tabela). O dado não vem para a
conversa: quem lê e calcula é o script. Não altera o banco.

| Parâmetro | Uso |
|---|---|
| `destino` | Obrigatório. Pasta relativa a `1 - COMERCIAL ZAPROMAQ`, sem letra de unidade; precisa existir. `01 - CRM\01 - CONTROLE` é recusado. |
| `tabelas` | Opcional: `clientes`, `contatos`, `oportunidades`, `log_alteracoes`, `listas`, `metas`, `planejamentos`, `visitas`, `municipios`. Padrão: as três primeiras. |

Formato: UTF-8, separador vírgula, todo campo entre aspas (texto com quebra
de linha fica dentro das aspas) — o formato que os scripts do padrão leem, datas `AAAA-MM-DD` ou `AAAA-MM-DDTHH:MM:SS`,
valor com ponto decimal. `crm-oportunidades.csv` traz `codigo_cliente`,
`empresa`, `codigo_contato` e os calculados `situacao`, `quadro`, `dias_parado`
e `ciclo`, na data do extrato. Cada chamada regrava os arquivos.

## Alterar — sempre nesta ordem

1. **Leia o registro** (`obter_atendimento` / `obter_cliente` / `obter_contato`) e guarde a `versao`.
2. **Mostre ao usuário o que vai mudar** (valor atual → valor novo) e confirme.
3. **Grave** com a `versao` lida, enviando só os campos que mudam, e `autor` com o
   nome de quem pediu.
4. Se voltar **"Conflito"**: alguém alterou o registro depois da leitura. Não repita
   às cegas: leia de novo, mostre a diferença ao usuário e só então refaça.

Regras que o conector confere e recusa com mensagem:

- Campos de lista (`etapa`, `responsavel`, `familia`, `origem`, `prioridade`,
  `motivo_desfecho`, `segmento`, `uf`, `categoria`, `qualificacao`) só aceitam valores de `listar_opcoes`.
- Datas em **AAAA-MM-DD**. Valor em reais como número (`850000.50`).
- Etapa **Pedido Fechado, Perdido ou Descartado** exige `motivo_desfecho` e `dt_desfecho`.
- Etapas de proposta e negociação exigem `familia`.
- Telefone: DDD + número, 10 ou 11 dígitos, **um número por campo**. Nome de quem
  atende e segundo número vão em `observacoes_acrescentar`.
- CNPJ é conferido pelo dígito verificador e não pode repetir outra empresa.
- `observacoes_acrescentar` só acrescenta uma linha datada; nada é apagado.

## Contatos

- **Cadastrar:** `criar_contato`. Antes, `buscar_contatos` para não duplicar a pessoa.
  Nome obrigatório (`GERAL` para o contato geral da empresa). O código `CT-CCCC-NNNN`
  é gerado pelo conector e fica congelado.
- **Pré-cliente:** cadastrar contato promove a empresa a cliente e atribui o próximo
  código. O conector recusa sem `promover_pre_cliente=true`; pergunte ao usuário antes.
- **Alterar:** `atualizar_contato` com a `versao` de `obter_contato` — nome, cargo,
  telefone, e-mail, `observacoes_acrescentar`.
- **Mudou de empresa:** o contato pertence a UMA empresa e não é transferido (ciclo 009): desative-o
  (`ativo=false`) e cadastre um novo na empresa nova com `criar_contato`.
- **Desativar / reativar:** `ativo=false` / `ativo=true`. Contato não se exclui.

## Empresas

- **Cadastrar:** `criar_cliente` — pré-cliente por padrão; `como_cliente=true` já dá código.
  Procure antes com `buscar_clientes`; nome repetido é recusado sem `ignorar_homonimo=true`.
- **Promover:** `promover_cliente` (pré-cliente → cliente, próximo código). Não se desfaz pelo sistema.
- **Desativar / reativar / excluir:** `alterar_situacao_cliente` com a `versao` de `obter_cliente`.
  Reativar a empresa reativa SÓ a empresa (ciclo 009): pergunte quais contatos ainda trabalham lá e
  reative cada um com `atualizar_contato` (`ativo=true`).
  Desativar leva os contatos junto e deixa os atendimentos abertos no funil (avise quantos).
  Só pré-cliente sem contato nem atendimento pode ser excluído; cliente com código nunca.
- Código e estágio não se editam à mão: só mudam pela promoção.

## Atendimentos

- **Abrir:** `abrir_atendimento` com o contato (`codigo_contato`), `etapa` e `responsavel`.
  Mesmas críticas de `atualizar_atendimento`; `dt_entrada` vem hoje se omitida. Retomada:
  `codigo_atendimento_anterior` — só atendimento da mesma empresa do contato (outra empresa é
  recusada, como no front). Pré-cliente exige `promover_pre_cliente=true`.
- **Trocar o contato:** `atualizar_atendimento` com `codigo_contato` — só contato ativo da mesma empresa.
- Atendimento não se exclui nem desativa: encerra-se pela etapa (Pedido Fechado, Perdido, Descartado).

## Visitas

O banco precisa estar no esquema 1.1 (`crm_status`). Um atendimento tem 0 a N visitas; a
situação gravada é PENDENTE, REALIZADA ou CANCELADA (a planilha mostra a pendente sem data como
A AGENDAR e com data como AGENDADA).

- **Criar:** `criar_visita` com o `codigo` do atendimento (PENDENTE). Recusa atendimento
  encerrado (etapa que encerra) e segunda visita pendente no mesmo atendimento. Com
  `id_planejamento`, a data tem de cair no período do planejamento. `tipo` vem da lista
  `TipoVisita`; `resultado`, da lista `ResultadoVisita` (`listar_opcoes`).
- **Realizar ou cancelar:** `alterar_situacao_visita` (ou `atualizar_visita` para outros campos)
  com a `versao` lida em `listar_visitas`. REALIZADA exige `data_realizada` e `relato` (o
  `alterar_situacao_visita` pede também `resultado`) e grava,
  na mesma transação, a linha `VISITA dd/MM/aaaa HH:mm - ...` no histórico do atendimento e a
  última interação dele. CANCELADA exige `motivo_cancelamento`. Visita não se exclui.
- `obter_atendimento` traz o resumo das visitas do atendimento.

## Planejamento por região

1. Ache a cidade-base: `resolver_cidade` (nome + UF) ou `buscar_municipios`; use o código IBGE.
2. `criar_planejamento`: descrição, `cod_municipio`, `raio_km` de 1 a 500 e período
   (`data_inicio` até `data_fim`). O código `PL-AAAA-NNN` é gerado e não muda.
3. `listar_candidatos`: atendimentos em aberto até o raio, na ordem de visita (qualificação, dias
   parado, km); cidades não reconhecidas vêm à parte.
4. `gerar_visitas` com os códigos escolhidos: primeiro com `simular` (padrão `true`), mostre o
   resultado ao usuário e só então repita com `simular=false`. Se um atendimento for recusado,
   nada é gravado. Até 100 por chamada.
5. `exportar_roteiro`: grava o CSV em `Documentos\CRM Zapromaq\Roteiros` (linha BASE e uma por
   visita pendente) e devolve só o caminho e o total. A ordem do roteiro não volta ao sistema.
6. `lancar_datas` e `lancar_resultados` em lote (até 100, cada item com `id` e `versao` lidas;
   `simular` padrão `true`, tudo ou nada).
7. `alterar_situacao_planejamento`: encerrar com visitas pendentes é recusado, salvo
   `cancelar_pendentes=true` (confirme com o usuário: cancela as pendentes com o motivo
   "planejamento encerrado"). Reabrir volta para ABERTO. `atualizar_planejamento` só altera
   planejamento ABERTO.

## Listas e metas

- `gerenciar_lista`: `incluir`, `ativar`, `desativar`, `ordenar` ou `definir_papel` (só na lista
  Etapa: `encerra`, `resultado` GANHO/PERDA, `quadro` FUNIL/PROSPECCAO, `exige_familia`). Nada é
  apagado nem renomeado. Mexer em Etapa ou Responsavel muda o funil de todos: confirme antes.
  Regras: UF só das 27 (C05); resultado só em etapa que encerra (C04); não desativa a última etapa
  ativa de ganho ou de perda (C03). Mudar `encerra` de etapa com atendimentos devolve a contagem e
  pede `confirmar=true` depois de confirmar com o usuário.
- `listar_opcoes` com `detalhar=true`: inclui os inativos, o uso de cada valor e, na Etapa, o papel.
- `gravar_meta`: cria ou altera a meta de um mês (faturamento, pedidos, propostas, contatos).

## O que confirmar sempre com o usuário antes de gravar

Promoção de pré-cliente, desativação ou exclusão de empresa, desativação de contato,
alteração de lista e de meta, gravação em lote (`gerar_visitas`, `lancar_datas`,
`lancar_resultados` com `simular=false`) e encerramento de planejamento com pendentes.
Mostre o valor atual e o novo.

## Contexto do cliente e do atendimento

O texto completo fica nos arquivos `CONTEXTO-GERAL.md` (pasta do cliente) e
`CONTEXTO-ATENDIMENTO.md` (pasta do atendimento). No banco vai só o ponteiro:
`atualizar_contexto_cliente` (`ctx_resumo` de uma linha, `ctx_familia`, `ctx_arquivo`)
e `atualizar_contexto_atendimento`. Esses campos são só da IA e não conflitam com
o vendedor: podem ser gravados sem ler a versão.

## Resultado das gravações

Toda gravação entra no log de auditoria com origem **IA**. Informe ao usuário o
que foi gravado (campos, valor anterior e novo, versão nova).
