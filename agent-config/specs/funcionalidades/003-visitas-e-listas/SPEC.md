# Spec: visitas, planejamentos, listas padrão e conector completo

> Situação: revisada após GATE 1 (versão final para aprovação)

## Objetivo

1. **Visitas ligadas ao atendimento.** Uma oportunidade tem várias visitas ao longo do tempo,
   consultadas e editadas pelo botão **Visitas (n)** da ficha do atendimento.
2. **Planejamento por região**, em formulário próprio: cidade-base + raio → atendimentos em aberto
   dentro do raio → gerar visitas → exportar CSV para montar o roteiro fora do sistema → lançar
   data e hora → lançar resultados → encerrar. A ordem do roteiro não volta ao sistema.
3. **Listas editáveis e papel das etapas.** No Início, em **Gerenciamento**: incluir,
   desativar/reativar, ordenar e editar o papel das etapas (encerra, resultado, quadro, exige
   família). Não existe renomear.
4. **Claude com paridade.** Tudo o que o vendedor faz, o Claude faz pelo conector
   `crm-zapromaq`, com as mesmas regras.
5. **Montagem única e independente.** O usuário copia a pasta da base para a empresa, roda
   `SISTEMA\MONTAR-FRONTEND.bat` e, no fim, banco e CRM estão atualizados. Qualquer falha: nada muda.

## Decisões

- **D-a:** o CSV exportado e as respostas do conector mantêm nome e telefone do contato.
- **D-b:** todo teste de banco usa o banco de teste da pasta (criado por
  `IA\teste\criar-banco-teste.ps1`, com `config.ambiente = TESTE`).
- **D-c:** lista é texto na tabela única `listas`; o papel das etapas fica em colunas; o programa
  não compara mais texto fixo de etapa.
- **D-d:** os dados do IBGE são preparados nesta estação e versionados; a montagem só os consome.
- **D-e:** BASE da rede verificada; sem ajustes na empresa desde 04/10.
- **D-f:** a montagem confere a trava (`.laccdb`) antes de qualquer alteração; se houver, para e
  mostra as estações. Nunca apaga a trava.
- **D-g (tudo ou nada):** a migração roda numa cópia do banco; o front é montado e testado contra
  a cópia; o banco de produção só é trocado no fim, se ninguém estiver usando e o hash dele for o
  mesmo do backup.
- **D-h:** sem renomear. Troca de vendedor: pelo Claude, um registro por vez
  (`buscar_atendimentos` → `atualizar_atendimento`; `atualizar_cadastro_cliente`;
  `atualizar_visita`; `atualizar_planejamento`), com placar e tratamento de conflito na
  `SKILL.md`; depois, `gerenciar_lista` desativa o valor antigo.
- **D-i (duplicatas, variantes e órfãos):** comparação de lista sempre pela **chave** (maiúsculas,
  sem acento, espaços aparados e colapsados). Na migração 003:
  1. Grupos com a mesma (tipo, chave): fica a variante com mais registros usando; empate → ativa;
     empate → menor `id`. Os registros com as outras variantes passam a ter o texto da que ficou,
     com versão +1 e log por registro (ação `UNIFICACAO`, usuário `MIGRACAO 003`). As linhas
     duplicadas são removidas com log (ação `UNIFICACAO`, valor antigo registrado).
  2. Valor usado em registro sem linha na lista (pela chave): entra como **inativo**, com o texto
     mais usado; as outras variantes são unificadas como no item 1.
  3. Etapas especiais ausentes entram inativas, com o papel da tabela "Papel inicial".
  4. Tudo vai para o relatório `SISTEMA\DADOS\execucao\logs\migracao-003-AAAAMMDD-HHMMSS.txt` e
     aparece na mensagem final da montagem.
- **D-j:** a ficha aceita valor ativo ou o próprio valor já gravado. Valor inativo gravado aparece
  como `<valor> (inativo)` e é gravado sem o sufixo. Vale no front, no cadastro em lote e no conector.
- **D-k:** formulário próprio `frmPlanejamento`, sem depender da ficha genérica.

## Contratos

### Tabelas (esquema 1.1)

- `municipios`: `cod_ibge` LONG PK, `nome` TEXT(80), `nome_busca` TEXT(80), `uf` TEXT(2),
  `lat` DOUBLE, `lon` DOUBLE, `fonte_carga` TEXT(60); índice (`uf`, `nome_busca`).
- `planejamentos`: `id` AUTOINCREMENT, `codigo` TEXT(12) único (`PL-AAAA-NNN`), `controle` LONG,
  `descricao` TEXT(120), `cod_municipio_base` LONG, `raio_km` INTEGER, `data_inicio` DATETIME,
  `data_fim` DATETIME, `responsavel` TEXT(60), `situacao` TEXT(20) (`ABERTO` | `ENCERRADO`),
  `observacoes` MEMO, `versao` LONG, `criado_em`, `alterado_em` DATETIME, `alterado_por` TEXT(50).
- `visitas`: `id` AUTOINCREMENT, `id_oportunidade` LONG obrigatório, `id_planejamento` LONG,
  `situacao` TEXT(20) (`PENDENTE` | `REALIZADA` | `CANCELADA`), `data_agendada` DATETIME,
  `hora_agendada` TEXT(20), `responsavel` TEXT(60), `tipo` TEXT(40), `observacoes` MEMO,
  `data_realizada` DATETIME, `resultado` TEXT(40), `relato` MEMO, `motivo_cancelamento` TEXT(120),
  `versao`, `criado_em`, `alterado_em`, `alterado_por`. Índices: (`id_oportunidade`, `situacao`),
  `id_planejamento`, `data_agendada`. Situação exibida: `A AGENDAR` (PENDENTE sem data),
  `AGENDADA` (PENDENTE com data), `REALIZADA`, `CANCELADA`.
- `listas` (colunas novas): `chave` TEXT(120) obrigatória, `encerra` YESNO, `resultado` TEXT(10)
  (`GANHO` | `PERDA` | nulo), `quadro` TEXT(12) (`PROSPECCAO` | `FUNIL` | nulo), `exige_familia`
  YESNO. Índice único (`tipo`, `chave`). YESNO gravado sempre explicitamente.
- Listas novas: `TipoVisita`, `ResultadoVisita`.

### Papel inicial das etapas (gravado pela migração 003)

| Etapa | encerra | resultado | quadro | exige_familia |
|---|---|---|---|---|
| Pedido Fechado | Sim | GANHO | FUNIL | Não |
| Perdido | Sim | PERDA | FUNIL | Não |
| Descartado | Sim | PERDA | PROSPECCAO | Não |
| Contato Inicial, Sem Retorno, Retorno Agendado | Não | — | PROSPECCAO | Não |
| Elaboração da Proposta, Negociação, Proposta em Análise, Proposta em Stand By | Não | — | FUNIL | Sim |
| Qualquer outra etapa | Não | — | FUNIL | Não |

### Arquivos de referência (dentro da base)

- `SISTEMA\DADOS\banco\referencia\municipios.csv` — UTF-8 com BOM, `;`, colunas
  `cod_ibge;nome;nome_busca;uf;lat;lon`, ponto decimal.
- `SISTEMA\DADOS\banco\referencia\municipios.fonte.json` — URLs, data do download, SHA-256 das
  fontes e do CSV, total, contagem por UF, divergências resolvidas, licenças (IBGE: dado público;
  coordenadas `kelvins/municipios-brasileiros`: MIT).
- `SISTEMA\DADOS\banco\migracoes\003-municipios-visitas-papel-listas.ps1` — `EsquemaDe='1.0'`,
  `EsquemaPara='1.1'`, SHA-256 do CSV como constante.
- `SISTEMA\DADOS\MANIFESTO.json` — gerado pelo empacotamento: lista e SHA-256 de todos os
  arquivos da base.

### Telas

- `frmListas` (Gerenciamento): combo da lista; grade com Valor, Ordem, Ativo, Em uso e, em Etapa,
  Encerra, Resultado, Quadro, Exige família; botões Incluir, Desativar/Reativar, Subir, Descer,
  Editar papel, Fechar.
- `frmPlanejamento`: cabeçalho (seletor de planejamento, descrição, cidade-base com busca, raio,
  início, fim, responsável, situação, observações; Novo, Editar, Salvar, Cancelar,
  Encerrar/Reabrir) e páginas (MultiPage) **Visitas** (grade, filtro de data, Marcar todos,
  Exportar, Lançar), **Candidatos** (grade com marcação, Gerar visitas) e **Lançamento** (visita
  atual com Data, Hora, Situação, Tipo, Resultado, Relato, Motivo; Anterior, Salvar e próxima,
  Fechar).
- `frmVinculo`: modo novo `municipio`.
- Ficha da visita: `frmCRM` com a tabela `visitas`, aberta pelo botão `cmdAtendimento` (legenda
  `Visitas (n)`) da ficha do atendimento.
- Aba `Planejamentos`: seletor na célula `C4` (validação de lista), grade de visitas só de leitura,
  botões Abrir planejamento, Atualizar, Início; duplo clique numa visita abre `frmPlanejamento`
  na página Lançamento, naquela visita.
- Início: botões `Gerenciamento` (macro `AbrirGerenciamento`) e `Planejamentos` (macro
  `GradePlanejamentos`).

### Conector (25 → 41)

Novas: `listar_planejamentos`, `obter_planejamento`, `criar_planejamento`,
`atualizar_planejamento`, `alterar_situacao_planejamento`, `listar_candidatos`, `gerar_visitas`,
`listar_visitas`, `criar_visita`, `atualizar_visita`, `lancar_datas`, `lancar_resultados`,
`alterar_situacao_visita`, `exportar_roteiro`, `buscar_municipios`, `resolver_cidade`.
Alteradas: `gerenciar_lista` (ações `incluir`, `ativar`, `desativar`, `ordenar`,
`definir_papel`), `listar_opcoes` (`detalhar`), `obter_atendimento` (resumo de visitas),
`crm_status` (esquemas), `exportar_base` (`planejamentos`, `visitas`, `municipios`).
Lotes (`gerar_visitas`, `lancar_datas`, `lancar_resultados`): `simular` padrão `true`, até 100 itens.

### Versão e conflito por tipo de gravação

| Gravação | Controle |
|---|---|
| `visitas`, `planejamentos` | `WHERE id=? AND versao=?`; 0 linhas = conflito |
| V6/V7 no atendimento | versão do atendimento na mesma transação; conflito desfaz a visita também |
| `listas` | `WHERE id=?` + valores antigos de `ativo`, `ordem`, `encerra`, `resultado`, `quadro`, `exige_familia` |
| Gerar visitas / lotes | uma transação; V8 conferida de novo dentro dela; qualquer conflito desfaz tudo |
| Código `PL-` | `MAX(controle)+1` na transação + índice único; colisão → nova tentativa (até 3) |
| Encerrar planejamento | versão de cada visita cancelada |

## Montagem única (`SISTEMA\MONTAR-FRONTEND.bat`)

O banco é sempre `BASE\crm_zapromaq.accdb` (ao lado de `SISTEMA\`). Passos, nesta ordem; a tela
mostra `[n/12] <passo>... OK` ou `PAROU NO PASSO n/12 (<passo>): <motivo>`.

1. **Pacote:** confere `MANIFESTO.json` (arquivos presentes e SHA-256), o CSV contra o
   `.fonte.json` e contra a constante da migração.
2. **Ambiente:** Excel 2019 64 bits; `CRM_Zapromaq.xlsm` da raiz livre.
3. **Trava:** `crm_zapromaq.laccdb` existe → para e mostra as estações da trava.
4. **Esquema:** lê `config.versao_esquema` e lista as migrações pendentes.
5. **Backup:** banco e `.xlsm` em `SISTEMA\DADOS\execucao\backup\antes-da-publicacao\AAAA-MM-DD_HHMMSS\`;
   guarda o SHA-256 do banco (H0).
6. **Cópia de trabalho:** `SISTEMA\DADOS\execucao\publicacao\crm_zapromaq.accdb`.
7. **Migrações na cópia** (dados em transação; estrutura sem transação; erro → a cópia é descartada).
8. **Conferências na cópia:** carga dos municípios, papel das etapas, índice único de `listas`,
   nenhum valor em uso sem linha na lista, esquema 1.1.
9. **Montar o front:** `Config!B2` provisório = cópia; compilação e autoteste.
10. **Teste de uso** contra a cópia.
11. **Troca do banco:** grava `Config!B2` definitivo (UNC do banco de produção) e confere no
    pacote; confere de novo a trava e se o SHA-256 do banco = H0; troca o arquivo de produção pela
    cópia (`[IO.File]::Replace`, com o banco anterior guardado no backup). Hash diferente → para:
    "alguém gravou durante a montagem; rode de novo". Nada foi trocado.
12. **Publicar:** copia o `.xlsm`, grava `config.versao_front`, `VERSAO.txt` e `VERSAO-FRONT.txt`;
    só então a versão sobe. Falha aqui → o banco volta ao backup, se o hash dele não mudou desde
    a troca; se mudou, para e avisa (o CRM antigo não grava no 1.1).

Mensagem final: `PRONTO. CRM <versão> publicado; banco no esquema 1.1. Municípios: <n>. Listas:
<u> unificações, <o> valores incluídos como inativos (relatório: <arquivo>). Backup: <pasta>.
Usuários do Claude Desktop: fechem e abram o aplicativo e reinstalem a skill
IA\instrucoes\crm-zapromaq.zip.`

`-Teste`: os mesmos 12 passos sobre o banco da pasta, só se `config.ambiente = TESTE` (senão
recusa); o passo 12 grava em `SISTEMA\DADOS\execucao\teste\` e não sobe a versão.

### Pasta levada para a empresa

Gerada por `agent-config\ferramentas\empacotar-base.ps1` em `dist\BASE-v<X.Y>\BASE\` (e no `.zip`),
a partir do commit. **Fica de fora:** `crm_zapromaq.accdb`, `*.laccdb`, `CRM_Zapromaq.xlsm`,
`VERSAO-FRONT.txt`, `SISTEMA\DADOS\execucao\**` (só os `.gitkeep`), `IA\logs\`, `*.bak-*`,
`SISTEMA\DADOS\docs\CHANGELOG.md`. **Entra:** `INICIAR-CRM.bat`, `SISTEMA\` (`MONTAR-FRONTEND.bat`,
`DADOS\` com `VERSAO.txt`, `MANIFESTO.json`, `fonte\`, `build\`, `banco\` com `esquema\`,
`migracoes\`, `referencia\`, `lib\`, `operacao\`, `testes\`, `ferramentas\`, `docs\`), `IA\`
(com `instrucoes\*.zip` gerados). O usuário copia por cima da base da empresa.

## Casos de regra (mesmo número no teste do front e no do conector)

C01 chave: "pedido  fechado ", "PEDIDO FECHADO" e "Pedído Fechado" têm a mesma chave.
C02 incluir valor com chave existente → recusa. C03 desativar a última etapa ativa GANHO (ou
PERDA) → recusa. C04 `resultado` com `encerra` = Não → recusa. C05 UF fora das 27 → recusa.
C06 gravar valor inativo já gravado → aceita; outro inativo → recusa. C07 etapa com encerra = Sim →
situação "Encerrado" e fora da fila. C08 V2. C09 V3 data fora do período. C10 V4 sem relato.
C11 V5 sem motivo. C12 V6 texto `VISITA 14/10/2026 09:30`. C13 V8 segunda pendente. C14 P1 raio
0 e 501. C15 P2. C16 P3 sem `cancelar_pendentes`. C17 unificação escolhe a variante mais usada.
C18 órfão entra inativo. C19 trava presente → montagem para no passo 3. C20 hash do banco mudou →
para no passo 11.

## Regras de visita (V1–V9)

Decisões registradas na Fase 3 para fechar os pontos que a lista de casos deixava em aberto.

- **V1** Cada visita pertence a exatamente um atendimento (`id_oportunidade` obrigatório); um
  atendimento tem 0 a N visitas.
- **V2 (C08)** Não se cria visita para atendimento encerrado (etapa com `encerra` = Sim).
- **V3 (C09)** Com planejamento, `data_agendada` tem de estar entre `data_inicio` e `data_fim`
  do planejamento. Sem planejamento, qualquer data.
- **V4 (C10)** Situação REALIZADA exige `data_realizada` e `relato` preenchidos.
- **V5 (C11)** Situação CANCELADA exige `motivo_cancelamento`.
- **V6 (C12)** Ao registrar REALIZADA, o histórico do atendimento (`oportunidades.observacoes`)
  ganha a linha `VISITA dd/MM/aaaa HH:mm - <tipo> - <resultado>: <relato resumido>`, na mesma
  transação da visita.
- **V7** Ao registrar REALIZADA, `oportunidades.ultima_interacao` recebe `data_realizada`
  (mesma transação, mesma versão do atendimento).
- **V8 (C13)** Um atendimento tem no máximo uma visita PENDENTE por vez: a segunda é recusada.
- **V9** `Visitas (n)` conta todas as visitas do atendimento, em qualquer situação; a lista
  mostra as pendentes primeiro.

## Critérios de aceite

### Fase 1 — Base, migração 003 e montagem única (publicação A)

1. `municipios.csv` e `.fonte.json` existem; o CSV passa nas checagens (total oficial, código de 7
   dígitos único, 27 UFs, lat −34..6, lon −74..−28) e o hash bate.
2. Função de chave idêntica no front e no conector (C01).
3. Migração 003 numa cópia do banco de teste: estrutura, papel inicial, carga, unificação (C17),
   órfãos (C18), índice único, relatório. Segunda execução não altera nada.
4. Banco de teste com casos fictícios: variantes de caixa e acento, duplicata, órfão, etapa
   especial ausente.
5. Montagem: os 12 passos na ordem; C19 e C20; erro forçado em cada um dos passos 1 a 11 → banco e
   `.xlsm` com o mesmo hash de antes e mensagem com o passo; `-Teste` recusa banco PRODUCAO.
6. Front e conector leem o papel da lista (C07); nenhum texto fixo de etapa no programa (checagem
   no `verificar.py`).
7. Ficha, cadastro em lote e conector seguem D-j (C06).
8. Conector aceita 1.0 e 1.1; front `VERSAO_ESQUEMA` = 1.1.
9. Empacotamento gera a pasta e o `MANIFESTO.json`; a conferência de proibidos passa.

### Fase 2 — Gerenciamento de listas (publicação A)

1. `frmListas` aberto pelo Início; incluir (C02, C05), desativar/reativar (C03), subir/descer,
   editar papel (C04); mudar `encerra` de etapa em uso mostra quantos atendimentos mudam e pede
   confirmação.
2. Log na mesma transação (INSERIR, ALTERAR, DESATIVACAO, REATIVACAO); conflito conforme a tabela.
3. Cache limpo na estação; o feed limpa nas outras em até 25 s.
4. `gerenciar_lista` e `listar_opcoes` com as mesmas regras (C02–C05).

### Fase 3 — Visitas no atendimento (publicação B)

1. Botão `Visitas (n)`; com visitas abre o conjunto; sem visitas abre Novo ligado ao atendimento.
2. V1–V9 (C08–C13) no front e no conector.
3. Feed atualiza a ficha aberta e a aba Oportunidades depois da V6.

### Fase 4 — Planejamentos (publicação C)

1. `frmPlanejamento`: criar e editar (P1, C14), período (P2, C15), encerrar/reabrir (P3, C16), P4.
2. Candidatos: ordem por qualificação, dias parado e km; "pendente em PL-…"; cidades não
   reconhecidas à parte; Gerar visitas numa transação.
3. Lançamento na página própria, com Salvar e próxima (V3–V7).
4. Aba `Planejamentos` com seletor e duplo clique; feed atualiza a aba e o formulário aberto.
5. Haversine (6.371 km), mesmo município = 0 km; cidade do cliente por UF + `nome_busca`.
6. Ferramentas de planejamento, candidatos, lotes e municípios com as mesmas regras.

### Fase 5 — Exportação (publicação C)

1. `Documentos\CRM Zapromaq\Roteiros\PL-AAAA-NNN_aaaa-mm-dd_a_aaaa-mm-dd.csv`: UTF-8 com BOM, `;`,
   AAAA-MM-DD, ponto decimal, linha `BASE` primeiro, uma linha por visita PENDENTE.
2. `exportar_roteiro` devolve só o caminho e o número de linhas.
3. Conector exige 1.1; com banco 1.0, gravar é recusado com mensagem clara.

## Itens 🖥

- **Bloqueiam a publicação:** A — `frmListas` em 1366×768 e ficha com valor inativo marcado;
  `MONTAR-FRONTEND.bat -Teste` completo nesta estação. B — ficha da visita e botão Visitas (n).
  C — `frmPlanejamento` (três páginas) e aba Planejamentos; CSV aberto no Excel 2019.
- **Não bloqueiam (só conferíveis na empresa, durante a publicação):** leitura dos nomes de
  estação na trava; `[IO.File]::Replace` no compartilhamento da rede (há alternativa de mover e
  renomear); tempo da carga pela rede; relatório de unificação com os dados reais.

## Limites

- **Sempre:** 100% local; Excel 2019; dados fictícios no repositório; log e versão; nenhuma
  exclusão de registro de negócio; SQL parametrizado; data como parâmetro.
- **Perguntar antes:** publicar na empresa; mudar contrato do conector além do listado aqui.
- **Nunca:** alterar o banco de produção fora do fluxo da montagem; `-Teste` em banco PRODUCAO;
  dado de cliente no repositório; apagar `.laccdb`.

## Questões em aberto

Nenhuma. O que depende de dado real (duplicatas, variantes, órfãos, nomes na trava, tempo da
carga) é tratado pela própria montagem e aparece no relatório e na mensagem final.
