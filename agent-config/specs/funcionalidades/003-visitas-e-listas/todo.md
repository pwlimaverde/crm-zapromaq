# Tarefas (refeitas em 12/10/2026)

Ciclo: teste falha → código → `verificar-tudo` no nível indicado → commit. Plano e ordem em
[plan.md](plan.md). 🖥B = bloqueia a publicação; 🖥N = não bloqueia.

## Fase 1: concluir (publicação A)

- [x] 1.1 Base de municípios do IBGE (`803b9d6`)
- [x] 1.2 Chave de lista C01 nos dois lados (`fbc8701`)
- [x] 1.3 Estrutura do esquema 1.1 (`d4042ab`)
- [x] 1.4 Casos problemáticos no banco de teste (`7f5fd47`)
- [x] 1.5 Migração 003: carga, unificação, órfãos, papel (`62fe3c6`)
- [x] 1.6 Rotina tudo ou nada, C19 e C20 (`d2748f1`)
- [x] 1.7 Montagem em 12 passos (`f959adb`); falha forçada automatizada nos passos 9 a 12 ainda pendente (1.7b)
- [x] 1.8 Papel das etapas lido da lista; valor inativo, C06 e C07 (`e633c12`)
- [x] 1.9 Esquema 1.1 no front e no conector; banco de demonstração (`998f8d7`)
- [x] 1.10 MANIFESTO.json e conferência no passo 1 (`fd815e1`)
- [x] 1.7b Falha forçada automatizada nos passos 9 a 12 — `Conferir` em `testar-build-lib.ps1`
- [x] 1.10b Pacote em pasta vazia com banco de teste roda sem outro comando
- [ ] 1.11 Documentação da Fase 1 (`banco-access`, `build-operacao-powershell`, `ferramentas-dev`, `contrato-banco`, `levar-para-a-empresa`, `implantacao`, `CLAUDE.md` Comandos, `PLANO.md` D21+) — `verificar.py`
- [ ] 🖥B `MONTAR-FRONTEND.bat -Teste` completo nesta estação; trava falsa

## Fase 2: gerenciamento de listas (publicação A)

- [x] 2.1 Regras C03, C04, C05 nos dois lados (`d4a8c5a`)
- [x] 2.5a `gerenciar_lista` com `definir_papel` e `listar_opcoes` com `detalhar` no conector (`d4a8c5a`), testes em `testar-mcp`
- [x] 2.2a Gravação no front (incluir, ativar, desativar, ordenar), log e conflito — `testar-uso` + `Confere` (C02)
- [x] 2.2b `definir_papel` no front: C04, contagem de atendimentos, `confirmar` — `testar-uso`
- [x] 2.3 Feed e cache de `listas` (`modFeed`) — `Confere`
- [x] 2.4a `frmListas` layout (JSON, `.txt`, prévia) — `verificar.py`
- [x] 2.4b (código pronto; interação só 🖥) `frmListas` código (combo, grade, botões, confirmação de encerra) — `testar-uso` + `verificar.py`
- [x] 2.4c (botão Gerenciamento no Início; 🖥 pendente) Botão Gerenciamento no Início — 🖥B em 1366×768
- [ ] 2.5b Documentação e `SKILL.md`, `IA\LEIA-ME.md`, `front-vba.md`, `visual-layout.md`, `ia-mcp.md`
- [ ] Checkpoint A: `verificar-tudo -Completo`, 🖥B acima, autorização do usuário

## Fase 3: visitas no atendimento (publicação B)

- [x] 3.1 (C08–C13 nos dois lados; V1–V9 na spec) Regras V2–V5, V8 e texto V6 (C08–C12) nos dois lados — `Confere` e `testar-mcp`
- [x] 3.2a Gravação no conector: `criar_visita`, `atualizar_visita`, V6/V7 na transação — `testar-mcp`
- [x] 3.2b Gravação no front (`modVisita.bas`) — `testar-uso`
- [x] 3.3a Botão `Visitas (n)` na ficha do atendimento — `verificar.py` e prévia (🖥 pendente)
- [~] 3.3b Tela de visitas (`frmVisitas`: lista e ficha) feita; falta o feed de `visitas` que atualiza a tela aberta — `testar-uso` + prévia; 🖥B
- [x] 3.4 Ferramentas `listar_visitas`, `alterar_situacao_visita`, resumo em `obter_atendimento` — `testar-mcp`
- [ ] Checkpoint B: 🖥B acima; banco exclusivo na janela (A3)

## Fase 4: planejamentos (publicação C)

- [x] 4.1 `nome_busca` e haversine — `Confere` e `testar-mcp`
- [x] 4.2 P1–P4 e ordem dos candidatos (C14–C16) — `Confere` e `testar-mcp`
- [ ] 4.3a Gravar planejamento (`PL-`, período, raio 1–500), VBA e conector — `testar-uso`
- [ ] 4.3b Encerrar e reabrir (P3) — `testar-uso`
- [ ] 4.4 Gerar visitas em transação, V8 dentro dela, corrida recusada — `testar-uso`
- [ ] 4.5 Lançamento de datas e resultados (rotina) — `testar-uso` e `testar-mcp`
- [ ] 4.6 `clsGrade` com marcação — `verificar.py` e `testar-uso`
- [ ] 4.7a `frmPlanejamento` layout (cabeçalho e três páginas) — prévia e `verificar.py`
- [ ] 4.7b Páginas Visitas e Candidatos — `testar-uso`
- [ ] 4.7c Página Lançamento — `testar-uso`
- [ ] 4.7d `frmVinculo` modo `municipio` — `verificar.py` e `testar-uso`
- [ ] 4.8 Aba Planejamentos (seletor, duplo clique, feed) — `testar-uso`; 🖥B
- [ ] 4.9a Ferramentas de planejamento (5) — `testar-mcp`
- [ ] 4.9b Candidatos, `gerar_visitas`, lotes, `buscar_municipios`, `resolver_cidade` — `testar-mcp`
- [ ] Checkpoint C: 🖥B das três páginas em 1366×768; autorização

## Fase 5: exportação (publicação C)

- [ ] 5.1 Linha do CSV (depende de A2) — `Confere`
- [ ] 5.2 Gravação do arquivo em `modRelatorio.bas` — `testar-uso`; 🖥B abrir no Excel 2019
- [ ] 5.3 `exportar_roteiro`, `exportar_base`; conector com 41 ferramentas — `testar-mcp`

## Documentação final (com a publicação C)

- [ ] `SPEC.md` carimbada com `> Concluída em dd/mm/aaaa — feature/visitas-e-listas`
- [ ] `PLANO.md` com as decisões D21+ e o estado de cada fase

## Decisões pendentes do usuário

- [ ] A1 Etapa com papel editável e UF só validada na inclusão (confirmar até o checkpoint A)
- [ ] A2 Nome e telefone do contato no CSV e no conector (até a Fase 5)
- [ ] A3 Janela da publicação B com banco exclusivo (até o checkpoint B)
