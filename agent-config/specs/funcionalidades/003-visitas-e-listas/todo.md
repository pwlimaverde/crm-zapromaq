# Tarefas: visitas, planejamentos, listas padrão e conector completo

Ciclo: teste falha → código → `verificar-tudo` no nível certo → commit. Testes de banco só no
banco de teste da pasta. 🖥B = bloqueia a publicação; 🖥N = não bloqueia.

## Fase 1 — Base, migração 003 e montagem única (publicação A)

- [x] 1.1 `gerar-municipios.py` + `municipios.csv` + `municipios.fonte.json`
  - Teste: `verificar.py` · Arquivos: `agent-config\ferramentas\gerar-municipios.py`, `banco\referencia\*`
- [x] 1.2 Função de chave de lista (C01) e checagem de casos compartilhados
  - Teste: `Confere` + `testar-mcp.ps1` + `verificar.py` · Arquivos: `modValidacao.bas`, `Regras.ps1`, `verificar.py`
- [x] 1.3 DDL e `modSchema`
  - Teste: `verificar.py` · Arquivos: `banco\esquema\criar-banco.ps1`, `modSchema.bas`
- [ ] 1.4 Banco de teste com variantes, duplicata, órfão e etapa especial ausente
  - Teste: (infraestrutura do 1.5) · Arquivos: `IA\teste\criar-banco-teste.ps1`
- [ ] 1.5 Migração 003 (C17, C18, papel, índice, idempotência, relatório)
  - Teste: caso que aplica no banco de teste · Arquivos: `banco\migracoes\003-municipios-visitas-papel-listas.ps1`
- [ ] 1.6 Rotina de atualização do banco (trava, backup + H0, cópia, conferências, troca, volta)
  - Teste: `Conferir` em `testar-build-lib.ps1` (C19, C20) · Arquivos: `banco\aplicar-migracoes.ps1`, `lib\Comum.ps1`
- [ ] 1.7 Montagem nos 12 passos, `Config!B2` provisório/definitivo, mensagens, `-Teste` só TESTE
  - Teste: `Conferir` em `testar-build-lib.ps1` · Arquivos: `build\montar-frontend.ps1`, `MONTAR-FRONTEND.bat`, `testes\testar-uso.ps1`
  - 🖥B `MONTAR-FRONTEND.bat -Teste` completo nesta estação · 🖥N `Replace` e nomes da trava na rede
- [ ] 1.8 Papel das etapas lido da lista; valor inativo (C06, C07)
  - Teste: `Confere` + `testar-uso.ps1` + `testar-mcp.ps1` + `verificar.py` · Arquivos: `modCalc`, `modCRM`, `modMenu`, `modGrade`, `modAcoes`, `modPainel`, `modRelatorio`, `modValidacao`, `modLote`, `frmCRM.txt`, `Regras.ps1`, `Ferramentas.ps1`
  - 🖥B ficha com valor inativo marcado
- [ ] 1.9 Conector 1.0/1.1; front 1.1; banco de demonstração
  - Teste: `testar-mcp.ps1` + `testar-uso.ps1` · Arquivos: `modConfig.bas`, `Banco.ps1`, `servidor.ps1`, `criar-banco-demo.ps1`
- [ ] 1.10 Empacotamento em pasta + `MANIFESTO.json`
  - Teste: `Conferir` em `testar-build-lib.ps1` (passo 1) + checagem no `verificar-tudo` · Arquivos: `agent-config\ferramentas\empacotar-base.ps1`

## Fase 2 — Gerenciamento (publicação A)

- [ ] 2.1 Regras (C02–C05, `encerra` em uso) — `Confere` · `modListas.bas` (novo)
- [ ] 2.2 Gravação com log e conflito — `testar-uso.ps1` · `modListas.bas`
- [ ] 2.3 Feed e cache — `Confere` · `modFeed.bas`
- [ ] 2.4 `frmListas` + botão Gerenciamento — `verificar.py` + prévia · `frmListas.txt`, `frmListas.json`, `abas.json` · 🖥B 1366×768
- [ ] 2.5 `gerenciar_lista` e `listar_opcoes` — `testar-mcp.ps1` · `Ferramentas.ps1`, `SKILL.md`

## Fase 3 — Visitas no atendimento (publicação B)

- [ ] 3.1 Regras C08–C13 — `Confere` · `modVisita.bas` (novo)
- [ ] 3.2 Gravação V6/V7 e conflito — `testar-uso.ps1` · `modVisita.bas`, `modCRM.bas`
- [ ] 3.3 Ficha da visita, botão Visitas (n), feed — `Confere` + `verificar.py` · `frmCRM.txt`, `modSchema.bas`, `modFeed.bas` · 🖥B
- [ ] 3.4 Ferramentas de visita — `testar-mcp.ps1` · `Ferramentas.ps1`, `SKILL.md`

## Fase 4 — Planejamentos (publicação C)

- [ ] 4.1 `nome_busca` + haversine — `Confere` · `modPlanejamento.bas` (novo)
- [ ] 4.2 P1–P4 e ordem (C14–C16) — `Confere`
- [ ] 4.3 Gravação de planejamento — `testar-uso.ps1`
- [ ] 4.4 Gerar visitas — `testar-uso.ps1`
- [ ] 4.5 Lançamento (rotina) — `testar-uso.ps1`
- [ ] 4.6 `clsGrade` com marcação — `verificar.py` + `testar-uso.ps1` · `clsGrade.cls`
- [ ] 4.7 `frmPlanejamento` + `frmVinculo` modo `municipio` — `verificar.py` + prévia · `frmPlanejamento.txt`, `frmPlanejamento.json`, `frmVinculo.txt` · 🖥B três páginas
- [ ] 4.8 Aba Planejamentos + feed — `testar-uso.ps1` · `abas.json`, `modGrade.bas`, `modFeed.bas`, `ThisWorkbook.txt` · 🖥B seletor e duplo clique
- [ ] 4.9 Ferramentas de planejamento e municípios — `testar-mcp.ps1` · `Ferramentas.ps1`, `SKILL.md`

## Fase 5 — Exportação (publicação C)

- [ ] 5.1 Linha do CSV — `Confere` · `modRelatorio.bas`
- [ ] 5.2 Gravação do arquivo — `testar-uso.ps1` · 🖥B abrir no Excel 2019
- [ ] 5.3 `exportar_roteiro`, `exportar_base`, só 1.1 — `testar-mcp.ps1` · `Ferramentas.ps1`, `Banco.ps1`

## Documentação (por fase, no mesmo commit da fatia que muda a stack)

- [ ] Fase 1: `banco-access.md`, `build-operacao-powershell.md`, `ferramentas-dev.md`, `contrato-banco.md`, `levar-para-a-empresa.md`, `implantacao.md`, `CLAUDE.md` (Comandos), `PLANO.md`
- [ ] Fase 2: `front-vba.md`, `visual-layout.md`, `ia-mcp.md`, `SKILL.md`, `IA\LEIA-ME.md`
- [ ] Fases 3–5: `front-vba.md`, `visual-layout.md`, `ia-mcp.md`, `SKILL.md`, `IA\LEIA-ME.md`, `contrato-banco.md`
- [ ] Final: `SPEC.md` carimbada com `> Concluída em dd/mm/aaaa — feature/<nome>`
