# Plano: visitas, planejamentos, listas padrão e conector completo

Spec: [SPEC.md](SPEC.md)

## Publicações

| Publicação | Fases | Exige da estação |
|---|---|---|
| A | 1 + 2 | Excel 2019 64 bits; ninguém com o CRM aberto (trava conferida); Claude Desktop reiniciado e skill reinstalada |
| B | 3 | Excel 2019; ninguém com o CRM aberto; skill reinstalada |
| C | 4 + 5 | Idem B |

Em todas: o usuário copia `dist\BASE-v<X.Y>\BASE\` por cima da base da empresa e roda só
`SISTEMA\MONTAR-FRONTEND.bat`.

## Ordem e dependências

**Fase 1 — Base, migração 003 e montagem única (A)**
- 1.1 Script de preparação + `municipios.csv` + `.fonte.json` — depende de: —
- 1.2 Função de chave de lista (front e conector) + checagem de casos compartilhados — depende de: —
- 1.3 Estrutura: DDL de `municipios`, `planejamentos`, `visitas`, colunas novas de `listas`; `modSchema` — depende de: 1.1
- 1.4 Banco de teste com casos problemáticos fictícios — depende de: 1.3
- 1.5 Migração 003 (estrutura, papel, carga, unificação, órfãos, índice, relatório) — depende de: 1.2, 1.3, 1.4
- 1.6 Rotina de atualização do banco (trava com nomes, backup + H0, cópia, migração na cópia, conferências, troca segura, volta ao backup) — depende de: 1.5
- 1.7 Montagem nos 12 passos, `Config!B2` provisório/definitivo, mensagens, `-Teste` só com TESTE — depende de: 1.6
- 1.8 Papel das etapas lido da lista (front e conector) + valor inativo na ficha, no lote e no conector — depende de: 1.5
- 1.9 Conector 1.0/1.1, front 1.1, banco de demonstração — depende de: 1.5, 1.8
- 1.10 Empacotamento: pasta pronta + `MANIFESTO.json`; passo 1 da montagem confere — depende de: 1.7

**Fase 2 — Gerenciamento (A)**
- 2.1 Regras puras (C02–C05, mudança de `encerra` em uso, mapa de uso, ordem) — depende de: 1.8
- 2.2 Gravação com log e conflito — depende de: 2.1
- 2.3 Feed e cache de `listas` — depende de: 2.2
- 2.4 `frmListas` + botão Gerenciamento — depende de: 2.2, 2.3
- 2.5 `gerenciar_lista` (`definir_papel`) e `listar_opcoes` (`detalhar`) — depende de: 2.1, 2.2

**Fase 3 — Visitas no atendimento (B)**
- 3.1 Regras V2–V5, V8, texto da V6 (C08–C13) — depende de: 1.9
- 3.2 Gravação com V6/V7 e conflito — depende de: 3.1
- 3.3 Ficha da visita + botão Visitas (n) + feed de `visitas` — depende de: 3.2
- 3.4 Ferramentas de visita — depende de: 3.2

**Fase 4 — Planejamentos (C)**
- 4.1 `nome_busca` + haversine — depende de: 1.9
- 4.2 P1–P4 e ordem dos candidatos (C14–C16) — depende de: 4.1
- 4.3 Gravação de planejamento (código PL-, encerrar/reabrir) — depende de: 4.2
- 4.4 Gerar visitas (transação, V8 de novo dentro dela) — depende de: 4.2, 3.2
- 4.5 Lançamento de datas e resultados (rotina sem tela) — depende de: 4.3, 3.2
- 4.6 `clsGrade` com marcação (seleção múltipla) — depende de: —
- 4.7 `frmPlanejamento` (três páginas) + `frmVinculo` modo `municipio` — depende de: 4.3, 4.4, 4.5, 4.6
- 4.8 Aba Planejamentos + feed de `planejamentos` e `visitas` — depende de: 4.7
- 4.9 Ferramentas de planejamento, candidatos, lotes e municípios — depende de: 4.3, 4.4, 4.5

**Fase 5 — Exportação (C)**
- 5.1 Linha do CSV — depende de: 4.7
- 5.2 Gravação do arquivo — depende de: 5.1
- 5.3 `exportar_roteiro`, `exportar_base`, conector só 1.1 — depende de: 5.2, 4.9

Front e conector de uma mesma regra andam na mesma fase; a regra e os casos C vêm primeiro.

## Teste que falha primeiro

| Fatia | Teste (tabela do `PROJETO.md`) |
|---|---|
| 1.1 | checagem nova em `verificar.py` (CSV, hash, faixas) |
| 1.2 | `Confere` (C01) + caso em `testar-mcp.ps1` (C01) + `verificar.py` (todo C em ambos os lados) |
| 1.3 | `verificar.py` (`modSchema` × DDL) |
| 1.4, 1.5 | caso que aplica a 003 no banco de teste (`criar-banco-teste.ps1`): C17, C18, papel, índice, idempotência |
| 1.6, 1.7, 1.10 | `Conferir` em `testes\testar-build-lib.ps1` (C19, C20, falha em cada passo, `-Teste` com PRODUCAO, manifesto) |
| 1.8 | `Confere` (C06, C07) + `testar-uso.ps1` + `testar-mcp.ps1` + `verificar.py` (nenhum texto fixo de etapa) |
| 1.9 | `testar-mcp.ps1` + `testar-uso.ps1` |
| 2.1, 2.3, 3.1, 4.1, 4.2, 5.1 | `Confere` em `modAutoteste` |
| 2.2, 3.2, 4.3, 4.4, 4.5, 5.2 | `testes\testar-uso.ps1` |
| 2.4, 3.3, 4.6, 4.7, 4.8 | `verificar.py` (controles × layout) + `testar-uso.ps1` + prévia + 🖥 |
| 2.5, 3.4, 4.9, 5.3 | `IA\teste\testar-mcp.ps1` |

## Checkpoints

- Depois de 1.7: `MONTAR-FRONTEND.bat -Teste` completo; erro forçado em cada passo; trava falsa.
- Depois de 1.10: copiar `dist\BASE-v<X.Y>\BASE\` para uma pasta vazia com um banco de teste e
  rodar a montagem: tudo pronto sem outro comando.
- Depois de 2.4, 3.3, 4.8: telas em 1366×768 (🖥 bloqueante).
- Antes de cada publicação: `verificar-tudo -Completo` com a saída anexada.

## Riscos

| Risco | Mitigação |
|---|---|
| Falha depois da migração deixa banco 1.1 com CRM antigo | Migração e testes na cópia; troca só no passo 11 |
| Alguém grava durante a montagem | Hash H0 conferido antes da troca; se mudou, nada é trocado |
| Trava antiga sem ninguém usando | Mostra as estações e orienta; nunca apaga |
| `Replace` não suportado no compartilhamento | Alternativa: mover produção → backup e cópia → produção, desfazendo se o segundo falhar |
| Unificação em massa inesperada | Regra determinística, log por registro, relatório na mensagem final |
| Etapa com texto diferente no banco real | Comparação pela chave; especial ausente é criada inativa com papel |
| Texto de etapa esquecido no programa | `verificar.py` recusa texto fixo de etapa em `fonte\` e em `IA\mcp\lib` |
| Regras divergentes entre front e conector | Casos C numerados, conferidos dos dois lados |
| Pasta copiada com sobras de desenvolvimento | Pasta gerada pelo empacotamento + manifesto conferido no passo 1 |
| Corrida na V8 entre dois usuários | Conferência dentro da transação; o segundo recebe recusa |
| `-Teste` atingir produção | Recusa se `config.ambiente` ≠ TESTE |

## Documentação (arquivo exato)

- `agent-config\doc_dev\planejamento\PLANO.md` — decisões D-a a D-k como D21 em diante
- `agent-config\specs\stacks\banco-access.md` — migração na cópia, troca segura, `chave` e papel
- `agent-config\specs\stacks\build-operacao-powershell.md` — os 12 passos, `-Teste`, manifesto
- `agent-config\specs\stacks\front-vba.md` — `frmListas`, `frmPlanejamento`, D-j, papel das etapas
- `agent-config\specs\stacks\ia-mcp.md` — 41 ferramentas, esquemas 1.0/1.1 → 1.1
- `agent-config\specs\stacks\visual-layout.md` — layouts novos e marcação da grade
- `agent-config\specs\stacks\ferramentas-dev.md` — `gerar-municipios.py`, empacotamento em pasta
- `src\BASE\SISTEMA\DADOS\docs\contrato-banco.md` — tabelas novas e colunas de `listas`
- `src\BASE\SISTEMA\DADOS\docs\levar-para-a-empresa.md` — copiar a pasta e rodar a montagem
- `src\BASE\SISTEMA\DADOS\docs\implantacao.md` — mensagens da montagem e o que fazer em cada parada
- `src\BASE\IA\LEIA-ME.md` e `src\BASE\IA\instrucoes\crm-zapromaq\SKILL.md`
- `agent-config\CLAUDE.md` — seção Comandos: `-Teste` agora migra o banco de teste
