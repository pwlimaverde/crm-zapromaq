# Plano (refeito em 12/10/2026): visitas, planejamentos, listas e conector

Spec: [SPEC.md](SPEC.md) · Tarefas: [todo.md](todo.md) · Branch: `feature/visitas-e-listas`

## Situação ao refazer o plano

| Fase | Estado | Commits |
|---|---|---|
| 1 Base, migração 003, montagem única | **Concluída no código**; faltam o teste automatizado da falha forçada (passos 9–12), a documentação e os 🖥 | `803b9d6` … `fd815e1` |
| 2 Gerenciamento de listas | **Parcial**: regras puras nos dois lados (C03–C05), `gerenciar_lista` com `definir_papel`, `listar_opcoes` com `detalhar`. Falta a gravação do front, o feed, o `frmListas` e a documentação | `d4a8c5a` |
| 3 Visitas no atendimento | Não iniciada | — |
| 4 Planejamentos | Não iniciada | — |
| 5 Exportação | Não iniciada | — |

Contagem atual do conector: 25 ferramentas. Meta: 41.

## Decisões que o plano assume (confirmar antes da fase indicada)

| # | Assunção | Confirmar até |
|---|---|---|
| A1 | Etapa tem papel editável; UF só é validada na inclusão (C05), sem restringir ativar/desativar | Fase 2 (publicação A) |
| A2 | D-a: CSV e respostas do conector mantêm nome e telefone do contato | Fase 5 |
| A3 | Publicação B (visitas) exige banco exclusivo na janela de publicação | Publicação B |
| A4 | Cada publicação passa pelo `/ship` e só sai com autorização explícita | Cada publicação |

## Princípios do replanejamento

- **Fatia = um comportamento verificável.** Cada fatia tem um teste que falha primeiro, um
  commit e a prova (`verificar-tudo` no nível indicado). Nenhuma fatia passa de um dia de trabalho.
- **Regra antes da tela.** Regra pura (C01–C20) vem primeiro, nos dois lados, com o mesmo número.
- **Tela em três passos:** layout (JSON + controles + prévia) → código de eventos → calibração 🖥.
  Isso separa o que dá para provar aqui do que só a estação confirma.
- **Checkpoint por publicação.** A, B e C terminam com `verificar-tudo -Completo`, teste de uso
  e lista de 🖥. A publicação não começa sem o aval do usuário.

## Fase 1: concluir (sem mudar o banco)

Ordem: 1.7b → 1.10b → 1.11 (docs) → checkpoint 🖥 da montagem.

- **1.7b** Teste automatizado da falha forçada nos passos 9 a 12 (variável de teste controla o
  passo que falha). Teste que falha primeiro: `Conferir` em `testar-build-lib.ps1`.
- **1.10b** Pacote em pasta vazia com banco de teste: a montagem roda sem outro comando.
  Teste: caso em `testar-build-lib.ps1` (ou roteiro manual registrado).
- **1.11** Documentação: `banco-access.md`, `build-operacao-powershell.md`, `ferramentas-dev.md`,
  `contrato-banco.md`, `levar-para-a-empresa.md`, `implantacao.md`, `CLAUDE.md` (Comandos) e
  `PLANO.md` (D21+). Checagem: `verificar.py` (links e nomes citados existem).
- **Checkpoint 🖥 (bloqueia a publicação A):** `MONTAR-FRONTEND.bat -Teste` completo nesta
  estação; trava falsa; troca no compartilhamento fica para a estação da empresa.

## Fase 2: gerenciamento de listas (publicação A)

Ordem: 2.2a → 2.2b → 2.3 → 2.4a → 2.4b → 2.4c → 2.5 → checkpoint.

- **2.2a Gravação no front, valores** (`modListas.bas`): incluir, ativar, desativar, ordenar,
  com log na mesma transação e conflito pelo estado antigo (`WHERE id=? AND ativo=?`).
  Teste que falha primeiro: `testar-uso.ps1` (caso novo) e `Confere` de C02 (chave existente).
- **2.2b Papel no front**: `definir_papel` com C04, contagem de atendimentos e
  `confirmar`. Teste que falha primeiro: `testar-uso.ps1` (encerra muda com 3 atendimentos
  recusa sem confirmar).
- **2.3 Feed e cache**: `modFeed` atualiza `listas` nas outras estações em até 25 s e o cache
  de `ObterLista` é limpo. Teste: `Confere` do cache (chave limpa ao receber evento).
- **2.4a `frmListas` layout**: `formularios\frmListas.json` e `.txt` (nomes de controle), prévia
  e `verificar.py` (controles × layout). Sem código de evento ainda.
- **2.4b `frmListas` código**: combo da lista, grade, Incluir, Desativar/Reativar, Subir, Descer,
  Editar papel, confirmação de encerra. Teste: `testar-uso.ps1` nos casos da grade e `verificar.py`.
- **2.4c Início**: botão Gerenciamento (`AbrirGerenciamento`). 🖥 em 1366×768.
- **2.5 Conector e documentação**: `SKILL.md`, `IA\LEIA-ME.md`, `front-vba.md`, `visual-layout.md`,
  `ia-mcp.md`. Teste: `testar-mcp.ps1` (já cobre `gerenciar_lista` e `listar_opcoes`).
- **Checkpoint A (bloqueia a publicação):** `verificar-tudo -Completo`; 🖥 de `frmListas` e de
  ficha com valor inativo marcado; autorização do usuário.

## Fase 3: visitas no atendimento (publicação B)

Ordem: 3.1 → 3.2a → 3.2b → 3.3a → 3.3b → 3.4 → checkpoint.

- **3.1 Regras V2–V5, V8 e texto V6** (C08–C12), puras nos dois lados.
  Teste: `Confere` (VBA) e `testar-mcp` (conector).
- **3.2a Gravação no conector**: `criar_visita` e `atualizar_visita`, com V6/V7 na mesma
  transação e conflito que desfaz a visita junto. Teste: `testar-mcp`.
- **3.2b Gravação no front** (`modVisita.bas`): mesmas regras, mesmo conflito. Teste: `testar-uso`.
- **3.3a Botão `Visitas (n)`** na ficha do atendimento, com a contagem. Layout + `verificar.py`.
- **3.3b Ficha da visita** (`frmCRM`, tabela `visitas`) e feed de `visitas` (atualiza a ficha
  aberta e a aba Oportunidades depois da V6). Teste: `testar-uso` e prévia.
- **3.4 Ferramentas de visita**: `listar_visitas`, `criar_visita`, `atualizar_visita`,
  `alterar_situacao_visita`, e resumo de visitas em `obter_atendimento`. Teste: `testar-mcp`.
- **Checkpoint B (bloqueia a publicação):** 🖥 da ficha da visita e do botão `Visitas (n)`;
  banco exclusivo na janela de publicação (A3).

## Fase 4: planejamentos (publicação C)

Ordem: 4.1 → 4.2 → 4.3a → 4.3b → 4.4 → 4.5 → 4.6 → 4.7a → 4.7b → 4.7c → 4.7d → 4.8 → 4.9a → 4.9b → checkpoint.

- **4.1 `nome_busca` e haversine** (6.371 km; mesmo município = 0). Teste: `Confere` e `testar-mcp`.
- **4.2 P1–P4 e ordem dos candidatos** (C14–C16). Teste: `Confere` e `testar-mcp`.
- **4.3a Gravar planejamento** (código `PL-`, período, raio 1–500), VBA e conector. Teste: `testar-uso`.
- **4.3b Encerrar e reabrir** (P3, sem `cancelar_pendentes` por padrão). Teste: `testar-uso`.
- **4.4 Gerar visitas**: transação única, V8 de novo dentro dela. Teste: `testar-uso` (corrida
  simulada: a segunda geração recebe recusa).
- **4.5 Lançamento de datas e resultados** (rotina, sem tela). Teste: `testar-uso` e `testar-mcp`.
- **4.6 `clsGrade` com marcação** (seleção múltipla). Teste: `verificar.py` e `testar-uso`.
- **4.7a `frmPlanejamento` layout** (cabeçalho e três páginas). Prévia e `verificar.py`.
- **4.7b Página Visitas e Candidatos** (grade, Marcar todos, Gerar visitas, Exportar). `testar-uso`.
- **4.7c Página Lançamento** (Salvar e próxima, Anterior). `testar-uso`.
- **4.7d `frmVinculo` modo `municipio`**. `verificar.py` e `testar-uso`.
- **4.8 Aba Planejamentos**: seletor em `C4`, duplo clique abre a página Lançamento, feed de
  `planejamentos`. `testar-uso`. 🖥 seletor e duplo clique.
- **4.9a Ferramentas de planejamento**: `listar_planejamentos`, `obter_planejamento`,
  `criar_planejamento`, `atualizar_planejamento`, `alterar_situacao_planejamento`. `testar-mcp`.
- **4.9b Candidatos, lotes e municípios**: `listar_candidatos`, `gerar_visitas`, `lancar_datas`,
  `lancar_resultados` (`simular` padrão true, até 100 itens), `buscar_municipios`,
  `resolver_cidade`. `testar-mcp`.
- **Checkpoint C (bloqueia a publicação):** 🖥 das três páginas em 1366×768; aba e duplo clique.

## Fase 5: exportação (publicação C)

Ordem: 5.1 → 5.2 → 5.3 → checkpoint.

- **5.1 Linha do CSV** (UTF-8 com BOM, `;`, ponto decimal). Depende de A2 para nome e telefone.
  Teste: `Confere`.
- **5.2 Gravação do arquivo** (`modRelatorio.bas`) na pasta de exportação. Teste: `testar-uso`.
  🖥 abrir no Excel 2019.
- **5.3 Ferramentas**: `exportar_roteiro` (devolve só o caminho), `exportar_base` (incluindo
  `planejamentos`, `visitas`, `municipios`). Conector chega a 41 ferramentas. Teste: `testar-mcp`.
- **Checkpoint C final:** `verificar-tudo -Completo`; contagem de ferramentas = 41.

## Checkpoints de publicação

| Publicação | Fases | Exige |
|---|---|---|
| A | 1 + 2 | `verificar-tudo -Completo`; 🖥 de `frmListas`; autorização; `MONTAR-FRONTEND.bat` na empresa |
| B | 3 | Idem A; banco exclusivo na janela (A3); 🖥 da ficha da visita |
| C | 4 + 5 | Idem; 🖥 das três páginas, da aba e do CSV no Excel 2019 |

## Riscos e mitigação

| Risco | Mitigação |
|---|---|
| Formulário novo estoura 1366×768 | Layout no JSON + prévia antes do código; 🖥 cedo (2.4c) |
| Regra divergente entre VBA e conector | Mesmo número C; `verificar.py` confere os dois lados |
| Gravação de planejamento corrompe visitas | Transação única e V8 dentro dela; teste de corrida |
| Escopo cresce (fases 3–5 são grandes) | Cada publicação é independente; A pode sair antes das outras |
| Dependência de decisão aberta (A1, A2) | Assunção registrada; muda só uma fatia, não o plano |
