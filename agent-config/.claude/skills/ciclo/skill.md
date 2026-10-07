---
name: ciclo
description: Orquestra uma funcionalidade ou correção de ponta a ponta neste projeto (diagnóstico, spec e plano em fatias, TDD fatia por fatia com autocorreção limitada, verificação, revisão por personas e integração em develop), com dois portões de aprovação humana. Use quando o usuário pedir /ciclo ou quiser que uma demanda seja conduzida do diagnóstico ao merge sem chamar cada comando. Não use para correção pequena e bem definida (/spec curto + /build) nem para publicar na empresa (/ship).
---

# Ciclo (orquestrador de desenvolvimento)

## Visão geral

Conduz uma demanda do **CRM Zapromaq** pelo processo inteiro de `DIRETRIZES-IA.md`
(`/spec → /plan → /build → /test → /review → /code-simplify → /ship`, nível "integrar") sem
exigir um comando por etapa. São **5 estágios com 2 portões** de aprovação humana; entre eles
o agente trabalha sozinho, mas **dentro** das paradas obrigatórias do projeto, que valem sempre.

Esta skill não repete o que as outras já dizem: ela indica **qual** usar em cada estágio e
**onde** parar. As regras do projeto moram em `CLAUDE.md`, `specs/PROJETO.md` e
`specs/stacks/*.md`.

Caminhos relativos a `agent-config/` (onde o Claude Code é aberto); a raiz do git é `..`.

**Glossário**

- **GATE** — portão: ponto em que o agente para e só segue com aprovação explícita do usuário.
- **RED / GREEN** — escrever primeiro o teste e vê-lo falhar (RED); depois o código mínimo que o
  faz passar (GREEN).
- **Fatia vertical** — pedaço pequeno da demanda que atravessa as camadas necessárias e pode ser
  testado e commitado sozinho.
- **Persona** — agente revisor especializado (`.claude/agents/`), chamado com o diff da branch.

## Convivência com o roteador (Jev)

Durante um ciclo, **esta skill prevalece** sobre qualquer indicação do roteador
(`[seletor de skill] …` ou `[roteador] … Delegue ao subagente …`). Depois do `/ciclo`, o
roteador não trata a mensagem seguinte como continuação: uma resposta curta como "aprovado"
pode receber outra skill indicada ou uma ordem de delegação. Ignore-as enquanto o ciclo estiver
em andamento e **não delegue a outro agente** as respostas de portão (o agente delegado não
fala com o usuário). Ao pedir aprovação, sugira ao usuário responder começando com `>>`
(ex.: `>> aprovado`), que faz o roteador deixar a mensagem passar direto.

## Quando usar

- O usuário digitou `/ciclo <demanda>` ou pediu para conduzir a demanda "de ponta a ponta".
- Funcionalidade nova ou correção que pede spec, várias fatias e revisão.

**Quando não usar**

- Correção pequena e bem definida: `/spec` curto (objetivo + critério de aceite + teste) e `/build`.
- Plano já aprovado, só falta executar: `/build auto`.
- Publicar na empresa (release, pacote, `MONTAR-FRONTEND.bat`): `/ship`. O `ciclo` termina
  com a branch integrada em `develop`, não com versão publicada.

## Os 5 estágios

```
[demanda do usuário]
        │
        ▼
Estágio 1  Diagnóstico: contexto, BASE da rede em dia?, ambiguidades, skills e agentes
        │
        ▼
Estágio 2  Spec e plano persistidos em specs/funcionalidades/NNN-nome/
        │
        ▼
🛑 GATE 1  aprovação explícita da spec e do plano
        │
        ▼
Estágio 3  Branch git-flow + fatia por fatia: RED → GREEN → regressão → commit
        │     (autocorreção limitada; paradas obrigatórias continuam valendo)
        ▼
Estágio 4  verificar-tudo no nível certo + revisão por personas + simplificação
        │
        ▼
Estágio 5  Relatório de entrega (com itens 🖥)
        │
        ▼
🛑 GATE 2  "Posso finalizar a branch, fazer o merge em develop e o push?"
        │
        ▼
        ferramentas\finalizar-branch.ps1 -Enviar
```

## Paradas obrigatórias (valem em todos os estágios)

O fluxo autônomo **nunca** passa por cima de `DIRETRIZES-IA.md` › "Quando a IA para e chama o
humano" nem de `specs/PROJETO.md` › Limites › "Perguntar antes". Em resumo:

1. Teste quebrado **sem correção óbvia** (ver a autocorreção no Estágio 3).
2. Spec ambígua, conflito de regra ou decisão de negócio/interface.
3. Ação irreversível ou sensível: migração **estrutural** do banco, qualquer toque no banco de
   produção, `MONTAR-FRONTEND.bat` sem `-Teste`, push na `main`, pacote para a empresa
   (`empacotar-base.ps1`), mudança de contrato (`Início!B7`, `VERSAO-FRONT.txt`, nomes de
   controle, ferramentas MCP, códigos `CT-`/`AT-`), ferramenta nova de desenvolvimento.

Nesses casos, pare, explique e entregue o comando pronto para o humano decidir/rodar.
Se as listas daquelas fontes mudarem, elas prevalecem sobre este resumo.

## Detalhamento

### Estágio 1 — Diagnóstico

1. **Contexto, nesta ordem:** `CLAUDE.md` → `specs/PROJETO.md` → `specs/stacks/<stack>.md` das
   camadas tocadas → `doc_dev/planejamento/PLANO.md` (decisões, correções, pendências).
2. **BASE da rede em dia?** A versão da rede tem prioridade (`PROJETO.md` › Limites, D15).
   Se nem o usuário nem o `PLANO.md` indicarem sincronização recente, pergunte se houve ajuste
   na empresa desde o último ciclo; se houve, a BASE de lá volta ao repositório **antes** de
   qualquer mudança (`trazer-da-rede.ps1` ou comparação em três vias).
3. **Ambiguidades:** requisito ausente ou conflito de regra → `interview-me` ou perguntas
   diretas (1–2 por vez, pt-BR), antes de escrever a spec.
4. **Mapa de skills e agentes do ciclo** (apresentar no GATE 1):

   | Estágio | Skills / agentes |
   |---|---|
   | Spec | `spec-driven-development` (`interview-me` se vago) |
   | Plano | `planning-and-task-breakdown` |
   | Construção | `incremental-implementation`, `test-driven-development`; `source-driven-development` para recurso de VBA/ACE/PowerShell/MCP (referências offline em `doc_dev/planejamento/referencias/`) |
   | Falhas | `debugging-and-error-recovery` |
   | Verificação | `ferramentas\verificar-tudo.ps1` |
   | Revisão | agentes `code-reviewer`, `security-auditor`, `test-engineer`; skill `code-simplification` |
   | Entrega | `git-workflow-and-versioning`; `shipping-and-launch` só no nível "integrar" |

### Estágio 2 — Spec e plano persistidos

1. Pasta `specs/funcionalidades/NNN-nome/` (NNN = último número + 1), a partir de `_modelo/`
   (ver `specs/funcionalidades/README.md`).
2. Artefatos:
   - `SPEC.md` — problema, critérios de aceite objetivos, escopo, limites e referências às
     specs de stack.
   - `plan.md` — fatias verticais finas, dependências, teste que falha primeiro de cada fatia
     (tabela "TDD por stack" de `PROJETO.md`) e itens que só se conferem na estação (`🖥`).
   - `todo.md` — caixas de checagem, uma por tarefa.

   Os arquivos ficam **sem commit** até a branch existir: sem versionar, eles acompanham o
   `git flow feature start` e entram no primeiro commit dela (nada se commita em `develop`).
3. Mostre no chat: diagnóstico, mapa de skills/agentes e o checklist das fatias.
4. **🛑 GATE 1** — aguarde aprovação explícita ("aprovado", "pode iniciar"). Antes disso, não
   crie branch nem escreva código. Pedido de ajuste → revise os arquivos e volte ao GATE 1.

### Estágio 3 — Execução fatia por fatia

1. **Branch:** `git flow feature start <nome>` (correção: `git flow bugfix start <nome>`).
   O primeiro commit da branch leva `SPEC.md`, `plan.md` e `todo.md`.
2. **Para cada fatia do `plan.md`:**
   1. **RED** — o teste que falha primeiro, pela tabela "TDD por stack" de `PROJETO.md` (se
      ela mudar, vale ela):

      | Mudança em | Teste que falha primeiro |
      |---|---|
      | Lógica VBA pura | `Confere` em `fonte/modulos/modAutoteste.bas` (`-Completo`) |
      | Regra estática VBA/esquema/encoding | checagem em `ferramentas/verificacao/verificar.py` |
      | Build, versão, pacote, B7 | `Conferir` em `testes/testar-build-lib.ps1` |
      | Banco / SQL / MCP de dados | caso em `IA/teste/testar-mcp.ps1` |
      | Migração de banco nova | caso que a aplica no banco de teste (`criar-banco-teste.ps1`) e confere o resultado (`specs/stacks/banco-access.md` › Testes) |
      | MCP de desenvolvimento | caso em `IA/teste/testar-mcp-dev.ps1` |
      | Grades e Painel com dados | `testes/testar-uso.ps1`, rodado no build (`-Completo`) |
      | Aparência | prévia (`gerar-visual.ps1`) + item 🖥 |
      | Ferramenta de desenvolvimento (`agent-config/`) | `roteador/testar_roteador.py` ou checagem no `verificar-tudo` |

      Confirme que ele **falha** pelo motivo esperado.
   2. **GREEN** — o código mínimo que o faz passar.
   3. **Regressão** — `ferramentas\verificar-tudo.ps1 -Rapido` (e o teste específico da stack).
   4. **Commit** — só os arquivos da fatia, mensagem em pt-BR com o porquê e as linhas de
      atribuição da sessão; marque a tarefa no `todo.md`.
   5. Só passe para a próxima fatia com a anterior verde e commitada.
3. **Autocorreção limitada:**
   - Erro com causa óbvia (sintaxe, encoding, caminho, nome trocado, teste mal escrito na
     própria fatia): corrija sozinho com `debugging-and-error-recovery` e reexecute, sem
     interromper o usuário — **até 3 tentativas por falha**.
   - Sem causa óbvia, ou estourou o limite: **pare e pergunte**, mostrando o erro, o que foi
     tentado e a hipótese atual.
   - **Nunca** fique verde pulando, apagando ou enfraquecendo teste, checagem do `verificar.py`,
     `Confere`/`Conferir` ou nível do `verificar-tudo`.
4. **Itens 🖥** (Excel 2019 / rede) não param o fluxo: implemente o que dá para provar aqui e
   anote o item para o relatório do Estágio 5.

### Estágio 4 — Verificação e revisão

1. **Verificação no nível da definição de pronto** (`PROJETO.md`):
   ```powershell
   powershell -File ferramentas\verificar-tudo.ps1            # padrão
   powershell -File ferramentas\verificar-tudo.ps1 -Completo  # se tocou fonte\, build\ ou layout
   ```
   A saída sem falha é a prova; ela vai no relatório.
2. **Revisão por personas** (agentes de `.claude/agents/`; podem rodar em paralelo, cada um
   com o diff da branch e a `SPEC.md`):
   - `code-reviewer` — cinco eixos + seção **Revisão** das specs de stack tocadas; nomes e
     comentários em pt-BR (`PROJETO.md` › Estilo); sem código morto.
   - `security-auditor` — SQL parametrizado, data como parâmetro, versão otimista, log na
     mesma transação (`modDB.LogEm`), nenhum dado real (repositório público), 100% offline.
   - `test-engineer` — os testes provam o comportamento? casos de borda, falsos positivos,
     checagem contornada.
   - `code-simplification` — reduzir sem mudar comportamento, respeitando `simplify-ignore` e
     as decisões registradas.
3. **Achados:** críticos e importantes são corrigidos (cada correção com teste e commit) e o
   `verificar-tudo` roda de novo; opcionais vão para o relatório para o usuário decidir.

### Estágio 5 — Relatório e GATE 2

1. **Antes do relatório**, feche a definição de pronto, no último commit da branch:
   - spec de stack atualizada se a stack mudou; `PLANO.md` se mudou fase/decisão/pendência;
   - `todo.md` todo marcado;
   - carimbo no topo da `SPEC.md`: `> Concluída em dd/mm/aaaa — feature/<nome>` (o nome da
     branch identifica o merge; a versão publicada entra quando o `/ship` publicar).
2. **Relatório de entrega:**
   - resumo das mudanças e arquivos criados/alterados;
   - saída do `verificar-tudo` (nível usado);
   - parecer das personas e o que foi corrigido;
   - branch e lista de commits;
   - **itens 🖥** que o usuário precisa conferir na estação — "pronto aqui" não é "pronto na
     empresa".
3. **🛑 GATE 2** — pergunte:
   > *"Tudo desenvolvido, testado e revisado. Posso finalizar a branch, fazer o merge em
   > `develop` e o push para o repositório remoto?"*
4. **Com a autorização:**
   ```powershell
   powershell -File ferramentas\finalizar-branch.ps1 -Enviar
   ```
   Confirme o merge `--no-ff` em `develop`, a remoção da branch (local e remota) e o push.

## Exemplos

```
/ciclo adicionar o campo "segmento" (lista) na ficha de clientes
/ciclo corrigir o filtro da grade de oportunidades que ignora a situação "Perdida"
/ciclo nova ferramenta MCP para listar oportunidades paradas há mais de 30 dias
/ciclo renomear a coluna "obs" de contatos para "observacao" (migração de banco)
/ciclo avisar no feed quando um atendimento mudar de responsável
```

No exemplo da migração, renomear coluna é migração **estrutural**: o ciclo planeja e testa no
banco de teste, mas para e pergunta antes de qualquer coisa que toque o banco de produção, e o
MCP de dados precisa acompanhar `config.versao_esquema`. No do feed, a mudança passa por
`modFeed`/`log_alteracoes`: toda gravação continua com log na mesma transação.

## Racionalizações comuns

| Racionalização | Realidade |
|---|---|
| "O teste está atrapalhando, vou relaxar a asserção" | Enfraquecer o teste é proibido; pare e pergunte. |
| "É só mais uma tentativa" (4ª, 5ª...) | Passou de 3 sem causa óbvia: pare e mostre o que sabe. |
| "A spec está clara o bastante, dispenso o GATE 1" | O GATE 1 é obrigatório, mesmo para spec curta. |
| "verificar-tudo -Rapido basta" | O nível final segue a definição de pronto (padrão ou `-Completo`). |
| "O item 🖥 passou aqui no 365" | Esta máquina não é a estação com Excel 2019; liste o item. |

## Sinais de alerta

- Código escrito antes do GATE 1 ou commit em `develop`/`main`.
- Fatia sem teste que falhou antes.
- Mudança em teste/checagem no mesmo commit que "conserta" a falha que ele apontava.
- Relatório sem a saída do `verificar-tudo` ou sem a lista 🖥.
- Toque em contrato (B7, nomes de controle, ferramentas MCP, `CT-`/`AT-`) sem pergunta.
- Resposta de portão delegada a outro agente ou trocada por outra skill indicada pelo roteador.

## Verificação

- [ ] GATE 1 aprovado antes de qualquer branch ou código.
- [ ] Cada fatia: teste falhou → passou → `-Rapido` verde → commit.
- [ ] `verificar-tudo` no nível certo, sem falha, com a saída no relatório.
- [ ] Revisão das três personas + simplificação, achados críticos/importantes resolvidos.
- [ ] Specs de stack, `PLANO.md` e `todo.md` em dia; `SPEC.md` carimbada; itens 🖥 listados.
- [ ] GATE 2 autorizado antes do `finalizar-branch.ps1 -Enviar`.
