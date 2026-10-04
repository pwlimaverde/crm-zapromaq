# Diretrizes de engenharia com IA

Protocolo de trabalho deste projeto, baseado em
[`addyosmani/agent-skills`](https://github.com/addyosmani/agent-skills) (*Beyond Vibe Coding*).
O original está em `vendor/agent-skills`; o que é deste projeto, em `adaptacao/` e `specs/`.

## Objetivo

Acabar com o ciclo de quebras do desenvolvimento sem estrutura. A IA não pega atalho, não gera
código antes de entender e não decide sozinha o que é ambíguo. O trabalho anda em **fatias
pequenas, com teste primeiro, especificação clara, verificação real e git disciplinado**.

## Os quatro pilares

1. **Contexto antes do código.** A IA trabalha dentro do contexto explícito: `CLAUDE.md` →
   `specs/PROJETO.md` → `specs/stacks/<stack>.md` → spec da funcionalidade.
2. **Spec antes de implementar.** Nada de funcionalidade sem objetivo, critérios de aceite,
   escopo e limites escritos (`specs/funcionalidades/NNN-nome/SPEC.md`).
3. **Fatias pequenas, commits atômicos.** Uma tarefa por vez; cada avanço comprovado vira um
   commit na branch do git-flow. O commit é o botão de desfazer.
4. **Verificação real.** "Parece certo" não conta. Pronto = teste novo passando +
   `ferramentas/verificar-tudo.ps1` sem falha (com a saída como prova) — ver a definição de
   pronto em `specs/PROJETO.md`.

## O ciclo

```
DEFINIR      PLANEJAR     CONSTRUIR    VERIFICAR    REVISAR      SIMPLIFICAR      ENTREGAR
/spec   ->   /plan   ->   /build  ->   /test   ->   /review  ->  /code-simplify -> /ship
```

| Comando | O que faz aqui | Produz |
|---|---|---|
| `/ciclo` | orquestrador de ponta a ponta: diagnóstico, plano, TDD autônomo com autocorreção, revisão modular e portão de entrega | fluxo completo com gates |
| `/spec` | entrevista (1–2 perguntas por vez, pt-BR) até ficar claro | `specs/funcionalidades/NNN-nome/SPEC.md` |
| `/plan` | fatias verticais com teste e verificação por tarefa; 🖥 = conferir na estação | `plan.md`, `todo.md` na mesma pasta |
| `/build` | abre a branch (`git flow feature start`), uma tarefa: teste que falha → código mínimo → suíte → commit | commits atômicos |
| `/build auto` | todas as tarefas após **uma** aprovação do plano; para nas situações abaixo | idem |
| `/test` | `verificar-tudo` (`-Rapido` / padrão / `-Completo`) | prova em texto |
| `/review` | cinco eixos + seção **Revisão** das specs de stack; revisores `code-reviewer`, `security-auditor`, `test-engineer` | achados com `arquivo:linha` |
| `/code-simplify` | reduz complexidade sem mudar comportamento; respeita `simplify-ignore` e as decisões registradas | commits |
| `/ship` | integrar (`finalizar-branch.ps1 -Enviar`) ou publicar na empresa (release + pacote) | merge, tag `vX.Y` |
| `/constraints` | barra de qualidade (aqui: os níveis do `verificar-tudo`) | — |

### Quando a IA para e chama o humano

1. Teste quebrado sem correção óbvia.
2. Spec ambígua, conflito de regra ou decisão de negócio/interface.
3. Ação irreversível ou sensível: migração **estrutural** do banco, banco de produção, publicar
   (`MONTAR-FRONTEND.bat` sem `-Teste`), push na `main`, pacote para a empresa, mudança de
   contrato (B7, `VERSAO-FRONT.txt`, nomes de controle, ferramentas MCP, códigos `CT-`/`AT-`).
   A IA entrega o comando pronto para o humano rodar.

## Habilidades centrais

- **Engenharia de contexto:** hierarquia `CLAUDE.md` > `specs/PROJETO.md` > spec de stack > spec
  da funcionalidade > conversa. Em dúvida numa sessão longa, reler a spec — não supor.
- **Depuração em 5 passos:** reproduzir → localizar → reduzir a um teste mínimo → corrigir →
  proteger com teste de regressão. Antes de "corrigir" algo estranho, procurar o motivo em
  `docs/legado/estado-migracao.md` e `PLANO.md`.
- **Fonte oficial:** documentação da **versão exata** (Excel 2019, ACE 16, PowerShell 5.1, MCP)
  em `doc_dev/planejamento/referencias/`, offline — não a memória do modelo.

## Responsabilidades

| Humano | IA |
|---|---|
| Problema de negócio e restrições | Conduzir a entrevista e redigir a spec |
| Aprovar o plano | Quebrar em fatias com dependências |
| Decidir o ambíguo e o irreversível | Escrever o teste antes do código |
| Exigir a prova (saída do `verificar-tudo`) | Rodar comandos, testes, commits atômicos e revisões |
| Conferir os itens 🖥 na estação/empresa | Listar o que só se confere lá |

## Começar uma tarefa

> "Siga as diretrizes do projeto e execute `/spec` para a nova funcionalidade: …"

Para correção pequena e bem definida: `/spec` curto (objetivo + critério de aceite + teste) e
`/build` direto.
