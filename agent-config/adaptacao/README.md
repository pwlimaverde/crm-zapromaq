# Adaptação do agent-skills ao CRM Zapromaq

O projeto [addyosmani/agent-skills](https://github.com/addyosmani/agent-skills) (licença MIT)
está **na íntegra** em `agent-config/vendor/agent-skills`, como *git subtree*. Ele nunca é
editado aqui: assim cada versão nova dele entra com um comando, sem conflito.

O que é deste projeto fica nesta pasta e é **somado** ao original:

| Arquivo | Vai para |
|---|---|
| `skills/<nome>.md` | fim de `.claude/skills/<nome>/SKILL.md` |
| `commands/<nome>.md` | fim de `.claude/commands/<nome>.md` |
| `commands/_todos.md` | fim de **todos** os comandos |
| `agents/<nome>.md` | fim de `.claude/agents/<nome>.md` |
| `excluidos.txt` | itens do original que não são instalados |
| `proprios/{skills,commands,agents}/` | itens só deste projeto, copiados como estão (ex.: `/roteador`) |

`agent-config/.claude/` (o que o Claude Code lê) é **gerado** por
`ferramentas/instalar-skills.ps1`. Não edite lá: edite aqui e rode o script.
No texto gerado, `agent-skills:<skill>` (nome de plugin) vira `<skill>`.
O **Antigravity** espelha e consome este mesmo ecossistema por referência via
`GEMINI.md` e `.agents/skills.json`, mantendo uma única fonte da verdade e zero
arquivos duplicados.

## Princípio

O original diz **como** trabalhar (spec → plano → TDD → revisão → entrega). A adaptação só
diz **onde e com quê** neste projeto: caminhos, comandos de teste, git-flow, restrições
inegociáveis. O conhecimento do projeto não se repete nos trechos: mora em
`agent-config/specs/` (projeto e stacks), e os trechos apontam para lá.

## Comandos

```powershell
powershell -File agent-config\ferramentas\instalar-skills.ps1            # regera .claude
powershell -File agent-config\ferramentas\instalar-skills.ps1 -Conferir  # .claude em dia? (o verificar-tudo roda)
powershell -File agent-config\ferramentas\atualizar-agent-skills.ps1     # traz a versão nova do original (numa branch)
```

## Ao atualizar o original

`atualizar-agent-skills.ps1` lista o que mudou e avisa quando mudou algo que tem adaptação.
Releia essas skills: se o original passou a cobrir o que o trecho daqui dizia, enxugue o
trecho; se renomeou ou removeu, o `instalar-skills` para e aponta o arquivo órfão.

## O que não foi instalado e por quê

Ver `excluidos.txt`. Os ganchos (`vendor/agent-skills/hooks`) também não estão ligados:
o `simplify-ignore.sh` reescreve o arquivo no lugar a cada leitura — num fonte VBA em
Windows-1252 + CRLF isso arrisca corromper acentos e fim de linha. Os marcadores
`simplify-ignore` valem por instrução (ver `skills/code-simplification.md`).
