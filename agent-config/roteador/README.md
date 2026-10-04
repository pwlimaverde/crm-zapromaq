# Roteador de modelo e seletor de skill (Jev)

Adaptado do roteador do `smart-core-assistant-v2` (`smart-agent-config/.claude/hooks/roteador.py`).
Um gancho `UserPromptSubmit` do Claude Code (Python só com biblioteca padrão) faz **uma**
chamada ao **Jev** (TypeSafe System One) por mensagem e decide em código:

- **skill** do agent-skills instalada em `.claude/skills` (ou nenhuma);
- **nível** (simples/rotina/difícil) → modelo do subagente (haiku/sonnet/opus);
- **especialista**: `code-reviewer`, `security-auditor`, `test-engineer` (de `.claude/agents`)
  ou os nativos `general-purpose`, `Plan`, `Explore`;
- **risco**, **continuação** da tarefa anterior e se o **entregável é texto**.

| Modo (`/roteador <modo>`) | O que faz |
|---|---|
| `skills` (padrão) | Só o seletor: indica a skill ("invoque a skill X"); não repete na mesma sessão. |
| `on` | Seletor + delegação: a sessão principal (no Haiku, `/model haiku`) delega ao especialista com o `model` do nível. |
| `off` | Nada. |

`/roteador status` mostra o modo, se a chave chega ao gancho, o catálogo e a última decisão.
A linha de status do Claude Code mostra a decisão de cada mensagem.

**Atalhos** (no início da mensagem): `>>` passa direto; `#rapido`, `#padrao`, `#profundo`
forçam o nível; `$<skill>` força a skill. Comandos (`/spec`, `/plan`, `/build`...) não são
roteados: o próprio comando já invoca a skill.

## Arquivos

```
roteador/
  roteador.py             o gancho (também --cli on|skills|off|status)
  roteador.config.json    política: limiares, níveis, pares concorrentes, mapas, descrição do projeto
  testar_roteador.py      testes SEM rede (rodam no verificar-tudo -Rapido)
  avaliar_roteador.py     calibração COM o Jev: matriz de nível, acerto de skill, latência
  avaliar_casos.json      casos da calibração (fictícios)
  estado/                 GERADO, fora do git: modo, decisoes.jsonl, status.txt, catalogo.json
.claude/settings.json     registra o gancho e a linha de status
adaptacao/proprios/commands/roteador.md   o /roteador (instalado em .claude/commands pelo instalar-skills)
```

## Política (`roteador.config.json` — ajuste aqui, não no código)

Risco > 0,7 ou confiança do nível < 0,6 sobe um nível (máximo +1); continuação > 0,55 herda
nível, especialista e skill (nunca desce); skill aceita com chance > 0,5 e confiança ≥ 0,65,
com desempate para pares parecidos (depuração × TDD, spec × plano, revisão × simplificação,
revisão × segurança, migração × implementação) por palavras-gatilho; skill nova aceita vence
a herdada. Falha do Jev (timeout de 4 s, HTTP, chave ausente) → sonnet + `general-purpose`,
sem skill. `using-agent-skills` e `context-engineering` ficam fora da escolha (são meta).

## Chave e privacidade

`TYPESAFE_API_KEY` vem do bloco `env` do `~/.claude/settings.json` global (o Claude Code a
repassa ao gancho). Nunca no repositório nem no log. Antes de sair da máquina (e antes do
log), a mensagem é **mascarada**: segredos, e-mails, IPs e — por ser um CRM — **CNPJ, CPF,
telefone** e sequências longas de dígitos. Só a mensagem mascarada (até 2000 caracteres), a
descrição do projeto e as descrições das skills vão ao Jev; o log guarda 200 caracteres.

## Calibração

```powershell
py -3 agent-config\roteador\avaliar_roteador.py
```

Em 04/10/2026 (jev-1.13.0, 34 casos): skill 33/34 (nenhum falso positivo), nível 28/34
(4 acima do esperado, 2 abaixo), continuação 3/4, Jev ~430 ms (p95 ~600 ms), partida do
processo ~87 ms.

## Contrato do gancho

stdin = JSON do Claude Code (`prompt`, `session_id`); stdout =
`hookSpecificOutput.additionalContext` ou nada; efeitos = uma linha no JSONL e o
`status.txt`. Registrado em *exec form* (`py -3 ${CLAUDE_PROJECT_DIR}/roteador/roteador.py`)
— `py`, não `python3` (alias da Microsoft Store). Nunca derruba a mensagem: qualquer erro
sai em silêncio e aparece como "erro interno" na linha de status. `ROTEADOR_ESTADO` troca a
pasta de estado (os testes usam uma temporária).
