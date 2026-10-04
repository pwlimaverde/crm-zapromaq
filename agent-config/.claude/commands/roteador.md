---
description: "Liga, desliga ou mostra o roteador de modelo e o seletor de skill com o Jev (on | skills | off | status)"
disable-model-invocation: true
argument-hint: on | skills | off | status
allowed-tools: Bash(py -3:*)
---

!`py -3 "${CLAUDE_PROJECT_DIR:-.}/roteador/roteador.py" --cli $ARGUMENTS`

Mostre ao usuário a saída acima, sem comentar nem executar mais nada.
