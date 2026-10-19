# Spec: Homologação e diagnóstico mínimo (ciclo 005)

> Situação: rascunho (planejado no ciclo 004; ainda não aprovado nem executado)

## Objetivo

Tratar os itens DEV-01 (resto), DEV-25 (diagnóstico), DEV-26 da revisão de 18/10/2026. Depende de: 004.

## Critérios de aceite

- [ ] execucao\homologacao\ e config.homologacao, obrigatórios para testes que gravam (tpr)
- [ ] equivalência de caminho canônico UNC x letra x relativo (tbl)
- [ ] CaminhoBanco sem fallback silencioso (Confere no modAutoteste)
- [ ] diagnóstico do MCP com estação, usuário técnico e ferramenta, um arquivo por estação (testar-mcp)
- [ ] checar-ambiente grava o relatório e verifica conector, IA e OneDrive (tbl)

## Decisões já assumidas

As decisões DA-0 a DA-12 e Dec 1 a 10 estão em `..\004-preparacao-da-rede\SPEC.md`.

## Limites

- Sempre: teste primeiro; teste que grava só em banco TESTE; nada de dado real no repositório.
- Nunca: publicar na rede nem tocar o banco de produção daqui.

## Questões em aberto

Nenhuma conhecida; revisar a spec ao iniciar o ciclo.