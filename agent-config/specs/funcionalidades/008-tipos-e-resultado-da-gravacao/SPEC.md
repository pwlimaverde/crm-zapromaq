# Spec: Tipos e resultado da gravação (ciclo 008)

> Situação: rascunho (planejado no ciclo 004; ainda não aprovado nem executado)

## Objetivo

Tratar os itens DEV-05, DEV-08 (estados), DEV-22 da revisão de 18/10/2026. Depende de: 005.

## Critérios de aceite

- [ ] parse que distingue vazio, válido e inválido (matriz 31/02, abc, 12.50, 1,5)
- [ ] estados não enviada / falhou / confirmada / indeterminada, sem "nada foi gravado" falso
- [ ] erro nativo e repetição só antes do commit (testar-mcp com falha injetada)

## Decisões já assumidas

As decisões DA-0 a DA-12 e Dec 1 a 10 estão em `..\004-preparacao-da-rede\SPEC.md`.

## Limites

- Sempre: teste primeiro; teste que grava só em banco TESTE; nada de dado real no repositório.
- Nunca: publicar na rede nem tocar o banco de produção daqui.

## Questões em aberto

Nenhuma conhecida; revisar a spec ao iniciar o ciclo.