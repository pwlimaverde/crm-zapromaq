# Spec: Esquema 1.2 (ciclo 010)

> Situação: rascunho (planejado no ciclo 004; ainda não aprovado nem executado)

## Objetivo

Tratar os itens DEV-08 (id persistente), DEV-11, DEV-12, DEV-25 (campo) da revisão de 18/10/2026. Depende de: 004, 006, 007.

## Critérios de aceite

- [ ] migração 004 estrutural pela montagem, inventário antes (testar-migracao-004)
- [ ] índice único em controle (só se o inventário achar zero duplicados)
- [ ] id_operacao e ctx_versao
- [ ] usuario_tecnico e estacao no log

## Decisões já assumidas

As decisões DA-0 a DA-12 e Dec 1 a 10 estão em `..\004-preparacao-da-rede\SPEC.md`.

## Limites

- Sempre: teste primeiro; teste que grava só em banco TESTE; nada de dado real no repositório.
- Nunca: publicar na rede nem tocar o banco de produção daqui.

## Questões em aberto

Nenhuma conhecida; revisar a spec ao iniciar o ciclo.