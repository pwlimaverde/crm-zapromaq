# Spec: Backup coordenado e alerta (ciclo 006)

> Situação: rascunho (planejado no ciclo 004; ainda não aprovado nem executado)

## Objetivo

Tratar os itens DEV-02, DEV-21, DEV-13 (geração) da revisão de 18/10/2026. Depende de: 004 F8.

## Critérios de aceite

- [ ] trava de backup com dono e sessão e recuperação de trava abandonada (tpr, dois processos)
- [ ] nome único, arquivo temporário promovido só depois de conferido, log por execução
- [ ] manifesto em toda cópia; retenção nunca apaga a última validada
- [ ] ver_log backup lê o backup.log (testar-mcp-dev)
- [ ] último backup validado em crm_status, estado_sistema e Início; alerta acima de 36 h
- [ ] config.geracao_base renovada ao restaurar

## Decisões já assumidas

As decisões DA-0 a DA-12 e Dec 1 a 10 estão em `..\004-preparacao-da-rede\SPEC.md`.

## Limites

- Sempre: teste primeiro; teste que grava só em banco TESTE; nada de dado real no repositório.
- Nunca: publicar na rede nem tocar o banco de produção daqui.

## Questões em aberto

Nenhuma conhecida; revisar a spec ao iniciar o ciclo.