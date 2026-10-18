# Plano: Backup coordenado e alerta (rascunho)

Spec: [SPEC.md](SPEC.md)

## Fatias (teste que falha primeiro, depois prova)

- trava de backup com dono e sessão e recuperação de trava abandonada (tpr, dois processos)
- nome único, arquivo temporário promovido só depois de conferido, log por execução
- manifesto em toda cópia; retenção nunca apaga a última validada
- ver_log backup lê o backup.log (testar-mcp-dev)
- último backup validado em crm_status, estado_sistema e Início; alerta acima de 36 h
- config.geracao_base renovada ao restaurar

## Depende de

004 F8