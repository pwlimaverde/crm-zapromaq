# Spec: Regras iguais nos três canais (ciclo 009)

> Concluída em 21/10/2026 - feature/regras-entre-canais (plano revisado de 19/10/2026, seção 5, Etapa 2)

## Objetivo

- **DEV-10** — matriz de campos versionada, compartilhada pela ficha, pelo lote e pelo conector; normalização da
  seção 5.1 (maiúsculas em todo texto comercial, com as exceções); transições de etapa (encerrar em Perdido ou
  Descartado limpa a próxima ação; reabrir limpa motivo, desfecho e retomada e orienta "use Duplicar/novo
  atendimento"); `retomar_em` só em Perdido e Descartado; `dt_entrada` = data do lançamento na criação; proposta e
  desfecho anteriores à entrada viram aviso; contato homônimo na mesma empresa gera aviso. Migração de dados de
  normalização com prévia (origem MIGRACAO), aplicada pela montagem.
- **DEV-11** — contato fixo por empresa: a ficha não troca a empresa de contato gravado; o conector remove
  `codigo_cliente_destino` de `atualizar_contato` (nota de mudança); na transação de criar/alterar atendimento o
  contato tem de existir, estar ativo e ser da empresa do atendimento (front e conector).
- **DEV-12** — lote: exatamente 11 colunas; CNPJ com dígito inválido é ERRO; UF, segmento e qualificação
  conferidos contra as listas; notas inteiras 1–3 pelo conversor do DEV-09; limites de tamanho; mesmo validador
  da ficha antes do INSERT; identificador de lote e de linha no detalhe do log (estável no reprocessamento).
- **DEV-15** — versão lida nas ações (desativar, reativar, excluir) na ficha; reativar a empresa reativa só a
  empresa (ficha e conector, descrição de `alterar_situacao_cliente` atualizada); conflito na ficha mostra, por
  campo, o valor lido, o atual e o digitado, sem perder a digitação.

## Critérios de aceite

- [x] Matriz `SISTEMA\DADOS\docs\matriz-campos.json` (versão, campo, tabela, tipo, tamanho, normalização, lista,
  obrigatoriedade, editabilidade) e o `verificar.py` confere que ficha (`modSchema`), conector (`TamanhoTexto`,
  defs) e DDL concordam com ela.
- [x] Casos compartilhados (front `Confere` + conector `testar-mcp`, cobrança do `verificar.py`):
  C22 texto comercial em maiúsculas, com acento, espaço duplo colapsado e quebra de linha preservada;
  C23 e-mail minúsculo e um só endereço (o segundo vai para observações);
  C24 URL, caminho e e-mail dentro das observações preservados;
  C25 encerrar em Perdido/Descartado limpa `prox_acao` e `dt_prox_acao`;
  C26 `retomar_em` fora de Perdido/Descartado é recusada;
  C27 reabrir limpa motivo, data de desfecho e retomada e devolve a orientação;
  C28 proposta/desfecho antes da entrada é aviso, não recusa;
  C29 contato de outra empresa ou inativo é recusado no atendimento;
  C30 lote: CNPJ com DV inválido é ERRO;
  C31 lote/nota fracionária e UF fora da lista são ERRO;
  e, sem número de caso (só o conector tem o fluxo testável sem tela), reativar a empresa não reativa os contatos (DEV-15).
- [x] Ficha: `retomar_em` recusado ao salvar fora de Perdido/Descartado (C26); `dt_entrada` nasce com a data do lançamento nos três canais.
- [x] Conector: `atualizar_contato` sem `codigo_cliente_destino` (nota de mudança); `alterar_situacao_cliente`
  reativa só a empresa; transições e avisos iguais ao front; normalização da seção 5.1 em todo texto.
- [x] Migração `005-normalizacao-secao-5-1.ps1` (só dados): normaliza os textos existentes pela mesma regra, grava
  cada diferença no log com origem MIGRACAO e a prévia no relatório; testada em banco fictício com casos legados.
- [x] Ações da ficha com versão lida: desativar/reativar/excluir com ficha antiga (versão mudou) são recusados sem
  mudar o registro nem os dependentes.
- [x] Conflito na ficha: lista campo, lido, atual e digitado; o usuário escolhe recarregar ou manter a digitação.

## Decisões assumidas

| # | Decisão | Premissa | O que mudaria |
|---|---|---|---|
| 009-1 | "Texto comercial" = todo campo T/M que não seja e-mail, CNPJ, telefone, lista, `ctx_arquivo`. Inclui nome, cargo, observações, próxima ação, `ctx_resumo`. Palavras que são URL (`://`, `www.`), caminho (`\`) ou e-mail (`@`) ficam como digitadas. | Seção 5.1 do plano, que prevalece sobre "nomes preservam digitação" (D24). | Se o comercial quiser nomes como digitados, `nome` sai da lista. |
| 009-2 | E-mail com mais de um endereço: o primeiro fica no campo, os demais vão para observações com o prefixo `E-MAIL ADICIONAL:`. | Seção 5.1. | — |
| 009-3 | Reabrir (sair de etapa que encerra para etapa aberta) limpa motivo, desfecho e retomada sem perguntar e mostra a orientação ("reabrir é para corrigir encerramento indevido; interesse novo é novo atendimento - use Duplicar"). | Seção 5.2 ("permitida, sem bloqueio"). | — |
| 009-4 | `dt_entrada` nasce com a data do lançamento nos três canais e continua editável (correção de lançamento); legado sem data continua sem data. A ficha não desabilita `retomar_em` ao trocar a etapa (os controles são criados em tempo de execução, sem evento): a regra vale no domínio, ao salvar, nos três canais. | Seção 5.2 sem retropreenchimento; o plano não manda travar a entrada. | Travar a entrada exige só `editavel = 0` no `modSchema` e o INSERT passar a data. |
| 009-5 | Migração de normalização é a 005 (só dados), aplicada pela montagem depois da 004 estrutural (ciclo 010). | Numeração pedida para o ciclo 010 (004 = esquema 1.2). | — |
| 009-6 | Identificador de lote = `LOTE-AAAAMMDD-HHMMSS-<estação>`; linha = número da linha; entra no detalhe do log de inserção. Reprocessar a mesma colagem cai em DUPLICADO por CNPJ ou SUSPEITO por nome (não duplica). | Idempotência completa vem com o identificador de operação (ciclo 010). | — |
