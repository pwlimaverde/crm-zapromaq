# Plano revisado de desenvolvimento (19/10/2026): sequência de execução

Documento-base de aceite: `C:\Users\pwlim\.claude\uploads\20e12264-dd94-4437-9e8d-890aee46c5e7\9c9d743c-PLANO-DESENVOLVIMENTO-CRM-PW_1.md` (DEV-01 a DEV-25, etapas 0 a 5, seções 5, 7, 8, 9 e 10). Este arquivo é a sequência de execução e as decisões assumidas. Ele substitui o plano antigo (DEV-01 a DEV-28, ciclos 004 a 014).

## Estado de partida

- `develop` = `70b0847`, com os ciclos 004, 005, 006, 007, 008 e 009 integrados. `origin/develop` ainda está em `c955a1e`: fazer `git push origin develop` depois de cada ciclo concluído.
- Ciclo 010 (`feature/esquema-1-2`): commits de `8d0c477` até `8195a76` (spec, migração 004, atendimento anterior, log de auditoria, versão do contexto, identificador de operação). Pode haver alterações não commitadas: se houver, commitar por tema antes de qualquer outra coisa.
- A versão do front continua **1.6**: a publicação só será feita depois de tudo concluído. Não publicar nada na rede daqui.

## Sequência

Uma branch `feature/<nome>` por ciclo, a partir da `develop` atualizada. Cada ciclo: spec (SPEC.md, plan.md, todo.md) a partir de `specs\funcionalidades\_modelo\`; teste que falha primeiro; código; `verificar-tudo.ps1 -Rapido`; commit em pt-BR; ao fim, `verificar-tudo.ps1 -Completo` sem falha; `finalizar-branch.ps1` **sem** `-Enviar`; `git push origin develop`.

| Ciclo | Itens | Conteúdo |
|---|---|---|
| 010 | DEV-06 B, DEV-13, DEV-16, DEV-17 | Esquema 1.2: migração 004, controle único, atendimento anterior sem ciclo, versão do contexto, identificador de operação. Fechar: verificar, finalizar, push. |
| 011 | DEV-18, DEV-19, DEV-20 | Feed: cursor só avança depois de aplicar; dependências e virada do dia; filtro e ordem reaplicados. |
| 012 | DEV-21, DEV-22, DEV-23 | Agenda pela `dt_prox_acao`; "retomadas a avaliar"; paginação com corte estável; extratos só em `EXECUCAO\extratos`; roteiro CSV fora de Documentos. |
| 013 | DEV-24 e matriz de consumidores (DEV-07) | Texto literal em grade e CSV neutralizado; skills da seção 6 do documento-base. Sem fila de exceções. |
| 014 | DEV-25 | Medir compactação numa cópia; adotar só com ganho. |
| 015 | DEV-26 | Lote atualiza duplicados por CNPJ: botão em lote e ação unitária por duplo clique; campo vazio não apaga; observações acrescentadas. Depende de DEV-12, DEV-09 e DEV-10. |
| 016 | DEV-27 | Rota lista todos os clientes do raio; "Gerar visitas" só aceita atendimento lançado; atalho para criar atendimento; `listar_candidatos` aditivo com `sem_atendimento`. Depende de DEV-11. |
| 017 | DEV-28 | Duplicar atendimento: `DadosParaDuplicar` (copia cliente, responsável, origem, prioridade, categoria, tipo_venda, máquina, família, última interação; zera valor, orçamento, proposta, agenda, desfecho, motivo, retomada, tentativas; `dt_entrada` = hoje); `id_atendimento_anterior` aponta o pai; nada grava até Salvar; conector `duplicar_atendimento` aditivo. Depende de DEV-11 e DEV-16. |

Depois dos 17 ciclos:
1. Revisão final de tudo o que foi implementado nesta rodada (skill `code-review-and-quality`; nos cinco eixos: correção, legibilidade, arquitetura, segurança e desempenho). Corrigir o que for real, rodar `verificar-tudo.ps1 -Completo` de novo, commit e merge.
2. `git push origin develop`.
3. Relatório final em texto: por ciclo, branch, hashes, saída do `verificar-tudo -Completo`, decisões assumidas novas, e itens 🖥 para o usuário rodar `MONTAR-FRONTEND.bat -Teste` e os testes práticos.

## Decisões assumidas (substituem perguntas)

1. DEV-26, DEV-27 e DEV-28 entram no fim da numeração; DEV-01 a DEV-25 não são renumerados.
2. Duplicado no lote só por CNPJ; nome parecido continua SUSPEITO e não é atualizado. Com CNPJ repetido, vale a primeira linha.
3. Campo vazio no lote não apaga o valor existente.
4. Observações do lote são acrescentadas, nunca substituídas.
5. Nome da empresa pelo lote substitui o cadastrado quando o CNPJ bate.
6. Empresa inativa não é atualizada pelo lote.
7. Lote grava linha a linha, cada uma em sua transação, com resultado por linha.
8. Duplo clique em duplicado pede confirmação mostrando a diferença antes de atualizar.
9. Ação de log "ALTERAR LOTE" para separar a origem na auditoria.
10. "Próxima etapa" do planejamento de rota = Gerar visitas; é a única passagem que grava e exige atendimento.
11. A rota inclui pré-clientes e exclui empresas inativas.
12. O atendimento criado pelo atalho volta marcado na rota.
13. O atalho liga o anterior só em Perdido ou Descartado com `retomar_em`.
14. `sem_atendimento` é campo novo do `listar_candidatos`, limitado a 200 linhas (aditivo).
15. Duplicar não grava nada até Salvar.
16. Na duplicação, valor, orçamento, data da proposta e próxima ação são zerados.
17. Vínculo da duplicação usa `id_atendimento_anterior`, sem migração nova.
18. Duplicar atendimento encerrado é permitido, com etapa inicial.
19. Conector ganha `duplicar_atendimento` (aditivo); a regra da cópia fica num lugar só, com teste.
20. Seções 5 do documento-base prevalecem sobre Dec 4, 8, 9, 10, DA-2 e as decisões do ciclo 005.
21. Itens 🖥 não bloqueiam os ciclos; ficam listados para o teste prático do usuário.

## Restrições (valem para todos os ciclos)

- Nada de dado real no repositório: exemplos fictícios.
- Nada publicado na rede; nenhum teste grava fora de banco TESTE; não alterar `BASE\crm_zapromaq.accdb` nem `execucao\` da raiz.
- Contrato MCP só com mudança aditiva ou nota de mudança; `VersaoServidor` coerente.
- Fontes VBA com acento só por `agent-config\ferramentas\editar-vba.ps1` ou Edit/Write; cp1252 + CRLF; `.ps1` em UTF-8 com BOM.
- Excel 2019 nas estações; front 1366×768; nada de recurso do Microsoft 365.
- Commits terminam com `Co-Authored-By: Claude Haiku 5.5 <noreply@anthropic.com>` e `Claude-Session: https://claude.ai/code/session_01QUHkJL5vKtWC5PbebLNVhb`.
- Não rodar o roteiro da estação (seção 7 do plano antigo / `implantacao.md`).
- Processos: só encerrar processos que a própria execução iniciou (Excel e PowerShell de teste). O processo PowerShell PID 5072 pertence a outra sessão: não tocar.
