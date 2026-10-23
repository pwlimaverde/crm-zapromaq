# Spec: Feed e tela (ciclo 011)

> Situação: aprovada em 21/10/2026 (plano revisado de 19/10/2026, Etapa 3; itens DEV-18, DEV-19 e DEV-20 do
> documento-base). Sem migração: usa `config.geracao_base` (ciclo 006) e o `log_alteracoes` como está.

## Objetivo

A tela de cada estação acompanha o banco sem perder evento, sem ficar presa depois de uma restauração, sem
mostrar situação de ontem e sem desfazer o filtro e a ordem que o vendedor montou.

- **DEV-18 — feed confiável** (`modFeed`): o cursor (`mUltimoLog`) só avança **depois** de aplicar todas as abas e
  avisos; falha no meio deixa o lote pendente e a próxima rodada reaplica (reescrever linha é idempotente).
  Commit fora de ordem: id do log que faltou na faixa lida vira **lacuna** e é procurado de novo nas rodadas
  seguintes por até 5 minutos (transação do CRM dura segundos; depois disso a lacuna é de transação desfeita).
  Geração da base (`config.geracao_base`) diferente da guardada, ou `MAX(id)` menor que o cursor (banco restaurado):
  cursores e lacunas zerados e as abas carregadas são remontadas.
- **DEV-19 — dependências e virada do dia**: mapa de dependências num lugar só (`modFeed.Dependentes`): cliente →
  contatos, oportunidades e Painel; contato → oportunidades e Painel; oportunidade → Painel; lista → cache das
  combos. Recargas agrupadas por tabela (uma consulta de ids dependentes por tabela, um `AtualizarRegistros` por
  aba). Virada do dia sem gravação: as abas carregadas são remontadas (situação e dias parado). Ficha em edição
  continua só com aviso (`AvisoExterno`). Separados: última consulta ao banco (`modConfig`), última carga efetiva
  da tela e data do cálculo diário (`modFeed.UltimaCarga`, `modFeed.DiaDoCalculo`), as duas só marcadas depois de
  aplicar com sucesso.
- **DEV-20 — filtro e ordem depois de atualização parcial** (`modGrade`): depois de reescrever as linhas, uma vez
  por lote, reaplica ordenação, AutoFiltro (inclui pesquisa e situação) e mantém a seleção pelo id. Linha nova ou
  excluída em outra estação: a aba é remontada preservando filtro, pesquisa, ordem e seleção (antes só avisava).

## Critérios de aceite

- [ ] `modFeed.NovasLacunas` e `modFeed.ReferenciaInvalida` (lógica pura) com `Confere` no autoteste: lacuna
  aberta, achada, vencida; geração trocada; cursor à frente do banco.
- [ ] `modFeed.Dependentes` com `Confere` (cliente, contato, oportunidade, lista, tabela sem dependente).
- [ ] testar-uso: erro injetado na primeira aba não perde o evento (a rodada seguinte aplica); log gravado com id
  menor que o cursor (commit fora de ordem) aparece na rodada seguinte; geração trocada remonta e não fica presa.
- [ ] testar-uso: telefone do contato alterado em outra estação aparece na linha do atendimento (Oportunidades).
- [ ] testar-uso: com filtro por responsável e ordenação por valor, alteração externa do responsável tira a linha
  da tela e alteração do valor a reposiciona, sem perder os critérios; registro novo de outra estação entra na
  grade com o filtro mantido.
- [ ] Falha de aplicação não marca a tela como atualizada (`UltimaCarga` não muda; testar-uso).
- 🖥 Duas estações: alteração numa aparece na outra em até 25 s com filtro e ordem mantidos; virada da meia-noite
  com o CRM aberto muda a situação "vence hoje" → "atrasada".

## Decisões assumidas

| # | Decisão | Premissa | O que mudaria |
|---|---|---|---|
| 011-1 | Lacuna esperada por 5 min, no máximo 200 guardadas. | Transações do front/conector duram segundos. | Transação longa (lote grande) no futuro: subir o prazo. |
| 011-2 | Linha nova/excluída remonta a aba (com filtro, ordem e seleção preservados) em vez de só avisar. | `Montar` já guarda e reaplica filtro, pesquisa e ordem; DEV-20 pede tratar novas e excluídas. | Grade muito grande e lenta: voltar ao aviso. |
| 011-3 | Virada do dia remonta as abas carregadas uma vez, na primeira rodada do dia. | Situação e dias parado são calculados no front. | — |
| 011-4 | Painel não é redesenhado pelo feed: fica marcado como velho e se atualiza ao ser aberto. | Redesenhar o Painel custa segundos e ele não fica aberto o tempo todo. | — |

## Limites

- Sempre: teste primeiro; teste que grava só em banco TESTE; nada de dado real no repositório.
- Nunca: publicar na rede nem tocar o banco de produção daqui; MsgBox no caminho do feed.

## Questões em aberto

Nenhuma.
