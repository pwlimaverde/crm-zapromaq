# Spec: Rota com todos os clientes do raio (ciclo 016)

> Situação: aprovada em 23/10/2026 (plano revisado de 19/10/2026: DEV-27, decisões assumidas 10 a 14). Sem migração.
> Depende de DEV-11 (contato fixo por empresa; contato do atendimento conferido).

## Objetivo

O planejamento de rota só mostrava atendimentos em aberto: a empresa do raio sem atendimento (pré-cliente da
qualificação, cliente sem negócio aberto) ficava invisível e o vendedor montava a rota por fora.

- **Rota lista todos do raio** — além dos candidatos (atendimentos em aberto), a página Candidatos mostra as empresas
  **ativas** do raio **sem atendimento em aberto**, inclusive pré-clientes (decisão 11), marcadas "(sem)".
  Front: `modPlanejamento.EmpresasSemAtendimento`; conector: `listar_candidatos` ganha `sem_atendimento` (aditivo,
  até 200 linhas, com `quantidade_sem_atendimento` e `sem_atendimento_truncado`; decisão 14).
- **"Gerar visitas" só aceita atendimento lançado** (decisão 10): é a única passagem que grava; linha sem atendimento
  não é marcada e o front e o conector recusam id/código que não seja atendimento.
- **Atalho "Criar atendimento"** na página Candidatos: para a empresa sem atendimento selecionada, abre a ficha de
  atendimento novo já com o contato (o único ativo; se houver vários ou nenhum, a escolha de contato oferece
  cadastrar). O anterior é ligado **só** se a empresa tiver Perdido ou Descartado com `retomar_em` (o mais recente;
  decisão 13). Ao voltar, o atendimento criado aparece **marcado** na rota (decisão 12).

## Critérios de aceite

- [ ] testar-mcp: `sem_atendimento` traz o pré-cliente do raio sem atendimento (km 0 em Curitiba), não traz a empresa
  inativa nem empresa que já tem atendimento em aberto; contadores coerentes; `candidatos` não muda.
- [ ] testar-uso: `EmpresasSemAtendimento` com as mesmas regras; `GerarVisitas` com id 0 é recusado sem gravar;
  `AnteriorParaRetomada` devolve o Perdido com retomada (e 0 sem); `ContatoParaAtalho` devolve o único contato ativo.
- [ ] Contrato: `VersaoServidor` 1.10.0, nota de mudança aditiva, matriz de consumidores revista, skill de dados.
- 🖥 frmPlanejamento: linhas "(sem)", botão "Criar atendimento", volta marcado.

## Decisões assumidas

As decisões 10 a 14 do plano revisado. Novas:

| # | Decisão | Premissa | O que mudaria |
|---|---|---|---|
| 016-1 | Empresas sem atendimento vêm **depois** dos candidatos, na ordem qualificação → km; a lista e a ordem dos candidatos não mudam. | Contrato de `listar_candidatos` aditivo; o vendedor visita primeiro quem tem negócio aberto. | Misturar as duas listas numa ordem só. |
| 016-2 | Atendimento em aberto de empresa inativa continua candidato (como hoje); a exclusão de inativas vale para a lista nova. | 5.2: atendimento aberto continua aberto com a empresa desativada. | Tirar também esses da rota. |
| 016-3 | O atalho usa o fluxo existente da ficha (contato escolhido, promoção do pré-cliente com confirmação). | Mesma regra de criação da ficha (DEV-11, 5.2). | — |

## Limites

- Nunca: gerar visita sem atendimento; criar atendimento sem contato.
