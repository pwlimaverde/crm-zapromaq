# Spec: Pendências finais do plano revisado (ciclo 018)

> Situação: aprovada em 26/10/2026 (pedido do usuário: "implementar tudo de ponta a ponta"). Itens do
> documento-base (seção 5.2 e DEV-02, 03, 04, 06A, 07, 08, 10, 12, 21, 28) que ficaram de fora dos ciclos 010 a 017.
> Uma branch `feature/<nome>` por pendência, a partir da `develop`.

## Objetivo

Fechar o que o documento-base de aceite pede e ainda não está no código, sem publicar nada (a 1.6 sai depois).

## Critérios de aceite (uma pendência por linha; teste que falha primeiro entre parênteses)

- [ ] P1 (DEV-28) Botão "Duplicar" na ficha de oportunidades, mesma rotina do menu da grade (`modGrade.AbrirDuplicacao`);
      contato inativo: a ficha abre com a empresa da origem e a escolha do contato já nessa empresa.
      (`Confere` de `modRegras.RotuloBotaoAcao`; `testar-uso` abre a duplicação pela ficha sem gravar)
- [ ] P2 (5.2) `dt_prox_acao` nova ou alterada anterior a hoje é recusada nos três canais; a vencida já gravada fica.
      (`Confere`; caso C34 na matriz; `testar-mcp`) - conector 1.12.0, nota de mudança.
- [ ] P3 (DEV-21) "Retomada a avaliar" só em Perdido e Descartado, igual no front e no conector. (`Confere`; `testar-mcp`)
- [ ] P4 (5.2, DEV-10) `retomar_em` desabilitado na ficha fora de Perdido e Descartado (`modRegras.RetomadaHabilitada`).
- [ ] P5 (DEV-03) Restauração sem `-Backup` usa a última cópia **validada**. (`testar-recuperacao`)
- [ ] P6 (DEV-08) `config.versao_minima_front` e `config.versao_minima_conector`, gravadas pela montagem e conferidas no
      ponto de escrita do front e do conector. (`Confere`; `testar-mcp`)
- [ ] P7 (DEV-07) A skill do Claude Desktop cita todas as ferramentas do catálogo; o `verificar.py` falha se faltar uma.
- [ ] P8 (DEV-02) `checar-ambiente` com versão do front instalada, pasta local e sincronização, Codex e escrita na pasta
      do banco; matriz de estações em `implantacao.md`. (`testar-preparacao-rede`)
- [ ] P9 (DEV-12) Lote: INDETERMINADA quando a falha é no commit; NÃO PROCESSADA para as linhas não alcançadas. (`testar-uso`)
- [ ] P10 (5.2) Contato homônimo no conector pede confirmação (`confirmar_homonimo`), como a ficha. (`testar-mcp`)
- [ ] P11 (DEV-04) Checksum por migração gravado em `config` e conferido ao aplicar; assinatura do VBA vira pendência
      no PLANO (depende de certificado da TI). (`testar-build-lib`)
- [ ] P12 (DEV-06A) Linha de diagnóstico do front com fuso. (`Confere`)
- [ ] P13 (5.2) Metas "não definida": só se houver tela de meta; senão, decisão registrada no PLANO.

## Decisões assumidas

| # | Decisão | Premissa | O que mudaria |
|---|---|---|---|
| 018-1 | O botão de ação da ficha de oportunidades vira "Duplicar" (o "Atend." desabilitado não tinha uso). | Mesmo espaço no layout; substitui a decisão 017-3. | — |
| 018-2 | Contato inativo na duplicação: a ficha mostra a empresa da origem e o botão do contato lista só os contatos dessa empresa. | O atendimento duplicado é da mesma empresa. | — |

| 018-3 | `dt_prox_acao` no passado: recusada só quando nova ou alterada; encerrar em Perdido/Descartado limpa a data e não é conferido (C34). | 5.2: "datas já vencidas permanecem para cobrança". | — |
| 018-4 | "Retomada a avaliar" só em Perdido e Descartado (resultado PERDA) nos dois canais; Pedido Fechado com `retomar_em` legado fica "Encerrado". | 5.2: retomada só existe em Perdido/Descartado; a seção do conector já filtrava PERDA. | — |
| 018-5 | O conector 1.12.0 acumula todas as mudanças de contrato do ciclo (uma nota de mudança só), porque nada foi publicado entre elas. | Publicação única da 1.6. | Publicar no meio do ciclo exigiria versões separadas. |

| 018-6 | Homônimo no `criar_contato` usa o parâmetro `ignorar_homonimo` (o mesmo nome do `criar_cliente`), não um `confirmar_homonimo` novo. | Um nome só para a mesma confirmação no conector. | — |
| 018-7 | Lote: só commit indeterminado ou escrita recusada antes de começar (manutenção, esquema, versão mínima) interrompem; erro de regra da linha não. | Recusa geral valeria para todas as linhas; erro da linha não. | — |

(As demais decisões de cada item ficam na seção do item, abaixo, à medida que forem tomadas.)

## Stacks tocadas

front-vba, banco-access, build-operacao-powershell, ia-mcp, ferramentas-dev.

## Limites

- Nunca: publicar na rede; gravar fora de banco TESTE; dado real; editar `modTema.bas` à mão.
