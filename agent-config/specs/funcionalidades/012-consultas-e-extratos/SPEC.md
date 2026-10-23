# Spec: Agenda, consultas e extratos (ciclo 012)

> Situação: aprovada em 22/10/2026 (plano revisado de 19/10/2026, Etapa 3; DEV-21, DEV-22 e DEV-23 do documento-base;
> seção 5.2 "Agenda do aberto" e 5.3 "Extratos" e "Cópia local do front"). Sem migração.

## Objetivo

- **DEV-21 — agenda e retomadas.** A agenda do atendimento aberto é só a `dt_prox_acao` (front, conector e
  relatórios): "Retomar hoje" deixa de existir para aberto. Atendimento **encerrado** (Perdido/Descartado) cuja
  `retomar_em` chegou passa a ter a situação **"Retomada a avaliar"** (sem reabrir nada): no front aparece no filtro
  de situação da grade e da ficha; no conector, `atendimentos_por_proxima_acao` ganha a seção `retomadas_a_avaliar`
  (encerrados com `retomar_em` até o fim do período, cada um uma vez), e a situação nova vale em
  `buscar_atendimentos`/`obter_atendimento`. Data futura nunca aparece como vencida.
- **DEV-22 — paginação e extrato com corte definido.** `atendimentos_por_proxima_acao` pagina por **cursor**
  (`cursor` / `proximo_cursor`): chave de ordenação completa com desempate único (código do atendimento), filtros e
  geração da base dentro do cursor. A página seguinte continua **depois da chave** do último item: inserção ou exclusão
  entre páginas não duplica nem omite os demais; cursor com filtro diferente ou de outra geração é recusado. A leitura
  é dinâmica e a resposta diz isso (`consistencia = "dinamica"` + limitação). `a_partir_de` continua aceito
  (compatibilidade) com aviso. `exportar_base` lê cada tabela em consulta separada: o manifesto declara
  `consistencia = "nao_comprovada"` e registra o `MAX(id)` do log antes e depois; se mudou durante a leitura, refaz
  até 2 vezes e informa se não conseguiu um intervalo sem gravação (isso também não prova corte, só reduz a chance).
- **DEV-23 — extratos e cópia local sem dados.**
  - Conector: `exportar_base` grava sempre em `SISTEMA\DADOS\execucao\extratos\<AAAAMMDD-HHMMSS>-<sufixo>\` e devolve o
    caminho; o parâmetro `destino` **sai** (nota de mudança). Manifesto `crm-extrato.json` com solicitante (`autor`),
    estação, finalidade, instante, corte e SHA-256 de cada arquivo. `exportar_roteiro` grava na mesma raiz, com manifesto.
    Ferramenta nova `descartar_extratos` (aditiva; `simular` = true por padrão): apaga só subpastas da raiz de extratos
    com manifesto válido, criadas antes do mês anterior, que não estejam em processamento (`em_processamento` no
    manifesto); registra o descarte em `execucao\extratos\descartes.log`. Caminho canônico conferido: nada fora da raiz.
  - Front: cópia local em `%LOCALAPPDATA%\CRM Zapromaq` (fora de Documentos/OneDrive) pelo `INICIAR-CRM.bat` e pela
    instalação a partir do modelo; a cópia antiga conhecida (`Documentos\CRM Zapromaq\CRM_Zapromaq.xlsm`) é removida
    pelo lançador ao instalar. O arquivo nunca guarda dados: salvar e salvar como são cancelados também na cópia local,
    o fechamento não pergunta se salva (descarta), `EnableAutoRecover = False` só neste arquivo; sem banco, as abas
    de dados são esvaziadas e o aviso diz isso. Roteiro do front vai para a raiz de extratos (com manifesto).

## Critérios de aceite

- [ ] `Confere`: `SituacaoDaEtapa` — aberto com `retomar_em` vencida segue a próxima ação; encerrado com `retomar_em`
  ≤ hoje = "Retomada a avaliar"; encerrado com retomada futura = "Encerrado".
- [ ] testar-mcp: mesma regra no conector (`Get-Situacao`); seção `retomadas_a_avaliar` com o encerrado do período uma
  vez; aberto aparece pela próxima ação; nenhuma data futura em `vencida`.
- [ ] testar-mcp: paginação por cursor — inserção e exclusão entre páginas sem duplicar/omitir; cursor com filtro
  diferente recusado; geração trocada recusada; resposta com `consistencia`.
- [ ] testar-mcp: `exportar_base` sem `destino` grava na raiz de extratos com manifesto (autor, estação, finalidade,
  hashes conferidos, consistência); `destino` informado é recusado; `descartar_extratos` simula e apaga só o vencido com
  manifesto válido, nunca pasta sem manifesto nem fora da raiz.
- [ ] testar-uso: roteiro do front na raiz de extratos com manifesto; `BeforeSave` cancela na cópia local; sem banco
  as abas de dados ficam vazias.
- [ ] Contrato: `VersaoServidor` 1.9.0, `FerramentasEsperadas` + `descartar_extratos`, nota de mudança no `IA\LEIA-ME.md`,
  skill de dados atualizada.
- 🖥 Estação com Documentos no OneDrive: INICIAR-CRM instala em `%LOCALAPPDATA%`, remove a cópia antiga, o arquivo
  fechado com "Salvar" não guarda grade; fechamento forçado não deixa recuperação.

## Decisões assumidas

| # | Decisão | Premissa | O que mudaria |
|---|---|---|---|
| 012-1 | Situação nova "Retomada a avaliar" substitui "Retomar hoje" (mesma cor). | 5.2: retomada só em Perdido/Descartado; interesse novo = atendimento novo. | — |
| 012-2 | `retomadas_a_avaliar` inclui as retomadas vencidas antes do período (ainda não avaliadas) até o fim dele. | "cuja retomar_em chegou". | Mostrar só as do período. |
| 012-3 | Paginação dinâmica por chave (keyset), não snapshot: o item alterado entre páginas pode mudar de lugar; a resposta avisa. | O ACE não dá leitura isolada; tabela temporária no banco operacional é vedada. | — |
| 012-4 | `exportar_base` não promete corte: `nao_comprovada`, com até 2 releituras se o log mudou. | DEV-22: igualdade do último id não prova snapshot. | Corte por cópia validada do banco, se exigido. |
| 012-5 | `finalidade` opcional no `exportar_base` (padrão "análise"); `autor` já é o solicitante declarado. | Não quebrar mais do que o `destino`. | Tornar obrigatória. |
| 012-6 | Descarte por ferramenta explícita (`descartar_extratos`), não por agendador. | Sem agendador (CLAUDE.md). | — |
| 012-7 | Front: salvar é sempre cancelado na cópia local (os dados vivem no banco); o lançador continua sendo quem instala. | 5.3: cópia local não guarda dados. | — |

## Limites

- Sempre: teste primeiro; teste que grava só em banco TESTE; nada de dado real no repositório.
- Nunca: publicar na rede nem tocar o banco de produção daqui; apagar arquivo fora da raiz de extratos.
