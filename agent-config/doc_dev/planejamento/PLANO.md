# Plano de desenvolvimento — CRM Zapromaq

**Revisão:** 01 · **Data:** 19/09/2026 · **Situação:** aprovado, pronto para iniciar a Fase 0
**Referências técnicas:** [`referencias/README.md`](referencias/README.md) (documentação oficial baixada em 19/09/2026)

## Como usar este plano

- Cada item é uma caixa `[ ]`. Só se marca `[x]` quando a **verificação** do item passou — não quando o código foi escrito. Ao marcar, anote a data: `[x] ... (19/09)`.
- Itens que dependem da estação (Excel 2019) estão marcados com **🖥 estação**. O resto é feito e verificado na máquina de desenvolvimento.
- Decisão nova ou mudança de rumo entra na seção **Decisões** com data e motivo antes de virar código.
- Retomada depois de pausa: refazer a verificação do último item marcado, nunca confiar na memória de onde parou.

---

## 1. Objetivo

Modernizar o CRM Comercial Zapromaq **mantendo o modelo de execução atual** (planilha `.xlsm` como front + banco Access `.accdb` numa pasta compartilhada, sem internet) e transformá-lo num projeto profissional: fonte versionada, build reproduzível, versão automática, testes, documentação, visual moderno e acesso seguro da IA (Claude Desktop) aos dados.

## 2. Restrições inegociáveis

| # | Restrição |
|---|---|
| R1 | 100% local, sem internet. Nada baixado em tempo de execução, nada instalado nas estações, nada de OCX/DLL de terceiros. |
| R2 | Estações: Windows + **Excel 2019 64 bits** (versão travada). Nada de recursos exclusivos do Microsoft 365 (matrizes dinâmicas, `LET`, `XLOOKUP`, `LAMBDA`…). |
| R3 | A pasta `BASE` é o entregável: copiada inteira para a rede, tem de montar e rodar. Na raiz de `BASE` ficam **só** `crm_zapromaq.accdb`, `CRM_Zapromaq.xlsm`, `SISTEMA\` e `IA\`. |
| R4 | O `.xlsm` mantém o **nome fixo** `CRM_Zapromaq.xlsm`. O `.bat` das estações (fora do projeto) copia a versão nova comparando o **texto inteiro** de `Início!B7`: `Ambiente: PRODUCAO   \|   Versao X.Y   \|   Publicada em dd/MM/yyyy HH:mm`. Célula e padrão **não mudam**. |
| R5 | Tudo que monta o `.xlsm` fica na rede e roda com o que existe numa estação: **PowerShell 5.1 + Edge + Excel 2019**. Node, Python e afins só no desenvolvimento, nunca como pré-requisito. |
| R6 | Nada que exija administrador (agendador no servidor, serviços, instaladores). |
| R7 | Máquina de desenvolvimento: sem Excel, com liberdade para ferramentas gratuitas; ficará sem rede. Build e teste do front só numa estação. |
| R8 | Estações em **1366×768**, escala do Windows variável. |

## 3. Decisões

| # | Data | Decisão | Motivo |
|---|---|---|---|
| D01 | 18/09 | Stack mantida: Excel `.xlsm` + Access `.accdb` + PowerShell | Restrições R1–R6 |
| D02 | 18/09 | Estrutura `BASE\{accdb, xlsm, SISTEMA\, IA\}`; `SISTEMA\{MONTAR-FRONTEND.bat, DADOS\}` | Definida pelo usuário |
| D03 | 18/09 | Versão automática: cada build publicado sobe a casa final (1.4 → 1.5); só é gravada se o build der certo | Fim da edição manual em `modConfig` |
| D04 | 18/09 | Build: `.bat` → PowerShell 64 bits → Excel via COM; monta em arquivo temporário, valida e só então substitui o `.xlsm` da raiz | Build falho não pode deixar a raiz sem front |
| D05 | 18/09 | IA via **servidor MCP local em PowerShell** (stdio), catálogo fechado de ferramentas, gravação pelo ACE com controle de versão e log `origem='IA'` | Imediato, sem admin, sem risco de corromper o `.accdb` |
| D06 | 18/09 | Fila JSON, agendador, `mdbtools` e botão "Aplicar pendências" **eliminados** | Substituídos por D05 |
| D07 | 18/09 | Atualização das estações por **feed de `log_alteracoes`** (`MAX(id)` a cada 20–30 s) reescrevendo só as linhas alteradas | Tela nunca mente sobre alteração de outro usuário ou da IA |
| D08 | 19/09 | Base do código = `old/MIGRACAO-CRM` (validada). Do `EDITADO` só os **mockups** | Código do `EDITADO` bagunçado e quebrado |
| D09 | 19/09 | Visual: **"desenhar em HTML, executar em MSForms"** — decoração vira imagem embutida no build; texto, campos e ícones (Segoe MDL2 Assets) são nativos; grade própria substitui a ListBox | Único jeito de visual moderno sem instalar nada |
| D10 | 19/09 | Gerador visual **sem Node**: JS roda no próprio Edge; PowerShell orquestra o Edge headless | Ajustes futuros na empresa (R5) |
| D11 | 19/09 | Layout desenhado para **1000 × 505 pt a 100%**; ao abrir, o formulário aplica `Zoom` conforme a DPI | 1366×768 com escala variável (R8) |
| D12 | 19/09 | **MCP de desenvolvimento** separado do MCP de dados, para ajustes pequenos pelo Claude Desktop da empresa | Praticidade sem instalar nada |
| D13 | 19/09 | Repositório git na raiz `crm-zapromaq\`: `BASE\` (entregável), `doc_dev\` (documentação de desenvolvimento), `old\` (legado: só textos versionados) | `doc_dev` não pode ir para a rede (R3) |
| D14 | 04/10 | Repositório público reorganizado: `src\BASE` (entregável), `agent-config\` (Claude Code, specs, ferramentas, `doc_dev`); `old\` fora; histórico anterior em bundle fora do repositório; git-flow | Versionar no GitHub sem dado real |
| D15 | 04/10 | A BASE vai para a empresa como pacote do commit (`empacotar-base.ps1`), extraído **por cima** da BASE da rede; a versão da rede tem prioridade e volta ao repositório antes de cada ciclo | Ajustes feitos lá pelo Claude Desktop |
| D16 | 04/10 | Processo de desenvolvimento do **agent-skills** (Addy Osmani, MIT): original na íntegra em `agent-config\vendor` (subtree), adaptação em `agent-config\adaptacao`, `.claude` gerado; specs do projeto e das stacks em `agent-config\specs` | Atualizar o original sem perder a adaptação |
| D17 | 06/10 | Cadastro em lote: **SUSPEITO começa fora** e só entra por duplo clique; OK também alterna. Reverte a decisão do comercial de 17/09 | Relatos de 18 a 29/09: o suspeito que entrava sozinho virava cadastro repetido que ninguém revia (`specs/funcionalidades/001-demandas-6-5-3-4`) |
| D18 | 06/10 | Atendimento anterior usa a coluna **existente** `oportunidades.id_atendimento_anterior` (esquema 1.0, já gravada pela IA): sem migração 004; só da mesma empresa e nunca o próprio, no front e na IA | Duas colunas com o mesmo sentido; migração estrutural evitável; a 003 segue reservada |
| D19 | 06/10 | O front, como a IA, **só troca o contato do atendimento dentro da mesma empresa** (código AT- congelado); empresa errada = encerrar como Descartado e abrir um novo apontando para ele | Corrige o defeito R15 |

## 4. Correções trazidas pela revisão com a documentação (19/09)

| # | O que a documentação mostrou | Impacto no plano |
|---|---|---|
| C01 | **MCP 2026-07-28** tornou o protocolo sem estado: sai o `initialize`, entra `server/discover` obrigatório, `_meta` com `io.modelcontextprotocol/protocolVersion` em toda requisição, `resultType` em todo resultado, `ttlMs`/`cacheScope` no `tools/list`, erro `-32022` para versão não suportada, `ping` removido. [12-mcp](referencias/12-mcp-especificacao/) | O servidor do exemplo é só "legado" (≤ 2025-11-25). Vira **dual-era**: atende `initialize` (legado) **e** requisições modernas. Item 1.6 |
| C02 | Chromium 132+ removeu o headless antigo; `--headless` já é o modo novo. [10-edge](referencias/10-edge-headless/) | Usar `--headless` (sem `=new`) |
| C03 | Skills: envio por **Customize › Skills** (ZIP) e exigem **Code execution** ligado em Settings › Capabilities. Nome até 64 caracteres, minúsculas/hífen, sem "claude"/"anthropic"; descrição até 1024. [13-claude](referencias/13-claude-desktop/) | Corrigir `IA\LEIA-ME.md`; `crm-zapromaq` já atende às regras |
| C04 | `Picture`: em tempo de desenho se atribui pela página de propriedades (persiste no arquivo); em execução, só `LoadPicture`. Formatos: **bmp, gif, jpg** (sem PNG). [01-msforms](referencias/01-vba-msforms/) | Fundo em **BMP**, atribuído **em tempo de desenho pelo build** (`VBComponents(..).Designer`), nunca `LoadPicture` de disco em execução |
| C05 | `OnTime`: hora arredondada ao segundo; cancelar exige o **mesmo** `Procedure` e `EarliestTime`; a rotina não pode estar em classe nem formulário; só roda com o Excel em modo Pronto. [04-excel](referencias/04-excel-modelo-objetos/) | Guardar a hora agendada numa variável de módulo; rotina em módulo padrão; usar `LatestTime` para não empilhar; testar com ficha aberta |
| C06 | `TOP n` do Access **não desempata**: devolve mais de n linhas se houver empate. [06-access-sql](referencias/06-access-sql-ace/) | Todo `ORDER BY` com `TOP` termina numa chave única (`id`/`codigo`) |
| C07 | SQL no VBA usa **data em formato EUA**. [06-access-sql](referencias/06-access-sql-ace/) | Confirma a regra: data sempre como parâmetro tipado |
| C08 | `OLE DB Services=-4` desliga o pool de sessões do OLE DB. [08-dotnet](referencias/08-dotnet-oledb/) | Mantido nos scripts PowerShell (solta o `.laccdb` ao fechar) |
| C09 | Windows PowerShell 5.1 lê script sem BOM como ANSI. [09-powershell](referencias/09-powershell-5.1/) | `.ps1` sempre **UTF-8 com BOM**; fontes VBA em **Windows-1252 + CRLF** (importação do VBE é ANSI) |
| C10 | Documentação do ribbon traz o `customUI` 2006/01 (`customUI/customUI.xml`, relação `.../2006/relationships/ui/extensibility`). [05-ribbon](referencias/05-excel-ribbon-customui/) | Se a faixa própria for adotada, usar esse formato documentado (compatível com 2019) |
| C11 | Office 2019: suporte estendido **terminou em 14/10/2025**. ACE 2016 Redistributable (publicado em 18/11/2025) exige remover versões anteriores do ACE antes de instalar. [14-office](referencias/14-office-2019/) · [06-ace](referencias/06-access-sql-ace/) | Risco registrado (seção 8); instruções de instalação do ACE na Fase 0 |

## 4b. Defeitos do legado encontrados na auditoria do código (19/09)

O legado funcionava, mas a leitura linha a linha das fontes achou defeitos reais. Todos corrigidos na Fase 3 (ou onde indicado).

| # | Defeito | Efeito para o usuário | Correção |
|---|---|---|---|
| A01 | Cadastro novo gravado sem `ativo` (Sim/Não no Access não aceita nulo: vira **Falso**) | Todo cliente/contato criado pela ficha ou pelo lote nascia **INATIVO** | `modCRM.InserirEm` grava `ativo = True`; migração **002** reativa os afetados (só quem nunca foi desativado de propósito) |
| A02 | Depois de salvar, `Buscar` zerava a seleção da lista | "Salvar e próx." com a ficha aberta pelo menu **pulava para o 1º registro** | `frmCRM`: reseleciona o registro gravado |
| A03 | Contador "X de Y" usava `SpecialCells(visível).Rows.Count` | Com filtro, o número de linhas visíveis saía **errado** (só o 1º bloco) | `SUBTOTAL(103)` na coluna de id |
| A04 | Erro na montagem da grade saía sem reproteger a aba | Grade ficava **editável** (a origem das perdas da 3.6) | `Proteger` no tratamento de erro |
| A05 | Grade com **uma** linha: `Value2` não é matriz | Erro ao voltar da ficha | Embrulha em matriz 1×1 |
| A06 | Desativar/reativar em cascata em 2-3 conexões sem transação; exclusão gravava o log **antes** do DELETE | Cascata pela metade em queda de rede; log afirmando exclusão que não ocorreu | Uma transação; log por registro, depois da operação; versão sobe |
| A07 | Promoção e criação do atendimento em transações separadas; `MAX+1` calculado fora da transação que insere | Empresa promovida sem atendimento; dois vendedores ao mesmo tempo → **erro de chave duplicada** na tela | Tudo numa transação; colisão repete com o próximo número |
| A08 | Log de alteração só dizia "alteração pelo formulário" | Auditoria sem saber **o que** mudou | Uma linha por campo alterado, com valor antigo e novo; nada mudou = nada gravado |
| A09 | Preencher combos disparava `Change` → `Buscar` | Abrir a ficha fazia **5 a 7 consultas** à rede | Trava `mCarregando`: uma consulta por ação |
| A10 | `CStr(Boolean)` comparado com "False" | Depende do idioma do Office | `CBool` |
| A11 | `modLote.NaLista` consultava o banco **por linha**; telefone fora da regra passava | Lote lento; telefone inválido gravado | Lista lida uma vez; telefone fora de 10–11 dígitos vai para observações e a linha fica SUSPEITA |
| A12 | `modPainel.ContarPorSituacao` nunca era chamado | Painel sem o bloco de situação de acompanhamento | Gravado na `BasePainel` |
| A13 | `Config!B4` gravado a **cada consulta** | Arquivo sempre "alterado"; eventos de célula à toa | No máximo 1×/minuto, com eventos desligados |
| A14 | `agendar-backups` com `RunLevel Highest` | Exigia administrador (viola R6) | `RunLevel Limited`, logon interativo |
| A15 | Zoom calculado só pela DPI (item 4.9) | Em processo sem suporte a DPI o Windows informa 96 dpi: zoom errado | Zoom pela **área útil em pontos** (px × 72 / dpi), igual nos dois modos (checar-ambiente já usa) |
| A16 | `AplicarPendencias` com caminho fixo da estrutura antiga | Botão quebrado fora daquela árvore | Removido (D06) |
| A18 | Aba **Painel** nunca era desenhada: `modPainel` só gravava a `BasePainel` e as fórmulas do Painel ficaram na planilha 3.6 | Painel sempre em branco | `modPainel.DesenharPainel`: cartões, tabelas e gráficos nativos a cada atualização; atualiza sozinho ao abrir a aba (mais de 15 min); botões pelo build |
| A17 | `TOP 50 ... ORDER BY empresa` sem desempate (C06) | Mais de 50 linhas e ordem instável | Desempate por `id` em todas as ordenações |

## 4c. Revisão geral de 19/09 (defeitos, usabilidade e tempo de resposta)

| # | Achado | Efeito | Correção |
|---|---|---|---|
| R01 | `NovaConexao` tentava ACE 16 e depois 12 a cada conexão | Rede fora do ar: até 30 s de Excel travado por ação (e a cada 25 s com o feed) | Provedor memorizado na primeira conexão |
| R02 | `Pronto()` abria 2 conexões (teste + esquema) antes de CADA botão | Latência somada a toda ação | Sem teste se houve consulta há < 60 s; tabela config em cache de 5 min (1 consulta na abertura) |
| R03 | Uma consulta de lista por combo ao abrir a ficha | Ficha abria com 8–10 idas à rede | `ObterLista` com cache de 10 min: uma consulta traz todas as listas |
| R04 | Trocar de registro na ficha consultava o banco (resumo do vínculo) | Contrariava a regra "clicar na linha não consulta"; navegação lenta | Resumo montado da linha da grade (colunas `estagio`/`codigo_cliente` na consulta) |
| R05 | Coluna procurada pelo nome, célula a célula (`Campo`/`CampoDaLinha`) | Montar grade de 500 linhas: centenas de milhares de comparações | Índice nome→posição (`modDB.IndiceDeCampos`) |
| R06 | Painel com ~20 conexões | Atualização lenta | Uma conexão (`AbrirLeitura`/`ConsultarEm`/`FecharLeitura`) |
| R07 | Regra condicional `=OR(...)` na grade | Excel em português lê fórmula de formatação no idioma local: Perdido/Descartado nunca esmaeciam | Duas regras sem função |
| R08 | `Config!B4` gravado a cada minuto marcava o arquivo como alterado | Com o feed, TODO fechamento perguntava se salva | Gravação preserva `ThisWorkbook.Saved` |
| R09 | Botão direito cancelado em qualquer célula das abas de grade | Não dava para copiar telefone/e-mail | Menu próprio só na área de dados, com **Copiar** |
| R10 | `_ % [` na busca da ficha e `* ? ~` na pesquisa da aba eram curingas | Resultado errado | Escapados (`PadraoLike`, `TextoFiltro`) |
| R11 | Grade própria sem teclado (a ListBox tinha) | Regressão de usabilidade | Setas, PgUp/PgDn, Home/End, Enter abre |
| R12 | Ficha aceitava aderência/porte fora de 1–3 (o lote recusava) | Dado inconsistente | Crítica na ficha |
| R13 | Situação GRAVADO do lote sem cor | Selo neutro | Cor no tema |
| R15 | (06/10) "Procurar" na edição de atendimento deixava escolher contato de outra empresa; o `Atualizar` grava `id_contato` mas não `id_cliente` | Atendimento apontando para duas empresas | Troca só na empresa do atendimento (`frmVinculo.EscolherContatoDe`) e checagem ao salvar (`modCRM.CriticarTrocaContato`) — D19 |
| R16 | (06/10) Ficha guardava um único id de vínculo por tabela (`mVinculoID`) | Um segundo campo K gravaria o id do contato nele | Vínculo por campo (`modSchema.CamposK`, `mVinculos`) |
| R14 | Backup periódico é cópia de arquivo com usuários conectados | Cópia pode sair inconsistente (registrado como AVISO no log) | **Pendente/recomendação**: exportação tabela a tabela para um banco novo, ou manter o diário (exclusivo) como cópia de referência |

## 5. Arquitetura alvo

```
crm-zapromaq\                     repositório git (máquina de desenvolvimento)
  CLAUDE.md
  .gitignore
  BASE\                           ENTREGÁVEL: copiado inteiro para a rede
    crm_zapromaq.accdb            banco (fora do git)
    CRM_Zapromaq.xlsm             front montado (fora do git)
    SISTEMA\
      MONTAR-FRONTEND.bat         único ponto de entrada do build
      DADOS\
        VERSAO.txt                número da versão; só o build altera
        fonte\
          modulos\ classes\ formularios\ pasta-de-trabalho\   VBA em Windows-1252 + CRLF
          layout\                 tema (cores, fontes, espaços) + posição dos elementos de cada tela
          assets\                 logo, imagens-fonte; assets\gerado\ = fundos BMP gerados (versionados)
        build\                    montar-frontend.ps1 e funções; gerador visual (PowerShell + páginas JS)
        banco\esquema\            criar-banco.ps1 (esquema 1.0)
        banco\migracoes\          001-..., 002-...  (sobem config.versao_esquema)
        operacao\                 backup, restauração, agendamento (+ scripts\)
        ferramentas\              só desenvolvimento: verificações, registro da migração 3.6
        testes\
        docs\                     estudo, dicionário, contrato do banco, CHANGELOG.md
        execucao\                 gerado, fora do git: backup\ historico-front\ logs\ previa\
    IA\
      INSTALAR-MCP.bat
      mcp\                        servidor.ps1 + lib\ (dados)
      instrucoes\crm-zapromaq\    SKILL.md
      teste\                      banco fictício + teste automático
      logs\                       fora do git
  doc_dev\planejamento\           este plano + referências oficiais
  old\                            legado (só textos no git; binários ignorados)
```

---

## 6. Fases

### Fase 0 — Preparação do ambiente de desenvolvimento

- [x] 0.1 Instalar o **Access Database Engine 2016 Redistributable x64** (remover versões anteriores do ACE antes, conforme a página oficial) e confirmar `Microsoft.ACE.OLEDB.16.0` no PowerShell 64 bits — (19/09) ACE 16.0 instalado; provedores 12.0 e 16.0 registrados
- [x] 0.2 `git init` em `crm-zapromaq\`; `.gitignore` na raiz do repositório (retirar o de `BASE\`, que viola R3) — (19/09) git em crm-zapromaq/, .gitignore na raiz do repositório
- [x] 0.3 Mover `exemplo\BASE` para `crm-zapromaq\BASE` (o exemplo vira o início do projeto real) — (19/09)
- [x] 0.4 Versionar os textos de `old\` (fontes, scripts, documentos); ignorar `.xlsm`, `.accdb`, logs, `Thumbs.db` e `lib\mdbtools` — (19/09) textos de old/ versionados; binários, logs e mdbtools ignorados
- [x] 0.5 Primeiro commit: estrutura + `CLAUDE.md` + `doc_dev\` — (19/09) commit 3020bfd, tag inicio
- [x] 0.6 Rodar `IA\teste\CRIAR-BANCO-TESTE.bat` e `TESTAR-MCP.bat` completos (15 testes de banco pendentes) — (19/09) TUDO CERTO (protocolo + banco), também com as migrações 001/002 aplicadas; corrigido criar-banco-teste (função Cli colidia com o apelido de Clear-Item)

**Verificação:** `TESTAR-MCP.bat` termina em "TUDO CERTO"; `git status` limpo; nenhum binário no histórico.

### Fase 1 — Estrutura real e migração do código validado

- [x] 1.1 Copiar as fontes de `old/MIGRACAO-CRM/.../10 - FONTES-VBA` para `fonte\modulos|formularios|pasta-de-trabalho`, **sem alterar conteúdo** (mesma codificação) — (19/09) cópia byte a byte conferida com cmp; commit 72e0090
- [x] 1.2 Migrar scripts de `20 - SCRIPTS` para `build\`, `banco\`, `operacao\`, `ferramentas\` com **caminhos relativos a `BASE`** (fim do "sobe cinco níveis") — (19/09) lib/Comum.ps1 acha a raiz de BASE pela posição do script
- [x] 1.3 Unificar `criar-banco.ps1` (esquema 1.0 = produção atual, com índice único de CNPJ) e registrar as migrações já aplicadas em produção como histórico — (19/09) esquema 1.0 = produção (mesmos nomes de índice e FK); migrações 001 e 002 em banco/migracoes
- [x] 1.4 Levar `conferir-vba.py`, `conferir-sql-gerado.py`, `conferir-calculados.py` para `ferramentas\verificacao` e ajustar caminhos — (19/09) substituídos por verificar.py (inclui a lógica do conferir-vba) + verificar-sintaxe-vba.py (parser ANTLR/MS-VBAL, zero falso positivo no legado)
- [x] 1.5 Documentos do legado (estudo, dicionário, estado da migração, montagem, distribuição) para `DADOS\docs` — (19/09) docs/legado
- [x] 1.6 **MCP dual-era (C01):** `server/discover`; aceitar `_meta` moderno por requisição (versão `2026-07-28`); `resultType: "complete"` em todo resultado; `ttlMs` e `cacheScope` no `tools/list`; `-32022` com `supported`; manter `initialize` para clientes legados; atualizar `testar-mcp.ps1` com os dois caminhos — (19/09) 18 testes de protocolo passando (legado e 2026-07-28)
- [x] 1.7 Corrigir `IA\LEIA-ME.md` (C03: Customize › Skills + Code execution) — (19/09)
- [x] 1.8 `docs\contrato-banco.md`: regras que front e IA compartilham (versão de esquema, versão de registro, log, normalização, datas, conexões curtas) — (19/09) docs/contrato-banco.md
- [x] 1.9 Verificações estáticas passando: `conferir-vba.py` sem problemas; analisador do PowerShell sem erros em todos os `.ps1`; `conferir-sql-gerado.py` sem divergência — (19/09) verificar.py: sem problemas

**Verificação:** as três verificações do 1.9 limpas; teste do MCP nos dois protocolos.

### Fase 2 — Build novo (sem mudar nada visível) — 🖥 estação na verificação

- [x] 2.1 `MONTAR-FRONTEND.bat`: chama o PowerShell 64 bits pelo caminho completo (`System32`/`Sysnative`) e confere o resultado pelo arquivo, não só pelo código de saída — (19/09)
- [x] 2.2 Pré-checagens com mensagem clara: Excel presente, **"Confiar no acesso ao modelo de objeto do projeto do VBA"** ligado, `.xlsm` da raiz não está em uso, trava contra dois builds simultâneos — (19/09) Excel ausente e acesso ao VBA tratados antes de alterar qualquer configuração; testado aqui
- [x] 2.3 Montagem em arquivo temporário em `execucao\`; importação das fontes lidas em Windows-1252 — (19/09)
- [x] 2.4 Validação antes de publicar: compilação do projeto VBA sem erro, todos os módulos presentes, `Início!B7` no padrão exato (R4) conferido por expressão regular — (19/09) compilação via VBE + modAutoteste com vigia; pacote conferido pelo zip (testar-build-lib.ps1: 17 testes)
- [x] 2.5 Publicação: `.xlsm` anterior → `execucao\historico-front\CRM_Zapromaq_v<versão>.xlsm`; novo → raiz com o nome fixo — (19/09)
- [x] 2.6 Versão: lê `VERSAO.txt`, calcula a próxima (1.4 → 1.5), injeta em `modConfig.VERSAO_FRONT` no momento da montagem, escreve `B7`, grava `config.versao_front` no banco e **só então** grava `VERSAO.txt` e o `CHANGELOG.md` — (19/09)
- [x] 2.7 `Config!B2` preenchido automaticamente com o caminho **UNC** de `BASE\crm_zapromaq.accdb` (unidade mapeada convertida), porque o `.xlsm` roda de `Documentos` — (19/09)
- [x] 2.8 Excel sempre encerrado no `finally`, inclusive em erro; log em `execucao\logs` — (19/09) caminho de falha testado aqui: raiz intacta, opção do VBA restaurada, trava solta
- [x] 2.9 Opção `-Teste`: monta em `execucao\` sem publicar nem subir versão — (19/09)
- [ ] 2.10 🖥 **Ida à estação nº 1:** montar o front atual pela estrutura nova, com o INICIAR-CRM real, e conferir que a cópia automática funciona e que tudo se comporta igual à versão 1.4

**Verificação:** estação abre a versão 1.5 pelo atalho de sempre; comportamento idêntico ao 1.4; build falho não altera a raiz.

### Fase 3 — Correções de base no front — 🖥 estação na verificação

- [x] 3.1 Todo caminho de gravação registra `log_alteracoes` **dentro da transação** (hoje `modAcoes.Registrar` grava fora dela, com `On Error Resume Next`) — (19/09) modDB.AbrirTransacao/LogEm; ativar, desativar, excluir, promover, inserir e alterar com log por campo na mesma transação
- [x] 3.2 `modCRM.PromoverCliente` com SQL parametrizado (hoje concatena valores) — (19/09) e numeração MAX+1 dentro da transação com nova tentativa em colisão
- [x] 3.3 Remover `modMenu.AplicarPendencias`, o botão "Aplicar pendências", `modConfig.PastaFila` e qualquer referência à fila e ao `mdbtools` — (19/09) e o botão saiu de abas.json
- [x] 3.4 Toda consulta com `TOP` termina o `ORDER BY` numa chave única (C06) — (19/09)
- [x] 3.5 Revisão de `On Error Resume Next` que engole erro de gravação — (19/09) erros de gravação agora sobem com mensagem; Registrar silencioso removido

**Verificação:** conferência de que cada `INSERT`/`UPDATE`/`DELETE` do front gera log na mesma transação; testes de gravação na estação.

### Fase 4 — Sistema visual

- [x] 4.1 `fonte\layout\tema.*`: cores da marca (azul `#172A67`, azul secundário `#153F71`, verde `#58B030`, neutros), fontes (Segoe UI), espaçamentos, raios, sombras
- [x] 4.2 Formato de layout por tela (dados), com posição em **pontos** e tipo de cada elemento: decoração, controle nativo, componente próprio
- [x] 4.3 Página de prévia (HTML + JS rodando no Edge): desenha a tela a partir dos dados, em 100/125/150%, com camada de contorno dos controles e dados de exemplo
- [x] 4.4 Verificações automáticas na prévia: sobreposição, texto que não cabe, tela maior que 1000 × 505 pt
- [x] 4.5 Gerador em PowerShell: Edge `--headless` (C02) produz os **fundos BMP** (C04) e exporta a tabela de posições e o módulo de tema VBA
- [x] 4.6 Build atribui o fundo **em tempo de desenho** (`VBComponents(..).Designer`) e cria os controles nas posições exportadas
- [x] 4.7 Biblioteca de componentes VBA: botão (normal/sobre/desabilitado), grade virtual com linhas alternadas e situação colorida, selo de status, campo, seção, abas
- [x] 4.8 Ícones por glifo da **Segoe MDL2 Assets** (Windows 10 e 11); tabela de glifos usados em `layout\icones.*`
- [x] 4.9 Zoom automático (A15): área útil da tela em pontos (`SystemParametersInfo(SPI_GETWORKAREA)` × 72 / `GetDeviceCaps(LOGPIXELSX)`, `Declare PtrSafe`) ÷ tamanho do formulário; máximo 100
- [x] 4.10 **Tela de calibração**: todos os tipos de controle sobre um fundo gerado
- [ ] 4.11 🖥 **Ida à estação nº 2:** print da calibração em 100% e 125%; ajustar a prévia até coincidir

**Verificação:** diferença entre prévia e print da estação aceita visualmente; fundo nítido a 100% e aceitável a 125%/150%.

### Fase 5 — Front novo

- [x] 5.1 Redesenho do `frmCRM` a partir do mockup, cabendo em 1000 × 505 pt
- [x] 5.2 Grade própria no lugar da `lstReg`; ficha em seções; barra de ações
- [x] 5.3 `frmLote` no mesmo tema
- [x] 5.4 Abas Início, Clientes, Contatos, Oportunidades, Painel com o mesmo tema (formas e formatação pelo build)
- [ ] 5.5 (Opcional) Guia própria na faixa de opções (C10) — **adiado**: as abas já têm botões da marca; customUI exige editar o zip do .xlsm e só se confere na estação
- [ ] 5.6 🖥 **Teste de uso** na estação com vendedor real

**Verificação:** critérios de aceite da seção 7 + aprovação visual do usuário.

### Fase 6 — Atualização automática das estações — 🖥 estação

- [x] 6.1 `modFeed` em módulo padrão (C05): guarda hora agendada, consulta `MAX(id)` de `log_alteracoes`, busca os registros alterados numa consulta só, reescreve as linhas (`AtualizarUmaLinha`) — (19/09) conferido estaticamente; comportamento real na 6.5
- [x] 6.2 Registro novo → aviso na linha de status, sem remontar a grade — (19/09) conferido estaticamente; comportamento real na 6.5
- [x] 6.3 Ficha aberta alterada por outro → aviso na ficha — (19/09) conferido estaticamente; comportamento real na 6.5
- [x] 6.4 Cancelar o agendamento no `Workbook_BeforeClose` (senão o Excel reabre o arquivo) — (19/09) conferido estaticamente; comportamento real na 6.5
- [ ] 6.5 🖥 Testar com duas estações e com a IA gravando

### Fase 7 — MCP de desenvolvimento

- [x] 7.1 Ferramentas: ler/alterar só `SISTEMA\DADOS\fonte\`, gerar prévia devolvendo **imagem**, montar `.xlsm` de teste, publicar só a pedido — (19/09) IA\mcp-dev, 9 ferramentas; TESTAR-MCP-DEV: 21 testes passando (montar/publicar só na estação)
- [x] 7.2 `INSTALAR-MCP-DEV.bat` (mesmo instalador, outro servidor) — (19/09) instalar-mcp.ps1 -Dev; protocolo extraído para mcp\lib\Protocolo.ps1 (18 testes do MCP de dados continuam passando)
- [x] 7.4 Manutenção completa pelo Claude Desktop (sem Claude Code na empresa): `estado_sistema`, `listar_migracoes`, `ler_migracao`, `criar_migracao` (dados/estrutura, recusa SQL destrutivo sem `CONFIRMO`), `aplicar_migracoes` (com `simular`), `verificar_projeto` (verificar.py, com alternativa em PowerShell sem Python) — (20/09) 15 ferramentas; validação de `array` acrescentada ao Protocolo; skill `IA\instrucoes\crm-zapromaq-dev`; TESTAR-MCP-DEV: 31 testes passando
- [x] 7.3 Regra de sincronização rede ↔ desenvolvimento documentada — (19/09) docs\sincronizar-rede.md + ferramentas\sincronizar\trazer-da-rede.ps1 (testado com -Simular)

### Fase 8 — Implantação

Roteiro passo a passo, com conferências e volta atrás: `BASE\SISTEMA\DADOS\docs\implantacao.md` (19/09). Os itens abaixo são executados na rede da empresa.

- [ ] 8.1 Backup da pasta de produção atual e do `.accdb`
- [ ] 8.2 Copiar `BASE` para a rede; mover o `.accdb` de produção para a raiz de `BASE`
- [ ] 8.3 Montar o front na rede; conferir `B7` e a cópia automática nas estações
- [ ] 8.4 Instalar MCP e skill nas máquinas com Claude Desktop
- [ ] 8.5 Reagendar backups com os caminhos novos
- [ ] 8.6 Restauração de backup testada ("backup não testado não é backup")

---

## 7. Testes e critérios de aceite

**Automáticos (máquina de desenvolvimento):** `conferir-vba.py`, analisador do PowerShell, `conferir-sql-gerado.py`, `TESTAR-MCP.bat` (protocolo legado e moderno + banco), verificações da prévia.

**Na estação:** montagem, compilação, `B7`, cópia pelo INICIAR-CRM, calibração visual, uso real.

**Critérios de aceite herdados do estudo de migração (continuam valendo):**

- [ ] Empresa sem código permanece na fila e pode ser promovida pelo fluxo definido
- [ ] Ordenar ou filtrar a grade não altera código nem vínculo
- [ ] Dois usuários (ou usuário e IA) no mesmo registro: conflito tratado, nunca sobrescrita silenciosa
- [ ] Falha depois de gravar e antes de confirmar não duplica registro
- [ ] Cancelar não grava; fechar com alteração pendente avisa
- [ ] Opção de lista desativada continua legível no histórico e não é oferecida em lançamento novo
- [ ] Data brasileira, moeda, acento, apóstrofo, campo vazio e observação longa preservados
- [ ] Os códigos `AT` e `CT` existentes continuam idênticos
- [ ] Toda gravação (front e IA) aparece em `log_alteracoes` na mesma transação
- [ ] Alteração feita em outra estação ou pela IA aparece sem clicar em Atualizar (Fase 6)

## 8. Riscos e pendências abertas

| # | Item | Situação |
|---|---|---|
| P01 | Onde fica o `INICIAR-CRM.bat` na rede? No legado ficava na raiz de `BASE`, o que conflita com R3 | **Confirmar com o usuário** antes da Fase 2 |
| P02 | Versão do Windows das estações (MDL2 existe no 10 e no 11) | Confirmar |
| P03 | Claude Desktop atual fala protocolo legado, moderno ou os dois? | Resolvido pelo dual-era (1.6); confirmar no teste real |
| P04 | Conector MCP local disponível no modo Cowork | Testar |
| P05 | `OnTime` com ficha modal aberta: roda ou espera? | Testar na Fase 6 |
| P06 | Nitidez do fundo a 125%/150% | Medir na calibração (4.11); plano B: fundo em três resoluções |
| P07 | Office 2019 sem atualizações de segurança desde 14/10/2025 | Risco aceito pela empresa; registrar |
| P08 | ACE redistribuível e Office Click-to-Run na máquina de desenvolvimento podem conflitar | **Ocorreu** (Excel 365 instalado, 04/10/2026): ACE, MCP e build `-Teste` funcionando juntos. Se o provedor ACE falhar aqui, suspeitar disto primeiro |
| P09 | Mouse wheel em listas não é suportado sem subclassing (derruba o Excel) | Aceito; navegação por barra, teclado e botões |

## 9. Regras de trabalho

1. **Fonte é texto; `.xlsm` e `.accdb` são produto.** Nunca editar o `.xlsm` à mão.
2. VBA em **Windows-1252 + CRLF**; `.ps1` em **UTF-8 com BOM**; Markdown e JSON em UTF-8 sem BOM.
3. SQL sempre parametrizado; data sempre parâmetro; `TOP` sempre com chave única no fim do `ORDER BY`.
4. Conexão curta por operação do usuário, nunca por linha.
5. Mudança de estrutura do banco = migração numerada + `versao_esquema` + atualização do MCP da IA na mesma versão.
6. Ajuste feito na rede (Claude Desktop da empresa) é trazido de volta para o repositório antes de qualquer desenvolvimento novo.
7. Commits pequenos, com mensagem dizendo o porquê; `CHANGELOG.md` atualizado pelo build a cada publicação.
8. Referências: antes de usar um recurso novo de VBA/Excel/ADO/MCP, conferir em `referencias\`. Atualizar com `python referencias\atualizar-referencias.py` quando houver internet.
9. Git-flow: `feature/`/`bugfix/` a partir de `develop`; publicação por `release/X.Y` com tag `vX.Y` na `main`; `hotfix/` a partir da `main`. Nada direto em `main`/`develop`.
10. Repositório público: nenhum dado real em código, teste ou documentação.
