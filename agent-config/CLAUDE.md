# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Layout do repositório

O Claude Code é aberto **nesta pasta** (`agent-config/`) e trabalha sobre a pasta acima. Raiz do git: `..` (`crm-zapromaq/`).

```
crm-zapromaq\
  README.md
  agent-config\       este CLAUDE.md, DIRETRIZES-IA.md, specs\, ferramentas\, doc_dev\ (plano e referências),
                      vendor\agent-skills (original, subtree), adaptacao\, .claude\ (GERADO)
  src\BASE\           o entregável (estrutura fixa, nunca reorganizar)
```

Nas seções abaixo, `BASE\` significa `..\src\BASE\`. O legado (`old/`) saiu do repositório; as menções a ele ficam como histórico — a cópia do estudo está em `..\src\BASE\SISTEMA\DADOS\docs\legado\`.

## Processo de desenvolvimento (agent-skills)

**Leia `DIRETRIZES-IA.md`.** Todo trabalho segue `/spec → /plan → /build → /test → /review → /code-simplify → /ship` (comandos e skills em `.claude\`, gerados de `vendor\agent-skills` + `adaptacao\`):

- **Contexto, nesta ordem:** este arquivo → `specs\PROJETO.md` (objetivo, limites, **definição de pronto**, TDD por stack) → `specs\stacks\<stack>.md` das camadas tocadas → `specs\funcionalidades\NNN-nome\SPEC.md`.
- **Spec antes de código;** plano em fatias verticais; **teste que falha primeiro** (tabela em `specs\PROJETO.md`); um commit por tarefa.
- **Pronto = prova:** `powershell -File ferramentas\verificar-tudo.ps1` (`-Rapido` | padrão | `-Completo`) sem falha, com a saída mostrada.
- **Parar e perguntar** em teste sem correção óbvia, spec ambígua e ação irreversível (lista em `DIRETRIZES-IA.md`).
- `.claude\skills|commands|agents|references` são **gerados**: mude `adaptacao\` e rode `ferramentas\instalar-skills.ps1`. `vendor\agent-skills` nunca se edita (atualizar: `ferramentas\atualizar-agent-skills.ps1`); o `CLAUDE.md`/`AGENTS.md`/`.claude` de dentro dele **não** valem para este projeto.
- Mudou uma stack? Atualize a spec dela no mesmo commit.

## Plano e referências

- **Checklist de desenvolvimento:** `doc_dev/planejamento/PLANO.md` — fases, decisões (D01…), correções (C01…), pendências (P01…) e regras de trabalho. Marcar itens só depois da verificação.
- **Documentação oficial offline:** `doc_dev/planejamento/referencias/` (índice em `README.md`; atualizar com `python atualizar-referencias.py`). Consultar antes de usar recurso de VBA/MSForms/Excel/ADO/ACE/PowerShell 5.1/MCP.

## Contexto

CRM comercial da Zapromaq (vendas de máquinas industriais). O objetivo deste repositório é **modernizar o sistema legado** que está em `old/` e transformá-lo em um projeto profissional, mantendo o mesmo modelo de execução.

### Restrições de implantação (inegociáveis)

- **100% local, sem internet.** Nada de web app, CDN, API externa, download em tempo de execução, pacote instalado na estação ou OCX/DLL de terceiros. As estações têm só Windows + Office (64 bits). **A versão do Excel é travada no 2019**: nada de recursos do Microsoft 365 (matrizes dinâmicas, `LET`, `XLOOKUP`, `LAMBDA`, etc.) nem no front nem nas fórmulas geradas pelo build.
- **Modelo de execução:** uma pasta compartilhada no servidor contém o banco `.accdb` e o `.xlsm` montado. O usuário copia o `.xlsm` dessa pasta para a própria máquina e usa a cópia local como front, acessando os dados pela rede.
- **Tudo que é necessário para montar o `.xlsm` fica na pasta da rede:** fontes VBA, scripts de build, assets (logo, ícones) e dados de configuração. O build roda numa estação a partir da rede, com o que já vem no Windows (PowerShell 5.1 + automação COM do Excel).
- **O entregável é uma pasta.** Copiar a pasta inteira para a rede tem de bastar para montar e rodar — sem caminhos absolutos da máquina de desenvolvimento e sem letra de unidade (use caminhos relativos à pasta ou UNC).
- **A máquina de desenvolvimento tem Excel do Microsoft 365** (desde 04/10/2026) e o motor ACE: dá para montar com `MONTAR-FRONTEND.bat -Teste`, compilar e rodar autoteste e teste de uso aqui. Não é o Excel 2019 das estações — recurso que só existe no 365 passa aqui e quebra lá; a conferência final continua numa estação com 2019.
- **Ferramentas de desenvolvimento são livres** (git, Python, linters, testes, geradores) desde que fiquem só no desenvolvimento: nada disso pode ser pré-requisito para montar ou usar o sistema na rede.

Repositório git na raiz `crm-zapromaq/`, público no GitHub (`pwlimaverde/crm-zapromaq`), com **git-flow** (merge `--no-ff`):

- `main` = o que está publicado; `develop` = integração. Nunca commitar direto em nenhuma das duas.
- Trabalho novo: `git flow feature start <nome>` (sai de `develop`, volta para ela); correção de `develop`: `bugfix/`.
- Publicação: `git flow release start X.Y` → build na rede → `git flow release finish X.Y` (merge na `main`, tag `vX.Y`). A tag é o número que o build gravou em `VERSAO.txt`; `hotfix/` sai da `main` para correção urgente do publicado.
- **Finalizar com `agent-config\ferramentas\finalizar-branch.ps1 [-Enviar]`**, não com `git flow ... finish`: o GitFlow .NET 2.3.0 instalado (winget `Kubis1982.GitFlow`) falha em todo merge `--no-ff` ("Value cannot be null. (Parameter 'message')") e deixa o merge pela metade; o script faz o mesmo finish (feature/bugfix → develop; release/hotfix → main + tag `vX.Y` + develop) e conclui um merge que a ferramenta tenha deixado pendente. O `start` da ferramenta funciona.
- **A versão da rede tem prioridade.** O usuário ajusta o sistema na empresa pelo Claude Desktop; antes de mexer aqui, a BASE de lá volta para o repositório (comparação em três vias com o último commit; o que mudou lá entra como está). Para levar: `empacotar-base.ps1` e extrair **por cima** da BASE da rede — nunca apagá-la (banco, planilha, `crm_zapromaq.ico` e `execucao\` só existem lá); lá só se roda o `MONTAR-FRONTEND.bat`, que não publica sem o banco e nunca repete versão (parte da maior entre `VERSAO.txt` e `VERSAO-FRONT.txt`).
- Testar o build fora da rede só com `MONTAR-FRONTEND.bat -Teste`: sem `-Teste` ele publica e sobe a versão de `VERSAO.txt`.
- Repositório público: **nenhum dado real** (cliente, CNPJ, telefone, endereço, valor de carteira, nome de estação) em código, teste ou documentação — exemplo é sempre fictício. `old/`, bancos e planilhas ficam no `.gitignore`. O histórico anterior (com `old/`) foi guardado fora do repositório em `crm-zapromaq-historico-ate-2026-10-04.bundle`.
- `.gitattributes` com `* -text`: o git nunca converte fim de linha (VBA em cp1252 + CRLF entra byte a byte).

Documentação, comentários de código e mensagens estão em pt-BR.

## Estrutura alvo (aprovada)

O entregável fica em `../src/BASE/`; `agent-config/` (com `doc_dev/`) ficam fora dele e nunca vão para a rede.

```
BASE\                      pasta copiada inteira para a rede
  crm_zapromaq.accdb       banco
  CRM_Zapromaq.xlsm        front; nome fixo (o INICIAR-CRM.bat compara VERSAO-FRONT.txt)
  SISTEMA\
    MONTAR-FRONTEND.bat    único ponto de entrada do build
    DADOS\                 VERSAO.txt, fonte\, build\, banco\ (esquema, migracoes), operacao\,
                           ferramentas\, testes\, docs\, execucao\ (gerado, fora do git)
  IA\                      servidor MCP (PowerShell) usado pelo Claude Desktop; independente do front
```

- **Build:** `.bat` chama o PowerShell 64 bits pelo caminho completo; monta em arquivo temporário, valida (compila, confere `Início!B7`) e só então substitui o `.xlsm` da raiz. A versão sobe sozinha a cada montagem (1.4 → 1.5) e só é gravada em `VERSAO.txt` se o build deu certo. O texto de `Início!B7` (`Ambiente: X   |   Versao X.Y   |   Publicada em dd/MM/yyyy HH:mm`) **não pode mudar de padrão nem de célula**.
- **IA:** servidor MCP precisa ser **dual-era**: protocolo legado (`initialize`, ≤ 2025-11-25) e o atual 2026-07-28 (sem estado: `server/discover`, `_meta` por requisição, `resultType`, `ttlMs`/`cacheScope`). Sem fila nem agendador (exigiriam administrador e não são imediatos). O Claude Desktop inicia `IA\mcp\servidor.ps1` por stdio; catálogo fechado de ferramentas, gravação com controle de versão e `log_alteracoes.origem = 'IA'`. Leitura e escrita pelo ACE; `mdbtools` e o botão "Aplicar pendências" do front saem. O servidor confere `config.versao_esquema` antes de gravar — migração de esquema exige atualizar `IA` junto.
- **Atualização das estações:** `modFeed` (ver *Front*). Todo caminho de gravação registra log dentro da transação (Fase 3).
- **Base do código:** `old/MIGRACAO-CRM` (validada em produção). Do `EDITADO` só se aproveitam os mockups em `12 - DESIGNER/referencia/`; seus módulos `modZpm*`/`modUI*`, classes e scripts `montar-lab-*` são descartados.
- **Front visual (aprovado):** "desenhar em HTML, executar em MSForms". Fonte única em `fonte\layout\` (tokens de tema + posição dos elementos de cada tela). **Sem Node no fluxo**: a prévia é JavaScript rodando no próprio Edge (dados de layout carregados como `.js`, pois `fetch` não funciona em `file://`); PowerShell 5.1 orquestra o Edge headless (`msedge --headless --screenshot` / `--dump-dom`; o modo antigo saiu do Chromium 132+) para gerar as imagens de fundo, o módulo de tema VBA e a tabela de posições do build. Tudo precisa rodar numa estação só com Windows + Edge + Excel 2019. Decoração (cartões, sombras, degradês) vira imagem BMP **embutida no formulário pelo build, em tempo de desenho** (`Picture` só aceita bmp/gif/jpg) — nunca `LoadPicture` de disco em tempo de execução (o `.xlsm` roda de `Documentos`). Texto e ícones são nativos: rótulos transparentes e glifos da fonte **Segoe MDL2 Assets**. Grade própria de rótulos substitui a ListBox. Estações: **1366×768**, escala do Windows variável → projetar a 100% e, ao abrir, `modTela.Encaixar` ajusta `UserForm.Zoom` pela área útil da tela em pontos (A15; máximo 100).
- **MCP de desenvolvimento:** ver *IA*.
- Arquivos `.ps1` são gravados em **UTF-8 com BOM** (o PowerShell 5.1 lê sem BOM como ANSI). No servidor MCP nada pode ir para o stdout além de mensagens JSON-RPC.
- Testes do MCP: `IA\teste\TESTAR-MCP.bat -SoProtocolo` (sem banco) ou completo após `CRIAR-BANCO-TESTE.bat` (exige ACE).

## O sistema legado (`old/`)

Duas cópias da mesma árvore de trabalho:

- `old/MIGRACAO-CRM/MIGRACAO-CRM/` — versão estável que foi para produção em 16/09/2026. Contém o documento de referência `2026-09-16_ESTUDO-MIGRACAO-CRM.md` (decisões de arquitetura, modelo de dados, regras de negócio, critérios de aceite). **Leia-o antes de redesenhar qualquer coisa.**
- `old/MIGRACAO-CRM-EDITADO/` — ramificação mais recente (front v1.4, 18/09/2026) com um redesenho visual experimental: módulos `modZpm*`, `modUI*`, `clsZpmButton`, assets em `12 - DESIGNER/`, scripts `montar-lab-v*.ps1` e `montar-crm-real-teste-v1.ps1`. Aqui o `montar-frontend.ps1` grava por padrão em `30 - TESTES\CRM_Zapromaq_UI_TESTE.xlsm`, não no modelo de produção.

`40 - DICIONARIO/estado-migracao.md` é o diário de decisões e defeitos corrigidos (o que foi tentado, por que falhou). Consulte-o antes de "corrigir" algo que parece estranho no código — quase sempre há um motivo registrado ali.

Estrutura de cada cópia: `10 - FONTES-VBA` (fonte do front), `20 - SCRIPTS` (PowerShell/Python/.bat), `30 - TESTES`, `40 - DICIONARIO`, `99 - HISTORICO` (versões substituídas com sufixo de data; `.xlsm` e logs gerados são saída, não fonte).

### Stack atual

| Camada | Tecnologia |
|---|---|
| Banco | Access `.accdb` numa pasta de rede (UNC), motor `Microsoft.ACE.OLEDB.16.0`, Office **64 bits** |
| Frontend | Excel `.xlsm` com VBA + ADO em late binding; uma cópia local por estação |
| Build / operação | PowerShell 5.1 (`System.Data.OleDb`, automação COM do Excel) disparado por `.bat` |
| Sessão de IA | Python (`openpyxl`) + `mdbtools` (binários Linux em `20 - SCRIPTS/lib/mdbtools`) — **somente leitura** do banco |

Tabelas: `clientes`, `contatos`, `oportunidades`, `listas`, `metas`, `log_alteracoes`, `config` (DDL em `criar-banco.ps1` e seção 5 do estudo).

### Arquitetura do front VBA (legado; o código em `BASE` partiu daqui)

- `modConfig` — caminho UNC do banco (aba `Config!B2`), ambiente (`B3`), `VERSAO_FRONT`, e o estado global `gModoModelo` (arquivo aberto da rede → nada grava) e `gGravados` (ids salvos pela ficha, lidos pela grade ao fechar).
- `modDB` — único ponto de acesso: `Consultar`, `Executar`, `InserirComId` (`@@IDENTITY` na mesma conexão/transação), `AtualizarComVersao`, `P(tipo, valor)` para parâmetros.
- `modSchema` — metadados que **geram** a ficha: `campo|Rótulo|tipo|lista|seção|obrigatório|editável`. `CamposDe` alimenta o formulário e o SQL de `modCRM.Inserir/Atualizar`; `ColunasPlanilha` define as colunas das abas de grade; `ColunasDe` a lista do formulário.
- `modCRM` — regras de negócio (códigos, promoção, SQL das listagens). `modCalc` — campos calculados (situação, quadro, dias parado, ciclo, meses), que não existem no banco. `modValidacao` — normalização de texto na gravação. `modGrade` — abas Clientes/Contatos/Oportunidades como grade somente leitura com AutoFiltro. `modPainel` — ~10 `GROUP BY` gravados na aba oculta `BasePainel`. `modAcoes`, `modLote` (+ `frmLote`), `modMenu` (macros dos botões).
- `frmCRM.txt` — código do UserForm; os campos da ficha são criados em tempo de execução a partir de `modSchema`. Os nomes dos controles fixos (`lstReg`, `fraCampos`, `cmdSalvar`…) são contrato com o código.
- `ThisWorkbook.txt` — proteção da cópia modelo e os eventos das abas de grade (`Workbook_Sheet*`, porque abas criadas por automação não têm CodeName).

O `.xlsm` **não é fonte**: é montado do zero pelo build. O modelo fica na rede; `INICIAR-CRM.bat` (raiz de `BASE`, no projeto desde 04/10/2026) copia para `Documentos\CRM Zapromaq\` da estação quando `Início!B7` muda.

No legado, a IA gravava por uma fila JSON (`BASE\FILA\`, `aplicar-fila.ps1`) com `mdbtools` só leitura: **eliminado** (D06), substituído pelo MCP.

## Comandos

Tudo roda com PowerShell 5.1 (`powershell.exe`, não `pwsh`) a partir de `..\src\BASE\`. Python só no desenvolvimento.

```powershell
# Verificação estática de TUDO (VBA: cp1252/CRLF, sintaxe por parser ANTLR MS-VBAL, blocos, chamadas modX.Y,
# nomes públicos repetidos; controles citados no código x layout JSON; modSchema x DDL; .ps1; .bat; JSON)
python SISTEMA\DADOS\ferramentas\verificacao\verificar.py

# Parte visual: prévias (execucao\previa\<tela>-100|125|150.png), fundos (fonte\assets\gerado), modTema.bas
powershell -File SISTEMA\DADOS\build\gerar-visual.ps1 [-Tela frmCRM] [-Tolerante]

# Build (só numa estação com Excel 2019): -Teste monta em execucao\teste sem publicar nem subir versão
SISTEMA\MONTAR-FRONTEND.bat [-Teste]

# Testes
powershell -File SISTEMA\DADOS\testes\testar-build-lib.ps1           # funções do build (versão, B7, pacote zip)
IA\teste\TESTAR-MCP.bat -SoProtocolo                                   # MCP de dados, sem banco (18 testes)
IA\teste\CRIAR-BANCO-TESTE.bat ; IA\teste\TESTAR-MCP.bat               # completo (exige ACE)
IA\teste\TESTAR-MCP-DEV.bat [-SemPrevia]                               # MCP de desenvolvimento, numa cópia (21 testes)

# Banco
SISTEMA\DADOS\banco\APLICAR-MIGRACOES.bat                              # migrações numeradas (config.versao_esquema)
powershell -File SISTEMA\DADOS\testes\criar-banco-demo.ps1 [-Recriar]  # banco fictício (20 de cada) para ver o front
powershell -File SISTEMA\DADOS\banco\esquema\criar-banco.ps1 -Banco <accdb> -Recriar

# Rede -> repositório (antes de qualquer ciclo novo): só fonte\ ...
powershell -File SISTEMA\DADOS\ferramentas\sincronizar\trazer-da-rede.ps1 -Rede \\servidor\...\BASE [-Simular]
# ... ou a BASE inteira recebida da empresa: comparar em três vias com o último commit (a rede tem prioridade)

# Repositório -> rede: pacote do commit (sem banco/planilha/logs), extraído POR CIMA da BASE de lá
powershell -File ..\..\agent-config\ferramentas\empacotar-base.ps1    # -> dist\BASE-vX.Y-....zip na raiz do repositório
```

Na máquina de desenvolvimento o VBA é conferido pelo `verificar.py` e pela prévia, e o build `-Teste` (Excel 365) compila e roda o autoteste (`modAutoteste`) e o teste de uso; a palavra final é o build numa estação com Excel 2019. Roteiro de implantação: `SISTEMA\DADOS\docs\implantacao.md`.

## Front (estado atual, Fases 4–6)

- **Layout é dado:** `fonte\layout\tema.json` (cores, fontes, variantes de botão, cores de situação, `moldura`), `icones.json` (glifos Segoe MDL2), `amostras.json` (dados fictícios da prévia) e `formularios\<tela>.json` (v2: `canvas` = área interna em pt, `fundo` = decoração que vira BMP, `controles` nativos, `componentes` botao/icone/selo/grade). Nomes de controle são contrato com `formularios\<tela>.txt` (o `verificar.py` confere).
- **Gerados, não editar:** `fonte\assets\gerado\*` (fundos PNG + `manifesto.json` com hash das entradas) e `modulos\modTema.bas`. O build roda `gerar-visual` sozinho se o manifesto estiver velho, converte o PNG em BMP 96 dpi e embute no `Designer.Picture` por um módulo VBA temporário.
- **Componentes (fonte\classes):** `clsUI` liga tudo pelo `Tag` (`ui:botao:<variante>`, `ui:grade`); botão = três Labels `bg_<nome>`, `<nome>` (texto; `cmdX_Click`/`.Enabled` continuam valendo) e `ico_<nome>`; depois de mudar `Enabled`, `mUI.Pintar`. `clsGrade` substitui a ListBox (API parecida: `ColumnWidths`, `Cabecalho`, `Carregar matriz`, `ListIndex`...) e chama `UI_GradeClique`/`UI_GradeDuploClique` no formulário; clique no fundo/ícone do botão chama `UI_Acao(nome)`. `modTela.Encaixar` aplica o zoom pela área útil antes do `Show`.
- **Feed (modFeed):** `Application.OnTime` a cada 25 s, `MAX(id)` de `log_alteracoes`; reescreve só as linhas alteradas (`modGrade.AtualizarRegistros`), avisa novos/excluídos na linha de status e a ficha aberta (`frmCRM.AvisoExterno`). Iniciado no `Workbook_Open`, cancelado no `BeforeClose`. Por isso **toda gravação precisa de log na mesma transação** (`modDB.LogEm`).
- **Tempo de resposta (regras novas):** caches por sessão — listas (10 min, `modCRM.ObterLista`), tabela config (5 min, `modConfig.ValorConfig`), provedor ACE memorizado; `modMenu.Pronto` não sonda o banco se houve consulta há < 60 s; várias leituras numa operação usam **uma** conexão (`modDB.AbrirLeitura`/`ConsultarEm`/`FecharLeitura`); coluna de consulta por dicionário (`modDB.IndiceDeCampos`), nunca varrendo `gCampos`. Revisão completa em `PLANO.md` §4c.
- `frmCalibracao` (Alt+F8 › `AbrirCalibracao`) mostra todos os controles sobre um fundo gerado, para comparar com a prévia na estação.

## IA

- `IA\mcp\servidor.ps1` (dados, 25 ferramentas; também instalável no Codex por `IA\INSTALAR-CLAUDE-E-CODEX.bat`) e `IA\mcp-dev\servidor-dev.ps1` (desenvolvimento, 15 ferramentas: lê/grava só `SISTEMA\DADOS\fonte`, prévia como imagem, migrações do banco — listar/ler/criar/aplicar —, conferência estática, estado do sistema, montagem de teste, publicação com confirmação) usam o mesmo núcleo `IA\mcp\lib\Protocolo.ps1`. Instalação: `IA\INSTALAR-MCP.bat` / `IA\INSTALAR-MCP-DEV.bat`.
- Skills do Claude Desktop em `IA\instrucoes\`: `crm-zapromaq` (dados) e `crm-zapromaq-dev` (manutenção: onde mexer, pacote de campo novo, quando publicar).
- Ajuste feito na rede pelo MCP de desenvolvimento fica em `execucao\dev\alteracoes.log` (+ cópias) e volta ao repositório com `trazer-da-rede.ps1` (`docs\sincronizar-rede.md`).

## Regras que o legado aprendeu na prática (preservar em qualquer stack)

- **Conexão curta por operação do usuário, nunca por linha.** Um ciclo abre/consulta/fecha custa ~26 ms na rede; consulta dentro de laço trava a tela.
- **SQL sempre parametrizado.** A base tem `INDUSTRIA D'ANGELO LTDA`.
- **Data nunca como literal no SQL** — o ACE lê `#mm/dd/yyyy#` e a estação usa dd/mm; o erro é silencioso.
- `NZ()` não existe via OLE DB; nulo se trata no código, não no SQL. `ORDER BY` do Access põe nulo na frente.
- **Controle de versão otimista** em todo `UPDATE` (`WHERE id=? AND versao=?`; 0 linhas = conflito, nunca sobrescrever). Dado e `log_alteracoes` na mesma transação.
- Códigos `CT-CCCC-NNNN` / `AT-CCCC-NNNN` são **congelados** na criação: nomeiam ~480 pastas de atendimento em `02 - CLIENTES`. Nunca recalcular nem reformatar.
- `clientes.id` é a identidade interna; `codigo_cliente` só existe para `Cliente`. `estagio` é derivado (sem código → `Pré-cliente`), nunca digitável. Primeiro atendimento promove o pré-cliente na mesma transação (`MAX+1`).
- Normalização de texto num ponto só (`modValidacao`): empresa/cidade/UF/máquina/orçamento em MAIÚSCULAS; nomes preservam digitação; e-mail minúsculo; CNPJ e telefone só dígitos (telefone 10–11); memo como digitado; acento preservado. Campo vazio grava `NULL` (índice único de CNPJ aceita vários nulos, não várias strings vazias).
- Os campos `ctx_*` só são escritos pela IA (MCP), nunca pelo formulário — é o que evita conflito entre as duas escritas.
- `Microsoft.Jet.OLEDB.4.0` é proibido (gera JET4 com extensão `.accdb`).

## Armadilhas do ambiente VBA/Excel (se ainda for mexer no legado)

- Fontes VBA misturam ASCII e **Windows-1252**; leia/grave em cp1252 com CRLF, senão os acentos corrompem na importação.
- Máximo de 25 continuações de linha por instrução (por isso `modSchema` usa um `Add` por campo).
- Cor no Excel é `R + G*256 + B*65536`. Paleta da marca: azul `#172A67`, azul secundário `#153F71`, verde `#58B030`.
- Nome de forma no Excel tem até 31 caracteres; botões usam `Placement = xlFreeFloating`.
- Nada de célula mesclada na área de dados da grade (quebra AutoFiltro); nada de OCX/DLL externos, URL ou subclassing (`SetWindowsHookEx`) no UserForm.
- Layout precisa caber em 1366×768 (1 pt do UserForm ≈ 1,333 px a 96 dpi).
- A versão não se edita à mão: o build lê `SISTEMA\DADOS\VERSAO.txt`, injeta em `modConfig.VERSAO_FRONT` na cópia de montagem e só grava a nova versão se tudo deu certo; o `.xlsm` anterior vai para `execucao\historico-front`.
