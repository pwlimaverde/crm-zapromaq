## Ajuste de atendimentos_por_proxima_acao — 28/09/2026 (1.4.1)

Teste real da 1.4.0: páginas de 54 a 60 mil caracteres passam do que o cliente
MCP mostra numa resposta. A ferramenta ganhou teto próprio de 35.000 caracteres
(`$script:LimiteCaracteresPlanejamento`; `atendimentos_por_interacao` segue com
60.000) e o resumo passou a ir só na primeira página (`a_partir_de` = 0), para
as seguintes levarem mais registros. `instalar-mcp.ps1` passa a anunciar 25
ferramentas (estava 11). Arquivos: `mcp\lib\Ferramentas.ps1`, `mcp\servidor.ps1`
(1.4.1), `mcp\instalar-mcp.ps1`, `instrucoes\crm-zapromaq\SKILL.md` e `.zip`, este
LEIA-ME; cópias com sufixo `.bak-20260928-080300`. Também precisa ir ao
repositório de desenvolvimento à mão.

## Ferramenta atendimentos_por_proxima_acao — 28/09/2026

Nova ferramenta somente leitura no conector de dados (`crm-zapromaq`), versão
1.4.0, para o relatório semanal 02 — Planejamento de Atendimentos: atendimentos
em aberto (etapa fora de Pedido Fechado/Perdido/Descartado; etapa vazia conta
como aberta) com `dt_prox_acao` num período de hoje em diante (data_inicio/data_fim
ou semana ISO; sem parâmetro, de hoje até o domingo da semana corrente). Período
encerrado é recusado; período iniciado antes de hoje começa hoje (`ajustado`);
limite de 92 dias. `incluir_vencidas` (padrão true) traz as ações anteriores a
hoje com `vencida: true`. Resumo por responsável e da equipe: total, vencidas,
por dia, por etapa, por quadro, valor no funil e abertos sem próxima ação
(carteira toda). Mesmo teto de 60.000 caracteres de `atendimentos_por_interacao`,
mas sem perder registro: pagina com `a_partir_de` / `proximo`. LEFT JOIN em
contatos: atendimento sem contato também entra.

Arquivos: `mcp\lib\Ferramentas.ps1` (bloco PLANEJAMENTO e entrada no catálogo),
`mcp\servidor.ps1` (1.4.0), `instrucoes\crm-zapromaq\SKILL.md` e `.zip`, este
LEIA-ME. Cópias anteriores com sufixo `.bak-20260928-075531`. Sem alteração de esquema
do banco, do front nem das outras ferramentas. Sintaxe conferida pelo analisador
do PowerShell e lógica testada com banco simulado antes de gravar.

Gravado direto na rede, fora do conector `crm-zapromaq-dev` (que não alcança `IA\`):
**precisa ser levado ao repositório de desenvolvimento à mão**. O conector carrega
a versão nova quando o Claude Desktop é fechado por completo e reaberto; a skill
da conta precisa ser reinstalada a partir do `.zip`.

## Ferramenta exportar_base — 23/09/2026

Nova ferramenta no conector de dados (`crm-zapromaq`), versão 1.3.0, para análise
por script (PRJ-COM-006): grava `crm-<tabela>.csv` e `crm-extrato.json` na pasta
`destino` (relativa a `1 - COMERCIAL ZAPROMAQ`, localizada pelo nome da pasta,
nunca por letra de unidade; tem de existir; `01 - CRM\01 - CONTROLE` é recusado)
e devolve só o envelope. Tabelas: clientes, contatos, oportunidades (padrão),
log_alteracoes, listas, metas. CSV UTF-8 sem BOM, vírgula, campos entre aspas, CRLF
(o formato de `00 - SCRIPTS\_comum.py`);
atendimentos com os calculados da mesma regra do front. Não grava no banco; o
arquivo é escrito em `.tmp` e trocado inteiro.

Arquivos: `mcp\lib\Ferramentas.ps1` (bloco EXTRATO e entrada no catálogo),
`mcp\servidor.ps1` (1.3.0), `instrucoes\crm-zapromaq\SKILL.md` e `.zip`, este
LEIA-ME. Cópias anteriores com sufixo `.bak-20260923-*`. Sintaxe conferida pelo
analisador do PowerShell e lógica testada com banco simulado; **não testada contra
o ACE** — a primeira chamada real é o teste.

Gravado direto na rede, fora do conector `crm-zapromaq-dev` (que não alcança `IA\`):
**precisa ser levado ao repositório de desenvolvimento à mão**. O conector carrega
a versão nova quando o Claude Desktop é fechado por completo e reaberto; a skill
da conta precisa ser reinstalada a partir do `.zip`.

## Ferramenta atendimentos_por_interacao — 22/09/2026

Nova ferramenta somente leitura no conector de dados (`crm-zapromaq`), para o
relatório semanal comercial: devolve numa chamada os atendimentos abertos e
encerrados com `ultima_interacao` no período (semana ISO ou data_inicio/data_fim;
sem parâmetro, a semana ISO anterior), filtro opcional por responsável, resumo
por responsável (total, por etapa, por situação) e JSON compacto sem campos
vazios. Acima de 60.000 caracteres corta com `truncado: true` e a quantidade real.

Arquivos: `mcp\lib\Ferramentas.ps1` (função `Invoke-AtendimentosPorInteracao` e
entrada no catálogo), `mcp\servidor.ps1` (versão 1.2.0),
`instrucoes\crm-zapromaq\SKILL.md` e `.zip`, este LEIA-ME. Cópias anteriores com
sufixo `.bak-20260922-*`. Sem alteração de esquema do banco nem das ferramentas
existentes; `Protocolo.ps1` não foi tocado (a ferramenta monta o próprio conteúdo).

Gravado direto na rede, fora do conector `crm-zapromaq-dev` (que não alcança `IA\`):
**precisa ser levado ao repositório de desenvolvimento à mão** — o
`trazer-da-rede.ps1` só cobre `SISTEMA\DADOS\fonte`. O conector carrega a versão
nova quando o Claude Desktop é fechado por completo e reaberto.

## Ajuste do ambiente do Codex — 21/09/2026

O Codex pode iniciar o MCP sem variáveis de caminhos que o OLE DB usa.
A ausência dos caminhos de componentes compartilhados fez a enumeração de
provedores falhar com 0x8007007E, reportado pelo conector como ACE não encontrado.
O instalador agora registra explicitamente as variáveis de caminhos do Windows,
incluindo CommonProgramFiles e CommonProgramW6432, apenas para este MCP.
Não altera a configuração do Claude nem as permissões gerais do Codex.

Teste realizado: processo com ambiente limpo e somente as variáveis registradas
pelo instalador executou crm_status com sucesso, ACE 16.0 e esquema 1.0.
A confirmação dentro do Codex requer reabrir o aplicativo para reiniciar o MCP.
Nenhum registro de negócio foi alterado. O teste externo anterior herdava o
ambiente completo e não reproduzia esta diferença do aplicativo.
# Instalação local no Claude Desktop e no Codex

Atualização de 21/09/2026. O programa MCP roda na máquina do usuário e é iniciado
pelo aplicativo; não exige hospedagem, serviço Windows nem porta de internet.
A estação precisa acessar a pasta de rede e ter o motor ACE de 64 bits.

- `INSTALAR-CLAUDE-E-CODEX.bat`: registra o MCP de dados nos dois aplicativos e
  instala a skill no Codex. Ambos os aplicativos precisam estar instalados.
- `INSTALAR-MCP.bat`: mantém a instalação apenas no Claude Desktop.
- Para somente Codex: executar `mcp\instalar-mcp.ps1 -Aplicativo Codex` no Windows PowerShell.
- Feche e reabra os aplicativos depois da instalação. No Claude, encerre também
  pelo ícone da bandeja. Não é necessário deixar um terminal aberto.
- No Claude, importe `instrucoes\crm-zapromaq.zip` em Customize > Skills > Upload
  a skill (ou a opção equivalente da versão instalada). A importação da skill é
  separada do registro do MCP; o instalador não automatiza essa interface.
- Para conferir sem alterar registros, peça ao assistente para executar somente
  `crm_status` pelo conector `crm-zapromaq`.

O instalador preserva backups das configurações alteradas. O caminho passado a
`powershell.exe -File` usa `ProviderPath`, sem o prefixo de provedor PowerShell
que impedia a inicialização no Claude.

Verificação nesta estação: protocolo aprovado; comandos registrados no Claude e
no Codex iniciaram o MCP, listaram 11 ferramentas e executaram `crm_status` com
sucesso (ACE 16.0, esquema 1.0). Nenhum registro de negócio foi alterado nesse teste.
A apresentação do conector dentro dos aplicativos precisa ser conferida após
reabri-los. O conector de desenvolvimento já existente no Claude também teve seu
caminho corrigido; ele não foi instalado no Codex e não é necessário ao uso do CRM.

---
# IA — acesso do Claude Desktop ao CRM

Esta pasta dá ao Claude Desktop leitura e gravação no banco do CRM **sem risco de
corromper o arquivo**: toda operação passa por um servidor MCP local, em PowerShell,
que usa o mesmo motor do front (ACE) e as mesmas regras de gravação.

Há um segundo servidor, **de desenvolvimento** (`mcp-dev\`), só para quem mantém o
sistema: ajusta as telas e o código do front, sem tocar no banco. Veja a seção
*MCP de desenvolvimento* no fim.

```
IA\
  INSTALAR-MCP.bat        registra o servidor no Claude Desktop deste usuário
  INSTALAR-MCP-DEV.bat    registra o MCP de desenvolvimento (só para quem mantém)
  mcp\
    servidor.ps1          servidor MCP (stdio); o Claude Desktop inicia sozinho
    instalar-mcp.ps1      edita a configuração do Claude Desktop (com backup)
    lib\Protocolo.ps1     núcleo do protocolo MCP (dual-era), usado pelos dois servidores
    lib\Banco.ps1         conexão curta, parâmetros tipados, transação, log
    lib\Regras.ps1        normalização e campos calculados — espelho do front
    lib\Ferramentas.ps1   catálogo fechado de operações (25 ferramentas)
  mcp-dev\
    servidor-dev.ps1      MCP de desenvolvimento: fontes do front, prévia, montagem
    lib\Catalogo.ps1       catálogo fechado (15 ferramentas)
    lib\Fontes.ps1         caminhos, listar/ler fontes
    lib\Gravacao.ps1       gravar_fonte
    lib\Previa.ps1         prévia, montagem, publicação, logs
    lib\Processos.ps1      execução dos scripts de build/visual
    lib\Migracoes.ps1      estado, migrações, conferência
                          (dividido em arquivos: o antivírus bloqueava o arquivo único)
  instrucoes\crm-zapromaq\SKILL.md       skill de uso do CRM (dados)
  instrucoes\crm-zapromaq-dev\SKILL.md   skill de manutenção do sistema
  teste\                  banco de teste fictício e teste automático do servidor
  logs\                   diagnóstico do servidor (gerado; fora do git)
```

## Como funciona

1. O Claude Desktop inicia `servidor.ps1` e conversa com ele por stdio (JSON-RPC).
   O servidor fala as duas gerações do protocolo MCP: a atual (**2026-07-28**, sem
   estado, `server/discover` e versão em `_meta` a cada requisição) e a anterior
   (`initialize`, até 2025-11-25). Serve para qualquer versão do Claude Desktop.
2. A IA só enxerga as ferramentas do catálogo — não existe SQL livre.
3. Cada chamada abre a conexão, executa e fecha. Gravação: confere a versão do
   esquema, confere a versão do registro (se alguém alterou depois da leitura,
   recusa), normaliza, grava e registra em `log_alteracoes` com origem `IA`,
   tudo numa transação.
4. O banco padrão é `BASE\crm_zapromaq.accdb` (dois níveis acima de `mcp\`).

Não precisa de administrador, serviço, tarefa agendada nem instalação: usa o
Windows PowerShell 5.1 que já vem no Windows. Exige **Office 64 bits** (ou o
*Access Database Engine 2016 Redistributable x64*) na máquina do Claude Desktop.

## Teste em outra máquina (exemplo)

1. Copie a pasta `BASE` inteira para a máquina com o Claude Desktop.
2. `IA\teste\CRIAR-BANCO-TESTE.bat` — cria `BASE\crm_zapromaq.accdb` com dados fictícios.
3. `IA\teste\TESTAR-MCP.bat` — conversa com o servidor como o Claude Desktop faria,
   nos dois protocolos; tem de terminar em **TUDO CERTO**.
   (`TESTAR-MCP.bat -SoProtocolo` testa só o protocolo, sem banco.)
4. `IA\INSTALAR-MCP.bat` — registra o servidor. Feche o Claude Desktop por completo
   (inclusive o ícone perto do relógio) e abra de novo.
5. No campo de mensagem: **+ › Conectores** — `crm-zapromaq` com 11 ferramentas.
6. Instale a skill (conforme a documentação do Claude, em `doc_dev`):
   - ligue **Code execution and file creation** em **Settings › Capabilities**
     (skills dependem disso; em plano Team/Enterprise, o administrador liga em
     *Organization settings › Skills*);
   - compacte a **pasta** `instrucoes\crm-zapromaq` (a pasta inteira, com o
     `SKILL.md` dentro) em `.zip`;
   - envie em **Customize › Skills › Upload a skill**.
7. Peça, por exemplo: *"quais atendimentos estão com ação atrasada?"* ou
   *"passe a próxima ação do AT-0001-0001 para sexta-feira"*.

Para rodar de novo do zero, repita o passo 2 (recria o banco).

## Na produção

A pasta `BASE` vai inteira para o servidor. Em cada máquina que usar o Claude
Desktop, rode `IA\INSTALAR-MCP.bat` **a partir da pasta da rede**: o instalador
grava o caminho UNC (`\\servidor\...`), então não depende de letra de unidade.
Atualizar o servidor MCP = substituir os arquivos na rede; o Claude Desktop pega
a versão nova ao ser reaberto.

## Problemas comuns

| Sintoma | O que conferir |
|---|---|
| Conector não aparece | Claude Desktop foi fechado por completo? Rode `INSTALAR-MCP.bat` de novo e veja em quais arquivos ele gravou |
| Aparece com erro | Log do Claude: `%APPDATA%\Claude\logs\mcp-server-crm-zapromaq.log` (versão da Store: `%LOCALAPPDATA%\Packages\Claude_*\LocalCache\Roaming\Claude\logs`) e `IA\logs\mcp-AAAAMMDD.log` |
| "Motor ACE não encontrado" | Office 32 bits ou sem Office: instale o Access Database Engine 2016 x64 |
| "Banco não encontrado" | O `.accdb` precisa estar em `BASE\`, ao lado das pastas `SISTEMA` e `IA` |
| "Gravação bloqueada: esquema" | O banco foi migrado para um esquema novo: atualize a pasta `IA` junto com o sistema |
| Ferramentas não aparecem no modo Cowork | Ainda não confirmado se conectores locais ficam disponíveis no Cowork: teste primeiro numa conversa comum e depois confira em **Conectores** dentro do Cowork |

## MCP de desenvolvimento (`crm-zapromaq-dev`)

Para ajustes pequenos no front feitos pelo Claude Desktop da empresa — mudar a
posição de um botão, um texto, uma cor do tema, uma regra de tela — sem Node,
sem editor e sem a máquina de desenvolvimento — e também para a manutenção do
banco por migração (campo novo, correção de dados) e para publicar a versão nova.

| Ferramenta | O que faz |
|---|---|
| `listar_fontes`, `ler_fonte` | lista e lê `SISTEMA\DADOS\fonte\` (devolve a `versao` do arquivo) |
| `gravar_fonte` | grava com a `versao` lida (recusa se o arquivo mudou), guarda cópia do anterior e registra quem, quando e por quê; VBA sai em Windows-1252 + CRLF |
| `gerar_previa` | desenha a tela pelo Edge e devolve a **imagem** + o relatório (texto que não cabe, sobreposição) |
| `verificar_layout` | só o relatório da conferência, de todas as telas |
| `montar_teste` | monta `SISTEMA\DADOS\execucao\teste\CRM_Zapromaq.xlsm` (compila e roda o autoteste) sem publicar |
| `publicar` | monta e publica na raiz (sobe a versão); exige `confirmacao = "PUBLICAR"` |
| `listar_alteracoes`, `ver_log` | registro das gravações; final do último log (build, visual, migracoes, backup, ambiente) |
| `estado_sistema` | versões (próxima publicação, publicada, esquema), migrações pendentes, banco em uso, Excel na máquina |
| `listar_migracoes`, `ler_migracao` | migrações do banco e se já foram aplicadas |
| `criar_migracao` | escreve a próxima migração numerada a partir dos comandos SQL; `dados` (em transação) ou `estrutura` (campo/tabela nova, sobe `versao_esquema`) |
| `aplicar_migracoes` | aplica as pendentes, com backup antes; `simular` só mostra |
| `verificar_projeto` | conferência estática em PowerShell (sintaxe VBA fica com o `montar_teste`) |

Acrescentar campo é um pacote: migração de estrutura + `modSchema.bas` +
consulta em `modCRM.bas` + `criar-banco.ps1`, depois `verificar_projeto`,
`montar_teste` e `publicar`. O passo a passo está na skill
`instrucoes\crm-zapromaq-dev`, que vai instalada no Claude Desktop de quem mantém.

Fica de fora de propósito: `build\`, `IA\`, a raiz de `BASE` e os arquivos
gerados (`assets\gerado\`, `modTema.bas` — mudam pelo `layout\` + `gerar_previa`).

Instalação: `IA\INSTALAR-MCP-DEV.bat` na máquina de quem mantém o sistema
(precisa de Edge; `montar_teste` e `publicar` precisam de Excel). Instale também
a skill `instrucoes\crm-zapromaq-dev` como a de dados (compacte a pasta em `.zip`
e envie em *Customize › Skills*): é ela que ensina a ordem das coisas — ler com
versão, migração para mexer no banco, conferir, montar teste e só então publicar.
Teste: `IA\teste\TESTAR-MCP-DEV.bat` (numa cópia; nada do projeto muda).

**Todo ajuste feito na rede precisa voltar ao repositório de desenvolvimento**
antes do próximo ciclo: `SISTEMA\DADOS\docs\sincronizar-rede.md`.

