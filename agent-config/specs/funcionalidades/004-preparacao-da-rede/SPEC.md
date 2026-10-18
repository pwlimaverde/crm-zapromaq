# Spec: Preparação da rede (ciclo 004)

> Situação: aprovada em 18/10/2026 (pedido explícito do usuário: revisar o plano, executar o ciclo sem pendências e sem perguntas)

## Objetivo

Levar a `develop` (spec 003 já integrada) para a rede da empresa, onde o banco ainda está no esquema 1.0 e **não recebeu nenhuma migração da versão anterior**. Depois deste ciclo:

- nenhum caminho altera o banco de produção fora da montagem (`MONTAR-FRONTEND.bat`);
- toda alteração estrutural acontece numa cópia, é conferida e só então troca o banco;
- existe backup validado (restaurável) antes da publicação e um ensaio sem efeito algum sobre a produção;
- o conector (MCP) com banco antigo recusa com mensagem clara em vez de erro de SQL;
- a publicação é recuperável até o ponto de não retorno (a gravação do `VERSAO-FRONT.txt`);
- os scripts de teste que vão no pacote não conseguem apagar nem gravar no banco de produção.

Nota: o Relatório de análise e o Adendo 01 citados pelo documento-base não estavam disponíveis; o ciclo parte do plano de 514 linhas e do código atual.

## Achados que motivam o ciclo (conferidos no código)

N1 passo 10 grava na cópia do banco real; N2 `APLICAR-MIGRACOES` aplica DDL no banco de produção; N3 `versao_front` gravada depois da troca impede a reversão; N4 `.xlsm` antigo movido antes de confirmar o novo; N5 conector novo aceita 1.0 mas lê colunas do 1.1; N6 o caminho 1.0 sem registros de migração para 1.1 não tem teste; N7 `CRIAR-BANCO-TESTE`/`TESTAR-MCP`/`criar-banco-demo -Recriar` apontam para o banco da raiz; N8 janela entre a conferência de trava e o `Replace`; N9 documentação desatualizada.

## Decisões assumidas

| # | Decisão |
|---|---|
| DA-0 | A rede é tratada como igual ao git (sem ajustes desde 04/10). A estação confere `execucao\dev\alteracoes.log` e `IA\**\*.bak-*` posteriores a 04/10/2026; havendo, não extrai e traz a BASE de volta. Não bloqueia o ciclo. |
| DA-1 | Marcador de teste = `config.ambiente = TESTE`. A pasta `execucao\homologacao\` com `config.homologacao` entra no ciclo 005. |
| DA-2 | RPO 24 h, RTO 4 h, alerta se não houver backup validado há mais de 36 h (constantes em `lib\Backup.ps1`). Cópia fora do servidor = dependência de TI. |
| DA-3 | Migração estrutural só pela montagem: `APLICAR-MIGRACOES` e `aplicar_migracoes` recusam estrutural em banco que não seja TESTE (as de dados 001/002 continuam no lugar). |
| DA-4 | `MONTAR-FRONTEND.bat -Ensaio` roda os passos 1 a 8 numa cópia em `execucao\ensaio\`, sem Excel, sem trocar o banco, e gera relatório. |
| DA-5 | Na publicação real, o passo 10 roda `testar-uso -SoLeitura` e compara impressão digital da cópia antes/depois. O teste que grava só roda no `-Teste`. |
| DA-6 | Contrato MCP só muda de forma aditiva (`atualizacao_pendente`, `esquemas_suportados`; `VersaoServidor` 1.5.0). |
| DA-7 | Ponto de não retorno do passo 12 = gravação do `VERSAO-FRONT.txt` (por último). `config.versao_front` passa a ser gravado na cópia, antes da troca. |
| DA-8 | Backup validado = cópia restaurada em pasta descartável, aberta pelo ACE e conferida (esquema, contagens) contra o manifesto. O arquivo de backup nunca é aberto. |
| DA-9 | Nome: `004-preparacao-da-rede`, `feature/preparacao-da-rede`. |
| DA-10 | Os portões 1 e 2 do ciclo ficam cobertos pelo pedido explícito do usuário. |
| DA-11 | Testes novos com ACE em `SISTEMA\DADOS\testes\testar-preparacao-rede.ps1` (sem Excel), ligado ao `verificar-tudo` padrão com o `testar-migracao-003.ps1`. |
| DA-12 | A1 papel da etapa editável e UF validada só na inclusão (como está); A2 mantém nome e telefone; A3 janela única de implantação cobre o banco exclusivo. |
| Dec 1-10 | Atendimento aberto sem agenda = aviso (sem bloqueio); atendimento antigo fica na empresa histórica; reativar empresa não reativa contatos; prioridade de retomada numa função só; valor histórico desativado preservado; controle global por tabela (índice único só se o inventário achar zero duplicados); um e-mail por campo; "entrada" = primeiro contato; normalização atual mantida; extratos só em `01 - CRM\02 - ANALISES`, retenção 30 dias só listada. (Aplicam-se aos ciclos 005 a 014.) |

## Critérios de aceite (coluna "Prova" das fatias)

- [ ] F0 `verificar-tudo` padrão mostra no RESUMO as etapas `testar-migracao-003` e `testar-preparacao-rede`.
- [ ] F1 `criar-banco -Recriar`, `criar-banco-teste` e `criar-banco-demo -Recriar` recusam banco PRODUCAO ou ilegível antes de apagar (hash igual); aceitam inexistente e TESTE.
- [ ] F2 `testar-mcp` e `testar-uso` (modo completo) recusam banco não TESTE antes da primeira gravação (hash igual).
- [ ] F3 migração estrutural só com `-CopiaDeTrabalho` dentro de `execucao\publicacao` ou `execucao\ensaio`; PRODUCAO 1.0 sem a opção recusa antes de aplicar; TESTE 1.0 chega ao 1.1; `Invoke-PreparacaoDoBanco` passa a opção.
- [ ] F4 `Invoke-EnsaioDoBanco` e `montar-frontend -Ensaio` aprovam banco "como a rede" (1.0, PRODUCAO, sem `migracao.*`), registram 001/002/003 na cópia, produção com hash igual, relatório gerado.
- [ ] F5 publicação roda o teste de uso só de leitura e compara impressão; `-Teste` roda leitura e depois completo; `-Completo` mostra "só leitura... impressão igual".
- [ ] F6 gancho que simula gravação depois de preparar o `.novo` faz a troca recusar sem trocar nada.
- [ ] F7 falha forçada em cada subpasso do passo 12 devolve raiz, `VERSAO.txt` e `VERSAO-FRONT.txt` (hash igual); sucesso grava na ordem certa; `versao_front` na cópia.
- [ ] F8 manifesto de backup (instante, sha256, esquema, contagens); restauração aprovada = 0, adulterada = 1; escolha pelo manifesto/nome, não por `LastWriteTime`.
- [ ] F9 conector exige 1.1; com 1.0 só `crm_status` responde (`atualizacao_pendente=true`), demais com `isError` e "esquema 1.0"; gravação recusada; `VersaoServidor 1.5.0`.
- [ ] F10 documentação e roteiro atualizados; `verificar.py` sem falha.
- [ ] `verificar-tudo -Completo` sem falha, saída mostrada.

## Stacks tocadas

build-operacao-powershell, banco-access, ia-mcp, ferramentas-dev (specs de stack atualizadas na F10).

## Comandos

`MONTAR-FRONTEND.bat -Ensaio` (novo); `aplicar-migracoes.ps1 -CopiaDeTrabalho <caminho>` (novo); `testar-uso.ps1 -SoLeitura` (novo); `restaurar-banco.ps1 -Backup -Referencia` (ampliado).

## Estrutura

`lib\Comum.ps1`, `lib\AtualizacaoBanco.ps1`, `lib\Backup.ps1` (novo), `build\montar-frontend.ps1`, `build\lib\Pacote.ps1`, `banco\aplicar-migracoes.ps1`, `banco\esquema\criar-banco.ps1`, `testes\testar-uso.ps1`, `testes\testar-preparacao-rede.ps1` (novo), `testes\criar-banco-demo.ps1`, `operacao\scripts\restaurar-banco.ps1`, `IA\mcp\lib\*`, `IA\mcp-dev\lib\Catalogo.ps1`, `IA\teste\*`, documentação.

## Estratégia de teste

Tabela das fatias em [plan.md](plan.md). `tpr` = `testar-preparacao-rede.ps1` (ACE), `tbl` = `testar-build-lib.ps1`.

## Limites

- Sempre: teste que grava só em banco TESTE/fictício; `.ps1` em UTF-8 com BOM; fonte VBA em cp1252/CRLF.
- Perguntar antes: nada neste ciclo (decisões assumidas acima).
- Nunca: tocar `BASE\crm_zapromaq.accdb` ou `execucao\` da raiz; abrir o arquivo de backup; publicar na rede; dado real no repositório.

## Itens para a estação (🖥)

`[IO.File]::Replace` no compartilhamento; nomes das estações lidos da trava; tempo de cópia e migração pela rede (medido no ensaio); `-SoLeitura` com dados reais; 🖥B da 003 (`frmListas`, Visitas, `frmPlanejamento`, aba Planejamentos em 1366x768; CSV no Excel 2019). Roteiro completo em `docs\implantacao.md` (F10).

## Questões em aberto

Nenhuma.
