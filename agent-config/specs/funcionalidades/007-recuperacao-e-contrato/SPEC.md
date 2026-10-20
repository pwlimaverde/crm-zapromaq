# Spec: Recuperação, manutenção e contrato do conector (ciclo 007)

> Concluída em 19/10/2026 - feature/recuperacao-e-contrato (plano revisado de 19/10/2026, seção 5, Etapa 1)

## Objetivo

Resto da Etapa 1 do *Plano de Desenvolvimento CRM PW* (19/10/2026) que é só código:

- **DEV-07** — teste de contrato com a lista exata das 41 ferramentas; `crm_status` informa versão e capacidades.
- **DEV-08** (guarda mínima) — sinal comum de manutenção (`config.manutencao`) respeitado no ponto central de escrita do front e do conector, junto com o esquema; a montagem liga o sinal enquanto mexe no banco.
- **DEV-03** (resto) — exclusividade efetiva durante toda a cópia do backup; recusa controlada para quem tenta gravar no meio da cópia; "janela atendida" para disparos simultâneos; aviso por janela da seção 5.3.
- **DEV-06 parte A** — diagnóstico do front em `EXECUCAO\logs\`, retenção do mês corrente e do anterior nos dois canais, falha ao gravar o diagnóstico do conector aparece no `crm_status`.
- **DEV-04** (resto) — volta como nova versão maior com motivo registrado; manifesto de release único (commit, versões do front, conector e esquema, migrações, hashes).
- **DEV-05** — conector publicado em pasta de versão imutável, lançador fixo que lê um ponteiro; volta = reapontar; volta incompatível recusada.

A publicação 1.6 não é feita aqui (depende do banco real e da estação): fica como item 🖥.

## Critérios de aceite

- [x] DEV-07: `testar-mcp` falha se faltar ou sobrar ferramenta em relação à lista exata (41); cada ferramenta tem título, descrição, `inputSchema` e `annotations`; `crm_status` devolve `versao_conector`, `capacidades` (nomes do catálogo) e `quantidade_ferramentas`; `instalar-mcp.ps1` anuncia a mesma contagem do catálogo (verificar.py).
- [x] DEV-08: com `config.manutencao` preenchida, toda gravação do conector é recusada com "manutenção" e nada muda (hash); leitura continua; no front, `modDB.AbrirTransacao` confere esquema e manutenção **dentro** da transação e recusa (mensagem pura `modDB.MotivoRecusaEscrita`, `Confere`); a ficha mantém a digitação (caminho de erro existente). O `aplicar-migracoes` (no lugar) liga o sinal com dono e desliga no fim; migração estrutural interrompida deixa o sinal ligado (falha segura) até `-LiberarManutencao`. A montagem **não** grava o sinal na produção (contrato D22: falha antes do passo 11 deixa a produção byte a byte igual); ali a proteção é a trava livre + hash H0 + esquema conferido no ponto de escrita (ficha aberta antes da troca não grava no esquema novo).
- [x] DEV-03: o backup segura o arquivo do banco com `FileShare.Read` durante a cópia inteira (cópia feita pelo próprio fluxo aberto); com o arquivo seguro, uma conexão ACE abre só leitura e a gravação falha — o conector devolve a mensagem controlada "cópia de segurança em andamento… nada foi gravado" (marca `execucao\backup\copiando.txt`), e o front traduz o mesmo erro (`modDB.MensagemEscritaRecusada`, `Confere`).
- [x] DEV-03: quatro disparos simultâneos `-Modo Agendado` para a mesma janela geram **uma** captura; os demais registram "janela atendida" e saem com 0; nenhum arquivo sobrescrito; logs por execução íntegros.
- [x] DEV-03: aviso por janela (seção 5.3), igual em `lib\Backup.ps1` (estado do sistema), no `crm_status` e no front (`modBackupAviso`): (a) a última janela encerrada (9h, 12h30, 16h, seg-sex; encerrada 30 min depois do início) sem nenhuma cópia; (b) dia útil cujas 16h30 já passaram sem cópia **completa validada** (cópia simples não apaga o aviso b). Estação desligada num dia útil → aviso no próximo acesso.
- [x] DEV-06A: o front grava `EXECUCAO\logs\front-<ESTACAO>-AAAAMM.log` com instante, estação, usuário técnico, versão, tela/operação, registro, erro nativo e resultado, sem conteúdo comercial; apaga os próprios arquivos anteriores ao mês anterior; falha de escrita não bloqueia. O conector guarda só o mês corrente e o anterior (pelo nome) e o `crm_status` informa `diagnostico.falhas` e `diagnostico.ultima_falha`.
- [x] DEV-04: `MONTAR-FRONTEND.bat -Volta -Motivo "<texto>"` publica o conteúdo do pacote extraído (versão do `MANIFESTO.json`) como **nova versão maior**, recusa sem motivo, sem pacote, com pacote de versão ≥ publicada, ou com esquema esperado pela fonte diferente do banco; toda publicação grava `execucao\releases\release-<versão>.json` com commit, versões do front, conector e esquema, migrações aplicadas, hashes do `.xlsm` e dos scripts de `build\` e `operacao\`, tipo (`publicacao`|`volta`) e motivo. Prova: `testar-build-lib`.
- [x] DEV-05: `IA\mcp\publicar-conector.ps1` copia `servidor.ps1`, `lib\` e a skill para `IA\conector\versoes\<versão>\` (temporária → conferida → renomeada; pasta existente com conteúdo diferente é recusada) e só promove o ponteiro `IA\conector\ATUAL.txt` (troca atômica) depois do `testar-mcp` num banco de homologação temporário; `-Voltar <versão>` reaponta e recusa versão que não suporta o esquema do banco; `IA\mcp\lancador.ps1` (registrado pelo instalador) lê o ponteiro uma vez e carrega a pasta; sem ponteiro, carrega `IA\mcp\servidor.ps1`. `crm_status` informa `pasta_conector`. Prova: `testar-conector-versoes.ps1`.

## Stacks tocadas

front-vba, ia-mcp, build-operacao-powershell, banco-access (só dado de `config`, sem migração).

## Estrutura

- `IA\mcp\lib\Banco.ps1`, `Protocolo.ps1`, `Ferramentas.ps1`, `servidor.ps1` (1.6.0), `lancador.ps1` (novo), `publicar-conector.ps1` (novo), `instalar-mcp.ps1`, `IA\PUBLICAR-CONECTOR.bat` (novo).
- `IA\teste\testar-mcp.ps1`, `IA\teste\testar-conector-versoes.ps1` (novo).
- `SISTEMA\DADOS\lib\Backup.ps1`, `AtualizacaoBanco.ps1`, `operacao\scripts\backup-banco.ps1`, `banco\aplicar-migracoes.ps1`.
- `SISTEMA\DADOS\build\montar-frontend.ps1`, `build\lib\Release.ps1` (novo), `SISTEMA\MONTAR-FRONTEND.bat`.
- `SISTEMA\DADOS\testes\testar-build-lib.ps1`, `testar-recuperacao.ps1` (novo).
- Front: `modDB.bas`, `modDiag.bas` (novo), `modBackupAviso.bas`, `modAutoteste.bas`, `frmCRM.txt`.
- `agent-config\ferramentas\verificar-tudo.ps1` (etapas novas), `verificar.py` (contagem do instalador).
- Docs: `contrato-banco.md`, `implantacao.md`, `IA\LEIA-ME.md`, specs de stack, `PLANO.md`.

## Estratégia de teste

| Fatia | Teste que falha primeiro |
|---|---|
| DEV-07 | `testar-mcp.ps1` (lista exata, `crm_status.capacidades`) e `verificar.py` (contagem do instalador) |
| DEV-08 | `testar-mcp.ps1` (manutenção recusa gravação) e `Confere` `MotivoRecusaEscrita` |
| DEV-03 | `testar-recuperacao.ps1` (exclusividade com ACE real, janela atendida com 4 processos, aviso por janela) e `Confere` no front |
| DEV-06A | `testar-recuperacao.ps1` (retenção e `crm_status.diagnostico`) e `Confere` `NomeArquivoDiag`/`DeveApagarDiag` |
| DEV-04 | `testar-build-lib.ps1` (`Test-PedidoVolta`, `New-ManifestoRelease`) |
| DEV-05 | `testar-conector-versoes.ps1` |

Itens 🖥: publicação 1.6 pela exceção da seção 4; ensaio de restauração cronometrado; volta do front e do conector ensaiadas na estação; aviso de backup na abertura do front.

## Decisões assumidas

| # | Decisão | Premissa | O que mudaria |
|---|---|---|---|
| 007-1 | Exclusividade = o backup abre o `.accdb` com `FileShare.Read` e copia pelo próprio fluxo. Experimento aqui: com o arquivo assim seguro, o ACE abre a conexão **só leitura** e a gravação falha com "A operação deve usar uma consulta atualizável". | É o único bloqueio efetivo sem administrador nem serviço; ler durante o backup continua possível. | Se a TI oferecer cópia de sombra (VSS), troca-se o mecanismo, mantendo a marca e a mensagem. |
| 007-2 | A mensagem controlada depende da marca `execucao\backup\copiando.txt` (criada antes de segurar o arquivo, apagada depois). Sem a marca, o erro "consulta atualizável" segue o caminho normal de erro. | O mesmo erro sem backup tem outra causa (DEV-14). | — |
| 007-3 | Achado registrado para o DEV-14: as quatro ocorrências de "consulta atualizável" coincidem com o padrão de quem lê o arquivo com compartilhamento só de leitura (cópia de arquivo), e uma delas foi às 16h20 (tentativa de backup). | Hipótese reforçada pelo experimento; não reproduzida com o backup antigo. | Confirmação depende do diagnóstico da parte A em produção. |
| 007-4 | Janela de backup encerra 30 min depois do início (tentativas em 0, 10 e 20 min + 10 min de cópia e validação); uma cópia conta para a janela se o instante estiver em [início, início + 60 min). Dia útil = segunda a sexta (feriado não é conhecido). | Horários 9h, 12h30, 16h e tentativas do `agendar-backups.ps1`. | Feriado gera aviso falso; uma lista de feriados em `config` resolveria. |
| 007-5 | Aviso (a) olha só a última janela encerrada; aviso (b) olha o último dia útil cujas 16h30 já passaram. | É o que alguém precisa saber ao abrir o sistema. | — |
| 007-6 | "Janela atendida": execução `Agendado` que encontra a trava ocupada espera a outra terminar (até `-EsperaTravaSeg`, padrão 25 min) e, se já existe cópia daquela janela (manifesto com `janela`), registra "janela atendida" e sai com 0; manual e outras sem janela continuam saindo com 4. | Disparos simultâneos são tarefas de estações diferentes para o mesmo horário. | — |
| 007-7 | Sinal de manutenção = `config.manutencao` (texto `estação\|usuário\|PID\|início\|motivo`), sem migração. Dono morto não libera sozinho (falha segura); `aplicar-migracoes.ps1 -LiberarManutencao` limpa à mão. A montagem não usa o sinal (gravar na produção quebraria o hash H0 e a garantia "falha antes do passo 11 não altera a produção"). | Liberar sozinho poderia abrir a escrita no meio de uma migração interrompida. | Se a montagem passar a migrar no lugar, ela liga o sinal antes do H0. |
| 007-8 | O front confere esquema e manutenção numa consulta dentro da transação, em `modDB.AbrirTransacao` (todas as gravações passam por ali). | Uma consulta a mais por gravação (~ms) é barata. | — |
| 007-9 | Diagnóstico do front: um arquivo por estação e por mês; o conector mantém um por estação e por dia; retenção pelos dois = mês corrente e anterior, pela data do nome. | Seção 5.3. | — |
| 007-10 | Volta (DEV-04) = extrair o pacote anterior por cima da BASE e rodar `MONTAR-FRONTEND.bat -Volta -Motivo "..."`; a versão sobe (nunca repete) e o manifesto registra `conteudo_de`. | O pacote anterior existe em `dist\` / no repositório (tag). | — |
| 007-11 | DEV-05: pasta `IA\conector\` é da rede (fora do git e do pacote); o pacote continua levando `IA\mcp\` como fonte; quem publica roda `IA\PUBLICAR-CONECTOR.bat` depois da montagem. | Separar fonte de versão em uso. | Se a montagem tiver de publicar o conector junto, o passo entra no fim do `montar-frontend`. |
| 007-12 | `VersaoServidor` sobe para 1.6.0 (mudança aditiva no `crm_status`; recusa por manutenção é erro novo). Nota de mudança no `IA\LEIA-ME.md`. | Contrato só aditivo. | — |

## Limites

- Sempre: teste primeiro; teste que grava só em banco TESTE de homologação, em pasta temporária; nada de dado real.
- Nunca: publicar na rede, tocar o banco de produção, alterar `BASE\crm_zapromaq.accdb` ou `execucao\` da raiz.

## Questões em aberto

Nenhuma: RPO/RTO e responsável pelo alerta são da gestão (item 🖥); cópia fora do servidor é da TI.

## Itens 🖥 (estação de implantação)

1. Publicação 1.6 pela exceção da seção 4 do plano: todos os clientes e conectores fechados, acesso exclusivo
   comprovado, cópia validada imediatamente antes (`FAZER-BACKUP.bat` + `TESTAR-RESTAURACAO.bat`), `MONTAR-FRONTEND.bat -Ensaio`,
   depois `MONTAR-FRONTEND.bat`; conferir `execucao\releases\release-1.6.json` (commit do pacote, esquema 1.1, migrações 001-003).
2. `IA\PUBLICAR-CONECTOR.bat` e, em cada estação com IA, `IA\INSTALAR-MCP.bat` (registra o lançador); `crm_status`
   com `pasta_conector` em `IA\conector\versoes\1.6.0` e 41 ferramentas.
3. Ensaio de restauração cronometrado (manual em `docs\implantacao.md` §11) e RPO/RTO definidos pela gestão.
4. Aviso de backup: abrir o front com a pasta de backup sem a cópia da última janela e ver a mensagem; `crm_status` com `motivos_backup`.
5. Exclusividade: durante `FAZER-BACKUP.bat -Modo Diario`, tentar salvar uma ficha e ver "Cópia de segurança em andamento" (a ficha mantém o que foi digitado).
6. Manutenção: `config.manutencao` ligada por `aplicar-migracoes` num banco TESTE; tentar salvar na ficha e no Claude Desktop e ver a recusa.
7. Diagnóstico: forçar um erro de gravação (ex.: banco TESTE só leitura) e conferir `execucao\logs\front-<ESTACAO>-AAAAMM.log`.
8. Volta do front e do conector ensaiadas numa cópia da BASE (`-Volta -Motivo`; `PUBLICAR-CONECTOR.bat -Voltar`).
