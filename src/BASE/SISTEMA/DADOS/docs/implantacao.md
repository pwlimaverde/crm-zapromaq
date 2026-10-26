# Implantação na rede - roteiro da versão 1.6 (ciclo 004, preparação da rede)

Roteiro para levar a `develop` (spec 003 + ciclo 004) para a rede, onde o banco está no
**esquema 1.0 e ainda não recebeu nenhuma migração** (001, 002 e 003 vão juntas). A versão da rede
é a 1.5; a publicação nova é a **1.6**. Feito **fora do horário de uso**, com todos fora do CRM e do
Claude Desktop. Cada passo tem a conferência que libera o seguinte: não pule conferência.

Legenda: 🖥 = numa estação com Excel 2019 · 🌐 = na pasta da rede

Regra que vale em tudo: **o banco de produção só é alterado pelo `MONTAR-FRONTEND.bat`** (cópia,
migração na cópia, conferências, troca). `APLICAR-MIGRACOES.bat` e a ferramenta `aplicar_migracoes`
recusam migração estrutural no banco de produção, e `CRIAR-BANCO-TESTE`, `TESTAR-MCP` e
`criar-banco-demo -Recriar` recusam banco que não seja TESTE. Mesmo assim, **não rode scripts de
teste na rede**.

## 0. Na máquina de desenvolvimento 🖥

1. `powershell -File agent-config\ferramentas\empacotar-base.ps1` gera `dist\BASE-v<versão>-<data>-<commit>.zip`.
   Levar esse arquivo.

## 1. Janela de implantação 🌐

1. Todos fecham o CRM **e o Claude Desktop**.
2. Conferir que `BASE\crm_zapromaq.laccdb` não existe.

## 2. Conferência de divergência 🌐

1. Abrir `BASE\SISTEMA\DADOS\execucao\dev\alteracoes.log` e procurar `BASE\IA\**\*.bak-*`.
2. Se houver qualquer item com data posterior a 04/10/2026, **não extrair**: compactar a BASE da rede
   e trazê-la para o desenvolvimento (`docs\sincronizar-rede.md`). A versão da rede tem prioridade.

## 3. Backup independente 🌐

1. Copiar a pasta `BASE` **inteira** (`.accdb`, `.xlsm`, `VERSAO-FRONT.txt`, `crm_zapromaq.ico`,
   `execucao\`) para `...\01 - CONTROLE\BACKUP-ANTES-DA-1.6-AAAAMMDD\`.
2. Conferir que o tamanho do `.accdb` copiado é igual ao original.

## 4. Extração 🌐

1. Extrair o `.zip` **por cima** da `BASE`, respondendo "Substituir". **Nunca apagar a BASE antes.**

## 5. Ambiente 🖥

1. `SISTEMA\DADOS\testes\CHECAR-AMBIENTE.bat`: Excel 2019 64 bits, ACE, acesso ao VBA.
2. Em **cada estação** (DEV-02), rodar o mesmo `CHECAR-AMBIENTE.bat` e guardar o relatório
   (`execucao\logs\ambiente-*.txt`, um por estação). Desde o ciclo 018 ele confere também: conector no
   Claude Desktop e no Codex, versão do front instalada x publicada, a pasta `%LOCALAPPDATA%\CRM Zapromaq`
   (e se está sincronizada ou redirecionada para a rede) e a **escrita na pasta do banco** (o ACE cria o
   `.laccdb` ali; sem escrita, a estação não grava). O que depender do servidor vai para a TI como pedido.
3. Preencher a matriz de estações abaixo com o resultado. Publicação que exige pré-requisito novo lista
   antes as estações afetadas.

### Matriz de estações compatíveis

| Estação | Papel | Excel 2019 64 bits | ACE 64 bits | Escrita na pasta do banco | `%LOCALAPPDATA%` local | Front instalado | Claude Desktop / Codex | Acesso ao modelo VBA | Relatório (data) |
|---|---|---|---|---|---|---|---|---|---|
| (estação do build) | build, publicação | obrigatório | obrigatório | obrigatória | sim | — | opcional | **sim** (só aqui) | |
| (estação comercial) | uso do CRM | obrigatório | obrigatório | obrigatória | obrigatório | versão publicada | conforme o uso de IA | **não** | |

Regras da matriz: acesso ao modelo de objeto do VBA só onde há build ou teste; fontes e publicação só para
quem mantém o sistema; backups com leitura restrita (pedido para a TI). Sem nome real de estação no
repositório: a matriz preenchida fica na pasta da rede, junto com os relatórios (`execucao\logs`).

## 6. Backup restaurável 🖥

1. `SISTEMA\DADOS\operacao\TESTAR-RESTAURACAO.bat -Backup "<BACKUP-ANTES-DA-1.6...>\crm_zapromaq.accdb" -Referencia "BASE\crm_zapromaq.accdb"`.
2. Tem de terminar em "RESTAURAÇÃO APROVADA" (código 0). Se reprovar, **parar**. Backup sem manifesto
   (como este, copiado à mão) é conferido pela abertura da cópia restaurada e pela comparação com a
   referência.

## 7. Ensaio 🖥

1. `SISTEMA\MONTAR-FRONTEND.bat -Ensaio` (precisa do banco fechado; não precisa de Excel; nunca troca o banco).
2. Esperado: `ENSAIO APROVADO`, esquema 1.0 -> 1.1 na cópia, migrações 001, 002 e 003 registradas na
   cópia, hash do banco de produção igual antes e depois.
3. Ler `SISTEMA\DADOS\execucao\ensaio\relatorio-*.txt`: unificações de listas, valores incluídos como
   inativos, contagens e tempo. **Esse relatório tem dado real: não vai para o repositório.**
4. Se o tempo do ensaio pela rede for alto, planejar a janela com folga.

## 8. Publicação 🖥

1. `SISTEMA\MONTAR-FRONTEND.bat` (sem `-Teste`, sem `-Ensaio`).
2. Esperado: `PRONTO. CRM 1.6 publicado; banco no esquema 1.1 ...`. O teste de uso da publicação é só de
   leitura e confere que a cópia do banco não mudou.

## 9. Conferências 🖥

1. `VERSAO-FRONT.txt`: linha 1 = `1.6`; linha 2 = `Ambiente: PRODUCAO   |   Versao 1.6   |   Publicada em dd/MM/yyyy HH:mm`.
2. No log `execucao\logs\build-*.txt`: a linha `Inicio!B7 = ...` é idêntica à linha 2 e aparece
   `config.versao_front = 1.6`.
3. Na estação piloto, pelo `INICIAR-CRM`: instala a 1.6; `Início!B7` igual à linha 2; as grades carregam; a ficha abre.

## 10. Claude Desktop 🖥

1. `IA\PUBLICAR-CONECTOR.bat` (ciclo 007): copia o conector para `IA\conector\versoes\<versão>`, roda o
   `testar-mcp` num banco de homologação temporário e troca o ponteiro `IA\conector\ATUAL.txt`.
2. Em cada estação com IA: `IA\INSTALAR-MCP.bat` (ou `INSTALAR-CLAUDE-E-CODEX.bat`) uma vez, para registrar o
   lançador fixo `IA\mcp\lancador.ps1`; reabrir e reinstalar `IA\instrucoes\crm-zapromaq.zip`.
3. `crm_status` deve mostrar `versao_conector` igual à do ponteiro, `pasta_conector` em `IA\conector\versoes\...`,
   `quantidade_ferramentas 41`, `versao_esquema 1.1`, `gravacao_liberada true` e `atualizacao_pendente false`. Com o banco ainda no 1.0 o conector responde só o `crm_status`, com
   `atualizacao_pendente true`; as outras ferramentas dizem "banco aguardando atualização".

## 11. Backups da rotina 🖥

1. `SISTEMA\DADOS\operacao\AGENDAR-BACKUPS.bat`, depois `FAZER-BACKUP.bat`, depois `TESTAR-RESTAURACAO.bat`.
2. **Manual de restauração cronometrado (DEV-03):** com cronômetro, restaurar a última cópia validada numa
   pasta de homologação (`TESTAR-RESTAURACAO.bat`) e anotar aqui: tempo total, instante do corte recuperado
   (`instante_local` do `.manifesto.json`) e perda potencial (tempo desde o corte). A gestão define RPO, RTO e
   quem responde ao aviso de backup; cópia fora do servidor é pedido à TI (ou limitação aceita, registrada).
3. Exclusividade (ciclo 007): durante a cópia completa (segundos), quem tentar gravar recebe "Cópia de
   segurança em andamento… nada foi gravado"; ler continua normal. Aviso no Início e no `crm_status` quando
   uma janela (9h, 12h30, 16h) termina sem cópia ou o dia útil termina sem cópia completa validada.

## 12. Itens 🖥 da spec 003 e do ciclo 004

1. Telas em 1366x768: `frmListas`, Visitas, `frmPlanejamento` e a aba Planejamentos; abrir o CSV do
   roteiro no Excel 2019.
2. Anotar: `[IO.File]::Replace` no compartilhamento funcionou; os nomes das estações apareceram na
   mensagem de trava quando havia alguém conectado; tempos de cópia e de migração pela rede.

## 13. De volta ao desenvolvimento

1. Registrar no `PLANO.md` o resultado da Fase 8.
2. `git flow release start 1.6`, depois `finalizar-branch.ps1 -Branch release/1.6 -Enviar` (main + tag `v1.6`).
   A `main` ainda não tem tag da 1.5: registrar no CHANGELOG qual commit gerou a 1.5 na rede, se for identificado.

## Se algo falhar

- **Ensaio ou passos 1 a 10 da montagem:** nada foi alterado. Corrigir a causa e repetir.
- **Passo 11 (troca do banco):** a mensagem diz que nada foi trocado. Se disser que a produção não
  conferiu, fechar todos e copiar `execucao\backup\antes-da-publicacao\<data-hora>\crm_zapromaq.accdb`
  para a raiz.
- **Passo 12 antes do `VERSAO-FRONT.txt`:** a montagem devolve sozinha o `.xlsm` anterior, o
  `VERSAO.txt` e o banco. O log diz o que foi devolvido. O `VERSAO-FRONT.txt` é gravado por último:
  depois dele a publicação vale (ponto de não retorno).
- **Problema depois de publicado:**
  - **volta do front (DEV-04, ciclo 007):** extrair por cima da BASE o pacote da versão boa e rodar
    `SISTEMA\MONTAR-FRONTEND.bat -Volta -Motivo "<por que>"`: publica aquele conteúdo como versão NOVA, maior
    (as estações recebem pelo atalho), e registra em `execucao\releases\release-<versão>.json`. É recusada se o
    front antigo esperar outro esquema (o banco não volta de esquema sobre gravações novas);
  - **volta do conector (DEV-05):** `IA\PUBLICAR-CONECTOR.bat -Voltar <versão> -Motivo "..."`;
  - preferir corrigir para frente (hotfix 1.7): a migração só acrescenta e o front 1.5 recusa gravar no 1.1;
  - volta total só se ninguém gravou ainda: todos fecham; copiar da pasta `antes-da-publicacao` o banco
    e o `.xlsm`, e restaurar o `VERSAO-FRONT.txt` do backup do passo 3. Em cada estação que já tinha a
    1.6, apagar `%LOCALAPPDATA%\CRM Zapromaq\CRM_Zapromaq.xlsm` (até a 1.6: `Documentos\CRM Zapromaq`): o
    `INICIAR-CRM` reinstala quando não há cópia local;
  - se já houve gravação no 1.1, não voltar o banco; corrigir para frente.

## Depois

- Acompanhar os logs em `SISTEMA\DADOS\execucao\logs\` e `IA\logs\` na primeira semana.
- Ajustes pequenos: MCP de desenvolvimento na rede; mudanças maiores: repositório. Em ambos os casos,
  a regra de sincronização (`docs\sincronizar-rede.md`).
