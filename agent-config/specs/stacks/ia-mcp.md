# Stack: IA — servidores MCP em PowerShell e skills do Claude Desktop

## Papel

Dois conectores MCP por stdio, iniciados pelo Claude Desktop (e opcionalmente pelo Codex):
- **dados** (`IA\mcp\servidor.ps1`, `crm-zapromaq`): consultar e alterar o CRM — 41 ferramentas (16 novas da spec 003: visitas,
  planejamentos, candidatos, lotes, municípios e roteiro; esquemas 1.0 e 1.1 aceitos).
- **desenvolvimento** (`IA\mcp-dev\servidor-dev.ps1`, `crm-zapromaq-dev`): manter o sistema na
  empresa — fontes, prévia, migrações, conferência, montagem de teste e publicação.

## Versões e ambiente

PowerShell 5.1; ACE para o MCP de dados. Protocolo **dual-era**: legado (`initialize`,
≤ 2025-11-25) e 2026-07-28 (sem estado: `server/discover`, `_meta` por requisição, `resultType`,
`ttlMs`/`cacheScope`). Núcleo comum: `IA\mcp\lib\Protocolo.ps1`. Sem fila nem agendador.

## Onde fica

```
src/BASE/IA/
  mcp/servidor.ps1 + lib/{Protocolo,Banco,Regras,Ferramentas}.ps1 + instalar-mcp.ps1
  mcp-dev/servidor-dev.ps1 + lib/{Fontes,Gravacao,Previa,Processos,Migracoes,Conferencia,Catalogo}.ps1
  instrucoes/crm-zapromaq/SKILL.md, instrucoes/crm-zapromaq-dev/SKILL.md   (o .zip é gerado pelo pacote)
  INSTALAR-MCP.bat, INSTALAR-MCP-DEV.bat, INSTALAR-CLAUDE-E-CODEX.bat
  teste/  criar-banco-teste.ps1, testar-mcp.ps1, testar-mcp-dev.ps1 (+ .bat)
  logs/   GERADO (fora do git)
```

Ferramentas de dados: `crm_status`, `listar_opcoes`, `buscar_clientes`, `obter_cliente`,
`criar_cliente`, `promover_cliente`, `alterar_situacao_cliente`, `atualizar_cadastro_cliente`,
`atualizar_contexto_cliente`, `buscar_contatos`, `obter_contato`, `criar_contato`,
`atualizar_contato`, `buscar_atendimentos`, `obter_atendimento`, `abrir_atendimento`,
`atualizar_atendimento`, `atualizar_contexto_atendimento`, `atendimentos_por_interacao`,
`atendimentos_por_proxima_acao`, `ultimas_alteracoes`, `gerenciar_lista`, `listar_metas`,
`gravar_meta`, `exportar_base`.
Desenvolvimento: `estado_sistema`, `listar_fontes`, `ler_fonte`, `gravar_fonte`, `gerar_previa`,
`verificar_layout`, `verificar_projeto`, `montar_teste`, `publicar`, `listar_alteracoes`,
`listar_migracoes`, `ler_migracao`, `criar_migracao`, `aplicar_migracoes`, `ver_log`.

## Comandos

```powershell
src\BASE\IA\teste\TESTAR-MCP.bat -SoProtocolo           # sem banco
src\BASE\IA\teste\CRIAR-BANCO-TESTE.bat ; src\BASE\IA\teste\TESTAR-MCP.bat
src\BASE\IA\teste\TESTAR-MCP-DEV.bat [-SemPrevia]
src\BASE\IA\INSTALAR-MCP.bat | INSTALAR-MCP-DEV.bat | INSTALAR-CLAUDE-E-CODEX.bat   # na rede, por usuário
```

## Estilo e regras

- **Nada no stdout além de JSON-RPC** (diagnóstico vai para stderr/log). `.ps1` UTF-8 com BOM.
- Catálogo fechado; cada ferramenta com esquema de entrada (parâmetro desconhecido é recusado) e
  anotações (`readOnlyHint`, `destructiveHint`...). Erro de ferramenta volta como `isError`, não
  como erro de protocolo.
- Gravação: confere `config.versao_esquema` antes; `UPDATE` com versão; log com `origem = 'IA'` e
  autor. Os campos `ctx_*` são escritos **só** por aqui.
- Mesmas regras de normalização do front (`modValidacao` ⇄ `Regras.ps1`) e dos vínculos: troca de
  contato e atendimento anterior (`abrir_atendimento`, `codigo_atendimento_anterior`) só dentro
  da mesma empresa (`modCRM.CriticarTrocaContato`/`CriticarAnterior` no front).
- O MCP de desenvolvimento grava só em `SISTEMA\DADOS\fonte`, guarda cópia e registra em
  `execucao\dev\alteracoes.log`; publicar exige confirmação.
- Bibliotecas do mcp-dev carregadas por **lista fixa** de nomes (arquivo sobrando na rede é inofensivo).

## Testes

`testar-mcp.ps1` faz o papel do cliente (as duas gerações do protocolo, leitura, gravação,
conflito de versão) e restaura o banco no fim; `testar-mcp-dev.ps1` roda numa cópia. Ferramenta
nova = caso novo no teste + contagem de ferramentas atualizada (`instalar-mcp.ps1` mostra o total).

## Limites

- Sempre: migração de esquema atualiza o MCP na mesma versão; skill (`SKILL.md`) atualizada
  quando ferramenta muda.
- Perguntar antes: ferramenta nova que grava ou apaga; mudar nome/esquema de ferramenta existente.
- Nunca: escrever no stdout fora do JSON-RPC; SQL montado por concatenação; apagar registro com
  vínculo; fila/agendador que exija administrador.

## Revisão

- [ ] Esquema de entrada fechado e validado; anotações corretas.
- [ ] Gravação com versão, log e transação; leitura sem efeito colateral.
- [ ] Teste cobre a ferramenta; SKILL.md e contagem atualizadas.

## Referências offline

`referencias/` 12-mcp-especificacao, 13-claude-desktop, 09-powershell-5.1, 08-dotnet-oledb.

## Esquema suportado (spec 004)

- O conector (1.5.0) exige o esquema 1.1 (`EsquemasSuportados = @('1.1')`). Com outro esquema só o `crm_status`
  responde (`atualizacao_pendente`, `esquemas_suportados`, `gravacao_liberada`); as demais ferramentas voltam
  `isError` com "esquema X". O `Protocolo` passa o nome da ferramenta ao `AntesDaFerramenta`.
- `aplicar_migracoes` (MCP de desenvolvimento) recusa migração estrutural em banco de produção.

## Contrato, manutenção e publicação por versão (ciclo 007)

- Contrato exato: `testar-mcp.ps1` tem `$script:FerramentasEsperadas`; o `verificar.py` cruza com o catálogo e com a
  contagem do `instalar-mcp.ps1`. Ferramenta nova = lista + nota de mudança no `IA\LEIA-ME.md` + `VersaoServidor` novo.
- `Invoke-Transacao` confere `config.manutencao` e o esquema antes da 1ª escrita (`ConferirEscrita`, dentro da
  transação); erro de escrita com `execucao\backup\copiando.txt` presente vira a mensagem controlada do backup.
- Erros lançados dentro de método de objeto são desembrulhados antes de voltar à IA (`Get-MensagemInterna`).
- Diagnóstico: mês corrente e anterior pela data do nome; escrita compartilhada com nova tentativa; falhas no `crm_status`.
- Publicação: `publicar-conector.ps1` (pasta de versão imutável + ponteiro `IA\conector\ATUAL.txt`), lançador fixo
  `lancador.ps1` registrado nos clientes; `testar-conector-versoes.ps1` prova publicação, interrupção e volta.
- O aviso de backup por janela é cópia fiel de `SISTEMA\DADOS\lib\Backup.ps1` (o `verificar.py` compara).
- Ciclo 008: `ConvertTo-Moeda`/`ConvertTo-Inteiro` (Regras.ps1) em vez de `[decimal]`/`[int]`; `Invoke-Transacao` repete
  a transação inteira só em bloqueio (`Test-ErroTransitorio`), antes do commit; falha no commit vira "resultado indeterminado".
- Ciclo 009 (1.7.0): `ConvertTo-Normalizado` da seção 5.1 (atenção: função que devolve lista precisa de `@()` no chamador);
  `Get-TransicaoEtapa`, `Test-RetomadaPermitida`, `Test-ContatoDoAtendimento`; `atualizar_contato` sem transferência.
- Ciclo 010 (1.8.0): esquema 1.2 (`EsquemasSuportados = @('1.2')`); `ctx_versao` exigida em `atualizar_contexto_*`
  (conflito devolve o contexto atual); `codigo_atendimento_anterior` sem ciclo; `id_operacao` nas criações (tabela
  `operacoes`, 60 dias); log com `estacao`, `usuario_tecnico`, `versao_cliente` e `autor_declarado`.
- Ciclo 012 (1.9.0): situação "Retomada a avaliar" (Get-Situacao); planejamento com `retomadas_a_avaliar` e paginação
  por chave (`Get-ChavePlanejamento` ordinal + `cursor` com filtro e geração da base; `Array.Sort(chaves, itens)` pelo
  PowerShell não reordena os itens - ordenar as chaves e buscar por dicionário); extratos em `execucao\extratos`
  (`New-PastaExtrato`, manifesto, `descartar_extratos`); `exportar_base` sem `destino`.
- Ciclo 013: matriz de consumidores `IA\teste\consumidores.json` (skill -> ferramenta -> parâmetros) conferida contra o
  `tools/list` no testar-mcp; atualizar `versao_conector` dela a cada `VersaoServidor` novo.
- Ciclo 016 (1.10.0): `listar_candidatos.sem_atendimento` (aditivo, até 200) em `Get-CandidatosPlanejamento`.
