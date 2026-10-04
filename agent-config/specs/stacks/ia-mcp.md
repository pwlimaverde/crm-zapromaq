# Stack: IA — servidores MCP em PowerShell e skills do Claude Desktop

## Papel

Dois conectores MCP por stdio, iniciados pelo Claude Desktop (e opcionalmente pelo Codex):
- **dados** (`IA\mcp\servidor.ps1`, `crm-zapromaq`): consultar e alterar o CRM — 25 ferramentas.
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
- Mesmas regras de normalização do front (`modValidacao` ⇄ `Regras.ps1`).
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
