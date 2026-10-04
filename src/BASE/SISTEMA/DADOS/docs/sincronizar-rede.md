# Sincronização rede ↔ desenvolvimento

O front tem **duas origens de alteração**:

1. a máquina de desenvolvimento (repositório git, Claude Code) — mudanças grandes;
2. a pasta da rede, pelo **MCP de desenvolvimento** no Claude Desktop da empresa
   (`IA\mcp-dev`) — ajustes pequenos, feitos e publicados na hora.

Sem regra, a próxima cópia do repositório para a rede apagaria os ajustes feitos
lá. A regra (PLANO, regra de trabalho 6):

> **Antes de qualquer desenvolvimento novo, a rede volta para o repositório.**

## Da rede para o repositório (início de cada ciclo)

Na máquina de desenvolvimento, com acesso à pasta da rede:

```powershell
# 1. ver o que mudou, sem copiar nada
powershell -File BASE\SISTEMA\DADOS\ferramentas\sincronizar\trazer-da-rede.ps1 -Rede \\servidor\...\BASE -Simular
# 2. trazer
powershell -File BASE\SISTEMA\DADOS\ferramentas\sincronizar\trazer-da-rede.ps1 -Rede \\servidor\...\BASE
# 3. conferir e registrar
git diff
python BASE\SISTEMA\DADOS\ferramentas\verificacao\verificar.py
git commit -am "Ajustes feitos na rede em dd/mm (MCP de desenvolvimento)"
```

O script compara `SISTEMA\DADOS\fonte\` pelos hashes: copia o que é novo ou
diferente na rede e só **lista** o que existe apenas no repositório (nunca apaga).
Ele também mostra o registro `execucao\dev\alteracoes.log` da rede — quem mudou o
quê e por quê — que ajuda a escrever a mensagem do commit.

## Do repositório para a rede (fim de cada ciclo)

Só depois do passo acima. Copie para a rede as pastas de `BASE\SISTEMA\DADOS`
que vêm do repositório (`fonte`, `build`, `banco`, `operacao`, `lib`, `docs`,
`testes`, `ferramentas`) e a pasta `BASE\IA` — **nunca** `execucao\` (é da rede:
logs, backups, histórico de publicações, registro do MCP) nem o `.accdb`/`.xlsm`
da raiz. Em seguida, `SISTEMA\MONTAR-FRONTEND.bat` na rede publica a versão nova.

## Conflitos

Se o mesmo arquivo foi alterado nos dois lados, o `trazer-da-rede` sobrescreve a
cópia do repositório com a da rede, e o `git diff` mostra a diferença contra o
último commit: junte as duas alterações à mão antes do commit. As cópias
anteriores de cada gravação feita na rede ficam em
`execucao\dev\copias\<data-hora>\` na própria rede.
