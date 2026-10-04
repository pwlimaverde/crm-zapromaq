# Levar o sistema para o computador da empresa

Roteiro da cópia desta estação (desenvolvimento) para a máquina da empresa, onde
a manutenção passa a ser feita pelo **Claude Desktop** (não há Claude Code lá).

## Como a BASE vai para a empresa

A BASE inteira vai num pacote `.zip` gerado na máquina de desenvolvimento:

```powershell
powershell -File agent-config\ferramentas\empacotar-base.ps1
# -> dist\BASE-v<versão>-<data>-<commit>.zip
```

O pacote sai do **commit** (só o que está versionado) e já leva as skills do
Claude Desktop em `.zip`. Ele **não** leva o que é só da rede, e por isso essas
coisas continuam lá depois da troca:

| Fica na rede | O que é |
|---|---|
| `crm_zapromaq.accdb` | o banco de produção |
| `CRM_Zapromaq.xlsm`, `VERSAO-FRONT.txt` | a planilha publicada e a versão dela (escritos pelo build) |
| `crm_zapromaq.ico` | ícone do atalho das estações |
| `SISTEMA\DADOS\execucao\` | backups, histórico de publicações, logs, registro do MCP |
| `IA\logs\` | logs dos conectores |

## Passo a passo na empresa

1. **Antes de tudo, traga para cá o que foi alterado lá** (ajustes pelo Claude
   Desktop): mande a BASE da rede compactada para o desenvolvimento e espere ela
   entrar no repositório. O pacote substitui os arquivos de lá; o que foi feito
   na rede e não voltou para o repositório se perde.
2. Todos fechando o CRM (a cópia local em `Documentos` pode ficar aberta; o
   `.xlsm` da rede, não).
3. **Extraia o `.zip` por cima da pasta `BASE` da rede** — na pasta que contém a
   `BASE` (`01 - CRM\01 - CONTROLE`), e responda **"Substituir os arquivos no
   destino"**. **Nunca apague a `BASE` antes**: o banco e a planilha publicada
   estão nela e não vêm no pacote.
4. `SISTEMA\MONTAR-FRONTEND.bat` — confere o ambiente, faz backup, aplica as
   migrações pendentes do banco, monta, compila, roda autoteste e teste de uso e
   publica a versão seguinte. Se qualquer etapa falhar, nada é publicado e a
   planilha de antes continua valendo.
5. As estações atualizam sozinhas pelo `INICIAR-CRM` (ele compara a primeira
   linha de `VERSAO-FRONT.txt` com a versão instalada).

Proteções do build para esse caminho: sem o `crm_zapromaq.accdb` na raiz ele
**não publica**; e a numeração parte da **maior** entre `VERSAO.txt` e a versão
já publicada em `VERSAO-FRONT.txt`, então um pacote com `VERSAO.txt` mais
antigo nunca repete um número (o que faria as estações acharem que não há nada
novo).

Arquivo que saiu do projeto continua na rede depois da extração (ela só
acrescenta e substitui). Os conectores carregam as bibliotecas por nome, então
isso não atrapalha.

## Deixar o Claude Desktop pronto para manter o sistema

Na máquina de quem mantém:

1. `IA\teste\TESTAR-MCP-DEV.bat` — tem de terminar em **TUDO CERTO**.
2. `IA\INSTALAR-MCP.bat` (dados) e `IA\INSTALAR-MCP-DEV.bat` (manutenção),
   rodados **a partir da pasta da rede** (gravam o caminho UNC). Feche o Claude
   Desktop por completo e abra de novo.
3. Instale as duas skills (Settings › Capabilities › *Code execution and file
   creation* ligado; depois *Customize › Skills › Upload a skill* com o `.zip`
   de cada pasta):
   - `IA\instrucoes\crm-zapromaq` — usar o CRM (consultar e alterar registros);
   - `IA\instrucoes\crm-zapromaq-dev` — manter o sistema.

Com isso o Desktop consegue, sozinho: ver o estado do sistema e das versões,
ler e alterar as fontes do front, gerar prévia das telas, **criar e aplicar
migrações do banco** (campo novo, tabela nova, correção de dados), conferir o
projeto, montar um `.xlsm` de teste e publicar a versão nova.

O que ele **não** faz e continua sendo trabalho desta estação: mexer em `build\`,
na própria pasta `IA\` e refatoração grande. Nesse caso, traga a BASE da rede de volta
para o repositório antes de começar (`docs\sincronizar-rede.md`).

## Máquina da empresa: o que precisa ter

| Para | Exige |
|---|---|
| Montar e publicar o `.xlsm` | Excel 2019 64 bits |
| Prévia das telas | Edge (já vem no Windows) |
| Banco e conectores MCP | Office 64 bits ou o Access Database Engine 2016 x64 |
| Conferência estática completa | Python (opcional; sem ele o conector faz a básica e o build ainda compila e roda o autoteste) |
