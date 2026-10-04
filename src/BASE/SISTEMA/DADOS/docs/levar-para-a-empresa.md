# Levar o sistema para o computador da empresa

Roteiro da cópia desta estação (desenvolvimento) para a máquina da empresa, onde
a manutenção passa a ser feita pelo **Claude Desktop** (não há Claude Code lá).

## O que copiar

Só duas pastas, inteiras, para dentro da `BASE` da rede:

```
SISTEMA\      build, fontes, banco (esquema e migrações), testes, docs
IA\           os dois conectores MCP + as skills do Claude Desktop
```

O que **não** se copia: o `.accdb` e o `CRM_Zapromaq.xlsm` — os de verdade já
estão na rede. `VERSAO-FRONT.txt` também não: quem o escreve é o build, na
primeira publicação.

Esta cópia sai com `SISTEMA\DADOS\execucao\` vazia (só `.gitkeep`), sem logs e
sem o banco de ensaio: nada daqui se mistura com a produção.

## Versão

`SISTEMA\DADOS\VERSAO.txt` está em **1.4**, que é a versão em uso na empresa. O
primeiro `MONTAR-FRONTEND.bat` lá vai publicar a **1.5** e só então gravar 1.5 no
arquivo. O `CHANGELOG.md` está limpo: a primeira linha dele será a 1.5 real.

## Passo a passo na empresa

1. Backup manual do `crm_zapromaq.accdb` e do `CRM_Zapromaq.xlsm` atuais.
2. Copiar `SISTEMA` e `IA` para dentro da `BASE` da rede.
3. Todos fechando o CRM (a cópia local em `Documentos` pode ficar aberta; o
   `.xlsm` da rede, não).
4. `SISTEMA\MONTAR-FRONTEND.bat` — faz tudo: confere o ambiente, backup, aplica
   as migrações pendentes, monta, compila, autoteste, teste de uso e publica a
   1.5. Se qualquer etapa falhar, nada é publicado.
5. Testar a cópia publicada na sua máquina e avisar a equipe para abrir o
   `INICIAR-CRM`.

Detalhes e o que fazer se der errado: `SISTEMA\LEIA-ME-PRIMEIRO.md` e
`docs\implantacao.md`. O `INICIAR-CRM.bat` das estações lê a primeira linha de
`VERSAO-FRONT.txt` (trecho pronto em `docs\versao-para-o-iniciar-crm.md`).

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
na própria pasta `IA\` e refatoração grande. Nesse caso, traga os ajustes feitos
na rede de volta com `ferramentas\sincronizar\trazer-da-rede.ps1`
(`docs\sincronizar-rede.md`) antes de começar.

## Máquina da empresa: o que precisa ter

| Para | Exige |
|---|---|
| Montar e publicar o `.xlsm` | Excel 2019 64 bits |
| Prévia das telas | Edge (já vem no Windows) |
| Banco e conectores MCP | Office 64 bits ou o Access Database Engine 2016 x64 |
| Conferência estática completa | Python (opcional; sem ele o conector faz a básica e o build ainda compila e roda o autoteste) |
