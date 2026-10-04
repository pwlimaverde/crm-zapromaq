# Publicar o CRM na rede — o que fazer

Vale para a **primeira troca** (sair do front legado) e para toda publicação
seguinte. Quem faz: uma estação com **Excel 2019 64 bits**.

## Onde cada coisa fica

```
BASE\                        (a pasta da rede que já existe)
  crm_zapromaq.accdb         o banco de produção — fica onde está
  CRM_Zapromaq.xlsm          o front; o MONTAR-FRONTEND substitui por aqui
  INICIAR-CRM.bat            atalho das estações (fora deste projeto): compara a
                             versão e copia o .xlsm novo para Documentos
  SISTEMA\                   ← a pasta que você copia para cá
  IA\                        ← opcional: conector do Claude Desktop
```

## Primeira troca

1. **Backup manual seu**, por segurança: copie `crm_zapromaq.accdb` e o
   `CRM_Zapromaq.xlsm` atual para uma pasta sua (o build também faz backup
   sozinho, mas esta cópia é a sua rede de proteção).
2. **Copie a pasta `SISTEMA` inteira** para dentro de `BASE` (e a `IA`, se for
   usar o Claude Desktop).
3. **Peça para todos fecharem o CRM.** As cópias locais em `Documentos` podem
   ficar abertas; o que não pode é o `.xlsm` da rede aberto.
4. Abra a pasta `SISTEMA` e **execute `MONTAR-FRONTEND.bat`**. Ele faz tudo:
   confere o ambiente, faz backup do banco e do front, aplica as migrações
   pendentes, monta o `.xlsm` novo, compila, roda o autoteste, confere o pacote,
   roda o teste de uso com os dados reais e só então publica na raiz, subindo a
   versão (1.4 → 1.5).
   **Se qualquer etapa falhar, nada é publicado**: o `.xlsm` da raiz continua o
   de antes e o backup fica guardado.
5. **Teste você primeiro:** copie o `CRM_Zapromaq.xlsm` publicado para a sua
   máquina, abra e confira as abas, a ficha e o Painel.
6. Estando tudo certo, **avise a equipe para abrir o `INICIAR-CRM`**. Ele vê que
   a versão subiu, copia o arquivo novo para `Documentos` de cada um e abre.

## Publicações seguintes

Só o passo 4, e depois o 5 e o 6. A sequência é a mesma.

## Se algo der errado depois de publicar

- **Voltar o front:** as versões anteriores ficam em
  `SISTEMA\DADOS\execucao\historico-front\`. Copie a que você quer de volta para
  a raiz como `CRM_Zapromaq.xlsm` e peça para abrirem o `INICIAR-CRM` de novo.
- **Voltar o banco:** a cópia de antes da publicação está em
  `SISTEMA\DADOS\execucao\backup\antes-da-publicacao\<data-hora>\`.
  Atenção: o que foi lançado depois da publicação está só no banco atual;
  restaurar o banco desfaz esses lançamentos.
- **O log de cada publicação** fica em `SISTEMA\DADOS\execucao\logs\build-*.txt`.
  Se falhar, é o primeiro lugar para olhar (e o que me mandar).

## Rotinas que valem a pena logo depois

- `SISTEMA\DADOS\operacao\AGENDAR-BACKUPS.bat` — backup automático do banco.
- `SISTEMA\DADOS\operacao\TESTAR-RESTAURACAO.bat` — restaura o último backup numa
  cópia e confere. Backup não testado não é backup.
- `IA\INSTALAR-MCP.bat` — conector do Claude Desktop, em cada máquina que usar.

Roteiro completo, com as conferências de cada passo:
`SISTEMA\DADOS\docs\implantacao.md`.

Primeira vez nesta máquina (copiar `SISTEMA` e `IA` para cá e deixar o Claude
Desktop pronto para manter o sistema): `SISTEMA\DADOS\docs\levar-para-a-empresa.md`.
