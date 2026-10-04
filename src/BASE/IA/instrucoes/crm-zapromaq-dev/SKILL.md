---
name: crm-zapromaq-dev
description: Mantém o sistema CRM Zapromaq (front Excel VBA + banco Access) pelas ferramentas do conector crm-zapromaq-dev. Use quando o usuário pedir para alterar uma tela, um texto, uma cor, uma regra do front, acrescentar campo no cadastro, criar ou aplicar migração de banco, conferir o projeto, montar uma versão de teste ou publicar uma versão nova do CRM.
---

# Manutenção do CRM Zapromaq

O sistema é uma pasta na rede (`BASE`) com o banco `crm_zapromaq.accdb`, o front
`CRM_Zapromaq.xlsm` (montado do zero a cada publicação) e as fontes em
`SISTEMA\DADOS`. **Toda manutenção passa pelas ferramentas do conector
`crm-zapromaq-dev`**: não edite arquivos por Python, não abra o `.accdb` por
outro caminho e não altere o `.xlsm` à mão — ele é gerado, não é fonte.

O conector de dados (`crm-zapromaq`) é outro: serve para consultar e alterar
*registros*. Este aqui altera o *sistema*.

## Sempre comece por `estado_sistema`

Ele diz a versão que a próxima publicação vai gerar, a versão publicada na rede,
a versão do esquema do banco, as migrações pendentes, se o banco está em uso
(alguém com o CRM aberto) e se a máquina tem Excel e Python. Decida a partir daí
e conte ao usuário o que encontrou.

## Alterar o front (telas, textos, regras)

1. `listar_fontes` / `ler_fonte` — guarde a `versao` devolvida.
2. `gravar_fonte` com essa `versao` e um `motivo` curto. Se vier **Conflito**,
   leia de novo e refaça a alteração; nunca repita às cegas.
3. `gerar_previa` (imagem da tela) quando mexer em layout; `verificar_projeto`
   sempre; `montar_teste` para compilar de verdade e rodar o autoteste.
4. `publicar` só quando o usuário mandar, com `confirmacao = "PUBLICAR"`.

O que fica de fora de propósito: `build\`, `IA\`, a raiz de `BASE` e os arquivos
gerados (`assets\gerado\`, `modulos\modTema.bas`) — estes mudam alterando
`layout\tema.json` / `layout\icones.json` e rodando `gerar_previa`.

Onde ficam as coisas:

| Quero mudar | Arquivo |
|---|---|
| Posição/tamanho de controle, botão, grade | `layout\formularios\<tela>.json` |
| Cores, fontes, variantes de botão | `layout\tema.json` |
| Ícones (glifos Segoe MDL2) | `layout\icones.json` |
| Comportamento da tela | `formularios\<tela>.txt` |
| Campos da ficha e colunas das abas | `modulos\modSchema.bas` |
| Regras de negócio e consultas | `modulos\modCRM.bas` |
| Campos calculados (situação, dias parado) | `modulos\modCalc.bas` |
| Normalização na gravação | `modulos\modValidacao.bas` |

Regras do VBA que o conector já cobre, mas que mudam como você escreve:
Windows-1252 (caractere fora da página é recusado — use `ChrW$(&H2192)`),
máximo de 25 continuações de linha por instrução, Excel **2019** (nada de
`XLOOKUP`, `LET`, `LAMBDA`, matrizes dinâmicas) e nada de OCX/DLL externo.

## Alterar o banco (campo novo, tabela nova, correção de dados)

Mudança de banco **nunca** é feita à mão: vira uma migração numerada, que roda
uma vez por banco e fica registrada. Use `listar_migracoes` para ver o que já
existe e `criar_migracao` para escrever a próxima.

**Migração de dados** (`tipo = "dados"`): corrigir valores, preencher campo,
acertar lista. Roda dentro de uma transação e pode ser aplicada com gente usando
o CRM.

**Migração de estrutura** (`tipo = "estrutura"`, exige `esquema_para`):
acrescentar campo ou tabela. Exige o banco exclusivo (ninguém com o CRM aberto)
e sobe a `versao_esquema`. Acrescentar um campo é um pacote, não um comando
solto — faça tudo antes de publicar:

1. `criar_migracao` com o `ALTER TABLE ... ADD COLUMN ...` e o `esquema_para`
   (ex.: 1.1 → 1.2).
2. `ler_fonte`/`gravar_fonte` em `modulos\modSchema.bas`: a linha do campo em
   `CamposDe` (rótulo, tipo, lista, seção, obrigatório, editável) e, se ele deve
   aparecer na aba, em `ColunasPlanilha`.
3. Se o campo entra em consulta, acerte o `SELECT` correspondente em
   `modulos\modCRM.bas`.
4. Acrescente o campo em `banco\esquema\criar-banco.ps1` (é ele que cria um banco
   novo do zero — sem isso, banco novo nasce diferente da produção).
5. `verificar_projeto` (confere `modSchema` contra o DDL), `montar_teste`, e só
   então `aplicar_migracoes` + `publicar`.

O conector recusa SQL destrutivo (`DROP`, `DELETE`/`UPDATE` sem `WHERE`) a menos
que você mostre ao usuário o que vai acontecer e ele confirme — aí repita com
`confirmacao = "CONFIRMO"`.

`aplicar_migracoes` faz backup do banco antes e aceita `simular = true` para só
mostrar o que faria. `publicar` também aplica sozinho o que estiver pendente.

## Publicar

`publicar` faz a sequência inteira: backup do banco e do front, migrações
pendentes, montagem do `.xlsm`, compilação, autoteste, conferência do pacote,
teste de uso e, só se tudo passar, substitui o `.xlsm` da raiz e sobe a versão.
Se qualquer etapa falhar, nada é publicado.

Depois de publicar, diga ao usuário para:

1. copiar o `CRM_Zapromaq.xlsm` publicado para a máquina dele e conferir;
2. avisar a equipe para abrir o `INICIAR-CRM` (ele vê a versão nova em
   `VERSAO-FRONT.txt` e copia o arquivo para cada estação).

Deu errado depois de publicado: as versões anteriores do front estão em
`SISTEMA\DADOS\execucao\historico-front\` e o banco em
`execucao\backup\antes-da-publicacao\<data-hora>\`. Use `ver_log` para ler o
final do último log (`build`, `visual`, `migracoes`, `backup`, `ambiente`).

## Coisas que exigem a máquina certa

`montar_teste` e `publicar` precisam de **Excel 2019 64 bits** instalado;
`gerar_previa` precisa do **Edge**; `verificar_projeto` é completo com Python e
cai numa conferência básica sem ele (nesse caso, confie no `montar_teste`, que
compila e roda o autoteste de verdade). `estado_sistema` diz o que a máquina
tem.

## O que este conector não faz

Alterar registros do CRM (é o conector `crm-zapromaq`), excluir migração já
aplicada, mexer em `build\`, na pasta `IA\` ou no `INICIAR-CRM.bat` das estações.
Para isso, fale com quem mantém o repositório de desenvolvimento.

Toda gravação feita aqui fica em `execucao\dev\alteracoes.log` (com cópia do
arquivo anterior) e precisa voltar ao repositório antes do próximo ciclo —
`SISTEMA\DADOS\docs\sincronizar-rede.md`.
