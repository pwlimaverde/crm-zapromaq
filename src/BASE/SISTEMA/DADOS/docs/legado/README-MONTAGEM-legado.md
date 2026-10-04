# Montagem do frontend

> **Caminho curto: rode `20 - SCRIPTS\MONTAR-FRONTEND.bat`.**
> Ele cria o `.xlsm` inteiro — abas, Config, os oito módulos, o formulário com os
> 20 controles posicionados, o código colado e os botões da tela inicial — e salva
> na sua Área de Trabalho.
>
> Exige uma caixa marcada no Excel, uma única vez:
> **Arquivo › Opções › Central de Confiabilidade › Configurações da Central de
> Confiabilidade › Configurações de Macro › [x] Confiar no acesso ao modelo de
> objeto do projeto do VBA.** Sem ela o Excel bloqueia criação de código por
> script — é proteção contra macro que se reescreve sozinha, não defeito.
>
> O resto deste documento é o passo a passo manual, que continua valendo como
> conferência e como plano B se o script falhar em alguma estação.

---

Passo a passo para levantar o `CRM_Zapromaq.xlsm` a partir das fontes desta pasta.
Fonte é o `.bas`; o que existe só dentro do `.xlsm` se perde quando o arquivo corrompe.

## 1. Arquivo

1. Novo arquivo Excel, salvo como **`CRM_Zapromaq.xlsm`** (Pasta de Trabalho Habilitada para Macros), em pasta **local**, nunca na rede.
2. Abas: `Início`, `Clientes`, `Contatos`, `Oportunidades`, `Painel`, `BasePainel` (oculta), `Config` (oculta).

## 2. Aba Config

| Célula | Conteúdo |
|---|---|
| `A2` | `Caminho do banco` |
| `B2` | `\\servidor\COMERCIAL\1 - COMERCIAL ZAPROMAQ\01 - CRM\01 - CONTROLE\BASE\crm_zapromaq.accdb` |
| `A3` | `Ambiente` |
| `B3` | `PRODUCAO` — use `TESTE` enquanto estiver homologando |
| `A4` | `Última consulta` |
| `B4` | (fica em branco; o sistema preenche) |

O caminho é **UNC**. Não use `G:` — a letra existe na estação, não no script.

## 3. Módulos

`Alt+F11` → **Arquivo › Importar Arquivo**, um por um, nesta ordem:

`modConfig.bas` · `modDB.bas` · `modValidacao.bas` · `modCalc.bas` · `modSchema.bas` · `modCRM.bas` · `modGrade.bas` · `modPainel.bas` · `modMenu.bas`

Não é preciso marcar referência ao ADO: o código usa late binding.

## 4. Formulário

**Inserir › UserForm**. Em Propriedades: `(Name)` = **`frmCRM`**, `Width` = 700, `Height` = 520.

Controles fixos, com estes nomes exatos:

| Nome | Tipo | Left | Top | Width | Height | Caption |
|---|---|---|---|---|---|---|
| `lblTitulo` | Label | 8 | 6 | 480 | 16 | — |
| `txtBusca` | TextBox | 8 | 26 | 240 | 18 | — |
| `cmdBuscar` | CommandButton | 252 | 26 | 56 | 18 | Buscar |
| `lblFiltro1` | Label | 316 | 29 | 40 | 14 | Etapa: |
| `cboFiltro1` | ComboBox | 356 | 26 | 140 | 18 | — |
| `lblFiltro2` | Label | 502 | 29 | 34 | 14 | Resp.: |
| `cboFiltro2` | ComboBox | 536 | 26 | 100 | 18 | — |
| `chkAbertas` | CheckBox | 8 | 46 | 130 | 14 | somente em aberto |
| `lblCabecalho` | Label | 8 | 62 | 670 | 11 | — |
| `lstReg` | ListBox | 8 | 74 | 670 | 120 | — |
| `lblStatus` | Label | 8 | 198 | 480 | 14 | — |
| `fraCampos` | Frame | 8 | 216 | 676 | 240 | Dados |
| `lblModo` | Label | 8 | 462 | 200 | 14 | — |
| `cmdVinculo` | CommandButton | 206 | 460 | 68 | 20 | Procurar |
| `cmdNovo` | CommandButton | 278 | 460 | 56 | 20 | Novo |
| `cmdEditar` | CommandButton | 338 | 460 | 56 | 20 | Editar |
| `cmdSalvar` | CommandButton | 398 | 460 | 62 | 20 | Salvar |
| `cmdCancelar` | CommandButton | 464 | 460 | 66 | 20 | Cancelar |
| `cmdAtendimento` | CommandButton | 534 | 460 | 74 | 20 | Promover |
| `cmdFechar` | CommandButton | 612 | 460 | 66 | 20 | Fechar |

> O `cmdAtendimento` troca de rótulo sozinho conforme a tela e o registro:
> **Promover** num pré-cliente, **+ Contato** num cliente, **+ Atend.** num contato.

Em `lstReg`, marque `ColumnHeads` = False (o cabeçalho é o `lblCabecalho`).

Duplo clique no formulário e cole **todo** o conteúdo de `frmCRM.txt`.

> Os campos de cada ficha **não** são desenhados: o formulário os cria a partir de `modSchema`, agrupados por seção.

## 5. Botões da aba Início

| Botão | Macro |
|---|---|
| Clientes | `GradeClientes` |
| Contatos | `GradeContatos` |
| Oportunidades | `GradeOportunidades` |
| Fila da manhã | `FilaDaManha` |
| Atualizar indicadores | `AtualizarIndicadores` |
| Aplicar pendências | `AplicarPendencias` |
| Sobre | `SobreOSistema` |

## 5.1 Abas de grade — Clientes, Contatos e Oportunidades

Cada uma dessas três abas recebe **o mesmo código de folha**, o conteúdo de `FolhaGrade.txt`, colado no objeto da planilha (não num módulo). `modGrade` descobre a tabela pelo nome da aba, então não existe uma versão por aba para manter.

Sem esse código a grade até carrega, mas **deixa de ser somente leitura**: não há duplo clique para abrir a ficha nem bloqueio de digitação na célula.

Botões de cada aba, na ordem, todos na linha 3:

| Botão | Macro |
|---|---|
| Atualizar | `GradeAtualizar` |
| Ver tudo | `GradeVerTudo` |
| Abrir ficha | `GradeAbrirFicha` |
| Limpar filtros | `GradeLimparFiltros` |
| Nova ficha | `AbrirClientes` / `AbrirContatos` / `AbrirOportunidades` |
| Início | `GradeVoltar` |

Layout fixo, definido em `modGrade`: linha 1 guarda o modo (`MINIMO` / `COMPLETO`), linha 2 é o título, linha 3 é a faixa dos botões, linha 4 é o status, linha 6 é o cabeçalho da tabela e os dados começam na 7. A coluna A é só respiro.

**Três coisas que não se alteram sem perder a garantia:**

1. **Nada de célula mesclada na área de dados.** AutoFiltro e ordenação do Excel não funcionam sobre mesclagem — foi por isso que o layout do modelo visual não pôde ser copiado como estava. O visual vem de estilo de tabela e cor.
2. **A grade é somente leitura.** Digitar na célula não passa por validação nem pelo controle de versão. Quem edita é a ficha.
3. **Uma consulta por clique em Atualizar.** A grade carrega a tabela inteira e quem filtra dali em diante é o Excel, em memória.

As colunas de cada modo saem de `modSchema.ColunasPlanilha`. Campo acrescentado ali tem de existir na consulta correspondente de `modCRM`, fora os dois calculados: `codigo_visual` e `situacao`.

## 6. Antes de distribuir

1. `Ambiente` = `TESTE` e banco de teste: lançar, editar e encerrar um atendimento.
2. Conferir os campos calculados de uma amostra contra a planilha 3.6.
3. Abrir o mesmo atendimento em duas estações e salvar nas duas: a segunda tem de recusar, não sobrescrever.
4. Só então `Ambiente` = `PRODUCAO`, apontando para `crm_zapromaq.accdb`.
5. Uma cópia local por estação. **Nunca a mesma cópia na rede** — era esse o problema que a migração veio resolver.

## 7. Ao alterar qualquer módulo

Exporte de volta para esta pasta na mesma sessão (**Arquivo › Exportar Arquivo**) e suba a `VERSAO_FRONT` em `modConfig`. A versão anterior vai para `99 - HISTORICO` com sufixo de data.
