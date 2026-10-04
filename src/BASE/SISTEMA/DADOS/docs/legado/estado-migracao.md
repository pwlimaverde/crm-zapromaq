# Estado da migração — **CONCLUÍDA em 16/09/2026**

Atualizado em 16/09/2026, após a primeira tentativa de carga.

Retomada: refazer a verificação do último passo concluído, nunca confiar na memória de onde parou.

| # | Passo | Estado | Verificação |
|---|---|---|---|
| 1 | Teste complementar e tabela de funções | **concluído** | `30 - TESTES\resultado-complementar-20260916-090110.txt` — 17 de 18 construções passaram; `NZ` não existe |
| 2 | Dicionário de dados | **concluído** | `40 - DICIONARIO\dicionario-dados.md` — todas as colunas com destino ou descarte declarado |
| — | Extração e normalização | **concluído** | `40 - DICIONARIO\carga\` — 855 / 496 / 487 / 127; soma de valor conferida; 6 exceções registradas |
| 3 | `criar-banco.ps1` | **executado com sucesso** | 7 tabelas, 16 índices, banco de 324 KB |
| 4 | `carregar.ps1` | **executado com sucesso** | 855 / 496 / 487 / 127 em 9 s (1ª tentativa reprovou — ver "Correções da 1ª carga") |
| 5 | `aplicar-integridade.ps1` | **executado com sucesso** | 3 constraints aplicadas; exclusão de cliente com contato recusada pelo banco |
| 6 | `conferir-base.ps1` | **aprovado, zero divergência** | 29 comparações OK, incluindo soma de valor e distribuição por etapa |
| 7 | Campos calculados (`modCalc`) | **concluído e conferido** | `conferir-calculados.py`: 487 de 487 iguais à planilha nos sete campos |
| 8 | Módulos VBA e formulário | **montados, compilados e testados** | `MONTAR-FRONTEND.bat` gera o `.xlsm` completo; compilação sem erro; lançamento, edição e proteção da cópia modelo funcionando |
| 9 | `modPainel` e aba BasePainel | **aprovado em 16/09/2026** | 48 indicadores conferidos contra o banco, zero divergência |
| — | Leitura do banco pela sessão de IA | **conferida** | `mdbtools` leu o `.accdb` (ACE12) e devolveu os mesmos 855 / 496 / 487 / 127 e a mesma soma — contagem independente do ACE, como o passo 11 exige |
| 10 | `aplicar-fila.ps1` | **executado com sucesso** | 3 comandos aplicados, 2 recusados pelos motivos certos; contexto e log conferidos no banco pelo `mdbtools` |
| 11a | Teste de **edição simultânea** | **aprovado em 16/09/2026** | dois usuários no mesmo atendimento: o segundo foi recusado, sem sobrescrita |
| 11b | Backup e restauração | **agendado e restauração testada** | tarefas às 09:00, 12:30 e 16:00 (seg-sex) nesta estação — revisado em 23/09/2026; restauração conferida |
| 12 | Virada | **concluída em 16/09/2026** | produção com 852 / 500 / 489, soma de valor conferida, conferência sem divergência |

## Correções da 1ª carga (16/09/2026)

A primeira execução do `1-MIGRAR-TESTE.bat` parou no rollback de clientes: *"O campo é muito pequeno para aceitar a quantidade de dados"*. A transação desfez tudo — nada ficou pela metade. Duas causas, ambas de dado real que nenhum estudo previu:

| Achado | Correção |
|---|---|
| 344 cadastros de cliente e 12 de contato com mais de um telefone no campo, quase sempre com o nome de quem atende junto — e um com CNPJ no meio | Campo recebe **só dígitos** do primeiro número válido; nome e números adicionais vão para `observacoes` em linha marcada `[migração 16/09/2026]`. No front, telefone fora de 10–11 dígitos é recusado na gravação |
| 5 atendimentos com `PRÓXIMA AÇÃO` entre 272 e 389 caracteres | `prox_acao` passou de `TEXT(255)` para **MEMO**, e no schema do formulário virou campo de texto longo |

O `extrair.py` agora confere todo valor contra o limite da coluna e **aborta antes de gerar a carga** se algum estourar.

## Exceções em aberto

Seis contatos sem nome na origem, gravados como `GERAL`:
`CT-0194-0204`, `CT-0208-0221`, `CT-0466-0480`, `CT-0467-0481`, `CT-0481-0493`, `CT-0482-0494`.
Se algum tiver nome conhecido, corrigir na planilha e reextrair antes da carga de produção.

## Teste da fila — 16/09/2026

Pacote com 5 comandos, dois deles inválidos de propósito. Resultado conferido no banco por leitura independente:

- `ctx_cliente` gravou resumo, família, ponteiro do `.md`, data e autor no cliente 11
- `ctx_atendimento` gravou no AT-0011-0148
- `cad_cliente` com valor já igual não gravou nada, como deve
- `cad_cliente` no campo `versao` foi recusado: não está na lista de sete campos permitidos
- `ctx_cliente` no cliente 99999 foi recusado: não existe
- `log_alteracoes` com 2 registros, ação `CONTEXTO`, origem `FILA`, autor preservado
- As 855 empresas e os 487 atendimentos intactos — a fila não encostou em mais nada
- Arquivo movido para `RECUSADOS` com o `.motivo.txt` dizendo que 3 comandos já foram gravados

## Correção da proteção da cópia modelo (16/09/2026)

A primeira versão marcava a aba Início com um aviso e trocava a ação dos botões quando o arquivo era aberto da rede. Bastou salvar uma vez para o aviso virar permanente **no arquivo que todas as estações copiam** — os botões passaram a só repetir o aviso.

Proteção que modifica aquilo que protege não é proteção. Refeita: o bloqueio agora é só em memória (`modConfig.gModoModelo`), nada é escrito no arquivo, as macros se recusam a rodar e `Workbook_BeforeSave` cancela o salvamento explicando o motivo. Abrir da rede e cancelar a cópia virou o modo previsto para compilar e ler código.

## Depois da virada — pendências

1. ~~`config.ambiente` gravado como `TESTE`~~ — **corrigido em 16/09/2026**; o `criar-banco.ps1` passou a aceitar `-Ambiente`.
2. Agendar o backup também numa segunda estação, ou no servidor, enquanto depender de máquina ligada.
3. Mover a planilha 3.6 para uma pasta de consulta, fora do uso diário.

## Pendência resolvida antes da virada

`AT-0217-0220` (Contato Inicial) está **sem responsável** na base migrada. Não trava nada, mas some de qualquer visão filtrada por vendedor. Atribuir antes da carga de produção, na planilha 3.6, e reextrair.

## Congelamento

A partir daqui, nenhuma alteração de estrutura na `CRM-Comercial-Zapromaq-3.6.xlsm`.
Coluna nova na origem invalida o dicionário e obriga a refazer do passo 2.

## 17/09/2026 - grade: id oculta, travamento, ordenacao, pesquisa, icones

Decisoes do comercial nesta data:
- Cliente desativado: oportunidades em aberto FICAM COMO ESTAO (o aviso informa quantas sao).
- SUSPEITO (sem CNPJ, nome parecido) so MARCA a linha; nao e barreira dura.
- Indice unico de CNPJ aplicado agora na producao.

Conferencia da base de producao (leitura mdbtools, sem abrir o banco):
- 853 registros (488 Cliente, 365 Pre-cliente) - 1 a mais que na carga, cadastro novo na estacao.
- 392 com CNPJ, 461 com CNPJ NULO (nao vazio), 0 duplicado, 0 com formato invalido.
- Unico par de nome repetido: dois registros do mesmo grupo, com CNPJs de mesma raiz -
  matriz e filial. Nao e duplicata, e derruba a ideia de usar a raiz do CNPJ como chave.

Paleta da marca, extraida por contagem de pixel do logo institucional:
azul #172A67 RGB(23,42,103) - azul secundario #153F71 RGB(21,63,113) - verde #58B030 RGB(88,176,48).
Nao existe manual de marca na pasta; tipografia segue como pendencia.

Fontes alteradas (conferir-vba.py: sem problemas):
- modGrade.bas reescrito - colunas ocultas _id e _b, linha 4 e a pesquisa, LIN_CAB passou de 6 para 7,
  protecao da aba, ordenacao por script no clique do cabecalho, duas colunas de icone.
- modAcoes.bas novo - desativacao de cliente em cascata nos contatos e exclusao de pre-cliente.
- FolhaGrade.txt - SelectionChange e BeforeRightClick.
- montar-frontend.ps1 - importa modAcoes; mensagem inicial foi de B4 para B5.
- aplicar-indice-cnpj.ps1 e APLICAR-INDICE-CNPJ.bat novos.

Pendente de execucao na estacao (nada disso rodou):
1. APLICAR-INDICE-CNPJ.bat com o CRM fechado em todas as estacoes.
2. MONTAR-FRONTEND.bat com o .xlsm fechado.
3. Analisador de sintaxe do PowerShell nao rodou em montar-frontend.ps1 nem em
   aplicar-indice-cnpj.ps1 - so conferencia de balanceamento. Nao existe PowerShell no
   ambiente onde o script foi escrito.

Nao construido ainda: cadastro e exclusao em lote de pre-clientes, repaginado dos botoes.

### 17/09/2026 - correcao: eventos das abas saem do modulo de aba

Sintoma: montagem falhava nas tres abas com "Subscrito fora do intervalo" e,
com o diagnostico ampliado, "componente nao localizado (CodeName lido: '')" -
o projeto tinha so EstaPastaDeTrabalho e Planilha1.

Causa, conferida no XML do .xlsm salvo (leitura do zip, sem abrir o Excel):
das sete abas, SO A PRIMEIRA tem o atributo codeName. Aba criada por
automacao com Worksheets.Add sai sem CodeName, e sem CodeName nao existe
VBComponent de aba - nao ha onde colar codigo. Salvar nao cria o atributo.

Correcao: os quatro eventos passaram para ThisWorkbook.txt como
Workbook_SheetBeforeDoubleClick, Workbook_SheetSelectionChange,
Workbook_SheetBeforeRightClick e Workbook_SheetChange, todos guardados por
modGrade.TabelaDaAba(Sh). O objeto da pasta de trabalho e alcancavel por
$wb.CodeName, que funciona. Um lugar so, sem copia por aba.
FolhaGrade.txt esta em 99 - HISTORICO, descontinuado.

Segundo defeito: $sh.TextFrame.Characters(1, 255) estoura em botao recem
criado (tem menos de 255 caracteres). Era o caminho de reserva do texto do
botao, usado quando TextFrame2 nao responde - por isso caiu so em dois
botoes da aba Oportunidades. Agora usa Characters() sem intervalo.

Terceiro defeito, mesma montagem: nome de forma no Excel cabe em 31 caracteres.
btnOportunidadesGradeLimparFiltros e btnOportunidadesAbrirOportunidades tem 34, e so
por isso caiam - nas outras duas abas o nome cabia. O OnAction era aplicado antes do
nome, entao os botoes funcionavam e ficavam sem nome: o log acusava falha de botao que
respondia. Agora o nome e btn + tres primeiras letras da aba + indice (btnOpo4), 7
caracteres. Nenhuma fonte VBA referencia botao por nome - conferido por busca.

### 17/09/2026 - estilizacao, decidida para dentro do script

Decisao do comercial: a estilizacao mora no SCRIPT, nao numa planilha estilizada a mao.
Motivo tecnico: modGrade roda ws.Cells.Clear a cada Atualizar, entao formatacao manual na
area de dados dura um clique; e planilha estilizada a mao como fonte obrigaria refazer o
visual a cada mudanca de VBA. Escopo da formatacao condicional: sobria, destaque nas acoes,
Perdido e Descartado esmaecidos.

Corrigido:
- modGrade: as cinco constantes de cor estavam TODAS erradas (o valor do Excel e
  R + G*256 + B*65536, nao hexadecimal lido na ordem). O cinza de fundo saia quase branco-rosa.
- modGrade: formatacao condicional nao pintava nada - FormatConditions.Add(xlTextString, , texto)
  exige operador e o erro ficava engolido pelo On Error Resume Next. Agora e xlExpression com
  formula. Sete situacoes: atrasada, retomar hoje e nesta semana com fundo; em dia, encerrado,
  sem etapa e sem proxima acao so por cor de texto. Linha de Perdido/Descartado esmaecida em cinza.
- modGrade: cabecalho das colunas de icone mostra o proprio glifo.
- modSchema: larguras da 3.6 - Empresa 50 (era 44), Cidade 22 (era 18), e as das outras abas.
- montar-frontend: logo institucional na aba Inicio; botoes deixaram de ser controle de
  formulario cinza e passaram a forma arredondada com cor da marca e Segoe UI branco; faixa
  de botoes de 76 pt com passo 80 (seis botoes cabem, o ultimo nao corta); botao a 36 pt de
  altura para nao invadir a linha da pesquisa; hierarquia de cor na aba Inicio.
- frmCRM: passo de linha de 21 para 24 pt, memo de 42 para 54, Segoe UI 9 em tudo, borda
  chapada em vez de relevo 3D, rotulo em cinza e secao no azul da marca, acao principal em azul.

Limite que permanece: UserForm do VBA nao tem canto arredondado, sombra nem controle
customizado. Sai o aspecto de 1998; nao vira app web.

### 17/09/2026 - botoes esticados, logo, ordenacao e ordem das colunas

- Botoes esticando e cortando: forma no Excel nasce com Placement = mover e dimensionar
  com a celula, e modGrade muda a largura das colunas a cada Atualizar. Agora
  Placement = 3 (xlFreeFloating) em todo botao e no logo.
- Logo nao aparecia: o caminho por contagem de '..' errou um nivel (apontava para
  01 - PADROES\02 - MARKETING, que nao existe). Agora e busca por nome do arquivo a
  partir da raiz da pasta, com profundidade 4.
- Ordenacao da grade: clientes e contatos vinham por nome de empresa. Agora a consulta
  ja devolve ORDENADO POR CODIGO; pre-cliente nao tem codigo e vai depois dos clientes,
  por nome. Oportunidades por codigo do atendimento, decrescente.
- Ordem das colunas, seguindo a logica da 3.6 (o que se filtra vem primeiro):
  Clientes ....... Codigo, Estagio, Resp., Empresa, Cidade, UF
  Oportunidades .. Atendimento, Situacao, Etapa, Resp., Empresa, Contato, Valor, Prox. acao
  Contatos ....... inalterado (Codigo, Empresa, Contato, Cargo, Telefone)
  Mudar a ordem = mover uma linha em modSchema.ColunasPlanilha.

### 17/09/2026 - acoes saem da grade e vao para a ficha

- Colunas de icone (editar / desativar) REMOVIDAS da grade. Duplo clique na linha abre a
  ficha, que e onde se ve o registro inteiro antes de decidir. O menu do botao direito
  perdeu a acao de desativar pelo mesmo motivo; ficou com abrir ficha, atualizar e limpar.
- Ficha ganhou cmdStatus: um botao com tres caras, conforme modAcoes.RotuloStatus -
  Excluir (pre-cliente), Desativar (registro ativo) ou Reativar (registro inativo).
  Desabilitado em oportunidades: atendimento se encerra pela Etapa, nao por ativo.
  lblModo mostra [INATIVO] ao lado do codigo.
- modAcoes: novas AcaoStatus, RotuloStatus, ReativarCliente e AlternarSimples.
  Reativar cliente traz junto os contatos que a cascata desativou - senao o cliente volta
  sem ninguem para ligar, e a falta so aparece na hora de lancar atendimento.
- modCRM: as consultas de clientes e contatos passaram a trazer o campo ativo, que a ficha
  precisa para saber qual cara o botao tem.
- Caixa de pesquisa padronizada: ocupava uma celula, e a largura da celula muda por aba
  (Empresa 50 em Clientes, Situacao 18 em Oportunidades) - a mesma caixa aparecia de tres
  tamanhos. Agora mescla ate cerca de 210 pt em qualquer aba. Mesclagem so na faixa de
  cabecalho; na area de dados continua proibida.
- Logo pequeno (116x40) a direita da faixa de botoes das tres abas de grade.
- Tons dos botoes: escala derivada do azul da marca, do escuro ao claro, Nova ficha no
  verde (unica que cria registro) e Inicio/Sobre em cinza.

### 17/09/2026 - grade legivel

- Caixa de pesquisa com span FIXO: rotulo em B, caixa de C a E nas tres abas. A versao
  anterior mesclava ate somar 210 pt, e como a largura das colunas muda por aba, a caixa
  mudava de tamanho - era o mesmo defeito de antes por outro caminho.
- CNPJ na grade agora sai com mascara (tipo J = modValidacao.FormatarCNPJ). O banco continua
  guardando so digitos.
- Observacoes (tres abas), Categoria e Proxima acao SAIRAM da grade e do Ver tudo - texto de
  240 caracteres empurra as outras colunas. Continuam na ficha, inteiros, e continuam
  alcancaveis pela pesquisa: entram na coluna oculta de busca mesmo sem coluna visivel,
  junto com maquina, CNPJ e e-mail.
- Area de dados com ShrinkToFit: nome comprido encolhe na celula em vez de sumir atras da
  coluna seguinte. Alinhamento vertical ao centro.

### 17/09/2026 - cadastro em lote de pre-clientes

Fontes novas: modLote.bas e frmLote.txt. Ponto de entrada modMenu.CadastroEmLote, botao
"Cadastro em lote" na aba Inicio (verde, como Nova ficha - as duas criam registro).

Entrada: bloco copiado da planilha de qualificacao, separado por TABULACAO, na ordem
confirmada: aderencia, porte, qualificacao, empresa, cidade, estado, segmento, telefone,
e-mail, CNPJ, observacoes. Telefone e CNPJ entram so com digitos.

Quatro estados por linha:
  OK ......... cadastra
  SUSPEITO ... cadastra e diz por que desconfiou; duplo clique alterna para IGNORAR
               (CNPJ com digito verificador que nao fecha, nome parecido com cadastrado,
                nome repetido no proprio lote, segmento fora da lista)
  DUPLICADO .. nao cadastra (CNPJ ja na base ou repetido no lote)
  ERRO ....... nao cadastra (sem empresa, UF invalida, aderencia/porte fora de 1 a 3,
               CNPJ com numero de digitos errado, linha com menos de 11 colunas)
O estado pior manda: conferencia posterior nao rebaixa ERRO para SUSPEITO.

A base inteira e lida em UMA consulta para montar os indices de CNPJ e de nome normalizado.
Conferir por linha seria uma ida a rede por empresa - a 26 ms, 200 linhas dariam 5 s so nisso.

Campo vazio grava NULO, nunca cadeia vazia: o indice unico de CNPJ aceita varios nulos e
recusaria a segunda vazia.

conferir-vba.py ganhou duas checagens:
  - Private nao conta como nome repetido (cmdFechar_Click em dois formularios nao e conflito)
  - declaracao de modulo depois do primeiro procedimento, que foi o defeito de compilacao
    do frmCRM nesta data e passava como OK.

PENDENCIA: nao existe lista "Qualificacao" na tabela listas - a ficha mostra o combo vazio.
Os valores em uso na base sao Prioridade alta, Fila padrao, Reserva, Dado inconsistente,
Reprovado e Insuficiente para qualificar (328 registros sem qualificacao).

### 17/09/2026 - edicao em sequencia e sincronia grade/ficha

DEFEITO GRAVE corrigido: editar pela ficha e fechar deixava a LINHA DA GRADE com o valor
antigo ate alguem clicar em Atualizar - a tela mentia. Agora modConfig.gGravou e ligado pela
ficha a cada gravacao (salvar, criar, ativar/desativar/excluir) e lido por modGrade quando a
ficha fecha; se gravou, a grade e remontada. Variavel de modulo e nao propriedade do
formulario porque Unload destroi o objeto e levaria a resposta junto.

Edicao em lote, as duas rotas, conforme decisao do comercial:
1. NAVEGACAO PELO CONJUNTO DA ABA - modGrade.IDsVisiveis devolve os ids das linhas visiveis
   na ordem da tela (o que sobrou do filtro, como foi ordenado) e passa para frmCRM.Abrir.
   A ficha ganhou < e > e o contador "3 de 27". Ao receber conjunto, a ficha zera os proprios
   filtros - senao um id da aba poderia nao existir na lista dela.
   Isso cobre o pedido original: filtro combinado de QUALQUER numero de colunas, porque quem
   filtra e o AutoFiltro nativo da aba.
2. FILTROS E ORDENACAO NA FICHA - terceiro combo (Origem em oportunidades, Qualificacao em
   clientes) e combo "Ordenar por", com a ordenacao feita no BANCO:
   oportunidades: atendimento / proxima acao / etapa+proxima acao / valor / entrada / empresa
   clientes: codigo / empresa / UF e cidade / aderencia
   contatos: codigo / empresa / contato
   Data nula vai por ultimo: em Access o ORDER BY joga Nulo na frente, e "sem proxima acao"
   no topo empurraria para baixo o que tem prazo.
3. Botao "Salvar e prox." grava e ja abre o proximo do conjunto.

Formulario passou a 780 x 588 para caber tudo sem espremer a lista.

PENDENCIA MANTIDA: nao existe lista "Qualificacao" na tabela listas - o combo fica so com o
item vazio. Valores em uso: Prioridade alta, Fila padrao, Reserva, Dado inconsistente,
Reprovado, Insuficiente para qualificar.

### 17/09/2026 - filtro sobrevive ao Atualizar; linha certa ao reabrir

Defeitos relatados no teste: ao fechar a ficha a grade atualizava mas PERDIA os filtros, e ao
refazer o filtro e clicar na primeira linha a ficha abria OUTRO registro.

Causa: eu mandava remontar a aba inteira depois de gravar. Montar destroi a tabela e recria -
leva junto filtro, ordenacao e a referencia de linha.

Correcoes:
- Depois de gravar, so a LINHA do registro e reescrita (AtualizarUmaLinha). Filtro, ordenacao,
  congelamento e a posicao da tela ficam como estavam. Remontagem completa so quando a linha
  nao pode ser resolvida - registro excluido, por exemplo.
- modConfig.gGravados guarda a LISTA de ids gravados, nao um Sim/Nao: com "Salvar e prox." o
  vendedor edita varios registros sem fechar a ficha, e atualizar so o primeiro deixaria o
  resto da tela desatualizado.
- Montar passou a GUARDAR e REAPLICAR os filtros de coluna, o texto da pesquisa e a ordenacao.
  Criterio simples e filtro de lista voltam; filtro de cor e "os 10 primeiros" nao - a
  reaplicacao e silenciosa de proposito.
- IDdaLinha e ValorDaLinha passaram a ler a celula pelo ENDERECO FISICO (ws.Cells(alvo.Row,...))
  em vez de posicao relativa dentro da tabela. Abrir a ficha do registro errado e o pior
  defeito que esta grade pode ter, e a conta relativa dependia de a tabela nunca se deslocar.

### 17/09/2026 - versao 1.2 do front

1. CONTROLE DE VERSAO. modConfig.VERSAO_FRONT passou a 1.2 e e a unica fonte do numero.
   O montar-frontend le esse valor da propria fonte e grava em config.versao_front no banco;
   o Workbook_Open compara e avisa quando a copia da estacao ficou para tras. Nao bloqueia:
   versao de front diferente nao corrompe dado - quem corrompe e esquema incompativel, e isso
   o EsquemaCompativel ja barra. A aba Inicio mostra versao e data de publicacao.
   O NOME DO ARQUIVO CONTINUA FIXO de proposito: MONTAR-FRONTEND.bat, conferir-painel.ps1,
   3-VIRAR-PARA-PRODUCAO.bat e a rotina de copia local apontam para CRM_Zapromaq.xlsm, e
   atalho na estacao tambem. Nome com data quebraria tudo isso e ainda dependeria de alguem
   reparar na data; a comparacao com o banco avisa sozinha.
2. "Salvar e prox." ja entra em modo de edicao no proximo registro.
3. Formulario de 780x588 para 830x500, uma faixa de botoes so. A segunda faixa ficava
   escondida em tela de notebook e UserForm nao tem rolagem.
4. Rotulo do memo caia em cima do campo anterior quando o memo vinha logo depois de um campo
   da coluna da esquerda - era o risco vertical atras do campo Observacoes. Agora o memo fecha
   a linha antes de se desenhar.

NAO FEITO, com motivo: rolagem do mouse dentro do UserForm. VBA nao tem evento de roda; o
unico caminho e SetWindowsHookEx (subclassing), que e a causa mais comum de fechamento do
Excel sem aviso quando o hook nao e liberado - inclusive ao entrar em modo de depuracao.
Nao vale o risco num arquivo que a equipe usa o dia inteiro. O frame ja tem barra de rolagem
lateral, e com a janela mais alta sobra menos o que rolar.

### 17/09/2026 - travada ao fechar a ficha

Sintoma: editar varios atendimentos e fechar a ficha deixava a planilha segundos travada.

Causa: AtualizarGravados chamava a atualizacao UMA VEZ POR REGISTRO, e cada chamada refazia a
consulta da tabela inteira. Dez atendimentos editados eram dez consultas de 490 linhas.

Nao foi corrigido mudando QUANDO dispara. Disparar a cada Salvar daria as mesmas dez
consultas, distribuidas, travando a ficha no meio da edicao - espalha a espera, nao reduz.

Correcao: UMA consulta para todos os ids gravados; ids distintos (com "Salvar e prox." o mesmo
registro pode ser salvo duas vezes); coluna de id lida em bloco em vez de celula por celula;
cada linha escrita de uma vez como matriz; ScreenUpdating desligado no meio.
De N consultas + N*nc escritas de celula para 1 consulta + N escritas de linha.

Armadilha evitada no caminho: a variavel do For Each nao pode ser reaproveitada para o Split
dentro do laco - reatribuir o enumerador corrompe a iteracao do Dictionary.

PONTO AINDA ABERTO, se a lentidao voltar: frmCRM chama Buscar() depois de CADA Salvar, e isso
e uma consulta completa por salvamento com a ficha aberta. Atualizar so a linha da lista do
formulario exigiria recalcular os campos derivados (situacao) a partir do que esta na tela.
Nao foi medido - nao mexer sem medir.

### 17/09/2026 - formulario dividido na vertical

A ficha de edicao ficava EMBAIXO da lista: para ver campo o usuario rolava, e perdia a lista
de vista. Agora e divisao vertical:
  esquerda (8 a 506) - busca, tres filtros, ordenacao, lista de 298 pt e o status
  direita  (514 a 984) - so o frame de campos, 398 pt de altura
  rodape (linha unica) - registro, ativar/desativar, navegacao e as acoes

Janela de 1000 x 505 pt. Nao passei disso de proposito: ponto do UserForm vale 1,333 pixel a
96 dpi, entao 1000 pt sao 1333 px e ainda cabem em 1366x768, que e o piso de tela que existe
em estacao de fabrica. UserForm nao tem rolagem para compensar janela maior que a tela.

Campos redimensionados para o frame de 470: duas colunas de 130 pt com rotulo de 96
(eram 186 com a segunda coluna em 344, o que media 650 e nao cabe mais); memo com 356.
As medidas viraram Const no inicio da rotina - ajustar layout agora e mudar numero, nao caçar
literal espalhado.

modSchema.ColunasDe (lista do formulario) recalculado para caber nos 498 pt:
clientes 479 - contatos 498 - oportunidades 498. Contato saiu da lista de oportunidades
(aparece na ficha ao lado) e Situacao entrou: e por ela que se escolhe o que atender primeiro.

### 17/09/2026 - acabamento do formulario dividido

1. Barra de rolagem do frame ficava POR CIMA da segunda coluna de campos. Ela ocupa cerca de
   16 pt, entao a largura util do frame de 470 e 450, nao 470. Medidas refeitas:
   rotulo 92, campo 118, coluna 1 em 232, memo 346. Ultimo campo termina em 446.
2. lblTitulo removido do frmCRM: a barra de titulo da janela ja diz qual tela e, e o rotulo
   repetia a mesma frase duas linhas abaixo. O espaco virou altura de lista (318 pt) e de
   frame (416 pt). O frmLote mantem o proprio titulo - la a janela e uma so.
3. Cabecalho da lista desalinhava porque a conta usava 5 pt por caractere. Em Consolas a
   largura do caractere vale 0,6 do corpo da fonte: 4,8 pt a 8 pt. Com 5, o erro de 0,2 pt por
   caractere acumulava e a ultima coluna saia deslocada. Titulo cortado agora leva ponto.
4. Larguras da lista: a ListBox tem 498 pt, mas a rolagem vertical come 16 - a soma cabe em
   478. Era isso que cortava a coluna Situacao. clientes 466, contatos 478, oportunidades 478.

### 17/09/2026 - alinhamento da lista e arranjo dos campos

1. CABECALHO DA LISTA virou uma ListBox de uma linha, desabilitada, com as MESMAS larguras de
   coluna da lista de dados - alinhamento exato por construcao. Era um Label de fonte fixa e
   nunca casou: a ListBox usa fonte proporcional, soma folga interna por coluna, e qualquer
   erro de fracao de ponto por caractere acumulava ao longo da linha. Duas tentativas de
   acertar a conta (5 pt e 4,8 pt por caractere) falharam pelo mesmo motivo - o problema nao
   era a constante, era a abordagem.
2. LARGURAS pela largura do CONTEUDO, nao do titulo: "AT-0489-0491" tem 12 caracteres e ocupa
   cerca de 62 pt em Segoe UI 9, e eu dava 86 - o vao entre Atendimento e Empresa ficava maior
   que os outros, e era isso que parecia desalinhado. Prox. acao saiu da lista de
   oportunidades: Situacao ja diz se esta atrasada, nesta semana ou em dia, que e a decisao
   tomada olhando a lista; a data exata se le na ficha.
   Somas: clientes 476, contatos 478, oportunidades 478.
3. FICHA - titulo de secao agora FECHA a linha pendente antes de se desenhar (era o
   "3. Prospeccao" colado no campo E-mail), e campo cujo conteudo nao cabe em 118 pt ocupa a
   LINHA INTEIRA: empresa, email, maquina, orcamento, ctx_resumo, ctx_familia. Coluna estreita
   com texto cortado nao economiza espaco - a linha ao lado fica vazia do mesmo jeito.
4. RODAPE com o grupo de botoes centrado (140 a 854) e altura da janela de 505 para 480, para
   a faixa de botoes nao ficar solta longe da borda.

### 18/09/2026 - versao 1.3: navegacao entre fichas

1. BOTOES CORTADOS. Height do UserForm nao e area util: barra de titulo e bordas comem cerca
   de 35 pt. Com 480 a faixa em 436 saia pela metade. Passou a 496.
2. ROTULO CORTADO. Campo que ocupa a linha inteira agora tem rotulo de 120 pt, nao 92 -
   "Contexto do atendimento" nao cabia. Sobra largura nessa linha, entao o rotulo pode crescer.
3. DUPLO CLIQUE ABRINDO OUTRO ITEM - causa encontrada. SelecionarID nao fazia nada quando o id
   nao estava na lista, e a ficha seguia mostrando o PRIMEIRO registro como se fosse o clicado.
   Acontecia com atendimento encerrado enquanto "somente em aberto" estava marcado. Agora
   SelecionarID limpa os filtros, refaz a consulta e procura de novo; se ainda nao achar, avisa.
   Registro errado na tela calado e pior do que aviso.
4. NAVEGACAO ENTRE FICHAS - dois botoes de rotulo variavel no rodape:
   clientes ....... Contatos | Atend.
   contatos ....... Cliente  | Atend.
   oportunidades .. Cliente  | Contato
   Abrem OUTRA ficha ja com o conjunto relacionado carregado, entao Anterior e Proximo
   percorrem os contatos daquele cliente ou os atendimentos daquele contato, sem filtro.
   modCRM.IDsDeContatosDoCliente / IDsDeAtendimentosDoCliente / IDsDeAtendimentosDoContato -
   uma consulta por clique, so da coluna id.
5. VERSAO_FRONT 1.3.

### 18/09/2026 - "Era esperada uma matriz"

frmCRM.VinculoAtual tinha um parametro chamado "campo" e chamava a funcao Campo() do proprio
modulo. O VBA NAO diferencia maiusculas em identificador: dentro da rotina o nome local ganha,
e Campo(...) virou indexacao de texto. Parametro renomeado para nomeCampo.

E a MESMA armadilha ja registrada para o PowerShell ($MEMO e $memo sao a mesma variavel) -
mudou a linguagem, nao a causa.

conferir-vba.py ganhou a checagem: parametro ou Dim com o nome de uma rotina do modulo, usado
como chamada dentro dela. Conferida contra o defeito real: reintroduzindo a colisao numa copia
das fontes, o script acusa. Terceira checagem que o verificador ganhou depois de deixar passar
um erro de compilacao - as outras duas sao declaracao de modulo fora de lugar e Private
contado como nome repetido.

### 18/09/2026 - versao 1.4: atualizacao sem o usuario ir na rede

INICIAR-CRM.bat, em 01 - CRM\01 - CONTROLE\BASE, copiado UMA VEZ para a area de trabalho de
cada estacao. Nao precisa ficar na pasta da rede: localiza a rede sozinho, tentando a propria
pasta, depois a unidade mapeada, depois o UNC - primeira que existir vence, sem letra de
unidade fixa no codigo.

Copia so se a versao da rede for diferente (robocopy decide), e abre a copia local.

A copia de trabalho saiu de C:\CRM Zapromaq\ para Documentos\CRM Zapromaq\. Duas razoes:
estacao com politica restritiva recusa escrita na raiz do C:, e Documentos e a pasta que o
usuario acha sozinho. O caminho vem do REGISTRO (Shell Folders\Personal no .bat,
WScript.Shell.SpecialFolders("MyDocuments") no VBA) e nao de %USERPROFILE%\Documents: com
OneDrive ligado a pasta e redirecionada e o caminho montado a mao aponta para o lugar errado.

Falhas previstas: rede fora do ar abre a copia local com aviso; CRM aberto na estacao impede
a substituicao e o .bat diz para fechar o Excel; sem copia e sem rede, nao abre e explica.

O aviso de copia desatualizada agora manda usar o atalho, nao ir na pasta da rede.
LEIA-ME-DISTRIBUICAO.md reescrito. VERSAO_FRONT 1.4.

### 18/09/2026 - lancador como arquivo unico + icone

INICIAR-CRM.bat fica SO na pasta da rede. Nenhuma estacao recebe copia dele: o que vai para a
maquina e um ATALHO apontando para ele na rede. Corrigir o lancador passa a ser mexer num
arquivo so - com copia por estacao, qualquer correcao exigiria voltar em todas.

.bat nao tem icone proprio: o Windows usa o icone do interpretador de comandos. Quem tem
icone e o ATALHO. Entao o proprio lancador cria o atalho "CRM Zapromaq" em
Documentos\CRM Zapromaq na primeira execucao, via VBScript temporario, apontando para o .bat
da rede, com WindowStyle 7 (minimizado).

crm_zapromaq.ico gerado a partir do logo institucional, em 16, 32, 48, 64, 128 e 256 px, sem
dependencia externa - PNG decodificado e reescrito em Python puro, com zlib. O icone e COPIADO
para a maquina: atalho com icone em caminho de rede fica generico quando a rede nao responde.

pushd "%~dp0" no inicio do .bat: aberto por caminho de rede, o CMD avisa que nao suporta UNC e
cai em C:\Windows. pushd mapeia uma unidade temporaria e o aviso nao aparece.

RESSALVA NAO RESOLVIDA: o logo tem proporcao 1307x452, bem horizontal. Em 16 px a faixa fica
com 5 px de altura e o icone pequeno tende a virar um risco. Nao foi possivel conferir
visualmente daqui. Se ficar ruim na estacao, a troca e substituir o .ico na pasta BASE.
