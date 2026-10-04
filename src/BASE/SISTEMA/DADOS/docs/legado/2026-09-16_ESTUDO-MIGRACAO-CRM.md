# Migração do CRM Comercial Zapromaq — planilha compartilhada para Access + Excel

**Data:** 16/09/2026 · **Revisão:** 00 · **Natureza:** estudo consolidado de migração, base para a construção
**Substitui:** `RELATORIO-PROJETO.md` (estudo inicial) e `2026-09-16_RELATORIO-FRONTEND-EXCEL-ACCESS.md` (estudo de frontend), que passam a ser material de origem
**Base funcional:** `CRM-Comercial-Zapromaq-3.6.xlsm`, lido por script em 16/09/2026
**Evidência de viabilidade:** `01 - PADROES\00 - IA\00 - SCRIPTS\teste-accdb\resultado-teste-20260916-082018.txt`

---

## 1. Decisão

Migrar. O modo Pasta de Trabalho Compartilhada do Excel já falhou várias vezes durante a implantação, e o mecanismo não tem correção: ele trava o arquivo inteiro, faz merge silencioso e corrompe quando a rede oscila. A decisão está tomada; este documento define o que se constrói, em que ordem, e como se sabe que cada etapa terminou.

**Arquitetura:** banco Access (`.accdb`) na pasta de rede, frontend Excel (`.xlsm`) em cópia local por estação, gravação por ADO/OLE DB, e um executor PowerShell na estação para o que vem da sessão de IA.

---

## 2. O que foi medido, e não presumido

Tudo nesta seção é resultado de execução, não estimativa.

### 2.1 Ambiente da estação (estação do comercial, 16/09/2026)

| Verificação | Resultado |
|---|---|
| Provedor OLE DB | `Microsoft.ACE.OLEDB.16.0` respondeu em processo 64 bits |
| Arquitetura | Em 32 bits, ACE 16 e 12 retornam "Classe não registrada" — o Office é **64 bits** |
| Arquivo de bloqueio | `.laccdb` criado na pasta de rede — o compartilhamento permite o mecanismo de lock |
| Caminho UNC | Funciona; o banco não depende de letra de unidade mapeada |
| `Microsoft.Jet.OLEDB.4.0` | Presente, mas **proibido**: cria arquivo formato JET4 com extensão `.accdb` — erro silencioso |

### 2.2 Comportamento do banco

| Teste | Resultado |
|---|---|
| `INSERT` parametrizado com acento e apóstrofo | Gravado e lido idêntico |
| Memo de 8.880 caracteres | Gravado e lido **idêntico** — o `CONTEXTO-GERAL.md` cabe |
| `@@IDENTITY` na mesma conexão | Devolve o id correto |
| `UPDATE` com controle de versão | 1 linha afetada; repetição com versão antiga devolve **0** |
| Transação com dado + auditoria | Commit conjunto; rollback desfaz |
| Índice único | Recusa duplicata |
| Data inválida (`#02/31/2026#`) | Recusada |
| Índice único com vários NULL | **Aceita** — as 372 pré-clientes convivem num índice único de `codigo_cliente` |
| Chave estrangeira | Aplicada; exclusão de cliente com contato e inserção de contato órfão **recusadas pelo banco** |
| Ordenação com acento | Correta: ABETO, ACO NOBRE, ACOS LEVES, ACUCAR MODELO |
| Função `NZ()` | **Não existe** via OLE DB — ver 5.2 |

### 2.3 Desempenho na pasta de rede

| Operação | Medido | Leitura |
|---|---|---|
| 500 inserções em transação única | 976 ms — **1,95 ms/registro** | A carga inicial de 1.466 registros roda em ~3 s |
| 50 inserções avulsas | 141 ms — 2,82 ms/registro | Transação também é mais rápida, não só mais segura |
| Busca `LIKE` sobre 554 registros | 16 ms | Extrapolado para 5.000: ~150 ms |
| Busca por índice único | 0 ms | Indexar o que se busca resolve |
| **Ciclo abre / consulta / fecha** | **25,9 ms** | **Ver 2.4 — é o número que governa o desenho do front** |
| `GROUP BY` por etapa, 500 registros | 46 ms | Os ~10 blocos do Painel ficam bem abaixo de 500 ms |
| `GROUP BY` por responsável, com filtro | 4 ms | Agregação com `WHERE` é praticamente gratuita |
| Tamanho do arquivo, 554 registros | 276 KB | Com os dados reais e os contextos, poucos MB. O teto de 2 GB não é limite prático |

### 2.4 A consequência dos 26 ms

A regra "conexão curta: abre, executa, fecha" custa **26 ms por ciclo na rede**. Por ação do usuário é imperceptível. Dentro de um laço, é fatal: uma grade de 500 linhas que abra conexão por linha trava a tela por 13 segundos.

> **Regra de construção, não recomendação:** conexão curta é **por operação do usuário**, nunca por linha, por campo ou por combo. Uma consulta devolve o conjunto; não se fazem N consultas para montar uma lista.

Isso invalida o padrão do estudo inicial, onde `ResumoContato`, `ObterLista` e `EmpresaDoCliente` são chamados dentro de laços de montagem de tela.

### 2.5 Estado da base em 16/09/2026

| Aba | Registros | Observação |
|---|---|---|
| Clientes | **855 empresas** | 483 com `CÓD. CLIENTE`, **372 sem** |
| Contatos | 496 | Códigos `CT-CCCC-NNNN` |
| Oportunidades | 487 | Códigos `AT-CCCC-NNNN` |
| Visitas | **0** | Aba oculta, sem uso — fora do escopo |

**Integridade conferida por script:** nenhum contato aponta para cliente inexistente; nenhuma oportunidade aponta para contato inexistente; nenhum código `AT` ou `CT` duplicado; nenhuma oportunidade sem etapa. **A base está limpa** — o risco de registros órfãos previsto no estudo inicial não se materializou.

**`CÓD. CLIENTE`:** valores de 1 a 484, 483 distintos, contíguos exceto o 53. O comportamento é de **sequencial do próprio CRM**, não de código importado de outro sistema. O desenho abaixo passa a gerar esse número pelo banco.

**Campos praticamente sem uso:** `ORDEM` (4 preenchidos em 855) e `SITUAÇÃO CADASTRO` (2 em 855). Descontinuados — ver 11.3.

**Sujeira encontrada:** 1 registro com `ADERÊNCIA` preenchida só com espaços. Tratado na carga.

**Distribuição das etapas:** Contato Inicial 257 · Proposta em Stand By 137 · Negociação 43 · Perdido 10 · Sem Retorno 10 · Descartado 8 · Proposta em Análise 7 · Levantamento Técnico 6 · Pedido Fechado 5 · Retorno Agendado 3 · Elaboração da Proposta 1.

### 2.6 Leitura do banco pela sessão de IA

Confirmado por execução em 16/09/2026: o `mdbtools` rodando na VM Linux da sessão abriu o `.accdb` criado pelo ACE 16.0 (formato reportado: `ACE12`), listou as tabelas, exportou as 554 linhas em CSV, devolveu o memo de 8.880 caracteres inteiro e preservou o apóstrofo de `INDUSTRIA D'ANGELO LTDA`.

**A sessão lê o banco direto. Não há exportação por demanda, nem espelho obrigatório, nem espera por alguém abrir o Excel.**

Três condições, tratadas no plano:

1. O `mdbtools` é reimplementação independente, não o motor da Microsoft. Na homologação, a contagem que ele devolve tem de bater com a que o ACE reporta (etapa 7).
2. Ele **só lê**. Escrita vinda da sessão passa pela fila (seção 9).
3. A VM é recriada a cada sessão: os binários ficam versionados em `01 - PADROES\00 - IA\00 - SCRIPTS\lib\mdbtools` (136 KB), não dependendo de rede liberada no dia.

---

## 3. Arquitetura

```
\\servidor\COMERCIAL\1 - COMERCIAL ZAPROMAQ\
  01 - CRM\01 - CONTROLE\
      BASE\crm_zapromaq.accdb           <- único arquivo de dados
      BASE\99 - BACKUP\                  <- cópias horárias + diária com janela exclusiva
      BASE\FILA\                         <- comandos vindos da sessão de IA
      BASE\FILA\APLICADOS\               <- comandos já processados
      CRM_Zapromaq.xlsm                  <- modelo oficial do frontend (controlado)
  02 - CLIENTES\NNNN - RAZAO\            <- pasta do cliente: contexto e documentos
```

Na estação: cópia local de `CRM_Zapromaq.xlsm`, em pasta local, **nunca na rede**.

| Camada | Tecnologia | Instalação necessária |
|---|---|---|
| Banco | Access `.accdb`, motor ACE 16.0 | nenhuma — vem com o Office |
| Frontend | Excel `.xlsm`, VBA + ADO | nenhuma |
| Executor da fila | PowerShell 5.1 + `System.Data.OleDb` | **nenhuma — já vem no Windows** |
| Sessão de IA | Python + `mdbtools`, na VM Linux | nenhuma na estação |

Nenhuma linha de Python vai para as estações. O Python roda onde já roda hoje.

O executor usa `System.Data.OleDb`, o mesmo caminho OLE DB do VBA do front — mesma semântica de tipo, data e transação nos dois lados. Um caminho ODBC teria semântica distinta e criaria divergência para caçar depois.

---

## 4. Codificação de clientes — a estratégia

### 4.1 O problema

Hoje a ausência de `CÓD. CLIENTE` é o que separa a fila de qualificação (372 empresas) dos clientes prospectados (483). É um marcador por omissão: funciona, mas não é declarado, não é filtrável de forma explícita e não sobrevive a um `id` interno.

### 4.2 A estratégia

Três campos, com papéis que não se misturam:

| Campo | Tipo | Papel |
|---|---|---|
| `id` | AUTOINCREMENT | Identidade interna. Existe para **toda** empresa, inclusive as 372 da fila. Nunca muda, nunca aparece na tela. É a FK dos contatos |
| `codigo_cliente` | LONG, único, aceita nulo | O número de 4 dígitos que já existe. **Só quem é Cliente tem.** Forma `CT-` e `AT-`, e nomeia a pasta |
| `estagio` | TEXT(12) | `Pré-cliente` ou `Cliente`. É o que o vendedor vê e filtra |

### 4.3 A regra que amarra os três

> `estagio` **não é digitável**. É derivado na gravação: sem `codigo_cliente` → `Pré-cliente`; com → `Cliente`.

Sem isso, alguém marca "Cliente" numa empresa sem código, e o atendimento nasce sem base para o código e sem pasta. O campo existe para consulta e filtro na grade, não como decisão livre — mesmo princípio da coluna `PRIORIDADE`, que hoje é fórmula e não se digita.

### 4.4 A promoção

**Abrir o primeiro atendimento de um pré-cliente promove a empresa automaticamente**, na mesma transação: atribui o próximo `codigo_cliente` (`MAX + 1`), muda o `estagio` para `Cliente` e grava o log. O vendedor não digita código, e não existe estado intermediário.

Também há promoção manual, pelo botão *Promover a cliente* na ficha, para quem quer cadastrar a empresa antes de abrir negociação.

Não há caminho de volta automático: rebaixar um cliente a pré-cliente é operação de exceção, feita pela gestão, e só é permitida se a empresa não tiver contato nem atendimento.

### 4.5 Por que não dar código a todas as 855

Porque o código nomeia pasta, e pasta se cria sob demanda (POP-COM-001, R1). Dar código a uma empresa que talvez nunca seja trabalhada é gerar ruído e consumir numeração. A fila de qualificação é justamente o conjunto que ainda não mereceu isso.

### 4.6 Os códigos existentes não mudam

Este é o ponto que o estudo inicial errava ao propor derivar `CT-00000` do `id`.

`CT-CCCC-NNNN` e `AT-CCCC-NNNN` nomeiam **479 pastas de atendimento** em `02 - CLIENTES` e estão escritos dentro dos arquivos de contexto. Qualquer mudança de formato desliga o CRM da pasta.

No banco:

- `contatos.controle` e `oportunidades.controle` recebem, na carga, o `CONTROLE` que a planilha já tem. Para registros novos, `MAX(controle) + 1`.
- `contatos.codigo` e `oportunidades.codigo` guardam o código completo **como texto congelado no ato da criação** — não recalculado.

O congelamento importa: se um contato for transferido para outra empresa, o código não pode mudar, ou o nome da pasta do atendimento deixa de corresponder ao registro.

---

## 5. Modelagem do banco

```sql
CREATE TABLE clientes (
  id                  AUTOINCREMENT PRIMARY KEY,
  codigo_cliente      LONG,                    -- NULL enquanto Pré-cliente; único quando preenchido
  estagio             TEXT(12)  NOT NULL,      -- 'Pré-cliente' | 'Cliente' (derivado)
  empresa             TEXT(120) NOT NULL,
  cnpj                TEXT(20),
  cidade              TEXT(60),
  uf                  TEXT(2),
  segmento            TEXT(60),
  telefone            TEXT(30),          -- só dígitos, 10 ou 11
  email               TEXT(120),
  responsavel         TEXT(60),
  aderencia           INTEGER,                 -- 1 a 3
  porte               INTEGER,                 -- 1 a 3
  qualificacao        TEXT(40),
  observacoes         MEMO,
  ctx_resumo          TEXT(255),               -- contexto: síntese de uma linha
  ctx_familia         TEXT(60),                -- família aderente confirmada
  ctx_arquivo         TEXT(255),               -- caminho relativo do CONTEXTO-GERAL.md
  ctx_atualizado_em   DATETIME,
  ctx_atualizado_por  TEXT(50),
  ativo               YESNO,
  versao              LONG NOT NULL,
  criado_em           DATETIME,
  alterado_em         DATETIME,
  alterado_por        TEXT(50)
);

CREATE TABLE contatos (
  id                  AUTOINCREMENT PRIMARY KEY,
  codigo              TEXT(15) NOT NULL,       -- CT-CCCC-NNNN congelado
  controle            LONG NOT NULL,
  id_cliente          LONG NOT NULL,           -- FK -> clientes.id
  nome                TEXT(120) NOT NULL,
  cargo               TEXT(60),
  telefone            TEXT(30),          -- só dígitos, 10 ou 11
  email               TEXT(120),
  observacoes         MEMO,
  ativo               YESNO,
  versao              LONG NOT NULL,
  criado_em           DATETIME,
  alterado_em         DATETIME,
  alterado_por        TEXT(50)
);

CREATE TABLE oportunidades (
  id                      AUTOINCREMENT PRIMARY KEY,
  codigo                  TEXT(15) NOT NULL,   -- AT-CCCC-NNNN congelado
  controle                LONG NOT NULL,
  id_contato              LONG NOT NULL,       -- FK -> contatos.id
  id_cliente              LONG NOT NULL,       -- congelado na criação
  id_atendimento_anterior LONG,                -- retomada
  orcamento               TEXT(30),
  etapa                   TEXT(40) NOT NULL,
  responsavel             TEXT(60),
  maquina                 TEXT(255),
  familia                 TEXT(60),
  categoria               TEXT(60),
  tipo_venda              TEXT(40),
  valor                   CURRENCY,
  origem                  TEXT(40),
  prioridade              TEXT(20),
  dt_entrada              DATETIME,
  dt_proposta             DATETIME,
  ultima_interacao        DATETIME,
  tentativas              INTEGER,
  prox_acao               MEMO,
  dt_prox_acao            DATETIME,
  retomar_em              DATETIME,
  motivo_desfecho         TEXT(60),
  dt_desfecho             DATETIME,
  observacoes             MEMO,
  ctx_resumo              TEXT(255),           -- contexto do atendimento
  ctx_arquivo             TEXT(255),
  ctx_atualizado_em       DATETIME,
  ctx_atualizado_por      TEXT(50),
  versao                  LONG NOT NULL,
  criado_em               DATETIME,
  alterado_em             DATETIME,
  alterado_por            TEXT(50)
);

CREATE TABLE listas (
  id     AUTOINCREMENT PRIMARY KEY,
  tipo   TEXT(30)  NOT NULL,
  valor  TEXT(120) NOT NULL,
  ordem  INTEGER,
  ativo  YESNO
);

CREATE TABLE metas (
  id  AUTOINCREMENT PRIMARY KEY,
  ano INTEGER NOT NULL, mes INTEGER NOT NULL,
  meta_faturamento CURRENCY, meta_pedidos INTEGER,
  meta_propostas INTEGER, meta_contatos INTEGER,
  alterado_em DATETIME, alterado_por TEXT(50)
);

CREATE TABLE log_alteracoes (
  id AUTOINCREMENT PRIMARY KEY,
  tabela TEXT(30), id_registro LONG, campo TEXT(40),
  valor_antigo MEMO, valor_novo MEMO,
  acao TEXT(20), usuario TEXT(50), quando DATETIME, origem TEXT(20)
);
```

**Índices:** `clientes(codigo_cliente)` único · `clientes(empresa)` · `clientes(cnpj)` · `clientes(estagio)` · `contatos(id_cliente)` · `contatos(codigo)` único · `contatos(nome)` · `oportunidades(id_contato)` · `oportunidades(id_cliente)` · `oportunidades(codigo)` único · `oportunidades(etapa)` · `oportunidades(dt_prox_acao)` · `oportunidades(orcamento)` · `listas(tipo)` · `metas(ano,mes)` único.

**Integridade referencial** aplicada após a carga: não se apaga cliente com contato, nem contato com oportunidade.

**Campos calculados, que não existem no banco** e são computados na exibição: `SITUAÇÃO`, `QUADRO`, `DIAS PARADO`, `CICLO`, `MÊS PROPOSTA`, `MÊS ENTRADA`, `MÊS DESFECHO`. As mesmas regras da planilha 3.6, reescritas em VBA uma única vez.

**Campos que vêm por JOIN e não se repetem:** empresa, cidade, UF, segmento, contato, cargo, telefone, e-mail.

### 5.1 Padrão de gravação de texto

Normalização **na gravação**, num único ponto (`modValidacao.Normalizar`), aplicada igual no front e no executor. O mesmo valor entra igual, venha de onde vier.

| Grupo de campo | Regra | Campos |
|---|---|---|
| Identificação e cadastro curto | **MAIÚSCULAS**, `Trim`, espaço duplo colapsado | `empresa`, `cidade`, `uf`, `maquina`, `orcamento` |
| Nome de pessoa e cargo | `Trim` e espaço duplo; **preserva a digitação** | `nome`, `cargo`, `responsavel` |
| E-mail | **minúsculas**, `Trim` | `email` |
| CNPJ e telefone | **só dígitos** no banco; máscara na exibição; telefone com 10 ou 11 dígitos, **um número por campo** | `cnpj`, `telefone` |
| Lista suspensa | valor exato da tabela `listas`, sem transformação | `etapa`, `familia`, `origem`, `categoria`, `prioridade`, `motivo_desfecho`, `qualificacao` |
| Texto livre e memo | **como digitado**, só `Trim` nas pontas | `observacoes`, `prox_acao`, `ctx_resumo` |

**Por que não tudo em caixa alta.** O motivo usual — padronizar busca e evitar duplicata — não se aplica: o ACE compara texto **sem diferenciar maiúsculas**, então `LIKE '%tubular%'` já encontra `TUBULARES EXEMPLO`. O ganho é visual, e nos campos longos o custo é real: um `CONTEXTO-ATENDIMENTO` de 8.000 caracteres em caixa alta fica ilegível, e texto em caixa alta não volta para capitalização correta de forma automática — siglas, preposições e nomes como `INDUSTRIA D'ANGELO` se perdem. E-mail em maiúsculas também não: a parte local é formalmente sensível a caixa, e o campo é copiado direto para o cliente de e-mail.

Onde o padrão vale, ele vale inteiro: razão social, cidade e UF gravam em maiúsculas, como já vêm do cartão CNPJ e como já estão nos nomes das pastas de cliente.

**Acentuação é preservada.** Maiúscula com acento é gravada e comparada normalmente; remover acento perderia dado sem ganho.

**A normalização não renomeia pasta.** Se o nome normalizado divergir do nome da pasta em `02 - CLIENTES`, a pasta permanece — POP-COM-001. Renomear quebra ponteiro.

**Telefone é o caso que a carga revelou.** Na planilha, 344 cadastros de cliente e 12 de contato traziam mais de um número no campo, e quase sempre o nome de quem atende junto: `5133330000 | FULANO: 51999990000`, `(51) 3333-1000 / (51) 3333-2000 / (51) 99999-3000`, e um deles com um CNPJ no meio. Aplicar "só dígitos" sobre isso produziria um número que não existe e que ninguém consegue discar.

Regra: a carga extrai **o primeiro número válido** para o campo (12 ou 13 dígitos começando com 55 perdem o código do país) e manda **o nome e os números adicionais para `observacoes`**, numa linha marcada `[migração 16/09/2026]`. Nada se perde, e o campo fica discável.

No front, a mesma regra vira crítica: telefone com menos de 10 ou mais de 11 dígitos é recusado na gravação, com a mensagem dizendo onde o nome e o segundo número devem ir. É o que impede o problema de voltar.

**`prox_acao` é MEMO, não `TEXT(255)`.** Cinco atendimentos já têm entre 272 e 389 caracteres nesse campo — o vendedor escreve parágrafos, não uma linha.

**Na carga inicial**, a normalização é aplicada e o relatório informa quantos registros mudaram por campo, para conferência antes da virada. O `extrair.py` confere todo valor contra o limite da coluna do banco e **aborta** se algum estourar: descobrir isso no meio da carga custa a migração inteira.

### 5.2 Vocabulário de SQL disponível

O ACE via OLE DB **não é o Access**: funções da biblioteca do Access não existem fora dele. Levantado por execução em 16/09/2026 — tabela fechada, não hipótese:

| Construção | Resultado | Uso no sistema |
|---|---|---|
| `IIF` | OK | Campos calculados e agregações condicionais |
| `IS NULL` | OK | Pré-clientes, campos não preenchidos |
| **`NZ`** | **FALHA** — "Função 'NZ' indefinida na expressão" | Tratar nulo do lado do cliente |
| `NOW`, `DATE` | OK | Carimbos e fila da manhã |
| `DATEDIFF` | OK | Dias parado, ciclo |
| `DATEADD` | OK | Retomar em |
| `FORMAT` | OK | Mês de referência |
| `YEAR` / `MONTH` | OK | Mês de referência — alternativa ao `FORMAT` |
| `LEN`, `LEFT`, `MID`, `VAL` | OK | Decompor `AT-CCCC-NNNN` |
| `&` (concatenação) | OK | Montar código visual |
| `SWITCH` | OK | Quadro, situação |
| `TOP n` | OK | Paginação da grade |
| Subconsulta, `LEFT JOIN`, `HAVING` | OK | Consultas do Painel |

**Regra que vem daí:** tratamento de nulo é feito em VBA ou PowerShell, nunca em SQL. Em vez de `NZ(MAX(codigo_cliente),0)`, usa-se `MAX(codigo_cliente)` e o nulo é tratado ao receber.

### 5.3 Datas: nunca literal, sempre parâmetro

O teste devolveu `FORMAT(#01/03/2026#,'yyyymm')` = **202601** e `DATEADD('d',7,#01/01/2026#)` = **08/01/2026**. O literal de data no SQL do ACE é **mm/dd/yyyy**, enquanto a estação escreve e exibe dd/mm/yyyy.

Consequência: um literal `#03/01/2026#` digitado pensando em "3 de janeiro" vira **1º de março**, sem erro, sem aviso, e a oportunidade aparece no mês errado do Painel.

> **Regra:** nenhuma data entra em SQL como literal. Sempre parâmetro tipado (`OleDbType.Date` no PowerShell, `adDate` no VBA). Vale para consulta, gravação e filtro de Painel — sem exceção para "só um `WHERE` rápido".

---

## 6. O contexto do cliente

**Decidido:** o texto continua no arquivo; o banco guarda ponteiro, resumo e carimbo.

| Onde | O que fica |
|---|---|
| `02 - CLIENTES\NNNN - RAZAO\CONTEXTO-GERAL.md` | O texto integral, com FATO / INFERÊNCIA / HIPÓTESE e fonte |
| `...\02 - ATENDIMENTOS\AT-CCCC-NNNN\CONTEXTO-ATENDIMENTO.md` | O quadro tático da negociação |
| `clientes.ctx_*` e `oportunidades.ctx_*` | Resumo de uma linha, família aderente, caminho do arquivo, data e autor |

O memo de 8.880 caracteres atravessa o ACE intacto — a escolha não é técnica, é de desenho. O arquivo permanece a fonte porque é escrito direto pela sessão, sem passar por fila nem esperar estação; é legível sem ferramenta, comparável entre versões, e é sobre ele que o POP-COM-001 e a IT-COM-002 estão escritos.

O que entra no banco é o que serve à tela e à consulta: o vendedor vê o resumo na grade e responde "quais clientes têm aderência confirmada em Curvadora de Tubos" por SQL, sem abrir 466 arquivos.

**A R5 do POP-COM-001 permanece intacta:** CRM dono do estado, pasta dona do conteúdo. O `ctx_resumo` é ponteiro com etiqueta, não segunda fonte.

Quem escreve: a skill `zapromaq-contexto-cliente` grava o `.md` como faz hoje e emite um comando de fila atualizando os quatro campos `ctx_*`. Se o comando ainda não foi aplicado, o carimbo no banco fica atrasado em relação ao arquivo — e é exatamente por isso que o carimbo existe.

---

## 7. Frontend

### 7.1 Princípios

1. **Conexão curta por operação do usuário**, nunca por linha (seção 2.4).
2. **Consulta e edição separadas.** A grade é de leitura; editar exige ação explícita, que relê o registro antes de abrir.
3. **Controle de versão em toda gravação.** `UPDATE ... WHERE id = ? AND versao = ?`. Zero linhas afetadas significa que alguém alterou: cancela, mostra a diferença e preserva o que foi digitado. Nunca sobrescreve em silêncio.
4. **Dado e auditoria na mesma transação.** Comprovado no teste; o log deixa de ser comando solto.
5. **Cor nunca é o único sinal.** "Atrasada", "Hoje", "Encerrada" também em texto.
6. **Nada de OCX externo**, ListView ou controle de calendário que exija instalação.
7. Layout validado em 1366 × 768, e em escala do Windows de 100%, 125% e 150%.

### 7.2 Telas

| Tela | Conteúdo |
|---|---|
| **Início** | Fila da manhã (ações vencidas e de hoje), atalhos, data da última consulta bem-sucedida |
| **Oportunidades** | Grade filtrada por responsável, etapa e próxima ação; ficha em cinco seções |
| **Clientes** | Grade com filtro por estágio (Pré-cliente / Cliente), responsável, UF, segmento; ficha com identificação, localização, prospecção, observações e contexto |
| **Contatos** | Grade e ficha com busca de empresa por nome, CNPJ ou código |
| **Painel** | Indicadores — seção 8 |
| **Listas** | Administração das opções, acesso do gestor |

Ficha de Oportunidades, em cinco seções: Cliente e contato (leitura, vindos por JOIN) · Negociação · Acompanhamento · Próxima ação e desfecho · Observações e contexto. Cabeçalho fixo com código, empresa, etapa e situação; rodapé fixo com Salvar, Cancelar e mensagem de estado.

### 7.3 Módulos VBA

| Módulo | Responsabilidade |
|---|---|
| `modConfig` | Caminho UNC do banco, versão do front, compatibilidade de esquema |
| `modDB` | Conexão, **comandos parametrizados**, transação, liberação de recursos |
| `modSchema` | Metadados dos campos: seção, ordem, largura, obrigatoriedade, editabilidade, vínculo |
| `modCRM` | Regras de cliente, contato, oportunidade, códigos e promoção |
| `modCalc` | Situação, quadro, dias parado, ciclo, meses |
| `modPainel` | Consultas agregadas |
| `modValidacao` | Campos, datas, valores e vínculos |
| `modMenu` | Navegação e abertura de fichas |

**Parâmetros ADO, sempre.** Concatenar valor em SQL é o que quebra com apóstrofo em razão social — e `INDUSTRIA D'ANGELO LTDA` já está na base.

### 7.4 O que muda em relação ao estudo inicial

| Ponto | Estudo inicial | Aqui |
|---|---|---|
| Recuperar id novo | `SELECT MAX(id)` em conexão nova | `@@IDENTITY` na mesma conexão, dentro da transação |
| Log | Comando solto após o INSERT | Mesma transação do dado |
| Edição simultânea | Não tratada | `versao` conferida no `WHERE` |
| SQL | Literais concatenados | Parâmetros ADO |
| Chave de cliente | Código obrigatório como PK | `id` interno + código opcional |
| Abas espelho | Peça central | **Eliminadas** — seção 8 |

---

## 8. Painel

Refeito. As abas espelho e as fórmulas `INDEX/MATCH` são descartadas: a versão 3.6 carrega cerca de 90.000 fórmulas só em Oportunidades, custo que não faz sentido preservar quando o dado está em banco.

**Como passa a funcionar:** `modPainel` executa cerca de dez consultas `GROUP BY` no banco, numa única conexão, e escreve o resultado numa aba `BasePainel` oculta — algumas dezenas de linhas, não 5.000. O `Painel` lê dali.

Blocos mantidos: fila de qualificação (pré-clientes por aderência e porte), prospecção, funil por etapa, situação de acompanhamento, família, categoria, responsável, origem e indicadores gerais, com metas vindas da tabela `metas`.

Regras:

- Data e hora da última atualização visíveis no topo.
- Botão *Atualizar indicadores* separado de *Salvar*.
- Os totais vêm de agregação no banco, **nunca da página de 100 registros da grade**.
- Se a atualização falhar, mantém o conjunto anterior marcado como desatualizado. Não mostra número velho como atual.
- Clicar num indicador abre Oportunidades com o filtro correspondente.

`Gráficos`, `BaseGraficos`, `Visitas` e a aba `Metas` antiga saem do escopo desta fase. As dependências que o Painel 3.6 tinha em `Visitas` (G84) e `Metas` (G103:G105) são substituídas: as metas passam à tabela `metas`, e o indicador de visitas sai do Painel nesta fase, em vez de exibir dado congelado.

---

## 9. Executor e fila

A sessão de IA lê o banco direto, mas **não escreve nele** — nenhuma biblioteca fora do Windows implementa o protocolo de bloqueio do ACE, e gravar por fora enquanto os vendedores usam o front é risco de corromper o arquivo.

**Fluxo:**

1. A sessão grava um `.json` em `BASE\FILA\`, com carimbo, autor e a lista de operações.
2. O executor (`aplicar-fila.ps1`) roda na estação — por tarefa agendada e por botão no front.
3. Para cada comando: valida, aplica em transação com controle de versão, grava o log com valor antigo e novo.
4. Move o arquivo para `APLICADOS\` com o resultado anexado. Comando recusado vai para `APLICADOS\RECUSADOS\` com o motivo.

**Os campos `ctx_*` não entram no controle de versão, de propósito.** Eles não são editáveis pelo formulário (`editavel = 0` no schema), então o `UPDATE` do vendedor nunca os toca, e o `UPDATE` da fila nunca toca nos campos dele. Campos disjuntos, sem conflito possível: a fila pode gravar contexto com a ficha aberta na tela de alguém, e nenhum dos dois perde trabalho. O `cad_cliente`, que mexe em campo do vendedor, **incrementa a versão** como qualquer outra gravação.

O executor **nunca apaga registro** e só altera os campos declarados no comando. O conjunto de operações permitidas é fechado: atualizar `ctx_*` de cliente, atualizar `ctx_*` de oportunidade, e corrigir campo cadastral de cliente. Qualquer outra coisa é recusada.

Scripts: `aplicar-fila.ps1` e `conferir-base.ps1`. Algo entre 200 e 300 linhas somadas — não é um segundo sistema.

---

## 10. Backup, manutenção e limites

| Item | Definição |
|---|---|
| Backup horário | Cópia para `99 - BACKUP`, retenção 48 h |
| Backup diário | Fora do expediente, **com janela exclusiva** (ninguém conectado), retenção 30 dias |
| Restauração | Testada na etapa 7 e repetida trimestralmente. Backup não testado não é backup |
| Compactação | Mensal, com todos fora do sistema |
| Rede | Cabeada onde possível. Wi-Fi caindo durante gravação é a causa residual |
| Segurança | Permissão de rede por usuário na pasta do banco. **Não existe segurança por usuário no `.accdb`**: quem tem escrita na pasta contorna o front. Proteger abas é organização, não segurança |
| Anexos | Documento e PDF ficam na pasta do cliente, **nunca dentro do banco** |

---

## 11. Migração

### 11.1 Ordem

Listas → Clientes → Contatos → Oportunidades. Uma transação por tabela; falha em qualquer ponto desfaz a tabela inteira.

### 11.2 Dicionário — Clientes

| Origem (3.6) | Destino | Regra |
|---|---|---|
| `CÓD. CLIENTE` | `codigo_cliente` | Vazio → NULL, e `estagio` = `Pré-cliente` |
| — | `estagio` | Derivado do código |
| `EMPRESA` | `empresa` | Obrigatório; linha sem empresa não migra |
| `CIDADE`, `ESTADO`, `SEGMENTO`, `TELEFONE GERAL`, `E-MAIL GERAL`, `CNPJ` | idem | Direto |
| `RESPONSÁVEL` | `responsavel` | Direto |
| `ADERÊNCIA`, `PORTE` | `aderencia`, `porte` | Só numérico 1–3; o registro com espaços vira NULL |
| `QUALIFICAÇÃO` | `qualificacao` | Direto |
| `OBSERVAÇÕES` | `observacoes` | Direto |
| `PRIORIDADE` | — | Fórmula; recalculada na exibição |
| `VALIDAÇÃO DO CNPJ` | — | Fórmula; revalidada no front |
| `ORDEM`, `SITUAÇÃO CADASTRO` | — | **Descontinuados** — ver 11.3 |

### 11.3 Campos descontinuados

`ORDEM` tem 4 valores em 855 linhas; `SITUAÇÃO CADASTRO` tem 2. Nenhum dos dois está em uso efetivo. `ativo` (YESNO) assume o papel da situação cadastral, com `True` para todos na carga. Se `ORDEM` voltar a ser necessária, entra como campo novo — uma linha no schema e uma coluna no banco.

### 11.4 Dicionário — Contatos e Oportunidades

Contatos: `CÓD. CONTATO` → `codigo` (congelado) · `CONTROLE` → `controle` · `CÓD. CLIENTE` → resolvido para `clientes.id` pelo mapa · `CONTATO` → `nome` · `CARGO`, `TELEFONE`, `E-MAIL`, `OBSERVAÇÕES` diretos · `EMPRESA` descartada (vem por JOIN).

Oportunidades: `ATENDIMENTO` → `codigo` (congelado) · `CONTROLE` → `controle` · `CÓD. CONTATO` → `id_contato` pelo mapa · `CÓD. CLIENTE` → `id_cliente` · `ATENDIMENTO ANTERIOR` → `id_atendimento_anterior` (hoje vazio em todas) · `ORÇAMENTO`, `ETAPA`, `RESPONSÁVEL`, `MÁQUINA / DESCRIÇÃO`, `FAMÍLIA`, `VALOR`, `ORIGEM`, `PRIORIDADE`, as seis datas, `TENTATIVAS`, `PRÓXIMA AÇÃO`, `MOTIVO DE DESFECHO`, `OBSERVAÇÕES`, `CATEGORIA`, `TIPO DE VENDA` diretos · `SITUAÇÃO`, `QUADRO`, `DIAS PARADO`, `CICLO`, `MÊS *` e as colunas de JOIN descartadas.

> **O mapeamento é por nome de coluna, lido do cabeçalho da linha 3.** Nunca por posição. Foi o que quebrou o `modMigrar` do estudo inicial: ele lia por índice fixo do layout antigo e teria gravado e-mail no campo de valor, sem erro visível.

### 11.5 Conferência obrigatória

Depois da carga, antes de qualquer uso: contagem por tabela igual à origem; soma de `VALOR` idêntica; contagem por etapa idêntica à distribuição de 2.5; nenhum código `CT`/`AT` alterado em relação à planilha; nenhum órfão. Divergência de um registro interrompe a virada.

---

## 12. Roteiro de execução

A migração é feita **numa única sessão de trabalho**, de ponta a ponta. Este roteiro existe para não se perder no meio: são doze passos em ordem obrigatória, cada um com uma verificação que diz se pode seguir. Passo sem verificação cumprida não avança — volta.

**Antes de começar:** ninguém usando a planilha, e cópia de `CRM-Comercial-Zapromaq-3.6.xlsm` guardada em `99 - HISTORICO` com a data. É o ponto de retorno.

| # | Passo | Como saber que terminou |
|---|---|---|
| 1 | ~~Rodar `teste-complementar.ps1`~~ — **concluído em 16/09/2026** | Tabela de funções fechada na seção 5.2; regra de data na 5.3 |
| 2 | Fechar o dicionário de dados em `40 - DICIONARIO` | Todo campo das abas Clientes, Contatos e Oportunidades tem destino ou descarte declarado, por **nome de coluna** |
| 3 | `criar-banco.ps1` — tabelas, índices e listas | Tabelas criadas; `listas` populada a partir da aba `Listas` da 3.6 |
| 4 | `migrar.ps1` — carga em banco de **teste**, uma transação por tabela | Contagem por tabela igual à origem; soma de `VALOR` idêntica; distribuição por etapa igual à de 2.5; nenhum `CT`/`AT` alterado; nenhum órfão |
| 5 | Aplicar integridade referencial e conferir de novo | Exclusão de cliente com contato recusada pelo banco |
| 6 | `conferir-base.ps1` — relatório de normalização de texto | Quantos registros mudaram por campo; nenhuma mudança inesperada |
| 7 | Front: `modConfig`, `modDB`, `modSchema`, `modValidacao`, `modCalc` | `modCalc` reproduz situação, quadro, dias parado, ciclo e meses **idênticos** aos da planilha, numa amostra de 20 oportunidades |
| 8 | Front: ficha e grade de **Oportunidades** | Lançar, editar e encerrar um atendimento; conflito de versão tratado; nenhuma conexão aberta dentro de laço |
| 9 | Front: **Clientes**, **Contatos**, **Listas** e promoção | Pré-cliente promovido recebe o próximo código e muda de estágio na mesma transação |
| 10 | `modPainel` e aba `BasePainel` | Cada indicador confere com a mesma consulta rodada à parte |
| 11 | `aplicar-fila.ps1` e binários do `mdbtools` versionados | Comando de fila aplicado com log de valor antigo e novo; comando inválido recusado com motivo; contagem do `mdbtools` igual à do ACE |
| 12 | Teste de falha, backup e virada | Dois usuários no mesmo atendimento tratados; queda de rede durante gravação não duplica nem perde; backup restaurado; carga final conferida |

**Os três pontos onde se para e se pensa**, em vez de seguir no impulso:

- **Depois do passo 6** — o banco carregou e confere? Se não, o problema está no dicionário, e insistir só propaga o erro.
- **Depois do passo 10** — os números batem com a planilha? Divergência de indicador aqui é divergência de regra, não de tela.
- **Depois do passo 12** — a virada só acontece com backup restaurado com sucesso. Backup não testado não é backup.

**Regra de congelamento:** a partir do passo 2, nenhuma alteração de estrutura na planilha 3.6. Coluna nova na origem no meio do processo invalida o dicionário e obriga a recomeçar do passo 2.

**Se a sessão precisar ser interrompida**, o estado fica registrado em `40 - DICIONARIO\estado-migracao.md`: último passo concluído, verificação que passou e o que falta. A retomada começa refazendo a verificação do último passo, nunca confiando na memória de onde parou.

---

## 13. Critérios de aceite

- Empresa sem código permanece na fila e pode ser promovida pelo fluxo definido.
- Ordenar ou filtrar a grade não altera código nem vínculo.
- Busca paginada não reduz o total do Painel.
- Dois usuários editando o mesmo atendimento recebem tratamento de conflito, sem sobrescrita silenciosa.
- Falha depois de gravar e antes de confirmar não duplica registro na nova tentativa.
- Cancelar não grava; fechar com alteração pendente avisa.
- Opção de lista desativada continua legível no histórico e não é oferecida em lançamento novo.
- Cliente ou contato com vínculo não é excluído pelo fluxo comum.
- Data brasileira, moeda, acento, apóstrofo, campo vazio e observação longa preservados.
- Falha na atualização mantém o último painel completo, marcado como desatualizado.
- Restauração de backup preserva registros, relações e auditoria.
- Os 479 códigos `AT` e os 496 `CT` continuam idênticos após a migração.

---

## 14. Riscos aceitos

| Risco | Por que é aceito |
|---|---|
| `.accdb` em pasta de rede continua sendo arquivo compartilhado; queda de rede durante gravação ainda pode corromper | O risco cai muito em relação à pasta compartilhada, mas não chega a zero. Mitigação: rede cabeada, backup horário, compactação mensal, restauração testada |
| Não há segurança por usuário no Access | Controle por permissão de rede na pasta. Se isolamento por vendedor virar requisito, a arquitetura precisa ser revista |
| Manutenção concentrada em uma pessoa | Documentação nas fontes `.md`, SQL fora dos botões, módulos com responsabilidade separada. Reduz, não elimina |
| `mdbtools` é reimplementação independente | Conferência de contagem contra o ACE na etapa 7, e repetida quando a versão mudar |
| Migração futura para SQL Server Express não é só trocar a string de conexão | Consultas, tipos, identidade e transação exigem adaptação. Concentrar o acesso em `modDB` e `modCRM` reduz o trabalho |

---

## 15. Onde ficam os arquivos do sistema

Tudo o que constrói e mantém o CRM fica em `01 - PADROES\00 - IA\MIGRACAO-CRM`. É de lá que sai qualquer alteração futura — o `.xlsm` distribuído e o `.accdb` em produção são resultado, não fonte.

```
01 - PADROES\00 - IA\MIGRACAO-CRM\
  2026-09-16_ESTUDO-MIGRACAO-CRM.md    este documento
  10 - FONTES-VBA\                      .bas e .frm exportados do VBE — a fonte editável
  20 - SCRIPTS\                         criar-banco.ps1, migrar.ps1, aplicar-fila.ps1,
                                        conferir-base.ps1, exportar-espelho.ps1
  30 - TESTES\                          teste-ace.ps1, teste-complementar.ps1,
                                        checar-ambiente.ps1 e os relatorios de execucao
  40 - DICIONARIO\                      dicionario-dados.md, listas-iniciais.csv,
                                        estado-migracao.md
  99 - HISTORICO\                       versao substituida, com sufixo de data
```

**Regras:**

1. **A fonte é o `.bas`, não o VBE.** Alterou módulo no Excel, exporta para `10 - FONTES-VBA` na mesma sessão. Código que só existe dentro do `.xlsm` se perde quando o arquivo corrompe.
2. **A versão do front aparece na tela** e é a mesma gravada em `modConfig`. Front com esquema incompatível não grava.
3. **Versão substituída vai para `99 - HISTORICO`** com sufixo de data, conforme a regra transitória em vigor.
4. **Relatório de teste não se apaga.** É a evidência de que a decisão foi medida, não presumida.
5. Nada aqui escreve letra de unidade. O caminho é o UNC, guardado na aba de configuração do front.

### Material de origem

| Arquivo | Papel |
|---|---|
| `RELATORIO-PROJETO.md` | Estudo inicial. Arquitetura aproveitada; mapeamento de colunas, formato de código e recuperação de id **não** aproveitados |
| `2026-09-16_RELATORIO-FRONTEND-EXCEL-ACCESS.md` | Estudo de frontend. Princípios de interface, contrato de gravação e critérios de aceite aproveitados |
| `30 - TESTES\resultado-teste-20260916-082018.txt` | Teste de viabilidade executado — fonte dos números da seção 2 |
| `CRM-Comercial-Zapromaq-3.6.xlsm` | Base funcional |
