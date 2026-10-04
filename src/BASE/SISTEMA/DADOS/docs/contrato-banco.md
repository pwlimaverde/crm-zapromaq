# Contrato do banco — regras que todo cliente do `crm_zapromaq.accdb` segue

Dois sistemas gravam no mesmo arquivo Access, ao mesmo tempo: o **front** (VBA na
planilha, uma cópia por estação) e a **IA** (servidor MCP em `BASE\IA`, pelo Claude
Desktop). Não há servidor de banco para impor regra: quem garante a integridade é
este contrato. Qualquer terceiro cliente futuro segue as mesmas regras.

## 1. Motor e conexão

- Só o motor **ACE** (`Microsoft.ACE.OLEDB.16.0`, fallback 12.0) grava no arquivo.
  Nada de mdbtools, UCanAccess, Jackcess ou cópia de arquivo para gravar: só o ACE
  respeita o bloqueio do `.laccdb`. `Microsoft.Jet.OLEDB.4.0` é proibido.
- **Conexão curta**: abre, executa, fecha — por operação do usuário ou por chamada
  de ferramenta. Nunca uma conexão por linha, campo ou item de combo (26 ms por ciclo
  na rede). Scripts .NET usam `OLE DB Services=-4` para não segurar o arquivo em pool.

## 2. SQL

- **Sempre parametrizado** (`?` posicional). A base tem `INDUSTRIA D'ANGELO LTDA`.
- **Data nunca como literal**: o SQL do ACE lê `#mm/dd/yyyy#`; a estação escreve dd/mm.
  Data é sempre parâmetro tipado (`adDate` / `OleDbType.Date`).
- `NZ()` não existe via OLE DB: nulo se trata no código.
- `TOP n` **não desempata**: todo `ORDER BY` com `TOP` termina numa chave única (`id`, `codigo`).
- `ORDER BY` do Access põe nulo na frente: para "vazio por último", `IIf(x Is Null,1,0), x`.

## 3. Versões

| Controle | Onde | Regra |
|---|---|---|
| Versão do **esquema** | `config.versao_esquema` | Front grava só se for exatamente a esperada (`modConfig.VERSAO_ESQUEMA`); a IA grava só se estiver na lista que conhece. Mudou a estrutura → migração em `banco\migracoes` + publicar front **e** IA juntos. |
| Versão do **registro** | coluna `versao` | Todo `UPDATE` de campo do vendedor é `... SET ..., versao=versao+1 WHERE id=? AND versao=?`. Zero linhas = alguém alterou depois da leitura: **não sobrescrever**, avisar e reler. |
| Versão do **front** | `config.versao_front` | Gravada pelo build ao publicar; a cópia local compara e avisa se ficou para trás. |

Campos `ctx_*` são escritos só pela IA e nunca pelo formulário: não entram no
controle de versão (não há conflito possível entre os dois).

## 4. Auditoria

- **Toda gravação** (inserção, alteração, ativação, exclusão, promoção, contexto)
  grava `log_alteracoes` **na mesma transação** do dado. Sem exceção: o log é também o
  "feed" que faz as estações se atualizarem sozinhas.
- `origem`: `FRONT` (planilha) ou `IA` (servidor MCP). `usuario`: usuário do Windows
  (front) ou `IA (nome de quem pediu)`.
- Alteração de registro existente: uma linha por **campo alterado**, com valor antigo e novo.

## 5. Normalização (ponto único por cliente: `modValidacao.Normalizar` / `IA\mcp\lib\Regras.ps1`)

| Campos | Regra |
|---|---|
| `empresa`, `cidade`, `uf`, `maquina`, `orcamento` | MAIÚSCULAS, trim, espaço duplo colapsado |
| `nome`, `cargo`, `responsavel` | trim; preserva a digitação |
| `email` | minúsculas |
| `cnpj`, `telefone` | só dígitos; telefone com 10 ou 11 dígitos, um número por campo |
| listas (`etapa`, `familia`, `origem`, `prioridade`, `motivo_desfecho`, `segmento`, `uf`, `qualificacao`, `categoria`) | valor exato da tabela `listas` |
| memo (`observacoes`, `prox_acao`, `ctx_resumo`) | como digitado, só trim |

Texto vazio grava **NULL**, nunca cadeia vazia (o índice único de CNPJ aceita vários nulos).
Acento é preservado.

## 6. Códigos

- `CT-CCCC-NNNN` e `AT-CCCC-NNNN` são **congelados** na criação e nomeiam pastas em
  `02 - CLIENTES`: nunca recalcular nem reformatar.
- `clientes.codigo_cliente` só existe para `Cliente`; `estagio` é derivado (sem código →
  `Pré-cliente`) e nunca digitado. Promoção atribui `MAX+1` na mesma transação, com
  `WHERE codigo_cliente IS NULL` para dois usuários não levarem o mesmo número.
- Números sequenciais (`controle`, `codigo_cliente`) são calculados **dentro da transação
  que insere**; colisão no índice único é repetida com o próximo número, não mostrada como erro.

## 7. O que nenhum cliente faz

- Apagar cliente com contato ou contato com atendimento (as chaves estrangeiras recusam).
- Excluir atendimento (encerra-se pela etapa).
- Gravar fora do ACE, fora de transação ou sem log.
