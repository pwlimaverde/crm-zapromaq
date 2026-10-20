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
| **Manutenção** (ciclo 007, DEV-08) | `config.manutencao` | `estação\|usuário\|PID\|início\|motivo`. Preenchida, front (`modDB.AbrirTransacao`) e conector (`Invoke-Transacao`, antes da 1ª escrita) recusam **toda** gravação, conferindo dentro da transação junto com o esquema; a leitura continua. Liga/desliga: `aplicar-migracoes.ps1` (no lugar); `-LiberarManutencao` desliga à mão. |
| **Backup em curso** (ciclo 007, DEV-03) | arquivo `SISTEMA\DADOS\execucao\backup\copiando.txt` | Enquanto existe, a cópia completa segura o `.accdb` (`FileShare.Read`): o ACE abre só leitura e a gravação falha antes de gravar; front e conector mostram "cópia de segurança em andamento… nada foi gravado". |

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
| listas (`etapa`, `familia`, `origem`, `prioridade`, `motivo_desfecho`, `segmento`, `uf`, `qualificacao`, `categoria`; nas visitas `tipo` = `TipoVisita` e `resultado` = `ResultadoVisita`) | valor exato da tabela `listas`; valor inativo só se já era o gravado no registro |
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
- Como o `AT-` guarda o código do cliente, o atendimento **não troca de empresa**: a troca de
  contato e o atendimento anterior (`oportunidades.id_atendimento_anterior`, opcional, nunca o
  próprio) só aceitam a mesma empresa — no front (`modCRM.CriticarTrocaContato`,
  `CriticarAnterior`) e na IA (`atualizar_atendimento`, `abrir_atendimento`).

## 7. O que nenhum cliente faz

- Apagar cliente com contato ou contato com atendimento (as chaves estrangeiras recusam).
- Excluir atendimento (encerra-se pela etapa).
- Gravar fora do ACE, fora de transação ou sem log.
- Apagar ou renomear valor de lista (só incluir, desativar/reativar e reordenar).
- Excluir visita ou planejamento (visita se cancela; planejamento se encerra).

## 8. Esquema 1.1 (migração 003)

O front espera exatamente `1.1`; o conector aceita `1.0` e `1.1` (`EsquemasSuportados` em
`IA\mcp\lib\Banco.ps1`). Num banco 1.0 as ferramentas de visitas, planejamentos e municípios não
têm as tabelas; `exportar_roteiro` recusa com a mensagem do esquema.

### Tabelas novas

| Tabela | Colunas principais | Observações |
|---|---|---|
| `municipios` | `cod_ibge` (PK), `nome`, `nome_busca`, `uf`, `lat`, `lon`, `fonte_carga` | Base de referência do IBGE (`banco\referencia\municipios.csv`), carregada pela migração; só leitura para front e IA. Índice (`uf`, `nome_busca`). |
| `planejamentos` | `id`, `codigo` (único, `PL-AAAA-NNN`), `controle`, `descricao`, `cod_municipio_base`, `raio_km`, `data_inicio`, `data_fim`, `responsavel`, `situacao` (`ABERTO`/`ENCERRADO`), `observacoes`, `versao`, `criado_em`, `alterado_em`, `alterado_por` | Código gerado na criação (`MAX(controle)+1` dentro da transação) e congelado. |
| `visitas` | `id`, `id_oportunidade` (obrigatório), `id_planejamento`, `situacao` (`PENDENTE`/`REALIZADA`/`CANCELADA`), `data_agendada`, `hora_agendada`, `responsavel`, `tipo`, `observacoes`, `data_realizada`, `resultado`, `relato`, `motivo_cancelamento`, `versao`, `criado_em`, `alterado_em`, `alterado_por` | Situação exibida: `A AGENDAR` (pendente sem data), `AGENDADA` (pendente com data), `REALIZADA`, `CANCELADA`. |

### `listas`: chave e papel das etapas

- Colunas novas: `chave` (obrigatória; índice único `tipo` + `chave`), `encerra`, `resultado`
  (`GANHO`/`PERDA`/nulo), `quadro` (`PROSPECCAO`/`FUNIL`/nulo), `exige_familia`. Sim/Não sempre
  gravado explicitamente.
- **Chave** = maiúsculas, sem acento, espaços aparados e colapsados (`modValidacao.ChaveLista` /
  `Regras.ps1`): dois valores com a mesma chave no mesmo tipo são recusados (C02).
- Listas novas: `TipoVisita` e `ResultadoVisita`.
- O programa **não compara texto fixo de etapa**: encerrar, resultado (ganho/perda), quadro e
  exigência de família vêm das colunas da etapa. Regras: resultado só em etapa que encerra (C04);
  não se desativa a última etapa ativa GANHO ou PERDA (C03); UF só entre as 27 (C05). Papel inicial
  gravado pela migração: Pedido Fechado (encerra, GANHO, FUNIL), Perdido (encerra, PERDA, FUNIL),
  Descartado (encerra, PERDA, PROSPECCAO); Contato Inicial, Sem Retorno e Retorno Agendado
  (PROSPECCAO); Elaboração da Proposta, Negociação, Proposta em Análise e Proposta em Stand By
  (FUNIL, exige família); as demais, FUNIL.

### Regras de gravação

| Gravação | Regra |
|---|---|
| Visita | Atendimento encerrado não recebe visita (C08); com planejamento, a data agendada cai no período dele (C09); REALIZADA exige data e relato (C10) e grava, na mesma transação, a linha `VISITA dd/MM/aaaa HH:mm - ...` no histórico do atendimento e a `ultima_interacao` (V6/V7, com a versão do atendimento); CANCELADA exige motivo (C11); no máximo uma PENDENTE por atendimento (C13). |
| Planejamento | Raio de 1 a 500 km (C14), cidade-base existente em `municipios`, período com início até o fim (C15); edição só enquanto ABERTO; encerrar com pendentes só cancelando-as (C16, motivo "planejamento encerrado"). |
| Gerar visitas e lotes (conector) | Uma transação; C13 conferida de novo dentro dela; qualquer recusa desfaz tudo. Lotes de até 100 itens, `simular` = true por padrão. |
| `visitas`, `planejamentos` | `WHERE id=? AND versao=?`; zero linhas = conflito. |
| `listas` | `WHERE id=?` com os valores antigos de `ativo`, `ordem`, `encerra`, `resultado`, `quadro` e `exige_familia`. |

Toda gravação dessas tabelas entra em `log_alteracoes` na mesma transação (tabelas `visitas`,
`planejamentos`, `listas`): é o que faz as outras estações atualizarem as telas abertas e a aba
Planejamentos (feed, até 25 s).
