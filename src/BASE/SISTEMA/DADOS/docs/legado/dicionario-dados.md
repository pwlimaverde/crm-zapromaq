# Dicionário de dados — migração do CRM 3.6

**Gerado por `gerar-dicionario.py` em 16/09/2026 12:19**, conferido contra os cabeçalhos reais de `crm36.xlsm`.

O mapeamento é **por nome de coluna**, nunca por posição. Destino `—` significa descarte declarado.

## Clientes

17 colunas na origem, 17 declaradas.

| Coluna de origem | Destino | Tipo | Regra |
|---|---|---|---|
| `CÓD. CLIENTE` | `codigo_cliente` | LONG | Vazio → NULL e estagio = Pré-cliente |
| `EMPRESA` | `empresa` | TEXT(120) | MAIÚSCULAS, trim, espaço duplo colapsado. Obrigatório |
| `CIDADE` | `cidade` | TEXT(60) | MAIÚSCULAS, trim |
| `ESTADO` | `uf` | TEXT(2) | MAIÚSCULAS, 2 caracteres |
| `SEGMENTO` | `segmento` | TEXT(60) | Valor da lista, sem transformação |
| `TELEFONE GERAL` | `telefone` | TEXT(30) | Só dígitos; máscara na exibição |
| `E-MAIL GERAL` | `email` | TEXT(120) | minúsculas, trim |
| `CNPJ` | `cnpj` | TEXT(20) | Só dígitos; diferente de 14 → NULL com aviso |
| `RESPONSÁVEL` | `responsavel` | TEXT(60) | Valor da lista |
| `ADERÊNCIA` | `aderencia` | INTEGER | Só 1, 2 ou 3; qualquer outra coisa → NULL |
| `PORTE` | `porte` | INTEGER | Só 1, 2 ou 3 |
| `QUALIFICAÇÃO` | `qualificacao` | TEXT(40) | Valor da lista |
| `OBSERVAÇÕES` | `observacoes` | MEMO | Como digitado, só trim nas pontas |
| `PRIORIDADE` | `—` | — | Fórmula na origem; recalculada na exibição a partir de qualificacao |
| `VALIDAÇÃO DO CNPJ` | `—` | — | Fórmula na origem; revalidada no front |
| `ORDEM` | `—` | — | Descontinuado: 4 valores preenchidos em 855 linhas |
| `SITUAÇÃO CADASTRO` | `ativo` | YESNO | Descontinuado como texto; ativo = True para todos na carga |

## Contatos

10 colunas na origem, 10 declaradas.

| Coluna de origem | Destino | Tipo | Regra |
|---|---|---|---|
| `CÓD. CONTATO` | `codigo` | TEXT(15) | Congelado como está. Nomeia vínculo e pasta |
| `CONTROLE` | `controle` | LONG | Preservado; novos recebem MAX+1 |
| `CÓD. CLIENTE` | `id_cliente` | LONG | Resolvido para clientes.id pelo mapa da carga |
| `CONTATO` | `nome` | TEXT(120) | Como digitado; vazio → GERAL, registrado como exceção |
| `CARGO` | `cargo` | TEXT(60) | Como digitado |
| `TELEFONE` | `telefone` | TEXT(30) | Só dígitos |
| `E-MAIL` | `email` | TEXT(120) | minúsculas |
| `OBSERVAÇÕES` | `observacoes` | MEMO | Como digitado |
| `EMPRESA` | `—` | — | Fórmula na origem; vem por JOIN |
| `SITUAÇÃO CADASTRO` | `ativo` | YESNO | ativo = True na carga |

## Oportunidades

40 colunas na origem, 40 declaradas.

| Coluna de origem | Destino | Tipo | Regra |
|---|---|---|---|
| `ATENDIMENTO` | `codigo` | TEXT(15) | Congelado. Nomeia a pasta do atendimento |
| `CONTROLE` | `controle` | LONG | Preservado; novos recebem MAX+1 |
| `CÓD. CONTATO` | `id_contato` | LONG | Resolvido para contatos.id |
| `CÓD. CLIENTE` | `id_cliente` | LONG | Congelado na criação |
| `ATENDIMENTO ANTERIOR` | `id_atendimento_anterior` | LONG | Resolvido pelo código; hoje vazio em todas as 487 |
| `ORÇAMENTO` | `orcamento` | TEXT(30) | MAIÚSCULAS, trim |
| `ETAPA` | `etapa` | TEXT(40) | Valor da lista. Obrigatório |
| `RESPONSÁVEL` | `responsavel` | TEXT(60) | Valor da lista |
| `MÁQUINA / DESCRIÇÃO` | `maquina` | TEXT(255) | MAIÚSCULAS, trim |
| `FAMÍLIA DE MÁQUINA` | `familia` | TEXT(60) | Valor da lista |
| `CATEGORIA` | `categoria` | TEXT(60) | Valor da lista |
| `TIPO DE VENDA` | `tipo_venda` | TEXT(40) | Como digitado |
| `VALOR (R$)` | `valor` | CURRENCY | Vazio permanece NULL — nunca vira zero |
| `ORIGEM` | `origem` | TEXT(40) | Valor da lista |
| `PRIORIDADE` | `prioridade` | TEXT(20) | Valor da lista |
| `DATA DE ENTRADA` | `dt_entrada` | DATETIME | Parâmetro tipado, nunca literal |
| `DATA DA PROPOSTA` | `dt_proposta` | DATETIME | Idem |
| `ÚLTIMA INTERAÇÃO` | `ultima_interacao` | DATETIME | Idem |
| `TENTATIVAS` | `tentativas` | INTEGER | Vazio → NULL |
| `PRÓXIMA AÇÃO` | `prox_acao` | TEXT(255) | Como digitado |
| `DATA PRÓX. AÇÃO` | `dt_prox_acao` | DATETIME | Idem |
| `RETOMAR EM` | `retomar_em` | DATETIME | Idem |
| `MOTIVO DE DESFECHO` | `motivo_desfecho` | TEXT(60) | Valor da lista |
| `DATA DO DESFECHO` | `dt_desfecho` | DATETIME | Idem |
| `OBSERVAÇÕES` | `observacoes` | MEMO | Como digitado |
| `SITUAÇÃO` | `—` | — | Calculado na exibição (modCalc) |
| `QUADRO` | `—` | — | Calculado na exibição |
| `DIAS PARADO` | `—` | — | Calculado na exibição |
| `CICLO (DIAS)` | `—` | — | Calculado na exibição |
| `MÊS PROPOSTA` | `—` | — | Calculado na exibição |
| `MÊS ENTRADA` | `—` | — | Calculado na exibição |
| `MÊS DESFECHO` | `—` | — | Calculado na exibição |
| `EMPRESA` | `—` | — | JOIN com clientes |
| `CIDADE` | `—` | — | JOIN com clientes |
| `ESTADO` | `—` | — | JOIN com clientes |
| `SEGMENTO` | `—` | — | JOIN com clientes |
| `CONTATO` | `—` | — | JOIN com contatos |
| `CARGO` | `—` | — | JOIN com contatos |
| `TELEFONE` | `—` | — | JOIN com contatos |
| `E-MAIL` | `—` | — | JOIN com contatos |

## Conferência

Todas as colunas das três abas têm destino ou descarte declarado. Nenhuma coluna declarada está ausente da origem.
