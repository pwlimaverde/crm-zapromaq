---
name: crm-zapromaq
description: Consulta e atualiza o CRM Comercial Zapromaq (clientes, contatos, atendimentos AT-0000-0000) pelas ferramentas do conector crm-zapromaq. Use quando o usuário pedir para ver, procurar, conferir ou alterar dados do CRM, atualizar etapa, próxima ação ou valor de um atendimento, corrigir cadastro de cliente ou registrar o resumo de contexto de um cliente ou atendimento.
---

# CRM Comercial Zapromaq

O banco do CRM é um arquivo Access na rede, usado ao mesmo tempo pelos vendedores
pela planilha do front. **Todo acesso é feito pelas ferramentas do conector
`crm-zapromaq`.** Nunca abra, copie ou grave o `.accdb` por outro caminho
(Python, mdbtools, cópia de arquivo): só o conector usa o motor que respeita o
bloqueio do arquivo, e gravar por fora pode corromper o banco de todos.

## Antes de começar

Chame `crm_status`. Se `gravacao_liberada` vier `false`, o banco está numa versão
de esquema que este conector não conhece: faça só leituras e avise o usuário.

## Ler

| Para | Ferramenta |
|---|---|
| Achar uma empresa | `buscar_clientes` (nome, cidade, segmento, CNPJ ou código) |
| Ficha da empresa, contatos e atendimentos | `obter_cliente` (`codigo_cliente`; pré-cliente não tem código: use `id_cliente`) |
| Atendimentos de um vendedor, etapa ou situação | `buscar_atendimentos` |
| Ficha de um atendimento | `obter_atendimento` com `AT-0000-0000` |
| Valores aceitos em listas | `listar_opcoes` |
| Quem alterou o quê | `ultimas_alteracoes` |

A **situação** (Ação atrasada, Retomar hoje, Ação nesta semana, Em dia, Sem próxima
ação, Encerrado) é calculada com a mesma regra da planilha.

## Alterar — sempre nesta ordem

1. **Leia o registro** (`obter_atendimento` / `obter_cliente`) e guarde a `versao`.
2. **Mostre ao usuário o que vai mudar** (valor atual → valor novo) e confirme.
3. **Grave** com a `versao` lida, enviando só os campos que mudam, e `autor` com o
   nome de quem pediu.
4. Se voltar **"Conflito"**: alguém alterou o registro depois da leitura. Não repita
   às cegas: leia de novo, mostre a diferença ao usuário e só então refaça.

Regras que o conector confere e recusa com mensagem:

- Campos de lista (`etapa`, `responsavel`, `familia`, `origem`, `prioridade`,
  `motivo_desfecho`, `segmento`, `uf`, `categoria`) só aceitam valores de `listar_opcoes`.
- Datas em **AAAA-MM-DD**. Valor em reais como número (`850000.50`).
- Etapa **Pedido Fechado, Perdido ou Descartado** exige `motivo_desfecho` e `dt_desfecho`.
- Etapas de proposta e negociação exigem `familia`.
- Telefone: DDD + número, 10 ou 11 dígitos, **um número por campo**. Nome de quem
  atende e segundo número vão em `observacoes_acrescentar`.
- CNPJ é conferido pelo dígito verificador e não pode repetir outra empresa.
- `observacoes_acrescentar` só acrescenta uma linha datada; nada é apagado.

Não existe ferramenta para excluir registro, criar cliente ou atendimento, nem para
mudar código ou estágio de cliente: isso é feito pelo front, pelos vendedores.

## Contexto do cliente e do atendimento

O texto completo fica nos arquivos `CONTEXTO-GERAL.md` (pasta do cliente) e
`CONTEXTO-ATENDIMENTO.md` (pasta do atendimento). No banco vai só o ponteiro:
`atualizar_contexto_cliente` (`ctx_resumo` de uma linha, `ctx_familia`, `ctx_arquivo`)
e `atualizar_contexto_atendimento`. Esses campos são só da IA e não conflitam com
o vendedor: podem ser gravados sem ler a versão.

## Resultado das gravações

Toda gravação entra no log de auditoria com origem **IA**. Informe ao usuário o
que foi gravado (campos, valor anterior e novo, versão nova).
