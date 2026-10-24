# Spec: Duplicar atendimento (ciclo 017)

> Situação: aprovada em 24/10/2026 (plano revisado de 19/10/2026: DEV-28, decisões assumidas 15 a 19). Sem migração.
> Depende de DEV-11 (contato fixo/ativo) e DEV-16 (anterior sem ciclo).

## Objetivo

O vendedor abre um negócio novo para a mesma empresa e máquina (nova proposta, novo ano) e hoje redigita tudo.

- **Front**: "Duplicar atendimento" (menu da linha na grade de Atendimentos) abre a ficha em **atendimento novo** já
  preenchido por `modCRM.DadosParaDuplicar` — copia contato (se ativo), responsável, origem, prioridade, categoria,
  tipo de venda, máquina, família e última interação; **zera** valor, orçamento, data da proposta, próxima ação e sua
  data, desfecho, motivo, retomada e tentativas; `dt_entrada` = hoje; etapa = a **etapa inicial** (primeira etapa ativa
  que não encerra, pela ordem da lista); `id_atendimento_anterior` aponta o atendimento de origem. **Nada grava até o
  Salvar** (decisão 15); duplicar encerrado é permitido (decisão 18).
- **Conector**: ferramenta nova `duplicar_atendimento` (aditiva, decisão 19): `codigo` do atendimento de origem,
  `simular` (padrão true: devolve só os dados da cópia), `id_operacao`, `autor`, `promover_pre_cliente`; gravar usa o
  mesmo caminho do `abrir_atendimento` (regras de contato, empresa, retomada, log, identificador de operação).
- **A regra da cópia num lugar só por canal, igual nos dois**: `modCRM.CAMPOS_DUPLICAR` e `$script:CamposDuplicar`;
  o `verificar.py` confere que as duas listas são iguais.

## Critérios de aceite

- [ ] testar-uso: `DadosParaDuplicar` copia os campos da lista, não traz os zerados, põe `dt_entrada` = hoje, etapa
  inicial e o anterior = origem; nada é gravado (contagem igual); contato inativo não é copiado.
- [ ] testar-mcp: `duplicar_atendimento` simula sem gravar; gravar cria o atendimento com os campos copiados, zerados,
  `dt_entrada` hoje, etapa inicial e `codigo_atendimento_anterior` = origem; origem encerrada é aceita; repetir com o
  mesmo `id_operacao` não duplica.
- [ ] verificar.py: lista de campos igual no front e no conector.
- [ ] Contrato: `VersaoServidor` 1.11.0, 43 ferramentas, nota, matriz, skill.
- 🖥 Menu da linha "Duplicar atendimento" abre a ficha preenchida; Cancelar não grava.

## Decisões assumidas

As decisões 15 a 19 do plano revisado. Novas:

| # | Decisão | Premissa | O que mudaria |
|---|---|---|---|
| 017-1 | O contato é copiado só se estiver ativo; senão a ficha pede o contato (fluxo normal). | DEV-11. | — |
| 017-2 | Etapa inicial = primeira etapa ativa que não encerra, pela `ordem` da lista. | Nenhum nome de etapa no código (spec 003). | Etapa inicial configurável. |
| 017-3 | No front, a entrada é o menu da linha da grade (sem botão novo na ficha). | O layout da ficha está no limite de 1366×768. | Botão na ficha, se houver espaço. |

## Limites

- Nunca: gravar na duplicação antes do Salvar; copiar valor, proposta, agenda ou desfecho.
