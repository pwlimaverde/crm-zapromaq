# Spec: Lote atualiza duplicados por CNPJ (ciclo 015)

> Situação: aprovada em 23/10/2026 (plano revisado de 19/10/2026: DEV-26, decisões assumidas 2 a 9). Sem migração.
> Depende de DEV-12 (lote com as regras do cadastro), DEV-09 (conversão) e DEV-10 (normalização).

## Objetivo

Hoje a linha do lote cujo CNPJ já existe na base vira DUPLICADO e não entra. O comercial recebe a planilha de
qualificação com dados mais novos dessas empresas e precisa levá-los ao cadastro sem redigitar:

- **Botão "Atualizar duplicados"** (frmLote): atualiza todas as linhas DUPLICADO por CNPJ da base, cada uma na sua
  transação, com resultado por linha (ATUALIZADO, SEM MUDANÇA, CONFLITO, ERRO).
- **Duplo clique** numa linha DUPLICADO atualizável mostra a **diferença** (campo: atual → novo) e pede confirmação
  antes de atualizar só aquela empresa.
- Regras (decisões do plano): duplicado só por CNPJ (nome parecido continua SUSPEITO e nunca atualiza); CNPJ repetido no
  próprio lote: vale a primeira linha, a repetida não atualiza; **campo vazio no lote não apaga**; **observações são
  acrescentadas** (linha nova; se o texto já está lá, nada muda); **nome da empresa do lote substitui** o cadastrado;
  **empresa inativa não é atualizada**; log com ação **"ALTERAR LOTE"** (um registro por campo, antigo e novo, na
  mesma transação), com versão lida no momento da atualização (conflito não sobrescreve).
- Campos que o lote atualiza: empresa, cidade, UF, segmento, telefone, e-mail, qualificação, aderência, porte e
  observações (acréscimo). CNPJ é a chave (não muda). Contato e atendimento continuam fora do lote.

## Critérios de aceite

- [ ] `Confere` em `modLote.MudancasDoLote`: vazio não apaga; empresa substitui; observação acrescentada numa linha
  nova; observação já presente não duplica; valor igual não entra.
- [ ] testar-uso: lote com CNPJ de empresa ativa → DUPLICADO atualizável; "Atualizar duplicados" grava só os campos
  preenchidos, acrescenta a observação, sobe a versão e loga "ALTERAR LOTE"; empresa inativa com o mesmo CNPJ não é
  atualizada; linha repetida no lote não atualiza; segunda execução = SEM MUDANÇA; diferença da linha antes de
  atualizar lista os campos.
- 🖥 frmLote na estação: botão, duplo clique com a confirmação da diferença, resumo.

## Decisões assumidas

As decisões 2 a 9 do plano revisado. Novas:

| # | Decisão | Premissa | O que mudaria |
|---|---|---|---|
| 015-1 | A diferença é recalculada contra o banco **na hora** de atualizar (não na conferência), com a versão lida ali. | Entre Conferir e Atualizar outra estação pode ter gravado. | — |
| 015-2 | Linha DUPLICADO com outro erro de campo (ERRO) não atualiza. | ERRO tem precedência (`Peso`). | — |
| 015-3 | Sem mudança no conector: o lote é só do front. | DEV-26 fala do lote. | Ferramenta de lote no conector, se pedida. |

## Limites

- Nunca: apagar valor existente pelo lote; atualizar por nome; transferir contato.
