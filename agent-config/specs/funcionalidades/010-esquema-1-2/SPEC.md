# Spec: Esquema 1.2 (ciclo 010)

> Situação: aprovada em 21/10/2026 (plano revisado de 19/10/2026, seção 5, Etapa 2)

## Objetivo

Migração **004 estrutural** (esquema 1.1 → 1.2), aplicada **só pela montagem** (cópia de trabalho, backup validado,
conferências, troca), e o código que passa a depender dela:

- **DEV-16** — índice único de `controle` em `contatos` e `oportunidades` (a alocação MAX+1 repete com a colisão,
  como já faz o código); `id_atendimento_anterior` validado: existe, mesma empresa, não aponta para si e **não
  forma ciclo** (front e conector); `atualizar_atendimento` passa a aceitar `codigo_atendimento_anterior`.
- **DEV-17** — `ctx_versao` (clientes e oportunidades), crescente e independente da `versao` comercial; gravar
  contexto exige a `ctx_versao` lida; conflito devolve o contexto atual.
- **DEV-10** (parte estrutural) — `ctx_familia` passa a 255 caracteres; nenhum campo trunca em silêncio.
- **DEV-06 parte B** — `log_alteracoes` ganha `usuario_tecnico`, `estacao`, `versao_cliente` e `autor_declarado`
  (o `usuario` continua como hoje, por compatibilidade).
- **DEV-13** (identificador) — tabela `operacoes` (`id_operacao` único, tabela, id, código, versão, hash do conteúdo,
  instante, origem, estação): criação com o mesmo identificador e o mesmo conteúdo devolve o resultado anterior sem
  gravar de novo; identificador com conteúdo diferente é recusado. Front (ficha e lote) e conector.

## Critérios de aceite

- [ ] `004-esquema-1-2.ps1`: EsquemaDe 1.1, EsquemaPara 1.2; inventário ANTES da estrutura (controles repetidos,
  anteriores órfãos, de outra empresa, apontando para si ou em ciclo): com qualquer achado, a migração **para** com a
  lista (na montagem: passo 7, banco de produção intacto; no ensaio: no relatório). Idempotente. Testada em banco
  fictício com os casos (testar-migracao-004).
- [ ] Banco criado do zero (criar-banco + migrações) e banco migrado têm as mesmas restrições (conferências do passo 8).
- [ ] Front `VERSAO_ESQUEMA = 1.2`; conector `EsquemasSuportados = @('1.2')`; montagem confere 1.2.
- [ ] C33 anterior que forma ciclo é recusado (front `Confere` e conector `testar-mcp`).
- [ ] DEV-17: duas sessões leem o mesmo contexto; a primeira grava, a segunda recebe conflito com o conteúdo atual;
  edição comercial em paralelo não gera conflito de contexto (testar-mcp).
- [ ] DEV-06B: toda linha nova de log (front e conector) tem estação e usuário técnico; o conector grava o autor
  declarado separado.
- [ ] DEV-13: `abrir_atendimento`, `criar_contato` e `criar_cliente` aceitam `id_operacao`; repetir a chamada não
  cria outro registro nem outro log; mesmo id com conteúdo diferente é recusado (testar-mcp). Ficha: o Salvar de
  registro novo usa um identificador que só muda quando a digitação muda; lote: identificador estável por linha.

## Decisões assumidas

| # | Decisão | Premissa | O que mudaria |
|---|---|---|---|
| 010-1 | Inventário com achado **para** a migração (não corrige sozinho). | Corrigir dado real é operação, não desenvolvimento (seção 1 do plano). | Se houver casos na rede, um script de correção assistida entra antes da 1.6. |
| 010-2 | FK física `id_atendimento_anterior → oportunidades.id` não é criada: o ACE não impede ciclo nem empresa diferente, e a FK autorreferente bloqueia a exclusão em lote de testes; a regra fica no domínio (front e conector) e na conferência do passo 8. | Plano: "integridade física onde o ACE permitir e validação de domínio no restante". | Criar a FK se a gestão preferir a trava física. |
| 010-3 | `id_operacao` opcional nas ferramentas (aditivo); retenção de 60 dias, limpa pelo próprio conector ao gravar. | Janela de reenvio de minutos/horas. | — |
| 010-4 | `ctx_versao` exigida na gravação de contexto (mudança que reforça concorrência: a versão nova recusa sem versão). Publicar junto com as skills (nota de mudança). | DEV-17 e seção 7 do plano. | — |
| 010-5 | Ficha: identificador de operação = GUID gerado na 1ª tentativa de salvar um registro novo e mantido enquanto a digitação não muda; lote: `LOTE|cnpj|nome normalizado|cidade`. | Reenvio após resultado indeterminado não duplica. | — |
