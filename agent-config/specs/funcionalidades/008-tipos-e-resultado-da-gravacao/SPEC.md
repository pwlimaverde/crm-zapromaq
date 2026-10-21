# Spec: Tipos, resultado da gravação e falha transitória (ciclo 008)

> Situação: aprovada em 19/10/2026 (plano revisado de 19/10/2026, seção 5, Etapa 2)

## Objetivo

- **DEV-09** — conversão de tipos sem perda silenciosa: vazio, válido e inválido são estados distintos; inválido
  não grava nem apaga o valor anterior e a ficha mantém o texto digitado; data em formato explícito; moeda com
  convenção única; inteiro sem fração e dentro do limite. Mesmas regras no lote e no conector.
- **DEV-13** (parte sem migração) — estados da gravação na ficha: recusada antes de gravar, falhou antes do
  commit (nada gravado), confirmada (e tela não atualizada), indeterminada (falha no commit); depois de confirmada,
  o front guarda id/código/versão e nunca diz "nada foi gravado" por falha de releitura; falha ao fechar a conexão
  é separada da confirmação; botão Salvar bloqueado durante a gravação. O identificador de operação (idempotência)
  exige a migração 1.2 e vai no ciclo 010.
- **DEV-14** — repetição automática só de erro transitório (bloqueio do ACE) e só antes do commit, com espera
  crescente e limite; mensagem distingue "não gravou, pode repetir", "resultado indeterminado" e erro de regra.

## Critérios de aceite

- [ ] Matriz C17–C21 (front `Confere` e conector `testar-mcp`; o `verificar.py` exige os dois lados):
  C17 data inexistente (31/02/2026, 2026-02-31) recusada; C18 moeda com mais de 2 casas recusada (12,505 / 12.505);
  C19 inteiro com fração recusado (1,5 / 1.5); C20 valor acima de 999.999.999.999,99 recusado; C21 valor negativo recusado.
- [ ] Front, texto digitado (`modValidacao.Converter`): `abc` inválido; `12.50` inválido ("use vírgula para os
  centavos"); `1,5` = 1,5; `1.250` = 1250; `1.250,50` = 1250,5; `1,250.50` (separadores mistos) inválido; `-10`
  convertido (a crítica recusa negativo); `0` válido; data `dd/mm/aaaa` e `aaaa-mm-dd`; vazio = Null (limpar de
  propósito); a ficha recusa salvar com campo inválido, mostrando campo e texto, sem apagar o que foi digitado.
- [ ] Lote: aderência/porte pelo mesmo conversor de inteiro (nota fracionária é erro).
- [ ] Conector: `ConvertTo-Moeda` e `ConvertTo-Inteiro` no atendimento, cadastro e metas (nada de `[int]1.5` = 2).
- [ ] DEV-13: `modDB.EstadoGravacao` (enviada → confirmada | indeterminada | desfeita); `MensagemResultadoGravacao`
  (pura, `Confere`) com as quatro mensagens; a ficha usa o estado no tratador de erro; releitura protegida depois de
  gravar; `Confirmar` trata o `Close` à parte; reentrada no Salvar bloqueada.
- [ ] DEV-14: conector repete a transação inteira em erro de bloqueio antes do commit (3 tentativas, 200/500/1000 ms),
  nunca depois; teste com duas conexões reais (uma segura a linha, o conector espera e grava); esgotado, "não gravou,
  pode repetir". Front: `ExecutarEm` repete o comando em erro de bloqueio (mesma transação, antes do commit).

## Estrutura

`modValidacao.bas`, `modDB.bas`, `modLote.bas`, `frmCRM.txt`, `modAutoteste.bas`; `IA\mcp\lib\Regras.ps1`, `Banco.ps1`,
`Ferramentas.ps1`; `IA\teste\testar-mcp.ps1`; `SISTEMA\DADOS\testes\testar-recuperacao.ps1` (rajada com bloqueio).

## Estratégia de teste

C17–C21 nos dois lados; `Confere` do conversor e das mensagens; `testar-recuperacao` com bloqueio real de linha;
`testar-uso` (build) segue verde. 🖥: digitar `12.50` no valor e ver a recusa sem perder o texto; salvar com a rede
caindo depois do commit (difícil de provocar: anotar se ocorrer).

## Decisões assumidas

| # | Decisão | Premissa | O que mudaria |
|---|---|---|---|
| 008-1 | Moeda: vírgula é o separador decimal; ponto só como milhar em grupos de 3 (`1.250.000`); `12.50` é inválido (não vira 1250 nem 12,5). Até 2 casas decimais; mais é inválido (sem arredondar em silêncio). | Estações em pt-BR; o erro histórico era `12.50` → `1250`. | Se o comercial digitar com ponto decimal, aceitar `12.50` quando houver exatamente 2 dígitos após o ponto. |
| 008-2 | Limite de moeda 999.999.999.999,99 (Currency do ACE suporta mais; o limite é de plausibilidade). | Máquinas industriais, nunca trilhões. | Subir o limite numa constante. |
| 008-3 | Inteiro: só dígitos com sinal opcional, sem separador; limite do Long (±2.147.483.647). | Campos inteiros são notas, tentativas e metas. | — |
| 008-4 | Data digitada: `dd/mm/aaaa` (dia e mês com 1 ou 2 dígitos) ou `aaaa-mm-dd`; `IsDate` não é mais usado (aceitava formatos ambíguos). | Mesma regra do conector (ISO). | — |
| 008-5 | O conector repete só erro de **bloqueio** (`bloqueado`, `locked`, `em uso`, `could not update`); "consulta atualizável" com backup em curso não repete (a cópia leva segundos, mas a sessão deve ver a mensagem e decidir). | Erro do backup já tem mensagem controlada "pode repetir". | Repetir também durante o backup, com espera maior. |
| 008-6 | Estado "indeterminado" só quando a falha é no próprio `CommitTrans`; falha no `Close` depois do commit é registrada no diagnóstico e o resultado é "confirmado". | No ACE o commit é síncrono no arquivo. | — |
