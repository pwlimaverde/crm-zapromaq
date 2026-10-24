# Tarefas: ciclo 017

- [x] Conector: `duplicar_atendimento` (simular, gravar pelo caminho do `abrir_atendimento`) - 1.11.0
- [x] Front: `modCRM.DadosParaDuplicar` (nada grava até o Salvar)
- [x] Tela: menu da linha "Duplicar atendimento" + `frmCRM.PrepararDuplicacao`
- [x] verificar.py: lista de campos igual nos dois canais; docs e `verificar-tudo -Completo`

## Revisão final da rodada (ciclos 010 a 017, cinco eixos)

Achados e destino (sem crítico aberto):

- **Corrigido (correção, 010):** o teste de uso DEV-13 travava o build -Teste (dicionário do PowerShell por
  Application.Run) - módulo de teste injetado.
- **Corrigido (correção, 010):** `testar-conector-versoes` não montava a versão "incompatível" com o conector no 1.2.
- **Corrigido (correção, 012):** `Array.Sort(chaves, itens)` pelo PowerShell não reordenava os itens (paginação
  por cursor perdia linha) - chaves ordenadas e busca por dicionário; teste de exclusão entre páginas.
- **Conferido (segurança, 012):** descarte de extratos não segue junção para fora da raiz - teste de regressão.
- **Corrigido (docs):** barra invertida virando tabulação (`IA\teste`) no LEIA-ME e no PLANO; contagem de ferramentas
  na spec de stack.
- **Corrigido (teste, 011):** referências COM a `ListObject` remontada e id explícito no AutoNumber derrubavam o Excel
  no teste de uso (não afeta o produto).
- **FYI (correção):** troca da geração da base remonta as abas, mas não avisa a ficha aberta (o controle de versão
  continua impedindo sobrescrita).
- **FYI (correção):** a situação "Retomada a avaliar" vale para qualquer etapa que encerra com `retomar_em` (legado
  de Pedido Fechado com retomada também aparece); a seção do conector filtra só PERDA.
- **FYI (legibilidade):** `modLote` passou de 700 linhas; o `frmCRM` já era grande - separar a atualização de
  duplicados num módulo próprio se crescer de novo.
- **FYI (desempenho):** `EmpresasSemAtendimento` lê todos os clientes e ordena em O(n²) no VBA; com ~500 empresas
  no raio fica abaixo de um segundo; acima disso, ordenar no SQL.
- **FYI (teste):** o `testar-mcp` não é repetível no mesmo banco (contexto gravado fica); no `verificar-tudo` o banco
  é sempre novo.

## Revisão completa (ciclos 007 a 017, diff c955a1e..develop), 24/10/2026

- **Corrigido (b361c2a):** conector repetia como "bloqueio" erro de regra que cita "bloqueado" e a recusa por esquema.
- **Corrigido (e443230):** log da migração 005 sem estação, usuário técnico e versão (DEV-06B).
- **Corrigido (7d1d24c):** lote gravava segmento/qualificação na caixa colada; operação de registro excluído
  devolvia id inexistente (lote com identificador estável por linha).
- **Risco aceito:** cascata de ativo (front `MudarAtivo`, conector `Set-AtivoEm`) grava log mesmo se o UPDATE da
  linha afetar 0 (corrida de milissegundos na mesma transação); conector mantém a operação de registro excluído
  como "repetida" (o id_operacao do conector é por chamada, não por conteúdo).
