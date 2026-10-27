# Tarefas: ciclo 018

- [x] P1 botão Duplicar na ficha (feature/botao-duplicar)
- [x] P2 próxima ação só hoje ou futuro (feature/prox-acao-futura)
- [x] P3 retomada a avaliar igual nos dois canais (feature/retomada-avaliar)
- [x] P4 retomar_em desabilitado fora de Perdido/Descartado (feature/retomada-habilitada)
- [x] P5 restauração pela última cópia validada (feature/restaurar-validada)
- [x] P6 versão mínima do cliente no ponto de escrita (feature/versao-minima)
- [x] P7 skill cobre o catálogo inteiro (feature/skill-catalogo)
- [x] P8 checar-ambiente completo (feature/checar-ambiente)
- [x] P9 estados por linha do lote (feature/lote-estados)
- [x] P10 contato homônimo no conector (feature/homonimo-conector)
- [x] P11 checksum das migrações (feature/checksum-migracoes)
- [x] P12 fuso no diagnóstico do front (feature/fuso-diag)
- [x] P13 metas "não definida": não aplicável (nenhum canal mostra meta x realizado) - D38 no PLANO
- [x] Documentação final (feature/docs-finais)
- [ ] `verificar-tudo -Completo` e push

## Revisão final do ciclo 018 (cinco eixos, diff 82686b0..develop), 27/10/2026

- **Corrigido durante o ciclo (correção, P1):** a duplicação com contato inativo abria a escolha genérica de
  empresa e contato antes da ficha aparecer (travou o teste de uso); agora abre com a empresa da origem.
- **Corrigido durante o ciclo (teste, P6):** o testar-mcp gravava a versão mínima com `$script:VersaoServidor`
  vazio (variável do processo do servidor, não do teste); passou a ler a versão do `servidor.ps1`.
- **Conferido (correção):** encerrar em Perdido/Descartado limpa `dt_prox_acao` antes da conferência C34 nos dois
  canais; a vencida gravada e reenviada igual não trava (teste dos dois lados).
- **Conferido (correção):** o lote só interrompe por commit indeterminado ou recusa antes de começar (estado
  `nao_enviada`); a recusa de identificador de operação reutilizado (também ERRO_ESCRITA_RECUSADA, mas depois de
  abrir a transação) continua só na linha.
- **Conferido (segurança):** SQL novo parametrizado (`DadosParaDuplicar`, homônimo, versão mínima); o
  `checar-ambiente` cria e apaga um arquivo temporário na pasta do banco (teste confere que não sobra).
- **Conferido (desempenho):** `modCalc.Situacao` passa a ler o resultado da etapa pelo mesmo cache de papel
  (`PapelEtapa`); a restauração calcula o hash só até achar a primeira cópia validada.
- **Risco aceito (operação, 🖥):** com `versao_minima_conector = 1.12.0` gravada pela publicação da 1.6, o
  conector antigo deixa de gravar até o 1.12.0 ser publicado - publicar os dois juntos (nota no `IA\LEIA-ME.md`).
- **FYI (legibilidade):** `frmCRM.txt` cresceu de novo (botão de ação, retomada); candidatos a sair para um
  módulo: `AjustarRetomada` e a escolha de contato por empresa.
