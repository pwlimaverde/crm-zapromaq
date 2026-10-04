---

## Neste projeto: superfície real

Não há servidor web, login nem segredo no sistema. Os riscos daqui são:

- **Injeção de SQL:** todo SQL parametrizado (`modDB.P`, `New-Parametro`); a base tem razão social
  com apóstrofo. Data nunca como literal (`#mm/dd/yyyy#` inverte dia e mês sem erro).
- **Integridade:** `UPDATE ... WHERE id=? AND versao=?` (0 linhas = conflito, nunca sobrescrever);
  dado e `log_alteracoes` na mesma transação.
- **MCP:** catálogo fechado de ferramentas, parâmetros validados por esquema, gravação com
  controle de versão e `origem = 'IA'`; o MCP de desenvolvimento só grava em `SISTEMA\DADOS\fonte`.
- **Vazamento:** repositório público — nenhum dado real, nome de estação ou valor de carteira;
  backups e logs da rede nunca entram no git.
- **Arquivos e .bat:** caminhos montados com `Join-Path`, aspas em todo caminho de `.bat`,
  nada de letra de unidade fixa; o `.accdb` não tem segurança por usuário (proteção é a
  permissão da pasta de rede).
- **Dependências:** nada de terceiros nas estações (sem OCX/DLL, sem pacote, sem internet).
