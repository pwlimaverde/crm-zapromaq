# Stack: banco — Access `.accdb` / ACE OLE DB 16 / ADO

## Papel

`crm_zapromaq.accdb` na raiz de `BASE` (pasta de rede, caminho UNC). Única fonte dos dados;
o front e a IA acessam por OLE DB.

## Versões e ambiente

- Provedor `Microsoft.ACE.OLEDB.16.0` (fallback 12.0), Office/ACE **64 bits**.
- **Proibido** `Microsoft.Jet.OLEDB.4.0` (gera JET4 com extensão `.accdb`, erro silencioso).
- VBA: ADO late binding. PowerShell: `System.Data.OleDb` (mesma semântica de tipo e data).
- Nesta máquina: ACE do Access Database Engine 2016 x64 + Excel 365.

## Modelo

Tabelas: `clientes`, `contatos`, `oportunidades`, `listas`, `metas`, `log_alteracoes`, `config`
(DDL em `banco/esquema/criar-banco.ps1`; dicionário em `docs/contrato-banco.md`).
`clientes.id` é a identidade; `codigo_cliente` só para `Cliente`; `estagio` derivado (sem código →
`Pré-cliente`). Códigos `CT-CCCC-NNNN`/`AT-CCCC-NNNN` congelados na criação (nomeiam ~480 pastas).
Campos calculados (situação, quadro, dias parado, ciclo, meses) não existem no banco.

## Onde fica

```
src/BASE/SISTEMA/DADOS/banco/
  esquema/criar-banco.ps1          DDL (banco novo)
  migracoes/NNN-descricao.ps1      migrações numeradas
  aplicar-migracoes.ps1 + APLICAR-MIGRACOES.bat
src/BASE/SISTEMA/DADOS/lib/Comum.ps1   Open-Banco, Invoke-Comando, Invoke-Escalar, New-Parametro
```

## Comandos

```powershell
powershell -File src\BASE\SISTEMA\DADOS\banco\esquema\criar-banco.ps1 -Banco <accdb> -Recriar
powershell -File src\BASE\SISTEMA\DADOS\banco\aplicar-migracoes.ps1 [-Banco <accdb>] [-Simular]
powershell -File src\BASE\SISTEMA\DADOS\testes\criar-banco-demo.ps1 -Recriar   # 20 de cada, fictício
powershell -File src\BASE\IA\teste\criar-banco-teste.ps1 -Recriar              # o dos testes do MCP
```

## Migrações

Arquivo `NNN-descricao.ps1` que devolve `@{ Id; Descricao; EsquemaDe; EsquemaPara; Aplicar = { param($cn,$tx,$log) } }`.
- Só **dados** (`EsquemaPara` vazio): roda em transação, pode rodar com gente usando, deve ser
  idempotente.
- **Estrutural** (`EsquemaPara`): exige banco exclusivo (sem `.laccdb`; DDL do ACE não é
  transacional), muda `config.versao_esquema` → front (`modConfig.VERSAO_ESQUEMA`) e MCP (que
  confere a versão antes de gravar) atualizam na **mesma** publicação.
- Backup antes de migrar; aplicadas ficam em `config` (`migracao.NNN`). O build aplica as pendentes.

## Estilo e regras

- **SQL sempre parametrizado** (`?` posicional; `modDB.P(tipo, valor)` / `New-Parametro`).
- **Data nunca literal** (`#mm/dd/yyyy#` lê ao contrário da estação, sem erro).
- `NZ()` não existe via OLE DB: nulo se trata no código. `ORDER BY` põe nulo na frente
  (`IIF(campo IS NULL,1,0)` antes).
- `UPDATE ... WHERE id=? AND versao=?`, `versao = versao + 1`; 0 linhas = conflito, nunca sobrescrever.
- Dado e `log_alteracoes` na mesma transação (`modDB.LogEm`); `origem` = `FORM`/`IA`/`MIGRACAO`...
- `@@IDENTITY` na mesma conexão/transação do `INSERT`.
- Campo vazio grava `NULL` (índice único de CNPJ aceita vários nulos, não várias strings vazias).
- `TOP n` sempre com chave única no fim do `ORDER BY`.

## Testes

Casos no `IA/teste/testar-mcp.ps1` sobre o banco fictício de `criar-banco-teste.ps1` (o
teste restaura o que altera). Migração nova: teste aplicando-a no banco de teste e conferindo
o resultado; rodar duas vezes prova a idempotência.

## Limites

- Sempre: conexão curta por operação do usuário (~26 ms por ciclo na rede); backup antes de
  migrar; idempotência nas migrações de dados.
- Perguntar antes: qualquer migração estrutural; apagar dado; mudar índice único; tocar o
  banco de produção.
- Nunca: Jet 4.0; consulta por linha; literal de data; recalcular códigos `CT-`/`AT-`;
  dado real em teste ou documento.

## Revisão

- [ ] Parâmetros com o tipo certo (data como `adDate`/`OleDbType.Date`, moeda como Currency).
- [ ] Controle de versão e log na transação.
- [ ] Migração com `EsquemaDe`/`EsquemaPara` corretos e idempotente.
- [ ] DDL, `modSchema`, consultas do `modCRM` e MCP coerentes (o `verificar.py` confere parte).

## Referências offline

`referencias/` 06-access-sql-ace, 07-ado, 08-dotnet-oledb.
