---

## Neste projeto: migração de banco

Mudança de dado ou de estrutura do `.accdb` é uma migração numerada em
`src/BASE/SISTEMA/DADOS/banco/migracoes/NNN-descricao.ps1` (`EsquemaDe`/`EsquemaPara`; estrutural
exige banco exclusivo e muda `config.versao_esquema`, o que obriga a atualizar front e MCP na
mesma versão). O build aplica as pendentes antes de publicar. Detalhes em
`specs/stacks/banco-access.md`. Códigos `CT-`/`AT-` são congelados: nunca migre o formato.
