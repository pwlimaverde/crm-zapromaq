# Registro da migração da planilha 3.6 para o banco (16/09/2026)

**Histórico — não executar.** Estes são os scripts que fizeram a migração da planilha
compartilhada `CRM-Comercial-Zapromaq-3.6.xlsm` para o `crm_zapromaq.accdb`, e os testes
de viabilidade que sustentaram a decisão. Ficam aqui como evidência do que foi medido e
de como os dados foram transformados (ver `docs\legado\estado-migracao.md`).

Os caminhos dentro deles são os da estrutura antiga (`01 - PADROES\00 - IA\MIGRACAO-CRM`,
"sobe cinco níveis"); não funcionam na estrutura nova e não foram adaptados de propósito.

| Arquivo | Papel na migração |
|---|---|
| `teste-ace.ps1`, `teste-complementar.ps1` | Viabilidade do ACE na rede e tabela de funções SQL (16/09) |
| `extrair.py`, `gerar-dicionario.py` | Planilha 3.6 → JSON de carga normalizado; dicionário de dados |
| `criar-banco.ps1` (antigo), `carregar.ps1`, `aplicar-integridade.ps1`, `conferir-base.ps1` | Criação, carga, chaves estrangeiras e conferência |
| `conferir-calculados.py` | `modCalc` × valores da planilha (487 de 487 iguais) |
| `aplicar-indice-cnpj.ps1` | Índice único de CNPJ em produção (17/09) |
| `aplicar-fila.ps1` | Executor da fila da IA — **descontinuado** (substituído pelo MCP em `BASE\IA`) |
| `*.bat` | Roteiro de execução da migração (teste → produção → virada) |

A estrutura que esses scripts produziram é o **esquema 1.0**, recriado hoje por
`banco\esquema\criar-banco.ps1`. Mudança de estrutura daqui em diante: `banco\migracoes`.
