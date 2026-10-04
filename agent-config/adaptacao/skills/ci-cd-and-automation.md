---

## Neste projeto: não há servidor de CI

O "pipeline" é local: `agent-config\ferramentas\verificar-tudo.ps1` no desenvolvimento e
`MONTAR-FRONTEND.bat` (backup → migrações → montagem → compilação → autoteste → teste de uso →
publicação) na rede. Um GitHub Actions só poderia rodar a parte estática (`verificar.py`,
`testar-build-lib`, MCP `-SoProtocolo`) num runner Windows — propor só com aprovação do usuário.
