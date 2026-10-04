## Neste projeto

Sem `package.json` nem CI: a barra de qualidade é a do `ferramentas\verificar-tudo.ps1` e das
seções **Revisão** das specs de stack. Se for criar `CONSTRAINTS.md`, ele fica em
`agent-config/specs/CONSTRAINTS.md` e os comandos registrados são os níveis do `verificar-tudo`.
Ferramentas que exigem internet ou Node (Semgrep em nuvem, Lighthouse, axe) não entram no
fluxo do sistema; no desenvolvimento, só com aprovação do usuário.
