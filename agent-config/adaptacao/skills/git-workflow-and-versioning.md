---

## Neste projeto: git-flow (substitui o trunk-based acima)

- `main` = o que está publicado na empresa (cada publicação com tag `vX.Y`); `develop` =
  integração. Nada de commit direto nelas.
- `git flow feature|bugfix start <nome>` a partir de `develop`; `release/X.Y` e `hotfix/<nome>`
  a partir da `main`. Merges `--no-ff`.
- **Finalizar sempre com** `powershell -File agent-config\ferramentas\finalizar-branch.ps1 [-Enviar]`:
  o `git flow ... finish` instalado (GitFlow .NET 2.3.0) quebra em todo merge `--no-ff`.
- Mensagens em pt-BR, dizendo o porquê; terminam com as linhas de atribuição da sessão.
  Nunca `--no-verify`.
- `.gitattributes` tem `* -text`: o git não converte fim de linha. Use editor/script que
  preserve CRLF e Windows-1252 nos fontes VBA (o `sed` do Git Bash remove CR).
- `agent-config/vendor/agent-skills` é *git subtree*: não edite; atualize com
  `ferramentas\atualizar-agent-skills.ps1`.
