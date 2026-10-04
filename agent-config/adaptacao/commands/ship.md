## Neste projeto: o que é "entregar"

Há dois níveis — confirme com o usuário qual é:

1. **Integrar a branch:** `verificar-tudo` (padrão; `-Completo` se tocou `fonte\`/`build\`) →
   revisão → `ferramentas\finalizar-branch.ps1 -Enviar` (merge `--no-ff` em `develop` + push).
2. **Publicar na empresa:** `git flow release start X.Y` (X.Y = a versão que a rede vai publicar:
   a de `VERSAO-FRONT.txt` de lá + 1) → `verificar-tudo -Completo` →
   `ferramentas\empacotar-base.ps1` → o usuário extrai o pacote **por cima** da BASE da rede e
   roda `MONTAR-FRONTEND.bat` → só então `finalizar-branch.ps1 -Enviar` (main + tag `vX.Y`).

Itens extras do go/no-go: pacote conferido pelo script; migrações revisadas (estrutural exige
banco exclusivo); a BASE da rede foi trazida antes (a versão da rede tem prioridade);
plano de volta = `execucao\backup\antes-da-publicacao\<data>` + `execucao\historico-front`.
Acessibilidade web e Core Web Vitals não se aplicam.
