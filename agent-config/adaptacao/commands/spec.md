## Neste projeto: onde e como fica a spec

- Salve em `agent-config/specs/funcionalidades/NNN-nome-curto/SPEC.md` (NNN = próximo número
  livre), partindo de `agent-config/specs/funcionalidades/_modelo/SPEC.md` — não em `SPEC.md` na raiz.
- Tech Stack, Commands, Project Structure, Code Style e Testing Strategy: **referencie** as specs
  de stack (`agent-config/specs/stacks/`) e escreva só o que for específico da funcionalidade.
- Boundaries: as restrições inegociáveis de `agent-config/specs/PROJETO.md` valem sempre; liste
  só as específicas.
- Entrevista em pt-BR, uma ou duas perguntas por vez. Regra de negócio e decisão de interface
  são do usuário; consulte também `src/BASE/SISTEMA/DADOS/docs/legado/` (estudo da migração
  e diário de decisões) antes de propor algo que o legado já decidiu.
- A spec é commitada junto com o início da branch (`git flow feature start <nome>`).
