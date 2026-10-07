# Funcionalidades

Uma pasta por trabalho: `NNN-nome-curto/` (NNN sequencial, 001, 002...), criada pelo `/spec`
a partir de `_modelo/`:

```
NNN-nome-curto/
  SPEC.md    o quê e por quê (critérios de aceite), referências às specs de stack
  plan.md    fatias, dependências, checkpoints (/plan)
  todo.md    lista de tarefas marcada durante o /build
```

Ao concluir, no último commit da branch (antes do merge), a SPEC recebe no topo:
`> Concluída em dd/mm/aaaa — feature/<nome>` (ou `bugfix/<nome>`). O nome da branch identifica
o merge em `develop` (`git log --merges --grep <nome>`); o hash não existe ainda nesse momento.
Quando a versão for publicada pelo `/ship`, acrescente `, versão publicada X.Y`.
A pasta fica como registro (não se apaga).
