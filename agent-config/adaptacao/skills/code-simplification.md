---

## Neste projeto: proteções e cercas de Chesterton

**Blocos protegidos.** Nunca altere o que estiver entre os marcadores (o gancho do original não
está ligado; vale por instrução):

```vb
' simplify-ignore-start: <motivo>
...
' simplify-ignore-end
```

Em PowerShell: `# simplify-ignore-start: <motivo>` / `# simplify-ignore-end`.

**Padrões intencionais — não "simplifique":**
- `modDB.Nz` próprio (o `NZ()` do Access não existe via OLE DB) e tratamento de nulo no código.
- Um `Add` por campo em `modSchema` (limite de 25 continuações de linha do VBA).
- Eventos das abas em `ThisWorkbook` (`Workbook_Sheet*`): aba criada por automação não tem CodeName.
- Uma conexão por operação do usuário (`AbrirLeitura`/`ConsultarEm`/`FecharLeitura`), caches por
  sessão, índice de campos por dicionário — são correções de tempo de resposta (PLANO §4c).
- `ErrorActionPreference = 'Continue'` em scripts que chamam `git`/`schtasks`: no PowerShell 5.1
  o stderr nativo vira exceção com `Stop`.
- Normalização só em `modValidacao`; `ctx_*` gravados só pela IA.

Antes de remover algo que parece sem motivo: `git log -S`, `docs/legado/estado-migracao.md` e
`PLANO.md` §3–4c. Guia de linguagem do original (TS/Python/React) não se aplica; aqui é VBA 7
e PowerShell 5.1 — ver as specs de stack.
