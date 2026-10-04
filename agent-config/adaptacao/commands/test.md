## Neste projeto: a suíte

```powershell
powershell -File ferramentas\verificar-tudo.ps1 -Rapido     # durante a tarefa
powershell -File ferramentas\verificar-tudo.ps1             # fim da tarefa (MCPs com banco)
powershell -File ferramentas\verificar-tudo.ps1 -Completo   # fim da funcionalidade / toda mudança em fonte\ ou build\
```

O teste que falha primeiro, por stack, está na tabela de `agent-config/specs/PROJETO.md`.
`browser-testing-with-devtools` não está instalado (não há front web). Esta máquina tem
Excel **365**: o `-Completo` compila e roda o autoteste aqui, mas a palavra final de
compatibilidade é o build numa estação com Excel **2019**.
