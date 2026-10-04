---

## Neste projeto: o número que governa

Um ciclo abre/consulta/fecha no `.accdb` da rede custa ~26 ms. Logo: **nenhuma consulta por
linha, por campo ou por combo**; uma consulta devolve o conjunto; várias leituras de uma
operação usam uma conexão (`modDB.AbrirLeitura`). Agregação no banco (`GROUP BY`), escrita de
linha inteira como matriz, `ScreenUpdating` desligado em lote. Medir antes de otimizar
(o build registra tempos do teste de uso). Detalhes: `specs/stacks/banco-access.md` e
`specs/stacks/front-vba.md`.
