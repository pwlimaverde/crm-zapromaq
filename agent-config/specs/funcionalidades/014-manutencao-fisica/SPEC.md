# Spec: Manutenção física - medir a compactação (ciclo 014)

> Situação: aprovada em 22/10/2026 (plano revisado de 19/10/2026: DEV-25 - "medir compactação numa cópia; adotar só
> com ganho"). Sem migração; nunca toca o banco de origem.

## Objetivo

Saber quanto do tamanho do banco é espaço não recuperado e se compactar melhora o uso, **medindo numa cópia**:

- `operacao\scripts\medir-compactacao.ps1` (+ `operacao\MEDIR-COMPACTACAO.bat`): copia o banco (ou uma cópia de backup
  validada indicada em `-Banco`) para `execucao\medicoes\<instante>\`, mede na cópia tamanho e tempos de **abrir**,
  **buscar** (consulta da grade de atendimentos com junções e `LIKE`), **gravar** (transação desfeita no fim) e
  **atualizar** (consulta do feed: `MAX(id)` + faixa do log) — N repetições, mediana e pior caso —, compacta a cópia
  pelo motor do Office (`DAO.DBEngine.120.CompactDatabase`), mede de novo e grava `medicao.json` + `medicao.txt` com a
  conclusão. As cópias são apagadas no fim (podem conter dado real); fica só o relatório.
- **Critério de ganho** (decisão 014-1): compactar reduz o tamanho em ≥ 20% **ou** a mediana de algum tempo em ≥ 10% e ≥ 5 ms (abaixo disso é ruído).
  Sem ganho: "não adotar" registrado. Com ganho: "avaliar rotina" - a rotina em janela de manutenção (DEV-08 +
  backup validado + troca controlada) **não** é implantada neste ciclo sem a medição no banco real (🖥 homologação).
- Procedimento de reparo em caso de corrupção (restaurar a última cópia validada; compactar/reparar só numa cópia) em
  `SISTEMA\DADOS\docs\manutencao-fisica.md`.

## Critérios de aceite

- [ ] `testes\testar-compactacao.ps1`: banco fictício inchado (inserções e exclusões no log) - a medição roda, o
  relatório tem tamanho e tempos antes/depois, o compactado é menor, a conclusão aplica o critério, o banco de origem
  não muda (hash igual) e nenhuma cópia fica para trás; banco em uso (`.laccdb`) é recusado sem `-Mesmo`.
- [ ] Entra no `verificar-tudo` (nível padrão).
- 🖥 Rodar `MEDIR-COMPACTACAO.bat` na rede sobre a última cópia validada (homologação) e registrar a conclusão no PLANO.

## Decisões assumidas

| # | Decisão | Premissa | O que mudaria |
|---|---|---|---|
| 014-1 | Ganho = tamanho −20% ou mediana −10% **e** −5 ms em algum tempo; a 1ª rodada de cada arquivo (aquecimento) é descartada. | Abaixo disso a rotina (janela exclusiva, troca) custa mais do que rende. | A gestão definir outro limiar. |
| 014-2 | Rotina de compactação não implantada sem a medição no banco real. | O banco fictício não representa a fragmentação da produção. | Medição na rede mostrar ganho: ciclo próprio para a rotina. |
| 014-3 | Medição só em cópia, com o banco de origem lido sem bloqueio; banco com `.laccdb` recusado (a cópia poderia sair inconsistente), salvo `-Mesmo`. | DEV-03: tamanho igual não prova consistência. | — |

## Limites

- Nunca: compactar, reparar ou trocar o banco de origem; deixar cópia com dado real fora do relatório.
