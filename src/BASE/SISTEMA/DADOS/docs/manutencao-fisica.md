# Manutenção física do banco (DEV-25, ciclo 014)

O `crm_zapromaq.accdb` cresce também com espaço que o Access não devolve sozinho (linhas apagadas e alteradas).
Compactar recupera esse espaço, mas exige o banco **exclusivo** e a troca do arquivo: só vale a pena com ganho medido.

## Medir (numa cópia, sem tocar o banco)

```
SISTEMA\DADOS\operacao\MEDIR-COMPACTACAO.bat -Banco <última cópia de backup validada (execucao\backup\...)>
```

- Copia o arquivo para `execucao\medicoes\compactacao-<instante>\`, mede tamanho e os tempos de **abrir**,
  **buscar** (consulta da grade de atendimentos), **gravar** (transação desfeita) e **atualizar** (rodada do feed),
  em N repetições (padrão 7, a primeira de aquecimento é descartada) - mediana e pior caso -, compacta a cópia
  (`DAO.DBEngine.120.CompactDatabase`, motor do próprio Office), mede de novo e grava `medicao.json` e `medicao.txt`.
- As cópias são apagadas no fim; o banco de origem é conferido por SHA-256 (não pode mudar).
- Banco em uso (`.laccdb`) é recusado - por isso a medição se faz sobre a cópia de backup validada.
- Teste automatizado: `testes\testar-compactacao.ps1` (banco fictício inchado), no `verificar-tudo` padrão.

**Critério:** compactar reduz o tamanho em pelo menos 20% **ou** a mediana de algum tempo em pelo menos 10% e 5 ms.

- Sem ganho → **não adotar** (conclusão registrada no `PLANO.md`).
- Com ganho → **avaliar a rotina**: compactação em janela de manutenção (`config.manutencao`, DEV-08), com backup
  completo validado antes, compactação sobre uma cópia, conferência (esquema, contagens, geração da base renovada) e
  troca controlada do arquivo, como na montagem. A rotina não existe ainda: só depois da medição no banco real.

Medição de referência no banco de demonstração (fictício, 1,5 MB, 23/10/2026): tamanho 9% menor e tempos de poucos
milissegundos sem diferença acima do ruído - **não adotar**. Ela não representa a produção: medir na rede (🖥).

## Reparo em caso de corrupção

1. Todos fora do CRM; ligar a manutenção (o próximo `MONTAR-FRONTEND.bat` ou `APLICAR-MIGRACOES.bat` liga e confere).
2. **Não** reparar o arquivo da produção no lugar. Primeiro, copiar o arquivo danificado para uma pasta de trabalho.
3. Caminho preferido: restaurar a **última cópia completa validada** (`operacao\TESTAR-RESTAURACAO.bat` confere a
   cópia numa pasta descartável; ver `implantacao.md`) e, com o banco restaurado no lugar, rodar
   `operacao\scripts\renovar-geracao.ps1 -Confirmo` (as estações recarregam tudo).
4. Só se não houver cópia validada recente: tentar compactar/reparar a **cópia** do arquivo danificado
   (`CompactDatabase` repara o que consegue), conferir esquema, vínculos e contagens contra o último manifesto de
   backup e, só então, trocar - registrando o que pode ter sido perdido desde a última cópia.
5. Registrar o ocorrido (instante, causa provável, cópia usada, perda potencial) no diário de operação.
