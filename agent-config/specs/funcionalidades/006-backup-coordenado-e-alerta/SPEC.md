# Spec: Backup coordenado e alerta (ciclo 006)

> Concluída em 18/10/2026 - feature/backup-coordenado-e-alerta

## Objetivo

DEV-02 (backup sem coordenação), DEV-21 (ninguém vê o estado do backup) e a geração da base de DEV-13.
Depende do ciclo 004 (F8: manifesto e validação).

## Critérios de aceite

- [x] Trava de backup com dono e sessão (`backup.lock`, criação atômica): a segunda execução sai com 4 citando o dono;
  trava abandonada (PID morto na mesma estação, ou mais de 60 min em qualquer estação) é recuperada. Prova: `testar-preparacao-rede` (processo filho real).
- [x] Nome único `crm_<tipo>_AAAAMMDD-HHMMSS_<ESTACAO>.accdb` (sufixo `-2`, `-3` se repetir); a cópia nasce `.tmp` e só é
  promovida depois de conferida; log por execução (`execucao\backup\logs`) e o consolidado `backup.log`. Prova: `testar-preparacao-rede`.
- [x] Manifesto em toda cópia e marca `.validado` só depois da restauração numa pasta descartável; a retenção nunca apaga o
  último validado e leva junto manifesto e marca de quem apaga. Prova: `testar-preparacao-rede`.
- [x] `ver_log backup` lê o `backup.log`; `estado_sistema` (dev) e `crm_status` mostram o último backup validado e o alerta
  (> 36 h). Prova: `testar-mcp-dev`, `testar-preparacao-rede`.
- [x] Alerta no Início: ao abrir, o front avisa se não há backup validado há mais de 36 h (`modBackupAviso`). Prova: 5 `Confere` no `modAutoteste` (165 testes).
- [x] `config.geracao_base`: nasce com o banco, é renovada a cada publicação (cópia, antes da troca) e por `renovar-geracao.ps1`
  depois de restaurar um backup sobre a produção (recusa banco em uso). Prova: `testar-preparacao-rede`.

## Decisões assumidas novas

- Marca de validação por arquivo ao lado (`<backup>.validado`), legível por qualquer consumidor (front, conector, dev) sem JSON complexo.
- Cópia simples com o CRM aberto pode não validar: fica sem marca e o log avisa (não é erro).
- Trava abandonada: 60 min; recuperada com aviso no log.
- Backups antigos (formato `HHMM`, sem manifesto) continuam válidos para a retenção e para restaurar, mas nunca contam como validados.
- Front: o aviso é um MsgBox na abertura, só quando atrasado; pasta ilegível não gera aviso.

## Itens 🖥

- Agendar (`AGENDAR-BACKUPS.bat`), rodar `FAZER-BACKUP.bat` e conferir `.manifesto.json` e `.validado` ao lado do `.accdb`.
- Abrir o front com o último validado velho (renomear a marca de teste na pasta de homologação) e ver o aviso.
- Duas estações rodando `FAZER-BACKUP.bat` ao mesmo tempo: a segunda recusa (código 4).

## Limites

Nunca abrir o arquivo de backup; nunca apagar o último validado.
