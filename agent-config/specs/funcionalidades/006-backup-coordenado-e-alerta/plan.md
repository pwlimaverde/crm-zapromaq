# Plano: Backup coordenado e alerta (ciclo 006)

Spec: [SPEC.md](SPEC.md)

1. Trava de backup (`Enter-/Exit-TravaBackup`) - `testar-preparacao-rede` com processo filho.
2. `backup-banco.ps1`: nome único, `.tmp`, manifesto, validado, logs - `testar-preparacao-rede`.
3. Retenção protege o último validado - `testar-preparacao-rede`.
4. Estado do backup no conector e no dev - `testar-preparacao-rede`, `testar-mcp-dev`.
5. Alerta no front (`modBackupAviso`) - `modAutoteste`.
6. `geracao_base` e `renovar-geracao.ps1` - `testar-preparacao-rede`.
