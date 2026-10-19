# Plano: Preparação da rede (ciclo 004)

Spec: [SPEC.md](SPEC.md)

`tpr` = `testar-preparacao-rede.ps1` (ACE); `tbl` = `testar-build-lib.ps1`.

## Ordem e dependências

| Fatia | O quê | Teste que falha primeiro | Dep. | Prova |
|---|---|---|---|---|
| F0 | Ligar `testar-migracao-003.ps1` e o `tpr` ao `verificar-tudo` padrão | `tpr` com o primeiro caso da F1 | — | etapas no RESUMO |
| F1 | `-Recriar` só apaga banco de TESTE (`Assert-PodeSubstituirBanco`) | `tpr`: PRODUCAO + `criar-banco -Recriar` recusa, hash igual | F0 | `-Rapido` e padrão |
| F2 | Testes que gravam recusam banco não TESTE | `tpr`: BASE temporária PRODUCAO + `testar-mcp` recusa | F1 | padrão |
| F3 | Migração estrutural só pela montagem (`-CopiaDeTrabalho`) | `tpr`: PRODUCAO 1.0 recusa; TESTE chega ao 1.1 | F0 | padrão |
| F4 | `Invoke-EnsaioDoBanco` e `montar-frontend -Ensaio` | `tpr`: banco "como a rede" aprovado, produção intacta | F3 | padrão |
| F5 | Passo 10 só leitura + impressão digital | `tbl` e `tpr` (impressão) | F2 | `-Completo` |
| F6 | Troca sem janela (`.novo` pronto antes da última conferência) | `tbl`: gancho de gravação tardia | — | `-Rapido` |
| F7 | Passo 12 recuperável (`Publicar-FrontNaRaiz`, `versao_front` na cópia) | `tbl`: falha forçada por subpasso | F6 | `-Completo` |
| F8 | Backup validado e restauração verificável (`lib\Backup.ps1`) | `tpr`: manifesto, adulteração, LastWriteTime invertido | F4 | padrão |
| F9 | Conector exige 1.1 com mensagem clara | `testar-mcp` com cópia 1.0 | F2 | padrão |
| F10 | Documentação e roteiro | `verificar.py` | F1 a F9 | `-Completo` final |

## Checkpoints

Depois da F3 (`-Rapido` e padrão), da F5 (`-Completo`), da F9 (padrão) e no fim (`-Completo`).

## Itens 🖥

Ver SPEC.md.

## Riscos

Alterar a cadeia de publicação sem Excel de produção: cobrir com fakes em `tbl` e com `-Completo` aqui (Excel 2019 de desenvolvimento). Ordem do passo 12 e do `Replace` só se provam no compartilhamento (🖥).
