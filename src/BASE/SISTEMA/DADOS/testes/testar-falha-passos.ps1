# testar-falha-passos.ps1 - prova da fatia 1.7b.
#
# Para cada passo 9 a 12 da montagem, força a falha (CRM_FALHA_NO_PASSO, so vale com -Teste) e confere:
#   1. a montagem para no passo certo ("PAROU NO PASSO n/12");
#   2. o banco de produção (aqui, o banco de teste da pasta) fica com o mesmo hash de antes.
# O passo 12 falha depois da troca do banco: o teste prova a volta ao backup.
# Roda só com Excel 2019 (a montagem compila o front). Não publica: -Teste grava em execucao\teste.

$ErrorActionPreference = 'Stop'
$base = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
$banco = Join-Path $base 'crm_zapromaq.accdb'
$montar = Join-Path $base 'SISTEMA\DADOS\build\montar-frontend.ps1'
if (-not (Test-Path -LiteralPath $banco)) { Write-Host "FALHA: banco de teste nao encontrado em $banco"; exit 1 }

$falhas = 0
foreach ($n in 9, 10, 11, 12) {
    $antes = (Get-FileHash -LiteralPath $banco -Algorithm SHA256).Hash
    $env:CRM_FALHA_NO_PASSO = [string]$n
    try {
        $saida = (& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $montar -Teste -SemTesteDeUso 2>&1 | Out-String)
    } finally {
        Remove-Item Env:CRM_FALHA_NO_PASSO -ErrorAction SilentlyContinue
    }
    $depois = (Get-FileHash -LiteralPath $banco -Algorithm SHA256).Hash
    $parou = ($saida -match ("PAROU NO PASSO " + $n + "/12"))
    $mesmo = ($antes -eq $depois)
    if ($parou -and $mesmo) {
        Write-Host ("  OK    falha forcada no passo " + $n + "/12: parou no passo certo e o banco ficou igual")
    } else {
        $falhas++
        Write-Host ("  FALHA falha forcada no passo " + $n + "/12: parou=" + $parou + " banco igual=" + $mesmo)
        Write-Host $saida
    }
}
if ($falhas -gt 0) { Write-Host ("RESULTADO: " + $falhas + " falha(s)"); exit 1 }
Write-Host 'TUDO CERTO.'
exit 0
