<#
  testar-recuperacao.ps1 - testes da Etapa 1 do plano de 19/10/2026 (ciclo 007), com ACE e sem Excel:
    DEV-08 sinal de manutenção (montagem e migrações), DEV-03 exclusividade durante a cópia, janela
    atendida e aviso por janela, DEV-06A diagnóstico e retenção.

  Tudo em pasta temporária e com bancos FICTÍCIOS (criar-banco-teste.ps1): nada de BASE é tocado.
#>
$ErrorActionPreference = 'Stop'
$dados = Split-Path -Parent $PSScriptRoot
$raizBase = Split-Path -Parent (Split-Path -Parent $dados)
. (Join-Path $dados 'lib\Comum.ps1')
. (Join-Path $dados 'lib\AtualizacaoBanco.ps1')

$ps = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('crm-trc-' + $PID)
New-Item -ItemType Directory -Path $tmp -Force | Out-Null

$script:falhas = 0
function Conferir([string]$nome, [bool]$ok, [string]$detalhe = '') {
    if ($ok) { Write-Host ('  OK    ' + $nome) -ForegroundColor Green }
    else { $script:falhas++; Write-Host ('  FALHA ' + $nome + '  ' + $detalhe) -ForegroundColor Red }
}
function Executar([string]$script, [string[]]$argumentos) {
    $pref = $ErrorActionPreference; $ErrorActionPreference = 'Continue'
    try {
        $saida = @(& $ps -NoProfile -ExecutionPolicy Bypass -File $script @argumentos 2>&1 | ForEach-Object { [string]$_ })
        return [pscustomobject]@{ codigo = $LASTEXITCODE; saida = ($saida -join "`n") }
    } finally { $ErrorActionPreference = $pref }
}
# BASE mínima em pasta temporária, com banco de homologação fictício na raiz.
function NovaBase([string]$nome) {
    $b = Join-Path $tmp $nome
    foreach ($rel in @('IA', 'SISTEMA\DADOS\lib', 'SISTEMA\DADOS\banco', 'SISTEMA\DADOS\build', 'SISTEMA\DADOS\testes', 'SISTEMA\DADOS\operacao')) {
        $dest = Join-Path $b $rel
        New-Item -ItemType Directory -Path (Split-Path -Parent $dest) -Force | Out-Null
        Copy-Item -LiteralPath (Join-Path $raizBase $rel) -Destination $dest -Recurse
    }
    Copy-Item -LiteralPath (Join-Path $dados 'VERSAO.txt') -Destination (Join-Path $b 'SISTEMA\DADOS\VERSAO.txt')
    $r = Executar (Join-Path $b 'IA\teste\criar-banco-teste.ps1') @('-Banco', (Join-Path $b 'crm_zapromaq.accdb'), '-Recriar')
    if ($r.codigo -ne 0) { throw ('não criou o banco de teste: ' + $r.saida) }
    return $b
}
function ValorConfig([string]$banco, [string]$chave) {
    $cn = Open-BancoLeitura $banco
    try { return [string](Invoke-Escalar $cn 'SELECT valor FROM config WHERE chave=?' @((New-Parametro 'texto' $chave)) $null) } finally { $cn.Close(); $cn.Dispose() }
}

try {
    Write-Host '--- DEV-08: sinal de manutenção (config.manutencao)'
    $b1 = NovaBase 'base-manutencao'
    $banco1 = Join-Path $b1 'crm_zapromaq.accdb'
    $cn = Open-Banco $banco1
    try {
        $s1 = Set-Manutencao $cn 'teste de manutenção'
        Conferir 'Set-Manutencao grava estação|usuário|PID|início|motivo' ((Get-Manutencao $cn) -eq $s1 -and $s1.Split('|').Count -eq 5 -and $s1.StartsWith($env:COMPUTERNAME + '|')) $s1
        Conferir 'Clear-Manutencao com outro dono não desliga' (-not (Clear-Manutencao $cn 'OUTRA|x|1|y|z') -and (Get-Manutencao $cn) -eq $s1)
        Conferir 'Clear-Manutencao do próprio dono desliga' ((Clear-Manutencao $cn $s1) -and (Get-Manutencao $cn) -eq '')
        [void](Set-Manutencao $cn 'dono que caiu')
    } finally { $cn.Close(); $cn.Dispose() }
    $r = Executar (Join-Path $b1 'SISTEMA\DADOS\banco\aplicar-migracoes.ps1') @('-Banco', $banco1, '-LiberarManutencao')
    Conferir 'aplicar-migracoes -LiberarManutencao desliga o sinal deixado por dono que caiu' ($r.codigo -eq 0 -and (ValorConfig $banco1 'manutencao') -eq '') $r.saida

    # FIM-DOS-CASOS
} finally { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue }

Write-Host ''
if ($script:falhas -gt 0) { Write-Host ('RESULTADO: ' + $script:falhas + ' falha(s)'); exit 1 }
Write-Host 'TUDO CERTO.'
exit 0
