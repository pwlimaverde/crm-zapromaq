<#
  testar-preparacao-rede.ps1 - testes da preparação da rede (spec 004), com ACE e sem Excel.

  Tudo acontece em pasta temporária e com bancos FICTÍCIOS: nada de BASE é tocado.
  Cada caso que recusa uma operação confere também que o banco ficou byte a byte igual (SHA-256).
#>
$ErrorActionPreference = 'Stop'
$dados = Split-Path -Parent $PSScriptRoot
$raizBase = Split-Path -Parent (Split-Path -Parent $dados)
. (Join-Path $dados 'lib\Comum.ps1')
. (Join-Path $dados 'lib\AtualizacaoBanco.ps1')

$ps = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$esquema = Join-Path $dados 'banco\esquema\criar-banco.ps1'
$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('crm-tpr-' + $PID)
New-Item -ItemType Directory -Path $tmp -Force | Out-Null

$script:falhas = 0
function Conferir([string]$nome, [bool]$ok, [string]$detalhe = '') {
    if ($ok) { Write-Host ('  OK    ' + $nome) -ForegroundColor Green }
    else { $script:falhas++; Write-Host ('  FALHA ' + $nome + '  ' + $detalhe) -ForegroundColor Red }
}
# Roda um script em outro PowerShell 5.1 e devolve código e saída (stderr incluído).
function Executar([string]$script, [string[]]$argumentos) {
    $pref = $ErrorActionPreference; $ErrorActionPreference = 'Continue'
    try {
        $saida = @(& $ps -NoProfile -ExecutionPolicy Bypass -File $script @argumentos 2>&1 | ForEach-Object { [string]$_ })
        return [pscustomobject]@{ codigo = $LASTEXITCODE; saida = ($saida -join "`n") }
    } finally { $ErrorActionPreference = $pref }
}
function NovoBanco([string]$caminho, [string]$ambiente) {
    $r = Executar $esquema @('-Banco', $caminho, '-Ambiente', $ambiente)
    if ($r.codigo -ne 0) { throw ('não criou o banco ' + $caminho + ': ' + $r.saida) }
}

try {
    Write-Host '--- F1: -Recriar só apaga banco de TESTE'
    $prod = Join-Path $tmp 'producao.accdb'
    NovoBanco $prod 'PRODUCAO'
    $h = Get-HashArquivo $prod
    $r = Executar $esquema @('-Banco', $prod, '-Recriar')
    Conferir 'criar-banco -Recriar recusa banco PRODUCAO' ($r.codigo -ne 0 -and $r.saida -match 'PRODUCAO') $r.saida
    Conferir 'banco PRODUCAO ficou igual (hash)' ((Get-HashArquivo $prod) -eq $h)

    $lixo = Join-Path $tmp 'lixo.accdb'
    [IO.File]::WriteAllText($lixo, 'isto não é um banco')
    $h2 = Get-HashArquivo $lixo
    $r = Executar $esquema @('-Banco', $lixo, '-Recriar')
    Conferir 'arquivo que não abre é recusado' ($r.codigo -ne 0) $r.saida
    Conferir 'arquivo ilegível ficou igual (hash)' ((Get-HashArquivo $lixo) -eq $h2)

    $teste = Join-Path $tmp 'teste.accdb'
    NovoBanco $teste 'TESTE'
    $r = Executar $esquema @('-Banco', $teste, '-Recriar')
    Conferir 'banco TESTE é recriado' ($r.codigo -eq 0) $r.saida

    $novo = Join-Path $tmp 'inexistente.accdb'
    $r = Executar $esquema @('-Banco', $novo, '-Recriar')
    Conferir 'arquivo inexistente é criado' ($r.codigo -eq 0 -and (Test-Path -LiteralPath $novo)) $r.saida

    $msg = Assert-PodeSubstituirBanco $prod
    Conferir 'Assert-PodeSubstituirBanco devolve o motivo para PRODUCAO' ($msg -match 'PRODUCAO')
    Conferir 'Assert-PodeSubstituirBanco aceita TESTE' ((Assert-PodeSubstituirBanco $teste) -eq '')
    Conferir 'Assert-PodeSubstituirBanco aceita inexistente' ((Assert-PodeSubstituirBanco (Join-Path $tmp 'nao-existe.accdb')) -eq '')

    # FIM-DOS-CASOS
} finally { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue }

Write-Host ''
if ($script:falhas -gt 0) { Write-Host ('RESULTADO: ' + $script:falhas + ' falha(s)'); exit 1 }
Write-Host 'TUDO CERTO.'
exit 0
