<#
  testar-publicacao-copia.ps1 - a publicação de verdade (MONTAR-FRONTEND sem -Teste, passos 1 a 12), mas
  numa CÓPIA de BASE em pasta temporária e com banco de demonstração FICTÍCIO. Prova de ponta a ponta,
  com o Excel de desenvolvimento, do que o -Teste não cobre (spec 004, F7):

    1. publicação ok: .xlsm na raiz, VERSAO.txt e VERSAO-FRONT.txt na versão nova, Início!B7 igual,
       config.versao_front gravada, banco com o esquema 1.1;
    2. segunda publicação: o .xlsm anterior vai para o histórico;
    3. falha forçada no último subpasso (VERSAO-FRONT.txt): o .xlsm, VERSAO.txt, VERSAO-FRONT.txt e o
       BANCO voltam como estavam (hash igual).

  Nada de BASE (a de verdade) é alterado: tudo roda em %TEMP%. Precisa do Excel 2019 e do ACE.
#>
$ErrorActionPreference = 'Stop'
$dados = Split-Path -Parent $PSScriptRoot
$baseOrigem = Split-Path -Parent (Split-Path -Parent $dados)
. (Join-Path $dados 'lib\Comum.ps1')
. (Join-Path $dados 'lib\AtualizacaoBanco.ps1')
$ps = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'

if ($null -eq [type]::GetTypeFromProgID('Excel.Application')) { Write-Host 'PULADO: Excel não instalado.'; exit 0 }

$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('crm-pubcopia-' + $PID)
$base = Join-Path $tmp 'BASE'
$script:falhas = 0
function Conferir([string]$nome, [bool]$ok, [string]$detalhe = '') {
    if ($ok) { Write-Host ('  OK    ' + $nome) -ForegroundColor Green }
    else { $script:falhas++; Write-Host ('  FALHA ' + $nome + '  ' + $detalhe) -ForegroundColor Red }
}
function Hash([string]$p) { if (Test-Path -LiteralPath $p) { return (Get-HashArquivo $p) } return '-' }
function Montar([string]$falhaSub) {
    $pref = $ErrorActionPreference; $ErrorActionPreference = 'Continue'
    try {
        if ($falhaSub) { $env:CRM_FALHA_SUBPASSO_12 = $falhaSub }
        $saida = @(& $ps -NoProfile -ExecutionPolicy Bypass -File (Join-Path $base 'SISTEMA\DADOS\build\montar-frontend.ps1') -TempoLimite 400 2>&1 | ForEach-Object { [string]$_ })
        return [pscustomobject]@{ codigo = $LASTEXITCODE; saida = ($saida -join "`n") }
    } finally { Remove-Item Env:CRM_FALHA_SUBPASSO_12 -ErrorAction SilentlyContinue; $ErrorActionPreference = $pref }
}

try {
    New-Item -ItemType Directory -Path $base -Force | Out-Null
    # cópia de BASE sem bancos, planilhas, travas nem execucao\
    & robocopy.exe $baseOrigem $base /E /XD execucao /XF '*.accdb' '*.xlsm' '*.laccdb' '*.ldb' 'VERSAO-FRONT.txt' /NFL /NDL /NJH /NJS /NP | Out-Null
    if ($LASTEXITCODE -ge 8) { throw 'robocopy falhou ao copiar BASE' }
    $banco = Join-Path $base 'crm_zapromaq.accdb'
    & $ps -NoProfile -ExecutionPolicy Bypass -File (Join-Path $base 'SISTEMA\DADOS\testes\criar-banco-demo.ps1') -Banco $banco -Recriar | Out-Null
    if ($LASTEXITCODE -ne 0) { throw 'não criou o banco de demonstração na cópia' }
    $versao0 = (Get-Content -LiteralPath (Join-Path $base 'SISTEMA\DADOS\VERSAO.txt') -Raw).Trim()
    $xlsm = Join-Path $base 'CRM_Zapromaq.xlsm'; $frontTxt = Join-Path $base 'VERSAO-FRONT.txt'; $versaoTxt = Join-Path $base 'SISTEMA\DADOS\VERSAO.txt'
    $hist = Join-Path $base 'SISTEMA\DADOS\execucao\historico-front'
    $v1 = '{0}.{1}' -f $versao0.Split('.')[0], ([int]$versao0.Split('.')[1] + 1)
    $v2 = '{0}.{1}' -f $versao0.Split('.')[0], ([int]$versao0.Split('.')[1] + 2)

    Write-Host '--- 1. publicação completa numa cópia (Excel)'
    $r = Montar ''
    Conferir ('montar-frontend sem -Teste termina com 0 e publica ' + $v1) ($r.codigo -eq 0 -and $r.saida -match 'PRONTO') ($r.saida.Substring([Math]::Max(0, $r.saida.Length - 1500)))
    Conferir ('VERSAO.txt = ' + $v1) ((Get-Content -LiteralPath $versaoTxt -Raw).Trim() -eq $v1)
    $linhas = @(Get-Content -LiteralPath $frontTxt)
    Conferir ('VERSAO-FRONT.txt: linha 1 = ' + $v1 + ' e linha 2 no padrão de Início!B7') ($linhas.Count -ge 2 -and $linhas[0] -eq $v1 -and $linhas[1] -match '^Ambiente: [A-Z]+   \|   Versao ' + [regex]::Escape($v1) + '   \|   Publicada em \d{2}/\d{2}/\d{4} \d{2}:\d{2}$') ($linhas -join ' / ')
    Conferir '.xlsm publicado na raiz' (Test-Path -LiteralPath $xlsm)
    $cn = Open-BancoLeitura $banco
    try { $vf = [string](Invoke-Escalar $cn "SELECT valor FROM config WHERE chave='versao_front'" @() $null); $esq = Get-VersaoEsquemaBanco $cn } finally { $cn.Close(); $cn.Dispose() }
    Conferir ('config.versao_front = ' + $v1 + ' e esquema 1.1') ($vf -eq $v1 -and $esq -eq '1.1') ("versao_front=$vf esquema=$esq")

    Write-Host '--- 2. segunda publicação: o anterior vai para o histórico'
    $hX1 = Hash $xlsm
    $r = Montar ''
    Conferir ('segunda publicação ' + $v2) ($r.codigo -eq 0 -and (Get-Content -LiteralPath $versaoTxt -Raw).Trim() -eq $v2) ($r.saida.Substring([Math]::Max(0, $r.saida.Length - 1500)))
    Conferir 'o .xlsm anterior está no histórico' (@(Get-ChildItem -LiteralPath $hist -Filter '*.xlsm' -ErrorAction SilentlyContinue | Where-Object { (Get-HashArquivo $_.FullName) -eq $hX1 }).Count -eq 1)

    Write-Host '--- 3. falha forçada no VERSAO-FRONT.txt: tudo volta, inclusive o banco'
    $antes = (Hash $xlsm) + '|' + (Hash $frontTxt) + '|' + (Hash $versaoTxt) + '|' + (Hash $banco)
    $r = Montar 'versao-front'
    $depois = (Hash $xlsm) + '|' + (Hash $frontTxt) + '|' + (Hash $versaoTxt) + '|' + (Hash $banco)
    Conferir 'a montagem falha no passo 12 (e diz que devolveu)' ($r.codigo -ne 0 -and $r.saida -match 'PAROU NO PASSO 12/12' -and $r.saida -match 'voltou ao backup') ($r.saida.Substring([Math]::Max(0, $r.saida.Length - 1200)))
    Conferir '.xlsm, VERSAO-FRONT.txt, VERSAO.txt e banco voltaram como estavam (hash)' ($antes -eq $depois) ($antes + ' != ' + $depois)
} finally { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue }

Write-Host ''
if ($script:falhas -gt 0) { Write-Host ('RESULTADO: ' + $script:falhas + ' falha(s)'); exit 1 }
Write-Host 'TUDO CERTO.'
exit 0
