<#
  testar-compactacao.ps1 - medição da compactação (DEV-25, ciclo 014), com ACE e sem Excel.

  Banco FICTÍCIO numa pasta temporária, inchado com linhas de log inseridas e apagadas (espaço que só a
  compactação recupera). Confere: a medição roda e grava o relatório (tamanho e tempos antes e depois), o
  compactado é menor, a conclusão aplica o critério, o banco de origem fica byte a byte igual, nenhuma cópia
  fica para trás e banco em uso (.laccdb) é recusado.
#>
$ErrorActionPreference = 'Stop'
$dados = Split-Path -Parent $PSScriptRoot
. (Join-Path $dados 'lib\Comum.ps1')

$ps = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$esquema = Join-Path $dados 'banco\esquema\criar-banco.ps1'
$medir = Join-Path $dados 'operacao\scripts\medir-compactacao.ps1'
$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('crm-tcp-' + $PID)
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

try {
    Conferir 'o script de medição existe' (Test-Path -LiteralPath $medir) $medir
    $banco = Join-Path $tmp 'crm_zapromaq.accdb'
    $r = Executar $esquema @('-Banco', $banco, '-Recriar')
    Conferir 'banco fictício criado' ($r.codigo -eq 0) $r.saida
    # inchar: linhas de log com texto longo, inseridas e apagadas numa transação
    $cn = Open-Banco $banco
    try {
        $tx = $cn.BeginTransaction()
        $longo = 'X' * 2000
        for ($i = 1; $i -le 2000; $i++) {
            [void](Invoke-Comando $cn ('INSERT INTO log_alteracoes (tabela, id_registro, campo, valor_antigo, valor_novo, acao, usuario, quando, origem) VALUES (?,?,?,?,?,?,?,?,?)') @(
                (New-Parametro 'texto' 'clientes'), (New-Parametro 'inteiro' $i), (New-Parametro 'texto' 'teste'), (New-Parametro 'memo' $longo), (New-Parametro 'memo' $longo),
                (New-Parametro 'texto' 'ALTERAR'), (New-Parametro 'texto' 'teste'), (New-Parametro 'data' (Get-Date)), (New-Parametro 'texto' 'FRONT')) $tx)
        }
        $tx.Commit()
        [void](Invoke-Comando $cn "DELETE FROM log_alteracoes WHERE campo='teste'" @() $null)
    } finally { $cn.Close(); $cn.Dispose() }
    $hash = (Get-FileHash -LiteralPath $banco -Algorithm SHA256).Hash
    $saida = Join-Path $tmp 'medicoes'
    $r = Executar $medir @('-Banco', $banco, '-Repeticoes', '3', '-Saida', $saida)
    Conferir 'a medição roda numa cópia' ($r.codigo -eq 0) $r.saida
    $pm = @(Get-ChildItem -LiteralPath $saida -Directory -ErrorAction SilentlyContinue) | Select-Object -First 1
    $m = $null
    if ($pm) { $m = Get-Content -LiteralPath (Join-Path $pm.FullName 'medicao.json') -Raw -Encoding UTF8 | ConvertFrom-Json }
    Conferir 'relatório com tamanho e tempos de abrir, buscar, gravar e atualizar, antes e depois' ($null -ne $m -and $m.antes.tamanho_bytes -gt 0 -and $null -ne $m.antes.abrir.mediana_ms -and $null -ne $m.depois.buscar.mediana_ms -and $null -ne $m.antes.gravar.pior_ms -and $null -ne $m.depois.atualizar.pior_ms) ($m | ConvertTo-Json -Compress -Depth 5)
    Conferir 'o compactado é menor (o espaço do log apagado volta)' ($null -ne $m -and [double]$m.depois.tamanho_bytes -lt [double]$m.antes.tamanho_bytes) ''
    $esperado = $false
    if ($m) {
        $esperado = ([double]$m.ganhos.tamanho_pct -ge 20)
        foreach ($k in 'abrir', 'buscar', 'gravar', 'atualizar') {
            if ([double]$m.ganhos.($k + '_mediana_pct') -ge 10 -and ([double]$m.antes.$k.mediana_ms - [double]$m.depois.$k.mediana_ms) -ge 5) { $esperado = $true }
        }
    }
    Conferir 'a conclusão aplica o critério (tamanho -20% ou mediana -10% e -5 ms)' ($null -ne $m -and [bool]$m.com_ganho -eq $esperado -and [string]$m.conclusao -match $(if ($esperado) { '^avaliar rotina' } else { '^não adotar' })) ([string]$m.conclusao)
    Conferir 'o banco de origem ficou byte a byte igual' ((Get-FileHash -LiteralPath $banco -Algorithm SHA256).Hash -eq $hash)
    Conferir 'nenhuma cópia do banco fica para trás (só o relatório)' ($pm -and @(Get-ChildItem -LiteralPath $pm.FullName -Filter '*.accdb').Count -eq 0 -and (Test-Path -LiteralPath (Join-Path $pm.FullName 'medicao.txt')))
    # banco em uso: o .laccdb existe
    $lac = [System.IO.Path]::ChangeExtension($banco, 'laccdb')
    Set-Content -LiteralPath $lac -Value 'EST-TESTE'
    try { $r = Executar $medir @('-Banco', $banco, '-Repeticoes', '1', '-Saida', $saida) } finally { Remove-Item -LiteralPath $lac -Force }
    Conferir 'banco em uso (.laccdb) é recusado' ($r.codigo -eq 2 -and $r.saida -match 'em uso') $r.saida
} finally {
    Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
}
Write-Host ''
if ($script:falhas -gt 0) { Write-Host ('RESULTADO: ' + $script:falhas + ' falha(s)'); exit 1 }
Write-Host 'TUDO CERTO.'
exit 0
