<#
  testar-conector-versoes.ps1 - DEV-05 (ciclo 007): conector publicado em pasta de versão imutável, ponteiro,
  lançador e volta. Tudo numa cópia temporária de IA (e do mínimo de SISTEMA) com banco de homologação fictício.
#>
$ErrorActionPreference = 'Stop'
$raizIA = Split-Path -Parent $PSScriptRoot
$raizBase = Split-Path -Parent $raizIA
$ps = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('crm-tcv-' + $PID)
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
# Uma chamada ao conector pelo lançador (como o Claude Desktop faz).
function ChamarLancador([string]$base, [string]$ferramenta, $argumentos) {
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $ps
    $psi.Arguments = '-NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "' + (Join-Path $base 'IA\mcp\lancador.ps1') + '"'
    $psi.UseShellExecute = $false; $psi.RedirectStandardInput = $true; $psi.RedirectStandardOutput = $true; $psi.RedirectStandardError = $true
    $psi.StandardOutputEncoding = New-Object System.Text.UTF8Encoding($false); $psi.CreateNoWindow = $true
    $p = [System.Diagnostics.Process]::Start($psi)
    $err = $p.StandardError.ReadToEndAsync()
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    $msg = [ordered]@{ jsonrpc = '2.0'; id = 1; method = 'tools/call'; params = [ordered]@{ name = $ferramenta; arguments = $argumentos } }
    $bt = $utf8.GetBytes((ConvertTo-Json -InputObject $msg -Depth 20 -Compress) + "`n")
    $p.StandardInput.BaseStream.Write($bt, 0, $bt.Length); $p.StandardInput.BaseStream.Flush()
    $linha = $p.StandardOutput.ReadLine()
    $p.StandardInput.Close(); if (-not $p.WaitForExit(8000)) { $p.Kill() }
    [void]$err.Result
    return ($linha | ConvertFrom-Json).result
}

try {
    $b = Join-Path $tmp 'BASE'
    foreach ($rel in @('IA', 'SISTEMA\DADOS\lib', 'SISTEMA\DADOS\banco', 'SISTEMA\DADOS\operacao')) {
        $dest = Join-Path $b $rel
        New-Item -ItemType Directory -Path (Split-Path -Parent $dest) -Force | Out-Null
        Copy-Item -LiteralPath (Join-Path $raizBase $rel) -Destination $dest -Recurse
    }
    Copy-Item -LiteralPath (Join-Path $raizBase 'SISTEMA\DADOS\VERSAO.txt') -Destination (Join-Path $b 'SISTEMA\DADOS\VERSAO.txt')
    foreach ($sobra in @('IA\conector', 'IA\logs')) { $p = Join-Path $b $sobra; if (Test-Path -LiteralPath $p) { Remove-Item -LiteralPath $p -Recurse -Force } }
    $banco = Join-Path $b 'crm_zapromaq.accdb'
    $r = Executar (Join-Path $b 'IA\teste\criar-banco-teste.ps1') @('-Banco', $banco, '-Recriar')
    if ($r.codigo -ne 0) { throw ('não criou o banco: ' + $r.saida) }
    $publicar = Join-Path $b 'IA\mcp\publicar-conector.ps1'
    $versoes = Join-Path $b 'IA\conector\versoes'
    $ponteiro = Join-Path $b 'IA\conector\ATUAL.txt'
    $versao = ([regex]::Match([IO.File]::ReadAllText((Join-Path $b 'IA\mcp\servidor.ps1')), "(?m)^\`$script:VersaoServidor = '([^']+)'")).Groups[1].Value

    Write-Host '--- sem ponteiro: o lançador carrega IA\mcp\servidor.ps1'
    $st = (ChamarLancador $b 'crm_status' ([ordered]@{})).content[0].text | ConvertFrom-Json
    Conferir 'sem publicação, o lançador usa IA\mcp' ([string]$st.pasta_conector -match '\\IA\\mcp$' -and $st.versao_conector -eq $versao) ([string]$st.pasta_conector)

    Write-Host '--- publicação interrompida não deixa versão mista nem troca o ponteiro'
    $env:CRM_FALHA_PUBLICAR_CONECTOR = 'antes-de-renomear'
    try { $r = Executar $publicar @('-SemTeste') } finally { Remove-Item Env:CRM_FALHA_PUBLICAR_CONECTOR -ErrorAction SilentlyContinue }
    Conferir 'falha antes de renomear: sem pasta da versão e sem ponteiro' ($r.codigo -ne 0 -and -not (Test-Path -LiteralPath (Join-Path $versoes $versao)) -and -not (Test-Path -LiteralPath $ponteiro)) $r.saida

    Write-Host '--- publicação: pasta imutável, manifesto, testar-mcp e ponteiro'
    $r = Executar $publicar @('-Motivo', 'primeira publicação')
    Conferir ('publicar cria versoes\' + $versao + ', roda o testar-mcp e aponta o ponteiro') ($r.codigo -eq 0 -and $r.saida -match 'testar-mcp aprovado' -and ([IO.File]::ReadAllText($ponteiro)).Trim() -eq $versao) $r.saida
    Conferir 'restos da promoção interrompida foram removidos' (@(Get-ChildItem -LiteralPath $versoes -Directory | Where-Object { $_.Name -match '\.tmp-' }).Count -eq 0)
    $pastaV = Join-Path $versoes $versao
    Conferir 'a pasta da versão tem servidor, lib, skill e manifesto' ((Test-Path (Join-Path $pastaV 'servidor.ps1')) -and (Test-Path (Join-Path $pastaV 'lib\Ferramentas.ps1')) -and (Test-Path (Join-Path $pastaV 'skill\SKILL.md')) -and (Test-Path (Join-Path $pastaV 'manifesto.json')))
    $st = (ChamarLancador $b 'crm_status' ([ordered]@{})).content[0].text | ConvertFrom-Json
    Conferir 'o lançador carrega a pasta da versão publicada' ([string]$st.pasta_conector -eq $pastaV) ([string]$st.pasta_conector)
    Conferir 'o conector da pasta de versão grava o diagnóstico em IA\logs (não dentro da versão)' (@(Get-ChildItem -LiteralPath (Join-Path $b 'IA\logs') -Filter 'mcp-*.log' -ErrorAction SilentlyContinue).Count -ge 1 -and -not (Test-Path -LiteralPath (Join-Path $pastaV 'logs')))
    $r = Executar $publicar @('-SemTeste')
    Conferir 'publicar de novo o mesmo conteúdo não muda nada' ($r.codigo -eq 0 -and $r.saida -match 'mesmo conteúdo') $r.saida
    Add-Content -LiteralPath (Join-Path $b 'IA\mcp\lib\Regras.ps1') -Value '# mudança sem subir a versão'
    $r = Executar $publicar @('-SemTeste')
    Conferir 'mesma versão com conteúdo diferente é recusada (versão publicada não muda)' ($r.codigo -ne 0 -and $r.saida -match 'outro conteúdo') $r.saida

    Write-Host '--- volta: reapontar, conferindo integridade e esquema'
    $velha = Join-Path $versoes '1.4.9'
    Copy-Item -LiteralPath $pastaV -Destination $velha -Recurse
    $r = Executar $publicar @('-Voltar', '1.4.9', '-Motivo', 'teste de volta')
    Conferir 'volta para versão íntegra e compatível reaponta o ponteiro' ($r.codigo -eq 0 -and ([IO.File]::ReadAllText($ponteiro)).Trim() -eq '1.4.9') $r.saida
    Conferir 'o histórico registra a volta e o motivo' ((Get-Content -LiteralPath (Join-Path $b 'IA\conector\historico.log') -Raw) -match 'VOLTA.*1\.4\.9.*teste de volta')
    $incompat = Join-Path $versoes '1.0.0'
    Copy-Item -LiteralPath $pastaV -Destination $incompat -Recurse
    $bancoInc = Join-Path $incompat 'lib\Banco.ps1'
    [IO.File]::WriteAllText($bancoInc, ([IO.File]::ReadAllText($bancoInc) -replace "(?m)^\`$script:EsquemasSuportados = @\('1\.1'\)", "`$script:EsquemasSuportados = @('1.0')"), (New-Object System.Text.UTF8Encoding($true)))
    $manInc = Get-Content -LiteralPath (Join-Path $incompat 'manifesto.json') -Raw | ConvertFrom-Json
    foreach ($a in $manInc.arquivos) { if ($a.caminho -eq 'lib\Banco.ps1') { $a.sha256 = (Get-FileHash -LiteralPath $bancoInc -Algorithm SHA256).Hash.ToLowerInvariant() } }
    [IO.File]::WriteAllText((Join-Path $incompat 'manifesto.json'), (ConvertTo-Json -InputObject $manInc -Depth 4))
    $r = Executar $publicar @('-Voltar', '1.0.0')
    Conferir 'volta para versão que não suporta o esquema do banco é recusada' ($r.codigo -ne 0 -and $r.saida -match 'incompatível' -and ([IO.File]::ReadAllText($ponteiro)).Trim() -eq '1.4.9') $r.saida
    Add-Content -LiteralPath (Join-Path $velha 'lib\Regras.ps1') -Value '# adulterado'
    $r = Executar $publicar @('-Voltar', '1.4.9')
    Conferir 'volta para versão adulterada é recusada' ($r.codigo -ne 0 -and $r.saida -match 'alterado') $r.saida
    $r = Executar $publicar @('-Voltar', '9.9.9')
    Conferir 'volta para versão não publicada é recusada' ($r.codigo -ne 0 -and $r.saida -match 'não foi publicada') $r.saida
} finally { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue }

Write-Host ''
if ($script:falhas -gt 0) { Write-Host ('RESULTADO: ' + $script:falhas + ' falha(s)'); exit 1 }
Write-Host 'TUDO CERTO.'
exit 0
