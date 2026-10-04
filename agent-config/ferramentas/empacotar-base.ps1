<#
  empacotar-base.ps1 - gera o .zip da pasta BASE para levar à empresa.

  O pacote sai do COMMIT atual (git archive de src/BASE), não da pasta de
  trabalho: só entra o que está versionado. Por isso nunca leva o banco, a
  planilha, VERSAO-FRONT.txt, logs, backups, .bak-* nem o banco de demonstração.
  As skills do Claude Desktop (IA\instrucoes\<nome>.zip) são geradas aqui,
  a partir dos SKILL.md.

  Na empresa: extrair o .zip POR CIMA da pasta BASE existente (substituir os
  arquivos) e rodar SISTEMA\MONTAR-FRONTEND.bat. O que é só da rede - o
  .accdb, o .xlsm publicado, VERSAO-FRONT.txt, crm_zapromaq.ico, a pasta
  SISTEMA\DADOS\execucao (backups, histórico, logs) e IA\logs - não está no
  pacote e continua lá. Roteiro: SISTEMA\DADOS\docs\levar-para-a-empresa.md.

  Uso: powershell -File agent-config\ferramentas\empacotar-base.ps1 [-Saida <pasta>]
  Padrão de saída: dist\ na raiz do repositório.
#>
param([string]$Saida = '')
# Continue: no PowerShell 5.1, stderr de comando nativo vira erro fatal com Stop.
# Os erros reais saem por throw, que sempre interrompe.
$ErrorActionPreference = 'Continue'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function G {
    $saida = & git @args 2>&1
    if ($LASTEXITCODE -ne 0) { throw ('git ' + ($args -join ' ') + "`n" + ($saida | Out-String)) }
    return $saida
}

$raiz = ((G rev-parse --show-toplevel) | Select-Object -First 1).Trim() -replace '/', '\'
Set-Location $raiz
if (-not (Test-Path -LiteralPath (Join-Path $raiz 'src\BASE\SISTEMA\MONTAR-FRONTEND.bat'))) { throw 'Não achei src\BASE neste repositório.' }

# o pacote é do commit: alteração não commitada ficaria de fora sem ninguém perceber
$pendentes = @(G status --porcelain --untracked-files=all -- src/BASE)
if ($pendentes.Count -gt 0) {
    $pendentes | Select-Object -First 15 | ForEach-Object { Write-Host ('  ' + $_) }
    throw 'Há alterações não commitadas em src\BASE (lista acima). Faça o commit antes de empacotar: o pacote sai do commit.'
}

$commit = ((G rev-parse --short HEAD) | Select-Object -First 1).Trim()
$ramo = ((G rev-parse --abbrev-ref HEAD) | Select-Object -First 1).Trim()
$versao = ((G show HEAD:src/BASE/SISTEMA/DADOS/VERSAO.txt) | Select-Object -First 1).Trim()
if ($versao -notmatch '^\d+\.\d+$') { throw ('VERSAO.txt inválido no commit: ' + $versao) }

if ([string]::IsNullOrWhiteSpace($Saida)) { $Saida = Join-Path $raiz 'dist' }
New-Item -ItemType Directory -Path $Saida -Force | Out-Null
$nome = 'BASE-v' + $versao + '-' + (Get-Date -Format 'yyyyMMdd-HHmm') + '-' + $commit + '.zip'
$zip = Join-Path $Saida $nome
if (Test-Path -LiteralPath $zip) { Remove-Item -LiteralPath $zip -Force }

G archive --format=zip --prefix=BASE/ -o $zip HEAD:src/BASE | Out-Null

# skills do Claude Desktop: um .zip por pasta de IA\instrucoes, com <nome>\SKILL.md dentro
$arq = [System.IO.Compression.ZipFile]::Open($zip, 'Update')
try {
    $skills = @($arq.Entries | Where-Object { $_.FullName -match '^BASE/IA/instrucoes/([^/]+)/SKILL\.md$' })
    foreach ($s in $skills) {
        $skill = ($s.FullName -split '/')[3]
        $mem = New-Object System.IO.MemoryStream
        $interno = New-Object System.IO.Compression.ZipArchive($mem, 'Create', $true)
        $ent = $interno.CreateEntry($skill + '/SKILL.md')
        $de = $s.Open(); $para = $ent.Open(); $de.CopyTo($para); $para.Dispose(); $de.Dispose()
        $interno.Dispose()
        $novo = $arq.CreateEntry('BASE/IA/instrucoes/' + $skill + '.zip')
        $para = $novo.Open(); $mem.Position = 0; $mem.CopyTo($para); $para.Dispose(); $mem.Dispose()
    }
    # conferência: nada que seja da rede ou dado pode estar no pacote
    $proibidos = @($arq.Entries | Where-Object {
        $_.FullName -match '\.(accdb|laccdb|mdb|ldb|xlsm|xlsx|xls|log|pdf|bmp)$' -or
        $_.FullName -match '(^|/)VERSAO-FRONT\.txt$' -or $_.FullName -match '\.bak-' -or
        ($_.FullName -match '/execucao/' -and $_.FullName -notmatch '/\.gitkeep$' -and $_.FullName -notmatch '/$') -or
        $_.FullName -match '^BASE/IA/logs/' })
    $total = @($arq.Entries | Where-Object { $_.FullName -notmatch '/$' }).Count
    $temIniciar = @($arq.Entries | Where-Object { $_.FullName -eq 'BASE/INICIAR-CRM.bat' }).Count -eq 1
} finally { $arq.Dispose() }

if ($proibidos.Count -gt 0) {
    Remove-Item -LiteralPath $zip -Force
    throw ('O pacote teria arquivos que não podem ir para a rede: ' + (($proibidos | ForEach-Object { $_.FullName }) -join ', '))
}
if (-not $temIniciar) { throw 'INICIAR-CRM.bat não entrou no pacote.' }

Write-Host ''
Write-Host ('Pacote: ' + $zip)
Write-Host ('  versão no VERSAO.txt: ' + $versao + '   commit ' + $commit + ' (' + $ramo + ')   ' + $total + ' arquivos, ' + $skills.Count + ' skill(s) .zip')
Write-Host ('  ' + [math]::Round((Get-Item -LiteralPath $zip).Length / 1KB) + ' KB')
Write-Host ''
Write-Host 'Na empresa: extraia POR CIMA da pasta BASE da rede (substituir os arquivos),'
Write-Host 'NUNCA apague a BASE antes - o banco e o .xlsm publicados estão nela.'
Write-Host 'Depois: SISTEMA\MONTAR-FRONTEND.bat. Roteiro: SISTEMA\DADOS\docs\levar-para-a-empresa.md'
if ($ramo -notin @('main', 'develop') -and $ramo -notlike 'release/*' -and $ramo -notlike 'hotfix/*') {
    Write-Host ''
    Write-Host ('AVISO: empacotado a partir de ' + $ramo + ', não de main/develop/release.') -ForegroundColor Yellow
}
