<#
  verificar-tudo.ps1 - a verificação do projeto (o "/test" e o critério de pronto).

  Níveis (cada um inclui o anterior):
    -Rapido    segundos/1 min: .claude em dia (instalar-skills -Conferir), verificar.py
               (VBA, esquema, .ps1, .bat, JSON), testar-build-lib, roteador (sem rede),
               protocolo do MCP
    (padrão)   + testes dos dois MCPs COM banco, numa cópia isolada de src\BASE
               (criar-banco-teste + testar-mcp; testar-mcp-dev) - não toca no banco
               de demonstração de src\BASE
    -Completo  + build -Teste em src\BASE (Excel): compila, roda o modAutoteste e o
               teste de uso (grades e Painel) com o banco de demonstração; cria esse
               banco se ele não existir. Nada é publicado e a versão não sobe.

  Sai 0 se tudo passou, 1 se algo falhou. Copiar a saída na resposta é a prova de
  "pronto" (definition of done: agent-config\specs\PROJETO.md).

  Uso: powershell -File agent-config\ferramentas\verificar-tudo.ps1 [-Rapido | -Completo]
#>
param([switch]$Rapido, [switch]$Completo)
# Continue: stderr de comando nativo não pode virar exceção (PowerShell 5.1);
# cada etapa é julgada pelo código de saída.
$ErrorActionPreference = 'Continue'

$ac   = Split-Path -Parent $PSScriptRoot
$repo = Split-Path -Parent $ac
$base = Join-Path $repo 'src\BASE'
$ps   = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$resultado = New-Object System.Collections.ArrayList
$inicio = Get-Date

function Etapa([string]$nome, [scriptblock]$acao) {
    Write-Host ''
    Write-Host ('=== ' + $nome) -ForegroundColor Cyan
    $t = Get-Date
    $saida = & $acao 2>&1 | ForEach-Object { "$_" }
    $codigo = $LASTEXITCODE
    # mostra o fim da saída e qualquer linha de falha
    $saida | Where-Object { $_ -match 'FALHA|ERRO|difere|problema|FALHOU|Traceback' } | Select-Object -First 30 | ForEach-Object { Write-Host ('  ' + $_) -ForegroundColor Red }
    $saida | Select-Object -Last 3 | ForEach-Object { Write-Host ('  ' + $_) }
    $ok = ($codigo -eq 0)
    [void]$resultado.Add([pscustomobject]@{ Etapa = $nome; Resultado = $(if ($ok) { 'OK' } else { 'FALHOU' }); Segundos = [int]((Get-Date) - $t).TotalSeconds })
}
function Pular([string]$nome, [string]$motivo) {
    Write-Host ''; Write-Host ('=== ' + $nome + ' - PULADO: ' + $motivo) -ForegroundColor Yellow
    [void]$resultado.Add([pscustomobject]@{ Etapa = $nome; Resultado = 'PULADO'; Segundos = 0 })
}

# ------------------------------------------------------------ rápido
Etapa '.claude em dia (skills + adaptação)' { & $ps -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'instalar-skills.ps1') -Conferir }
$py = Get-Command python -ErrorAction SilentlyContinue
if ($py) { Etapa 'verificar.py (estático)' { & $py.Source (Join-Path $base 'SISTEMA\DADOS\ferramentas\verificacao\verificar.py') } }
else { Pular 'verificar.py (estático)' 'Python não encontrado' }
# caractere de controle no meio do texto (ex.: "\f" de agent-config\ferramentas virando
# avanço de página quando um script de edição interpreta a barra) - já aconteceu
Etapa 'texto sem caractere de controle' {
    $ruins = 0
    foreach ($f in @(& git -C $repo ls-files -co --exclude-standard)) {
        if ($f -notmatch '\.(md|ps1|bat|bas|cls|txt|json|py|js|css|html|g4)$') { continue }
        $p = Join-Path $repo ($f -replace '/', '\')
        if (-not (Test-Path -LiteralPath $p -PathType Leaf)) { continue }
        $bytes = [System.IO.File]::ReadAllBytes($p)
        foreach ($x in $bytes) { if ($x -lt 32 -and $x -ne 9 -and $x -ne 10 -and $x -ne 13) { Write-Output ('FALHA caractere de controle ' + $x + ' em ' + $f); $ruins++; break } }
    }
    if ($ruins -gt 0) { cmd /c exit 1 } else { Write-Output 'nenhum caractere de controle'; cmd /c exit 0 }
}
Etapa 'testar-build-lib' { & $ps -NoProfile -ExecutionPolicy Bypass -File (Join-Path $base 'SISTEMA\DADOS\testes\testar-build-lib.ps1') }
if ($py) { Etapa 'roteador (sem rede)' { & $py.Source (Join-Path $ac 'roteador\testar_roteador.py') } }
else { Pular 'roteador (sem rede)' 'Python não encontrado' }
Etapa 'MCP: protocolo' { & $ps -NoProfile -ExecutionPolicy Bypass -File (Join-Path $base 'IA\teste\testar-mcp.ps1') -SoProtocolo }

# ------------------------------------------------------------ padrão
if (-not $Rapido) {
    # cópia isolada: os testes do MCP criam e alteram BASE\crm_zapromaq.accdb
    $copia = Join-Path ([System.IO.Path]::GetTempPath()) ('crm-verificar-' + $PID)
    $arquivos = @(& git -C $repo ls-files -co --exclude-standard -- src/BASE)
    foreach ($f in $arquivos) {
        $de = Join-Path $repo ($f -replace '/', '\')
        if (-not (Test-Path -LiteralPath $de -PathType Leaf)) { continue }
        $para = Join-Path $copia ($f.Substring(4) -replace '/', '\')
        New-Item -ItemType Directory -Path (Split-Path -Parent $para) -Force | Out-Null
        Copy-Item -LiteralPath $de -Destination $para
    }
    $baseCopia = Join-Path $copia 'BASE'
    Etapa 'MCP de dados com banco de teste' {
        & $ps -NoProfile -ExecutionPolicy Bypass -File (Join-Path $baseCopia 'IA\teste\criar-banco-teste.ps1') -Recriar | Out-Null
        if ($LASTEXITCODE -ne 0) { Write-Output 'ERRO ao criar o banco de teste'; return }
        & $ps -NoProfile -ExecutionPolicy Bypass -File (Join-Path $baseCopia 'IA\teste\testar-mcp.ps1')
    }
    Etapa 'MCP de desenvolvimento' { & $ps -NoProfile -ExecutionPolicy Bypass -File (Join-Path $baseCopia 'IA\teste\testar-mcp-dev.ps1') -SemPrevia }
    Remove-Item -LiteralPath $copia -Recurse -Force -ErrorAction SilentlyContinue
}

# ------------------------------------------------------------ completo
if ($Completo) {
    if ($null -eq [type]::GetTypeFromProgID('Excel.Application')) {
        Pular 'build -Teste (Excel)' 'Excel não instalado nesta máquina'
    } else {
        if (-not (Test-Path -LiteralPath (Join-Path $base 'crm_zapromaq.accdb'))) {
            Etapa 'banco de demonstração' { & $ps -NoProfile -ExecutionPolicy Bypass -File (Join-Path $base 'SISTEMA\DADOS\testes\criar-banco-demo.ps1') }
        }
        Etapa 'build -Teste (compilação, autoteste, teste de uso)' { & $ps -NoProfile -ExecutionPolicy Bypass -File (Join-Path $base 'SISTEMA\DADOS\build\montar-frontend.ps1') -Teste }
    }
}

Write-Host ''
Write-Host '=== RESUMO' -ForegroundColor Cyan
$resultado | Format-Table -AutoSize | Out-String -Width 200 | Write-Host
$falhou = @($resultado | Where-Object { $_.Resultado -eq 'FALHOU' }).Count
$nivel = if ($Completo) { 'completo' } elseif ($Rapido) { 'rápido' } else { 'padrão' }
Write-Host ('Nível ' + $nivel + ', ' + [int]((Get-Date) - $inicio).TotalSeconds + ' s. ' + $(if ($falhou) { $falhou.ToString() + ' etapa(s) FALHARAM.' } else { 'TUDO CERTO.' }))
if ($falhou) { exit 1 } else { exit 0 }
