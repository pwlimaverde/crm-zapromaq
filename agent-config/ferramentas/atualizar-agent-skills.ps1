<#
  atualizar-agent-skills.ps1 - traz a versão nova do addyosmani/agent-skills e reinstala.

    1. git subtree pull em agent-config\vendor\agent-skills (o original, nunca editado aqui)
    2. mostra o que mudou no original (skills, comandos, agentes, checklists)
    3. instalar-skills.ps1: regera agent-config\.claude com a adaptação por cima

  Se uma skill ou comando adaptado mudou de nome ou sumiu, o passo 3 para e diz
  qual arquivo de adaptacao\ ficou órfão. Releia o que mudou nas skills adaptadas
  (lista do passo 2) e ajuste adaptacao\ se o original passou a dizer outra coisa.

  Rode numa branch (git flow feature start atualizar-agent-skills), com a árvore limpa.
  Uso: atualizar-agent-skills.ps1 [-Ramo main]
#>
param([string]$Ramo = 'main')
$ErrorActionPreference = 'Continue'
$origem = 'https://github.com/addyosmani/agent-skills.git'
$prefixo = 'agent-config/vendor/agent-skills'

function G {
    $saida = & git @args 2>&1
    if ($LASTEXITCODE -ne 0) { throw ('git ' + ($args -join ' ') + "`n" + ($saida | Out-String)) }
    return $saida
}

$repo = ((G rev-parse --show-toplevel) | Select-Object -First 1).Trim()
Set-Location $repo
if (@(G status --porcelain).Count -gt 0) { throw 'Há alterações não commitadas. O subtree pull precisa da árvore limpa.' }
$ramoAtual = ((G rev-parse --abbrev-ref HEAD) | Select-Object -First 1).Trim()
if ($ramoAtual -in @('main', 'develop')) { throw ('Você está em ' + $ramoAtual + '. Abra uma branch antes: git flow feature start atualizar-agent-skills') }

$antes = ((G rev-parse HEAD) | Select-Object -First 1).Trim()
Write-Host ('Trazendo ' + $origem + ' (' + $Ramo + ')...')
G subtree pull --prefix=$prefixo $origem $Ramo --squash -m ('Atualiza addyosmani/agent-skills (' + $Ramo + ')') | Out-Null
$depois = ((G rev-parse HEAD) | Select-Object -First 1).Trim()
if ($antes -eq $depois) { Write-Host 'Nada novo no original.'; exit 0 }

Write-Host ''
Write-Host 'O que mudou no original:'
G diff --stat $antes $depois -- ($prefixo + '/skills') ($prefixo + '/.claude/commands') ($prefixo + '/agents') ($prefixo + '/references') | ForEach-Object { Write-Host ('  ' + $_) }
$adaptados = @(Get-ChildItem (Join-Path $repo 'agent-config\adaptacao') -Recurse -Filter '*.md' | ForEach-Object { $_.BaseName })
$tocados = @(G diff --name-only $antes $depois -- $prefixo | Where-Object { $n = ($_ -split '/')[-2..-1]; $adaptados -contains ([IO.Path]::GetFileNameWithoutExtension($_)) -or $adaptados -contains $n[0] })
if ($tocados.Count -gt 0) {
    Write-Host ''
    Write-Host 'ATENÇÃO - mudaram itens que têm adaptação (releia e ajuste adaptacao\ se preciso):' -ForegroundColor Yellow
    $tocados | ForEach-Object { Write-Host ('  ' + $_) -ForegroundColor Yellow }
}

Write-Host ''
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'instalar-skills.ps1')
if ($LASTEXITCODE -ne 0) { throw 'instalar-skills falhou (veja acima). O subtree já foi atualizado; corrija adaptacao\ e rode instalar-skills.ps1.' }
Write-Host ''
Write-Host 'Próximo passo: revisar git status/git diff de agent-config\.claude e fazer o commit.'
