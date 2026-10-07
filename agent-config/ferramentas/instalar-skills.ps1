<#
  instalar-skills.ps1 - monta agent-config\.claude (o que o Claude Code lê) a partir de:

    vendor\agent-skills\   o projeto addyosmani/agent-skills NA ÍNTEGRA (git subtree,
                           nunca editado aqui; atualizar com atualizar-agent-skills.ps1)
    adaptacao\             o que é deste projeto, somado ao original:
      skills\<nome>.md     acrescentado ao fim de .claude\skills\<nome>\SKILL.md
      commands\<nome>.md   acrescentado ao fim de .claude\commands\<nome>.md
      commands\_todos.md   acrescentado ao fim de TODO comando
      agents\<nome>.md     acrescentado ao fim de .claude\agents\<nome>.md
      excluidos.txt        itens do original que não se aplicam (um por linha: skills\x, commands\x, agents\x)
      proprios\{skills,commands,agents}\   itens só deste projeto, copiados como estão (ex.: /roteador)

  Gera .claude\skills, .claude\commands, .claude\agents e .claude\references (os
  checklists que as skills citam como ../../references). O resto de .claude
  (settings*.json) não é tocado. No texto, "agent-skills:<skill>" (nome do plugin)
  vira "<skill>": aqui as skills são do projeto, não de um plugin.
  Também confere e assegura o espelhamento do Antigravity (.agents\skills.json
  e GEMINI.md), mantendo fonte única por referência, sem duplicatas.

  Gerado, não editar: para mudar, edite adaptacao\ e rode de novo.

  Uso: instalar-skills.ps1 [-Conferir]
    -Conferir  só confere se .claude e Antigravity estão em dia (sai 1 se não)
#>
param([switch]$Conferir)
$ErrorActionPreference = 'Stop'

$ac       = Split-Path -Parent $PSScriptRoot
$vendor   = Join-Path $ac 'vendor\agent-skills'
$adapt    = Join-Path $ac 'adaptacao'
$destino  = Join-Path $ac '.claude'
$pastas   = @('skills', 'commands', 'agents', 'references')
$utf8     = New-Object System.Text.UTF8Encoding($false)

if (-not (Test-Path -LiteralPath (Join-Path $vendor 'skills'))) { throw ('Original não encontrado em ' + $vendor) }

$excluidos = @{}
$arqExc = Join-Path $adapt 'excluidos.txt'
if (Test-Path -LiteralPath $arqExc) {
    foreach ($l in [System.IO.File]::ReadAllLines($arqExc, $utf8)) {
        $l = $l.Trim()
        if ($l -and -not $l.StartsWith('#')) { $excluidos[($l -replace '/', '\').ToLowerInvariant()] = $true }
    }
}

function Ler([string]$p) { return [System.IO.File]::ReadAllText($p, $utf8) }
function Gravar([string]$p, [string]$t) {
    New-Item -ItemType Directory -Path (Split-Path -Parent $p) -Force | Out-Null
    [System.IO.File]::WriteAllText($p, $t, $utf8)
}
function Ajustar([string]$t) { return ($t -replace 'agent-skills:', '') }
function Somar([string]$texto, [string[]]$extras) {
    foreach ($e in $extras) {
        if (Test-Path -LiteralPath $e) { $texto = $texto.TrimEnd() + "`n`n" + (Ler $e).Trim() + "`n" }
    }
    return $texto
}

$temp = Join-Path ([System.IO.Path]::GetTempPath()) ('instalar-skills-' + $PID)
if (Test-Path -LiteralPath $temp) { Remove-Item -LiteralPath $temp -Recurse -Force }
$usados = @{}

# --- skills: a pasta inteira (SKILL.md + references\, scripts\ ...), adaptação no SKILL.md
foreach ($d in Get-ChildItem -LiteralPath (Join-Path $vendor 'skills') -Directory) {
    if ($excluidos[('skills\' + $d.Name).ToLowerInvariant()]) { continue }
    foreach ($f in Get-ChildItem -LiteralPath $d.FullName -Recurse -File) {
        $rel = $f.FullName.Substring($d.FullName.Length + 1)
        $alvo = Join-Path $temp ('skills\' + $d.Name + '\' + $rel)
        if ($f.Extension -eq '.md') {
            $t = Ajustar (Ler $f.FullName)
            if ($rel -eq 'SKILL.md') {
                $extra = Join-Path $adapt ('skills\' + $d.Name + '.md')
                if (Test-Path -LiteralPath $extra) { $usados[$extra.ToLowerInvariant()] = $true }
                $t = Somar $t @($extra)
            }
            Gravar $alvo $t
        } else {
            New-Item -ItemType Directory -Path (Split-Path -Parent $alvo) -Force | Out-Null
            Copy-Item -LiteralPath $f.FullName -Destination $alvo
        }
    }
}

# --- comandos do Claude Code (os .toml são de outros agentes)
$todos = Join-Path $adapt 'commands\_todos.md'
foreach ($f in Get-ChildItem -LiteralPath (Join-Path $vendor '.claude\commands') -Filter '*.md' -File) {
    if ($excluidos[('commands\' + $f.BaseName).ToLowerInvariant()]) { continue }
    $extra = Join-Path $adapt ('commands\' + $f.Name)
    if (Test-Path -LiteralPath $extra) { $usados[$extra.ToLowerInvariant()] = $true }
    Gravar (Join-Path $temp ('commands\' + $f.Name)) (Somar (Ajustar (Ler $f.FullName)) @($extra, $todos))
}

# --- agentes (personas dos revisores)
foreach ($f in Get-ChildItem -LiteralPath (Join-Path $vendor 'agents') -Filter '*.md' -File) {
    if ($excluidos[('agents\' + $f.BaseName).ToLowerInvariant()]) { continue }
    $extra = Join-Path $adapt ('agents\' + $f.Name)
    if (Test-Path -LiteralPath $extra) { $usados[$extra.ToLowerInvariant()] = $true }
    Gravar (Join-Path $temp ('agents\' + $f.Name)) (Somar (Ajustar (Ler $f.FullName)) @($extra))
}

# --- itens próprios do projeto (não existem no original): copiados como estão
foreach ($sub in @('skills', 'commands', 'agents')) {
    $p = Join-Path $adapt ('proprios\' + $sub)
    if (-not (Test-Path -LiteralPath $p)) { continue }
    foreach ($f in Get-ChildItem -LiteralPath $p -Recurse -File) {
        $rel = $f.FullName.Substring($p.Length + 1)
        $alvo = Join-Path $temp ($sub + '\' + $rel)
        if (Test-Path -LiteralPath $alvo) {
            Remove-Item -LiteralPath $temp -Recurse -Force
            throw ('Item próprio com o mesmo nome de um do original: adaptacao\proprios\' + $sub + '\' + $rel)
        }
        New-Item -ItemType Directory -Path (Split-Path -Parent $alvo) -Force | Out-Null
        Copy-Item -LiteralPath $f.FullName -Destination $alvo
    }
}

# --- checklists compartilhados (../../references a partir de skills\<nome>\SKILL.md)
foreach ($f in Get-ChildItem -LiteralPath (Join-Path $vendor 'references') -File) {
    Gravar (Join-Path $temp ('references\' + $f.Name)) (Ajustar (Ler $f.FullName))
}

# adaptação apontando para algo que não existe (ou foi excluído) é erro: ficaria esquecida
foreach ($sub in @('skills', 'commands', 'agents')) {
    $p = Join-Path $adapt $sub
    if (-not (Test-Path -LiteralPath $p)) { continue }
    foreach ($f in Get-ChildItem -LiteralPath $p -Filter '*.md' -File) {
        if ($f.Name -eq '_todos.md') { continue }
        if (-not $usados[$f.FullName.ToLowerInvariant()]) {
            Remove-Item -LiteralPath $temp -Recurse -Force
            throw ('Adaptação sem item correspondente no original (renomeado, removido ou excluído?): adaptacao\' + $sub + '\' + $f.Name)
        }
    }
}

Gravar (Join-Path $temp 'LEIA-ME.md') @"
# agent-config\.claude — GERADO

skills\, commands\, agents\ e references\ são gerados por
``ferramentas\instalar-skills.ps1`` a partir de ``vendor\agent-skills`` (original,
addyosmani/agent-skills, licença MIT) + ``adaptacao\`` (o que é deste projeto).
Não edite aqui: edite ``adaptacao\`` e rode o script. Detalhes em ``adaptacao\README.md``.
"@

# --- comparar com o que está instalado
# Chave em minúsculas (o Windows não distingue), mas guarda o caminho com a grafia real:
# é ele que se usa para gravar, e grafia diferente (skill.md x SKILL.md) conta como mudança -
# o Claude Code e o roteador procuram SKILL.md, e num sistema que distingue maiúsculas o
# arquivo com outra grafia não seria achado.
function Mapa([string]$raiz) {
    $m = @{}
    if (-not (Test-Path -LiteralPath $raiz)) { return $m }
    foreach ($p in $pastas + @('LEIA-ME.md')) {
        $alvo = Join-Path $raiz $p
        if (Test-Path -LiteralPath $alvo -PathType Leaf) {
            $f = Get-Item -LiteralPath $alvo
            $m[$p.ToLowerInvariant()] = [pscustomobject]@{ Rel = $f.FullName.Substring($raiz.Length + 1); Hash = (Get-FileHash -LiteralPath $alvo).Hash }
            continue
        }
        if (-not (Test-Path -LiteralPath $alvo)) { continue }
        foreach ($f in Get-ChildItem -LiteralPath $alvo -Recurse -File) {
            $rel = $f.FullName.Substring($raiz.Length + 1)
            $m[$rel.ToLowerInvariant()] = [pscustomobject]@{ Rel = $rel; Hash = (Get-FileHash -LiteralPath $f.FullName).Hash }
        }
    }
    return $m
}
function Difere([string]$k) {
    $a = $atual[$k]; $n = $novo[$k]
    return (-not $a) -or ($a.Hash -ne $n.Hash) -or ($a.Rel -cne $n.Rel)
}
$novo = Mapa $temp; $atual = Mapa $destino
$dif = @($novo.Keys | Where-Object { Difere $_ } | ForEach-Object { $novo[$_].Rel }) +
       @($atual.Keys | Where-Object { -not $novo.ContainsKey($_) } | ForEach-Object { $atual[$_].Rel })

# --- Antigravity (espelho por referência, sem duplicatas)
$geminiMd    = Join-Path $ac 'GEMINI.md'
$antigravity = Join-Path $ac '.agents'
$skillsJson  = Join-Path $antigravity 'skills.json'

function Conferir-Antigravity {
    $erros = @()
    if (-not (Test-Path -LiteralPath $geminiMd)) {
        $erros += 'GEMINI.md ausente em agent-config\'
    }
    if (-not (Test-Path -LiteralPath $skillsJson)) {
        $erros += '.agents\skills.json ausente em agent-config\'
    } else {
        $conteudo = Ler $skillsJson
        if (-not ($conteudo -match '\.claude/skills')) {
            $erros += '.agents\skills.json não referencia .claude/skills'
        }
    }
    return $erros
}

function Assegurar-Antigravity {
    if (-not (Test-Path -LiteralPath $antigravity)) {
        New-Item -ItemType Directory -Path $antigravity -Force | Out-Null
    }
    if (-not (Test-Path -LiteralPath $skillsJson)) {
        $json = @'
{
  "entries": [
    {
      "path": "agent-config/.claude/skills"
    },
    {
      "path": ".claude/skills"
    }
  ]
}
'@
        Gravar $skillsJson $json
    }
}

if ($Conferir) {
    Remove-Item -LiteralPath $temp -Recurse -Force
    $errosAgy = Conferir-Antigravity
    if ($dif.Count -gt 0 -or $errosAgy.Count -gt 0) {
        $dif | Sort-Object | Select-Object -First 20 | ForEach-Object { Write-Host ('  difere: ' + $_) }
        $errosAgy | ForEach-Object { Write-Host ('  antigravity: ' + $_) }
        Write-Host ('.claude ou Antigravity DESATUALIZADO. Rode ferramentas\instalar-skills.ps1.') -ForegroundColor Red
        exit 1
    }
    Write-Host '.claude e Antigravity em dia com vendor\agent-skills + adaptacao.' -ForegroundColor Green
    exit 0
}

# Sincroniza arquivo a arquivo, sem apagar as pastas: o Claude Code observa
# .claude\skills e um terminal pode estar parado dentro dela - apagar a pasta
# inteira falha no meio e deixa a instalação vazia.
New-Item -ItemType Directory -Path $destino -Force | Out-Null
foreach ($k in $atual.Keys) {
    if (-not $novo.ContainsKey($k)) { Remove-Item -LiteralPath (Join-Path $destino $atual[$k].Rel) -Force }
}
foreach ($k in $novo.Keys) {
    if (-not (Difere $k)) { continue }
    $rel = $novo[$k].Rel
    $para = Join-Path $destino $rel
    # sobrescrever mantém a grafia antiga do nome no Windows: só apagando antes ela muda
    if ($atual[$k] -and $atual[$k].Rel -cne $rel) { Remove-Item -LiteralPath (Join-Path $destino $atual[$k].Rel) -Force }
    New-Item -ItemType Directory -Path (Split-Path -Parent $para) -Force | Out-Null
    Copy-Item -LiteralPath (Join-Path $temp $rel) -Destination $para -Force
}
foreach ($p in $pastas) {
    $alvo = Join-Path $destino $p
    if (-not (Test-Path -LiteralPath $alvo)) { continue }
    # pastas que ficaram vazias (item removido do original ou da adaptação)
    Get-ChildItem -LiteralPath $alvo -Recurse -Directory | Sort-Object { $_.FullName.Length } -Descending |
        Where-Object { @(Get-ChildItem -LiteralPath $_.FullName -Force).Count -eq 0 } |
        ForEach-Object { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue }
}
Remove-Item -LiteralPath $temp -Recurse -Force

Assegurar-Antigravity
$errosAgy = Conferir-Antigravity
if ($errosAgy.Count -gt 0) {
    $errosAgy | ForEach-Object { Write-Warning ('Antigravity: ' + $_) }
}

$n = @{}
foreach ($p in @('skills', 'commands', 'agents')) {
    $n[$p] = @(Get-ChildItem -LiteralPath (Join-Path $destino $p) | Where-Object { $_.PSIsContainer -or $_.Extension -eq '.md' }).Count
}
Write-Host ('.claude gerado: ' + $n.skills + ' skills, ' + $n.commands + ' comandos, ' + $n.agents + ' agentes; ' +
            $usados.Count + ' adaptação(ões) aplicada(s); ' + $dif.Count + ' arquivo(s) mudaram.')
Write-Host ('Antigravity sincronizado: GEMINI.md e .agents\skills.json ativos referenciando .claude\skills.')
