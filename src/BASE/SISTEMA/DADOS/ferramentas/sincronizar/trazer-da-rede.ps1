<#
  trazer-da-rede.ps1 - traz para o repositório de desenvolvimento os ajustes
  feitos NA REDE (MCP de desenvolvimento / Claude Desktop da empresa).
  Roda na máquina de DESENVOLVIMENTO. Regra 6 do plano: antes de qualquer
  desenvolvimento novo, a rede volta para o repositório.

  Compara SISTEMA\DADOS\fonte da rede com a do repositório (por hash):
    - arquivo só na rede ou diferente -> copiado para o repositório
    - arquivo só no repositório       -> listado (não apaga nada)
  Depois: git diff mostra o que mudou; rode verificar.py e faça o commit.
  O registro execucao\dev\alteracoes.log da rede diz quem mudou o quê e por quê.

  Uso: trazer-da-rede.ps1 -Rede \\servidor\...\BASE [-Simular]
#>
param(
    [Parameter(Mandatory = $true)][string]$Rede,
    [switch]$Simular
)
$ErrorActionPreference = 'Stop'

$dadosRepo = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$fonteRepo = Join-Path $dadosRepo 'fonte'
$fonteRede = Join-Path $Rede 'SISTEMA\DADOS\fonte'
if (-not (Test-Path -LiteralPath $fonteRede)) { throw ('Não achei ' + $fonteRede + ' (passe a pasta BASE da rede em -Rede).') }

function Get-Mapa([string]$raiz) {
    $m = @{}
    $r = [System.IO.Path]::GetFullPath($raiz).TrimEnd('\') + '\'
    foreach ($f in Get-ChildItem -LiteralPath $raiz -Recurse -File) {
        $m[$f.FullName.Substring($r.Length)] = (Get-FileHash -LiteralPath $f.FullName -Algorithm SHA256).Hash
    }
    return $m
}

$mapaRede = Get-Mapa $fonteRede
$mapaRepo = Get-Mapa $fonteRepo
$copiados = 0
foreach ($rel in ($mapaRede.Keys | Sort-Object)) {
    if ($mapaRepo.ContainsKey($rel) -and $mapaRepo[$rel] -eq $mapaRede[$rel]) { continue }
    $situacao = if ($mapaRepo.ContainsKey($rel)) { 'ALTERADO' } else { 'NOVO    ' }
    Write-Host ($situacao + '  ' + $rel)
    if (-not $Simular) {
        $destino = Join-Path $fonteRepo $rel
        New-Item -ItemType Directory -Path (Split-Path -Parent $destino) -Force | Out-Null
        Copy-Item -LiteralPath (Join-Path $fonteRede $rel) -Destination $destino -Force
    }
    $copiados++
}
foreach ($rel in ($mapaRepo.Keys | Sort-Object)) {
    if (-not $mapaRede.ContainsKey($rel)) { Write-Host ('SÓ AQUI   ' + $rel + '   (não existe na rede; nada foi apagado)') }
}

$log = Join-Path $Rede 'SISTEMA\DADOS\execucao\dev\alteracoes.log'
Write-Host ''
if ($copiados -eq 0) { Write-Host 'Rede e repositório iguais em fonte\.' }
elseif ($Simular) { Write-Host ($copiados.ToString() + ' arquivo(s) seriam trazidos (simulação: nada foi copiado).') }
else { Write-Host ($copiados.ToString() + ' arquivo(s) trazidos. Agora: git diff, python ferramentas\verificacao\verificar.py, commit.') }
if (Test-Path -LiteralPath $log) {
    Write-Host ''
    Write-Host 'Registro de alterações feitas pelo MCP de desenvolvimento (últimas 20):'
    Get-Content -LiteralPath $log -Encoding UTF8 | Select-Object -Last 20 | ForEach-Object {
        $a = $_ | ConvertFrom-Json
        Write-Host ('  ' + $a.quando + '  ' + $a.usuario + '  ' + $a.caminho + '  ' + $a.motivo)
    }
}
