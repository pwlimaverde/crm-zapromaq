<#
  instalar-mcp.ps1 - registra (ou remove) o servidor MCP do CRM no Claude Desktop
  DESTE usuário. Não exige administrador: só edita o arquivo de configuração do
  próprio usuário, com cópia de segurança antes.

  O Claude Desktop instalado pela Microsoft Store (MSIX) lê a configuração de
  uma pasta virtualizada, diferente da que o botão "Edit Config" abre. Por isso
  o script grava nas duas quando existirem.

  Uso: instalar-mcp.ps1 [-Dev] [-Remover]
    -Dev  registra o MCP de DESENVOLVIMENTO (IA\mcp-dev, "crm-zapromaq-dev"),
          que altera as fontes do front - só nas máquinas de quem mantém o sistema.
#>
param([switch]$Remover, [switch]$Dev, [ValidateSet('Claude','Codex','Ambos')][string]$Aplicativo = 'Claude')
$ErrorActionPreference = 'Stop'

if ($Dev) {
    $nomeServidor = 'crm-zapromaq-dev'
    $servidor = (Resolve-Path (Join-Path $PSScriptRoot '..\mcp-dev\servidor-dev.ps1')).ProviderPath
    $qtdFerramentas = 15
} else {
    $nomeServidor = 'crm-zapromaq'
    $servidor = (Resolve-Path (Join-Path $PSScriptRoot 'servidor.ps1')).ProviderPath
    $qtdFerramentas = 25
}

# Unidade mapeada (ex.: G:) vira caminho UNC: o Claude Desktop pode subir antes
# de a unidade ser reconectada, e o caminho UNC não depende da letra.
if ($servidor -match '^([A-Za-z]):\\') {
    $letra = $Matches[1] + ':'
    $disco = Get-WmiObject Win32_LogicalDisk -Filter ("DeviceID='" + $letra + "'") -ErrorAction SilentlyContinue
    if ($disco -and $disco.DriveType -eq 4 -and $disco.ProviderName) {
        $servidor = $disco.ProviderName + $servidor.Substring(2)
        Write-Host ('Unidade de rede convertida para UNC: ' + $servidor)
    }
}

$powershell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$entrada = [ordered]@{
    command = $powershell
    args = @('-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', $servidor)
}

if ($Aplicativo -in @('Claude','Ambos')) {
# arquivos de configuração candidatos
$alvos = New-Object System.Collections.ArrayList
$padrao = Join-Path $env:APPDATA 'Claude\claude_desktop_config.json'
if (Test-Path (Split-Path $padrao)) { [void]$alvos.Add($padrao) }
Get-ChildItem (Join-Path $env:LOCALAPPDATA 'Packages') -Directory -Filter 'Claude_*' -ErrorAction SilentlyContinue | ForEach-Object {
    $pasta = Join-Path $_.FullName 'LocalCache\Roaming\Claude'
    if (-not (Test-Path $pasta)) { New-Item -ItemType Directory -Path $pasta -Force | Out-Null }
    [void]$alvos.Add((Join-Path $pasta 'claude_desktop_config.json'))
}
if ($alvos.Count -eq 0) {
    throw 'Claude Desktop não encontrado para este usuário. Instale e abra o Claude Desktop uma vez antes.'
}

$utf8 = New-Object System.Text.UTF8Encoding($false)
foreach ($arq in $alvos) {
    $cfg = New-Object PSObject
    if (Test-Path $arq) {
        $texto = [IO.File]::ReadAllText($arq)
        if ($texto.Trim()) { $cfg = $texto | ConvertFrom-Json }
        Copy-Item $arq ($arq + '.bak-' + (Get-Date -Format 'yyyyMMdd-HHmmss'))
    }
    if (-not ($cfg.PSObject.Properties.Name -contains 'mcpServers')) {
        $cfg | Add-Member -NotePropertyName mcpServers -NotePropertyValue (New-Object PSObject)
    }
    if ($Remover) {
        $cfg.mcpServers.PSObject.Properties.Remove($nomeServidor)
        Write-Host ('removido de: ' + $arq)
    } else {
        $cfg.mcpServers | Add-Member -NotePropertyName $nomeServidor -NotePropertyValue $entrada -Force
        Write-Host ('registrado em: ' + $arq)
    }
    [IO.File]::WriteAllText($arq, (ConvertTo-Json -InputObject $cfg -Depth 50), $utf8)
}

Write-Host ''
Write-Host 'Feche o Claude Desktop POR COMPLETO (também o ícone perto do relógio) e abra de novo.'
if (-not $Remover) {
    Write-Host ('Depois, no campo de mensagem: "+" > Conectores > ' + $nomeServidor + ' deve aparecer com ' + $qtdFerramentas + ' ferramentas.')
}

} # configuracao Claude

if ($Aplicativo -in @('Codex','Ambos')) {
    $codexCmd = Get-Command codex -ErrorAction SilentlyContinue
    if (-not $codexCmd) {
        $codexCmd = Get-ChildItem (Join-Path $env:LOCALAPPDATA 'OpenAI\Codex\bin\*\codex.exe') -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1
    }
    if (-not $codexCmd) { throw 'Codex nao encontrado. Instale e abra o Codex antes.' }
    $codexExe = if ($codexCmd.Source) { $codexCmd.Source } else { $codexCmd.FullName }
    $codexRoot = if ($env:CODEX_HOME) { $env:CODEX_HOME } else { Join-Path $env:USERPROFILE '.codex' }
    $cfgCodex = Join-Path $codexRoot 'config.toml'
    if (Test-Path -LiteralPath $cfgCodex) {
        Copy-Item -LiteralPath $cfgCodex -Destination ($cfgCodex + '.bak-' + (Get-Date -Format 'yyyyMMdd-HHmmss-fff'))
    }
    if ($Remover) {
        & $codexExe mcp remove $nomeServidor
    } else {
        # O Codex inicia MCP com ambiente reduzido. OLE DB/ACE depende dos
        # caminhos de componentes compartilhados do Windows/Office.
        $envArgs = @()
        foreach ($nomeEnv in @('SystemRoot','SystemDrive','CommonProgramFiles','CommonProgramW6432','ProgramFiles','ProgramW6432','TEMP','TMP','USERPROFILE','LOCALAPPDATA','APPDATA')) {
            $valorEnv = [Environment]::GetEnvironmentVariable($nomeEnv)
            if (-not [string]::IsNullOrWhiteSpace($valorEnv)) {
                $envArgs += @('--env', ($nomeEnv + '=' + $valorEnv))
            }
        }
        & $codexExe mcp add $nomeServidor @envArgs -- $powershell @($entrada.args)
    }
    if ($LASTEXITCODE -ne 0) { throw 'Falha ao configurar o MCP no Codex.' }
    if (-not $Remover) {
        $origemSkill = Join-Path (Split-Path $PSScriptRoot -Parent) ('instrucoes\' + $nomeServidor + '\SKILL.md')
        $destinoSkill = Join-Path $codexRoot ('skills\' + $nomeServidor)
        if (-not (Test-Path -LiteralPath $destinoSkill)) { New-Item -ItemType Directory -Path $destinoSkill -Force | Out-Null }
        $arquivoSkill = Join-Path $destinoSkill 'SKILL.md'
        if (Test-Path -LiteralPath $arquivoSkill) {
            Copy-Item -LiteralPath $arquivoSkill -Destination ($arquivoSkill + '.bak-' + (Get-Date -Format 'yyyyMMdd-HHmmss-fff'))
        }
        Copy-Item -LiteralPath $origemSkill -Destination $arquivoSkill -Force
        Write-Host ('Skill instalada em: ' + $arquivoSkill)
    }
    Write-Host 'Reabra o Codex para carregar a conexao e a skill.'
}
