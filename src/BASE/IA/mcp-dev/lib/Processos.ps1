<#
  Processos.ps1 - execução de scripts auxiliares do MCP de desenvolvimento.

  Os scripts de build/visual/migração rodam num processo separado para que a
  saída deles nunca vaze no stdout do servidor (quebraria o protocolo MCP).

  Sem parâmetros de política na linha de comando: o processo filho herda a
  política do próprio servidor (variável de ambiente do processo), e usa o
  mesmo executável do PowerShell que está rodando este servidor.
#>

function Get-ExecutavelPowerShell {
    try { $p = (Get-Process -Id $PID).Path; if ($p) { return $p } } catch { }
    return (Join-Path $PSHOME 'powershell.exe')
}

# Roda um executável, captura saída e erro, respeita o limite de tempo.
function Invoke-ProcessoFilho([string]$exe, [string]$argumentos, [int]$limiteSeg) {
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $exe
    $psi.Arguments = $argumentos
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.RedirectStandardInput = $true
    $psi.StandardOutputEncoding = New-Object System.Text.UTF8Encoding($false)
    $psi.CreateNoWindow = $true
    $p = [System.Diagnostics.Process]::Start($psi)
    $p.StandardInput.Close()
    $out = $p.StandardOutput.ReadToEndAsync()
    $err = $p.StandardError.ReadToEndAsync()
    if (-not $p.WaitForExit($limiteSeg * 1000)) {
        try { $p.Kill() } catch { }
        throw ('Passou de ' + $limiteSeg + ' s e foi interrompido: ' + [System.IO.Path]::GetFileName($exe) + ' ' + $argumentos)
    }
    $p.WaitForExit()
    return [ordered]@{ codigo = $p.ExitCode; saida = ($out.Result + $err.Result) }
}

# Script .ps1 de SISTEMA\DADOS (build, visual, migrações).
function Invoke-ScriptFilho([string]$script, [string]$argumentos, [int]$limiteSeg) {
    return (Invoke-ProcessoFilho (Get-ExecutavelPowerShell) ('-NoLogo -NoProfile -File "' + $script + '" ' + $argumentos) $limiteSeg)
}

# Últimas n linhas de um texto (relatório curto para o modelo).
function Get-Final([string]$texto, [int]$n) {
    $l = @($texto -split "`r?`n" | Where-Object { $_.Trim() -ne '' })
    if ($l.Count -le $n) { return ($l -join "`n") }
    return ($l[($l.Count - $n)..($l.Count - 1)] -join "`n")
}
