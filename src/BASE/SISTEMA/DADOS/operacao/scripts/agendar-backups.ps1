<#
  agendar-backups.ps1 - cria (ou remove) as tarefas de backup do CRM nesta estação.

  NÃO exige administrador: a tarefa roda como o próprio usuário, com nível
  de execução LIMITADO e logon interativo - o Agendador de Tarefas permite
  isso a qualquer usuário. (A versão anterior usava RunLevel Highest, que é
  o que exigia administrador; o backup não precisa de privilégio elevado.)

  Consequência a conhecer: a tarefa só roda com esse usuário conectado na
  estação. Se a máquina estiver desligada no horário, roda assim que ligar.
  Pode ser instalado em mais de uma estação: o arquivo leva o nome da máquina.

  Uso: agendar-backups.ps1 [-Horarios '09:00','16:40'] [-Remover]
#>
param(
    [string[]]$Horarios = @('09:00', '16:40'),
    [switch]$Remover
)
$ErrorActionPreference = 'Stop'
. (Join-Path (Split-Path -Parent $PSScriptRoot) '..\lib\Comum.ps1')

$raiz = Get-RaizBase $PSScriptRoot
$script = ConvertTo-CaminhoUNC (Join-Path $PSScriptRoot 'backup-banco.ps1')
$ps = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'

Write-Host '=== AGENDAMENTO DE BACKUP - CRM Zapromaq ==='
Write-Host ('Estação.: ' + $env:COMPUTERNAME + '   usuário: ' + $env:USERDOMAIN + '\' + $env:USERNAME)
Write-Host ('Script..: ' + $script)
Write-Host ''

foreach ($h in $Horarios) {
    $nome = 'CRM Zapromaq - backup ' + $h.Replace(':', '')
    try { Unregister-ScheduledTask -TaskName $nome -Confirm:$false -ErrorAction SilentlyContinue } catch { }
    if ($Remover) { Write-Host ('removida: ' + $nome); continue }

    $acao = New-ScheduledTaskAction -Execute $ps -Argument ('-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "' + $script + '" -Modo Periodico')
    $gatilho = New-ScheduledTaskTrigger -Daily -At $h
    $config = New-ScheduledTaskSettingsSet -StartWhenAvailable -DontStopIfGoingOnBatteries -AllowStartIfOnBatteries `
                                           -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Minutes 30)
    # Nunca SYSTEM: essa conta acessa a rede como conta de máquina e não enxerga o compartilhamento.
    $principal = New-ScheduledTaskPrincipal -UserId ($env:USERDOMAIN + '\' + $env:USERNAME) -LogonType Interactive -RunLevel Limited
    Register-ScheduledTask -TaskName $nome -Action $acao -Trigger $gatilho -Settings $config -Principal $principal -Force | Out-Null
    Write-Host ('criada: ' + $nome.PadRight(34) + ' próxima execução: ' + (Get-ScheduledTaskInfo -TaskName $nome).NextRunTime)
}

if (-not $Remover) {
    $pastaBkp = Get-PastaExecucao $raiz 'backup'
    $antes = @(Get-ChildItem -LiteralPath $pastaBkp -Filter '*.accdb' -File).Count
    $nomeTeste = 'CRM Zapromaq - backup ' + $Horarios[0].Replace(':', '')
    Write-Host ''
    Write-Host '--- teste imediato (a tarefa precisa provar que enxerga a pasta da rede) ---'
    try {
        Start-ScheduledTask -TaskName $nomeTeste
        $fim = (Get-Date).AddSeconds(30)
        do { Start-Sleep -Seconds 2; $depois = @(Get-ChildItem -LiteralPath $pastaBkp -Filter '*.accdb' -File).Count }
        while ($depois -le $antes -and (Get-Date) -lt $fim)
        if ($depois -gt $antes) { Write-Host ('  OK: a tarefa gerou backup (' + $antes + ' -> ' + $depois + ' cópias).') }
        else { Write-Host '  ATENÇÃO: nenhuma cópia nova em 30 s. Confira execucao\backup\backup.log e o histórico da tarefa.' }
    } catch { Write-Host ('  não foi possível disparar o teste: ' + $_.Exception.Message.Split([char]13)[0]) }
}
