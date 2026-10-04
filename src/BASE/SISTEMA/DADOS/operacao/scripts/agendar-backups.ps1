<#
  agendar-backups.ps1 - cria (ou remove) as tarefas de backup do CRM nesta estação.

  NÃO exige administrador. Cria as tarefas pelo schtasks.exe com definição em XML
  (Register-ScheduledTask devolveu "Acesso negado" para usuário comum nesta rede;
  o schtasks /Create /XML é o caminho que o Windows aceita para o próprio usuário).
  A tarefa roda como o próprio usuário, nível LIMITADO, só com ele conectado.
  Se a máquina estiver desligada no horário, roda assim que ligar.

  Horários: 09:00, 12:30 e 16:00, de segunda a sexta. Cada tarefa chama
  backup-banco.ps1 -Modo Agendado: se o CRM estiver aberto, tenta de novo a cada
  10 min (3 tentativas) e, persistindo, faz a cópia simples. Retenção: mês
  corrente e anterior (ver backup-banco.ps1).

  Nomes: 'CRM Zapromaq - backup auto HHMM'. O prefixo "auto" evita colisão com
  tarefas antigas 'CRM Zapromaq - backup 0900/1640', que podem ter sido criadas
  como administrador e ficam invisíveis/intocáveis para usuário comum.

  Uso: agendar-backups.ps1 [-Horarios '09:00','12:30','16:00'] [-Remover]
#>
param(
    [string[]]$Horarios = @('09:00', '12:30', '16:00'),
    [switch]$Remover
)
$ErrorActionPreference = 'Stop'
. (Join-Path (Split-Path -Parent $PSScriptRoot) '..\lib\Comum.ps1')

$raiz = Get-RaizBase $PSScriptRoot
$script = ConvertTo-CaminhoUNC (Join-Path $PSScriptRoot 'backup-banco.ps1')
$ps = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$schtasks = Join-Path $env:SystemRoot 'System32\schtasks.exe'
$usuario = $env:USERDOMAIN + '\' + $env:USERNAME
$prefixo = 'CRM Zapromaq - backup auto '

function Invoke-Schtasks([string[]]$argumentos) {
    # schtasks escreve erro no stderr; com ErrorActionPreference=Stop o PowerShell 5.1
    # transformaria isso em exceção antes de lermos o código de saída.
    $eap = $ErrorActionPreference; $ErrorActionPreference = 'Continue'
    try { $saida = & $schtasks @argumentos 2>&1 | ForEach-Object { "$_" } }
    finally { $ErrorActionPreference = $eap }
    return New-Object PSObject -Property @{ Codigo = $LASTEXITCODE; Texto = (($saida | Where-Object { $_ }) -join ' ').Trim() }
}

function Esc([string]$t) { return [System.Security.SecurityElement]::Escape($t) }

function New-XmlTarefa([string]$horario, [string]$modo, [string]$descricao) {
    $gatilho = ''
    if ($horario) {
        $gatilho = @"
  <Triggers>
    <CalendarTrigger>
      <StartBoundary>$((Get-Date).ToString('yyyy-MM-dd'))T$($horario):00</StartBoundary>
      <Enabled>true</Enabled>
      <ScheduleByWeek>
        <DaysOfWeek><Monday /><Tuesday /><Wednesday /><Thursday /><Friday /></DaysOfWeek>
        <WeeksInterval>1</WeeksInterval>
      </ScheduleByWeek>
    </CalendarTrigger>
  </Triggers>
"@
    }
    $argumentos = '-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "' + $script + '" -Modo ' + $modo
    return @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Author>$(Esc $usuario)</Author>
    <Description>$(Esc $descricao)</Description>
  </RegistrationInfo>
$gatilho
  <Principals>
    <Principal id="Author">
      <UserId>$(Esc $usuario)</UserId>
      <LogonType>InteractiveToken</LogonType>
      <RunLevel>LeastPrivilege</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <StartWhenAvailable>true</StartWhenAvailable>
    <ExecutionTimeLimit>PT45M</ExecutionTimeLimit>
    <Enabled>true</Enabled>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>$(Esc $ps)</Command>
      <Arguments>$(Esc $argumentos)</Arguments>
    </Exec>
  </Actions>
</Task>
"@
}

function Register-Tarefa([string]$nome, [string]$xml) {
    $arq = Join-Path $env:TEMP ('crm-tarefa-' + [guid]::NewGuid().ToString('N') + '.xml')
    try {
        [System.IO.File]::WriteAllText($arq, $xml, [System.Text.Encoding]::Unicode)
        return (Invoke-Schtasks @('/Create', '/TN', $nome, '/XML', $arq, '/F'))
    } finally { Remove-Item -LiteralPath $arq -Force -ErrorAction SilentlyContinue }
}

Write-Host '=== AGENDAMENTO DE BACKUP - CRM Zapromaq ==='
Write-Host ('Estação.: ' + $env:COMPUTERNAME + '   usuário: ' + $usuario)
Write-Host ('Script..: ' + $script)
Write-Host ''

# ---- remove as tarefas de backup do CRM que este usuário enxerga (novas e antigas)
$lista = Invoke-Schtasks @('/Query', '/FO', 'CSV', '/NH')
$existentes = @()
if ($lista.Codigo -eq 0) {
    $existentes = @($lista.Texto -split '"\s*"|"' | Where-Object { $_ -like '\CRM Zapromaq - backup*' } | Select-Object -Unique)
}
foreach ($t in $existentes) {
    $r = Invoke-Schtasks @('/Delete', '/TN', $t, '/F')
    if ($r.Codigo -eq 0) { Write-Host ('removida: ' + $t.TrimStart('\')) }
    else { Write-Host ('NÃO removida: ' + $t.TrimStart('\') + ' -> ' + $r.Texto) }
}
# tarefas antigas podem existir sem aparecer na consulta (criadas como administrador)
foreach ($antiga in 'CRM Zapromaq - backup 0900', 'CRM Zapromaq - backup 1640') {
    if ($existentes -contains ('\' + $antiga)) { continue }
    $r = Invoke-Schtasks @('/Delete', '/TN', $antiga, '/F')
    if ($r.Codigo -eq 0) { Write-Host ('removida: ' + $antiga) }
    elseif ($r.Texto -match 'negado|denied') {
        Write-Host ('AVISO: existe a tarefa antiga "' + $antiga + '" criada como administrador.')
        Write-Host '       Ela não impede as novas, mas só um administrador consegue removê-la.'
    }
}
if ($Remover) { return }

$falhou = $false
foreach ($h in $Horarios) {
    $nome = $prefixo + $h.Replace(':', '')
    $xml = New-XmlTarefa $h 'Agendado' ('Backup do CRM Zapromaq - ' + $h + ', seg a sex. Log: SISTEMA\DADOS\execucao\backup\backup.log')
    $r = Register-Tarefa $nome $xml
    if ($r.Codigo -eq 0) {
        $prox = ''
        try { $prox = (Get-ScheduledTaskInfo -TaskName $nome -ErrorAction Stop).NextRunTime } catch { }
        Write-Host ('criada: ' + $nome.PadRight(36) + ' próxima execução: ' + $prox)
    } else {
        $falhou = $true
        Write-Host ('FALHOU: ' + $nome + ' -> ' + $r.Texto)
    }
}

if ($falhou) {
    Write-Host ''
    Write-Host 'ATENÇÃO: nem todas as tarefas foram criadas. Envie esta tela para análise.'
    return
}

$pastaBkp = Get-PastaExecucao $raiz 'backup'
$antes = @(Get-ChildItem -LiteralPath $pastaBkp -Filter '*.accdb' -File).Count
# Teste numa tarefa temporária em modo Periodico: o modo Agendado esperaria
# até 20 min se o CRM estiver aberto, e o teste precisa responder em 30 s.
$nomeTeste = 'CRM Zapromaq - teste do agendamento'
Write-Host ''
Write-Host '--- teste imediato (a tarefa precisa provar que enxerga a pasta da rede) ---'
try {
    $r = Register-Tarefa $nomeTeste (New-XmlTarefa '' 'Periodico' 'Teste temporário do agendamento de backup')
    if ($r.Codigo -ne 0) { throw ('não criou a tarefa de teste: ' + $r.Texto) }
    $r = Invoke-Schtasks @('/Run', '/TN', $nomeTeste)
    if ($r.Codigo -ne 0) { throw ('não disparou: ' + $r.Texto) }
    $fim = (Get-Date).AddSeconds(30)
    do { Start-Sleep -Seconds 2; $depois = @(Get-ChildItem -LiteralPath $pastaBkp -Filter '*.accdb' -File).Count }
    while ($depois -le $antes -and (Get-Date) -lt $fim)
    if ($depois -gt $antes) { Write-Host ('  OK: a tarefa gerou backup (' + $antes + ' -> ' + $depois + ' cópias).') }
    else { Write-Host '  ATENÇÃO: nenhuma cópia nova em 30 s. Confira execucao\backup\backup.log e o histórico da tarefa.' }
} catch {
    Write-Host ('  teste não concluído: ' + "$_".Split([char]13)[0])
} finally {
    Invoke-Schtasks @('/Delete', '/TN', $nomeTeste, '/F') | Out-Null
}
