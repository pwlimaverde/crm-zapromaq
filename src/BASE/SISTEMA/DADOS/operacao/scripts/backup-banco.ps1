<#
  backup-banco.ps1 - cópia de segurança do banco do CRM.

  Modos:
    -Modo Diario     exige EXCLUSIVIDADE (ninguém conectado: sem .laccdb). Sai com 2 se ocupado.
    -Modo Periodico  cópia simples, roda com gente usando (pode sair inconsistente).
    -Modo Agendado   o das tarefas agendadas: tenta o completo até -Tentativas vezes,
                     esperando -IntervaloMin entre elas; se o CRM continuar aberto,
                     faz a cópia simples com AVISO no log. Nunca termina sem cópia,
                     salvo falha de gravação.

  Nome do arquivo: crm_<diario|periodico>_AAAAMMDD-HHMM_<ESTACAO>.accdb
    diario    = cópia completa, feita sem ninguém conectado
    periodico = cópia simples, feita com o CRM aberto

  Retenção (todos os modos): mês corrente e o anterior (-Meses 2).
    Em setembro sobram agosto e setembro; em outubro, apaga agosto.
    A data vem do NOME do arquivo, não da data de modificação: Copy-Item preserva
    a data do banco original, então a data do arquivo não é a hora do backup.
    O expurgo só roda depois de uma cópia bem-sucedida, nunca deixa a pasta sem
    cópia e não mexe em subpastas (antes-da-publicacao) nem em arquivo fora do padrão.

  Destino: SISTEMA\DADOS\execucao\backup. Código de saída: 0 ok, 1 falha, 2 ocupado (modo diário).
#>
param(
    [ValidateSet('Periodico', 'Diario', 'Agendado')][string]$Modo = 'Periodico',
    [ValidateRange(1, 12)][int]$Tentativas = 3,
    [ValidateRange(1, 60)][int]$IntervaloMin = 10,
    [ValidateRange(1, 24)][int]$Meses = 2,
    [string]$PastaTeste = ''   # só para teste do expurgo: -PastaTeste <pasta> roda apenas o expurgo nela
)
$ErrorActionPreference = 'Stop'

function Anotar([string]$t) {
    $linha = (Get-Date -Format 'dd/MM/yyyy HH:mm:ss') + '  [' + $Modo + ']  [' + $env:COMPUTERNAME + ']  ' + $t
    Write-Host $linha
    Add-Content -LiteralPath $log -Value $linha -Encoding UTF8
}

function Invoke-Expurgo([string]$pasta) {
    $corte = (Get-Date -Day 1).Date.AddMonths(-($Meses - 1))
    $amanha = (Get-Date).Date.AddDays(1)
    $rx = '^crm_(diario|periodico)_(\d{8})-\d{4}_.+\.accdb$'
    $copias = @(Get-ChildItem -LiteralPath $pasta -Filter 'crm_*.accdb' -File | ForEach-Object {
        if ($_.Name -match $rx) {
            $d = [datetime]::MinValue
            if ([datetime]::TryParseExact($Matches[2], 'yyyyMMdd', [Globalization.CultureInfo]::InvariantCulture,
                                          [Globalization.DateTimeStyles]::None, [ref]$d)) {
                New-Object PSObject -Property @{ Arquivo = $_; Data = $d }
            }
        }
    })
    $futuras = @($copias | Where-Object { $_.Data -ge $amanha })
    if ($futuras.Count -gt 0) {
        # relógio da estação errado: apagar agora poderia levar cópias boas
        Anotar ('expurgo NÃO feito: ' + $futuras.Count + ' cópia(s) com data futura no nome. Confira o relógio da estação.')
        return
    }
    $velhos = @($copias | Where-Object { $_.Data -lt $corte })
    if ($velhos.Count -eq 0) { return }
    if (($copias.Count - $velhos.Count) -lt 1) {
        Anotar 'expurgo não feito: removeria a última cópia'
        return
    }
    $falhas = 0
    foreach ($v in $velhos) {
        try { Remove-Item -LiteralPath $v.Arquivo.FullName -Force }
        catch { $falhas++; Anotar ('expurgo: não removeu ' + $v.Arquivo.Name + ' -> ' + $_.Exception.Message.Split([char]13)[0]) }
    }
    Anotar ('expurgo: ' + ($velhos.Count - $falhas) + ' cópia(s) anterior(es) a ' + $corte.ToString('dd/MM/yyyy') + ' removida(s)')
}

# ---- modo de teste do expurgo: não toca no banco
if ($PastaTeste) {
    $log = Join-Path $PastaTeste 'backup.log'
    Invoke-Expurgo $PastaTeste
    exit 0
}

. (Join-Path (Split-Path -Parent $PSScriptRoot) '..\lib\Comum.ps1')

$raiz = Get-RaizBase $PSScriptRoot
$banco = Get-CaminhoBanco $raiz
$pastaBkp = Get-PastaExecucao $raiz 'backup'
$log = Join-Path $pastaBkp 'backup.log'
$trava = [System.IO.Path]::ChangeExtension($banco, 'laccdb')

if (-not (Test-Path -LiteralPath $banco)) {
    # tarefa agendada que falha em silêncio é pior que tarefa que não existe
    Anotar ('FALHOU: banco não encontrado -> ' + $banco)
    exit 1
}

# ---- decide o tipo de cópia
$temUsuario = Test-Path -LiteralPath $trava
if ($Modo -eq 'Agendado') {
    for ($i = 1; $temUsuario -and $i -lt $Tentativas; $i++) {
        Anotar ('CRM aberto (tentativa ' + $i + ' de ' + $Tentativas + '): nova tentativa em ' + $IntervaloMin + ' min')
        Start-Sleep -Seconds ($IntervaloMin * 60)
        $temUsuario = Test-Path -LiteralPath $trava
    }
    $tipo = if ($temUsuario) { 'periodico' } else { 'diario' }
} elseif ($Modo -eq 'Diario') {
    if ($temUsuario) {
        Anotar 'ABORTADO: há usuário conectado (.laccdb presente). O backup diário exige exclusividade.'
        exit 2
    }
    $tipo = 'diario'
} else {
    $tipo = 'periodico'
}
if ($temUsuario) { Anotar 'AVISO: há usuário conectado. Cópia simples: sai, mas pode não estar consistente.' }

# ---- cópia
$destino = Join-Path $pastaBkp ('crm_' + $tipo + '_' + (Get-Date -Format 'yyyyMMdd-HHmm') + '_' + $env:COMPUTERNAME + '.accdb')
try {
    Copy-Item -LiteralPath $banco -Destination $destino -Force
    $tamOrigem = (Get-Item -LiteralPath $banco).Length
    $tamCopia = (Get-Item -LiteralPath $destino).Length
    if ($tamCopia -ne $tamOrigem) {
        Anotar ('FALHA: cópia com ' + $tamCopia + ' bytes contra ' + $tamOrigem + ' do original. Removida.')
        Remove-Item -LiteralPath $destino -Force
        exit 1
    }
    Anotar ('OK: ' + (Split-Path -Leaf $destino) + '  (' + [math]::Round($tamCopia / 1MB, 2) + ' MB)')
} catch {
    Anotar ('FALHA ao copiar: ' + $_.Exception.Message.Split([char]13)[0])
    exit 1
}

# ---- expurgo: só depois de cópia bem-sucedida
Invoke-Expurgo $pastaBkp

$total = @(Get-ChildItem -LiteralPath $pastaBkp -Filter '*.accdb' -File)
Anotar ('na pasta: ' + $total.Count + ' cópia(s), ' + [math]::Round((($total | Measure-Object Length -Sum).Sum) / 1MB, 1) + ' MB')
exit 0
