<#
  backup-banco.ps1 - cópia de segurança do banco do CRM.

  Modos:
    -Modo Periodico  cópia rápida, retenção em horas (padrão 48). Roda com gente usando.
    -Modo Diario     exige EXCLUSIVIDADE (ninguém conectado: sem .laccdb), retenção em dias (padrão 30).

  Cópia com usuário conectado é melhor que nada, mas não é consistente: pode
  pegar o arquivo no meio de uma gravação. Por isso existe o modo diário, e
  por isso a restauração precisa ser testada (restaurar-banco.ps1).

  Destino: SISTEMA\DADOS\execucao\backup. Código de saída: 0 ok, 1 falha, 2 ocupado (modo diário).
#>
param(
    [ValidateSet('Periodico', 'Diario')][string]$Modo = 'Periodico',
    [int]$RetencaoHoras = 48,
    [int]$RetencaoDias = 30
)
$ErrorActionPreference = 'Stop'
. (Join-Path (Split-Path -Parent $PSScriptRoot) '..\lib\Comum.ps1')

$raiz = Get-RaizBase $PSScriptRoot
$banco = Get-CaminhoBanco $raiz
$pastaBkp = Get-PastaExecucao $raiz 'backup'
$log = Join-Path $pastaBkp 'backup.log'

function Anotar([string]$t) {
    $linha = (Get-Date -Format 'dd/MM/yyyy HH:mm:ss') + '  [' + $Modo + ']  [' + $env:COMPUTERNAME + ']  ' + $t
    Write-Host $linha
    Add-Content -LiteralPath $log -Value $linha -Encoding UTF8
}

if (-not (Test-Path -LiteralPath $banco)) {
    # tarefa agendada que falha em silêncio é pior que tarefa que não existe
    Anotar ('FALHOU: banco não encontrado -> ' + $banco)
    exit 1
}

$temUsuario = Test-Path -LiteralPath ([System.IO.Path]::ChangeExtension($banco, 'laccdb'))
if ($Modo -eq 'Diario' -and $temUsuario) {
    Anotar 'ABORTADO: há usuário conectado (.laccdb presente). O backup diário exige exclusividade.'
    exit 2
}
if ($temUsuario) { Anotar 'AVISO: há usuário conectado. A cópia sai, mas pode não estar consistente.' }

$destino = Join-Path $pastaBkp ('crm_' + $Modo.ToLower() + '_' + (Get-Date -Format 'yyyyMMdd-HHmm') + '_' + $env:COMPUTERNAME + '.accdb')
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

# ---- expurgo: nunca deixa a pasta sem nenhuma cópia daquele modo
$limite = if ($Modo -eq 'Periodico') { (Get-Date).AddHours(-$RetencaoHoras) } else { (Get-Date).AddDays(-$RetencaoDias) }
$padrao = 'crm_' + $Modo.ToLower() + '_*.accdb'
$todas = @(Get-ChildItem -LiteralPath $pastaBkp -Filter $padrao -File)
$velhos = @($todas | Where-Object { $_.LastWriteTime -lt $limite })
if ($velhos.Count -gt 0 -and ($todas.Count - $velhos.Count) -ge 1) {
    foreach ($v in $velhos) { Remove-Item -LiteralPath $v.FullName -Force }
    Anotar ('expurgo: ' + $velhos.Count + ' cópia(s) anterior(es) ao limite removida(s)')
} elseif ($velhos.Count -gt 0) {
    Anotar 'expurgo não feito: removeria a última cópia deste modo'
}

$total = @(Get-ChildItem -LiteralPath $pastaBkp -Filter '*.accdb' -File)
Anotar ('na pasta: ' + $total.Count + ' cópia(s), ' + [math]::Round((($total | Measure-Object Length -Sum).Sum) / 1MB, 1) + ' MB')
exit 0
