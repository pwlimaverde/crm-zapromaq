<#
  restaurar-banco.ps1 - restaura um backup numa CÓPIA DE TESTE e confere o que voltou.

  Backup que nunca foi restaurado não é backup. Este script nunca sobrescreve
  a produção: restaura numa cópia em execucao\backup e conta registros,
  soma de valor, auditoria e relações, para a pessoa comparar com o esperado.

  Uso: restaurar-banco.ps1 [-Backup <arquivo.accdb>]   (padrão: o mais recente)
#>
param([string]$Backup = '')
$ErrorActionPreference = 'Stop'
. (Join-Path (Split-Path -Parent $PSScriptRoot) '..\lib\Comum.ps1')

$raiz = Get-RaizBase $PSScriptRoot
$pastaBkp = Get-PastaExecucao $raiz 'backup'
$log = New-Log $raiz 'restauracao'

if ([string]::IsNullOrWhiteSpace($Backup)) {
    $ultimo = Get-ChildItem -LiteralPath $pastaBkp -Filter 'crm_*.accdb' -File | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($null -eq $ultimo) {
        $log.L('Nenhum backup encontrado em ' + $pastaBkp + '. Gere um com FAZER-BACKUP.bat.')
        $log.Salvar(); exit 1
    }
    $Backup = $ultimo.FullName
}
$destino = Join-Path $pastaBkp ('restauracao-teste-' + (Get-Date -Format 'yyyyMMdd-HHmm') + '.accdb')

$log.L('=== TESTE DE RESTAURAÇÃO ===')
$log.L('Backup..: ' + (Split-Path -Leaf $Backup) + '  (gerado em ' + (Get-Item -LiteralPath $Backup).LastWriteTime + ')')
$log.L('Destino.: ' + $destino)
Copy-Item -LiteralPath $Backup -Destination $destino -Force

try { $cn = Open-Banco $destino }
catch { $log.L('ERRO: o backup NÃO abriu - este backup não serve. ' + $_.Exception.Message); $log.Salvar(); exit 1 }

try {
    $log.L('')
    $log.L('--- o que voltou ---')
    foreach ($t in @('clientes', 'contatos', 'oportunidades', 'listas', 'log_alteracoes')) {
        $log.L('  ' + $t.PadRight(18) + (Invoke-Escalar $cn ('SELECT COUNT(*) FROM ' + $t) @()))
    }
    $log.L('  ' + 'soma de valor'.PadRight(18) + (Invoke-Escalar $cn 'SELECT SUM(valor) FROM oportunidades' @()))
    $log.L('  ' + 'esquema'.PadRight(18) + (Get-VersaoEsquemaBanco $cn))
    $log.L('')
    $log.L('--- relações preservadas ---')
    $log.L('  contatos órfãos ......: ' + (Invoke-Escalar $cn 'SELECT COUNT(*) FROM contatos WHERE id_cliente NOT IN (SELECT id FROM clientes)' @()))
    $log.L('  atendimentos órfãos ..: ' + (Invoke-Escalar $cn 'SELECT COUNT(*) FROM oportunidades WHERE id_contato NOT IN (SELECT id FROM contatos)' @()))
} finally { $cn.Close(); $cn.Dispose() }

$log.L('')
$log.L('Compare com o esperado. Se bater, o backup serve.')
$log.L('A cópia de teste pode ser apagada depois de conferida: ela não é backup nem produção.')
$log.Salvar()
