<#
  restaurar-banco.ps1 - restaura um backup numa CÓPIA DE TESTE e confere o que voltou.

  Backup que nunca foi restaurado não é backup. Este script nunca sobrescreve a produção e nunca abre o
  arquivo de backup: copia o backup para execucao\backup\restauracao-teste-*.accdb e abre só essa cópia.

  Escolha do backup (sem -Backup, ciclo 018): a última cópia VALIDADA (com .validado e manifesto, hash
  conferido); cópia simples só com -Backup. "Última" pelo manifesto, ou pelo NOME do arquivo
  (crm_<tipo>_AAAAMMDD-HHMM_<estação>.accdb); o LastWriteTime só entra em último caso, porque
  Copy-Item preserva a data do banco original.

  Conferência:
    - com manifesto (<backup>.manifesto.json): SHA-256, esquema e contagens têm de bater;
    - com -Referencia <banco.accdb>: as contagens da cópia restaurada são comparadas com as do banco de
      referência (lido em modo somente leitura). Se a referência está em uso, a divergência pode ser só
      gravação posterior ao backup: leia a mensagem.
  Código de saída: 0 tudo conferiu, 1 divergência ou backup que não abre, 2 sem backup para restaurar.

  Uso: restaurar-banco.ps1 [-Backup <arquivo.accdb>] [-Referencia <banco.accdb>]
#>
param([string]$Backup = '', [string]$Referencia = '')
$ErrorActionPreference = 'Stop'
$dadosLib = Join-Path (Split-Path -Parent $PSScriptRoot) '..\lib'
. (Join-Path $dadosLib 'Comum.ps1')
. (Join-Path $dadosLib 'AtualizacaoBanco.ps1')

$raiz = Get-RaizBase $PSScriptRoot
$pastaBkp = Get-PastaExecucao $raiz 'backup'
$log = New-Log $raiz 'restauracao'

if ([string]::IsNullOrWhiteSpace($Backup)) {
    # DEV-03 (ciclo 018): sem -Backup, a última cópia VALIDADA (completa, com manifesto e hash conferido);
    # a cópia simples mais nova, não verificada, só com -Backup explícito
    $ultimo = Get-BackupValidadoMaisRecente $pastaBkp
    if ($null -eq $ultimo) {
        $qualquer = Get-BackupMaisRecente $pastaBkp
        if ($null -eq $qualquer) { $log.L('Nenhum backup encontrado em ' + $pastaBkp + '. Gere um com FAZER-BACKUP.bat.') }
        else { $log.L('Nenhuma cópia validada em ' + $pastaBkp + '. A mais recente (' + $qualquer.Name + ') não foi validada: para restaurá-la mesmo assim, informe -Backup.') }
        $log.Salvar(); exit 2
    }
    $log.L('Sem -Backup: a última cópia validada é ' + $ultimo.Name + '.')
    $Backup = $ultimo.FullName
}
if (-not (Test-Path -LiteralPath $Backup)) { $log.L('Backup não encontrado: ' + $Backup); $log.Salvar(); exit 2 }
$destino = Join-Path $pastaBkp ('restauracao-teste-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.accdb')
$divergencias = New-Object System.Collections.ArrayList

$log.L('=== TESTE DE RESTAURAÇÃO ===')
$log.L('Backup..: ' + (Split-Path -Leaf $Backup) + '  (instante ' + (Get-InstanteDoBackup (Get-Item -LiteralPath $Backup)) + ')')
$log.L('Destino.: ' + $destino)

# 1. manifesto (SHA-256, esquema e contagens), numa cópia descartável
$manifesto = $Backup + '.manifesto.json'
if (Test-Path -LiteralPath $manifesto) {
    $val = Test-BackupRestauravel $Backup
    if ($val.ok) { $log.L('Manifesto: SHA-256, esquema ' + $val.esquema + ' e contagens conferem.') }
    else { foreach ($p in $val.problemas) { [void]$divergencias.Add('manifesto: ' + $p) } }
} else {
    $log.L('Sem manifesto (backup anterior ao manifesto): só a conferência de abertura e, se pedida, a da referência.')
}

# 2. cópia de teste + o que voltou
Copy-Item -LiteralPath $Backup -Destination $destino -Force
try { $cn = Open-BancoLeitura $destino }
catch { $log.L('ERRO: o backup NÃO abriu - este backup não serve. ' + $_.Exception.Message); $log.Salvar(); exit 1 }
try {
    $contagens = Get-ContagensDoBanco $cn
    $esquema = Get-VersaoEsquemaBanco $cn
    $log.L('')
    $log.L('--- o que voltou ---')
    foreach ($k in $contagens.Keys) { $log.L('  ' + ([string]$k).PadRight(20) + $contagens[$k]) }
    $log.L('  ' + 'esquema'.PadRight(20) + $esquema)
    $log.L('')
    $log.L('--- relações preservadas ---')
    $orf1 = [int](Invoke-Escalar $cn 'SELECT COUNT(*) FROM contatos WHERE id_cliente NOT IN (SELECT id FROM clientes)' @())
    $orf2 = [int](Invoke-Escalar $cn 'SELECT COUNT(*) FROM oportunidades WHERE id_contato NOT IN (SELECT id FROM contatos)' @())
    $log.L('  contatos órfãos ......: ' + $orf1)
    $log.L('  atendimentos órfãos ..: ' + $orf2)
    if ($orf1 -gt 0 -or $orf2 -gt 0) { [void]$divergencias.Add('há registros órfãos na cópia restaurada') }
} finally { $cn.Close(); $cn.Dispose() }

# 3. referência
if ($Referencia) {
    $cr = Open-BancoLeitura $Referencia
    try { $contRef = Get-ContagensDoBanco $cr; $esqRef = Get-VersaoEsquemaBanco $cr } finally { $cr.Close(); $cr.Dispose() }
    $log.L('')
    $log.L('--- comparação com a referência ---')
    if ($esqRef -ne $esquema) { [void]$divergencias.Add('esquema: backup ' + $esquema + ', referência ' + $esqRef) }
    foreach ($k in $contRef.Keys) {
        $obtido = if ($contagens.Contains($k)) { [int]$contagens[$k] } else { -1 }
        if ($obtido -ne [int]$contRef[$k]) { [void]$divergencias.Add($k + ': backup ' + $obtido + ', referência ' + $contRef[$k]) }
    }
    if ($divergencias.Count -eq 0) { $log.L('  contagens e esquema iguais aos da referência.') }
}

$log.L('')
if ($divergencias.Count -gt 0) {
    $log.L('DIVERGÊNCIA - o backup NÃO foi aprovado:')
    foreach ($d in $divergencias) { $log.L('  - ' + $d) }
    $log.L('A cópia de teste pode ser apagada: ela não é backup nem produção.')
    $log.Salvar(); exit 1
}
$log.L('RESTAURAÇÃO APROVADA. A cópia de teste pode ser apagada depois de conferida: ela não é backup nem produção.')
$log.Salvar()
exit 0
