<#
  testar-uso.ps1 - teste de fumaça do front montado, sem interação.

  Abre uma CÓPIA do .xlsm de teste (eventos desligados: o Workbook_Open com
  MsgBox não roda), monta as três grades e o Painel pelas próprias macros e
  confere o resultado: linhas carregadas, status, Painel preenchido. Exporta
  as abas em PDF (execucao\teste\uso-*.pdf) para conferência visual.
  Um vigia encerra o Excel se alguma caixa de mensagem travar a execução.

  Uso: testar-uso.ps1 [-Arquivo <xlsm>]   (padrão: execucao\teste\CRM_Zapromaq.xlsm)
#>
param([string]$Arquivo = '', [switch]$SemRelatorio)
$ErrorActionPreference = 'Stop'
$dados = Split-Path -Parent $PSScriptRoot
. (Join-Path $dados 'lib\Comum.ps1')
. (Join-Path $dados 'build\lib\Excel.ps1')
$raiz = Get-RaizBase $PSScriptRoot
if (-not $Arquivo) { $Arquivo = Join-Path $dados 'execucao\teste\CRM_Zapromaq.xlsm' }
$pasta = Get-PastaExecucao $raiz 'teste'
$copia = Join-Path $env:TEMP ('crm-uso-' + $PID + '.xlsm')
Copy-Item -LiteralPath $Arquivo -Destination $copia -Force

$falhas = 0
function Conferir([string]$nome, [bool]$ok, [string]$det = '') {
    if ($ok) { Write-Host ('  OK    ' + $nome) } else { $script:falhas++; Write-Host ('  FALHA ' + $nome + '  ' + $det) }
}

$xl = $null; $idExcel = 0; $sinal = Join-Path $env:TEMP ('crm-uso-' + $PID + '.ok')
try {
    $xl = New-Excel ([ref]$idExcel)
    $xl.DisplayAlerts = $false
    $xl.EnableEvents = $false
    $vigia = Start-Vigia $idExcel 240 $sinal
    $wb = $xl.Workbooks.Open($copia)
    $n = $wb.Name
    foreach ($m in 'GradeClientes', 'GradeContatos', 'GradeOportunidades') {
        $ini = Get-Date
        $xl.Run("'" + $n + "'!modGrade." + $m)
        $ms = [int]((Get-Date) - $ini).TotalMilliseconds
        $aba = $m.Substring(5)
        $ws = $wb.Worksheets.Item($aba)
        $linhas = if ($ws.ListObjects.Count -gt 0 -and $ws.ListObjects.Item(1).DataBodyRange) { $ws.ListObjects.Item(1).DataBodyRange.Rows.Count } else { 0 }
        Conferir ($aba + ': ' + $linhas + ' linha(s) em ' + $ms + ' ms') ($linhas -gt 0) ([string]$ws.Range('B5').Value2)
        Conferir ($aba + ': aba protegida') ([bool]$ws.ProtectContents)
        $ws.ExportAsFixedFormat(0, (Join-Path $pasta ('uso-' + $aba + '.pdf')))
    }
    $ini = Get-Date
    $xl.Run("'" + $n + "'!modPainel.AtualizarBasePainel")
    $ms = [int]((Get-Date) - $ini).TotalMilliseconds
    $p = $wb.Worksheets.Item('Painel')
    $tempos = [string]$xl.Run("'" + $n + "'!modPainel.TemposUltimaAtualizacao")
    Write-Host ('  tempos do Painel: ' + $tempos)
    # sem a marca final, o desenho quebrou no meio (e a caixa de erro prendeu o Excel)
    Conferir 'Painel desenhado até o fim' ($tempos -like '*paineis de baixo=*') $tempos
    Conferir 'Painel em menos de 5 s' ($ms -lt 5000) ($ms.ToString() + ' ms')
    Conferir ('Painel desenhado em ' + $ms + ' ms') ([string]$p.Range('B3').Value2 -like 'Atualizado em*') ([string]$p.Range('B3').Value2)
    Conferir ('Painel: cartão Empresas = ' + $p.Range('B6').Value2) ($p.Range('B6').Value2 -gt 0)
    Conferir ('Painel: ' + $p.ChartObjects().Count + ' gráfico(s)') ($p.ChartObjects().Count -ge 2)
    # "Ver tudo" com filtro e pesquisa ativos: as linhas ficam ocultas e a
    # grade ja saiu com o cabecalho da consulta anterior e colunas #N/D
    $ws = $wb.Worksheets.Item('Oportunidades')
    $lo = $ws.ListObjects.Item(1)
    $ws.Unprotect('zpm')
    $lo.Range.AutoFilter(3, [string]$lo.ListColumns.Item(3).DataBodyRange.Cells.Item(1, 1).Value2)
    $ws.Cells.Item(4, 3).Value2 = 'A'
    $ws.Protect('zpm')
    foreach ($passo in 'COMPLETO', 'MINIMO') {
        $xl.Run("'" + $n + "'!modGrade.GradeVerTudo")
        $cabs = @($wb.Worksheets.Item('Oportunidades').ListObjects.Item(1).ListColumns | ForEach-Object { $_.Name })
        $ruim = @($cabs | Where-Object { $_ -like '#N/D*' -or $_ -like 'Coluna*' -or $_ -like '*2' })
        Conferir ('Ver tudo (' + $passo + ') com filtro e pesquisa') ($ruim.Count -eq 0) ($cabs -join '|')
    }

    # relatorio: pasta com CSVs, PDF e LEIA-ME (no build, -SemRelatorio pula:
    # publicar nao pode criar pasta de relatorio em Documentos sem o usuario pedir)
    if (-not $SemRelatorio) {
    $pastaRel = [string]$xl.Run("'" + $n + "'!modRelatorio.GerarRelatorio", $true)
    foreach ($a in 'indicadores.csv', 'movimento_mensal.csv', 'atendimentos.csv', 'clientes.csv', 'painel.pdf', 'LEIA-ME.txt') {
        $arq = Join-Path $pastaRel $a
        Conferir ('relatório: ' + $a) ((Test-Path -LiteralPath $arq) -and (Get-Item -LiteralPath $arq).Length -gt 0) $arq
    }
    Write-Host ('  relatório em ' + $pastaRel)
    }
    $p.PageSetup.Orientation = 2; $p.PageSetup.Zoom = $false; $p.PageSetup.FitToPagesWide = 1; $p.PageSetup.FitToPagesTall = 1
    $p.ExportAsFixedFormat(0, (Join-Path $pasta 'uso-Painel.pdf'))
    $wb.Close($false)
} catch {
    $falhas++
    Write-Host ('ERRO: ' + $_.Exception.Message)
    if (-not (Get-Process -Id $idExcel -ErrorAction SilentlyContinue)) { Write-Host '  (o Excel travou numa caixa de mensagem e foi encerrado pelo vigia)' }
} finally {
    if ($vigia) { Stop-Vigia $vigia $sinal }
    Close-Excel $xl $idExcel
    Remove-Item -LiteralPath $copia -Force -ErrorAction SilentlyContinue
}
Write-Host ''
if ($falhas -eq 0) { Write-Host 'TUDO CERTO. PDFs em execucao\teste\uso-*.pdf'; exit 0 }
Write-Host ($falhas.ToString() + ' falha(s).'); exit 1
