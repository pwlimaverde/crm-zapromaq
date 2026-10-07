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
    # item 3: Clientes e Contatos abrem em Ativos; Todos mostra os inativos (em vermelho);
    # Limpar filtros volta para Ativos. Célula de situação: F4 (modGrade.COL_SITUACAO).
    foreach ($aba in 'Clientes', 'Contatos') {
        $ws = $wb.Worksheets.Item($aba); $lo = $ws.ListObjects.Item(1)
        $total = $lo.DataBodyRange.Rows.Count
        $vis = [int]$xl.WorksheetFunction.Subtotal(103, $lo.ListColumns.Item('_id').DataBodyRange)
        Conferir ($aba + ': abre em Ativos (' + $vis + ' de ' + $total + ')') ($vis -lt $total -and [string]$ws.Range('F4').Value2 -eq 'Ativos') ('situação=' + [string]$ws.Range('F4').Value2 + '; banco de demonstração sem inativos? rode testes\criar-banco-demo.ps1 -Recriar')
        $ws.Range('F4').Value2 = 'Todos'
        $xl.Run("'" + $n + "'!modGrade.FiltrarSituacao", $ws)
        $vis = [int]$xl.WorksheetFunction.Subtotal(103, $lo.ListColumns.Item('_id').DataBodyRange)
        Conferir ($aba + ': Todos mostra os ' + $total) ($vis -eq $total) ([string]$vis)
        $regras = 0
        foreach ($fc in $lo.DataBodyRange.FormatConditions) { if ([string]$fc.Formula1 -like '*INATIVO*') { $regras++ } }
        Conferir ($aba + ': regra vermelha dos inativos') ($regras -ge 1)
        $ws.Activate()
        $xl.Run("'" + $n + "'!modGrade.GradeLimparFiltros")
        $vis = [int]$xl.WorksheetFunction.Subtotal(103, $lo.ListColumns.Item('_id').DataBodyRange)
        Conferir ($aba + ': Limpar filtros volta para Ativos') ($vis -lt $total -and [string]$ws.Range('F4').Value2 -eq 'Ativos') ([string]$vis)
    }

    # item 6: os dados do contato (troca na ficha) vêm numa consulta só
    $lo = $wb.Worksheets.Item('Contatos').ListObjects.Item(1)
    $idCto = [int]$lo.ListColumns.Item('_id').DataBodyRange.Cells.Item(1, 1).Value2
    $nomeCto = [string]$lo.ListColumns.Item('Contato').DataBodyRange.Cells.Item(1, 1).Value2
    $dc = $xl.Run("'" + $n + "'!modCRM.DadosDoContato", $idCto)
    Conferir ('DadosDoContato(' + $idCto + ') traz nome e empresa') ($null -ne $dc -and [string]$dc.Item('nome') -eq $nomeCto -and [int]$dc.Item('id_cliente') -gt 0) $nomeCto

    # item 4: lista do atendimento anterior = os da mesma empresa, mais recentes primeiro,
    # sem o próprio. No banco de demonstração o cliente 1 tem 3; o id 1 é o mais recente.
    $la = $xl.Run("'" + $n + "'!modCRM.ListarAtendimentosDoCliente", 1, 0, '')
    $qt = if ($la -is [array]) { $la.GetLength(0) } else { 0 }
    Conferir ('anteriores do cliente 1: ' + $qt + ', o mais recente primeiro') ($qt -ge 2 -and [int]$la.GetValue(1, 1) -eq 1)
    $lb = $xl.Run("'" + $n + "'!modCRM.ListarAtendimentosDoCliente", 1, 1, '')
    $ids = if ($lb -is [array]) { @(for ($k = 1; $k -le $lb.GetLength(0); $k++) { [int]$lb.GetValue($k, 1) }) } else { @() }
    Conferir 'anteriores sem o próprio atendimento' ($ids.Count -eq $qt - 1 -and $ids -notcontains 1) ($ids -join ',')

    # item 5: lote fictício com uma linha de cada estado (nada é gravado: só Analisar)
    #   1 OK; 2 SUSPEITO (nome de empresa da base, sem CNPJ); 3 DUPLICADO (CNPJ da linha 1); 4 ERRO (3 colunas)
    $t = "`t"
    $lote = @(('2' + $t + '2' + $t + $t + 'LOTE TESTE UM LTDA' + $t + 'CURITIBA' + $t + 'PR' + $t + $t + $t + $t + '11222333000181' + $t),
              ('2' + $t + '2' + $t + $t + 'METALURGICA EXEMPLO' + $t + 'CURITIBA' + $t + 'PR' + $t + $t + $t + $t + $t),
              ('1' + $t + '1' + $t + $t + 'LOTE TESTE DOIS LTDA' + $t + 'CURITIBA' + $t + 'PR' + $t + $t + $t + $t + '11222333000181' + $t),
              ('1' + $t + '1' + $t + 'SO TRES COLUNAS')) -join "`r`n"
    $rl = "'" + $n + "'!modLote."
    $q = [int]$xl.Run($rl + 'Analisar', $lote)
    # foreach, não ForEach-Object: o $_ do pipeline chega ao COM embrulhado (PSObject) e o Run trava
    $est = @(foreach ($i in 1..4) { [string]$xl.Run($rl + 'Valor', $i, 'status') }) -join ','
    Conferir ('lote: 4 linhas nos 4 estados (' + $est + ')') ($q -eq 4 -and $est -eq 'OK,SUSPEITO,DUPLICADO,ERRO') $est
    Conferir 'lote: total inicial conta só a OK' ([int]$xl.Run($rl + 'ACadastrar') -eq 1) ([string]$xl.Run($rl + 'ACadastrar'))
    $ok1 = [bool]$xl.Run($rl + 'AlternarIgnorar', 1)
    Conferir 'lote: duplo clique na OK tira a linha' ($ok1 -and [int]$xl.Run($rl + 'ACadastrar') -eq 0)
    $ok2 = [bool]$xl.Run($rl + 'AlternarIgnorar', 2)
    Conferir 'lote: duplo clique na SUSPEITO inclui' ($ok2 -and [int]$xl.Run($rl + 'ACadastrar') -eq 1 -and -not [bool]$xl.Run($rl + 'Ignorada', 2))
    Conferir 'lote: DUPLICADO e ERRO não alternam' (-not [bool]$xl.Run($rl + 'AlternarIgnorar', 3) -and -not [bool]$xl.Run($rl + 'AlternarIgnorar', 4))

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
