<#
  conferir-painel.ps1 - confere os indicadores do Painel contra o banco.

  Le a aba BasePainel do .xlsm (o que o modPainel gravou) e roda as MESMAS
  agregacoes direto no banco, comparando valor a valor. Conferir o Painel
  olhando para o proprio Painel nao confere nada.

  Rode DEPOIS de clicar em "Atualizar indicadores" na sua copia local.
#>
param(
    [string]$Arquivo = "",
    [string]$Banco = ""
)
$ErrorActionPreference = 'Stop'
. (Join-Path (Split-Path -Parent $PSScriptRoot) 'lib\Comum.ps1')
$raizBase = Get-RaizBase $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($Banco)) { $Banco = Get-CaminhoBanco $raizBase }
$log = Join-Path (Get-PastaExecucao $raizBase 'logs') ('painel-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.txt')
$saida = New-Object System.Collections.ArrayList
function L($t) { [void]$saida.Add([string]$t); Write-Host $t }
function Fim { $saida | Set-Content -Path $log -Encoding UTF8; Write-Host ""; Write-Host ("Log: " + $log) }

L "=== CONFERENCIA DO PAINEL ==="
L ("Data....: " + (Get-Date -Format 'dd/MM/yyyy HH:mm:ss'))
L ("Banco...: " + $Banco)
L ""

# ---------- localizar a copia local ----------
if ([string]::IsNullOrWhiteSpace($Arquivo) -or -not (Test-Path $Arquivo)) {
    $candidatos = @(
        (Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'CRM Zapromaq\CRM_Zapromaq.xlsm'),
        (Join-Path ([Environment]::GetFolderPath('Desktop')) 'CRM_Zapromaq.xlsm'),
        (Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'CRM_Zapromaq.xlsm'),
        (Join-Path $env:USERPROFILE 'Downloads\CRM_Zapromaq.xlsm'),
        "C:\CRM\CRM_Zapromaq.xlsm"
    )
    $achado = $null
    foreach ($c in $candidatos) { if (Test-Path $c) { $achado = $c; break } }

    if ($null -eq $achado) {
        # procura no perfil do usuario, sem varrer o disco inteiro
        try {
            $achado = (Get-ChildItem -Path $env:USERPROFILE -Filter 'CRM_Zapromaq.xlsm' -File -Recurse -ErrorAction SilentlyContinue |
                       Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName
        } catch { }
    }

    if ($null -eq $achado) {
        # ultimo recurso: pede o arquivo
        try {
            Add-Type -AssemblyName System.Windows.Forms
            $dlg = New-Object System.Windows.Forms.OpenFileDialog
            $dlg.Title = "Selecione sua copia local do CRM_Zapromaq.xlsm"
            $dlg.Filter = "Excel com macro (*.xlsm)|*.xlsm"
            $dlg.InitialDirectory = $env:USERPROFILE
            if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $achado = $dlg.FileName }
        } catch { }
    }

    if ($null -eq $achado) {
        L "ERRO: nao encontrei sua copia local do CRM_Zapromaq.xlsm."
        L ""
        L "Rode assim, apontando o caminho:"
        L '  powershell -ExecutionPolicy Bypass -File conferir-painel.ps1 -Arquivo "C:\caminho\CRM_Zapromaq.xlsm"'
        Fim; exit 1
    }
    $Arquivo = $achado
    L ("copia local encontrada: " + $Arquivo)
}
L ("Arquivo.: " + $Arquivo)
if (-not (Test-Path $Banco)) { L "ERRO: banco nao encontrado."; Fim; exit 1 }

# ---------- ler a aba BasePainel ----------
$xl = $null; $wb = $null
$doPainel = @{}
$atualizado = ""
try {
    $xl = New-Excel ([ref]$idExcel)
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $xl.EnableEvents = $false          # nao dispara Workbook_Open
    $wb = $xl.Workbooks.Open($Arquivo, 0, $true)   # somente leitura
    $ws = $wb.Worksheets.Item('BasePainel')
    $ultima = $ws.Cells.Item($ws.Rows.Count, 1).End(-4162).Row   # xlUp
    for ($i = 2; $i -le $ultima; $i++) {
        $bloco = [string]$ws.Cells.Item($i, 1).Value2
        $chave = [string]$ws.Cells.Item($i, 2).Value2
        $v1    = $ws.Cells.Item($i, 3).Value2
        if ($bloco -eq 'ATUALIZADO_EM') { $atualizado = [string]$chave; continue }
        if ($bloco -eq '' -or $bloco -eq 'ordem_etapa') { continue }   # ordem da lista, nao e indicador
        $doPainel[($bloco + '|' + $chave)] = $v1
    }
    $wb.Close($false); $xl.Quit()
    L ("linhas lidas da aba BasePainel: " + $doPainel.Count)
} catch {
    L ("ERRO ao ler a planilha: " + $_.Exception.Message.Split([char]13)[0])
    if ($null -ne $wb) { try { $wb.Close($false) } catch { } }
    if ($null -ne $xl) { try { $xl.Quit() } catch { } }
    Fim; exit 1
}
if ($doPainel.Count -eq 0) {
    L ""
    L "A aba BasePainel esta vazia. Clique em 'Atualizar indicadores' na sua copia"
    L "local antes de rodar esta conferencia."
    Fim; exit 1
}

# ---------- consultar o banco ----------
$provider = $null
foreach ($prov in @('Microsoft.ACE.OLEDB.16.0','Microsoft.ACE.OLEDB.12.0')) {
    try { $t = New-Object System.Data.OleDb.OleDbConnection("Provider=$prov;Data Source=$Banco;"); $t.Open(); $t.Close(); $t.Dispose(); $provider = $prov; break } catch { }
}
if (-not $provider) { L "ERRO: nenhum provedor ACE."; Fim; exit 1 }
$cn = New-Object System.Data.OleDb.OleDbConnection("Provider=$provider;Data Source=$Banco;")
$cn.Open()
$abertas = " AND op.etapa NOT IN ('Pedido Fechado','Perdido','Descartado')"

function Consultar($sql) {
    $cmd = $cn.CreateCommand(); $cmd.CommandText = $sql
    $rd = $cmd.ExecuteReader()
    $r = @{}
    while ($rd.Read()) {
        $k = if ($rd.IsDBNull(0)) { "" } else { [string]$rd.GetValue(0) }
        $r[$k] = $rd.GetValue(1)
    }
    $rd.Close(); $cmd.Dispose()
    return $r
}
function Escalar($sql) {
    $cmd = $cn.CreateCommand(); $cmd.CommandText = $sql
    $v = $cmd.ExecuteScalar(); $cmd.Dispose()
    if ($null -eq $v -or $v -is [System.DBNull]) { return 0 }
    return $v
}

$doBanco = @{}
foreach ($t in @(
  @{b='total'; k='empresas';     s="SELECT COUNT(*) FROM clientes"},
  @{b='total'; k='clientes';     s="SELECT COUNT(*) FROM clientes WHERE estagio='Cliente'"},
  @{b='total'; k='preclientes';  s="SELECT COUNT(*) FROM clientes WHERE estagio='Pré-cliente'"},
  @{b='total'; k='contatos';     s="SELECT COUNT(*) FROM contatos"},
  @{b='total'; k='atendimentos'; s="SELECT COUNT(*) FROM oportunidades"},
  @{b='total'; k='abertos';      s="SELECT COUNT(*) FROM oportunidades op WHERE 1=1$abertas"},
  @{b='total'; k='valor_aberto'; s="SELECT SUM(op.valor) FROM oportunidades op WHERE 1=1$abertas"},
  @{b='total'; k='valor_ganho';  s="SELECT SUM(valor) FROM oportunidades WHERE etapa='Pedido Fechado'"}
)) { $doBanco[($t.b + '|' + $t.k)] = Escalar $t.s }

foreach ($par in @(
  @{b='etapa';       s="SELECT op.etapa, COUNT(*) FROM oportunidades op GROUP BY op.etapa"},
  @{b='responsavel'; s="SELECT op.responsavel, COUNT(*) FROM oportunidades op WHERE 1=1$abertas GROUP BY op.responsavel"},
  @{b='familia';     s="SELECT op.familia, COUNT(*) FROM oportunidades op WHERE 1=1$abertas GROUP BY op.familia"},
  @{b='categoria';   s="SELECT op.categoria, COUNT(*) FROM oportunidades op WHERE 1=1$abertas GROUP BY op.categoria"},
  @{b='origem';      s="SELECT op.origem, COUNT(*) FROM oportunidades op GROUP BY op.origem"},
  @{b='segmento';    s="SELECT cl.segmento, COUNT(*) FROM oportunidades op LEFT JOIN clientes cl ON op.id_cliente=cl.id WHERE 1=1$abertas GROUP BY cl.segmento"},
  @{b='mes_entrada'; s="SELECT YEAR(op.dt_entrada)*100+MONTH(op.dt_entrada), COUNT(*) FROM oportunidades op WHERE op.dt_entrada IS NOT NULL GROUP BY YEAR(op.dt_entrada)*100+MONTH(op.dt_entrada)"}
)) {
    $r = Consultar $par.s
    foreach ($k in $r.Keys) { $doBanco[($par.b + '|' + $k)] = $r[$k] }
}
$cn.Close(); $cn.Dispose()

# ---------- comparar ----------
L ""
L ("indicadores atualizados na planilha em: " + $atualizado)
L ""
L "--- comparacao ---"
$dif = 0; $ok = 0; $soNoPainel = 0; $soNoBanco = 0

foreach ($chave in ($doBanco.Keys | Sort-Object)) {
    $vb = $doBanco[$chave]
    if (-not $doPainel.ContainsKey($chave)) {
        $soNoBanco++
        L ("  FALTA NO PAINEL  " + $chave.PadRight(46) + "banco: " + $vb)
        continue
    }
    $vp = $doPainel[$chave]
    $iguais = $false
    if ($vb -is [double] -or $vb -is [decimal] -or $vp -is [double]) {
        $iguais = ([math]::Abs([double]$vb - [double]$vp) -lt 0.005)
    } else {
        $iguais = ([string]$vb -eq [string]$vp)
    }
    if ($iguais) { $ok++ }
    else {
        $dif++
        L ("  DIVERGENTE       " + $chave.PadRight(46) + "painel: " + $vp + "   banco: " + $vb)
    }
}
foreach ($chave in ($doPainel.Keys | Sort-Object)) {
    if (-not $doBanco.ContainsKey($chave)) {
        $soNoPainel++
        L ("  SO NO PAINEL     " + $chave.PadRight(46) + "painel: " + $doPainel[$chave])
    }
}

L ""
L ("conferem: " + $ok + "   divergentes: " + $dif + "   so no banco: " + $soNoBanco + "   so no painel: " + $soNoPainel)
L ""
if ($dif -eq 0 -and $soNoBanco -eq 0) {
    L "PAINEL OK - os indicadores conferidos batem com o banco."
    L "(mes_desfecho, fila e metas agrupam por mais de uma coluna e ficam de"
    L " fora desta comparacao; confira na tela se quiser fechar 100%)"
} else {
    L "CONFERIR - ha divergencia. Divergencia aqui e divergencia de REGRA,"
    L "nao de tela: o numero que o vendedor ve esta errado."
    Fim; exit 1
}
Fim
