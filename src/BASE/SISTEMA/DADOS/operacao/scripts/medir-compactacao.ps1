<#
  medir-compactacao.ps1 - mede, NUMA CÓPIA, quanto compactar o banco ganharia (DEV-25, ciclo 014).

  Nunca toca o banco de origem: copia para execucao\medicoes\<instante>\, mede a cópia (tamanho e os tempos de
  abrir, buscar, gravar e atualizar - N repetições, mediana e pior caso), compacta a cópia pelo motor do Office
  (DAO.DBEngine.120.CompactDatabase), mede de novo e grava medicao.json + medicao.txt com a conclusão.
  As cópias são apagadas no fim (podem ter dado real): fica só o relatório.

  Critério de ganho (spec 014, decisão 014-1): tamanho -20% ou mediana de algum tempo -10% E pelo menos 5 ms
  (abaixo disso é ruído de medição). A primeira rodada de cada arquivo aquece o cache e é descartada. Sem ganho, a
  conclusão é "não adotar"; com ganho, "avaliar rotina" - a rotina em janela de manutenção não é implantada
  sem a medição no banco real.

  Banco em uso (.laccdb) é recusado: a cópia poderia sair inconsistente. Prefira medir a última cópia de
  backup validada (-Banco <execucao\backup\...accdb>). -Mesmo aceita o banco em uso (só para ensaio).

  Uso: medir-compactacao.ps1 [-Banco <accdb>] [-Repeticoes 7] [-Mesmo] [-Saida <pasta>]
#>
param([string]$Banco = '', [int]$Repeticoes = 7, [switch]$Mesmo, [string]$Saida = '')
$ErrorActionPreference = 'Stop'
$dadosLib = Join-Path (Split-Path -Parent $PSScriptRoot) '..\lib'
. (Join-Path $dadosLib 'Comum.ps1')

$raiz = Get-RaizBase $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($Banco)) { $Banco = Get-CaminhoBanco $raiz }
if (-not (Test-Path -LiteralPath $Banco)) { Write-Host ('Banco não encontrado: ' + $Banco); exit 1 }
if ($Repeticoes -lt 1 -or $Repeticoes -gt 50) { Write-Host 'Repeticoes vai de 1 a 50.'; exit 1 }
if ((Test-BancoEmUso $Banco 1000) -and -not $Mesmo) {
    Write-Host ('RECUSADO: o banco está em uso (.laccdb): a cópia poderia sair inconsistente. Meça a última cópia de backup ' +
                'validada (-Banco <arquivo>) ou rode com todos fora do CRM. Nada foi alterado.'); exit 2
}

if ([string]::IsNullOrWhiteSpace($Saida)) { $Saida = Get-PastaExecucao $raiz 'medicoes' }
$pasta = Join-Path $Saida ('compactacao-' + (Get-Date -Format 'yyyyMMdd-HHmmss'))
New-Item -ItemType Directory -Path $pasta -Force | Out-Null
$copia = Join-Path $pasta 'copia.accdb'
$compactado = Join-Path $pasta 'compactado.accdb'

function Get-Mediana([double[]]$v) {
    $o = @($v | Sort-Object); $n = $o.Count
    if ($n -eq 0) { return 0 }
    if ($n % 2 -eq 1) { return [double]$o[[int](($n - 1) / 2)] }
    return ([double]$o[$n / 2 - 1] + [double]$o[$n / 2]) / 2
}

# Tempo (ms) de cada operação, N vezes, sobre o arquivo indicado.
function Medir([string]$arquivo) {
    $t = @{ abrir = @(); buscar = @(); gravar = @(); atualizar = @() }
    $sw = New-Object System.Diagnostics.Stopwatch
    # rodada 0: aquece cache e provedor; não entra na conta
    for ($i = -1; $i -lt $Repeticoes; $i++) {
        $sw.Restart(); $cn = Open-Banco $arquivo; $sw.Stop(); $t.abrir += $sw.Elapsed.TotalMilliseconds
        try {
            # buscar: a consulta da grade de atendimentos (junções + LIKE em texto)
            $sw.Restart()
            $c = New-ComandoBanco $cn ('SELECT o.id, o.codigo, o.etapa, c.empresa, ct.nome FROM (oportunidades o LEFT JOIN clientes c ON o.id_cliente = c.id) ' +
                                       'LEFT JOIN contatos ct ON o.id_contato = ct.id WHERE c.empresa LIKE ? ORDER BY o.id') @((New-Parametro 'texto' '%A%')) $null
            $r = $c.ExecuteReader(); while ($r.Read()) { }; $r.Close(); $c.Dispose()
            $sw.Stop(); $t.buscar += $sw.Elapsed.TotalMilliseconds
            # gravar: uma alteração com log, numa transação desfeita (a cópia não muda de conteúdo)
            $sw.Restart()
            $tx = $cn.BeginTransaction()
            try {
                [void](Invoke-Comando $cn 'UPDATE config SET valor=valor WHERE chave=?' @((New-Parametro 'texto' 'ambiente')) $tx)
                [void](Invoke-Comando $cn ('INSERT INTO log_alteracoes (tabela, id_registro, campo, valor_antigo, valor_novo, acao, usuario, quando, origem) ' +
                                           'VALUES (?,?,?,?,?,?,?,?,?)') @((New-Parametro 'texto' 'config'), (New-Parametro 'inteiro' 0), (New-Parametro 'texto' 'medicao'),
                                           (New-Parametro 'memo' ''), (New-Parametro 'memo' ''), (New-Parametro 'texto' 'MEDIR'), (New-Parametro 'texto' 'medicao'),
                                           (New-Parametro 'data' (Get-Date)), (New-Parametro 'texto' 'FRONT')) $tx)
            } finally { $tx.Rollback() }
            $sw.Stop(); $t.gravar += $sw.Elapsed.TotalMilliseconds
            # atualizar: a rodada do feed (MAX do log e a faixa dos últimos 200)
            $sw.Restart()
            $max = Invoke-Escalar $cn 'SELECT MAX(id) FROM log_alteracoes' @() $null
            if ($max -is [System.DBNull] -or $null -eq $max) { $max = 0 }
            $c = New-ComandoBanco $cn 'SELECT tabela, id_registro, usuario, origem FROM log_alteracoes WHERE id > ? AND id <= ?' @((New-Parametro 'inteiro' ([int]$max - 200)), (New-Parametro 'inteiro' ([int]$max))) $null
            $r = $c.ExecuteReader(); while ($r.Read()) { }; $r.Close(); $c.Dispose()
            $sw.Stop(); $t.atualizar += $sw.Elapsed.TotalMilliseconds
        } finally { $cn.Close(); $cn.Dispose() }
    }
    foreach ($k in @('abrir', 'buscar', 'gravar', 'atualizar')) { $t[$k] = @($t[$k] | Select-Object -Skip 1) }
    $res = [ordered]@{ tamanho_bytes = (Get-Item -LiteralPath $arquivo).Length }
    foreach ($k in @('abrir', 'buscar', 'gravar', 'atualizar')) {
        $res[$k] = [ordered]@{ mediana_ms = [math]::Round((Get-Mediana ([double[]]$t[$k])), 1); pior_ms = [math]::Round(([double[]]$t[$k] | Measure-Object -Maximum).Maximum, 1) }
    }
    return $res
}

function Get-Ganho([double]$antes, [double]$depois) {
    if ($antes -le 0) { return 0 }
    return [math]::Round((($antes - $depois) / $antes) * 100, 1)
}

$hashOrigem = (Get-FileHash -LiteralPath $Banco -Algorithm SHA256).Hash
$codigo = 0
try {
    # cópia lida sem bloquear o arquivo de origem
    $de = [System.IO.File]::Open($Banco, 'Open', 'Read', 'ReadWrite')
    try { $para = [System.IO.File]::Create($copia); try { $de.CopyTo($para) } finally { $para.Close() } } finally { $de.Close() }

    $antes = Medir $copia
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $dao = New-Object -ComObject DAO.DBEngine.120
    try { $dao.CompactDatabase($copia, $compactado) } finally { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($dao) }
    $sw.Stop()
    $depois = Medir $compactado

    $ganhos = [ordered]@{ tamanho_pct = (Get-Ganho $antes.tamanho_bytes $depois.tamanho_bytes) }
    foreach ($k in @('abrir', 'buscar', 'gravar', 'atualizar')) { $ganhos[$k + '_mediana_pct'] = (Get-Ganho $antes[$k].mediana_ms $depois[$k].mediana_ms) }
    $comGanho = ($ganhos.tamanho_pct -ge 20)
    foreach ($k in @('abrir', 'buscar', 'gravar', 'atualizar')) {
        if ($ganhos[$k + '_mediana_pct'] -ge 10 -and ($antes[$k].mediana_ms - $depois[$k].mediana_ms) -ge 5) { $comGanho = $true }
    }
    $conclusao = if ($comGanho) { 'avaliar rotina: há ganho medido; a rotina em janela de manutenção só depois da medição no banco real (homologação)' }
                 else { 'não adotar: sem ganho medido (tamanho < 20% e nenhum tempo com mediana 10% e 5 ms menor)' }

    $rel = [ordered]@{
        tipo = 'medicao-compactacao'; criado_em = (Get-Date -Format 'yyyy-MM-ddTHH:mm:ss'); estacao = [Environment]::MachineName
        origem = $Banco; origem_sha256 = $hashOrigem; repeticoes = $Repeticoes; compactar_ms = [math]::Round($sw.Elapsed.TotalMilliseconds, 1)
        antes = $antes; depois = $depois; ganhos = $ganhos; criterio = 'tamanho -20% ou mediana de algum tempo -10% e -5 ms'
        com_ganho = $comGanho; conclusao = $conclusao
    }
    [System.IO.File]::WriteAllText((Join-Path $pasta 'medicao.json'), (ConvertTo-Json -InputObject $rel -Depth 5), (New-Object System.Text.UTF8Encoding($false)))
    $txt = @('MEDIÇÃO DA COMPACTAÇÃO - ' + $rel.criado_em + ' - ' + $rel.estacao, 'Origem: ' + $Banco, '',
             ('Tamanho: ' + [math]::Round($antes.tamanho_bytes / 1MB, 2) + ' MB -> ' + [math]::Round($depois.tamanho_bytes / 1MB, 2) + ' MB (' + $ganhos.tamanho_pct + '% menor); compactar levou ' + $rel.compactar_ms + ' ms'))
    foreach ($k in @('abrir', 'buscar', 'gravar', 'atualizar')) {
        $txt += ($k + ': mediana ' + $antes[$k].mediana_ms + ' -> ' + $depois[$k].mediana_ms + ' ms (' + $ganhos[$k + '_mediana_pct'] + '%); pior ' + $antes[$k].pior_ms + ' -> ' + $depois[$k].pior_ms + ' ms')
    }
    $txt += @('', 'Critério: ' + $rel.criterio, 'CONCLUSÃO: ' + $conclusao)
    [System.IO.File]::WriteAllLines((Join-Path $pasta 'medicao.txt'), [string[]]$txt, (New-Object System.Text.UTF8Encoding($true)))
    $txt | ForEach-Object { Write-Host $_ }
    Write-Host ('Relatório: ' + $pasta)
} catch {
    Write-Host ('ERRO: ' + $_.Exception.Message); $codigo = 1
} finally {
    foreach ($f in @($copia, $compactado, [System.IO.Path]::ChangeExtension($copia, 'laccdb'), [System.IO.Path]::ChangeExtension($compactado, 'laccdb'))) {
        if (Test-Path -LiteralPath $f) { Remove-Item -LiteralPath $f -Force -ErrorAction SilentlyContinue }
    }
}
if (-not $Mesmo -and (Get-FileHash -LiteralPath $Banco -Algorithm SHA256).Hash -ne $hashOrigem) { Write-Host 'ERRO: o banco de origem mudou durante a medição.'; exit 1 }
exit $codigo
