<#
  testar-migracao-003.ps1 - aplica a migração 003 numa CÓPIA do banco de teste e confere o resultado.

  O banco de teste é criado em arquivo temporário (criar-banco-teste.ps1 -Banco) e a migração roda numa
  cópia dele; nada da pasta é alterado. Casos C17, C18, papel das etapas, índice único,
  idempotência e relatório (spec 003, Casos de regra e Critérios de aceite da Fase 1).
#>
param()
$ErrorActionPreference = 'Stop'
$raizIA = Split-Path -Parent $PSScriptRoot
$raizBase = Split-Path -Parent $raizIA
$origem = Join-Path ([System.IO.Path]::GetTempPath()) ('migracao-003-origem-' + [guid]::NewGuid().ToString('N') + '.accdb')
$pastaMig = Join-Path $raizBase 'SISTEMA\DADOS\banco\migracoes'
. (Join-Path $raizBase 'SISTEMA\DADOS\lib\Comum.ps1')

$script:falhas = 0
function Conferir([string]$nome, [bool]$ok, [string]$detalhe = '') {
    if ($ok) { Write-Host ('  OK    ' + $nome) -ForegroundColor Green }
    else { $script:falhas++; Write-Host ('  FALHA ' + $nome + '  ' + $detalhe) -ForegroundColor Red }
}
function NovoLog {
    $l = [pscustomobject]@{ linhas = New-Object System.Collections.ArrayList }
    $l | Add-Member -MemberType ScriptMethod -Name L -Value { param($t) [void]$this.linhas.Add([string]$t) }
    return $l
}
# Consulta: uma linha = um objeto com as colunas pelo nome (valor, encerra...); nulo vira $null
function Consulta($cn, [string]$sql, [array]$params = @()) {
    $c = New-ComandoBanco $cn $sql $params $null
    $r = $c.ExecuteReader()
    $saida = New-Object System.Collections.ArrayList
    while ($r.Read()) {
        $linha = [ordered]@{}
        for ($i = 0; $i -lt $r.FieldCount; $i++) {
            $v = $r.GetValue($i)
            if ($v -is [DBNull]) { $v = $null }
            $linha[$r.GetName($i).ToLowerInvariant()] = $v
        }
        [void]$saida.Add([pscustomobject]$linha)
    }
    $r.Close(); $c.Dispose()
    return , $saida
}

# banco de teste próprio (com os casos da 003), criado em arquivo temporário: não mexe no banco da pasta
& (Join-Path $PSScriptRoot 'criar-banco-teste.ps1') -Banco $origem -Recriar -CasosDaMigracao | Out-Null
$copia = Join-Path ([IO.Path]::GetTempPath()) ('migracao-003-' + [guid]::NewGuid().ToString('N') + '.accdb')
Copy-Item -LiteralPath $origem -Destination $copia
$cn = $null
try {
    $m = & (Join-Path $pastaMig '003-municipios-visitas-papel-listas.ps1')
    $cn = Open-Banco $copia
    if ((Get-VersaoEsquemaBanco $cn) -ne '1.0') { throw 'a cópia de teste não está no esquema 1.0' }

    Write-Host '--- aplicação da migração 003 na cópia'
    $log = NovoLog
    & $m.Aplicar $cn $null $log $pastaMig
    Conferir 'municípios carregados (5571)' ([int](Invoke-Escalar $cn 'SELECT COUNT(*) FROM municipios' @() $null) -eq 5571)
    Conferir 'nenhuma chave de lista vazia' ([int](Invoke-Escalar $cn 'SELECT COUNT(*) FROM listas WHERE chave IS NULL' @() $null) -eq 0)

    $etapas = Consulta $cn "SELECT valor, ativo, encerra, resultado, quadro, exige_familia FROM listas WHERE tipo='Etapa'"
    $valores = @($etapas | ForEach-Object { [string]$_.valor })

    Write-Host '--- C17: variante unificada na mais usada'
    Conferir 'PEDIDO FECHADO saiu da lista (comparação exata)' (-not ($valores -ccontains 'PEDIDO FECHADO'))
    Conferir 'Pedido Fechado continua, uma vez só' (@($valores | Where-Object { $_ -ceq 'Pedido Fechado' }).Count -eq 1)
    $atendimentos = Consulta $cn 'SELECT etapa FROM oportunidades'
    $etapasOpo = @($atendimentos | ForEach-Object { [string]$_.etapa } | Where-Object { $_ -ieq 'pedido fechado' })
    Conferir 'atendimentos com a variante agora usam Pedido Fechado' (@($etapasOpo | Where-Object { $_ -cne 'Pedido Fechado' }).Count -eq 0) ($etapasOpo -join ' / ')
    Conferir 'há atendimentos com a etapa Pedido Fechado (teste não é vácuo)' ($etapasOpo.Count -ge 3) ''
    Conferir 'unificação registrada no log' ([int](Invoke-Escalar $cn "SELECT COUNT(*) FROM log_alteracoes WHERE acao='UNIFICACAO'" @() $null) -gt 0)

    Write-Host '--- C18: valor usado sem linha entra inativo'
    $orfao = @($etapas | Where-Object { $_.valor -ceq 'Qualificando Antigo' })
    Conferir 'Qualificando Antigo existe como linha' ($orfao.Count -eq 1)
    Conferir 'Qualificando Antigo entra inativo' ($orfao.Count -eq 1 -and -not [bool]$orfao[0].ativo)

    Write-Host '--- etapa especial ausente entra inativa com o papel dela'
    $perdido = @($etapas | Where-Object { $_.valor -ceq 'Perdido' })
    Conferir 'Perdido existe, inativo, encerra, resultado PERDA' ($perdido.Count -eq 1 -and -not [bool]$perdido[0].ativo -and [bool]$perdido[0].encerra -and $perdido[0].resultado -eq 'PERDA')

    Write-Host '--- papel inicial das etapas'
    $pf = @($etapas | Where-Object { $_.valor -ceq 'Pedido Fechado' })[0]
    Conferir 'Pedido Fechado: encerra, GANHO, FUNIL, sem família' ([bool]$pf.encerra -and $pf.resultado -eq 'GANHO' -and $pf.quadro -eq 'FUNIL' -and -not [bool]$pf.exige_familia)
    $ng = @($etapas | Where-Object { $_.valor -ceq 'Negociação' })[0]
    Conferir 'Negociação: aberta, FUNIL, exige família' (-not [bool]$ng.encerra -and $ng.quadro -eq 'FUNIL' -and [bool]$ng.exige_familia)
    $ci = @($etapas | Where-Object { $_.valor -ceq 'Contato Inicial' })[0]
    Conferir 'Contato Inicial: aberta, PROSPECCAO' (-not [bool]$ci.encerra -and $ci.quadro -eq 'PROSPECCAO')
    Conferir 'nenhuma etapa sem papel gravado' (@($etapas | Where-Object { $null -eq $_.encerra -or $null -eq $_.quadro }).Count -eq 0)

    Write-Host '--- índice único e relatório'
    $rejeitou = $false
    try { [void](Invoke-Comando $cn "INSERT INTO listas (tipo, valor, chave, ordem, ativo) VALUES ('Etapa', 'TESTE DUPLICADO', 'PEDIDO FECHADO', 99, True)" @() $null) }
    catch { $rejeitou = $true }
    Conferir 'índice único (tipo, chave) recusa duplicata' $rejeitou
    Conferir 'relatório cita o órfão' (@($log.linhas | Where-Object { $_ -match 'Qualificando Antigo' }).Count -gt 0)
    Conferir 'relatório cita a unificação' (@($log.linhas | Where-Object { $_ -match 'PEDIDO FECHADO' }).Count -gt 0)

    Write-Host '--- idempotência: segunda aplicação não muda nada'
    $antes = @(
        [int](Invoke-Escalar $cn 'SELECT COUNT(*) FROM listas' @() $null),
        [int](Invoke-Escalar $cn 'SELECT COUNT(*) FROM municipios' @() $null),
        [int](Invoke-Escalar $cn 'SELECT COUNT(*) FROM log_alteracoes' @() $null),
        [int](Invoke-Escalar $cn 'SELECT SUM(versao) FROM oportunidades' @() $null))
    $log2 = NovoLog
    $erro = $null
    try { & $m.Aplicar $cn $null $log2 $pastaMig } catch { $erro = $_.Exception.Message }
    $depois = @(
        [int](Invoke-Escalar $cn 'SELECT COUNT(*) FROM listas' @() $null),
        [int](Invoke-Escalar $cn 'SELECT COUNT(*) FROM municipios' @() $null),
        [int](Invoke-Escalar $cn 'SELECT COUNT(*) FROM log_alteracoes' @() $null),
        [int](Invoke-Escalar $cn 'SELECT SUM(versao) FROM oportunidades' @() $null))
    Conferir 'segunda aplicação sem erro' ($null -eq $erro) $erro
    Conferir 'segunda aplicação não altera contagens nem versões' (($antes -join ',') -eq ($depois -join ','))
    Write-Host '--- conferências da cópia (passo 8 da montagem)'
    . (Join-Path $raizBase 'SISTEMA\DADOS\lib\AtualizacaoBanco.ps1')
    # o runner (aplicar-migracoes) grava a versão do esquema depois da migração estrutural
    Set-ConfigBanco $cn 'versao_esquema' '1.1' $null
    # a montagem aplica todas as pendentes: a 004 (esquema 1.2, ciclo 010) vem depois da 003
    $m4 = & (Join-Path $pastaMig '004-esquema-1-2.ps1')
    & $m4.Aplicar $cn $null (NovoLog) $pastaMig
    Set-ConfigBanco $cn 'versao_esquema' $script:EsquemaAtual $null
    $problemas = Get-ProblemasDoBanco $cn 5571
    Conferir 'cópia migrada passa nas conferências do passo 8' ($problemas.Count -eq 0) ($problemas -join '; ')
} finally {
    if ($cn) { $cn.Close(); $cn.Dispose() }
    Remove-Item -LiteralPath $copia -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $origem -Force -ErrorAction SilentlyContinue
}
Write-Host ''
if ($script:falhas -eq 0) { Write-Host 'TUDO CERTO.' -ForegroundColor Green; exit 0 }
Write-Host ($script:falhas.ToString() + ' falha(s).') -ForegroundColor Red; exit 1
