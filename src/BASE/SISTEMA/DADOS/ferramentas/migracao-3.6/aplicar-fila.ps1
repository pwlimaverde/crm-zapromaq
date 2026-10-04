<#
  aplicar-fila.ps1 - aplica no banco os comandos deixados pela sessao de IA.

  A sessao NAO grava no .accdb: nenhuma biblioteca fora do Windows implementa
  o protocolo de bloqueio do ACE, e gravar por fora enquanto os vendedores usam
  o front e risco de corromper o arquivo. Ela escreve um .json na FILA; este
  script e o unico que encosta no banco, e pelo mesmo motor do front.

  Conjunto de operacoes FECHADO. Qualquer outra coisa e recusada.
#>
param(
    [string]$Banco = "",
    [switch]$Simular
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Data

$base = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($Banco)) { $Banco = Join-Path $base 'crm_zapromaq_teste.accdb' }
$pastaFila = Join-Path (Split-Path -Parent $Banco) 'FILA'
$pastaOk   = Join-Path $pastaFila 'APLICADOS'
$pastaNao  = Join-Path $pastaOk 'RECUSADOS'
$log = Join-Path $base ('log-fila-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.txt')
$saida = New-Object System.Collections.ArrayList
function L($t) { [void]$saida.Add([string]$t); Write-Host $t }
function Fim { $saida | Set-Content -Path $log -Encoding UTF8; Write-Host ""; Write-Host ("Log: " + $log) }

L "=== APLICAR FILA - CRM Comercial Zapromaq ==="
L ("Data..: " + (Get-Date -Format 'dd/MM/yyyy HH:mm:ss') + "   " + $env:COMPUTERNAME + "\" + $env:USERNAME)
L ("Banco.: " + $Banco)
L ("Fila..: " + $pastaFila)
if ($Simular) { L "MODO SIMULACAO - nada sera gravado" }

if (-not (Test-Path $Banco)) { L "ERRO: banco nao encontrado."; Fim; exit 1 }
if (-not (Test-Path $pastaFila)) { L "ERRO: pasta da fila nao encontrada."; Fim; exit 1 }
foreach ($p in @($pastaOk, $pastaNao)) { if (-not (Test-Path $p)) { New-Item -ItemType Directory -Path $p | Out-Null } }

$arquivos = @(Get-ChildItem -Path $pastaFila -Filter '*.json' -File | Sort-Object Name)
if ($arquivos.Count -eq 0) { L ""; L "Nenhuma pendencia."; Fim; exit 0 }
L ("Arquivos na fila: " + $arquivos.Count)

$provider = $null
foreach ($prov in @('Microsoft.ACE.OLEDB.16.0','Microsoft.ACE.OLEDB.12.0')) {
    try { $t = New-Object System.Data.OleDb.OleDbConnection("Provider=$prov;Data Source=$Banco;"); $t.Open(); $t.Close(); $t.Dispose(); $provider = $prov; break } catch { }
}
if (-not $provider) { L "ERRO: nenhum provedor ACE."; Fim; exit 1 }

$cn = New-Object System.Data.OleDb.OleDbConnection("Provider=$provider;Data Source=$Banco;")
$cn.Open()
$tpTexto = [System.Data.OleDb.OleDbType]::VarWChar
$tpData  = [System.Data.OleDb.OleDbType]::Date
$tpInt   = [System.Data.OleDb.OleDbType]::Integer

function AddPar($cmd, $tipo, $valor) {
    $par = $cmd.CreateParameter(); $par.OleDbType = $tipo
    if ($null -eq $valor -or ($valor -is [string] -and $valor -eq '')) { $par.Value = [DBNull]::Value } else { $par.Value = $valor }
    [void]$cmd.Parameters.Add($par)
}
# campos cadastrais que a fila pode corrigir - lista fechada
$camposCadastro = @('empresa','cnpj','cidade','uf','segmento','telefone','email')

$totalOk = 0; $totalNao = 0

foreach ($arq in $arquivos) {
    L ""
    L ("--- " + $arq.Name + " ---")
    $recusas = New-Object System.Collections.ArrayList
    $aplicados = 0
    try {
        $pacote = Get-Content -Path $arq.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
    } catch {
        L ("  RECUSADO: JSON invalido - " + $_.Exception.Message.Split([char]13)[0])
        if (-not $Simular) { Move-Item $arq.FullName (Join-Path $pastaNao $arq.Name) -Force }
        $totalNao++
        continue
    }

    L ("  gerado em: " + $pacote.gerado_em + "   autor: " + $pacote.autor)
    $comandos = @($pacote.comandos)
    L ("  comandos: " + $comandos.Count)

    foreach ($c in $comandos) {
        $op = [string]$c.op
        try {
            switch ($op) {

                'ctx_cliente' {
                    $cmdB = $cn.CreateCommand()
                    $cmdB.CommandText = "SELECT id FROM clientes WHERE codigo_cliente=?"
                    AddPar $cmdB $tpInt ([int]$c.codigo_cliente)
                    $idCli = $cmdB.ExecuteScalar(); $cmdB.Dispose()
                    if ($null -eq $idCli -or $idCli -is [System.DBNull]) { throw ("cliente " + $c.codigo_cliente + " nao encontrado") }

                    if (-not $Simular) {
                        $trans = $cn.BeginTransaction()
                        $cmd = $cn.CreateCommand(); $cmd.Transaction = $trans
                        $cmd.CommandText = "UPDATE clientes SET ctx_resumo=?, ctx_familia=?, ctx_arquivo=?," +
                                           " ctx_atualizado_em=?, ctx_atualizado_por=? WHERE id=?"
                        AddPar $cmd $tpTexto $c.ctx_resumo
                        AddPar $cmd $tpTexto $c.ctx_familia
                        AddPar $cmd $tpTexto $c.ctx_arquivo
                        AddPar $cmd $tpData  (Get-Date)
                        AddPar $cmd $tpTexto ([string]$pacote.autor)
                        AddPar $cmd $tpInt   ([int]$idCli)
                        $n = $cmd.ExecuteNonQuery(); $cmd.Dispose()
                        if ($n -ne 1) { $trans.Rollback(); throw "UPDATE afetou $n linha(s)" }
                        $cmdL = $cn.CreateCommand(); $cmdL.Transaction = $trans
                        $cmdL.CommandText = "INSERT INTO log_alteracoes (tabela,id_registro,campo,valor_antigo,valor_novo,acao,usuario,quando,origem)" +
                                            " VALUES ('clientes',?,'ctx_resumo','',?,'CONTEXTO',?,?,'FILA')"
                        AddPar $cmdL $tpInt   ([int]$idCli)
                        AddPar $cmdL $tpTexto $c.ctx_resumo
                        AddPar $cmdL $tpTexto ([string]$pacote.autor)
                        AddPar $cmdL $tpData  (Get-Date)
                        [void]$cmdL.ExecuteNonQuery(); $cmdL.Dispose()
                        $trans.Commit()
                    }
                    $aplicados++
                    L ("    OK  ctx_cliente " + $c.codigo_cliente)
                }

                'ctx_atendimento' {
                    $cmdB = $cn.CreateCommand()
                    $cmdB.CommandText = "SELECT id FROM oportunidades WHERE codigo=?"
                    AddPar $cmdB $tpTexto ([string]$c.codigo)
                    $idOpo = $cmdB.ExecuteScalar(); $cmdB.Dispose()
                    if ($null -eq $idOpo -or $idOpo -is [System.DBNull]) { throw ("atendimento " + $c.codigo + " nao encontrado") }

                    if (-not $Simular) {
                        $trans = $cn.BeginTransaction()
                        $cmd = $cn.CreateCommand(); $cmd.Transaction = $trans
                        $cmd.CommandText = "UPDATE oportunidades SET ctx_resumo=?, ctx_arquivo=?," +
                                           " ctx_atualizado_em=?, ctx_atualizado_por=? WHERE id=?"
                        AddPar $cmd $tpTexto $c.ctx_resumo
                        AddPar $cmd $tpTexto $c.ctx_arquivo
                        AddPar $cmd $tpData  (Get-Date)
                        AddPar $cmd $tpTexto ([string]$pacote.autor)
                        AddPar $cmd $tpInt   ([int]$idOpo)
                        $n = $cmd.ExecuteNonQuery(); $cmd.Dispose()
                        if ($n -ne 1) { $trans.Rollback(); throw "UPDATE afetou $n linha(s)" }
                        $cmdL = $cn.CreateCommand(); $cmdL.Transaction = $trans
                        $cmdL.CommandText = "INSERT INTO log_alteracoes (tabela,id_registro,campo,valor_antigo,valor_novo,acao,usuario,quando,origem)" +
                                            " VALUES ('oportunidades',?,'ctx_resumo','',?,'CONTEXTO',?,?,'FILA')"
                        AddPar $cmdL $tpInt   ([int]$idOpo)
                        AddPar $cmdL $tpTexto $c.ctx_resumo
                        AddPar $cmdL $tpTexto ([string]$pacote.autor)
                        AddPar $cmdL $tpData  (Get-Date)
                        [void]$cmdL.ExecuteNonQuery(); $cmdL.Dispose()
                        $trans.Commit()
                    }
                    $aplicados++
                    L ("    OK  ctx_atendimento " + $c.codigo)
                }

                'cad_cliente' {
                    $campo = [string]$c.campo
                    if ($camposCadastro -notcontains $campo) { throw ("campo nao permitido pela fila: " + $campo) }

                    $cmdB = $cn.CreateCommand()
                    $cmdB.CommandText = "SELECT id FROM clientes WHERE codigo_cliente=?"
                    AddPar $cmdB $tpInt ([int]$c.codigo_cliente)
                    $idCli = $cmdB.ExecuteScalar(); $cmdB.Dispose()
                    if ($null -eq $idCli -or $idCli -is [System.DBNull]) { throw ("cliente " + $c.codigo_cliente + " nao encontrado") }

                    $cmdA = $cn.CreateCommand()
                    $cmdA.CommandText = "SELECT " + $campo + " FROM clientes WHERE id=?"
                    AddPar $cmdA $tpInt ([int]$idCli)
                    $antes = $cmdA.ExecuteScalar(); $cmdA.Dispose()
                    if ($antes -is [System.DBNull]) { $antes = '' }

                    if ([string]$antes -eq [string]$c.valor) {
                        L ("    --  cad_cliente " + $c.codigo_cliente + " " + $campo + ": ja esta igual, nada a fazer")
                        $aplicados++
                    }
                    elseif (-not $Simular) {
                        $trans = $cn.BeginTransaction()
                        $cmd = $cn.CreateCommand(); $cmd.Transaction = $trans
                        $cmd.CommandText = "UPDATE clientes SET " + $campo + "=?, versao=versao+1," +
                                           " alterado_em=?, alterado_por=? WHERE id=?"
                        AddPar $cmd $tpTexto ([string]$c.valor)
                        AddPar $cmd $tpData  (Get-Date)
                        AddPar $cmd $tpTexto ([string]$pacote.autor)
                        AddPar $cmd $tpInt   ([int]$idCli)
                        $n = $cmd.ExecuteNonQuery(); $cmd.Dispose()
                        if ($n -ne 1) { $trans.Rollback(); throw "UPDATE afetou $n linha(s)" }
                        $cmdL = $cn.CreateCommand(); $cmdL.Transaction = $trans
                        $cmdL.CommandText = "INSERT INTO log_alteracoes (tabela,id_registro,campo,valor_antigo,valor_novo,acao,usuario,quando,origem)" +
                                            " VALUES ('clientes',?,?,?,?,'ALTERAR',?,?,'FILA')"
                        AddPar $cmdL $tpInt   ([int]$idCli)
                        AddPar $cmdL $tpTexto $campo
                        AddPar $cmdL $tpTexto ([string]$antes)
                        AddPar $cmdL $tpTexto ([string]$c.valor)
                        AddPar $cmdL $tpTexto ([string]$pacote.autor)
                        AddPar $cmdL $tpData  (Get-Date)
                        [void]$cmdL.ExecuteNonQuery(); $cmdL.Dispose()
                        $trans.Commit()
                        $aplicados++
                        L ("    OK  cad_cliente " + $c.codigo_cliente + " " + $campo + ": '" + $antes + "' -> '" + $c.valor + "'")
                    }
                    else {
                        $aplicados++
                        L ("    OK  cad_cliente " + $c.codigo_cliente + " " + $campo + ": '" + $antes + "' -> '" + $c.valor + "'  (simulado)")
                    }
                }

                default {
                    throw ("operacao nao reconhecida: '" + $op + "'")
                }
            }
        } catch {
            $msg = $_.Exception.Message.Split([char]13)[0]
            [void]$recusas.Add(($op + ": " + $msg))
            L ("    RECUSADO  " + $op + " -> " + $msg)
        }
    }

    if ($Simular) {
        if ($recusas.Count -eq 0) {
            $totalOk++
            L ("  SIMULACAO: " + $aplicados + " comando(s) seriam aplicados, nenhuma recusa.")
        } else {
            $totalNao++
            L ("  SIMULACAO: " + $aplicados + " aplicavel(is), " + $recusas.Count + " recusado(s). Nada foi gravado nem movido.")
        }
        continue
    }

    if ($recusas.Count -eq 0) {
        Move-Item $arq.FullName (Join-Path $pastaOk $arq.Name) -Force
        $totalOk++
        L ("  aplicado: " + $aplicados + " comando(s). Arquivo movido para APLICADOS.")
    } else {
        $destino = Join-Path $pastaNao $arq.Name
        Move-Item $arq.FullName $destino -Force
        $motivo = New-Object System.Collections.ArrayList
        [void]$motivo.Add("Processado em " + (Get-Date -Format 'dd/MM/yyyy HH:mm:ss') + " por " + $env:USERNAME)
        [void]$motivo.Add("")
        [void]$motivo.Add("APLICADOS NO BANCO: " + $aplicados + " comando(s) - NAO reprocessar sem conferir")
        [void]$motivo.Add("RECUSADOS: " + $recusas.Count)
        [void]$motivo.Add("")
        foreach ($r in $recusas) { [void]$motivo.Add("  - " + $r) }
        $motivo | Set-Content -Path ($destino + '.motivo.txt') -Encoding UTF8
        $totalNao++
        L ("  " + $aplicados + " aplicado(s), " + $recusas.Count + " recusado(s). Arquivo movido para RECUSADOS com o motivo.")
    }
}

$cn.Close(); $cn.Dispose()
L ""
L ("Arquivos sem nenhuma recusa: " + $totalOk + "   |   com ao menos uma recusa: " + $totalNao)
L "(o contador e de ARQUIVOS, nao de comandos: um arquivo com 3 aplicados e 1 recusado conta como 1 com recusa)"
Fim
