<#
  carregar.ps1 - carga inicial do CRM Zapromaq no .accdb
  Le os arquivos JSON produzidos por extrair.py (sessao de IA) e insere,
  uma transacao por tabela. Nao interpreta planilha: a normalizacao ja veio pronta.
#>
param(
    [string]$Banco = "",
    [string]$Dados = ""
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Data

$base = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($Banco)) { $Banco = Join-Path $base 'crm_zapromaq_teste.accdb' }
if ([string]::IsNullOrWhiteSpace($Dados)) { $Dados = Join-Path $base 'carga' }
$log = Join-Path $base ('log-carregar-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.txt')
$saida = New-Object System.Collections.ArrayList
function L($t) { [void]$saida.Add([string]$t); Write-Host $t }
function Fim { $saida | Set-Content -Path $log -Encoding UTF8; Write-Host ""; Write-Host ("Log: " + $log) }

L "=== CARGA INICIAL - CRM Comercial Zapromaq ==="
L ("Data..: " + (Get-Date -Format 'dd/MM/yyyy HH:mm:ss') + "   " + $env:COMPUTERNAME + "\" + $env:USERNAME)
L ("Banco.: " + $Banco)
L ("Dados.: " + $Dados)
if (-not [Environment]::Is64BitProcess) { L "ERRO: rode em PowerShell 64 bits."; Fim; exit 1 }
if (-not (Test-Path $Banco)) { L "ERRO: banco nao encontrado. Rode criar-banco.ps1 antes."; Fim; exit 1 }

function LerJson($nome) {
    $caminho = Join-Path $Dados $nome
    if (-not (Test-Path $caminho)) { throw "arquivo de carga ausente: $caminho" }
    return (Get-Content -Path $caminho -Raw -Encoding UTF8 | ConvertFrom-Json)
}

$jListas = LerJson 'carga-listas.json'
$jClientes = LerJson 'carga-clientes.json'
$jContatos = LerJson 'carga-contatos.json'
$jOportunidades = LerJson 'carga-oportunidades.json'
L ("lidos: listas=" + $jListas.Count + " clientes=" + $jClientes.Count + " contatos=" + $jContatos.Count + " oportunidades=" + $jOportunidades.Count)

$provider = $null
foreach ($prov in @('Microsoft.ACE.OLEDB.16.0','Microsoft.ACE.OLEDB.12.0')) {
    try {
        $teste = New-Object System.Data.OleDb.OleDbConnection("Provider=$prov;Data Source=$Banco;")
        $teste.Open(); $teste.Close(); $teste.Dispose()
        $provider = $prov; break
    } catch { }
}
if (-not $provider) { L "ERRO: nenhum provedor ACE disponivel."; Fim; exit 1 }

$cn = New-Object System.Data.OleDb.OleDbConnection("Provider=$provider;Data Source=$Banco;Persist Security Info=False;")
$cn.Open()

$tpTexto = [System.Data.OleDb.OleDbType]::VarWChar
$tpMemo  = [System.Data.OleDb.OleDbType]::LongVarWChar
$tpData  = [System.Data.OleDb.OleDbType]::Date
$tpInt   = [System.Data.OleDb.OleDbType]::Integer
$tpMoeda = [System.Data.OleDb.OleDbType]::Currency
$tpBool  = [System.Data.OleDb.OleDbType]::Boolean

function Escalar($sql) {
    $cmd = $cn.CreateCommand(); $cmd.CommandText = $sql
    $v = $cmd.ExecuteScalar(); $cmd.Dispose()
    if ($null -eq $v -or $v -is [System.DBNull]) { return $null }
    return $v
}
function AddPar($cmd, $tipo, $valor) {
    $par = $cmd.CreateParameter()
    $par.OleDbType = $tipo
    if ($null -eq $valor -or ($valor -is [string] -and $valor -eq '')) { $par.Value = [DBNull]::Value }
    else { $par.Value = $valor }
    [void]$cmd.Parameters.Add($par)
}
function ComoData($txt) {
    if ([string]::IsNullOrWhiteSpace($txt)) { return $null }
    return [datetime]::ParseExact($txt, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
}

foreach ($tab in @('clientes','contatos','oportunidades','listas')) {
    $qt = [int](Escalar ("SELECT COUNT(*) FROM " + $tab))
    if ($qt -gt 0) { L ("ERRO: a tabela " + $tab + " ja tem " + $qt + " registro(s). A carga exige tabelas vazias."); $cn.Close(); Fim; exit 1 }
}

$agora = Get-Date
$autor = 'MIGRACAO'
$inicioTudo = Get-Date

# ------------------------------------------------------------------ LISTAS
L ""
L "--- listas ---"
$inicio = Get-Date
$trans = $cn.BeginTransaction()
try {
    foreach ($it in $jListas) {
        $cmd = $cn.CreateCommand(); $cmd.Transaction = $trans
        $cmd.CommandText = "INSERT INTO listas (tipo, valor, ordem, ativo) VALUES (?,?,?,?)"
        AddPar $cmd $tpTexto $it.tipo
        AddPar $cmd $tpTexto $it.valor
        AddPar $cmd $tpInt   ([int]$it.ordem)
        AddPar $cmd $tpBool  $true
        [void]$cmd.ExecuteNonQuery(); $cmd.Dispose()
    }
    $trans.Commit()
    L ("  " + $jListas.Count + " opcoes em " + [math]::Round(((Get-Date)-$inicio).TotalMilliseconds) + " ms")
} catch { $trans.Rollback(); L ("  ROLLBACK listas: " + $_.Exception.Message); $cn.Close(); Fim; exit 1 }

# ------------------------------------------------------------------ CLIENTES
L ""
L "--- clientes ---"
$inicio = Get-Date
$mapaCliente = @{}
$trans = $cn.BeginTransaction()
try {
    foreach ($reg in $jClientes) {
        $cmd = $cn.CreateCommand(); $cmd.Transaction = $trans
        $cmd.CommandText = "INSERT INTO clientes (codigo_cliente, estagio, empresa, cnpj, cidade, uf, segmento," +
            " telefone, email, responsavel, aderencia, porte, qualificacao, observacoes," +
            " ativo, versao, criado_em, alterado_em, alterado_por) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,1,?,?,?)"
        AddPar $cmd $tpInt   $reg.codigo_cliente
        AddPar $cmd $tpTexto $reg.estagio
        AddPar $cmd $tpTexto $reg.empresa
        AddPar $cmd $tpTexto $reg.cnpj
        AddPar $cmd $tpTexto $reg.cidade
        AddPar $cmd $tpTexto $reg.uf
        AddPar $cmd $tpTexto $reg.segmento
        AddPar $cmd $tpTexto $reg.telefone
        AddPar $cmd $tpTexto $reg.email
        AddPar $cmd $tpTexto $reg.responsavel
        AddPar $cmd $tpInt   $reg.aderencia
        AddPar $cmd $tpInt   $reg.porte
        AddPar $cmd $tpTexto $reg.qualificacao
        AddPar $cmd $tpMemo  $reg.observacoes
        AddPar $cmd $tpBool  $true
        AddPar $cmd $tpData  $agora
        AddPar $cmd $tpData  $agora
        AddPar $cmd $tpTexto $autor
        [void]$cmd.ExecuteNonQuery(); $cmd.Dispose()

        $cmdId = $cn.CreateCommand(); $cmdId.Transaction = $trans
        $cmdId.CommandText = "SELECT @@IDENTITY"
        $novoId = [int]$cmdId.ExecuteScalar(); $cmdId.Dispose()
        $mapaCliente[[int]$reg.seq] = $novoId
    }
    $trans.Commit()
    L ("  " + $jClientes.Count + " clientes em " + [math]::Round(((Get-Date)-$inicio).TotalSeconds,1) + " s")
} catch { $trans.Rollback(); L ("  ROLLBACK clientes: " + $_.Exception.Message); $cn.Close(); Fim; exit 1 }

# ------------------------------------------------------------------ CONTATOS
L ""
L "--- contatos ---"
$inicio = Get-Date
$mapaContato = @{}
$trans = $cn.BeginTransaction()
try {
    foreach ($reg in $jContatos) {
        $idCli = $mapaCliente[[int]$reg.cliente_seq]
        if (-not $idCli) { throw ("contato " + $reg.codigo + " sem cliente mapeado (seq " + $reg.cliente_seq + ")") }
        $cmd = $cn.CreateCommand(); $cmd.Transaction = $trans
        $cmd.CommandText = "INSERT INTO contatos (codigo, controle, id_cliente, nome, cargo, telefone, email," +
            " observacoes, ativo, versao, criado_em, alterado_em, alterado_por) VALUES (?,?,?,?,?,?,?,?,?,1,?,?,?)"
        AddPar $cmd $tpTexto $reg.codigo
        AddPar $cmd $tpInt   ([int]$reg.controle)
        AddPar $cmd $tpInt   $idCli
        AddPar $cmd $tpTexto $reg.nome
        AddPar $cmd $tpTexto $reg.cargo
        AddPar $cmd $tpTexto $reg.telefone
        AddPar $cmd $tpTexto $reg.email
        AddPar $cmd $tpMemo  $reg.observacoes
        AddPar $cmd $tpBool  $true
        AddPar $cmd $tpData  $agora
        AddPar $cmd $tpData  $agora
        AddPar $cmd $tpTexto $autor
        [void]$cmd.ExecuteNonQuery(); $cmd.Dispose()

        $cmdId = $cn.CreateCommand(); $cmdId.Transaction = $trans
        $cmdId.CommandText = "SELECT @@IDENTITY"
        $novoId = [int]$cmdId.ExecuteScalar(); $cmdId.Dispose()
        $mapaContato[[int]$reg.seq] = $novoId
    }
    $trans.Commit()
    L ("  " + $jContatos.Count + " contatos em " + [math]::Round(((Get-Date)-$inicio).TotalSeconds,1) + " s")
} catch { $trans.Rollback(); L ("  ROLLBACK contatos: " + $_.Exception.Message); $cn.Close(); Fim; exit 1 }

# ------------------------------------------------------------------ OPORTUNIDADES
L ""
L "--- oportunidades ---"
$inicio = Get-Date
$mapaOportunidade = @{}
$trans = $cn.BeginTransaction()
try {
    foreach ($reg in $jOportunidades) {
        $idCto = $mapaContato[[int]$reg.contato_seq]
        $idCli = $mapaCliente[[int]$reg.cliente_seq]
        if (-not $idCto) { throw ("atendimento " + $reg.codigo + " sem contato mapeado") }
        if (-not $idCli) { throw ("atendimento " + $reg.codigo + " sem cliente mapeado") }
        $cmd = $cn.CreateCommand(); $cmd.Transaction = $trans
        $cmd.CommandText = "INSERT INTO oportunidades (codigo, controle, id_contato, id_cliente, orcamento, etapa," +
            " responsavel, maquina, familia, categoria, tipo_venda, valor, origem, prioridade," +
            " dt_entrada, dt_proposta, ultima_interacao, tentativas, prox_acao, dt_prox_acao," +
            " retomar_em, motivo_desfecho, dt_desfecho, observacoes," +
            " versao, criado_em, alterado_em, alterado_por)" +
            " VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,1,?,?,?)"
        AddPar $cmd $tpTexto $reg.codigo
        AddPar $cmd $tpInt   ([int]$reg.controle)
        AddPar $cmd $tpInt   $idCto
        AddPar $cmd $tpInt   $idCli
        AddPar $cmd $tpTexto $reg.orcamento
        AddPar $cmd $tpTexto $reg.etapa
        AddPar $cmd $tpTexto $reg.responsavel
        AddPar $cmd $tpTexto $reg.maquina
        AddPar $cmd $tpTexto $reg.familia
        AddPar $cmd $tpTexto $reg.categoria
        AddPar $cmd $tpTexto $reg.tipo_venda
        AddPar $cmd $tpMoeda $reg.valor
        AddPar $cmd $tpTexto $reg.origem
        AddPar $cmd $tpTexto $reg.prioridade
        AddPar $cmd $tpData  (ComoData $reg.dt_entrada)
        AddPar $cmd $tpData  (ComoData $reg.dt_proposta)
        AddPar $cmd $tpData  (ComoData $reg.ultima_interacao)
        AddPar $cmd $tpInt   $reg.tentativas
        AddPar $cmd $tpMemo  $reg.prox_acao
        AddPar $cmd $tpData  (ComoData $reg.dt_prox_acao)
        AddPar $cmd $tpData  (ComoData $reg.retomar_em)
        AddPar $cmd $tpTexto $reg.motivo_desfecho
        AddPar $cmd $tpData  (ComoData $reg.dt_desfecho)
        AddPar $cmd $tpMemo  $reg.observacoes
        AddPar $cmd $tpData  $agora
        AddPar $cmd $tpData  $agora
        AddPar $cmd $tpTexto $autor
        [void]$cmd.ExecuteNonQuery(); $cmd.Dispose()

        $cmdId = $cn.CreateCommand(); $cmdId.Transaction = $trans
        $cmdId.CommandText = "SELECT @@IDENTITY"
        $novoId = [int]$cmdId.ExecuteScalar(); $cmdId.Dispose()
        $mapaOportunidade[[int]$reg.seq] = $novoId
    }
    $trans.Commit()
    L ("  " + $jOportunidades.Count + " oportunidades em " + [math]::Round(((Get-Date)-$inicio).TotalSeconds,1) + " s")
} catch { $trans.Rollback(); L ("  ROLLBACK oportunidades: " + $_.Exception.Message); $cn.Close(); Fim; exit 1 }

# --------------------------------------------------- retomadas (segunda passada)
$comAnterior = @($jOportunidades | Where-Object { $_.anterior_seq })
if ($comAnterior.Count -gt 0) {
    L ""
    L "--- retomadas ---"
    $trans = $cn.BeginTransaction()
    try {
        foreach ($reg in $comAnterior) {
            $cmd = $cn.CreateCommand(); $cmd.Transaction = $trans
            $cmd.CommandText = "UPDATE oportunidades SET id_atendimento_anterior=? WHERE id=?"
            AddPar $cmd $tpInt $mapaOportunidade[[int]$reg.anterior_seq]
            AddPar $cmd $tpInt $mapaOportunidade[[int]$reg.seq]
            [void]$cmd.ExecuteNonQuery(); $cmd.Dispose()
        }
        $trans.Commit()
        L ("  " + $comAnterior.Count + " vinculo(s) de retomada")
    } catch { $trans.Rollback(); L ("  ROLLBACK retomadas: " + $_.Exception.Message) }
} else {
    L ""
    L "--- retomadas: nenhuma na origem ---"
}

# ------------------------------------------------------------------ resumo
L ""
L "--- conferencia rapida ---"
foreach ($tab in @('listas','clientes','contatos','oportunidades')) {
    L ("  " + $tab.PadRight(16) + (Escalar ("SELECT COUNT(*) FROM " + $tab)))
}
L ("  soma de valor  " + (Escalar "SELECT SUM(valor) FROM oportunidades"))
L ("  pre-clientes   " + (Escalar "SELECT COUNT(*) FROM clientes WHERE estagio='Pré-cliente'"))
L ("  clientes       " + (Escalar "SELECT COUNT(*) FROM clientes WHERE estagio='Cliente'"))

$cn.Close(); $cn.Dispose()
L ""
L ("Carga concluida em " + [math]::Round(((Get-Date)-$inicioTudo).TotalSeconds,1) + " s")
L "Proximo passo: aplicar-integridade.ps1 e depois conferir-base.ps1"
Fim
