<#
  Banco.ps1 - acesso ao .accdb pelo motor ACE (o mesmo do front).

  Regras que valem aqui e no front:
    - conexao curta: abre, executa, fecha - por chamada de ferramenta, nunca por linha;
    - SQL sempre parametrizado ('?' posicional); data NUNCA como literal;
    - NZ() nao existe via OLE DB: nulo e tratado aqui, nao no SQL;
    - dado e log_alteracoes na mesma transacao.
#>

$script:ProvedorAce = $null

# Versoes de esquema com que este servidor sabe trabalhar. O front grava
# config.versao_esquema; se o banco estiver numa versao desconhecida, a
# gravacao e recusada (leitura continua liberada).
$script:EsquemasSuportados = @('1.0')

function Get-ProvedorAce {
    if ($script:ProvedorAce) { return $script:ProvedorAce }
    if (-not [Environment]::Is64BitProcess) {
        throw 'O servidor precisa rodar em PowerShell 64 bits (o ACE do Office e 64 bits).'
    }
    $instalados = @()
    try {
        $instalados = @((New-Object System.Data.OleDb.OleDbEnumerator).GetElements() |
                        ForEach-Object { $_.SOURCES_NAME })
    } catch { }
    foreach ($p in @('Microsoft.ACE.OLEDB.16.0', 'Microsoft.ACE.OLEDB.12.0')) {
        if ($instalados -contains $p) { $script:ProvedorAce = $p; return $p }
    }
    throw ('Motor ACE (Microsoft.ACE.OLEDB) nao encontrado nesta maquina. Instale o Office 64 bits ' +
           'ou o "Microsoft Access Database Engine 2016 Redistributable" x64.')
}

function Test-ErroDeBloqueio([System.Exception]$ex) {
    $m = $ex.Message
    return ($m -match 'bloquead|locked|em uso|in use|ja esta aberto|already opened|could not use|nao foi possivel usar')
}

function Open-Conexao {
    if (-not (Test-Path -LiteralPath $script:CaminhoBanco)) {
        throw ('Banco nao encontrado: ' + $script:CaminhoBanco)
    }
    $prov = Get-ProvedorAce
    # OLE DB Services=-4 desliga o pool: ao fechar, a conexao solta o arquivo de
    # verdade. Com pool, o .laccdb pode ficar preso depois do Close.
    $cs = "Provider=$prov;Data Source=$($script:CaminhoBanco);OLE DB Services=-4;Persist Security Info=False;"
    $tentativa = 0
    while ($true) {
        $tentativa++
        $cn = New-Object System.Data.OleDb.OleDbConnection($cs)
        try { $cn.Open(); return $cn }
        catch {
            $cn.Dispose()
            if ($tentativa -ge 4 -or -not (Test-ErroDeBloqueio $_.Exception)) { throw }
            Start-Sleep -Milliseconds (250 * $tentativa)
        }
    }
}

# Parametro tipado. Tipos: texto, memo, inteiro, moeda, numero, data, logico
function P([string]$tipo, $valor) {
    return @{ Tipo = $tipo; Valor = $valor }
}

function New-Comando($cn, [string]$sql, [array]$params, $tx) {
    $cmd = $cn.CreateCommand()
    $cmd.CommandText = $sql
    if ($tx) { $cmd.Transaction = $tx }
    foreach ($p in @($params)) {
        if ($null -eq $p) { continue }
        $par = $cmd.CreateParameter()
        switch ($p.Tipo) {
            'texto'   { $par.OleDbType = [System.Data.OleDb.OleDbType]::VarWChar }
            'memo'    { $par.OleDbType = [System.Data.OleDb.OleDbType]::LongVarWChar }
            'inteiro' { $par.OleDbType = [System.Data.OleDb.OleDbType]::Integer }
            'moeda'   { $par.OleDbType = [System.Data.OleDb.OleDbType]::Currency }
            'numero'  { $par.OleDbType = [System.Data.OleDb.OleDbType]::Double }
            'data'    { $par.OleDbType = [System.Data.OleDb.OleDbType]::Date }
            'logico'  { $par.OleDbType = [System.Data.OleDb.OleDbType]::Boolean }
            default   { throw ('Tipo de parametro desconhecido: ' + $p.Tipo) }
        }
        $v = $p.Valor
        if ($null -eq $v -or ($v -is [string] -and $v -eq '')) { $par.Value = [DBNull]::Value }
        else { $par.Value = $v }
        [void]$cmd.Parameters.Add($par)
    }
    return $cmd
}

# Valor do banco -> valor JSON. Data sem hora vira 'yyyy-MM-dd'.
function ConvertFrom-ValorBanco($v) {
    if ($null -eq $v -or $v -is [System.DBNull]) { return $null }
    if ($v -is [datetime]) {
        if ($v.TimeOfDay.TotalSeconds -eq 0) { return $v.ToString('yyyy-MM-dd') }
        return $v.ToString('yyyy-MM-ddTHH:mm:ss')
    }
    if ($v -is [decimal]) { return [double]$v }
    return $v
}

function Read-Linhas($cmd) {
    $linhas = New-Object System.Collections.ArrayList
    $rd = $cmd.ExecuteReader()
    try {
        while ($rd.Read()) {
            $l = [ordered]@{}
            for ($i = 0; $i -lt $rd.FieldCount; $i++) {
                $l[$rd.GetName($i)] = ConvertFrom-ValorBanco $rd.GetValue($i)
            }
            [void]$linhas.Add($l)
        }
    } finally { $rd.Close() }
    return ,$linhas
}

# Consulta de leitura: conexao propria, curta.
function Invoke-Consulta([string]$sql, [array]$params) {
    $cn = Open-Conexao
    try {
        $cmd = New-Comando $cn $sql $params $null
        try { return ,(Read-Linhas $cmd) } finally { $cmd.Dispose() }
    } finally { $cn.Close(); $cn.Dispose() }
}

function Invoke-Valor([string]$sql, [array]$params) {
    $cn = Open-Conexao
    try {
        $cmd = New-Comando $cn $sql $params $null
        try { return (ConvertFrom-ValorBanco $cmd.ExecuteScalar()) } finally { $cmd.Dispose() }
    } finally { $cn.Close(); $cn.Dispose() }
}

<#
  Executa um bloco dentro de UMA transacao. O bloco recebe um objeto com
  .Consultar(sql, params) .Valor(sql, params) .Executar(sql, params) -> linhas afetadas
  e .Log(tabela, id, campo, antigo, novo, acao).
  Qualquer throw dentro do bloco desfaz tudo.
#>
function Invoke-Transacao([scriptblock]$bloco) {
    $cn = Open-Conexao
    $tx = $null
    try {
        $tx = $cn.BeginTransaction()
        $ctx = New-Object PSObject -Property @{ Cn = $cn; Tx = $tx }
        $ctx | Add-Member -MemberType ScriptMethod -Name Consultar -Value {
            param($sql, $params)
            $c = New-Comando $this.Cn $sql $params $this.Tx
            try { return ,(Read-Linhas $c) } finally { $c.Dispose() }
        }
        $ctx | Add-Member -MemberType ScriptMethod -Name Valor -Value {
            param($sql, $params)
            $c = New-Comando $this.Cn $sql $params $this.Tx
            try { return (ConvertFrom-ValorBanco $c.ExecuteScalar()) } finally { $c.Dispose() }
        }
        $ctx | Add-Member -MemberType ScriptMethod -Name Executar -Value {
            param($sql, $params)
            $c = New-Comando $this.Cn $sql $params $this.Tx
            try { return $c.ExecuteNonQuery() } finally { $c.Dispose() }
        }
        $ctx | Add-Member -MemberType ScriptMethod -Name Log -Value {
            param($tabela, $idRegistro, $campo, $antigo, $novo, $acao)
            $sql = 'INSERT INTO log_alteracoes (tabela, id_registro, campo, valor_antigo, valor_novo,' +
                   ' acao, usuario, quando, origem) VALUES (?,?,?,?,?,?,?,?,?)'
            $c = New-Comando $this.Cn $sql @(
                (P 'texto' $tabela), (P 'inteiro' $idRegistro), (P 'texto' $campo),
                (P 'memo' (Format-ValorLog $antigo)), (P 'memo' (Format-ValorLog $novo)),
                (P 'texto' $acao), (P 'texto' $script:UsuarioLog), (P 'data' (Get-Date)),
                (P 'texto' 'IA')) $this.Tx
            try { [void]$c.ExecuteNonQuery() } finally { $c.Dispose() }
        }
        $resultado = & $bloco $ctx
        $tx.Commit()
        $tx = $null
        return $resultado
    } finally {
        if ($tx) { try { $tx.Rollback() } catch { } }
        $cn.Close(); $cn.Dispose()
    }
}

function Format-ValorLog($v) {
    if ($null -eq $v -or $v -is [System.DBNull]) { return '' }
    if ($v -is [datetime]) { return $v.ToString('dd/MM/yyyy') }
    if ($v -is [double] -or $v -is [decimal]) { return ([double]$v).ToString([Globalization.CultureInfo]::InvariantCulture) }
    return [string]$v
}

function Get-VersaoEsquema {
    return [string](Invoke-Valor "SELECT valor FROM config WHERE chave='versao_esquema'" @())
}

function Assert-EsquemaGravavel {
    $v = Get-VersaoEsquema
    if ($script:EsquemasSuportados -notcontains $v) {
        throw ("Gravacao bloqueada: o banco esta no esquema '" + $v + "' e este servidor conhece " +
               ($script:EsquemasSuportados -join ', ') + '. Atualize a pasta IA junto com o sistema.')
    }
}
