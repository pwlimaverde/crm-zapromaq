<#
  Comum.ps1 - funções compartilhadas pelos scripts de SISTEMA (build, banco, operação, testes).

  Carregar com:   . (Join-Path <pasta DADOS> 'lib\Comum.ps1')

  Regra de ouro: nenhum caminho absoluto e nenhuma letra de unidade. Tudo
  sai da posição do próprio script, subindo até achar a raiz de BASE (a
  pasta que contém SISTEMA\DADOS\VERSAO.txt). Copiar BASE inteira para
  outro lugar tem de bastar.

  Requisitos: Windows PowerShell 5.1 64 bits; motor ACE (Office 64 bits ou
  Access Database Engine 2016 Redistributable x64) para as funções de banco.
#>

Add-Type -AssemblyName System.Data

$script:NomeBanco = 'crm_zapromaq.accdb'
$script:NomeFront = 'CRM_Zapromaq.xlsm'

# ------------------------------------------------------------------ caminhos
function Get-RaizBase([string]$apartirDe) {
    $d = (Resolve-Path -LiteralPath $apartirDe).ProviderPath
    while ($d) {
        if (Test-Path -LiteralPath (Join-Path $d 'SISTEMA\DADOS\VERSAO.txt')) { return $d }
        $pai = Split-Path -Parent $d
        if ($pai -eq $d) { break }
        $d = $pai
    }
    throw ('Raiz de BASE não encontrada acima de ' + $apartirDe + ' (procurei SISTEMA\DADOS\VERSAO.txt).')
}

function Get-PastaDados([string]$raiz) { return (Join-Path $raiz 'SISTEMA\DADOS') }

function Get-PastaExecucao([string]$raiz, [string]$sub) {
    $p = Join-Path (Get-PastaDados $raiz) ('execucao\' + $sub)
    if (-not (Test-Path -LiteralPath $p)) { New-Item -ItemType Directory -Path $p -Force | Out-Null }
    return $p
}

# Unidade mapeada (G:\...) vira caminho UNC (\\servidor\...). O front roda a
# partir de Documentos, em qualquer estação: só o UNC vale para todas.
function ConvertTo-CaminhoUNC([string]$caminho) {
    $caminho = $caminho -replace '^Microsoft\.PowerShell\.Core\\FileSystem::', ''
    $cheio = [System.IO.Path]::GetFullPath($caminho)
    if ($cheio -match '^([A-Za-z]):\\') {
        $letra = $Matches[1] + ':'
        try {
            $disco = Get-WmiObject Win32_LogicalDisk -Filter ("DeviceID='" + $letra + "'") -ErrorAction Stop
            if ($disco -and $disco.DriveType -eq 4 -and $disco.ProviderName) {
                return ($disco.ProviderName.TrimEnd('\') + $cheio.Substring(2))
            }
        } catch { }
    }
    return $cheio
}

function Get-CaminhoBanco([string]$raiz) { return (Join-Path $raiz $script:NomeBanco) }
function Get-CaminhoFront([string]$raiz) { return (Join-Path $raiz $script:NomeFront) }

# ------------------------------------------------------------------ log
function New-Log([string]$raiz, [string]$nome) {
    $pasta = Get-PastaExecucao $raiz 'logs'
    $arq = Join-Path $pasta ($nome + '-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.txt')
    $obj = New-Object PSObject -Property @{ Arquivo = $arq; Linhas = (New-Object System.Collections.ArrayList) }
    $obj | Add-Member -MemberType ScriptMethod -Name L -Value {
        param($t)
        [void]$this.Linhas.Add([string]$t)
        Write-Host $t
    }
    $obj | Add-Member -MemberType ScriptMethod -Name Salvar -Value {
        $u = New-Object System.Text.UTF8Encoding($true)
        [System.IO.File]::WriteAllLines($this.Arquivo, [string[]]$this.Linhas.ToArray(), $u)
        Write-Host ''
        Write-Host ('Log: ' + $this.Arquivo)
    }
    return $obj
}

# ------------------------------------------------------------------ banco (ACE)
function Get-ProvedorAce {
    if (-not [Environment]::Is64BitProcess) { throw 'Rode no PowerShell 64 bits: o ACE do Office é 64 bits.' }
    $instalados = @()
    try { $instalados = @((New-Object System.Data.OleDb.OleDbEnumerator).GetElements() | ForEach-Object { $_.SOURCES_NAME }) } catch { }
    foreach ($p in @('Microsoft.ACE.OLEDB.16.0', 'Microsoft.ACE.OLEDB.12.0')) {
        if ($instalados -contains $p) { return $p }
    }
    throw 'Motor ACE não encontrado. Instale o Office 64 bits ou o Access Database Engine 2016 Redistributable x64.'
}

# Conexão curta. OLE DB Services=-4 desliga o pool (ver referencias/08-dotnet-oledb):
# ao fechar, o arquivo .laccdb é liberado de verdade.
function Open-Banco([string]$caminho) {
    if (-not (Test-Path -LiteralPath $caminho)) { throw ('Banco não encontrado: ' + $caminho) }
    $cn = New-Object System.Data.OleDb.OleDbConnection(
        'Provider=' + (Get-ProvedorAce) + ';Data Source=' + $caminho + ';OLE DB Services=-4;Persist Security Info=False;')
    $cn.Open()
    return $cn
}

# Parâmetro tipado: tipos texto, memo, inteiro, moeda, numero, data, logico
function New-Parametro([string]$tipo, $valor) { return @{ Tipo = $tipo; Valor = $valor } }

function New-ComandoBanco($cn, [string]$sql, [array]$params, $tx) {
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
            default   { throw ('Tipo de parâmetro desconhecido: ' + $p.Tipo) }
        }
        if ($null -eq $p.Valor -or ($p.Valor -is [string] -and $p.Valor -eq '')) { $par.Value = [DBNull]::Value } else { $par.Value = $p.Valor }
        [void]$cmd.Parameters.Add($par)
    }
    return $cmd
}

function Invoke-Escalar($cn, [string]$sql, [array]$params, $tx) {
    $c = New-ComandoBanco $cn $sql $params $tx
    try { $v = $c.ExecuteScalar() } finally { $c.Dispose() }
    if ($v -is [System.DBNull]) { return $null }
    return $v
}

function Invoke-Comando($cn, [string]$sql, [array]$params, $tx) {
    $c = New-ComandoBanco $cn $sql $params $tx
    try { return $c.ExecuteNonQuery() } finally { $c.Dispose() }
}

function Get-VersaoEsquemaBanco($cn) {
    return [string](Invoke-Escalar $cn "SELECT valor FROM config WHERE chave='versao_esquema'" @())
}

# Grava ou atualiza uma chave da tabela config (parametrizado).
function Set-ConfigBanco($cn, [string]$chave, [string]$valor, $tx) {
    $n = Invoke-Comando $cn 'UPDATE config SET valor=? WHERE chave=?' @((New-Parametro 'texto' $valor), (New-Parametro 'texto' $chave)) $tx
    if ($n -eq 0) {
        [void](Invoke-Comando $cn 'INSERT INTO config (chave, valor) VALUES (?, ?)' @((New-Parametro 'texto' $chave), (New-Parametro 'texto' $valor)) $tx)
    }
}
