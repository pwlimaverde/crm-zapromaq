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

# Forma canônica de um caminho para COMPARAR: absoluto, sem prefixo \\?\, sem ponto nem barra final,
# unidade mapeada vista como UNC, \\<esta máquina>\X$ visto como X:\. Compara-se sem diferença de caixa.
function Get-CaminhoCanonico([string]$caminho) {
    $c = $caminho -replace '^\\\\\?\\', ''
    # relativo é relativo à pasta ATUAL DO POWERSHELL (o .NET usa a do processo, que pode ser outra)
    if (-not [System.IO.Path]::IsPathRooted($c)) { $c = Join-Path (Get-Location).ProviderPath $c }
    $c = [System.IO.Path]::GetFullPath($c)
    $m =[regex]::Match($c, '^\\\\(localhost|127\.0\.0\.1|' + [regex]::Escape($env:COMPUTERNAME) + ')\\([A-Za-z])\$(\\.*)?$', 'IgnoreCase')
    if ($m.Success) { $c = $m.Groups[2].Value.ToUpperInvariant() + ':' + $m.Groups[3].Value }
    $c = ConvertTo-CaminhoUNC $c
    return $c.TrimEnd('\').ToUpperInvariant()
}

# Dois caminhos (UNC, letra, relativo, com ou sem caixa) apontam para o mesmo arquivo ou pasta?
function Test-MesmoCaminho([string]$a, [string]$b) {
    return ((Get-CaminhoCanonico $a) -eq (Get-CaminhoCanonico $b))
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

# Conexão SOMENTE LEITURA (Mode=Read): não cria .laccdb de gravação nem altera o arquivo.
function Open-BancoLeitura([string]$caminho) {
    if (-not (Test-Path -LiteralPath $caminho)) { throw ('Banco não encontrado: ' + $caminho) }
    $cn = New-Object System.Data.OleDb.OleDbConnection(
        'Provider=' + (Get-ProvedorAce) + ';Data Source=' + $caminho + ';Mode=Read;OLE DB Services=-4;Persist Security Info=False;')
    $cn.Open()
    return $cn
}

# Lê config.ambiente de um banco SEM alterá-lo (Mode=Read). Devolve o texto ('TESTE', 'PRODUCAO'...),
# ou $null se o arquivo não abre ou não tem a tabela config.
function Get-AmbienteDoBanco([string]$caminho) {
    $cn = $null
    try {
        $cn = Open-BancoLeitura $caminho
        $v = Invoke-Escalar $cn "SELECT valor FROM config WHERE chave='ambiente'" @() $null
        if ($null -eq $v) { return $null }
        return [string]$v
    } catch { return $null }
    finally { if ($cn) { $cn.Dispose() } }
}

# Trava de segurança de tudo o que APAGA ou SOBRESCREVE um banco (-Recriar, criar-banco-teste/demo):
# só se apaga arquivo inexistente ou banco com config.ambiente = TESTE. Banco de PRODUCAO, ou que
# não abre (não dá para provar que é de teste), é recusado ANTES de qualquer remoção.
# Devolve o motivo; vazio = pode substituir.
function Assert-PodeSubstituirBanco([string]$caminho) {
    if (-not (Test-Path -LiteralPath $caminho)) { return '' }
    $amb = Get-AmbienteDoBanco $caminho
    if ($null -eq $amb) {
        return ('RECUSADO: não foi possível ler config.ambiente de ' + $caminho + ' (arquivo ilegível ou sem tabela config); só se substitui banco com ambiente TESTE. Nada foi apagado.')
    }
    if ($amb -ne 'TESTE') {
        return ('RECUSADO: ' + $caminho + ' está com config.ambiente = ' + $amb + '; só se substitui banco com ambiente TESTE. Nada foi apagado.')
    }
    return ''
}

# Pasta dos bancos de homologação (execucao\homologacao): onde o desenvolvimento cria e destrói bancos fictícios.
function Get-PastaHomologacao([string]$raiz) { return (Get-PastaExecucao $raiz 'homologacao') }

# Banco de HOMOLOGAÇÃO = config.ambiente = TESTE **e** config.homologacao = 1 (declarado por quem criou o banco:
# criar-banco.ps1 -Ambiente TESTE, ou o MONTAR-FRONTEND -Teste). Todo teste que GRAVA exige isso: ambiente TESTE
# sozinho já foi escrito num banco de verdade por engano. Devolve o motivo; vazio = é de homologação.
function Assert-BancoDeHomologacao([string]$caminho) {
    $motivo = Assert-PodeSubstituirBanco $caminho
    if ($motivo) { return $motivo }
    if (-not (Test-Path -LiteralPath $caminho)) { return ('RECUSADO: o banco ' + $caminho + ' não existe.') }
    $cn = $null
    try {
        $cn = Open-BancoLeitura $caminho
        $v = Invoke-Escalar $cn "SELECT valor FROM config WHERE chave='homologacao'" @() $null
    } catch { return ('RECUSADO: não foi possível ler config.homologacao de ' + $caminho) }
    finally { if ($cn) { $cn.Dispose() } }
    if ([string]$v -ne '1') {
        return ('RECUSADO: ' + $caminho + ' está com ambiente TESTE mas sem config.homologacao = 1 (banco criado antes da spec 005 ou não declarado como de homologação). ' +
                'Recrie-o com criar-banco-teste.ps1 ou rode MONTAR-FRONTEND.bat -Teste, que declara a homologação. Nada foi gravado.')
    }
    return ''
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

# O banco está em uso (.laccdb presente)? O .laccdb pode demorar alguns instantes para sumir depois do último
# Close: espera até $esperarMs antes de decidir (ciclo 007: o ensaio recusava a própria cópia recém-fechada).
# Continua conservador: .laccdb que não some (resto de queda) conta como em uso.
function Test-BancoEmUso([string]$banco, [int]$esperarMs = 3000) {
    $laccdb = [System.IO.Path]::ChangeExtension($banco, 'laccdb')
    $limite = (Get-Date).AddMilliseconds($esperarMs)
    while ((Test-Path -LiteralPath $laccdb) -and (Get-Date) -lt $limite) { Start-Sleep -Milliseconds 200 }
    return (Test-Path -LiteralPath $laccdb)
}

# ------------------------------------------------------------------ manutenção (DEV-08)
# config.manutencao = 'estação|usuário|PID|início|motivo'. Preenchida, o front (modDB.AbrirTransacao) e o conector
# (Invoke-Transacao) recusam toda gravação; a leitura continua. Quem liga é quem mexe no banco (montagem,
# migrações) e desliga no fim. Dono que caiu NÃO libera sozinho (falha segura): a próxima montagem assume, ou
# aplicar-migracoes.ps1 -LiberarManutencao.
function New-TextoManutencao([string]$motivo) {
    return ($env:COMPUTERNAME + '|' + $env:USERNAME + '|' + $PID + '|' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + '|' + $motivo)
}
function Get-Manutencao($cn) {
    return [string](Invoke-Escalar $cn "SELECT valor FROM config WHERE chave='manutencao'" @() $null)
}
function Set-Manutencao($cn, [string]$motivo) {
    $t = New-TextoManutencao $motivo
    Set-ConfigBanco $cn 'manutencao' $t $null
    return $t
}
# Desliga. Com $soSeFor, só desliga se o sinal ainda é aquele (não apaga o de outra manutenção que assumiu).
function Clear-Manutencao($cn, [string]$soSeFor = '') {
    if ($soSeFor -and (Get-Manutencao $cn) -ne $soSeFor) { return $false }
    [void](Invoke-Comando $cn "DELETE FROM config WHERE chave='manutencao'" @() $null)
    return $true
}

# Chave de comparação de lista (C01): sem caixa, sem acento, espaços aparados e colapsados.
# Mesma regra de modValidacao.ChaveLista (front) e Regras.ps1 (conector): as três tem de concordar.
function Get-ChaveLista([string]$texto) {
    if ($null -eq $texto) { return '' }
    $s = $texto.Replace([string][char]160, ' ').Trim().ToUpperInvariant()
    while ($s.Contains('  ')) { $s = $s.Replace('  ', ' ') }
    $sb = New-Object System.Text.StringBuilder
    foreach ($c in $s.Normalize([Text.NormalizationForm]::FormD).ToCharArray()) {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($c) -ne [Globalization.UnicodeCategory]::NonSpacingMark) { [void]$sb.Append($c) }
    }
    return $sb.ToString()
}

# Onde cada tipo de lista é guardado (tabela:coluna). Única fonte: migração 003 e conferência.
# Uma lista sem coluna aqui não tem uso registrado (ex.: papel das etapas fica na própria listas).
function Get-ColunasDaLista {
    return @{
        UF = @('clientes:uf'); Segmento = @('clientes:segmento'); Qualificacao = @('clientes:qualificacao')
        Responsavel = @('clientes:responsavel', 'oportunidades:responsavel')
        Etapa = @('oportunidades:etapa'); Familia = @('oportunidades:familia'); Categoria = @('oportunidades:categoria')
        Origem = @('oportunidades:origem'); Prioridade = @('oportunidades:prioridade'); Motivo = @('oportunidades:motivo_desfecho')
    }
}
