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
# 1.1 = migração 003 (municípios, planejamentos, visitas, papel das listas). O conector lê colunas que só
# existem no 1.1, então exige 1.1: com o banco ainda no 1.0 (entre extrair o pacote e terminar a montagem)
# só o crm_status responde, dizendo que a atualização está pendente (spec 004, F9).
# 1.2 = migração 004 (ciclo 010): controle único, ctx_versao, autor técnico no log, tabela operacoes.
$script:EsquemasSuportados = @('1.2')

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

# Papel das etapas (spec 003, D-c): le a tabela listas (Encerra, Resultado, Quadro, Exige familia).
# Nenhum nome de etapa fica no codigo. Cache de 5 s; Limpar-PapelEtapas zera na hora (gerenciar_lista).
$script:PapelEtapas = $null
$script:PapelEtapasLidoEm = [datetime]::MinValue
function Get-PapelEtapas {
    if ($null -eq $script:PapelEtapas -or ((Get-Date) - $script:PapelEtapasLidoEm).TotalSeconds -ge 5) {
        $mapa = @{}
        $filas = Invoke-Consulta "SELECT chave, valor, encerra, resultado, quadro, exige_familia FROM listas WHERE tipo='Etapa'" @()
        foreach ($l in $filas) {
            $mapa[[string]$l.chave] = @{
                valor         = [string]$l.valor
                encerra       = [bool]$l.encerra
                resultado     = [string]$l.resultado
                quadro        = $(if ([string]::IsNullOrEmpty([string]$l.quadro)) { 'FUNIL' } else { [string]$l.quadro })
                exige_familia = [bool]$l.exige_familia
            }
        }
        $script:PapelEtapas = $mapa
        $script:PapelEtapasLidoEm = Get-Date
    }
    return $script:PapelEtapas
}
function Limpar-PapelEtapas { $script:PapelEtapas = $null }

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
function Invoke-TransacaoUmaVez([scriptblock]$bloco) {
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
        $ctx | Add-Member -NotePropertyName Conferida -NotePropertyValue $false
        # DEV-08: antes da PRIMEIRA escrita, dentro da transação, confere manutenção e esquema. Leitura em
        # transação (candidatos, roteiro) continua liberada durante a manutenção.
        $ctx | Add-Member -MemberType ScriptMethod -Name ConferirEscrita -Value {
            if ($this.Conferida) { return }
            $c = New-Comando $this.Cn "SELECT chave, valor FROM config WHERE chave IN ('versao_esquema','manutencao')" @() $this.Tx
            try { $linhas = Read-Linhas $c } finally { $c.Dispose() }
            $esq = ''; $man = ''
            foreach ($l in $linhas) { if ($l.chave -eq 'versao_esquema') { $esq = [string]$l.valor } else { $man = [string]$l.valor } }
            $motivo = Get-MotivoRecusaEscrita $esq $man
            if ($motivo) { throw $motivo }
            $this.Conferida = $true
        }
        $ctx | Add-Member -MemberType ScriptMethod -Name Executar -Value {
            param($sql, $params)
            $this.ConferirEscrita()
            $c = New-Comando $this.Cn $sql $params $this.Tx
            try { return $c.ExecuteNonQuery() } finally { $c.Dispose() }
        }
        $ctx | Add-Member -MemberType ScriptMethod -Name Log -Value {
            param($tabela, $idRegistro, $campo, $antigo, $novo, $acao)
            $this.ConferirEscrita()
            $sql ='INSERT INTO log_alteracoes (tabela, id_registro, campo, valor_antigo, valor_novo,' +
                   ' acao, usuario, quando, origem) VALUES (?,?,?,?,?,?,?,?,?)'
            $c = New-Comando $this.Cn $sql @(
                (P 'texto' $tabela), (P 'inteiro' $idRegistro), (P 'texto' $campo),
                (P 'memo' (Format-ValorLog $antigo)), (P 'memo' (Format-ValorLog $novo)),
                (P 'texto' $acao), (P 'texto' $script:UsuarioLog), (P 'data' (Get-Date)),
                (P 'texto' 'IA')) $this.Tx
            try { [void]$c.ExecuteNonQuery() } finally { $c.Dispose() }
        }
        try {
            $resultado = & $bloco $ctx
        } catch {
            # DEV-03: durante a cópia exclusiva do backup o ACE abre só leitura e a escrita falha; quem pediu
            # recebe a mensagem controlada, não o erro cru do motor
            $controlada = Get-MensagemEscritaRecusada (Get-MensagemInterna $_.Exception) (Test-BackupEmCurso)
            if ($controlada) { throw $controlada }
            throw
        }
        # DEV-13: falha no PRÓPRIO commit = resultado indeterminado (pode ter gravado): nunca repetir sozinho
        try { $tx.Commit() }
        catch {
            $tx = $null
            $ex = New-Object System.Exception(('Resultado indeterminado: a confirmação da gravação falhou (' + (Get-MensagemInterna $_.Exception) +
                ') e ela pode ter sido concluída. Leia o registro de novo (obter_*) e confira antes de repetir.'))
            $ex.Data['indeterminado'] = $true
            throw $ex
        }
        $tx = $null
        return $resultado
    } finally {
        if ($tx) { try { $tx.Rollback() } catch { } }
        # depois do commit, falha ao fechar não muda o resultado (vai só para o diagnóstico)
        try { $cn.Close(); $cn.Dispose() } catch { Write-DiagSeHouver ('fechar conexão: ' + $_.Exception.Message) }
    }
}

function Write-DiagSeHouver([string]$t) { if (Get-Command Write-Diag -ErrorAction SilentlyContinue) { Write-Diag $t } }

# DEV-14: erro transitório = registro/página bloqueada por outra conexão. Só ele se repete.
function Test-ErroTransitorio([string]$mensagem) {
    return ($mensagem -match 'bloquead|locked|could not update|não foi possível atualizar|nao foi possivel atualizar')
}

<#
  Executa o bloco numa transação (Invoke-TransacaoUmaVez) e REPETE a transação inteira quando ela falhou
  ANTES do commit por erro transitório (DEV-14): o rollback já aconteceu, então repetir não duplica nada.
  Espera crescente (250, 500, 1000, 2000 ms), no máximo 5 tentativas. Esgotado: "não gravou, pode repetir".
  Nunca repete erro de regra, conflito de versão, resultado indeterminado ou recusa por backup/manutenção.
#>
function Invoke-Transacao([scriptblock]$bloco) {
    $esperas = @(250, 500, 1000, 2000)
    for ($tentativaTx = 0; ; $tentativaTx++) {
        try { return (Invoke-TransacaoUmaVez $bloco) }
        catch {
            $ex = $_.Exception
            $msg = Get-MensagemInterna $ex
            $transitorio = (-not $ex.Data.Contains('indeterminado')) -and (Test-ErroTransitorio $msg)
            if (-not $transitorio) { throw }
            if ($tentativaTx -ge $esperas.Count) {
                throw ('Não gravou: o registro continuou bloqueado por outra gravação depois de ' + ($tentativaTx + 1) + ' tentativas (' + $msg +
                       '). Nada foi gravado; pode repetir em instantes.')
            }
            Write-DiagSeHouver ('transitório (tentativa ' + ($tentativaTx + 1) + '): ' + $msg)
            Start-Sleep -Milliseconds $esperas[$tentativaTx]
        }
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

# Todas as ferramentas, menos o crm_status, exigem um esquema suportado (leitura e gravação). Devolve o
# motivo; vazio = pode seguir. Erro ao ler o esquema não é motivo para recusar aqui: o handler relata o seu.
function Get-MotivoEsquemaNaoSuportado {
    try { $v = Get-VersaoEsquema } catch { return '' }
    if ($script:EsquemasSuportados -contains $v) { return '' }
    return ("Banco aguardando atualizacao: ele esta no esquema " + $v + " e este conector exige o esquema " +
            ($script:EsquemasSuportados -join ', ') + '. A atualizacao do banco e feita pela montagem (MONTAR-FRONTEND.bat) na estacao de implantacao; ' +
            'enquanto isso so o crm_status responde. Nada foi lido nem gravado.')
}

# Mensagem original de um erro lançado dentro de método de objeto ("Exceção ao chamar ..." embrulha).
function Get-MensagemInterna([System.Exception]$ex) {
    while ($ex -is [System.Management.Automation.MethodInvocationException] -and $null -ne $ex.InnerException) { $ex = $ex.InnerException }
    return $ex.Message
}

# DEV-03: a cópia exclusiva do backup deixa a marca execucao\backup\copiando.txt enquanto segura o arquivo.
function Test-BackupEmCurso {
    $marca = Join-Path (Split-Path -Parent $script:CaminhoBanco) 'SISTEMA\DADOS\execucao\backup\copiando.txt'
    return (Test-Path -LiteralPath $marca)
}

# DEV-03: erro de escrita com o backup segurando o arquivo vira a mensagem controlada (vazio = não é o caso).
# Mesma regra de modDB.MensagemEscritaRecusada (front).
function Get-MensagemEscritaRecusada([string]$erro, [bool]$backupEmCurso) {
    if (-not $backupEmCurso) { return '' }
    if ($erro -notmatch 'atualiz|updat|bloquead|locked|em uso|in use|somente leitura|read-only') { return '' }
    return ('Cópia de segurança (backup) do banco em andamento: gravação recusada, nada foi gravado. ' +
            'A leitura continua normal; tente gravar de novo em um ou dois minutos.')
}

# DEV-08: motivo para recusar uma escrita (vazio = pode gravar). Mesma regra de modDB.MotivoRecusaEscrita (front).
# config.manutencao = 'estação|usuário|PID|início|motivo', ligada pela montagem/migrações enquanto mexem no banco.
function Get-MotivoRecusaEscrita([string]$esquema, [string]$manutencao) {
    if (-not [string]::IsNullOrWhiteSpace($manutencao)) {
        $p = $manutencao.Split('|')
        $quem = if ($p.Count -ge 5) { ' (' + $p[4] + ', por ' + $p[0] + '\' + $p[1] + ' desde ' + $p[3] + ')' } else { '' }
        return ('Sistema em manutenção' + $quem + ': gravação recusada, nada foi gravado. Tente de novo quando a manutenção terminar.')
    }
    if ($script:EsquemasSuportados -notcontains $esquema) {
        return ("Gravação bloqueada: o banco está no esquema '" + $esquema + "' e este conector conhece " +
                ($script:EsquemasSuportados -join ', ') + '. Nada foi gravado. Atualize a pasta IA junto com o sistema.')
    }
    return ''
}

function Assert-EsquemaGravavel {
    $v = Get-VersaoEsquema
    if ($script:EsquemasSuportados -notcontains $v) {
        throw ("Gravacao bloqueada: o banco esta no esquema '" + $v + "' e este servidor conhece " +
               ($script:EsquemasSuportados -join ', ') + '. Atualize a pasta IA junto com o sistema.')
    }
}
