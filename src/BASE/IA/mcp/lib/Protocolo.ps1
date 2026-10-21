<#
  Protocolo.ps1 - núcleo MCP compartilhado pelos servidores do CRM
  (IA\mcp\servidor.ps1 = dados; IA\mcp-dev\servidor-dev.ps1 = desenvolvimento).

  Transporte stdio, JSON-RPC uma mensagem por linha, UTF-8. Dual-era:
  protocolo moderno 2026-07-28 (sem estado) e legado (initialize).
  NADA além de mensagem MCP pode sair no stdout.

  Quem usa define antes de chamar Start-ServidorMcp:
    $script:InfoServidor      [ordered]@{ name; title; version }
    $script:VersaoServidor    texto da versão (log)
    $script:Instrucoes        texto de instruções para o modelo
    $script:Ferramentas       catálogo: name, title, description, inputSchema, annotations, handler
    $script:AntesDaFerramenta (opcional) scriptblock chamado com os argumentos antes do handler
  e chama Initialize-LogMcp <pasta IA> <prefixo>.
#>
$script:VersoesModernas = @('2026-07-28')
$script:VersoesLegadas = @('2025-11-25', '2025-06-18', '2025-03-26', '2024-11-05')
$script:TodasVersoes = @($script:VersoesModernas + $script:VersoesLegadas)
$script:ChaveVersao = 'io.modelcontextprotocol/protocolVersion'
if (-not (Get-Variable -Name AntesDaFerramenta -Scope Script -ErrorAction SilentlyContinue)) { $script:AntesDaFerramenta = $null }

# ------------------------------------------------------------------ log
# Diagnóstico (spec 005): um arquivo POR ESTAÇÃO, IA\logs\<prefixo>-<ESTACAO>-AAAAMMDD.log, e cada linha traz
# [ESTACAO\usuário técnico] e, nas chamadas, o nome da ferramenta. Assim, vários usuários na mesma pasta de
# rede não misturam nem sobrescrevem os logs, e dá para saber quem (qual estação/conta) chamou o quê.
# Retenção (seção 5.3, ciclo 007): mês corrente e o anterior, pela DATA DO NOME (não pela data de modificação).
# Cada linha (DEV-06 parte A): instante com fuso, PID, estação\usuário técnico ([Environment]), versão do
# servidor, ferramenta, registro (só códigos), texto/erro nativo. Sem conteúdo comercial.
# Falha ao gravar o diagnóstico não bloqueia a operação: conta em $script:DiagFalhas e aparece no crm_status.
$script:ArquivoLog = $null
$script:IdentidadeLog = [Environment]::MachineName + '\' + [Environment]::UserName
$script:DiagFalhas = 0
$script:DiagUltimaFalha = ''
$script:RegistroAtual = ''

# Arquivos de log a apagar: <prefixo>-<ESTACAO>-AAAAMMDD.log com data anterior ao 1º dia do mês anterior.
function Get-LogsVencidos([string]$pastaLogs, [string]$prefixo, [datetime]$agora) {
    $corte = (Get-Date -Year $agora.Year -Month $agora.Month -Day 1).Date.AddMonths(-1)
    return @(Get-ChildItem -LiteralPath $pastaLogs -Filter ($prefixo + '-*.log') -File -ErrorAction SilentlyContinue | Where-Object {
        $d = [datetime]::MinValue
        ($_.Name -match '-(\d{8})\.log$') -and [datetime]::TryParseExact($Matches[1], 'yyyyMMdd', [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$d) -and $d -lt $corte
    })
}

function Initialize-LogMcp([string]$raizIA, [string]$prefixo) {
    $pastaLogs = Join-Path $raizIA 'logs'
    try {
        if (-not (Test-Path $pastaLogs)) { New-Item -ItemType Directory -Path $pastaLogs | Out-Null }
        $estacao = ([Environment]::MachineName -replace '[^A-Za-z0-9_-]', '_')
        $script:ArquivoLog = Join-Path $pastaLogs ($prefixo + '-' + $estacao + '-' + (Get-Date -Format 'yyyyMMdd') + '.log')
        Get-LogsVencidos $pastaLogs $prefixo (Get-Date) | Remove-Item -Force -ErrorAction SilentlyContinue
    } catch { $script:DiagFalhas++; $script:DiagUltimaFalha = 'pasta de logs: ' + $_.Exception.Message }
}
$script:FerramentaAtual = ''

function Write-Diag([string]$texto) {
    $quem = $script:IdentidadeLog + ' v' + $script:VersaoServidor
    if ($script:FerramentaAtual) { $quem += ' ' + $script:FerramentaAtual }
    if ($script:RegistroAtual) { $quem += ' ' + $script:RegistroAtual }
    $linha = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff zzz') + ' [' + $PID + '] [' + $quem + '] ' + $texto
    try { [Console]::Error.WriteLine($linha) } catch { }
    if ($script:ArquivoLog) {
        # várias sessões (Claude e Codex) na mesma estação escrevem no mesmo arquivo: compartilhado e com nova tentativa
        $bytes = (New-Object System.Text.UTF8Encoding($false)).GetBytes($linha + "`r`n")
        foreach ($espera in @(0, 30, 80, 150, 300)) {
            if ($espera) { Start-Sleep -Milliseconds $espera }
            try {
                $fs = New-Object System.IO.FileStream($script:ArquivoLog, [System.IO.FileMode]::Append, [System.IO.FileAccess]::Write, [System.IO.FileShare]::ReadWrite)
                try { $fs.Write($bytes, 0, $bytes.Length) } finally { $fs.Dispose() }
                return
            } catch { $ultima = $_.Exception.Message }
        }
        $script:DiagFalhas++
        $script:DiagUltimaFalha = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + ' ' + $ultima
    }
}

# Só os códigos do registro (nada de conteúdo comercial): codigo, codigo_cliente, codigo_contato, ids.
function Get-RegistroDosArgumentos($a) {
    if ($null -eq $a) { return '' }
    $partes = @()
    foreach ($n in @('codigo', 'codigo_cliente', 'id_cliente', 'codigo_contato', 'id_contato', 'id', 'id_planejamento')) {
        if ($a.PSObject.Properties.Name -contains $n -and $null -ne $a.$n) { $partes += ($n + '=' + [string]$a.$n) }
    }
    return ($partes -join ',')
}

# ------------------------------------------------------------------ E/S
$utf8 = New-Object System.Text.UTF8Encoding($false)
$script:Entrada = New-Object System.IO.StreamReader([Console]::OpenStandardInput(), $utf8)
$script:Saida = New-Object System.IO.StreamWriter([Console]::OpenStandardOutput(), $utf8)
$script:Saida.AutoFlush = $true
$script:Saida.NewLine = "`n"

function Send-Mensagem($obj) {
    # -Compress garante uma linha só: quebras de linha dos textos saem como \n
    $json = ConvertTo-Json -InputObject $obj -Depth 30 -Compress
    $script:Saida.WriteLine($json)
}

function New-ErroRpc([int]$codigo, [string]$mensagem, $dados = $null) {
    $ex = New-Object System.Exception($mensagem)
    $ex.Data['rpc'] = $codigo
    if ($null -ne $dados) { $ex.Data['dados'] = $dados }
    return $ex
}

function Send-Erro($id, [int]$codigo, [string]$mensagem, $dados = $null) {
    $erro = [ordered]@{ code = $codigo; message = $mensagem }
    if ($null -ne $dados) { $erro['data'] = $dados }
    Send-Mensagem ([ordered]@{ jsonrpc = '2.0'; id = $id; error = $erro })
}

# ------------------------------------------------------------------ validação de argumentos
# Confere o essencial do inputSchema: campos desconhecidos, obrigatórios,
# tipo, enum, tamanho e padrão. Erro aqui volta como erro de ferramenta,
# para a IA ler a mensagem e corrigir a chamada.
function Test-Argumentos($schema, $a) {
    $props = $schema.properties
    # objeto vazio ({}) devolve $null em .Properties.Name - por isso o ForEach
    $nomes = @(); if ($null -ne $a) { $nomes = @($a.PSObject.Properties | ForEach-Object { $_.Name }) }
    foreach ($n in $nomes) {
        if (-not $props.Contains($n)) {
            throw ("Parâmetro desconhecido: '" + $n + "'. Aceitos: " + (@($props.Keys) -join ', '))
        }
    }
    foreach ($r in @($schema.required)) {
        if ($r -and $nomes -notcontains $r) { throw ("Parâmetro obrigatório ausente: '" + $r + "'.") }
    }
    foreach ($n in $nomes) {
        $def = $props[$n]; $v = $a.$n
        $tipos = @($def.type)
        if ($null -eq $v) {
            if ($tipos -notcontains 'null') { throw ("Parâmetro '" + $n + "' não aceita null.") }
            continue
        }
        $ok = $false
        foreach ($t in $tipos) {
            switch ($t) {
                'string'  { if ($v -is [string]) { $ok = $true } }
                'integer' { if ($v -is [int] -or $v -is [long]) { $ok = $true } }
                'number'  { if ($v -is [int] -or $v -is [long] -or $v -is [double] -or $v -is [decimal]) { $ok = $true } }
                'boolean' { if ($v -is [bool]) { $ok = $true } }
                'array'   { if ($v -is [System.Collections.IEnumerable] -and -not ($v -is [string])) { $ok = $true } }
            }
        }
        if (-not $ok) { throw ("Parâmetro '" + $n + "' deve ser do tipo " + ($tipos -join ' ou ') + '.') }
        if ($tipos -contains 'array' -and $ok) {
            $itens = @($v)
            if ($itens.Count -eq 0) { throw ("Parâmetro '" + $n + "' não pode ser uma lista vazia.") }
            if ($def.Contains('maxItems') -and $itens.Count -gt $def.maxItems) { throw ("Parâmetro '" + $n + "' aceita no máximo " + $def.maxItems + ' itens.') }
            $tipoItem = ''; $maxItem = 0
            if ($def.Contains('items')) { $tipoItem = [string]$def.items.type; if ($def.items.Contains('maxLength')) { $maxItem = [int]$def.items.maxLength } }
            foreach ($i in $itens) {
                if ($tipoItem -eq 'string' -and -not ($i -is [string])) { throw ("Parâmetro '" + $n + "': cada item deve ser texto.") }
                if ($maxItem -gt 0 -and ([string]$i).Length -gt $maxItem) { throw ("Parâmetro '" + $n + "': item passa de " + $maxItem + ' caracteres.') }
            }
            continue
        }
        if ($def.Contains('enum') -and @($def.enum) -notcontains $v) {
            throw ("Parâmetro '" + $n + "' deve ser um de: " + (@($def.enum) -join ' | '))
        }
        if ($v -is [string]) {
            if ($def.Contains('maxLength') -and $v.Length -gt $def.maxLength) { throw ("Parâmetro '" + $n + "' passa de " + $def.maxLength + ' caracteres.') }
            if ($def.Contains('pattern') -and $v -notmatch $def.pattern) { throw ("Parâmetro '" + $n + "' fora do formato esperado (" + $def.pattern + ').') }
        }
        if ($v -is [int] -or $v -is [long] -or $v -is [double] -or $v -is [decimal]) {
            if ($def.Contains('minimum') -and $v -lt $def.minimum) { throw ("Parâmetro '" + $n + "' deve ser no mínimo " + $def.minimum + '.') }
            if ($def.Contains('maximum') -and $v -gt $def.maximum) { throw ("Parâmetro '" + $n + "' deve ser no máximo " + $def.maximum + '.') }
        }
    }
}

# ------------------------------------------------------------------ métodos
# Versão moderna pedida na requisição ($null quando a requisição é do jeito legado).
function Get-VersaoModerna($params) {
    if ($null -eq $params) { return $null }
    if (-not ($params.PSObject.Properties.Name -contains '_meta')) { return $null }
    $meta = $params._meta
    if ($null -eq $meta) { return $null }
    $p = $meta.PSObject.Properties[$script:ChaveVersao]
    if ($null -eq $p) { return $null }
    return [string]$p.Value
}

function Get-ListaFerramentas {
    return @(foreach ($f in $script:Ferramentas) {
        [ordered]@{ name = $f.name; title = $f.title; description = $f.description
                    inputSchema = $f.inputSchema; annotations = $f.annotations }
    })
}

function Invoke-Ferramenta($params) {
    $nome = [string]$params.name
    $f = $script:Ferramentas | Where-Object { $_.name -eq $nome } | Select-Object -First 1
    if (-not $f) { throw (New-ErroRpc -32602 ('Ferramenta desconhecida: ' + $nome)) }
    $a = $null
    if ($params.PSObject.Properties.Name -contains 'arguments') { $a = $params.arguments }
    $inicio = Get-Date
    $script:FerramentaAtual = $nome
    $script:RegistroAtual = ''
    try { $script:RegistroAtual = Get-RegistroDosArgumentos $a } catch { }
    try {
        # erro de validação volta como erro DE FERRAMENTA, para o modelo ler e corrigir
        Test-Argumentos $f.inputSchema $a
        if ($script:AntesDaFerramenta) { & $script:AntesDaFerramenta $a $nome }
        $resultado = & $f.handler $a
        Write-Diag ('OK   ' + $nome + ' (' + [int]((Get-Date) - $inicio).TotalMilliseconds + ' ms)')
        # ferramenta que devolve imagem (ou mais de um bloco) monta o próprio
        # content em '__conteudo'; o resto vira um bloco de texto com o JSON
        if ($resultado -is [System.Collections.IDictionary] -and $resultado.Contains('__conteudo')) {
            return [ordered]@{ content = @($resultado['__conteudo']); isError = $false }
        }
        $json = ConvertTo-Json -InputObject $resultado -Depth 20
        return [ordered]@{ content = @([ordered]@{ type = 'text'; text = $json }); isError = $false }
    } catch {
        # erro lançado dentro de método de objeto ($ctx.Executar...) vem embrulhado em "Exceção ao chamar ...":
        # a IA precisa ler a mensagem original
        $ex = $_.Exception
        while ($ex -is [System.Management.Automation.MethodInvocationException] -and $null -ne $ex.InnerException) { $ex = $ex.InnerException }
        $m = $ex.Message
        # erro nativo: tipo e código (HRESULT) junto da mensagem
        Write-Diag ('ERRO ' + $nome + ': ' + $m + ' [' + $ex.GetType().Name + ' 0x' + ('{0:X8}' -f $ex.HResult) + ']')
        return [ordered]@{ content = @([ordered]@{ type = 'text'; text = $m }); isError = $true }
    } finally { $script:FerramentaAtual = ''; $script:RegistroAtual = '' }
}

function Invoke-Metodo($msg) {
    $params = $null
    if ($msg.PSObject.Properties.Name -contains 'params') { $params = $msg.params }
    $metodo = [string]$msg.method
    $versao = Get-VersaoModerna $params

    # ---------------------------------------------------------- MODERNO (2026-07-28)
    if ($null -ne $versao -or $metodo -eq 'server/discover') {
        if ($null -ne $versao -and $script:VersoesModernas -notcontains $versao) {
            throw (New-ErroRpc -32022 'Unsupported protocol version' ([ordered]@{ supported = $script:TodasVersoes; requested = $versao }))
        }
        $meta = [ordered]@{ 'io.modelcontextprotocol/serverInfo' = $script:InfoServidor }
        switch ($metodo) {
            'server/discover' {
                return [ordered]@{
                    resultType = 'complete'; supportedVersions = $script:TodasVersoes
                    capabilities = [ordered]@{ tools = [ordered]@{} }
                    _meta = $meta; instructions = $script:Instrucoes
                    ttlMs = 3600000; cacheScope = 'public'
                }
            }
            'tools/list' {
                # a lista não varia por conexão nem por usuário: cacheável por 1 h
                return [ordered]@{ resultType = 'complete'; tools = (Get-ListaFerramentas); _meta = $meta; ttlMs = 3600000; cacheScope = 'public' }
            }
            'tools/call' {
                $r = Invoke-Ferramenta $params
                $saida = [ordered]@{ resultType = 'complete' }
                foreach ($k in $r.Keys) { $saida[$k] = $r[$k] }
                $saida['_meta'] = $meta
                return $saida
            }
            default { throw (New-ErroRpc -32601 ('Método não suportado: ' + $metodo)) }
        }
    }

    # ---------------------------------------------------------- LEGADO (<= 2025-11-25)
    switch ($metodo) {
        'initialize' {
            $pedido = [string]$params.protocolVersion
            $resposta = if ($script:VersoesLegadas -contains $pedido) { $pedido } else { $script:VersoesLegadas[0] }
            Write-Diag ('initialize (legado): cliente pediu ' + $pedido + ', respondendo ' + $resposta)
            return [ordered]@{
                protocolVersion = $resposta
                capabilities = [ordered]@{ tools = [ordered]@{ listChanged = $false } }
                serverInfo = $script:InfoServidor
                instructions = $script:Instrucoes
            }
        }
        'ping' { return [ordered]@{} }
        'tools/list' { return [ordered]@{ tools = (Get-ListaFerramentas) } }
        'tools/call' { return (Invoke-Ferramenta $params) }
        default { throw (New-ErroRpc -32601 ('Método não suportado: ' + $metodo)) }
    }
}

# ------------------------------------------------------------------ laço principal
function Start-ServidorMcp {
    Write-Diag ('servidor ' + $script:InfoServidor.name + ' ' + $script:VersaoServidor + ' iniciado')
    while ($true) {
        $linha = $script:Entrada.ReadLine()
        if ($null -eq $linha) { break }            # cliente fechou o stdin: encerrar
        if ([string]::IsNullOrWhiteSpace($linha)) { continue }
        try { $msg = $linha | ConvertFrom-Json }
        catch { Send-Erro $null -32700 'JSON inválido'; continue }

        $temId = $msg.PSObject.Properties.Name -contains 'id'
        $temMetodo = $msg.PSObject.Properties.Name -contains 'method'
        if (-not $temMetodo) { continue }          # resposta do cliente: nada a fazer
        if (-not $temId) {                         # notificação: não tem resposta
            Write-Diag ('notificação: ' + $msg.method)
            continue
        }
        try {
            $res = Invoke-Metodo $msg
            Send-Mensagem ([ordered]@{ jsonrpc = '2.0'; id = $msg.id; result = $res })
        } catch {
            $ex = $_.Exception
            $cod = if ($ex.Data.Contains('rpc')) { [int]$ex.Data['rpc'] } else { -32603 }
            $dados = if ($ex.Data.Contains('dados')) { $ex.Data['dados'] } else { $null }
            Write-Diag ('erro RPC ' + $cod + ': ' + $ex.Message)
            Send-Erro $msg.id $cod $ex.Message $dados
        }
    }
    Write-Diag 'stdin fechado; servidor encerrado'
}
