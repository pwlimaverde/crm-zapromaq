<#
  testar-mcp.ps1 - faz o papel do Claude Desktop: inicia o servidor MCP,
  conversa com ele pelo stdio e confere as respostas.

  -SoProtocolo : só protocolo, nas duas gerações (não precisa de banco nem ACE)
                 legado  = initialize (<= 2025-11-25)
                 moderno = sem estado, _meta por requisição, server/discover (2026-07-28)
  sem a opção  : roda também leituras e gravações no banco de teste,
                 incluindo o teste de conflito de versão.
#>
param([switch]$SoProtocolo)
$ErrorActionPreference = 'Stop'

$raizIA = Split-Path -Parent $PSScriptRoot
$servidor = Join-Path $raizIA 'mcp\servidor.ps1'
$ps = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'

$psi = New-Object System.Diagnostics.ProcessStartInfo
$psi.FileName = $ps
$psi.Arguments = '-NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "' + $servidor + '"'
$psi.UseShellExecute = $false
$psi.RedirectStandardInput = $true
$psi.RedirectStandardOutput = $true
$psi.RedirectStandardError = $true
$psi.StandardOutputEncoding = New-Object System.Text.UTF8Encoding($false)
$psi.CreateNoWindow = $true
$proc = [System.Diagnostics.Process]::Start($psi)
# stderr lido em segundo plano: se ninguém lê, o buffer enche e o servidor trava
$leituraErro = $proc.StandardError.ReadToEndAsync()
$utf8 = New-Object System.Text.UTF8Encoding($false)

$script:falhas = 0
$script:id = 0
function Enviar($metodo, $params, [switch]$Notificacao) {
    $msg = [ordered]@{ jsonrpc = '2.0' }
    if (-not $Notificacao) { $script:id++; $msg['id'] = $script:id }
    $msg['method'] = $metodo
    if ($null -ne $params) { $msg['params'] = $params }
    $bytes = $utf8.GetBytes((ConvertTo-Json -InputObject $msg -Depth 20 -Compress) + "`n")
    $proc.StandardInput.BaseStream.Write($bytes, 0, $bytes.Length)
    $proc.StandardInput.BaseStream.Flush()
    if ($Notificacao) { return $null }
    $linha = $proc.StandardOutput.ReadLine()
    if ($null -eq $linha) { throw 'O servidor fechou a saída (veja o stderr abaixo).' }
    $r = $linha | ConvertFrom-Json
    if ($r.id -ne $script:id) { throw ('Resposta fora de ordem: esperado id ' + $script:id + ', veio ' + $r.id) }
    return $r
}
function Conferir([string]$nome, [bool]$ok, [string]$detalhe = '') {
    if ($ok) { Write-Host ('  OK    ' + $nome) -ForegroundColor Green }
    else { $script:falhas++; Write-Host ('  FALHA ' + $nome + '  ' + $detalhe) -ForegroundColor Red }
}
function Chamar($ferramenta, $argumentos) {
    $r = Enviar 'tools/call' ([ordered]@{ name = $ferramenta; arguments = $argumentos })
    return $r.result
}
function Dados($resultado) { return ($resultado.content[0].text | ConvertFrom-Json) }
# parâmetros de uma requisição MODERNA: versão e identidade vão em _meta, a cada chamada
function Moderno($params, [string]$versao = '2026-07-28') {
    $p = [ordered]@{}
    if ($null -ne $params) { foreach ($k in $params.Keys) { $p[$k] = $params[$k] } }
    $p['_meta'] = [ordered]@{ 'io.modelcontextprotocol/protocolVersion' = $versao
                              'io.modelcontextprotocol/clientInfo' = [ordered]@{ name = 'testar-mcp'; version = '1.0' }
                              'io.modelcontextprotocol/clientCapabilities' = [ordered]@{} }
    return $p
}

try {
    Write-Host '--- protocolo moderno (2026-07-28, sem estado)'
    $r = Enviar 'server/discover' (Moderno $null)
    Conferir 'server/discover responde DiscoverResult' ($r.result.resultType -eq 'complete' -and @($r.result.supportedVersions) -contains '2026-07-28') ($r | ConvertTo-Json -Compress -Depth 6)
    Conferir 'discover anuncia também as versões legadas' (@($r.result.supportedVersions) -contains '2025-11-25')
    Conferir 'discover traz serverInfo em _meta' ($r.result._meta.'io.modelcontextprotocol/serverInfo'.name -eq 'crm-zapromaq')
    $r = Enviar 'tools/list' (Moderno $null '1900-01-01')
    Conferir 'versão não suportada volta -32022 com a lista suportada' ($r.error.code -eq -32022 -and @($r.error.data.supported) -contains '2026-07-28') ($r | ConvertTo-Json -Compress -Depth 6)
    $r = Enviar 'tools/list' (Moderno $null)
    Conferir 'tools/list moderno: resultType, ttlMs e cacheScope' ($r.result.resultType -eq 'complete' -and $r.result.ttlMs -ge 0 -and $r.result.cacheScope -eq 'public')
    Conferir ('tools/list moderno: ' + @($r.result.tools).Count + ' ferramentas') (@($r.result.tools).Count -ge 11)
    $r = Enviar 'tools/call' (Moderno ([ordered]@{ name = 'obter_atendimento'; arguments = [ordered]@{ codigo = 'XX' } }))
    Conferir 'tools/call moderno: erro de ferramenta com resultType' ($r.result.resultType -eq 'complete' -and $r.result.isError -eq $true)
    $r = Enviar 'metodo/inexistente' (Moderno $null)
    Conferir 'método desconhecido (moderno) volta -32601' ($r.error.code -eq -32601)

    Write-Host '--- protocolo legado (initialize)'
    $r = Enviar 'initialize' ([ordered]@{ protocolVersion = '2025-06-18'; capabilities = [ordered]@{}
                                          clientInfo = [ordered]@{ name = 'testar-mcp'; version = '1.0' } })
    Conferir 'initialize devolve a versão pedida' ($r.result.protocolVersion -eq '2025-06-18') ($r | ConvertTo-Json -Compress)
    Conferir 'servidor declara ferramentas' ($null -ne $r.result.capabilities.tools)
    [void](Enviar 'notifications/initialized' $null -Notificacao)
    $r = Enviar 'initialize' ([ordered]@{ protocolVersion = '2099-01-01'; capabilities = [ordered]@{}; clientInfo = [ordered]@{ name = 'x'; version = '1' } })
    Conferir 'versão desconhecida no initialize recebe a legada mais nova' ($r.result.protocolVersion -eq '2025-11-25')
    $r = Enviar 'ping' ([ordered]@{})
    Conferir 'ping' ($null -ne $r.result)
    $r = Enviar 'tools/list' ([ordered]@{})
    $nomes = @($r.result.tools | ForEach-Object { $_.name })
    Conferir ('tools/list: ' + $nomes.Count + ' ferramentas') ($nomes.Count -ge 11)
    Conferir 'ferramentas de leitura marcadas como somente leitura' (@($r.result.tools | Where-Object { $_.name -like 'buscar_*' -and -not $_.annotations.readOnlyHint }).Count -eq 0)
    $r = Enviar 'metodo/inexistente' ([ordered]@{})
    Conferir 'método desconhecido volta erro -32601' ($r.error.code -eq -32601)
    $r = Enviar 'tools/call' ([ordered]@{ name = 'nao_existe'; arguments = [ordered]@{} })
    Conferir 'ferramenta desconhecida volta erro -32602' ($r.error.code -eq -32602)
    $res = Chamar 'obter_atendimento' ([ordered]@{ codigo = 'XX-1' })
    Conferir 'argumento fora do formato vira erro de ferramenta' ($res.isError -eq $true) $res.content[0].text
    $res = Chamar 'buscar_clientes' ([ordered]@{ campo_que_nao_existe = 1 })
    Conferir 'parâmetro desconhecido é recusado' ($res.isError -eq $true -and $res.content[0].text -match 'desconhecido')

    if (-not $SoProtocolo) {
        Write-Host '--- leitura'
        $res = Chamar 'crm_status' ([ordered]@{})
        Conferir 'crm_status' (-not $res.isError) $res.content[0].text
        if ($res.isError) { throw 'Sem acesso ao banco - rode CRIAR-BANCO-TESTE.bat antes (e confira o ACE 64 bits).' }
        $st = Dados $res
        Conferir 'gravação liberada no esquema 1.0' ($st.gravacao_liberada -eq $true)

        $d = Dados (Chamar 'buscar_clientes' ([ordered]@{ texto = "d'angelo" }))
        Conferir 'busca com apóstrofo' ($d.quantidade -eq 1)
        $d = Dados (Chamar 'buscar_atendimentos' ([ordered]@{ situacao = ('A' + [char]0x00E7 + [char]0x00E3 + 'o atrasada') }))
        Conferir 'filtro por situação calculada' (@($d.atendimentos | Where-Object { $_.codigo -eq 'AT-0001-0001' }).Count -eq 1)
        $at = Dados (Chamar 'obter_atendimento' ([ordered]@{ codigo = 'AT-0001-0001' }))
        Conferir 'obter_atendimento traz versão e empresa' ($at.versao -ge 1 -and $at.empresa -eq 'METALURGICA EXEMPLO LTDA')

        Write-Host '--- gravação'
        $amanha = (Get-Date).AddDays(1).ToString('yyyy-MM-dd')
        $res = Chamar 'atualizar_atendimento' ([ordered]@{ codigo = 'AT-0001-0001'; versao = [int]$at.versao; dt_prox_acao = $amanha
                                                           observacoes_acrescentar = 'Teste automatizado do MCP'; autor = 'teste' })
        Conferir 'atualizar_atendimento grava' (-not $res.isError) $res.content[0].text
        $g = Dados $res
        Conferir 'versão sobe 1' ($g.versao -eq ([int]$at.versao + 1))
        $res = Chamar 'atualizar_atendimento' ([ordered]@{ codigo = 'AT-0001-0001'; versao = [int]$at.versao; tentativas = 9 })
        Conferir 'versão antiga é recusada (conflito)' ($res.isError -and $res.content[0].text -match 'Conflito')
        $res = Chamar 'atualizar_atendimento' ([ordered]@{ codigo = 'AT-0001-0001'; versao = [int]$g.versao; etapa = 'Perdido' })
        Conferir 'Perdido sem motivo é recusado' ($res.isError -and $res.content[0].text -match 'motivo')
        $res = Chamar 'atualizar_atendimento' ([ordered]@{ codigo = 'AT-0001-0001'; versao = [int]$g.versao; etapa = 'Etapa Inventada' })
        Conferir 'etapa fora da lista é recusada' ($res.isError -and $res.content[0].text -match 'lista')
        $res = Chamar 'atualizar_contexto_cliente' ([ordered]@{ codigo_cliente = 1; ctx_resumo = 'Cliente de teste; resumo gravado pelo MCP' })
        Conferir 'atualizar_contexto_cliente' (-not $res.isError) $res.content[0].text
        $cli = Dados (Chamar 'obter_cliente' ([ordered]@{ codigo_cliente = 2 }))
        $res = Chamar 'atualizar_cadastro_cliente' ([ordered]@{ codigo_cliente = 2; versao = [int]$cli.versao; telefone = '123' })
        Conferir 'telefone inválido é recusado' ($res.isError -and $res.content[0].text -match 'Telefone')
        $res = Chamar 'atualizar_cadastro_cliente' ([ordered]@{ codigo_cliente = 2; versao = [int]$cli.versao; cidade = '  caxias   do sul ' })
        $g = Dados $res
        Conferir 'cidade normalizada em maiúsculas' ((-not $res.isError) -and ($g.alterados[0].depois -eq 'CAXIAS DO SUL')) $res.content[0].text
        $log = Dados (Chamar 'ultimas_alteracoes' ([ordered]@{ limite = 10 }))
        Conferir 'log de auditoria com origem IA' (@($log.alteracoes | Where-Object { $_.origem -eq 'IA' }).Count -ge 3)
    }
} catch {
    $script:falhas++
    Write-Host ('ERRO: ' + $_.Exception.Message) -ForegroundColor Red
} finally {
    try { $proc.StandardInput.Close() } catch { }
    if (-not $proc.WaitForExit(5000)) { $proc.Kill() }
    $err = $leituraErro.Result
    if ($err) { Write-Host '--- stderr do servidor'; Write-Host $err.TrimEnd() }
}
Write-Host ''
if ($script:falhas -eq 0) { Write-Host 'TUDO CERTO.' -ForegroundColor Green; exit 0 }
Write-Host ($script:falhas.ToString() + ' falha(s).') -ForegroundColor Red; exit 1
