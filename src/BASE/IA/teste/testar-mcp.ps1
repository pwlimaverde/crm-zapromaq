<#
  testar-mcp.ps1 - faz o papel do Claude Desktop: inicia o servidor MCP,
  conversa com ele pelo stdio e confere as respostas.

  -SoProtocolo : só protocolo, nas duas gerações (não precisa de banco nem ACE)
                 legado  = initialize (<= 2025-11-25)
                 moderno = sem estado, _meta por requisição, server/discover (2026-07-28)
  sem a opção  : roda também leituras e gravações no banco de teste,
                 incluindo o teste de conflito de versão.
#>
#  -Servidor <servidor.ps1> e -Banco <accdb>: testam outra pasta de conector (publicar-conector.ps1, DEV-05)
#  contra outro banco de homologação; sem eles, IA\mcp\servidor.ps1 e o banco da raiz de BASE.
param([switch]$SoProtocolo, [string]$Servidor = '', [string]$Banco = '')
$ErrorActionPreference = 'Stop'

$raizIA = Split-Path -Parent $PSScriptRoot
$servidor = if ($Servidor) { $Servidor } else { Join-Path $raizIA 'mcp\servidor.ps1' }
$ps = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'

$psi = New-Object System.Diagnostics.ProcessStartInfo
$psi.FileName = $ps
$psi.Arguments = '-NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "' + $servidor + '"' + $(if ($Banco) { ' -Banco "' + $Banco + '"' } else { '' })
# servidor de uma pasta de versão encontra a pasta IA (logs) por aqui, como no lançador
$psi.EnvironmentVariables['CRM_RAIZ_IA'] = $raizIA
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
# DEV-07: o contrato do conector é ESTA lista. Mudou o catálogo? Muda aqui, com nota de mudança no IA\LEIA-ME.md
# (ferramenta, parâmetro, resposta e erro) e VersaoServidor novo.
$script:FerramentasEsperadas = @(
    'abrir_atendimento', 'alterar_situacao_cliente', 'alterar_situacao_planejamento', 'alterar_situacao_visita',
    'atendimentos_por_interacao', 'atendimentos_por_proxima_acao', 'atualizar_atendimento', 'atualizar_cadastro_cliente',
    'atualizar_contato', 'atualizar_contexto_atendimento', 'atualizar_contexto_cliente', 'atualizar_planejamento',
    'atualizar_visita', 'buscar_atendimentos', 'buscar_clientes', 'buscar_contatos', 'buscar_municipios',
    'criar_cliente', 'criar_contato', 'criar_planejamento', 'criar_visita', 'crm_status', 'descartar_extratos', 'duplicar_atendimento', 'exportar_base',
    'exportar_roteiro', 'gerar_visitas', 'gerenciar_lista', 'gravar_meta', 'lancar_datas', 'lancar_resultados',
    'listar_candidatos', 'listar_metas', 'listar_opcoes', 'listar_planejamentos', 'listar_visitas', 'obter_atendimento',
    'obter_cliente', 'obter_contato', 'obter_planejamento', 'promover_cliente', 'resolver_cidade', 'ultimas_alteracoes')
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

Write-Host '--- regras de lista (C01): chave, sem servidor'
. (Join-Path $raizIA 'mcp\lib\Regras.ps1')
Conferir 'C01 chave ignora espaços sobrando e caixa' ((ConvertTo-ChaveLista 'pedido  fechado ') -eq 'PEDIDO FECHADO')
Conferir 'C01 chave ignora acento (Pedído)' ((ConvertTo-ChaveLista 'Pedído Fechado') -eq 'PEDIDO FECHADO')
Conferir 'C01 chave tira til e cedilha' ((ConvertTo-ChaveLista 'Negociação Ç Ñ') -eq 'NEGOCIACAO C N')
Conferir 'C01 chave de valor vazio é vazia' ((ConvertTo-ChaveLista '   ') -eq '')

Write-Host '--- valor inativo (C06, D-j): só o que já está gravado no registro passa'
. (Join-Path $raizIA 'mcp\lib\Protocolo.ps1')
. (Join-Path $raizIA 'mcp\lib\Banco.ps1')
. (Join-Path $raizIA 'mcp\lib\Ferramentas.ps1')
$script:CaminhoBanco = if ($Banco) { $Banco } else { Join-Path (Split-Path -Parent $raizIA) 'crm_zapromaq.accdb' }
# Trava (spec 004, F2): daqui em diante o teste GRAVA no banco. Só roda se config.ambiente = TESTE;
# banco de PRODUCAO (ou que não se lê) é recusado antes de qualquer gravação. Fora de processo,
# para não misturar as funções de SISTEMA\DADOS\lib\Comum.ps1 com as do conector.
if (Test-Path -LiteralPath $script:CaminhoBanco) {
    $env:CRM_TRAVA_COMUM = Join-Path (Split-Path -Parent $raizIA) 'SISTEMA\DADOS\lib\Comum.ps1'
    $env:CRM_TRAVA_BANCO = $script:CaminhoBanco
    $motivoTrava = (& $ps -NoProfile -ExecutionPolicy Bypass -Command '. $env:CRM_TRAVA_COMUM; Assert-BancoDeHomologacao $env:CRM_TRAVA_BANCO' 2>&1 | Out-String).Trim()
    Remove-Item Env:CRM_TRAVA_COMUM, Env:CRM_TRAVA_BANCO -ErrorAction SilentlyContinue
    if ($motivoTrava) {
        Write-Host ('RECUSADO: o testar-mcp grava no banco e só roda em banco de TESTE. ' + $motivoTrava) -ForegroundColor Red
        try { $proc.StandardInput.Close(); if (-not $proc.WaitForExit(3000)) { $proc.Kill() } } catch { }
        exit 2
    }
}
function Recusa([scriptblock]$bloco) { try { $null = & $bloco; return [bool]$false } catch { return [bool]$true } }
# o banco de demonstração não tem etapa inativa: desativa 'Perdido' só durante estes casos e restaura
function Set-AtivoEtapa([string]$valor, [bool]$ativo) {
    $cn = Open-Conexao
    try { $c = New-Comando $cn 'UPDATE listas SET ativo=? WHERE tipo=? AND valor=?' @((P 'logico' $ativo), (P 'texto' 'Etapa'), (P 'texto' $valor)) $null; [void]$c.ExecuteNonQuery(); $c.Dispose() }
    finally { $cn.Close(); $cn.Dispose() }
}
try {
    Set-AtivoEtapa 'Perdido' $false
    Conferir 'C06 etapa ativa passa' ((Resolve-ValorLista $null 'Etapa' 'etapa' 'Negociação') -eq 'Negociação')
    Conferir 'C06 etapa inativa não gravada é recusada' (Recusa { Resolve-ValorLista $null 'Etapa' 'etapa' 'Perdido' 'Negociação' })
    Conferir 'C06 etapa inativa já gravada no registro passa' ((Resolve-ValorLista $null 'Etapa' 'etapa' 'Perdido' 'Perdido') -eq 'Perdido')
    Conferir 'C06 inativa gravada aceita o próprio texto, sem diferença de caixa' ((Resolve-ValorLista $null 'Etapa' 'etapa' 'perdido' 'Perdido') -eq 'Perdido')
    Conferir 'C06 inativa que é outra gravada no registro é recusada' (Recusa { Resolve-ValorLista $null 'Etapa' 'etapa' 'Perdido' 'Descartado' })
} finally { Set-AtivoEtapa 'Perdido' $true }

Write-Host '--- DEV-09: conversão sem perda silenciosa (mesma matriz C17-C21 do front)'
Conferir 'C17 data 2026-02-31 recusada' (Recusa { ConvertTo-Data 'dt_proposta' '2026-02-31' })
Conferir 'C18 moeda com 3 casas recusada (12.505)' (Recusa { ConvertTo-Moeda 'valor' 12.505 })
Conferir 'C18 moeda com 2 casas passa (12.5)' ((ConvertTo-Moeda 'valor' 12.5) -eq [decimal]12.5)
Conferir 'C19 inteiro com fração recusado (1.5)' (Recusa { ConvertTo-Inteiro 'tentativas' 1.5 })
Conferir 'C19 inteiro sem fração passa (3.0 = 3)' ((ConvertTo-Inteiro 'tentativas' 3.0) -eq 3)
Conferir 'C20 moeda acima do limite recusada' (Recusa { ConvertTo-Moeda 'valor' 1000000000000 })
Conferir 'C21 valor negativo recusado' (Recusa { ConvertTo-CampoAtendimento $null 'valor' -10 })
Conferir 'inteiro fora do Long recusado' (Recusa { ConvertTo-Inteiro 'meta_pedidos' 2147483648 })
Conferir 'moeda nula continua nula (limpar o campo)' ($null -eq (ConvertTo-Moeda 'valor' $null))

Write-Host '--- DEV-10: normalização da seção 5.1 (mesmos casos C22-C24 do front)'
Conferir 'C22 nome em maiúsculas, acento preservado e espaço colapsado' ((ConvertTo-Normalizado 'nome' ('  jo' + [char]0x00E3 + 'o   da  silva ')) -ceq ('JO' + [char]0x00C3 + 'O DA SILVA'))
Conferir 'C22 observações: maiúsculas e quebra de linha preservada' ((ConvertTo-Normalizado 'observacoes' ("linha  um  `r`n`r`n  linha dois")) -ceq "LINHA UM`r`n`r`nLINHA DOIS")
Conferir 'C22 campo de lista fica como na lista' ((ConvertTo-Normalizado 'etapa' ' Proposta em Análise ') -ceq 'Proposta em Análise')
$script:EmailsAdicionais = ''
Conferir 'C23 e-mail minúsculo e só o primeiro endereço' ((ConvertTo-Normalizado 'email' 'Ana@X.COM; b@y.com') -ceq 'ana@x.com' -and $script:EmailsAdicionais -eq 'b@y.com')
Conferir 'C23 e-mail com um endereço só fica inteiro' ((ConvertTo-Normalizado 'email' ' Compras@Exemplo.com.BR ') -ceq 'compras@exemplo.com.br' -and $script:EmailsAdicionais -eq '')
Conferir 'C23 e-mail adicional vai para as observações' ((Add-EmailsAdicionais 'OBS' 'b@y.com') -ceq "OBS`r`nE-MAIL ADICIONAL: b@y.com")
Conferir 'C24 URL, caminho e e-mail dentro das observações ficam como digitados' ((ConvertTo-Normalizado 'observacoes' 'ver http://site.com/Ab e C:\Pasta\Arq.txt com Ana@Ex.com') -ceq 'VER http://site.com/Ab E C:\Pasta\Arq.txt COM Ana@Ex.com')
Conferir 'C24 caminho de contexto como gravado' ((ConvertTo-Normalizado 'ctx_arquivo' ' Clientes\0001\Contexto.md ') -ceq 'Clientes\0001\Contexto.md')

Write-Host '--- DEV-10: transições de etapa e retomada (mesmos casos C25-C28 do front)'
$trP = Get-TransicaoEtapa $false $true $true
Conferir 'C25 encerrar em Perdido/Descartado limpa a próxima ação e a data' ((@($trP.limpar) -join ',') -eq 'prox_acao,dt_prox_acao' -and $trP.orientacao -eq '')
Conferir 'C25 encerrar em Pedido Fechado (GANHO) não limpa a próxima ação' (@((Get-TransicaoEtapa $false $true $false).limpar).Count -eq 0)
Conferir 'C26 retomada em etapa aberta é recusada' ((Test-RetomadaPermitida $false $false '2026-12-01' $null) -match 'C26')
Conferir 'C26 retomada em Perdido/Descartado passa' ($null -eq (Test-RetomadaPermitida $true $true '2026-12-01' $null))
Conferir 'C26 retomada antiga que não mudou não trava a edição' ($null -eq (Test-RetomadaPermitida $false $false '2026-12-01' '2026-12-01'))
$trR = Get-TransicaoEtapa $true $false $false
Conferir 'C27 reabrir limpa motivo, desfecho e retomada e orienta' ((@($trR.limpar) -join ',') -eq 'motivo_desfecho,dt_desfecho,retomar_em' -and $trR.orientacao -match 'atendimento NOVO')
$avC28 = Get-CriticaAtendimento ([ordered]@{ etapa = 'Negociação'; responsavel = 'Vendedor A'; dt_entrada = '2026-10-10'; dt_proposta = '2026-10-01'; familia = 'X'; dt_prox_acao = '2026-12-01' })
Conferir 'C28 proposta antes da entrada é aviso, não recusa' (@($avC28 | Where-Object { $_ -match 'proposta' }).Count -eq 1)

Write-Host '--- DEV-11: contato do atendimento (mesmo caso C29 do front)'
Conferir 'C29 contato de outra empresa é recusado' ((Test-ContatoDoAtendimento $true $true 7 8) -match 'outra empresa')
Conferir 'C29 contato inativo é recusado' ((Test-ContatoDoAtendimento $true $false 7 7) -match 'inativo')
Conferir 'C29 contato ativo da mesma empresa passa' ($null -eq (Test-ContatoDoAtendimento $true $true 7 7))
$schemaCt = ($script:Ferramentas | Where-Object { $_.name -eq 'atualizar_contato' }).inputSchema.properties
Conferir 'DEV-11 atualizar_contato não aceita mais codigo_cliente_destino (contato não é transferido)' (-not $schemaCt.Contains('codigo_cliente_destino'))

Write-Host '--- DEV-12: as regras do lote valem no conector (mesmos casos C30-C31 do front)'
Conferir 'C30 CNPJ com DV errado é recusado' (-not (Test-CnpjValido '11222333000182'))
Conferir 'C31 nota fracionária é recusada' (Recusa { ConvertTo-Inteiro 'aderencia' 1.5 })
Conferir 'C31 UF fora da lista é recusada' (-not (Test-UFValida 'XX'))

Write-Host '--- DEV-16: atendimento anterior sem ciclo (mesmo caso C33 do front)'
Conferir 'C33 anterior que fecha ciclo é recusado (A -> B -> A)' (Test-AnteriorFormaCiclo 10 20 @{ 20 = 10; 10 = 0 })
Conferir 'C33 corrente sem ciclo passa' (-not (Test-AnteriorFormaCiclo 10 20 @{ 20 = 30; 30 = 0 }))

Write-Host '--- seção 5.2: próxima ação só hoje ou futura (mesmo caso C34 do front)'
Conferir 'C34 próxima ação nova no passado é recusada' ((Test-ProxAcaoPermitida '2026-10-25' $null '2026-10-26') -match 'C34')
Conferir 'C34 próxima ação alterada para o passado é recusada' ((Test-ProxAcaoPermitida '2026-10-20' '2026-10-25' '2026-10-26') -match 'C34')
Conferir 'C34 próxima ação hoje passa' ($null -eq (Test-ProxAcaoPermitida '2026-10-26' $null '2026-10-26'))
Conferir 'C34 próxima ação futura passa' ($null -eq (Test-ProxAcaoPermitida '2026-11-03' '2026-10-01' '2026-10-26'))
Conferir 'C34 vencida já gravada que não mudou não trava' ($null -eq (Test-ProxAcaoPermitida '2026-10-01' '2026-10-01' '2026-10-26'))
Conferir 'C34 sem próxima ação passa' ($null -eq (Test-ProxAcaoPermitida $null '2026-10-01' '2026-10-26'))

Write-Host '--- DEV-03: mensagem controlada de escrita durante o backup (pura)'
Conferir 'backup em curso + erro de consulta atualizável vira a mensagem do backup' ((Get-MensagemEscritaRecusada 'A operação deve usar uma consulta atualizável.' $true) -match 'Cópia de segurança')
Conferir 'sem backup em curso o erro segue o caminho normal' ((Get-MensagemEscritaRecusada 'A operação deve usar uma consulta atualizável.' $false) -eq '')
Conferir 'erro de outro tipo com backup em curso não é traduzido' ((Get-MensagemEscritaRecusada 'Violação de chave' $true) -eq '')

Write-Host '--- DEV-03: aviso por janela (seção 5.3) - os mesmos casos de testar-recuperacao.ps1 e do modAutoteste'
function CJ([string]$quando, [string]$tipo, [bool]$val) { return [pscustomobject]@{ instante = [datetime]::ParseExact($quando, 'yyyy-MM-dd HH:mm', $null); tipo = $tipo; validado = $val } }
function AJ([string]$quando) { return [datetime]::ParseExact($quando, 'yyyy-MM-dd HH:mm', $null) }
$casosJ = @(
    @{ n = 'J1'; agora = '2026-10-19 10:00'; copias = @((CJ '2026-10-16 16:05' 'diario' $true), (CJ '2026-10-19 09:05' 'diario' $true)); motivos = 0 }
    @{ n = 'J2'; agora = '2026-10-19 10:00'; copias = @((CJ '2026-10-16 16:05' 'diario' $true)); motivos = 1 }
    @{ n = 'J3'; agora = '2026-10-19 09:20'; copias = @((CJ '2026-10-16 16:05' 'diario' $true)); motivos = 0 }
    @{ n = 'J4'; agora = '2026-10-19 17:00'; copias = @((CJ '2026-10-16 16:05' 'diario' $true), (CJ '2026-10-19 09:05' 'periodico' $true), (CJ '2026-10-19 12:35' 'periodico' $true), (CJ '2026-10-19 16:10' 'periodico' $true)); motivos = 1 }
    @{ n = 'J5'; agora = '2026-10-19 17:00'; copias = @((CJ '2026-10-19 16:10' 'diario' $true)); motivos = 0 }
    @{ n = 'J6'; agora = '2026-10-19 08:00'; copias = @((CJ '2026-10-15 16:05' 'diario' $true)); motivos = 2 }
    @{ n = 'J7'; agora = '2026-10-24 11:00'; copias = @((CJ '2026-10-23 16:05' 'diario' $true)); motivos = 0 }
    @{ n = 'J8'; agora = '2026-10-19 17:00'; copias = @((CJ '2026-10-19 16:10' 'diario' $false)); motivos = 1 }
)
foreach ($k in $casosJ) {
    $av = Get-AvisoBackupJanela (AJ $k.agora) $k.copias
    Conferir ('aviso de backup ' + $k.n + ': ' + $k.motivos + ' motivo(s)') (@($av.motivos).Count -eq $k.motivos -and $av.alerta -eq ($k.motivos -gt 0)) ($av | ConvertTo-Json -Compress)
}

Write-Host '--- regras de lista da Fase 2 (C03, C04, C05): puras, sem servidor'
Conferir 'C05 UF de uma das 27 passa' (Test-UFValida 'sp')
Conferir 'C05 UF fora das 27 é recusada' (-not (Test-UFValida 'XX'))
Conferir 'C05 UF com espaço fora é recusada' (-not (Test-UFValida ' SP '))
Conferir 'C04 resultado GANHO com encerra Não é recusado' ($null -ne (Test-PapelEtapaValido $false 'GANHO'))
Conferir 'C04 resultado PERDA com encerra Sim passa' ($null -eq (Test-PapelEtapaValido $true 'PERDA'))
Conferir 'C04 sem resultado e sem encerra passa' ($null -eq (Test-PapelEtapaValido $false ''))
Conferir 'C04 resultado fora de GANHO/PERDA é recusado' ($null -ne (Test-PapelEtapaValido $true 'TALVEZ'))
$papeis = @(
    [pscustomobject]@{ valor = 'Pedido Fechado'; resultado = 'GANHO'; ativo = $true }
    [pscustomobject]@{ valor = 'Perdido'; resultado = 'PERDA'; ativo = $true }
    [pscustomobject]@{ valor = 'Descartado'; resultado = 'PERDA'; ativo = $true }
    [pscustomobject]@{ valor = 'Negociação'; resultado = ''; ativo = $true }
)
Conferir 'C03 desativar a única GANHO é recusado' (-not (Test-DesativarEtapaPermitido $papeis 'Pedido Fechado'))
Conferir 'C03 desativar uma de duas PERDA passa' (Test-DesativarEtapaPermitido $papeis 'Perdido')
Conferir 'C03 desativar etapa sem resultado passa' (Test-DesativarEtapaPermitido $papeis 'Negociação')
$semDescartado = @($papeis | Where-Object { $_.valor -ne 'Descartado' })
Conferir 'C03 com a outra PERDA já inativa, a última PERDA é recusada' (-not (Test-DesativarEtapaPermitido $semDescartado 'Perdido'))

try {
Write-Host '--- gerenciar_lista (Fase 2): regras C03, C04, C05 e papel, no banco de demonstração'
function Mensagem([scriptblock]$bloco) { try { $null = & $bloco; return '' } catch { return [string]$_.Exception.Message } }
function Lista([string]$acao, [string]$tipo, [string]$valor, $extra = @{}) {
    $p = [ordered]@{ tipo = $tipo; valor = $valor; acao = $acao }
    foreach ($k in $extra.Keys) { $p[$k] = $extra[$k] }
    return (Invoke-GerenciarLista ([pscustomobject]$p))
}
Conferir 'C05 incluir UF inválida é recusado com a regra' ((Mensagem { Lista 'incluir' 'UF' 'XX' }) -match 'C05')
Conferir 'C04 definir_papel com resultado e encerra Não é recusado' ((Mensagem { Lista 'definir_papel' 'Etapa' 'Negociação' @{ encerra = $false; resultado = 'GANHO' } }) -match 'C04')
Conferir 'C03 desativar a única etapa GANHO é recusado' ((Mensagem { Lista 'desativar' 'Etapa' 'Pedido Fechado' }) -match 'C03')
Conferir 'C03 e C04: uma das duas PERDA pode ser desativada' ((Mensagem {
        try { Lista 'desativar' 'Etapa' 'Descartado' | Out-Null } finally { Lista 'ativar' 'Etapa' 'Descartado' | Out-Null } }) -eq '')
Conferir 'definir_papel que muda encerra com atendimentos pede confirmar' ((Mensagem { Lista 'definir_papel' 'Etapa' 'Negociação' @{ encerra = $true } }) -match 'confirmar')
$antes = Lista 'definir_papel' 'Etapa' 'Negociação' @{ quadro = 'FUNIL' }
Conferir 'definir_papel sem mudança devolve gravado false' ($antes.gravado -eq $false)
try {
    $muda = Lista 'definir_papel' 'Etapa' 'Negociação' @{ quadro = 'PROSPECCAO' }
    Conferir 'definir_papel muda quadro e devolve o papel novo' ($muda.gravado -eq $true -and $muda.papel.quadro -eq 'PROSPECCAO' -and @($muda.alterados) -contains 'quadro')
} finally {
    $volta = Lista 'definir_papel' 'Etapa' 'Negociação' @{ quadro = 'FUNIL' }
    Conferir 'definir_papel restaurou o quadro original (FUNIL)' ($volta.papel.quadro -eq 'FUNIL')
}
Conferir 'definir_papel em lista sem papel é recusado' ((Mensagem { Lista 'definir_papel' 'Responsavel' 'X' @{ quadro = 'FUNIL' } }) -match 'Etapa|não existe|Lista')
$det = Invoke-ListarOpcoes ([pscustomobject]@{ tipo = 'Etapa'; detalhar = $true })
$pf = @($det.itens | Where-Object { $_.valor -eq 'Pedido Fechado' }) | Select-Object -First 1
Conferir 'listar_opcoes detalhar traz o papel da etapa (Pedido Fechado = GANHO, encerra)' ($null -ne $pf -and $pf.resultado -eq 'GANHO' -and $pf.encerra -eq $true)
Conferir 'listar_opcoes detalhar traz o uso de cada valor' ($null -ne $pf -and $pf.uso -ge 0 -and $null -ne $pf.ativo)
Conferir 'listar_opcoes sem detalhar continua só com a lista de valores' (-not ((Invoke-ListarOpcoes ([pscustomobject]@{ tipo = 'Etapa' })).PSObject.Properties.Name -contains 'itens'))
# revisão final 003: o uso de cada valor sai de uma leitura por coluna, não de uma conexão por valor
$script:aberturas = 0
$abrirOriginal = ${function:Open-Conexao}
function Open-Conexao { $script:aberturas++; & $abrirOriginal }
try {
    $detR = Invoke-ListarOpcoes ([pscustomobject]@{ tipo = 'Responsavel'; detalhar = $true })
    $detU = Invoke-ListarOpcoes ([pscustomobject]@{ tipo = 'UF'; detalhar = $true })
} finally { ${function:Open-Conexao} = $abrirOriginal }
$nItens = @($detR.itens).Count + @($detU.itens).Count
Conferir ('listar_opcoes detalhar sem conexão por valor (' + $script:aberturas + ' conexões para ' + $nItens + ' valores)') ($script:aberturas -le 6 -and $nItens -gt 6) ''
$usoRef = [int](Invoke-Valor 'SELECT COUNT(*) FROM clientes WHERE responsavel=?' @((P 'texto' ([string]@($detR.itens)[0].valor)))) + [int](Invoke-Valor 'SELECT COUNT(*) FROM oportunidades WHERE responsavel=?' @((P 'texto' ([string]@($detR.itens)[0].valor))))
Conferir 'listar_opcoes detalhar: uso do responsável soma clientes e atendimentos' ([int]@($detR.itens)[0].uso -eq $usoRef) ([string]@($detR.itens)[0].uso + ' x ' + $usoRef)
# revisão final 003: valor incluído grava a chave (o papel é lido pela chave; a montagem recusa lista sem chave)
$maxLogChave = [int](Invoke-Valor 'SELECT MAX(id) FROM log_alteracoes' @())
try {
    Conferir 'C02 incluir variante com acento de valor existente (Pedído Fechado) é recusado' ((Mensagem { Lista 'incluir' 'Etapa' 'Pedído Fechado' }) -match 'já existe')
    $inc = Lista 'incluir' 'Etapa' 'Teste Revisão Final'
    Conferir 'incluir etapa grava' ($inc.gravado -eq $true)
    $chaveInc = [string](Invoke-Valor "SELECT chave FROM listas WHERE tipo='Etapa' AND valor=?" @((P 'texto' 'Teste Revisão Final')))
    Conferir 'incluir grava a chave da lista (TESTE REVISAO FINAL)' ($chaveInc -eq 'TESTE REVISAO FINAL') $chaveInc
    $pap = Lista 'definir_papel' 'Etapa' 'Teste Revisão Final' @{ encerra = $true; resultado = 'PERDA' }
    Limpar-PapelEtapas
    Conferir 'C07 etapa incluída e marcada encerra é lida como encerrada' ($pap.gravado -eq $true -and (Get-PapelEtapa 'Teste Revisão Final').encerra -eq $true) ($pap | ConvertTo-Json -Compress -Depth 4)
} finally {
    $cnL = Open-Conexao
    try {
        foreach ($sqlL in @("DELETE FROM listas WHERE tipo='Etapa' AND valor='Teste Revisão Final'", ('DELETE FROM log_alteracoes WHERE id > ' + $maxLogChave))) {
            $c = New-Comando $cnL $sqlL @() $null; [void]$c.ExecuteNonQuery(); $c.Dispose()
        }
    } finally { $cnL.Close(); $cnL.Dispose() }
    Limpar-PapelEtapas
}
Write-Host '--- regras de visita (C08 a C13): puras'
Conferir 'C08 visita em atendimento encerrado e recusada' ((Test-VisitaPodeCriar $true) -match 'C08')
Conferir 'C08 visita em atendimento aberto passa' ($null -eq (Test-VisitaPodeCriar $false))
Conferir 'C09 data fora do periodo e recusada' ((Test-DataNoPeriodo '2026-10-20' '2026-10-01' '2026-10-15') -match 'C09')
Conferir 'C09 data dentro do periodo passa' ($null -eq (Test-DataNoPeriodo '2026-10-05' '2026-10-01' '2026-10-15'))
Conferir 'C10 realizada sem relato e recusada' ((Test-RealizadaValida '' '2026-10-14') -match 'C10')
Conferir 'C10 realizada com relato e data passa' ($null -eq (Test-RealizadaValida 'Visitou' '2026-10-14'))
Conferir 'C11 cancelada sem motivo e recusada' ((Test-CanceladaValida '') -match 'C11')
Conferir 'C11 cancelada com motivo passa' ($null -eq (Test-CanceladaValida 'Cliente viajou'))
Conferir 'C12 texto do historico comeca com VISITA dd/MM/aaaa HH:mm' ((Get-TextoHistoricoVisita ([datetime]'2026-10-14 09:30') 'Visita' 'Positivo' 'Visitou') -eq 'VISITA 14/10/2026 09:30 - Visita - Positivo: Visitou')
Conferir 'C13 segunda pendente e recusada' ((Test-SegundaPendente 1) -match 'C13')
Conferir 'C13 sem pendente passa' ($null -eq (Test-SegundaPendente 0))
Write-Host '--- visitas (Fase 3.2a): criar, realizar e cancelar pelo conector, no banco de demonstração'
function Linhas-Visita([string]$sql, $params = @()) { $x = Invoke-Consulta $sql $params; while ($x -is [System.Collections.IList] -and $x.Count -eq 1 -and $x[0] -is [System.Collections.IList]) { $x = $x[0] }; return ,@($x) }
function Executa-Visita([string]$sql, $params = @()) {
    $cn = Open-Conexao
    try { $c = New-Comando $cn $sql $params $null; [void]$c.ExecuteNonQuery(); $c.Dispose() } finally { $cn.Close(); $cn.Dispose() }
}
$filaNeg = Linhas-Visita "SELECT codigo FROM oportunidades WHERE etapa='Negociação' ORDER BY id"
$codNeg = [string]$filaNeg[0].codigo
# atendimento numa etapa que encerra (pelo papel da lista): o banco de teste não tem 'Pedido Fechado'
# (só 'Perdido'), o de demonstração tem os dois; com o código vazio a recusa vinha do formato, não da C08
$filaFech = Linhas-Visita "SELECT o.codigo FROM oportunidades o INNER JOIN listas l ON o.etapa = l.valor WHERE l.tipo='Etapa' AND l.encerra=True ORDER BY o.id"
$codFech = [string]$filaFech[0].codigo
$opAntes = (Linhas-Visita "SELECT id, observacoes, ultima_interacao, versao FROM oportunidades WHERE codigo='$codNeg'")[0]
$maxLogV = [int](Invoke-Valor 'SELECT MAX(id) FROM log_alteracoes' @())
$criadas = @()
try {
    $c1 = Invoke-CriarVisita ([pscustomobject]@{ codigo = $codNeg; tipo = 'Visita'; responsavel = 'Vendedor A'; observacoes = 'teste visita' })
    $criadas += $c1.id
    Conferir 'criar_visita grava PENDENTE na versão 1' ($c1.gravado -and $c1.situacao -eq 'PENDENTE' -and $c1.versao -eq 1) ($c1 | ConvertTo-Json -Compress)
    Conferir 'C13: segunda visita pendente no mesmo atendimento e recusada' ((Mensagem { Invoke-CriarVisita ([pscustomobject]@{ codigo = $codNeg }) }) -match 'C13')
    Conferir ('C08: visita em atendimento encerrado (' + $codFech + ') e recusada') ($codFech -ne '' -and (Mensagem { Invoke-CriarVisita ([pscustomobject]@{ codigo = $codFech }) }) -match 'C08')
    Conferir 'V4/C10: realizada sem relato e recusada' ((Mensagem { Invoke-AtualizarVisita ([pscustomobject]@{ id = $c1.id; versao = 1; situacao = 'REALIZADA'; data_realizada = '2026-10-12' }) }) -match 'C10')
    Conferir 'conflito: versão lida antiga e recusada' ((Mensagem { Invoke-AtualizarVisita ([pscustomobject]@{ id = $c1.id; versao = 9; situacao = 'CANCELADA'; motivo_cancelamento = 'x' }) }) -match 'Conflito')
    $r1 = Invoke-AtualizarVisita ([pscustomobject]@{ id = $c1.id; versao = 1; situacao = 'REALIZADA'; data_realizada = '2026-10-12 09:30'; resultado = 'Positivo'; relato = 'Visitou e apresentou a proposta' })
    Conferir 'realizar a visita grava e sobe a versão para 2' ($r1.gravado -and $r1.versao -eq 2) ($r1 | ConvertTo-Json -Compress)
    $opDepois = (Linhas-Visita "SELECT observacoes, ultima_interacao FROM oportunidades WHERE codigo='$codNeg'")[0]
    Conferir 'V6: o histórico do atendimento ganhou a linha VISITA 12/10/2026 09:30' ([string]$opDepois.observacoes -match 'VISITA 12/10/2026 09:30') ([string]$opDepois.observacoes)
    Conferir 'V7: a última interação do atendimento é a data realizada' (([datetime]$opDepois.ultima_interacao).ToString('yyyy-MM-dd') -eq '2026-10-12') ([string]$opDepois.ultima_interacao)
    $c2 = Invoke-CriarVisita ([pscustomobject]@{ codigo = $codNeg; tipo = 'Retorno' })
    $criadas += $c2.id
    Conferir 'depois de realizada, cabe outra visita pendente (C13 só vale com pendente)' ($c2.gravado -eq $true)
    Conferir 'C11: cancelar sem motivo e recusado' ((Mensagem { Invoke-AtualizarVisita ([pscustomobject]@{ id = $c2.id; versao = 1; situacao = 'CANCELADA' }) }) -match 'C11')
    $r2 = Invoke-AtualizarVisita ([pscustomobject]@{ id = $c2.id; versao = 1; situacao = 'CANCELADA'; motivo_cancelamento = 'Cliente viajou' })
    Conferir 'cancelar com motivo grava' ($r2.gravado -eq $true -and $r2.situacao -eq 'CANCELADA')
} finally {
    # volta: apaga as visitas do teste, o histórico e a última interação do atendimento, e o log
    foreach ($i in $criadas) { Executa-Visita 'DELETE FROM visitas WHERE id=?' @((P 'inteiro' $i)) }
    Executa-Visita 'UPDATE oportunidades SET observacoes=?, ultima_interacao=?, versao=? WHERE id=?' @((P 'memo' $opAntes.observacoes), (P 'data' $opAntes.ultima_interacao), (P 'inteiro' ([int]$opAntes.versao)), (P 'inteiro' ([int]$opAntes.id)))
    Executa-Visita 'DELETE FROM log_alteracoes WHERE id > ?' @((P 'inteiro' $maxLogV))
}
Write-Host '--- ferramentas de visita (Fase 3.4): listar, alterar situação, resumo no atendimento'
$filaT7 = Linhas-Visita "SELECT codigo FROM oportunidades WHERE etapa='Negociação' ORDER BY id"
$codT7 = [string]$filaT7[0].codigo
$maxLogT7 = [int](Invoke-Valor 'SELECT MAX(id) FROM log_alteracoes' @())
$criadasT7 = @()
try {
    $v1 = Invoke-CriarVisita ([pscustomobject]@{ codigo = $codT7; tipo = 'Visita'; observacoes = 'teste 3.4' })
    $criadasT7 += $v1.id
    $lv = Invoke-ListarVisitas ([pscustomobject]@{ codigo = $codT7 })
    Conferir 'listar_visitas por atendimento traz a visita criada' (@($lv.visitas | Where-Object { [int]$_.id -eq [int]$v1.id }).Count -eq 1) ([string]$lv.quantidade)
    Conferir 'listar_visitas traz o contato (D-a: nome e telefone)' (@($lv.visitas | Where-Object { [int]$_.id -eq [int]$v1.id })[0].contato -ne $null) ''
    $lf = Invoke-ListarVisitas ([pscustomobject]@{ situacao = 'REALIZADA'; codigo = $codT7 })
    Conferir 'listar_visitas filtrado por situacao REALIZADA não traz a pendente' (@($lf.visitas | Where-Object { [int]$_.id -eq [int]$v1.id }).Count -eq 0) ''
    Conferir 'alterar_situacao_visita sem motivo ao cancelar é recusado (C11)' ((Mensagem { Invoke-AlterarSituacaoVisita ([pscustomobject]@{ id = $v1.id; versao = 1; situacao = 'CANCELADA' }) }) -match 'C11')
    $ac = Invoke-AlterarSituacaoVisita ([pscustomobject]@{ id = $v1.id; versao = 1; situacao = 'CANCELADA'; motivo_cancelamento = 'Cliente adiou' })
    Conferir 'alterar_situacao_visita cancela com motivo' ($ac.gravado -eq $true -and $ac.situacao -eq 'CANCELADA') ($ac | ConvertTo-Json -Compress)
    $at = Invoke-ObterAtendimento ([pscustomobject]@{ codigo = $codT7 })
    Conferir 'obter_atendimento traz o total de visitas e a lista' ([int]$at.visitas_total -ge 1 -and @($at.visitas).Count -ge 1) ([string]$at.visitas_total)
} finally {
    foreach ($i in $criadasT7) { Executa-Visita 'DELETE FROM visitas WHERE id=?' @((P 'inteiro' $i)) }
    Executa-Visita 'DELETE FROM log_alteracoes WHERE id > ?' @((P 'inteiro' $maxLogT7))
}
Write-Host '--- planejamento (Fase 4.1 e 4.2): nome_busca, distância e regras P1 a P3, puras'
Conferir 'nome_busca: sem acento e em maiúsculas (São Luís = SAO LUIS)' ((ConvertTo-NomeBusca 'São Luís') -eq 'SAO LUIS')
Conferir 'distância: mesma cidade é 0 km' ((Get-DistanciaKm -23.5505 -46.6333 -23.5505 -46.6333) -eq 0)
$dSpRj = Get-DistanciaKm -23.5505 -46.6333 -22.9068 -43.1729
Conferir 'distância: São Paulo a Rio fica entre 340 e 370 km (Haversine, raio 6.371 km)' ($dSpRj -gt 340 -and $dSpRj -lt 370) ([string]$dSpRj)
Conferir 'C14 P1: raio 0 é recusado' ((Test-RaioValido 0) -match 'C14')
Conferir 'C14 P1: raio 501 é recusado' ((Test-RaioValido 501) -match 'C14')
Conferir 'C14 P1: raio 1 e 500 passam' ($null -eq (Test-RaioValido 1) -and $null -eq (Test-RaioValido 500))
Conferir 'C15 P2: início depois do fim é recusado' ((Test-PeriodoValido '2026-10-20' '2026-10-10') -match 'C15')
Conferir 'C15 P2: período de um dia passa' ($null -eq (Test-PeriodoValido '2026-10-10' '2026-10-10'))
Conferir 'C16 P3: encerrar com pendentes sem cancelar é recusado' ((Test-EncerramentoPermitido 2 $false) -match 'C16')
Conferir 'C16 P3: encerrar com pendentes cancelando passa' ($null -eq (Test-EncerramentoPermitido 2 $true))
Conferir 'C16 P3: encerrar sem pendentes passa' ($null -eq (Test-EncerramentoPermitido 0 $false))
Conferir 'P4: código PL-AAAA-NNN com 3 dígitos' ((Get-CodigoPlanejamento 2026 7) -eq 'PL-2026-007')
Write-Host '--- planejamento (Fase 4.3): criar, validar, encerrar com pendentes (C16) e reabrir, no banco de demonstração'
$cidT9 = [int](Invoke-Valor 'SELECT MIN(cod_ibge) FROM municipios' @())
$maxLogT9 = [int](Invoke-Valor 'SELECT MAX(id) FROM log_alteracoes' @())
$filaT9 = Linhas-Visita "SELECT codigo FROM oportunidades WHERE etapa='Negociação' ORDER BY id"
$codT9 = [string]$filaT9[0].codigo
$planT9 = @()
$visT9 = @()
try {
    $p1 = Invoke-CriarPlanejamento ([pscustomobject]@{ descricao = 'TESTE PLANEJAMENTO'; cod_municipio = $cidT9; raio_km = 50; data_inicio = '2026-10-20'; data_fim = '2026-10-30' })
    $planT9 += $p1.id
    Conferir 'criar_planejamento grava ABERTO na versão 1 com código PL-AAAA-NNN' ($p1.gravado -and $p1.situacao -eq 'ABERTO' -and $p1.codigo -match '^PL-\d{4}-\d{3}$') ($p1 | ConvertTo-Json -Compress)
    Conferir 'C14: raio 0 é recusado ao criar' ((Mensagem { Invoke-CriarPlanejamento ([pscustomobject]@{ descricao = 'x'; cod_municipio = $cidT9; raio_km = 0; data_inicio = '2026-10-20'; data_fim = '2026-10-30' }) }) -match 'C14')
    Conferir 'C15: início depois do fim é recusado' ((Mensagem { Invoke-CriarPlanejamento ([pscustomobject]@{ descricao = 'x'; cod_municipio = $cidT9; raio_km = 10; data_inicio = '2026-10-30'; data_fim = '2026-10-20' }) }) -match 'C15')
    Conferir 'cidade-base inexistente é recusada' ((Mensagem { Invoke-CriarPlanejamento ([pscustomobject]@{ descricao = 'x'; cod_municipio = 1; raio_km = 10; data_inicio = '2026-10-20'; data_fim = '2026-10-30' }) }) -match 'Cidade-base')
    $v1 = Invoke-CriarVisita ([pscustomobject]@{ codigo = $codT9; tipo = 'Visita'; data_agendada = '2026-10-25'; id_planejamento = $p1.id })
    $visT9 += $v1.id
    Conferir 'C09: visita com data dentro do período do planejamento passa' ($v1.gravado -eq $true)
    Conferir 'C16: encerrar com visita pendente sem cancelar é recusado' ((Mensagem { Invoke-AlterarSituacaoPlanejamento ([pscustomobject]@{ id = $p1.id; versao = 1; situacao = 'ENCERRADO' }) }) -match 'C16')
    $e1 = Invoke-AlterarSituacaoPlanejamento ([pscustomobject]@{ id = $p1.id; versao = 1; situacao = 'ENCERRADO'; cancelar_pendentes = $true })
    Conferir 'encerrar cancelando pendentes grava e cancela a visita' ($e1.gravado -and $e1.situacao -eq 'ENCERRADO' -and [int]$e1.visitas_canceladas -eq 1) ($e1 | ConvertTo-Json -Compress)
    $vDepois = Linhas-Visita "SELECT situacao, motivo_cancelamento FROM visitas WHERE id=$($v1.id)"
    Conferir 'a visita cancelada ficou CANCELADA com o motivo do planejamento' ([string]$vDepois[0].situacao -eq 'CANCELADA' -and [string]$vDepois[0].motivo_cancelamento -eq 'planejamento encerrado') ''
    Conferir 'planejamento encerrado não se altera (atualizar_planejamento)' ((Mensagem { Invoke-AtualizarPlanejamento ([pscustomobject]@{ id = $p1.id; versao = 2; raio_km = 60 }) }) -match 'Planejamento encerrado')
    $r1 = Invoke-AlterarSituacaoPlanejamento ([pscustomobject]@{ id = $p1.id; versao = 2; situacao = 'ABERTO' })
    Conferir 'reabrir grava ABERTO e sobe a versão' ($r1.gravado -and $r1.situacao -eq 'ABERTO' -and $r1.versao -eq 3) ($r1 | ConvertTo-Json -Compress)
    $u1 = Invoke-AtualizarPlanejamento ([pscustomobject]@{ id = $p1.id; versao = 3; raio_km = 60 })
    Conferir 'atualizar_planejamento grava o raio novo' ($u1.gravado -and $u1.versao -eq 4) ($u1 | ConvertTo-Json -Compress)
    Conferir 'C14: atualizar com raio 501 é recusado' ((Mensagem { Invoke-AtualizarPlanejamento ([pscustomobject]@{ id = $p1.id; versao = 4; raio_km = 501 }) }) -match 'C14')
} finally {
    foreach ($i in $visT9) { Executa-Visita 'DELETE FROM visitas WHERE id=?' @((P 'inteiro' $i)) }
    foreach ($i in $planT9) { Executa-Visita 'DELETE FROM planejamentos WHERE id=?' @((P 'inteiro' $i)) }
    Executa-Visita 'DELETE FROM log_alteracoes WHERE id > ?' @((P 'inteiro' $maxLogT9))
}
Write-Host '--- candidatos e geração de visitas (Fase 4.4 e 4.9b): no banco de demonstração'
$cidCwb = [int](Invoke-Valor "SELECT MIN(cod_ibge) FROM municipios WHERE nome='CURITIBA' AND uf='PR'" @())
$maxLogT10 = [int](Invoke-Valor 'SELECT MAX(id) FROM log_alteracoes' @())
$planT10 = @(); $visT10 = @()
try {
    $pc = Invoke-CriarPlanejamento ([pscustomobject]@{ descricao = 'TESTE CANDIDATOS'; cod_municipio = $cidCwb; raio_km = 500; data_inicio = '2026-10-20'; data_fim = '2026-10-30' })
    $planT10 += $pc.id
    $lc = Invoke-ListarCandidatos ([pscustomobject]@{ id_planejamento = $pc.id })
    Conferir 'candidatos: os de Curitiba estão a 0 km' (@($lc.candidatos | Where-Object { $_.cidade -eq 'CURITIBA' } | Where-Object { $_.km -eq 0 }).Count -ge 1) ([string]$lc.quantidade)
    Conferir 'candidatos: fora do raio não entram (todos com km <= raio)' (@($lc.candidatos | Where-Object { $_.km -gt 500 }).Count -eq 0) ''
    $ordQ = @($lc.candidatos | ForEach-Object { [int]$_.ordem_qualificacao })
    $ordemOk = $true; for ($i = 1; $i -lt $ordQ.Count; $i++) { if ($ordQ[$i] -lt $ordQ[$i - 1]) { $ordemOk = $false } }
    Conferir 'candidatos: ordem de visita (qualificação não decresce)' ($ordQ.Count -ge 2 -and $ordemOk) ($ordQ -join ',')
    $alvo = [string]($lc.candidatos | Where-Object { -not $_.visita_pendente } | Select-Object -First 1).codigo
    $sim = Invoke-GerarVisitas ([pscustomobject]@{ id_planejamento = $pc.id; codigos = @($alvo); simular = $true })
    Conferir 'gerar_visitas simular confere sem gravar' ($sim.simulado -eq $true -and $sim.ok -eq 1 -and $sim.visitas_criadas.Count -eq 0) ($sim | ConvertTo-Json -Compress -Depth 4)
    $cntAntes = [int](Invoke-Valor 'SELECT COUNT(*) FROM visitas' @())
    $ger = Invoke-GerarVisitas ([pscustomobject]@{ id_planejamento = $pc.id; codigos = @($alvo); simular = $false })
    $visT10 += @($ger.visitas_criadas)
    Conferir 'gerar_visitas grava uma visita PENDENTE no planejamento' ($ger.visitas_criadas.Count -eq 1) ($ger | ConvertTo-Json -Compress -Depth 4)
    $encerrado = $codFech
    Conferir 'gerar_visitas: atendimento encerrado é recusado e nada é gravado (C08)' ($encerrado -ne '' -and (Mensagem { Invoke-GerarVisitas ([pscustomobject]@{ id_planejamento = $pc.id; codigos = @($alvo, $encerrado); simular = $false }) }) -match 'Nada gravado') ''
    Conferir 'gerar_visitas: a recusa desfez tudo (mesma quantidade de visitas)' ([int](Invoke-Valor 'SELECT COUNT(*) FROM visitas' @()) -eq ($cntAntes + 1)) ''
    Conferir 'gerar_visitas: segunda pendente no mesmo atendimento é recusada (C13)' ((Mensagem { Invoke-GerarVisitas ([pscustomobject]@{ id_planejamento = $pc.id; codigos = @($alvo); simular = $false }) }) -match 'Nada gravado') ''
    # DEV-27 (ciclo 016): a rota lista também as empresas ATIVAS do raio sem atendimento em aberto (pré-clientes inclusive)
    Executa-Visita "INSERT INTO clientes (estagio, empresa, cidade, uf, ativo, versao) VALUES ('Pré-cliente', 'TESTE ROTA PRE', 'CURITIBA', 'PR', True, 1)"
    Executa-Visita "INSERT INTO clientes (estagio, empresa, cidade, uf, ativo, versao) VALUES ('Pré-cliente', 'TESTE ROTA INATIVA', 'CURITIBA', 'PR', False, 1)"
    $lc2 = Invoke-ListarCandidatos ([pscustomobject]@{ id_planejamento = $pc.id })
    $sem2 = @($lc2.sem_atendimento)
    $preRota = @($sem2 | Where-Object { $_.empresa -eq 'TESTE ROTA PRE' })
    Conferir 'DEV-27 sem_atendimento traz o pré-cliente ativo do raio (km 0 em Curitiba)' ($preRota.Count -eq 1 -and [double]$preRota[0].km -eq 0 -and [int]$preRota[0].id_cliente -gt 0) ($sem2 | ConvertTo-Json -Compress -Depth 3)
    Conferir 'DEV-27 sem_atendimento não traz empresa inativa' (@($sem2 | Where-Object { $_.empresa -eq 'TESTE ROTA INATIVA' }).Count -eq 0)
    $empCand = @($lc2.candidatos | ForEach-Object { [string]$_.empresa })
    Conferir 'DEV-27 sem_atendimento não repete empresa que já é candidata (tem atendimento em aberto)' (@($sem2 | Where-Object { $empCand -contains [string]$_.empresa }).Count -eq 0) ''
    Conferir 'DEV-27 contadores do sem_atendimento coerentes (até 200)' ([int]$lc2.quantidade_sem_atendimento -ge $sem2.Count -and $sem2.Count -le 200 -and [bool]$lc2.sem_atendimento_truncado -eq ([int]$lc2.quantidade_sem_atendimento -gt 200)) ''
    Conferir 'DEV-27 a lista de candidatos não muda' ([int]$lc2.quantidade -eq [int]$lc.quantidade) ([string]$lc2.quantidade + ' x ' + [string]$lc.quantidade)
    Conferir 'DEV-27 gerar_visitas só aceita atendimento (código de cliente recusado, nada gravado)' ((Mensagem { Invoke-GerarVisitas ([pscustomobject]@{ id_planejamento = $pc.id; codigos = @('CL-0001'); simular = $false }) }) -match 'Nada gravado') ''
} finally {
    foreach ($i in $visT10) { Executa-Visita 'DELETE FROM visitas WHERE id=?' @((P 'inteiro' ([int]$i))) }
    foreach ($i in $planT10) { Executa-Visita 'DELETE FROM planejamentos WHERE id=?' @((P 'inteiro' $i)) }
    Executa-Visita 'DELETE FROM log_alteracoes WHERE id > ?' @((P 'inteiro' $maxLogT10))
    Executa-Visita "DELETE FROM clientes WHERE empresa IN ('TESTE ROTA PRE', 'TESTE ROTA INATIVA')"
}
Write-Host '--- lançamento em lote e leitura de planejamento (Fase 4.5, 4.9a, 4.9b): no banco de demonstração'
$cidT11 = [int](Invoke-Valor "SELECT MIN(cod_ibge) FROM municipios WHERE nome='CURITIBA' AND uf='PR'" @())
$maxLogT11 = [int](Invoke-Valor 'SELECT MAX(id) FROM log_alteracoes' @())
$filaT11 = Linhas-Visita "SELECT codigo FROM oportunidades WHERE etapa='Negociação' ORDER BY id"
$codT11 = [string]$filaT11[0].codigo
$planT11 = @(); $visT11 = @()
try {
    $pl11 = Invoke-CriarPlanejamento ([pscustomobject]@{ descricao = 'TESTE LOTE'; cod_municipio = $cidT11; raio_km = 100; data_inicio = '2026-10-20'; data_fim = '2026-10-30' })
    $planT11 += $pl11.id
    $v11 = Invoke-CriarVisita ([pscustomobject]@{ codigo = $codT11; tipo = 'Visita'; id_planejamento = $pl11.id })
    $visT11 += $v11.id
    $sim = Invoke-LancarDatas ([pscustomobject]@{ simular = $true; itens = @([pscustomobject]@{ id = $v11.id; versao = 1; data_agendada = '2026-10-25'; hora_agendada = '09:00' }) })
    Conferir 'lancar_datas simular confere sem gravar' ($sim.simulado -eq $true -and $sim.ok -eq 1) ($sim | ConvertTo-Json -Compress -Depth 4)
    $dataSim = Linhas-Visita "SELECT data_agendada FROM visitas WHERE id=$($v11.id)"
    Conferir 'lancar_datas simular não mudou a data' ($null -eq $dataSim[0].data_agendada) ''
    Conferir 'lancar_datas com data fora do período (C09) não grava nada' ((Mensagem { Invoke-LancarDatas ([pscustomobject]@{ simular = $false; itens = @([pscustomobject]@{ id = $v11.id; versao = 1; data_agendada = '2026-11-05' }) }) }) -match 'Nada gravado') ''
    $ger = Invoke-LancarDatas ([pscustomobject]@{ simular = $false; itens = @([pscustomobject]@{ id = $v11.id; versao = 1; data_agendada = '2026-10-25'; hora_agendada = '09:00' }) })
    Conferir 'lancar_datas grava a data e a hora' ($ger.recusadas -eq 0 -and $ger.ok -eq 1) ($ger | ConvertTo-Json -Compress -Depth 4)
    $dataDep = Linhas-Visita "SELECT data_agendada, hora_agendada, versao FROM visitas WHERE id=$($v11.id)"
    Conferir 'a visita ficou com a data 25/10/2026 09:00 e versão 2' (([datetime]$dataDep[0].data_agendada).ToString('yyyy-MM-dd') -eq '2026-10-25' -and [string]$dataDep[0].hora_agendada -eq '09:00' -and [int]$dataDep[0].versao -eq 2) ''
    Conferir 'lancar_resultados: cancelar sem motivo recusa o lote inteiro' ((Mensagem { Invoke-LancarResultados ([pscustomobject]@{ simular = $false; itens = @([pscustomobject]@{ id = $v11.id; versao = 2; situacao = 'CANCELADA' }) }) }) -match 'Nada gravado') ''
    $lp = Invoke-ListarPlanejamentos ([pscustomobject]@{ situacao = 'ABERTO' })
    Conferir 'listar_planejamentos traz o planejamento criado' (@($lp.planejamentos | Where-Object { [int]$_.id -eq [int]$pl11.id }).Count -eq 1) ''
    $op = Invoke-ObterPlanejamento ([pscustomobject]@{ codigo = $pl11.codigo })
    Conferir 'obter_planejamento traz o resumo de visitas (1 pendente)' ([int]$op.visitas.PENDENTE -eq 1 -and $op.cidade_base -eq 'Curitiba') ($op | ConvertTo-Json -Compress -Depth 3)
    $bm = Invoke-BuscarMunicipios ([pscustomobject]@{ texto = 'curitib'; uf = 'pr' })
    Conferir 'buscar_municipios acha Curitiba pelo nome sem diferença de caixa' (@($bm.municipios | Where-Object { [int]$_.cod_ibge -eq $cidT11 }).Count -eq 1) ''
    $rc = Invoke-ResolverCidade ([pscustomobject]@{ cidade = 'Curitiba'; uf = 'PR' })
    Conferir 'resolver_cidade resolve Curitiba/PR para um código' ($rc.resolvida -eq $true -and [int]$rc.municipio.cod_ibge -eq $cidT11) ''
    $rn = Invoke-ResolverCidade ([pscustomobject]@{ cidade = 'Nao existe'; uf = 'PR' })
    Conferir 'resolver_cidade sem resultado não resolve' ($rn.resolvida -eq $false) ''
} finally {
    foreach ($i in $visT11) { Executa-Visita 'DELETE FROM visitas WHERE id=?' @((P 'inteiro' $i)) }
    foreach ($i in $planT11) { Executa-Visita 'DELETE FROM planejamentos WHERE id=?' @((P 'inteiro' $i)) }
    Executa-Visita 'DELETE FROM log_alteracoes WHERE id > ?' @((P 'inteiro' $maxLogT11))
}
Write-Host '--- DEV-21: agenda do aberto pela próxima ação; encerrado com retomada vira "Retomada a avaliar"'
$hojeT = (Get-Date).Date
Conferir 'DEV-21 aberto com retomada vencida segue a próxima ação' ((Get-Situacao $true 'Negociação' $hojeT.AddDays(30) $hojeT.AddDays(-1)) -eq 'Em dia')
Conferir 'DEV-21 encerrado com retomada que chegou' ((Get-Situacao $true 'Perdido' $null $hojeT) -eq 'Retomada a avaliar')
Conferir 'DEV-21 encerrado com retomada futura continua Encerrado' ((Get-Situacao $true 'Perdido' $null $hojeT.AddDays(5)) -eq 'Encerrado')
# ciclo 018: só Perdido e Descartado (PERDA) viram "Retomada a avaliar", igual ao front e à seção retomadas_a_avaliar
Conferir 'DEV-21 Pedido Fechado com retomada que chegou continua Encerrado' ((Get-Situacao $true 'Pedido Fechado' $null $hojeT) -eq 'Encerrado')
Conferir 'DEV-21 Descartado com retomada que chegou é "Retomada a avaliar"' ((Get-Situacao $true 'Descartado' $null $hojeT) -eq 'Retomada a avaliar')
$fila21 = Linhas-Visita "SELECT TOP 2 id, codigo, etapa, responsavel, dt_prox_acao, retomar_em FROM oportunidades WHERE etapa IN (SELECT valor FROM listas WHERE tipo='Etapa' AND encerra=False) ORDER BY id"
$maxLog21 = [int](Invoke-Valor 'SELECT MAX(id) FROM log_alteracoes' @())
try {
    Executa-Visita 'UPDATE oportunidades SET etapa=?, responsavel=?, dt_prox_acao=NULL, retomar_em=? WHERE id=?' @((P 'texto' 'Perdido'), (P 'texto' 'TESTE RETOMADA 012'), (P 'data' $hojeT.AddDays(-3)), (P 'inteiro' ([int]$fila21[0].id)))
    Executa-Visita 'UPDATE oportunidades SET responsavel=?, dt_prox_acao=?, retomar_em=? WHERE id=?' @((P 'texto' 'TESTE RETOMADA 012'), (P 'data' $hojeT.AddDays(2)), (P 'data' $hojeT.AddDays(-1)), (P 'inteiro' ([int]$fila21[1].id)))
    $pl21 = (Invoke-AtendimentosPorProximaAcao ([pscustomobject]@{ responsavel = 'TESTE RETOMADA 012'; data_inicio = $hojeT.ToString('yyyy-MM-dd'); data_fim = $hojeT.AddDays(6).ToString('yyyy-MM-dd') })).__conteudo.text | ConvertFrom-Json
    $ret21 = @($pl21.retomadas_a_avaliar | ForEach-Object { [string]$_.codigo })
    $ate21 = @($pl21.atendimentos | ForEach-Object { [string]$_.codigo })
    Conferir 'DEV-21 encerrado com retomada aparece uma vez em retomadas_a_avaliar' ($ret21.Count -eq 1 -and $ret21[0] -eq [string]$fila21[0].codigo) ($pl21 | ConvertTo-Json -Compress -Depth 6)
    Conferir 'DEV-21 aberto aparece pela próxima ação e o encerrado não está na agenda' (($ate21 -contains [string]$fila21[1].codigo) -and -not ($ate21 -contains [string]$fila21[0].codigo)) ($ate21 -join ',')
    Conferir 'DEV-21 data futura nunca aparece como vencida' (@($pl21.atendimentos | Where-Object { $_.vencida -and [datetime]$_.dt_prox_acao -ge $hojeT }).Count -eq 0)
    $sit21 = @($pl21.retomadas_a_avaliar)[0].situacao
    Conferir 'DEV-21 a situação da retomada é "Retomada a avaliar"' ($sit21 -eq 'Retomada a avaliar') ([string]$sit21)
} finally {
    foreach ($o in $fila21) {
        Executa-Visita 'UPDATE oportunidades SET etapa=?, responsavel=?, dt_prox_acao=?, retomar_em=? WHERE id=?' @((P 'texto' $o.etapa), (P 'texto' $o.responsavel), (P 'data' (ConvertTo-DataOuNulo $o.dt_prox_acao)), (P 'data' (ConvertTo-DataOuNulo $o.retomar_em)), (P 'inteiro' ([int]$o.id)))
    }
    Executa-Visita 'DELETE FROM log_alteracoes WHERE id > ?' @((P 'inteiro' $maxLog21))
}

Write-Host '--- DEV-22: paginação do planejamento por cursor (inserção e exclusão entre páginas, filtro e geração)'
$limiteOriginal = $script:LimiteCaracteresPlanejamento
$arg22 = [ordered]@{ data_inicio = $hojeT.ToString('yyyy-MM-dd'); data_fim = $hojeT.AddDays(30).ToString('yyyy-MM-dd') }
function Arg22($extra) { $h = [ordered]@{}; foreach ($k in $arg22.Keys) { $h[$k] = $arg22[$k] }; foreach ($k in $extra.Keys) { $h[$k] = $extra[$k] }; return [pscustomobject]$h }
$todo22 = (Invoke-AtendimentosPorProximaAcao ([pscustomobject]$arg22)).__conteudo.text | ConvertFrom-Json
$cods22 = @($todo22.atendimentos | ForEach-Object { [string]$_.codigo })
Conferir 'DEV-22 resposta declara a leitura dinâmica' ($todo22.consistencia -eq 'dinamica') ([string]$todo22.consistencia)
$script:LimiteCaracteresPlanejamento = 2500
$sai22 = $null
try {
    $p1 = (Invoke-AtendimentosPorProximaAcao ([pscustomobject]$arg22)).__conteudo.text | ConvertFrom-Json
    $vistos22 = New-Object System.Collections.ArrayList
    foreach ($x in @($p1.atendimentos)) { [void]$vistos22.Add([string]$x.codigo) }
    Conferir 'DEV-22 a primeira página é parcial e traz proximo_cursor' ($p1.truncado -and $p1.proximo_cursor -and $vistos22.Count -lt $cods22.Count) ('pagina=' + $vistos22.Count + ' de ' + $cods22.Count)
    # exclusão entre páginas: um item já visto sai do conjunto (vira Perdido); com deslocamento, a página seguinte pularia um
    $sai22 = Linhas-Visita 'SELECT id, etapa, motivo_desfecho, dt_desfecho FROM oportunidades WHERE codigo=?' @((P 'texto' $vistos22[0]))
    Executa-Visita "UPDATE oportunidades SET etapa='Perdido' WHERE id=?" @((P 'inteiro' ([int]$sai22[0].id)))
    $cur = [string]$p1.proximo_cursor; $voltas = 0
    while ($cur -and $voltas -lt 50) {
        $pn = (Invoke-AtendimentosPorProximaAcao (Arg22 @{ cursor = $cur })).__conteudo.text | ConvertFrom-Json
        foreach ($x in @($pn.atendimentos)) { [void]$vistos22.Add([string]$x.codigo) }
        $cur = [string]$pn.proximo_cursor; $voltas++
    }
    $dup22 = @($vistos22 | Group-Object | Where-Object { $_.Count -gt 1 })
    $falta22 = @($cods22 | Where-Object { $vistos22 -notcontains $_ })
    Conferir 'DEV-22 exclusão entre páginas não duplica nem omite' ($dup22.Count -eq 0 -and $falta22.Count -eq 0 -and $vistos22.Count -eq $cods22.Count) ('duplicados=' + ($dup22.Name -join ',') + ' faltam=' + ($falta22 -join ','))
    $msg22 = Mensagem { Invoke-AtendimentosPorProximaAcao (Arg22 @{ cursor = [string]$p1.proximo_cursor; responsavel = 'Vendedor A' }) }
    Conferir 'DEV-22 cursor com filtro diferente é recusado' ($msg22 -match 'outra consulta') $msg22
    $g22 = [string](Invoke-Valor "SELECT valor FROM config WHERE chave='geracao_base'" @())
    Executa-Visita "UPDATE config SET valor='TESTE012GERACAO' WHERE chave='geracao_base'"
    try { $msg22 = Mensagem { Invoke-AtendimentosPorProximaAcao (Arg22 @{ cursor = [string]$p1.proximo_cursor }) } }
    finally { Executa-Visita "UPDATE config SET valor=? WHERE chave='geracao_base'" @((P 'texto' $g22)) }
    Conferir 'DEV-22 cursor de outra geração da base é recusado' ($msg22 -match 'trocada') $msg22
    Conferir 'DEV-22 cursor ilegível é recusado' ((Mensagem { Invoke-AtendimentosPorProximaAcao (Arg22 @{ cursor = 'xx' }) }) -match 'cursor')
} finally {
    $script:LimiteCaracteresPlanejamento = $limiteOriginal
    if ($sai22) { Executa-Visita 'UPDATE oportunidades SET etapa=? WHERE id=?' @((P 'texto' ([string]$sai22[0].etapa)), (P 'inteiro' ([int]$sai22[0].id))) }
}

Write-Host '--- DEV-22/23: extrato na raiz fixa, manifesto com hashes e consistência; descarte só do vencido com manifesto'
$raiz23 = Get-RaizExtratos
$antes23 = @(if (Test-Path -LiteralPath $raiz23) { Get-ChildItem -LiteralPath $raiz23 -Directory | ForEach-Object { $_.FullName } })
$criadas23 = New-Object System.Collections.ArrayList
try {
    Conferir 'DEV-23 destino informado é recusado (saiu do contrato)' ((Mensagem { Invoke-ExportarBase ([pscustomobject]@{ destino = 'QUALQUER' }) }) -match 'destino')
    $e23 = Invoke-ExportarBase ([pscustomobject]@{ tabelas = @('clientes', 'listas'); finalidade = 'teste de contrato'; autor = 'teste' })
    [void]$criadas23.Add([string]$e23.pasta)
    Conferir 'DEV-23 extrato gravado numa subpasta da raiz execucao\extratos' ((Split-Path -Parent ([string]$e23.pasta)) -eq $raiz23 -and (Test-Path -LiteralPath (Join-Path $e23.pasta 'crm-clientes.csv'))) ([string]$e23.pasta)
    $m23 = Get-Content -LiteralPath (Join-Path $e23.pasta 'crm-extrato.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    Conferir 'DEV-23 manifesto com solicitante, estação, finalidade, instante e corte' ($m23.tipo -eq 'extrato-crm' -and $m23.solicitante -eq 'teste' -and $m23.estacao -eq [Environment]::MachineName -and $m23.finalidade -eq 'teste de contrato' -and $m23.criado_em -and $null -ne $m23.corte.max_log_antes -and $m23.em_processamento -eq $false) ($m23 | ConvertTo-Json -Compress -Depth 5)
    $hashOk = ((Get-FileHash -LiteralPath (Join-Path $e23.pasta 'crm-clientes.csv') -Algorithm SHA256).Hash.ToLowerInvariant() -eq [string]$m23.tabelas.clientes.sha256)
    Conferir 'DEV-23 hash do arquivo confere com o manifesto' $hashOk
    Conferir 'DEV-22 extrato declara consistência não comprovada' ($m23.consistencia -eq 'nao_comprovada' -and $e23.consistencia -eq 'nao_comprovada') ([string]$m23.consistencia)
    # descarte: uma pasta vencida com manifesto, uma vencida SEM manifesto, uma em processamento e o extrato novo
    $velha = Join-Path $raiz23 '20200101-000000-aaaa'; New-Item -ItemType Directory -Path $velha -Force | Out-Null; [void]$criadas23.Add($velha)
    $mv = [ordered]@{ tipo = 'extrato-crm'; versao_manifesto = 1; criado_em = '2020-01-01T00:00:00'; em_processamento = $false }
    [System.IO.File]::WriteAllText((Join-Path $velha 'crm-extrato.json'), (ConvertTo-Json $mv), (New-Object System.Text.UTF8Encoding($false)))
    $semMan = Join-Path $raiz23 '20200101-000000-bbbb'; New-Item -ItemType Directory -Path $semMan -Force | Out-Null; [void]$criadas23.Add($semMan)
    $proc23 = Join-Path $raiz23 '20200101-000000-cccc'; New-Item -ItemType Directory -Path $proc23 -Force | Out-Null; [void]$criadas23.Add($proc23)
    $mv.em_processamento = $true
    [System.IO.File]::WriteAllText((Join-Path $proc23 'crm-extrato.json'), (ConvertTo-Json $mv), (New-Object System.Text.UTF8Encoding($false)))
    $s23 = Invoke-DescartarExtratos ([pscustomobject]@{ autor = 'teste' })
    Conferir 'DEV-23 descarte simula por padrão e não apaga' ($s23.simulado -eq $true -and (Test-Path -LiteralPath $velha) -and @($s23.descartar | Where-Object { $_ -like '*aaaa' }).Count -eq 1) ($s23 | ConvertTo-Json -Compress -Depth 4)
    $d23 = Invoke-DescartarExtratos ([pscustomobject]@{ simular = $false; autor = 'teste' })
    Conferir 'DEV-23 descarte apaga só o vencido com manifesto válido' (-not (Test-Path -LiteralPath $velha) -and (Test-Path -LiteralPath $semMan) -and (Test-Path -LiteralPath $proc23) -and (Test-Path -LiteralPath $e23.pasta)) ($d23 | ConvertTo-Json -Compress -Depth 4)
    Conferir 'DEV-23 descarte registrado em descartes.log' ((Get-Content -LiteralPath (Join-Path $raiz23 'descartes.log') -Raw) -match '20200101-000000-aaaa')
    # revisão final: junção dentro de uma pasta vencida não pode levar o descarte para fora da raiz
    $alvoJ = Join-Path $env:TEMP ('crm-fora-da-raiz-' + $PID); New-Item -ItemType Directory -Path $alvoJ -Force | Out-Null
    Set-Content -LiteralPath (Join-Path $alvoJ 'nao-apagar.txt') -Value 'fora da raiz'
    $velhaJ = Join-Path $raiz23 '20200101-000000-dddd'; New-Item -ItemType Directory -Path $velhaJ -Force | Out-Null; [void]$criadas23.Add($velhaJ)
    $mvJ = [ordered]@{ tipo = 'extrato-crm'; versao_manifesto = 1; criado_em = '2020-01-01T00:00:00'; em_processamento = $false }
    [System.IO.File]::WriteAllText((Join-Path $velhaJ 'crm-extrato.json'), (ConvertTo-Json $mvJ), (New-Object System.Text.UTF8Encoding($false)))
    [void](cmd /c mklink /J (Join-Path $velhaJ 'juncao') $alvoJ)
    try {
        [void](Invoke-DescartarExtratos ([pscustomobject]@{ simular = $false; autor = 'teste' }))
        Conferir 'revisão: o descarte apaga a pasta vencida sem seguir a junção (o arquivo de fora continua)' (-not (Test-Path -LiteralPath $velhaJ) -and (Test-Path -LiteralPath (Join-Path $alvoJ 'nao-apagar.txt'))) ''
    } finally { Remove-Item -LiteralPath $alvoJ -Recurse -Force -ErrorAction SilentlyContinue }
} finally {
    foreach ($c in $criadas23) { if (Test-Path -LiteralPath $c) { Remove-Item -LiteralPath $c -Recurse -Force } }
    $log23 = Join-Path $raiz23 'descartes.log'
    if ($antes23.Count -eq 0 -and (Test-Path -LiteralPath $log23)) { Remove-Item -LiteralPath $log23 -Force }
}
Write-Host '--- exportação do roteiro (Fase 5): CSV com BOM, ponto e vírgula, BASE primeiro, uma linha por visita pendente'
$cidT12 = [int](Invoke-Valor "SELECT MIN(cod_ibge) FROM municipios WHERE nome='CURITIBA' AND uf='PR'" @())
$maxLogT12 = [int](Invoke-Valor 'SELECT MAX(id) FROM log_alteracoes' @())
$filaT12 = Linhas-Visita "SELECT codigo FROM oportunidades WHERE etapa='Negociação' ORDER BY id"
$codT12 = [string]$filaT12[0].codigo
$pastaT12 = $null
$planT12 = @(); $visT12 = @()
try {
    $p12 = Invoke-CriarPlanejamento ([pscustomobject]@{ descricao = 'TESTE ROTEIRO'; cod_municipio = $cidT12; raio_km = 100; data_inicio = '2026-10-20'; data_fim = '2026-10-30' })
    $planT12 += $p12.id
    $v12 = Invoke-CriarVisita ([pscustomobject]@{ codigo = $codT12; tipo = 'Visita'; data_agendada = '2026-10-22'; hora_agendada = '09:30'; id_planejamento = $p12.id })
    $visT12 += $v12.id
    $ex = Invoke-ExportarRoteiro ([pscustomobject]@{ id_planejamento = $p12.id })
    $pastaT12 = Split-Path -Parent ([string]$ex.caminho)
    Conferir 'exportar_roteiro devolve o caminho e uma visita' ($ex.linhas -eq 1 -and (Test-Path -LiteralPath $ex.caminho)) ($ex | ConvertTo-Json -Compress)
    Conferir 'DEV-23 roteiro numa subpasta da raiz de extratos, com manifesto' ((Split-Path -Parent $pastaT12) -eq (Get-RaizExtratos) -and (Test-Path -LiteralPath (Join-Path $pastaT12 'crm-extrato.json'))) $pastaT12
    Conferir 'o nome do arquivo segue PL-AAAA-NNN_aaaa-mm-dd_a_aaaa-mm-dd.csv' ((Split-Path -Leaf $ex.caminho) -eq ($p12.codigo + '_2026-10-20_a_2026-10-30.csv')) (Split-Path -Leaf $ex.caminho)
    $bytes = [System.IO.File]::ReadAllBytes($ex.caminho)
    Conferir 'arquivo com BOM UTF-8' ($bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) ''
    $linhasT12 = @([System.IO.File]::ReadAllLines($ex.caminho, [System.Text.Encoding]::UTF8))
    Conferir 'cabeçalho separado por ponto e vírgula' ($linhasT12[0] -eq 'tipo;visita;atendimento;empresa;contato;telefone;cidade;uf;latitude;longitude;km_base;data_agendada;hora_agendada;tipo_visita;responsavel;situacao') $linhasT12[0]
    Conferir 'a segunda linha é a BASE, e a terceira é a visita' ($linhasT12.Count -eq 3 -and $linhasT12[1].StartsWith('"BASE"') -and $linhasT12[2].StartsWith('"VISITA"')) ($linhasT12 -join ' | ')
    Conferir 'a visita traz a data em AAAA-MM-DD e a hora' ($linhasT12[2] -match ';"2026-10-22";"09:30";') $linhasT12[2]
    Conferir 'latitude com ponto decimal (sem vírgula)' ($linhasT12[1] -match '-25\.\d+') $linhasT12[1]
} finally {
    if ($pastaT12 -and (Test-Path -LiteralPath $pastaT12)) { Remove-Item -LiteralPath $pastaT12 -Recurse -Force }
    foreach ($i in $visT12) { Executa-Visita 'DELETE FROM visitas WHERE id=?' @((P 'inteiro' $i)) }
    foreach ($i in $planT12) { Executa-Visita 'DELETE FROM planejamentos WHERE id=?' @((P 'inteiro' $i)) }
    Executa-Visita 'DELETE FROM log_alteracoes WHERE id > ?' @((P 'inteiro' $maxLogT12))
}
Write-Host '--- extrato da base (Fase 5.3): tabelas novas aceitas'
Conferir 'exportar_base aceita planejamentos, visitas e municipios (lista e consulta)' (@('planejamentos','visitas','municipios') | Where-Object { ($script:TabelasExtrato -contains $_) -and $script:SqlExtrato.ContainsKey($_) }).Count -eq 3 ''
    Write-Host '--- protocolo moderno (2026-07-28, sem estado)'
    $r = Enviar 'server/discover' (Moderno $null)
    Conferir 'server/discover responde DiscoverResult' ($r.result.resultType -eq 'complete' -and @($r.result.supportedVersions) -contains '2026-07-28') ($r | ConvertTo-Json -Compress -Depth 6)
    Conferir 'discover anuncia também as versões legadas' (@($r.result.supportedVersions) -contains '2025-11-25')
    Conferir 'discover traz serverInfo em _meta' ($r.result._meta.'io.modelcontextprotocol/serverInfo'.name -eq 'crm-zapromaq')
    $r = Enviar 'tools/list' (Moderno $null '1900-01-01')
    Conferir 'versão não suportada volta -32022 com a lista suportada' ($r.error.code -eq -32022 -and @($r.error.data.supported) -contains '2026-07-28') ($r | ConvertTo-Json -Compress -Depth 6)
    $r = Enviar 'tools/list' (Moderno $null)
    Conferir 'tools/list moderno: resultType, ttlMs e cacheScope' ($r.result.resultType -eq 'complete' -and $r.result.ttlMs -ge 0 -and $r.result.cacheScope -eq 'public')
    # DEV-07: contrato EXATO - ferramenta que some ou aparece sem constar aqui (e sem nota de mudança) quebra o teste
    $nomesModerno = @($r.result.tools | ForEach-Object { [string]$_.name })
    $faltam = @($script:FerramentasEsperadas | Where-Object { $nomesModerno -notcontains $_ })
    $sobram = @($nomesModerno | Where-Object { $script:FerramentasEsperadas -notcontains $_ })
    Conferir ('tools/list moderno: exatamente as ' + $script:FerramentasEsperadas.Count + ' ferramentas do contrato') ($faltam.Count -eq 0 -and $sobram.Count -eq 0 -and $nomesModerno.Count -eq $script:FerramentasEsperadas.Count) ('faltam: ' + ($faltam -join ',') + ' | sobram: ' + ($sobram -join ','))
    $incompletas = @($r.result.tools | Where-Object { -not $_.title -or -not $_.description -or $null -eq $_.inputSchema -or $_.inputSchema.type -ne 'object' -or $null -eq $_.annotations -or $null -eq $_.annotations.readOnlyHint })
    Conferir 'toda ferramenta tem título, descrição, inputSchema de objeto e annotations' ($incompletas.Count -eq 0) (($incompletas | ForEach-Object { $_.name }) -join ',')
    # DEV-07 (ciclo 013): matriz de consumidores (skills do comercial, fora de BASE) contra o catálogo publicado
    $mat = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'consumidores.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $porNome = @{}; foreach ($f in $r.result.tools) { $porNome[[string]$f.name] = $f }
    $quebras = @()
    foreach ($c in $mat.consumidores) {
        foreach ($pf in $c.ferramentas.PSObject.Properties) {
            if ($pf.Name -eq '*') { continue }
            if (-not $porNome.ContainsKey($pf.Name)) { $quebras += ($c.skill + ': ' + $pf.Name + ' não existe'); continue }
            $props = @($porNome[$pf.Name].inputSchema.properties.PSObject.Properties | ForEach-Object { $_.Name })
            foreach ($par in @($pf.Value)) { if ($props -notcontains $par) { $quebras += ($c.skill + ': ' + $pf.Name + '.' + $par + ' não existe') } }
        }
    }
    Conferir ('DEV-07 matriz de consumidores (' + @($mat.consumidores).Count + ' skills) bate com o catálogo') ($quebras.Count -eq 0) ($quebras -join '; ')
    $vServ = [regex]::Match([IO.File]::ReadAllText($servidor), "VersaoServidor = '([^']+)'").Groups[1].Value
    Conferir 'DEV-07 matriz revista na versão atual do conector' ([string]$mat.versao_conector -eq $vServ) ([string]$mat.versao_conector + ' x ' + $vServ)
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
    Conferir ('tools/list legado: as mesmas ' + $script:FerramentasEsperadas.Count + ' ferramentas do contrato') ((@($nomes | Sort-Object) -join ',') -eq (@($script:FerramentasEsperadas | Sort-Object) -join ',')) ($nomes -join ',')
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
        Conferir ('gravação liberada no esquema ' + $st.versao_esquema + ' (o conector exige ' + $script:EsquemasSuportados[0] + ')') ($st.versao_esquema -eq $script:EsquemasSuportados[0] -and $st.gravacao_liberada -eq $true -and $st.atualizacao_pendente -eq $false)
        # DEV-07: o consumidor (skill) descobre versão e capacidades pelo crm_status
        Conferir 'crm_status traz versao_conector, capacidades (o catálogo) e quantidade_ferramentas' ($st.versao_conector -match '^\d+\.\d+\.\d+$' -and [int]$st.quantidade_ferramentas -eq $script:FerramentasEsperadas.Count -and ((@($st.capacidades | Sort-Object) -join ',') -eq (@($script:FerramentasEsperadas | Sort-Object) -join ','))) ($res.content[0].text)

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
        $res = Chamar 'atualizar_atendimento' ([ordered]@{ codigo = 'AT-0001-0001'; versao = [int]$g.versao; etapa = 'Pedido Fechado' })
        Conferir 'Pedido Fechado sem motivo é recusado' ($res.isError -and $res.content[0].text -match 'motivo')
        $res = Chamar 'atualizar_atendimento' ([ordered]@{ codigo = 'AT-0001-0001'; versao = [int]$g.versao; etapa = 'Etapa Inventada' })
        Conferir 'etapa fora da lista é recusada' ($res.isError -and $res.content[0].text -match 'lista')
        $ontem = (Get-Date).AddDays(-1).ToString('yyyy-MM-dd')
        $res = Chamar 'atualizar_atendimento' ([ordered]@{ codigo = 'AT-0001-0001'; versao = [int]$g.versao; dt_prox_acao = $ontem; autor = 'teste' })
        Conferir 'C34 atualizar_atendimento recusa próxima ação no passado' ($res.isError -and $res.content[0].text -match 'C34') $res.content[0].text
        $res = Chamar 'abrir_atendimento' ([ordered]@{ codigo_contato = 'CT-0001-0001'; etapa = 'Contato Inicial'; responsavel = 'Vendedor A'
                                                       dt_prox_acao = $ontem; autor = 'teste' })
        Conferir 'C34 abrir_atendimento recusa próxima ação no passado' ($res.isError -and $res.content[0].text -match 'C34') $res.content[0].text
        # retomada (item 4): o anterior tem de ser da mesma empresa do contato (o front também recusa).
        # Recusado, nada é gravado: o teste não deixa atendimento novo no banco.
        $res = Chamar 'abrir_atendimento' ([ordered]@{ codigo_contato = 'CT-0001-0001'; etapa = 'Contato Inicial'; responsavel = 'Vendedor A'
                                                       codigo_atendimento_anterior = 'AT-0004-0005'; autor = 'teste' })
        Conferir 'anterior de outra empresa é recusado' ($res.isError -and $res.content[0].text -match 'outra empresa') $res.content[0].text
        $cx = Dados (Chamar 'obter_cliente' ([ordered]@{ codigo_cliente = 1 }))
        $res = Chamar 'atualizar_contexto_cliente' ([ordered]@{ codigo_cliente = 1; ctx_versao = [int]$cx.ctx_versao; ctx_resumo = 'Cliente de teste; resumo gravado pelo MCP' })
        Conferir 'atualizar_contexto_cliente com a ctx_versao lida' (-not $res.isError -and [int](Dados $res).ctx_versao -eq [int]$cx.ctx_versao + 1) $res.content[0].text
        # DEV-17: a segunda sessão leu a mesma versão; a primeira gravou; a segunda recebe conflito com o conteúdo atual
        $res = Chamar 'atualizar_contexto_cliente' ([ordered]@{ codigo_cliente = 1; ctx_versao = [int]$cx.ctx_versao; ctx_resumo = 'Outra sessão' })
        Conferir 'DEV-17 segunda sessão com a ctx_versao antiga recebe conflito e o contexto atual' ($res.isError -and $res.content[0].text -match 'Conflito de contexto' -and $res.content[0].text -match 'RESUMO GRAVADO PELO MCP') $res.content[0].text
        $res = Chamar 'atualizar_contexto_cliente' ([ordered]@{ codigo_cliente = 1; ctx_resumo = 'Sem versão' })
        Conferir 'DEV-17 contexto sem ctx_versao é recusado' ($res.isError -and $res.content[0].text -match 'ctx_versao') $res.content[0].text
        $cli1 = Dados (Chamar 'obter_cliente' ([ordered]@{ codigo_cliente = 1 }))
        $res = Chamar 'atualizar_cadastro_cliente' ([ordered]@{ codigo_cliente = 1; versao = [int]$cli1.versao; observacoes_acrescentar = 'edição comercial em paralelo'; autor = 'teste' })
        $res2 = Chamar 'atualizar_contexto_cliente' ([ordered]@{ codigo_cliente = 1; ctx_versao = [int]$cli1.ctx_versao; ctx_familia = 'Família de teste' })
        Conferir 'DEV-17 edição comercial em paralelo não gera conflito de contexto' ((-not $res.isError) -and (-not $res2.isError)) ($res.content[0].text + ' / ' + $res2.content[0].text)
        $cli = Dados (Chamar 'obter_cliente' ([ordered]@{ codigo_cliente = 2 }))
        $res = Chamar 'atualizar_cadastro_cliente' ([ordered]@{ codigo_cliente = 2; versao = [int]$cli.versao; telefone = '123' })
        Conferir 'telefone inválido é recusado' ($res.isError -and $res.content[0].text -match 'Telefone')
        $res = Chamar 'atualizar_cadastro_cliente' ([ordered]@{ codigo_cliente = 2; versao = [int]$cli.versao; cidade = '  caxias   do sul ' })
        $g = Dados $res
        Conferir 'cidade normalizada em maiúsculas' ((-not $res.isError) -and ($g.alterados[0].depois -eq 'CAXIAS DO SUL')) $res.content[0].text
        $ultLog = Linhas-Visita 'SELECT TOP 1 usuario, usuario_tecnico, estacao, versao_cliente, autor_declarado FROM log_alteracoes ORDER BY id DESC'
        Conferir 'DEV-06B log do conector: estação, usuário técnico, versão e autor declarado separados' ([string]$ultLog[0].estacao -eq [Environment]::MachineName -and [string]$ultLog[0].usuario_tecnico -eq [Environment]::UserName -and [string]$ultLog[0].versao_cliente -match '^CONECTOR \d' -and ((([string]$ultLog[0].usuario -eq 'IA') -and -not $ultLog[0].autor_declarado) -or (([string]$ultLog[0].usuario -match '^IA \((.+)\)$') -and [string]$ultLog[0].autor_declarado -eq $Matches[1]))) ($ultLog[0] | ConvertTo-Json -Compress)
        $log = Dados (Chamar 'ultimas_alteracoes' ([ordered]@{ limite = 10 }))
        Conferir 'log de auditoria com origem IA' (@($log.alteracoes | Where-Object { $_.origem -eq 'IA' }).Count -ge 3)

        # Devolve ao estado inicial o que muda o resultado dos testes acima, para o teste
        # poder rodar de novo no mesmo banco (sem isto, a 2a execução falhava: o AT-0001-0001
        # deixava de estar atrasado e a cidade do cliente 2 já estava normalizada).
        Write-Host '--- manutenção (DEV-08): config.manutencao recusa gravação no ponto de escrita; leitura continua'
        Conferir 'motivo puro: sem manutenção e esquema suportado grava' ((Get-MotivoRecusaEscrita $script:EsquemasSuportados[0] '') -eq '')
        Conferir 'motivo puro: manutenção recusa e cita o dono' ((Get-MotivoRecusaEscrita $script:EsquemasSuportados[0] 'EST-X|ana|12|2026-10-19 10:00|migração') -match 'manutenção.*EST-X')
        Conferir 'motivo puro: esquema desconhecido recusa' ((Get-MotivoRecusaEscrita '0.9' '') -match 'esquema')
        # revisão final: só o bloqueio do motor se repete; recusa e erro de regra (mesmo citando "bloqueado") não
        Conferir 'transitório: bloqueio do ACE (pt-BR e inglês) se repete' ((Test-ErroTransitorio 'Não foi possível atualizar; bloqueado no momento.') -and (Test-ErroTransitorio 'Could not update; currently locked.'))
        Conferir 'transitório: recusa por esquema não se repete' (-not (Test-ErroTransitorio (Get-MotivoRecusaEscrita '0.9' '')))
        Conferir 'transitório: erro de regra que cita "Bloqueado" não se repete' (-not (Test-ErroTransitorio "Campo 'etapa': 'Bloqueado' não está na lista Etapa."))
        $atB = Dados (Chamar 'obter_atendimento' ([ordered]@{ codigo = 'AT-0001-0001' }))
        $res = Chamar 'atualizar_atendimento' ([ordered]@{ codigo = 'AT-0001-0001'; versao = [int]$atB.versao; etapa = 'Etapa Bloqueada' })
        Conferir 'etapa fora da lista com "Bloqueada" no nome: erro de regra, sem repetir como bloqueio' ($res.isError -and $res.content[0].text -match 'lista' -and $res.content[0].text -notmatch 'continuou bloqueado') $res.content[0].text
        $atM = Dados (Chamar 'obter_atendimento' ([ordered]@{ codigo = 'AT-0001-0001' }))
        Executa-Visita "INSERT INTO config (chave, valor) VALUES ('manutencao', ?)" @((P 'texto' 'EST-TESTE|teste|1|2026-10-19 10:00|migração de teste'))
        try {
            $res = Chamar 'atualizar_atendimento' ([ordered]@{ codigo = 'AT-0001-0001'; versao = [int]$atM.versao; tentativas = 7; autor = 'teste' })
            Conferir 'manutenção ligada: gravação recusada com a mensagem de manutenção' ($res.isError -and $res.content[0].text -match 'manutenção') $res.content[0].text
            Conferir 'a mensagem chega limpa (sem "Exceção ao chamar")' ($res.content[0].text -notmatch 'chamar') $res.content[0].text
            $res = Chamar 'criar_cliente' ([ordered]@{ empresa = 'EMPRESA MANUTENCAO TESTE'; autor = 'teste' })
            Conferir 'manutenção ligada: criação também recusada' ($res.isError -and $res.content[0].text -match 'manutenção') $res.content[0].text
            $atM2 = Dados (Chamar 'obter_atendimento' ([ordered]@{ codigo = 'AT-0001-0001' }))
            Conferir 'manutenção ligada: leitura continua e o registro não mudou (versão igual)' ([int]$atM2.versao -eq [int]$atM.versao)
            Conferir 'manutenção ligada: nada foi criado' ([int](Invoke-Valor "SELECT COUNT(*) FROM clientes WHERE empresa='EMPRESA MANUTENCAO TESTE'" @()) -eq 0)
        } finally { Executa-Visita "DELETE FROM config WHERE chave='manutencao'" }
        $res = Chamar 'atualizar_atendimento' ([ordered]@{ codigo = 'AT-0001-0001'; versao = [int]$atM.versao; tentativas = [int]$atM.tentativas; observacoes_acrescentar = 'depois da manutenção'; autor = 'teste' })
        Conferir 'manutenção desligada: gravação volta a funcionar' (-not $res.isError) $res.content[0].text

        Write-Host '--- DEV-08 (ciclo 018): versão mínima do conector no ponto de escrita'
        Conferir 'motivo puro: conector abaixo da mínima recusa' ((Get-MotivoRecusaEscrita $script:EsquemasSuportados[0] '' '1.12.0' '1.11.0') -match 'exige a 1\.12\.0')
        Conferir 'motivo puro: conector na mínima grava' ((Get-MotivoRecusaEscrita $script:EsquemasSuportados[0] '' '1.12.0' '1.12.0') -eq '')
        Conferir 'motivo puro: 1.12.0 passa da mínima 1.9.0 (comparação numérica)' ((Get-MotivoRecusaEscrita $script:EsquemasSuportados[0] '' '1.9.0' '1.12.0') -eq '')
        Conferir 'transitório: recusa por versão mínima não se repete' (-not (Test-ErroTransitorio (Get-MotivoRecusaEscrita $script:EsquemasSuportados[0] '' '9.0.0' '1.0.0')))
        $atV = Dados (Chamar 'obter_atendimento' ([ordered]@{ codigo = 'AT-0001-0001' }))
        Executa-Visita "INSERT INTO config (chave, valor) VALUES ('versao_minima_conector', ?)" @((P 'texto' '99.0.0'))
        try {
            $res = Chamar 'atualizar_atendimento' ([ordered]@{ codigo = 'AT-0001-0001'; versao = [int]$atV.versao; tentativas = [int]$atV.tentativas + 1; autor = 'teste' })
            Conferir 'mínima acima deste conector: gravação recusada citando a versão' ($res.isError -and $res.content[0].text -match '99\.0\.0') $res.content[0].text
            $atV2 = Dados (Chamar 'obter_atendimento' ([ordered]@{ codigo = 'AT-0001-0001' }))
            Conferir 'mínima acima deste conector: leitura continua e nada mudou' ([int]$atV2.versao -eq [int]$atV.versao)
        } finally { Executa-Visita "DELETE FROM config WHERE chave='versao_minima_conector'" }
        Executa-Visita "INSERT INTO config (chave, valor) VALUES ('versao_minima_conector', ?)" @((P 'texto' $vServ))
        try {
            $st18 = Dados (Chamar 'crm_status' ([ordered]@{}))
            Conferir 'crm_status mostra a versão mínima do conector' ([string]$st18.versao_minima_conector -eq $vServ -and $vServ -match '^\d+\.\d+\.\d+$') ([string]$st18.versao_minima_conector + ' x ' + $vServ)
            $res = Chamar 'atualizar_atendimento' ([ordered]@{ codigo = 'AT-0001-0001'; versao = [int]$atV.versao; tentativas = [int]$atV.tentativas; observacoes_acrescentar = 'versão mínima atendida'; autor = 'teste' })
            Conferir 'mínima igual a este conector: gravação passa' (-not $res.isError) $res.content[0].text
        } finally { Executa-Visita "DELETE FROM config WHERE chave='versao_minima_conector'" }

        Write-Host '--- DEV-15: reativar a empresa reativa só a empresa; ação com versão antiga é recusada'
        $c15 = Dados (Chamar 'obter_cliente' ([ordered]@{ codigo_cliente = 2 }))
        $r15 = Dados (Chamar 'alterar_situacao_cliente' ([ordered]@{ codigo_cliente = 2; versao = [int]$c15.versao; acao = 'desativar'; autor = 'teste' }))
        Conferir 'desativar a empresa desativa os contatos (cascata)' ($r15.gravado -and [int]$r15.contatos_desativados -ge 1) ($r15 | ConvertTo-Json -Compress)
        $res = Chamar 'alterar_situacao_cliente' ([ordered]@{ codigo_cliente = 2; versao = [int]$c15.versao; acao = 'reativar'; autor = 'teste' })
        Conferir 'DEV-15 ação com a versão antiga é recusada (conflito)' ($res.isError -and $res.content[0].text -match 'Conflito') $res.content[0].text
        $c15b = Dados (Chamar 'obter_cliente' ([ordered]@{ codigo_cliente = 2 }))
        $r15 = Dados (Chamar 'alterar_situacao_cliente' ([ordered]@{ codigo_cliente = 2; versao = [int]$c15b.versao; acao = 'reativar'; autor = 'teste' }))
        $ct15 = Dados (Chamar 'obter_contato' ([ordered]@{ codigo_contato = 'CT-0002-0002' }))
        Conferir 'DEV-15 reativar a empresa não reativa os contatos' ($r15.gravado -and [int]$r15.contatos_reativados -eq 0 -and [int]$r15.contatos_inativos -ge 1 -and $ct15.ativo -eq $false) ($r15 | ConvertTo-Json -Compress)
        $res = Chamar 'atualizar_contato' ([ordered]@{ codigo_contato = 'CT-0002-0002'; versao = [int]$ct15.versao; ativo = $true; autor = 'teste' })
        Conferir 'o contato volta um a um (atualizar_contato ativo=true)' (-not $res.isError) $res.content[0].text

        Write-Host '--- DEV-13: identificador de operação (reenvio não duplica)'
        $maxLog13 = [int](Invoke-Valor 'SELECT MAX(id) FROM log_alteracoes' @())
        $nOp13 = [int](Invoke-Valor 'SELECT COUNT(*) FROM oportunidades' @())
        $arg13 = [ordered]@{ codigo_contato = 'CT-0001-0001'; etapa = 'Contato Inicial'; responsavel = 'Vendedor A'; id_operacao = 'TESTE-OP-0001-ABCDEF'; autor = 'teste' }
        $r1 = Dados (Chamar 'abrir_atendimento' $arg13)
        $r2 = Dados (Chamar 'abrir_atendimento' $arg13)
        Conferir 'DEV-13 reenvio com o mesmo id_operacao devolve o mesmo atendimento sem criar outro' ($r1.gravado -and $r2.repetida -and $r2.codigo -eq $r1.codigo -and [int](Invoke-Valor 'SELECT COUNT(*) FROM oportunidades' @()) -eq $nOp13 + 1) (($r1 | ConvertTo-Json -Compress) + ' / ' + ($r2 | ConvertTo-Json -Compress))
        Conferir 'DEV-13 o reenvio não gera log novo' ([int](Invoke-Valor 'SELECT COUNT(*) FROM log_alteracoes WHERE id > ?' @((P 'inteiro' $maxLog13))) -eq 1)
        $arg13.etapa = 'Negociação'; $arg13.familia = 'X'
        $res = Chamar 'abrir_atendimento' $arg13
        Conferir 'DEV-13 mesmo id_operacao com outro conteúdo é recusado' ($res.isError -and $res.content[0].text -match 'outro conteúdo') $res.content[0].text
        Executa-Visita 'DELETE FROM operacoes WHERE id_operacao=?' @((P 'texto' 'TESTE-OP-0001-ABCDEF'))
        Executa-Visita 'DELETE FROM oportunidades WHERE codigo=?' @((P 'texto' ([string]$r1.codigo)))
        Executa-Visita 'DELETE FROM log_alteracoes WHERE id > ?' @((P 'inteiro' $maxLog13))

        Write-Host '--- DEV-16: anterior pelo atualizar_atendimento (mesma empresa, sem ciclo)'
        $todos16 = Linhas-Visita 'SELECT codigo, id_cliente FROM oportunidades ORDER BY id'
        $grupo16 = $todos16 | Group-Object { [int]$_.id_cliente } | Where-Object { $_.Count -ge 2 } | Select-Object -First 1
        $par16 = @($grupo16.Group | Select-Object -First 2 | ForEach-Object { [string]$_.codigo })
        $aA = Dados (Chamar 'obter_atendimento' ([ordered]@{ codigo = $par16[0] }))
        $res = Chamar 'atualizar_atendimento' ([ordered]@{ codigo = $par16[0]; versao = [int]$aA.versao; codigo_atendimento_anterior = $par16[1]; autor = 'teste' })
        Conferir 'atualizar_atendimento grava o anterior da mesma empresa' (-not $res.isError) $res.content[0].text
        $aB = Dados (Chamar 'obter_atendimento' ([ordered]@{ codigo = $par16[1] }))
        $res = Chamar 'atualizar_atendimento' ([ordered]@{ codigo = $par16[1]; versao = [int]$aB.versao; codigo_atendimento_anterior = $par16[0]; autor = 'teste' })
        Conferir 'C33 anterior que fecharia ciclo é recusado no conector' ($res.isError -and $res.content[0].text -match 'ciclo') $res.content[0].text
        $aA = Dados (Chamar 'obter_atendimento' ([ordered]@{ codigo = $par16[0] }))
        $res = Chamar 'atualizar_atendimento' ([ordered]@{ codigo = $par16[0]; versao = [int]$aA.versao; codigo_atendimento_anterior = $null; autor = 'teste' })
        Conferir 'anterior removido com null' (-not $res.isError) $res.content[0].text

        Write-Host '--- DEV-28: duplicar atendimento (simular não grava; gravar copia, zera e liga o anterior)'
        $maxLog28 = [int](Invoke-Valor 'SELECT MAX(id) FROM log_alteracoes' @())
        $n28 = [int](Invoke-Valor 'SELECT COUNT(*) FROM oportunidades' @())
        $etIni = [string](Invoke-Valor "SELECT TOP 1 valor FROM listas WHERE tipo='Etapa' AND ativo=True AND encerra=False ORDER BY ordem, id" @())
        $pai28 = (Linhas-Visita 'SELECT * FROM oportunidades WHERE codigo=?' @((P 'texto' 'AT-0001-0001')))[0]
        $s28 = Invoke-DuplicarAtendimento ([pscustomobject]@{ codigo = 'AT-0001-0001' })
        Conferir 'DEV-28 duplicar simula por padrão e não grava' ($s28.simulado -eq $true -and [int](Invoke-Valor 'SELECT COUNT(*) FROM oportunidades' @()) -eq $n28) ($s28 | ConvertTo-Json -Compress -Depth 4)
        Conferir 'DEV-28 a cópia traz etapa inicial, o anterior e nada dos campos zerados' ($s28.dados.etapa -eq $etIni -and $s28.dados.codigo_atendimento_anterior -eq 'AT-0001-0001' -and -not $s28.dados.Contains('valor') -and -not $s28.dados.Contains('dt_prox_acao') -and -not $s28.dados.Contains('retomar_em')) ($s28.dados | ConvertTo-Json -Compress)
        $g28 = Invoke-DuplicarAtendimento ([pscustomobject]@{ codigo = 'AT-0001-0001'; simular = $false; id_operacao = 'TESTE-DUP-0001-ABCDEF'; autor = 'teste' })
        $novo28 = (Linhas-Visita 'SELECT * FROM oportunidades WHERE codigo=?' @((P 'texto' ([string]$g28.codigo))))[0]
        $copiaOk = $true; $difs28 = @()
        foreach ($c in $script:CamposDuplicar.Split(',')) { if ([string]$novo28[$c] -ne [string]$pai28[$c]) { $copiaOk = $false; $difs28 += $c } }
        Conferir 'DEV-28 gravar copia responsável, origem, prioridade, categoria, tipo de venda, máquina, família e última interação' ($g28.gravado -and $copiaOk) ($difs28 -join ',')
        Conferir 'DEV-28 gravar zera valor, orçamento, proposta, agenda, desfecho, retomada e tentativas' ($null -eq $novo28.valor -and $null -eq $novo28.orcamento -and $null -eq $novo28.dt_proposta -and $null -eq $novo28.prox_acao -and $null -eq $novo28.dt_prox_acao -and $null -eq $novo28.motivo_desfecho -and $null -eq $novo28.dt_desfecho -and $null -eq $novo28.retomar_em -and $null -eq $novo28.tentativas) ($novo28 | ConvertTo-Json -Compress)
        Conferir 'DEV-28 dt_entrada hoje, etapa inicial, mesmo contato e anterior = origem' ([string]$novo28.dt_entrada -eq (Get-Date).ToString('yyyy-MM-dd') -and [string]$novo28.etapa -eq $etIni -and [int]$novo28.id_contato -eq [int]$pai28.id_contato -and [int]$novo28.id_atendimento_anterior -eq [int]$pai28.id -and [string]$g28.duplicado_de -eq 'AT-0001-0001') ($novo28 | ConvertTo-Json -Compress)
        $g28b = Invoke-DuplicarAtendimento ([pscustomobject]@{ codigo = 'AT-0001-0001'; simular = $false; id_operacao = 'TESTE-DUP-0001-ABCDEF'; autor = 'teste' })
        Conferir 'DEV-28 repetir com o mesmo id_operacao não duplica de novo' ($g28b.repetida -and [string]$g28b.codigo -eq [string]$g28.codigo -and [int](Invoke-Valor 'SELECT COUNT(*) FROM oportunidades' @()) -eq $n28 + 1) ($g28b | ConvertTo-Json -Compress)
        $enc28 = [string](Invoke-Valor "SELECT TOP 1 o.codigo FROM oportunidades o WHERE o.etapa IN (SELECT valor FROM listas WHERE tipo='Etapa' AND encerra=True) ORDER BY o.id" @())
        $se28 = $null; if ($enc28) { $se28 = Invoke-DuplicarAtendimento ([pscustomobject]@{ codigo = $enc28 }) }
        Conferir 'DEV-28 duplicar atendimento encerrado é permitido (etapa inicial)' ($enc28 -and $se28.dados.etapa -eq $etIni) $enc28
        Conferir 'DEV-28 atendimento inexistente é recusado' ((Mensagem { Invoke-DuplicarAtendimento ([pscustomobject]@{ codigo = 'AT-9999-9999' }) }) -match 'encontrad')
        Executa-Visita 'DELETE FROM operacoes WHERE id_operacao=?' @((P 'texto' 'TESTE-DUP-0001-ABCDEF'))
        Executa-Visita 'DELETE FROM oportunidades WHERE codigo=?' @((P 'texto' ([string]$g28.codigo)))
        Executa-Visita 'DELETE FROM log_alteracoes WHERE id > ?' @((P 'inteiro' $maxLog28))

        Write-Host '--- restauração do banco de teste'
        # a data original é vencida (C34: o conector não grava próxima ação no passado): volta direto no banco
        Executa-Visita 'UPDATE oportunidades SET dt_prox_acao=? WHERE codigo=?' @((P 'data' (ConvertTo-DataOuNulo $at.dt_prox_acao)), (P 'texto' 'AT-0001-0001'))
        $atAgora = Dados (Chamar 'obter_atendimento' ([ordered]@{ codigo = 'AT-0001-0001' }))
        Conferir 'próxima ação do AT-0001-0001 restaurada' ([string]$atAgora.dt_prox_acao -eq [string]$at.dt_prox_acao) ([string]$atAgora.dt_prox_acao)
        # C34: a vencida já gravada não trava a gravação de outro campo (mesma data reenviada + outro campo)
        $res = Chamar 'atualizar_atendimento' ([ordered]@{ codigo = 'AT-0001-0001'; versao = [int]$atAgora.versao; dt_prox_acao = [string]$at.dt_prox_acao
                                                           tentativas = [int]$atAgora.tentativas + 1; autor = 'teste' })
        Conferir 'C34 próxima ação vencida já gravada não trava outro campo' (-not $res.isError) $res.content[0].text
        $atAgora = Dados (Chamar 'obter_atendimento' ([ordered]@{ codigo = 'AT-0001-0001' }))
        $res = Chamar 'atualizar_atendimento' ([ordered]@{ codigo = 'AT-0001-0001'; versao = [int]$atAgora.versao; tentativas = [int]$atAgora.tentativas - 1; autor = 'teste' })
        Conferir 'tentativas do AT-0001-0001 restauradas' (-not $res.isError) $res.content[0].text
        $cliAgora = Dados (Chamar 'obter_cliente' ([ordered]@{ codigo_cliente = 2 }))
        $res = Chamar 'atualizar_cadastro_cliente' ([ordered]@{ codigo_cliente = 2; versao = [int]$cliAgora.versao; cidade = [string]$cli.cidade; autor = 'teste' })
        Conferir 'cidade do cliente 2 restaurada' (-not $res.isError) $res.content[0].text
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
