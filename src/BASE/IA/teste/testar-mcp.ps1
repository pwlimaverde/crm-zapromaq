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
$script:CaminhoBanco = Join-Path (Split-Path -Parent $raizIA) 'crm_zapromaq.accdb'
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
$filaFech = Linhas-Visita "SELECT codigo FROM oportunidades WHERE etapa='Pedido Fechado' ORDER BY id"
$codFech = [string]$filaFech[0].codigo
$opAntes = (Linhas-Visita "SELECT id, observacoes, ultima_interacao, versao FROM oportunidades WHERE codigo='$codNeg'")[0]
$maxLogV = [int](Invoke-Valor 'SELECT MAX(id) FROM log_alteracoes' @())
$criadas = @()
try {
    $c1 = Invoke-CriarVisita ([pscustomobject]@{ codigo = $codNeg; tipo = 'Visita'; responsavel = 'Vendedor A'; observacoes = 'teste visita' })
    $criadas += $c1.id
    Conferir 'criar_visita grava PENDENTE na versão 1' ($c1.gravado -and $c1.situacao -eq 'PENDENTE' -and $c1.versao -eq 1) ($c1 | ConvertTo-Json -Compress)
    Conferir 'C13: segunda visita pendente no mesmo atendimento e recusada' ((Mensagem { Invoke-CriarVisita ([pscustomobject]@{ codigo = $codNeg }) }) -match 'C13')
    Conferir 'C08: visita em atendimento encerrado (Pedido Fechado) e recusada' ((Mensagem { Invoke-CriarVisita ([pscustomobject]@{ codigo = $codFech }) }) -match 'C08')
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
    $encerrado = [string](Linhas-Visita "SELECT codigo FROM oportunidades WHERE etapa='Pedido Fechado' ORDER BY id")[0].codigo
    Conferir 'gerar_visitas: atendimento encerrado é recusado e nada é gravado (C08)' ((Mensagem { Invoke-GerarVisitas ([pscustomobject]@{ id_planejamento = $pc.id; codigos = @($alvo, $encerrado); simular = $false }) }) -match 'Nada gravado') ''
    Conferir 'gerar_visitas: a recusa desfez tudo (mesma quantidade de visitas)' ([int](Invoke-Valor 'SELECT COUNT(*) FROM visitas' @()) -eq ($cntAntes + 1)) ''
    Conferir 'gerar_visitas: segunda pendente no mesmo atendimento é recusada (C13)' ((Mensagem { Invoke-GerarVisitas ([pscustomobject]@{ id_planejamento = $pc.id; codigos = @($alvo); simular = $false }) }) -match 'Nada gravado') ''
} finally {
    foreach ($i in $visT10) { Executa-Visita 'DELETE FROM visitas WHERE id=?' @((P 'inteiro' ([int]$i))) }
    foreach ($i in $planT10) { Executa-Visita 'DELETE FROM planejamentos WHERE id=?' @((P 'inteiro' $i)) }
    Executa-Visita 'DELETE FROM log_alteracoes WHERE id > ?' @((P 'inteiro' $maxLogT10))
}
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
        Conferir ('gravação liberada no esquema ' + $st.versao_esquema + ' (1.0 e 1.1)') (($st.versao_esquema -in @('1.0', '1.1')) -and $st.gravacao_liberada -eq $true)

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
        # retomada (item 4): o anterior tem de ser da mesma empresa do contato (o front também recusa).
        # Recusado, nada é gravado: o teste não deixa atendimento novo no banco.
        $res = Chamar 'abrir_atendimento' ([ordered]@{ codigo_contato = 'CT-0001-0001'; etapa = 'Contato Inicial'; responsavel = 'Vendedor A'
                                                       codigo_atendimento_anterior = 'AT-0004-0005'; autor = 'teste' })
        Conferir 'anterior de outra empresa é recusado' ($res.isError -and $res.content[0].text -match 'outra empresa') $res.content[0].text
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

        # Devolve ao estado inicial o que muda o resultado dos testes acima, para o teste
        # poder rodar de novo no mesmo banco (sem isto, a 2a execução falhava: o AT-0001-0001
        # deixava de estar atrasado e a cidade do cliente 2 já estava normalizada).
        Write-Host '--- restauração do banco de teste'
        $atAgora = Dados (Chamar 'obter_atendimento' ([ordered]@{ codigo = 'AT-0001-0001' }))
        $res = Chamar 'atualizar_atendimento' ([ordered]@{ codigo = 'AT-0001-0001'; versao = [int]$atAgora.versao; dt_prox_acao = [string]$at.dt_prox_acao; autor = 'teste' })
        Conferir 'próxima ação do AT-0001-0001 restaurada' (-not $res.isError) $res.content[0].text
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
