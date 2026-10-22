<#
  Regras.ps1 - normalizacao e calculos, identicos aos do front.

  Espelho de modValidacao.Normalizar e modCalc (VBA). Se a regra mudar la,
  muda aqui na mesma versao - o mesmo valor tem de entrar igual no banco,
  venha do vendedor ou da IA.
#>

# Papel de uma etapa (Get-PapelEtapas, da tabela listas). Etapa sem linha: aberta, funil, sem familia.
function Get-PapelEtapa([string]$etapa) {
    $mapa = Get-PapelEtapas
    $chave = ConvertTo-ChaveLista $etapa
    if ($mapa.ContainsKey($chave)) { return $mapa[$chave] }
    return @{ valor = $etapa; encerra = $false; resultado = ''; quadro = 'FUNIL'; exige_familia = $false }
}

function Get-SoDigitos([string]$s) {
    if ($null -eq $s) { return '' }
    return ($s -replace '[^0-9]', '')
}

# Normaliza pelo NOME do campo (seção 5.1 do plano de 19/10/2026, ciclo 009). Espelho de
# modValidacao.Normalizar (front), casos C22 a C24:
#   todo texto comercial em MAIÚSCULAS, aparado e com espaço horizontal repetido colapsado; observações, próxima
#   ação e contexto mantêm as quebras de linha; URL, caminho e e-mail dentro do texto ficam como digitados;
#   e-mail minúsculo e UM endereço (os demais ficam em $script:EmailsAdicionais, para ir às observações);
#   CNPJ e telefone só dígitos; lista = valor exato (só apara); ctx_arquivo = caminho como gravado.
# Devolve $null quando vazio - nunca cadeia vazia (indice unico de CNPJ).
$script:EmailsAdicionais = ''
$script:CamposLista = @('etapa', 'responsavel', 'familia', 'origem', 'categoria', 'prioridade', 'motivo_desfecho', 'qualificacao', 'segmento')
function ConvertTo-EspacosColapsados([string]$s) {
    $s = $s.Replace("`t", ' ').Trim()
    while ($s.Contains('  ')) { $s = $s.Replace('  ', ' ') }
    return $s
}
function ConvertTo-MaiusculasProtegidas([string]$s) {
    if ($s -eq '') { return '' }
    $pt = [Globalization.CultureInfo]'pt-BR'
    return ((@($s.Split(' ') | ForEach-Object {
        if ($_.Contains('@') -or $_.Contains('://') -or $_.Contains('\') -or $_.ToLowerInvariant().StartsWith('www.')) { $_ } else { $_.ToUpper($pt) }
    })) -join ' ')
}
function ConvertTo-TextoComercial([string]$s, [bool]$multilinha) {
    $linhas = $s.Replace("`r`n", "`n").Replace("`r", "`n").Split("`n")
    $saida = @($linhas | ForEach-Object { ConvertTo-MaiusculasProtegidas (ConvertTo-EspacosColapsados $_) })
    if ($multilinha) { return (($saida -join "`r`n").Trim("`r", "`n")) }
    return ((@($saida | Where-Object { $_ -ne '' })) -join ' ')
}
function Get-ListaEmails([string]$s) {
    return @(($s -replace '[;,/]', ' ').Split(' ') | Where-Object { $_.Trim() -ne '' } | ForEach-Object { $_.Trim().ToLowerInvariant() })
}
function ConvertTo-Normalizado([string]$campo, $valor) {
    if ($null -eq $valor) { return $null }
    $s = ([string]$valor).Replace([string][char]160, ' ')
    $c = $campo.ToLowerInvariant()
    if ($c -eq 'email') {
        $todos = @(Get-ListaEmails $s)   # @(): um endereço só voltaria como texto, e [0] seria a 1ª letra
        $s = if ($todos.Count -gt 0) { $todos[0] } else { '' }
        $script:EmailsAdicionais = if ($todos.Count -gt 1) { (@($todos | Select-Object -Skip 1)) -join '; ' } else { '' }
    }
    elseif ($c -in @('cnpj', 'telefone')) { $s = Get-SoDigitos $s }
    elseif ($script:CamposLista -contains $c) { $s = ConvertTo-EspacosColapsados $s }
    elseif ($c -eq 'ctx_arquivo') { $s = $s.Trim() }
    elseif ($c -in @('observacoes', 'prox_acao', 'ctx_resumo')) { $s = ConvertTo-TextoComercial $s $true }
    else { $s = ConvertTo-TextoComercial $s $false }
    if ($s.Length -eq 0) { return $null }
    return $s
}
# Endereços além do primeiro vão para as observações, numa linha própria (mesma regra do front).
function Add-EmailsAdicionais($observacoes, [string]$adicionais) {
    if (-not $adicionais) { return $observacoes }
    if ($null -eq $observacoes -or [string]$observacoes -eq '') { return ('E-MAIL ADICIONAL: ' + $adicionais) }
    if (([string]$observacoes).Contains($adicionais)) { return $observacoes }
    return ([string]$observacoes + "`r`n" + 'E-MAIL ADICIONAL: ' + $adicionais)
}

# Chave de comparacao de lista (C01): sem caixa, sem acento, espacos aparados e colapsados.
# Espelho de modValidacao.ChaveLista - a mesma regra nos dois lados. Dois valores com a mesma
# chave sao o mesmo valor (ex.: "Pedido Fechado" e "PEDIDO  FECHADO ").
function ConvertTo-ChaveLista([string]$texto) {
    if ($null -eq $texto) { return '' }
    $s = $texto.Replace([string][char]160, ' ').Trim().ToUpperInvariant()
    while ($s.Contains('  ')) { $s = $s.Replace('  ', ' ') }
    $sb = New-Object System.Text.StringBuilder
    foreach ($c in $s.Normalize([Text.NormalizationForm]::FormD).ToCharArray()) {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($c) -ne [Globalization.UnicodeCategory]::NonSpacingMark) { [void]$sb.Append($c) }
    }
    return $sb.ToString()
}

# Regras de lista da Fase 2 (C03, C04, C05). Espelho do modListas no front: mudar aqui exige mudar la.
$script:UFsBrasil = @('AC','AL','AP','AM','BA','CE','DF','ES','GO','MA','MT','MS','MG','PA','PB','PR','PE','PI','RJ','RN','RS','RO','RR','SC','SP','SE','TO')

# C05: UF só entra se for uma das 27 (maiúscula, sem espaço sobrando).
function Test-UFValida([string]$uf) {
    return ($script:UFsBrasil -ccontains $uf.ToUpperInvariant())
}

# C04: resultado (GANHO/PERDA) só existe em etapa que encerra. Devolve a mensagem ou $null.
function Test-PapelEtapaValido([bool]$encerra, [string]$resultado) {
    if ($resultado -notin @('', 'GANHO', 'PERDA')) { return 'Resultado deve ser GANHO, PERDA ou vazio.' }
    if ($resultado -ne '' -and -not $encerra) { return 'Resultado só pode ser GANHO ou PERDA em etapa que encerra o atendimento (C04).' }
    return $null
}

# C03: desativar uma etapa com resultado só é permitido se sobrar outra ATIVA com o mesmo resultado.
# $papeis: linhas com valor, resultado e ativo (todas as etapas).
function Test-DesativarEtapaPermitido($papeis, [string]$valor) {
    $chave = ConvertTo-ChaveLista $valor
    $alvo = @($papeis | Where-Object { (ConvertTo-ChaveLista $_.valor) -eq $chave }) | Select-Object -First 1
    if ($null -eq $alvo -or [string]$alvo.resultado -eq '') { return $true }
    $outras = @($papeis | Where-Object { $_.ativo -and (ConvertTo-ChaveLista $_.valor) -ne $chave -and [string]$_.resultado -eq [string]$alvo.resultado })
    return ($outras.Count -gt 0)
}
function Test-CnpjValido([string]$v) {
    $s = Get-SoDigitos $v
    if ($s.Length -ne 14) { return $false }
    if ($s -eq ([string]$s[0]) * 14) { return $false }
    $d = $s.ToCharArray() | ForEach-Object { [int][string]$_ }
    $soma = 0; $peso = 5
    for ($i = 0; $i -lt 12; $i++) { $soma += $d[$i] * $peso; $peso--; if ($peso -lt 2) { $peso = 9 } }
    $d1 = 11 - ($soma % 11); if ($d1 -ge 10) { $d1 = 0 }
    $soma = 0; $peso = 6
    for ($i = 0; $i -lt 13; $i++) { $soma += $d[$i] * $peso; $peso--; if ($peso -lt 2) { $peso = 9 } }
    $d2 = 11 - ($soma % 11); if ($d2 -ge 10) { $d2 = 0 }
    return ($d[12] -eq $d1 -and $d[13] -eq $d2)
}

function Get-CriticaTelefone($v) {
    $s = Get-SoDigitos ([string]$v)
    if ($s -eq '') { return '' }
    if ($s.Length -lt 10 -or $s.Length -gt 11) {
        return ('Telefone deve ter DDD mais o numero: 10 ou 11 digitos. Nome de quem atende ' +
                'e um segundo numero vao em observacoes.')
    }
    return ''
}

# Datas entram SO no formato ISO (AAAA-MM-DD): e o unico sem ambiguidade
# entre dd/mm e mm/dd. Devolve [datetime] ou $null; lanca se invalida.
function ConvertTo-Data([string]$campo, $valor) {
    if ($null -eq $valor -or [string]$valor -eq '') { return $null }
    $d = [datetime]::MinValue
    if (-not [datetime]::TryParseExact([string]$valor, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture,
                                       [Globalization.DateTimeStyles]::None, [ref]$d)) {
        throw ("Campo '" + $campo + "': data invalida '" + $valor + "'. Use o formato AAAA-MM-DD.")
    }
    return $d
}

# DEV-09 (ciclo 008): número sem perda silenciosa. Mesma matriz (C17 a C21) de modValidacao.Converter (front).
# Moeda: até 2 casas decimais e até 999.999.999.999,99; nada de arredondar em silêncio.
$script:MoedaMaxima = [decimal]999999999999.99
function ConvertTo-Moeda([string]$campo, $v) {
    if ($null -eq $v -or ($v -is [string] -and $v.Trim() -eq '')) { return $null }
    try { $d = [decimal]$v } catch { throw ("Campo '" + $campo + "': valor inválido '" + $v + "'.") }
    if ([decimal]::Round($d, 2) -ne $d) { throw ("Campo '" + $campo + "': mais de 2 casas decimais (" + $v + '). Nada foi gravado.') }
    if ([math]::Abs($d) -gt $script:MoedaMaxima) { throw ("Campo '" + $campo + "': valor acima do limite (999.999.999.999,99).") }
    return $d
}
# Inteiro: sem fração e dentro do Long do banco (nada de [int]1.5 = 2).
function ConvertTo-Inteiro([string]$campo, $v) {
    if ($null -eq $v -or ($v -is [string] -and $v.Trim() -eq '')) { return $null }
    try { $d = [decimal]$v } catch { throw ("Campo '" + $campo + "': número inválido '" + $v + "'.") }
    if ([decimal]::Truncate($d) -ne $d) { throw ("Campo '" + $campo + "': número inteiro, sem fração (" + $v + '). Nada foi gravado.') }
    if ($d -gt [int]::MaxValue -or $d -lt [int]::MinValue) { throw ("Campo '" + $campo + "': número fora do limite.") }
    return [int]$d
}

function ConvertTo-DataOuNulo($v) {
    if ($null -eq $v -or [string]$v -eq '') { return $null }
    if ($v -is [datetime]) { return $v.Date }
    return [datetime]::ParseExact(([string]$v).Substring(0, 10), 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
}

# Regras de etapa da seção 5.2 (ciclo 009, DEV-10). Espelho de modRegras (front), casos C25 a C27.
# Devolve @{ limpar = campos a limpar; orientacao = texto da reabertura ou '' }.
function Get-TransicaoEtapa([bool]$encerraAntiga, [bool]$encerraNova, [bool]$perdaNova) {
    $limpar = @()
    $orientacao = ''
    if ($encerraNova -and $perdaNova) { $limpar += @('prox_acao', 'dt_prox_acao') }
    if ($encerraAntiga -and -not $encerraNova) {
        $limpar += @('motivo_desfecho', 'dt_desfecho', 'retomar_em')
        $orientacao = ('Reabrir serve para corrigir um encerramento indevido: motivo, data do desfecho e retomada foram limpos ' +
                       '(os valores anteriores ficam no histórico). Interesse novo é atendimento NOVO: use duplicar_atendimento no encerrado.')
    }
    return @{ limpar = $limpar; orientacao = $orientacao }
}
# C26: retomar_em nova ou alterada só em Perdido/Descartado (encerra + PERDA). Devolve a mensagem ou $null.
function Test-RetomadaPermitida([bool]$encerraNova, [bool]$perdaNova, $retomarNova, $retomarLida) {
    $n = ConvertTo-DataOuNulo $retomarNova
    if ($null -eq $n) { return $null }
    if ($encerraNova -and $perdaNova) { return $null }
    if ((ConvertTo-DataOuNulo $retomarLida) -eq $n) { return $null }
    return 'A data de retomada só existe em atendimento Perdido ou Descartado. Para um atendimento aberto, use a próxima ação (dt_prox_acao) (C26).'
}

# C29 (DEV-11): o contato do atendimento existe, está ativo e é da empresa do atendimento. Espelho de
# modRegras.CriticarContatoDoAtendimento. Devolve a mensagem ou $null.
function Test-ContatoDoAtendimento([bool]$existe, [bool]$ativo, [int]$idClienteContato, [int]$idClienteAtendimento) {
    if (-not $existe) { return 'Contato não encontrado: nada foi gravado.' }
    if ($idClienteContato -ne $idClienteAtendimento) { return 'O contato é de outra empresa: o atendimento só usa contato da empresa dele. Nada foi gravado.' }
    if (-not $ativo) { return 'O contato está inativo: escolha um contato ativo da empresa. Nada foi gravado.' }
    return $null
}

# ---------------------------------------------------------------- modCalc
function Get-Situacao([bool]$temContato, [string]$etapa, $dtProxAcao, $retomarEm) {
    if (-not $temContato) { return '' }
    if ([string]::IsNullOrEmpty($etapa)) { return 'Sem etapa definida' }
    if ((Get-PapelEtapa $etapa).encerra) { return 'Encerrado' }
    $hoje = (Get-Date).Date
    $ret = ConvertTo-DataOuNulo $retomarEm
    if ($ret -and $ret -le $hoje) { return 'Retomar hoje' }
    $prox = ConvertTo-DataOuNulo $dtProxAcao
    if (-not $prox) { return ('Sem pr' + [char]0x00F3 + 'xima a' + [char]0x00E7 + [char]0x00E3 + 'o') }
    if ($prox -lt $hoje) { return ('A' + [char]0x00E7 + [char]0x00E3 + 'o atrasada') }
    # fim da semana corrente: domingo (Weekday vbMonday: segunda=1 .. domingo=7)
    $diaSemana = [int]$hoje.DayOfWeek; if ($diaSemana -eq 0) { $diaSemana = 7 }
    $fimSemana = $hoje.AddDays(7 - $diaSemana)
    if ($prox -le $fimSemana) { return ('A' + [char]0x00E7 + [char]0x00E3 + 'o nesta semana') }
    return 'Em dia'
}

function Get-Quadro([string]$etapa) {
    if ([string]::IsNullOrEmpty($etapa)) { return '' }
    if ((Get-PapelEtapa $etapa).quadro -eq 'PROSPECCAO') { return ('1. Prospec' + [char]0x00E7 + [char]0x00E3 + 'o') }
    return '2. Funil comercial'
}

function Get-DiasParado($ultimaInteracao) {
    $d = ConvertTo-DataOuNulo $ultimaInteracao
    if (-not $d) { return $null }
    return [int]((Get-Date).Date - $d).TotalDays
}

function Get-Ciclo($dtProposta, $dtDesfecho) {
    $a = ConvertTo-DataOuNulo $dtProposta; $b = ConvertTo-DataOuNulo $dtDesfecho
    if (-not $a -or -not $b) { return $null }
    return [int]($b - $a).TotalDays
}

function Format-CodigoCliente($codigo) {
    if ($null -eq $codigo) { return $null }
    return ([int]$codigo).ToString('0000')
}

# Regras de visita (C08 a C13, spec 003 secao "Regras de visita"). Espelho de modVisita (VBA).
# Devolvem a mensagem da recusa, ou $null quando a regra passa.
function Test-VisitaPodeCriar([bool]$atendimentoEncerrado) {
    if ($atendimentoEncerrado) { return 'Atendimento encerrado não recebe visita (C08).' }
    return $null
}

# datas em AAAA-MM-DD (comparação de texto é a de data nesse formato)
function Test-DataNoPeriodo([string]$data, [string]$inicio, [string]$fim) {
    if ($data -lt $inicio -or $data -gt $fim) { return ('Data ' + $data + ' fora do período do planejamento (' + $inicio + ' a ' + $fim + ') (C09).') }
    return $null
}

function Test-RealizadaValida([string]$relato, [string]$dataRealizada) {
    if ([string]::IsNullOrWhiteSpace($relato) -or [string]::IsNullOrWhiteSpace($dataRealizada)) { return 'Visita realizada exige data e relato (C10).' }
    return $null
}

function Test-CanceladaValida([string]$motivo) {
    if ([string]::IsNullOrWhiteSpace($motivo)) { return 'Visita cancelada exige motivo (C11).' }
    return $null
}

# Linha que vai para o histórico do atendimento quando a visita é realizada (C12).
function Get-TextoHistoricoVisita([datetime]$quando, [string]$tipo, [string]$resultado, [string]$relato) {
    $resumo = $relato
    if ($resumo.Length -gt 200) { $resumo = $resumo.Substring(0, 200) }
    return ('VISITA ' + $quando.ToString('dd/MM/yyyy HH:mm') + ' - ' + $tipo + ' - ' + $resultado + ': ' + $resumo)
}

# Só uma visita PENDENTE por atendimento por vez (C13).
function Test-SegundaPendente([int]$pendentesExistentes) {
    if ($pendentesExistentes -gt 0) { return 'Já há visita pendente neste atendimento: realize ou cancele antes (C13).' }
    return $null
}

# Planejamento (Fase 4, spec 003 secao "Regras de planejamento"). Espelho de modPlanejamento (VBA).

# Nome de município para comparar: a mesma chave das listas (sem acento, maiúscula, espaço).
function ConvertTo-NomeBusca([string]$nome) { return ConvertTo-ChaveLista $nome }

# Distância em km entre dois pontos (Haversine, raio da Terra 6.371 km).
function Get-DistanciaKm([double]$lat1, [double]$lon1, [double]$lat2, [double]$lon2) {
    $raio = 6371.0
    $graus = [math]::PI / 180
    $dLat = ($lat2 - $lat1) * $graus
    $dLon = ($lon2 - $lon1) * $graus
    $a = [math]::Sin($dLat / 2) * [math]::Sin($dLat / 2) +
         [math]::Cos($lat1 * $graus) * [math]::Cos($lat2 * $graus) * [math]::Sin($dLon / 2) * [math]::Sin($dLon / 2)
    return 2 * $raio * [math]::Asin([math]::Sqrt($a))
}

# P1 (C14): raio de 1 a 500 km.
function Test-RaioValido([int]$raio) {
    if ($raio -lt 1 -or $raio -gt 500) { return 'Raio deve ser de 1 a 500 km (C14).' }
    return $null
}

# P2 (C15): período preenchido e início não depois do fim (datas AAAA-MM-DD).
function Test-PeriodoValido([string]$inicio, [string]$fim) {
    if ([string]::IsNullOrWhiteSpace($inicio) -or [string]::IsNullOrWhiteSpace($fim)) { return 'Informe início e fim do período (C15).' }
    if ($inicio -gt $fim) { return 'O início é depois do fim do período (C15).' }
    return $null
}

# P3 (C16): não encerra com visita pendente sem cancelar_pendentes.
function Test-EncerramentoPermitido([int]$pendentes, [bool]$cancelarPendentes) {
    if ($pendentes -gt 0 -and -not $cancelarPendentes) { return ('Há ' + $pendentes + ' visita(s) pendente(s): cancele-as ou use cancelar_pendentes (C16).') }
    return $null
}

# P4: código PL-AAAA-NNN (NNN com 3 dígitos).
function Get-CodigoPlanejamento([int]$ano, [int]$numero) {
    return ('PL-' + $ano + '-' + $numero.ToString('000'))
}
