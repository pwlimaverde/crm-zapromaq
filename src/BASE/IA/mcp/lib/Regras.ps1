<#
  Regras.ps1 - normalizacao e calculos, identicos aos do front.

  Espelho de modValidacao.Normalizar e modCalc (VBA). Se a regra mudar la,
  muda aqui na mesma versao - o mesmo valor tem de entrar igual no banco,
  venha do vendedor ou da IA.
#>

$script:EtapasEncerradas = @('Pedido Fechado', 'Perdido', 'Descartado')
$script:EtapasProspeccao = @('Contato Inicial', 'Sem Retorno', 'Retorno Agendado', 'Descartado')
# etapas que exigem familia de maquina (modValidacao.CriticarOportunidade)
$script:EtapasComFamilia = @(
    ('Elabora' + [char]0x00E7 + [char]0x00E3 + 'o da Proposta'),
    ('Negocia' + [char]0x00E7 + [char]0x00E3 + 'o'),
    ('Proposta em An' + [char]0x00E1 + 'lise'),
    'Proposta em Stand By')

function Get-SoDigitos([string]$s) {
    if ($null -eq $s) { return '' }
    return ($s -replace '[^0-9]', '')
}

# Normaliza pelo NOME do campo. Campo desconhecido so leva Trim.
# Devolve $null quando vazio - nunca cadeia vazia (indice unico de CNPJ).
function ConvertTo-Normalizado([string]$campo, $valor) {
    if ($null -eq $valor) { return $null }
    $s = ([string]$valor).Replace([string][char]160, ' ').Trim()
    while ($s.Contains('  ')) { $s = $s.Replace('  ', ' ') }
    switch ($campo.ToLowerInvariant()) {
        { $_ -in @('empresa', 'cidade', 'uf', 'maquina', 'orcamento') } { $s = $s.ToUpper([Globalization.CultureInfo]'pt-BR') }
        'email'    { $s = $s.ToLowerInvariant() }
        'cnpj'     { $s = Get-SoDigitos $s }
        'telefone' { $s = Get-SoDigitos $s }
    }
    if ($s.Length -eq 0) { return $null }
    return $s
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

function ConvertTo-DataOuNulo($v) {
    if ($null -eq $v -or [string]$v -eq '') { return $null }
    if ($v -is [datetime]) { return $v.Date }
    return [datetime]::ParseExact(([string]$v).Substring(0, 10), 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
}

# ---------------------------------------------------------------- modCalc
function Get-Situacao([bool]$temContato, [string]$etapa, $dtProxAcao, $retomarEm) {
    if (-not $temContato) { return '' }
    if ([string]::IsNullOrEmpty($etapa)) { return 'Sem etapa definida' }
    if ($script:EtapasEncerradas -contains $etapa) { return 'Encerrado' }
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
    if ($script:EtapasProspeccao -contains $etapa) { return ('1. Prospec' + [char]0x00E7 + [char]0x00E3 + 'o') }
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
