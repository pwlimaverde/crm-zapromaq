<#
  Versao.ps1 - número da versão do front e o texto de Início!B7.

  CONTRATO com o .bat das estações (fora do projeto): ele compara o TEXTO
  INTEIRO de Início!B7 para decidir se copia a versão nova. Célula e padrão
  não mudam:  Ambiente: <AMB>   |   Versao <X.Y>   |   Publicada em <dd/MM/yyyy HH:mm>
#>

$script:PadraoB7 = '^Ambiente: (?<amb>[A-Z]+)   \|   Versao (?<ver>\d+\.\d+)   \|   Publicada em (?<data>\d{2}/\d{2}/\d{4} \d{2}:\d{2})$'

function Get-VersaoAtual([string]$pastaDados) {
    $arq = Join-Path $pastaDados 'VERSAO.txt'
    $v = ([System.IO.File]::ReadAllText($arq)).Trim()
    if ($v -notmatch '^\d+\.\d+$') { throw ('VERSAO.txt inválido: "' + $v + '" (esperado X.Y, ex.: 1.4)') }
    return $v
}

# Versão publicada na raiz de BASE (VERSAO-FRONT.txt, escrito pelo build ao publicar).
# $null se o arquivo não existe ou não começa por X.Y.
function Get-VersaoPublicada([string]$raiz) {
    $arq = Join-Path $raiz 'VERSAO-FRONT.txt'
    if (-not (Test-Path -LiteralPath $arq)) { return $null }
    $linha = @([System.IO.File]::ReadAllLines($arq))[0]
    if ($null -eq $linha) { return $null }
    $linha = $linha.Trim()
    if ($linha -notmatch '^\d+\.\d+$') { return $null }
    return $linha
}

# -1, 0 ou 1, comparando número a número (1.10 > 1.9)
function Compare-Versao([string]$a, [string]$b) {
    $pa = $a.Split('.'); $pb = $b.Split('.')
    for ($i = 0; $i -lt 2; $i++) {
        if ([int]$pa[$i] -gt [int]$pb[$i]) { return 1 }
        if ([int]$pa[$i] -lt [int]$pb[$i]) { return -1 }
    }
    return 0
}

# A versão de onde a próxima parte: a MAIOR entre VERSAO.txt e a publicada.
# A BASE pode chegar à rede num pacote com VERSAO.txt mais antigo que a última
# publicação feita lá; recomeçar dele repetiria um número já publicado e o
# INICIAR-CRM das estações acharia que não há nada novo.
function Get-VersaoDePartida([string]$pastaDados, [string]$raiz) {
    $doArquivo = Get-VersaoAtual $pastaDados
    $publicada = Get-VersaoPublicada $raiz
    if ($publicada -and (Compare-Versao $publicada $doArquivo) -gt 0) { return $publicada }
    return $doArquivo
}

# 1.4 -> 1.5 ; 1.9 -> 1.10 (casa final numérica, nunca "2.0" sozinho)
function Get-ProximaVersao([string]$versao) {
    if ($versao -notmatch '^(\d+)\.(\d+)$') { throw ('Versão inválida: ' + $versao) }
    return ($Matches[1] + '.' + ([int]$Matches[2] + 1))
}

function Set-VersaoArquivo([string]$pastaDados, [string]$versao) {
    [System.IO.File]::WriteAllText((Join-Path $pastaDados 'VERSAO.txt'), $versao + "`r`n", (New-Object System.Text.UTF8Encoding($false)))
}

function Format-TextoB7([string]$ambiente, [string]$versao, [datetime]$quando) {
    return ('Ambiente: ' + $ambiente.ToUpperInvariant() + '   |   Versao ' + $versao + '   |   Publicada em ' +
            $quando.ToString('dd/MM/yyyy HH:mm', [Globalization.CultureInfo]::InvariantCulture))
}

# Devolve $null se o texto obedece ao contrato, ou a descrição do problema.
function Test-TextoB7([string]$texto, [string]$versaoEsperada) {
    if ($texto -notmatch $script:PadraoB7) { return ('Início!B7 fora do padrão: "' + $texto + '"') }
    if ($versaoEsperada -and $Matches['ver'] -ne $versaoEsperada) {
        return ('Início!B7 com versão ' + $Matches['ver'] + ', esperado ' + $versaoEsperada)
    }
    return $null
}

# Troca o valor da constante VERSAO_FRONT no texto de modConfig.bas.
function Set-VersaoNoModConfig([string]$texto, [string]$versao) {
    $rx = '(?m)^(Public Const VERSAO_FRONT As String = ")[^"]*(")'
    if ($texto -notmatch $rx) { throw 'modConfig.bas: constante VERSAO_FRONT não encontrada' }
    return [regex]::Replace($texto, $rx, ('${1}' + $versao + '${2}'))
}
