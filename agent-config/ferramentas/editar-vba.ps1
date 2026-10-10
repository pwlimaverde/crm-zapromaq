<#
  editar-vba.ps1 - troca UM trecho exato num fonte VBA (cp1252 + CRLF) sem tocar
  no resto do arquivo.

  Por que existe: a ferramenta Edit do agente le o arquivo como UTF-8; num fonte
  cp1252 com acento ela troca CADA byte de acento do arquivo inteiro por U+FFFD e
  regrava tudo em UTF-8 (EF BF BD) - o acento original se perde. Arquivo VBA so
  com ASCII nao e afetado. O verificar.py recusa o resultado (rede de seguranca);
  esta ferramenta e o jeito de editar sem corromper.

  O que garante (senao, nao grava e sai 1):
    - o arquivo e cp1252 de verdade: sem BOM, sem U+FFFD, so CRLF;
    - o trecho antigo aparece exatamente uma vez;
    - o texto novo cabe em cp1252 (sem caractere perdido); LF vira CRLF;
    - depois de gravar, os bytes antes e depois do trecho sao os mesmos de antes.

  Uso (texto direto ou em arquivo UTF-8 - o arquivo evita problema de aspas e acento
  na linha de comando):
    powershell -File agent-config\ferramentas\editar-vba.ps1 -Arquivo <fonte> -Antigo '<texto>' -Novo '<texto>'
    powershell -File agent-config\ferramentas\editar-vba.ps1 -Arquivo <fonte> -AntigoArquivo a.txt -NovoArquivo n.txt
  -Simular mostra o que mudaria sem gravar. Quebra de linha final dos arquivos de texto e ignorada.
#>
param([Parameter(Mandatory = $true)][string]$Arquivo,
      [string]$Antigo, [string]$Novo, [string]$AntigoArquivo, [string]$NovoArquivo,
      [switch]$Simular)
$ErrorActionPreference = 'Stop'
$cp = [System.Text.Encoding]::GetEncoding(1252)

function Parar([string]$msg) { [Console]::Error.WriteLine('RECUSADO: ' + $msg + ' (nada foi gravado)'); exit 1 }
function Texto([string]$direto, [string]$arq, [string]$nome, [bool]$informado) {
    if ($arq) {
        if (-not (Test-Path -LiteralPath $arq)) { Parar ($nome + 'Arquivo nao existe: ' + $arq) }
        try { $t = [System.IO.File]::ReadAllText($arq, (New-Object System.Text.UTF8Encoding($false, $true))) }
        catch { Parar ($nome + 'Arquivo nao e UTF-8 valido: ' + $arq) }
        $t = $t.TrimStart([char]0xFEFF).TrimEnd("`r", "`n")
    } elseif ($informado) { $t = $direto }
    else { Parar ('informe -' + $nome + ' ou -' + $nome + 'Arquivo') }
    # qualquer quebra de linha vira CRLF, como no fonte
    return (($t -replace "`r`n", "`n") -replace "`r", "`n") -replace "`n", "`r`n"
}

if (-not (Test-Path -LiteralPath $Arquivo)) { Parar ('arquivo nao existe: ' + $Arquivo) }
$caminho = (Resolve-Path -LiteralPath $Arquivo).ProviderPath
$antes = [System.IO.File]::ReadAllBytes($caminho)

# 1. o arquivo e cp1252 + CRLF de verdade
if ($antes.Length -ge 3 -and $antes[0] -eq 0xEF -and $antes[1] -eq 0xBB -and $antes[2] -eq 0xBF) { Parar 'arquivo com BOM UTF-8 - nao e fonte cp1252' }
for ($i = 0; $i -le $antes.Length - 3; $i++) {
    if ($antes[$i] -eq 0xEF -and $antes[$i + 1] -eq 0xBF -and $antes[$i + 2] -eq 0xBD) { Parar 'arquivo ja tem U+FFFD (acento perdido) - restaure do git antes de editar' }
}
for ($i = 0; $i -lt $antes.Length; $i++) {
    if ($antes[$i] -eq 10 -and ($i -eq 0 -or $antes[$i - 1] -ne 13)) { Parar 'arquivo tem LF sem CR - nao e o formato do fonte VBA' }
}

$a = Texto $Antigo $AntigoArquivo 'Antigo' $PSBoundParameters.ContainsKey('Antigo')
$n = Texto $Novo $NovoArquivo 'Novo' $PSBoundParameters.ContainsKey('Novo')
if ($a -eq '') { Parar 'trecho antigo vazio' }

# 2. o texto novo cabe em cp1252 (Encoding 1252 troca o que nao cabe por "?")
if ($cp.GetString($cp.GetBytes($n)) -ne $n) { Parar 'texto novo tem caractere que nao existe em Windows-1252' }

# 3. o trecho aparece uma vez so
$texto = $cp.GetString($antes)
$pos = $texto.IndexOf($a, [System.StringComparison]::Ordinal)
if ($pos -lt 0) { Parar 'trecho antigo nao encontrado' }
if ($texto.IndexOf($a, $pos + 1, [System.StringComparison]::Ordinal) -ge 0) { Parar 'trecho antigo aparece mais de uma vez - inclua mais contexto' }

# cp1252 e 1 byte por caractere: a posicao no texto e a posicao em bytes
$bNovo = $cp.GetBytes($n)
$lenA = $cp.GetBytes($a).Length
$depois = New-Object byte[] ($antes.Length - $lenA + $bNovo.Length)
[System.Array]::Copy($antes, 0, $depois, 0, $pos)
[System.Array]::Copy($bNovo, 0, $depois, $pos, $bNovo.Length)
[System.Array]::Copy($antes, $pos + $lenA, $depois, $pos + $bNovo.Length, $antes.Length - $pos - $lenA)

$linha = ($texto.Substring(0, $pos) -split "`r`n").Count
if ($Simular) {
    Write-Output ('SIMULACAO: linha ' + $linha + ', ' + $lenA + ' byte(s) -> ' + $bNovo.Length + ' byte(s). Nada gravado.')
    exit 0
}
[System.IO.File]::WriteAllBytes($caminho, $depois)

# 4. confere o que ficou no disco: prefixo e sufixo intactos, trecho novo no lugar
$gravado = [System.IO.File]::ReadAllBytes($caminho)
$ok = ($gravado.Length -eq $depois.Length)
for ($i = 0; $ok -and $i -lt $gravado.Length; $i++) { if ($gravado[$i] -ne $depois[$i]) { $ok = $false } }
if (-not $ok) {
    [System.IO.File]::WriteAllBytes($caminho, $antes)
    Parar 'o arquivo gravado nao confere com o esperado - original restaurado'
}
Write-Output ('OK: ' + $caminho + ' linha ' + $linha + ', ' + $lenA + ' byte(s) -> ' + $bNovo.Length + ' byte(s); resto do arquivo intacto.')
exit 0
