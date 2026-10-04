<#
  testar-build-lib.ps1 - testes das funções do build que NÃO precisam de Excel:
  versão (VERSAO.txt, próxima versão, VERSAO_FRONT), contrato de Início!B7 e
  leitura do .xlsm pronto pelo zip. Roda em qualquer máquina com PowerShell 5.1.
#>
$ErrorActionPreference = 'Stop'
$dados = Split-Path -Parent $PSScriptRoot
. (Join-Path $dados 'lib\Comum.ps1')
. (Join-Path $dados 'build\lib\Versao.ps1')
. (Join-Path $dados 'build\lib\Pacote.ps1')

$script:falhas = 0
function Conferir([string]$nome, [bool]$ok, [string]$detalhe = '') {
    if ($ok) { Write-Host ('  OK    ' + $nome) -ForegroundColor Green }
    else { $script:falhas++; Write-Host ('  FALHA ' + $nome + '  ' + $detalhe) -ForegroundColor Red }
}

Write-Host '--- versão'
Conferir '1.4 -> 1.5' ((Get-ProximaVersao '1.4') -eq '1.5')
Conferir '1.9 -> 1.10' ((Get-ProximaVersao '1.9') -eq '1.10')
Conferir '2.0 -> 2.1' ((Get-ProximaVersao '2.0') -eq '2.1')
$erro = $null; try { Get-ProximaVersao 'x' | Out-Null } catch { $erro = $_ }
Conferir 'versão inválida recusada' ($null -ne $erro)
Conferir 'VERSAO.txt do projeto é X.Y' ((Get-VersaoAtual $dados) -match '^\d+\.\d+$')
Conferir '1.10 maior que 1.9' ((Compare-Versao '1.10' '1.9') -eq 1)
Conferir '1.5 igual a 1.5' ((Compare-Versao '1.5' '1.5') -eq 0)
Conferir '1.5 menor que 2.0' ((Compare-Versao '1.5' '2.0') -eq -1)
# versão de partida: VERSAO.txt x VERSAO-FRONT.txt publicada na raiz
$tmpV = Join-Path ([System.IO.Path]::GetTempPath()) ('crm-versao-' + $PID)
$tmpDados = Join-Path $tmpV 'SISTEMA\DADOS'
New-Item -ItemType Directory -Path $tmpDados -Force | Out-Null
try {
    Set-VersaoArquivo $tmpDados '1.5'
    Conferir 'sem VERSAO-FRONT.txt parte do VERSAO.txt' ((Get-VersaoDePartida $tmpDados $tmpV) -eq '1.5')
    [System.IO.File]::WriteAllText((Join-Path $tmpV 'VERSAO-FRONT.txt'), "1.7`r`nAmbiente: PRODUCAO`r`n")
    Conferir 'publicada maior vence (pacote com VERSAO.txt antigo)' ((Get-VersaoDePartida $tmpDados $tmpV) -eq '1.7')
    Set-VersaoArquivo $tmpDados '1.8'
    Conferir 'VERSAO.txt maior vence' ((Get-VersaoDePartida $tmpDados $tmpV) -eq '1.8')
    [System.IO.File]::WriteAllText((Join-Path $tmpV 'VERSAO-FRONT.txt'), "lixo`r`n")
    Conferir 'VERSAO-FRONT.txt inválido é ignorado' ((Get-VersaoDePartida $tmpDados $tmpV) -eq '1.8')
} finally { Remove-Item -LiteralPath $tmpV -Recurse -Force -ErrorAction SilentlyContinue }

$mod = [System.IO.File]::ReadAllText((Join-Path $dados 'fonte\modulos\modConfig.bas'), [System.Text.Encoding]::GetEncoding(1252))
$novo = Set-VersaoNoModConfig $mod '9.99'
Conferir 'VERSAO_FRONT trocada' ($novo -match 'Public Const VERSAO_FRONT As String = "9\.99"')
Conferir 'só a constante mudou' (($novo -replace '"9\.99"', '"X"') -eq ($mod -replace 'VERSAO_FRONT As String = "[^"]*"', 'VERSAO_FRONT As String = "X"'))

Write-Host '--- contrato de Início!B7'
$quando = [datetime]::new(2026, 9, 19, 8, 5, 0)
$b7 = Format-TextoB7 'producao' '1.5' $quando
Conferir 'texto exato' ($b7 -eq 'Ambiente: PRODUCAO   |   Versao 1.5   |   Publicada em 19/09/2026 08:05') $b7
Conferir 'padrão aceito' ($null -eq (Test-TextoB7 $b7 '1.5'))
Conferir 'versão divergente recusada' ($null -ne (Test-TextoB7 $b7 '1.6'))
Conferir 'espaçamento diferente recusado' ($null -ne (Test-TextoB7 ($b7 -replace '   \|', '  |') ''))
Conferir 'data com barra de outro formato recusada' ($null -ne (Test-TextoB7 ($b7 -replace '19/09/2026', '2026-09-19') ''))

Write-Host '--- leitura do pacote .xlsm (zip Open XML)'
$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('crm-teste-' + [guid]::NewGuid().ToString('N') + '.xlsm')
$zip = [System.IO.Compression.ZipFile]::Open($tmp, 'Create')
function Parte($nome, $texto) {
    $e = $zip.CreateEntry($nome); $w = New-Object System.IO.StreamWriter($e.Open(), (New-Object System.Text.UTF8Encoding($false)))
    $w.Write($texto); $w.Dispose()
}
Parte '[Content_Types].xml' '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/></Types>'
Parte 'xl/workbook.xml' '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Inicio" sheetId="1" r:id="rId1"/><sheet name="Clientes" sheetId="2" r:id="rId2"/></sheets></workbook>'
Parte 'xl/_rels/workbook.xml.rels' '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="x" Target="/xl/worksheets/sheet2.xml"/></Relationships>'
Parte 'xl/sharedStrings.xml' ('<?xml version="1.0"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>CRM</t></si><si><t>' + $b7 + '</t></si></sst>')
Parte 'xl/worksheets/sheet1.xml' '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="7"><c r="B7" t="s"><v>1</v></c></row><row r="8"><c r="B8" t="inlineStr"><is><t>inline</t></is></c></row></sheetData></worksheet>'
Parte 'xl/worksheets/sheet2.xml' '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>'
Parte 'xl/vbaProject.bin' 'x'
$zip.Dispose()
try {
    Conferir 'B7 por sharedStrings' ((Get-TextoCelulaXlsx $tmp 'Inicio' 'B7') -eq $b7)
    Conferir 'célula inlineStr' ((Get-TextoCelulaXlsx $tmp 'Inicio' 'B8') -eq 'inline')
    Conferir 'célula vazia devolve nulo' ($null -eq (Get-TextoCelulaXlsx $tmp 'Inicio' 'C9'))
    Conferir 'pacote válido' ((Test-PacoteXlsm $tmp @('Inicio', 'Clientes')).Count -eq 0)
    $p = Test-PacoteXlsm $tmp @('Inicio', 'Painel')
    Conferir 'aba faltando detectada' ($p.Count -eq 1 -and $p[0] -match 'Painel') ($p -join '; ')
} finally { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }

Write-Host ''
if ($script:falhas -eq 0) { Write-Host 'TUDO CERTO.' -ForegroundColor Green; exit 0 }
Write-Host ($script:falhas.ToString() + ' falha(s).') -ForegroundColor Red; exit 1
