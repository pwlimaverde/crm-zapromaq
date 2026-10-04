<#
  gerar-visual.ps1 - gera a parte visual do front a partir de fonte\layout.

  Roda com o que existe em qualquer Windows 10/11: PowerShell 5.1 + Edge.
  Sem Node, sem internet (a prévia é JavaScript rodando no próprio Edge).

    1. execucao\previa\dados-layout.js   tema + ícones + amostras + telas, para a prévia
    2. fonte\assets\gerado\<tela>-fundo.bmp
                                          só a DECORAÇÃO (cartões, faixas, sombras, logo),
                                          renderizada pelo Edge em modo headless a 96 dpi e
                                          convertida para BMP 24 bits (o Picture do MSForms
                                          aceita bmp/gif/jpg, não PNG - referencias/01-vba-msforms)
    3. execucao\previa\<tela>-100|125|150.png
                                          a tela completa como ficará, nas escalas do Windows
    4. conferência feita pela própria prévia: elemento fora da tela, sobreposto
       ou com texto que não cabe -> lista de problemas (sai com código 1)
    5. fonte\modulos\modTema.bas          cores, fontes, ícones e estados para o VBA (gerado)
    6. fonte\assets\gerado\manifesto.json hash das entradas: o build sabe se está velho

  Uso: gerar-visual.ps1 [-Tela frmCRM] [-Tolerante]   (-Tolerante: não falha por problema de layout)
#>
param([string]$Tela = '', [switch]$Tolerante)
$ErrorActionPreference = 'Stop'
$dados = Split-Path -Parent $PSScriptRoot
. (Join-Path $dados 'lib\Comum.ps1')
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Web

$raiz = Get-RaizBase $PSScriptRoot
$fonte = Join-Path $dados 'fonte'
$layout = Join-Path $fonte 'layout'
$pastaPrevia = Get-PastaExecucao $raiz 'previa'
$pastaGerado = Join-Path $fonte 'assets\gerado'
if (-not (Test-Path -LiteralPath $pastaGerado)) { New-Item -ItemType Directory -Path $pastaGerado | Out-Null }
$log = New-Log $raiz 'visual'
$utf8 = New-Object System.Text.UTF8Encoding($false)
function Ler([string]$p) { return [System.IO.File]::ReadAllText($p, [System.Text.Encoding]::UTF8) }

$log.L('=== GERADOR VISUAL - CRM Zapromaq ===')

# ------------------------------------------------------------ 1. dados da prévia
$arqTelas = @(Get-ChildItem -LiteralPath (Join-Path $layout 'formularios') -Filter '*.json' | Sort-Object Name |
              Where-Object { (Ler $_.FullName) -match '"canvas"\s*:' })
if ($Tela) { $arqTelas = @($arqTelas | Where-Object { $_.BaseName -eq $Tela }) }
if ($arqTelas.Count -eq 0) { throw 'nenhuma tela com "canvas" em fonte\layout\formularios' }
$partes = @()
foreach ($a in $arqTelas) { $partes += ('"' + $a.BaseName + '":' + (Ler $a.FullName)) }
$js = 'window.CRM_LAYOUT = {"tema":' + (Ler (Join-Path $layout 'tema.json')) + ',"icones":' + (Ler (Join-Path $layout 'icones.json')) +
      ',"amostras":' + (Ler (Join-Path $layout 'amostras.json')) + ',"telas":{' + ($partes -join ',') + '}};'
[System.IO.File]::WriteAllText((Join-Path $pastaPrevia 'dados-layout.js'), $js, $utf8)
$log.L('telas: ' + (@($arqTelas | ForEach-Object { $_.BaseName }) -join ', '))

# ------------------------------------------------------------ Edge headless
$edge = @("${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe", "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe") |
        Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $edge) { throw 'Microsoft Edge não encontrado (necessário para gerar o visual).' }
$perfil = Join-Path $pastaPrevia 'perfil-edge'
$pagina = 'file:///' + ((Join-Path $PSScriptRoot 'visual\previa.html') -replace '\\', '/')

# O modo headless novo é o padrão do Chromium 132+ (referencias/10-edge-headless): --headless basta.
function Invoke-Edge([string[]]$extra, [switch]$Saida) {
    $base = @('--headless', '--disable-gpu', '--hide-scrollbars', '--no-first-run', '--no-default-browser-check',
              '--disable-extensions', '--virtual-time-budget=4000', ('--user-data-dir=' + $perfil))
    $argumentos = $base + $extra
    if ($Saida) {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $edge; $psi.UseShellExecute = $false; $psi.RedirectStandardOutput = $true; $psi.RedirectStandardError = $true
        $psi.StandardOutputEncoding = [System.Text.Encoding]::UTF8
        $psi.Arguments = (($argumentos | ForEach-Object { '"' + $_ + '"' }) -join ' ')
        $p = [System.Diagnostics.Process]::Start($psi)
        $err = $p.StandardError.ReadToEndAsync()
        $out = $p.StandardOutput.ReadToEnd(); $p.WaitForExit(); [void]$err.Result
        return $out
    }
    $p = Start-Process -FilePath $edge -ArgumentList ($argumentos | ForEach-Object { '"' + $_ + '"' }) -PassThru -WindowStyle Hidden
    if (-not $p.WaitForExit(60000)) { $p.Kill(); throw 'o Edge não terminou a captura em 60 s' }
}

function Save-Captura([string]$url, [int]$larguraPx, [int]$alturaPx, [double]$escala, [string]$png) {
    Remove-Item -LiteralPath $png -Force -ErrorAction SilentlyContinue
    Invoke-Edge @(('--force-device-scale-factor=' + $escala.ToString([Globalization.CultureInfo]::InvariantCulture)),
                  ('--window-size=' + $larguraPx + ',' + $alturaPx), ('--screenshot=' + $png), $url)
    if (-not (Test-Path -LiteralPath $png)) { throw ('o Edge não gerou ' + $png) }
}

# ------------------------------------------------------------ 2-4. por tela
$problemasTotal = 0
foreach ($a in $arqTelas) {
    $def = (Ler $a.FullName) | ConvertFrom-Json
    $w = [int][math]::Ceiling([double]$def.canvas.largura * 96 / 72)
    $h = [int][math]::Ceiling([double]$def.canvas.altura * 96 / 72)
    $log.L('')
    $log.L('--- ' + $a.BaseName + '  (' + $def.canvas.largura + ' x ' + $def.canvas.altura + ' pt = ' + $w + ' x ' + $h + ' px a 96 dpi)')

    # fundo -> BMP 24 bits exatamente do tamanho da área interna
    $png = Join-Path $pastaPrevia ($a.BaseName + '-fundo.png')
    Save-Captura ($pagina + '?tela=' + $a.BaseName + '&modo=fundo&captura=1') $w $h 1 $png
    $img = [System.Drawing.Image]::FromFile($png)
    try {
        $bmp = New-Object System.Drawing.Bitmap($w, $h, [System.Drawing.Imaging.PixelFormat]::Format24bppRgb)
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.DrawImage($img, 0, 0, $w, $h); $g.Dispose()
        # guardado como PNG sem perda (bem menor para versionar); o build converte
        # para BMP na montagem, porque o Picture do MSForms não aceita PNG
        $destBmp = Join-Path $pastaGerado ($a.BaseName + '-fundo.png')
        $bmp.Save($destBmp, [System.Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose()
    } finally { $img.Dispose() }
    $log.L('  fundo: ' + $destBmp + '  (' + [math]::Round((Get-Item -LiteralPath $destBmp).Length / 1KB) + ' KB, 24 bits, ' + $w + ' x ' + $h + ' px)')

    # prévias completas nas escalas do Windows
    foreach ($e in @(1.0, 1.25, 1.5)) {
        $saidaPng = Join-Path $pastaPrevia ($a.BaseName + '-' + [int]($e * 100) + '.png')
        Save-Captura ($pagina + '?tela=' + $a.BaseName + '&modo=completo&captura=1') $w $h $e $saidaPng
    }
    Save-Captura ($pagina + '?tela=' + $a.BaseName + '&modo=completo&captura=1&contornos=1') $w $h 1 (Join-Path $pastaPrevia ($a.BaseName + '-contornos.png'))
    $log.L('  prévias: ' + $a.BaseName + '-100/125/150.png e -contornos.png em execucao\previa')

    # conferência: a prévia escreve o relatório num <pre>; --dump-dom devolve o HTML final
    $dom = Invoke-Edge @(($pagina + '?tela=' + $a.BaseName + '&modo=completo'), '--dump-dom') -Saida
    $m = [regex]::Match($dom, '<pre id="relatorio">(.*?)</pre>', 'Singleline')
    if (-not $m.Success -or -not $m.Groups[1].Value.Trim()) { throw ('a prévia de ' + $a.BaseName + ' não produziu relatório (erro de JavaScript?)') }
    $rel = [System.Web.HttpUtility]::HtmlDecode($m.Groups[1].Value) | ConvertFrom-Json
    if (@($rel.problemas).Count -eq 0) { $log.L('  conferência: ' + $rel.itens + ' elementos, sem problemas') }
    else {
        foreach ($p in $rel.problemas) { $log.L('  PROBLEMA: ' + $p.msg) }
        $problemasTotal += @($rel.problemas).Count
    }
}

# ------------------------------------------------------------ 5. modTema.bas (VBA)
function NomeConst([string]$n) { return ([regex]::Replace($n, '([a-z0-9])([A-Z])', '$1_$2')).ToUpperInvariant() }
function CorLong([string]$hex) {
    $hex = $hex.TrimStart('#')
    return ([Convert]::ToInt32($hex.Substring(0, 2), 16) + [Convert]::ToInt32($hex.Substring(2, 2), 16) * 256 + [Convert]::ToInt32($hex.Substring(4, 2), 16) * 65536)
}
$tema = (Ler (Join-Path $layout 'tema.json')) | ConvertFrom-Json
$icones = (Ler (Join-Path $layout 'icones.json')) | ConvertFrom-Json
$L = New-Object System.Collections.ArrayList
function A([string]$t) { [void]$L.Add($t) }
A 'Attribute VB_Name = "modTema"'
A "'=========================================================="
A "' modTema - GERADO por build\gerar-visual.ps1 a partir de"
A "'           fonte\layout\tema.json e icones.json."
A "' NAO EDITE A MAO: altere o JSON e rode o gerador."
A "'=========================================================="
A 'Option Explicit'
A ''
foreach ($p in $tema.fontes.PSObject.Properties) { A ('Public Const FONTE_' + (NomeConst $p.Name) + ' As String = "' + $p.Value + '"') }
A ''
foreach ($p in $tema.cores.PSObject.Properties) { A ('Public Const COR_' + (NomeConst $p.Name) + ' As Long = ' + (CorLong $p.Value) + '   '' ' + $p.Value) }
A ''
foreach ($p in $icones.PSObject.Properties) {
    if ($p.Name -like '_*') { continue }
    A ('Public Const ICO_' + (NomeConst $p.Name) + ' As Long = &H' + $p.Value[0] + '&   '' ' + $p.Value[1])
}
A ''
A "' Cor de uma parte (fundo, texto, hover, borda) de uma variante de botao. -1 = nao se aplica."
A 'Public Function CorBotao(ByVal variante As String, ByVal parte As String) As Long'
A '    CorBotao = -1'
A '    Select Case LCase$(variante) & "|" & LCase$(parte)'
foreach ($v in $tema.botoes.PSObject.Properties) {
    if ($v.Name -like '_*') { continue }
    foreach ($c in $v.Value.PSObject.Properties) { A ('        Case "' + $v.Name.ToLower() + '|' + $c.Name.ToLower() + '": CorBotao = ' + (CorLong $c.Value)) }
}
A '    End Select'
A 'End Function'
A ''
A "' Cor do selo de uma situacao (parte: texto ou fundo). Situacao desconhecida: azul."
A 'Public Function CorSituacao(ByVal situacao As String, ByVal parte As String) As Long'
A '    Select Case situacao & "|" & LCase$(parte)'
foreach ($s in $tema.situacoes.PSObject.Properties) {
    if ($s.Name -like '_*') { continue }
    foreach ($c in $s.Value.PSObject.Properties) { A ('        Case "' + $s.Name + '|' + $c.Name.ToLower() + '": CorSituacao = ' + (CorLong $c.Value)) }
}
A '        Case Else: If LCase$(parte) = "fundo" Then CorSituacao = COR_AZUL_CLARO Else CorSituacao = COR_AZUL'
A '    End Select'
A 'End Function'
A ''
A 'Public Function Glifo(ByVal codigo As Long) As String'
A '    Glifo = ChrW$(codigo)'
A 'End Function'
$cp1252 = [System.Text.Encoding]::GetEncoding(1252)
[System.IO.File]::WriteAllText((Join-Path $fonte 'modulos\modTema.bas'), (($L -join "`r`n") + "`r`n"), $cp1252)
$log.L('')
$log.L('modTema.bas gerado (' + $L.Count + ' linhas)')

# ------------------------------------------------------------ 6. manifesto
$entradas = [ordered]@{}
foreach ($f in @(Get-ChildItem -LiteralPath $layout -Recurse -File -Filter '*.json') + @(Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot 'visual') -File)) {
    $entradas[$f.FullName.Substring($dados.Length + 1)] = (Get-FileHash -LiteralPath $f.FullName -Algorithm SHA256).Hash
}
[System.IO.File]::WriteAllText((Join-Path $pastaGerado 'manifesto.json'),
    (ConvertTo-Json -InputObject ([ordered]@{ gerado_em = (Get-Date -Format 'yyyy-MM-dd HH:mm'); entradas = $entradas }) -Depth 4), $utf8)

$log.L('')
if ($problemasTotal -gt 0) {
    $log.L('RESULTADO: ' + $problemasTotal + ' problema(s) de layout (veja acima e execucao\previa\*-contornos.png).')
    $log.Salvar()
    if (-not $Tolerante) { exit 1 }
    exit 0
}
$log.L('RESULTADO: visual gerado, sem problemas de layout.')
$log.Salvar()
exit 0
