<#
  Excel.ps1 - automação do Excel para o build (COM, late binding).

  - Tema: valores "@nome" nos layouts viram cor (R + G*256 + B*65536) ou fonte.
  - Formulários: criados a partir de fonte\layout\formularios\*.json.
  - Vigia: processo separado que MATA o Excel do build se ele travar (ex.: caixa
    de erro de compilação invisível). Sem isso, um erro de sintaxe deixaria o
    build parado para sempre numa estação sem ninguém olhando.
#>

Add-Type -Namespace CrmBuild -Name Janela -MemberDefinition @'
[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint processId);
'@ -ErrorAction SilentlyContinue

# ------------------------------------------------------------------ tema
function Read-Json([string]$arquivo) {
    return ([System.IO.File]::ReadAllText($arquivo, [System.Text.Encoding]::UTF8) | ConvertFrom-Json)
}

function ConvertTo-CorExcel([string]$hex) {
    if ($hex -notmatch '^#([0-9A-Fa-f]{2})([0-9A-Fa-f]{2})([0-9A-Fa-f]{2})$') { throw ('cor inválida: ' + $hex) }
    return ([Convert]::ToInt32($Matches[1], 16) + [Convert]::ToInt32($Matches[2], 16) * 256 + [Convert]::ToInt32($Matches[3], 16) * 65536)
}

# "@azul" -> número de cor; "@padrao" -> nome de fonte; o resto passa como veio.
function Resolve-ValorTema($tema, $valor) {
    if ($valor -isnot [string] -or -not $valor.StartsWith('@')) { return $valor }
    $nome = $valor.Substring(1)
    $c = $tema.cores.PSObject.Properties[$nome]
    if ($c) { return (ConvertTo-CorExcel $c.Value) }
    $f = $tema.fontes.PSObject.Properties[$nome]
    if ($f) { return $f.Value }
    throw ('valor de tema desconhecido: ' + $valor)
}

# ------------------------------------------------------------------ processo
# Excel por ligação TARDIA. Com a biblioteca de tipos do Office registrada, o
# New-Object -ComObject devolve a classe de interoperabilidade e o PowerShell
# às vezes falha nela ("Elemento não encontrado", Hwnd nulo). O objeto COM puro
# funciona sempre. Devolve o objeto; o processo vem em $idProcesso.
function New-Excel([ref]$idProcesso) {
    $antes = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    $xl = [Activator]::CreateInstance([Type]::GetTypeFromProgID('Excel.Application'))
    $idProcesso.Value = Get-PidExcel $xl $antes
    return $xl
}

# Processo do Excel: pela janela e, se ela não responder, pelo processo que
# nasceu agora (o Excel pode ainda não ter janela ou a chamada falhar).
function Get-PidExcel($xl, $idsAntes = @()) {
    $id = 0
    try {
        $pid2 = [uint32]0
        [void][CrmBuild.Janela]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd, [ref]$pid2)
        $id = [int]$pid2
    } catch { $id = 0 }
    if ($id -le 0) {
        $novos = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $idsAntes -notcontains $_.Id })
        if ($novos.Count -gt 0) { $id = @($novos | Sort-Object StartTime)[-1].Id }
    }
    return [int]$id
}

# Mata o processo $alvo se o arquivo-sinal não aparecer em $segundos.
function Start-Vigia([int]$alvo, [int]$segundos, [string]$sinal) {
    $cmd = '$fim=(Get-Date).AddSeconds(' + $segundos + '); while((Get-Date) -lt $fim){ if(Test-Path -LiteralPath ''' +
           $sinal.Replace("'", "''") + '''){ exit 0 }; Start-Sleep -Milliseconds 500 }; Stop-Process -Id ' + $alvo + ' -Force -ErrorAction SilentlyContinue'
    $ps = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    return (Start-Process -FilePath $ps -ArgumentList @('-NoProfile', '-NonInteractive', '-WindowStyle', 'Hidden', '-Command', $cmd) -PassThru -WindowStyle Hidden)
}

function Stop-Vigia($vigia, [string]$sinal) {
    Set-Content -LiteralPath $sinal -Value 'ok'
    if ($vigia -and -not $vigia.HasExited) { [void]$vigia.WaitForExit(3000); if (-not $vigia.HasExited) { $vigia.Kill() } }
    Remove-Item -LiteralPath $sinal -Force -ErrorAction SilentlyContinue
}

# Encerra o Excel do build de qualquer jeito: Quit, solta o COM, e se o
# processo continuar vivo, mata. Excel órfão segura o arquivo e trava o próximo build.
function Close-Excel($xl, [int]$idProcesso) {
    if ($xl) {
        try { $xl.DisplayAlerts = $false } catch { }
        try { foreach ($w in @($xl.Workbooks)) { $w.Close($false) } } catch { }
        try { $xl.Quit() } catch { }
        try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl) } catch { }
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
    if ($idProcesso -gt 0) {
        $p = Get-Process -Id $idProcesso -ErrorAction SilentlyContinue
        if ($p -and -not $p.WaitForExit(8000)) { Stop-Process -Id $idProcesso -Force -ErrorAction SilentlyContinue }
    }
}

# ------------------------------------------------------------------ acesso ao projeto VBA
# "Confiar no acesso ao modelo de objeto do projeto do VBA" (HKCU, sem administrador).
# O build liga só durante a montagem e devolve o valor original ao terminar.
function Get-AcessoVBOM {
    $k = Get-ItemProperty 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security' -ErrorAction SilentlyContinue
    if ($k -and ($k.PSObject.Properties.Name -contains 'AccessVBOM')) { return [int]$k.AccessVBOM }
    return $null
}

function Set-AcessoVBOM($valor) {
    $chave = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
    if (-not (Test-Path $chave)) { New-Item -Path $chave -Force | Out-Null }
    if ($null -eq $valor) { Remove-ItemProperty -Path $chave -Name AccessVBOM -ErrorAction SilentlyContinue }
    else { Set-ItemProperty -Path $chave -Name AccessVBOM -Value ([int]$valor) -Type DWord }
}

# ------------------------------------------------------------------ controles
# Aplica "Font.Size" = 9, "BackColor" = "@azul" etc. num controle MSForms.
function Set-PropControle($ctl, [string]$nome, $valor, $tema) {
    $v = Resolve-ValorTema $tema $valor
    if ($nome.StartsWith('Font.')) { Set-FonteControle $ctl $nome.Substring(5) $v }
    else { $ctl.$nome = $v }
}

# Fonte de um controle MSForms no Designer, pelo PowerShell (conferido no Excel 2019):
#  - $ctl.Font.Name = ... falha ("referência de objeto não definida");
#  - FontName/FontSize/FontBold/FontItalic funcionam na maioria dos controles;
#  - Size é Currency (referencias/01-vba-msforms/font-size-bold.md): Int32 passa,
#    Double/Decimal não; fração vai em CurrencyWrapper;
#  - o Frame não tem FontName: vai pelo objeto Font via InvokeMember (Size em Decimal).
function ConvertTo-TamanhoFonte($v) {
    $d = [decimal]$v
    if ($d -eq [math]::Truncate($d)) { return [int]$d }
    return (New-Object Runtime.InteropServices.CurrencyWrapper($d))
}

function Set-FonteControle($ctl, [string]$parte, $valor) {
    $v = $valor
    if ($parte -eq 'Size') { $v = ConvertTo-TamanhoFonte $valor }
    try {
        $ctl.('Font' + $parte) = $v
        return
    } catch {
        if ($_.Exception.Message -notmatch 'Font') { throw }
    }
    $B = [Reflection.BindingFlags]
    $f = [System.__ComObject].InvokeMember('Font', $B::GetProperty, $null, $ctl, $null)
    if ($parte -eq 'Size') { $valor = [decimal]$valor }
    [void][System.__ComObject].InvokeMember($parte, $B::SetProperty, $null, $f, @($valor))
}

function Get-EstilosControle($def, $ctlDef) {
    $saida = [ordered]@{}
    $lista = @()
    if ($ctlDef.PSObject.Properties['estilo']) { $lista = @($ctlDef.estilo) }
    foreach ($e in $lista) {
        $est = $def.estilos.PSObject.Properties[$e]
        if (-not $est) { throw ('estilo "' + $e + '" não existe em ' + $def.nome) }
        foreach ($p in $est.Value.PSObject.Properties) { $saida[$p.Name] = $p.Value }
    }
    if ($ctlDef.PSObject.Properties['props']) { foreach ($p in $ctlDef.props.PSObject.Properties) { $saida[$p.Name] = $p.Value } }
    return $saida
}

# Propriedade do UserForm pela coleção Properties do VBE. O PowerShell não acerta o
# tipo do Value (Int32 e Double recusados com "não é possível converter em String",
# conforme a ordem das chamadas); texto funciona sempre e o VBE converte. Número vai
# em formato invariante (ponto decimal). Conferido no Excel 2019.
function Set-PropriedadeVBE($comp, [string]$nome, $valor) {
    if ($valor -is [double] -or $valor -is [int] -or $valor -is [long] -or $valor -is [single] -or $valor -is [decimal]) {
        $texto = ([double]$valor).ToString([Globalization.CultureInfo]::InvariantCulture)
    } else { $texto = [string]$valor }
    $comp.Properties.Item($nome).Value = $texto
}

# Cria o UserForm a partir do JSON de layout e cola o código (Windows-1252).
function Add-FormularioDeLayout($vbProj, [string]$arqLayout, [string]$pastaFonte, $tema, $log) {
    $def = Read-Json $arqLayout
    $comp = $vbProj.VBComponents.Add(3)                 # vbext_ct_MSForm
    $comp.Name = $def.nome
    foreach ($p in $def.propriedades.PSObject.Properties) {
        Set-PropriedadeVBE $comp $p.Name (Resolve-ValorTema $tema $p.Value)
    }
    $d = $comp.Designer
    $falhas = 0
    foreach ($c in $def.controles) {
        try {
            $ctl = $d.Controls.Add('Forms.' + $c.tipo + '.1', $c.nome, $true)
            $ctl.Left = [double]$c.pos[0]; $ctl.Top = [double]$c.pos[1]
            $ctl.Width = [double]$c.pos[2]; $ctl.Height = [double]$c.pos[3]
            if ($c.PSObject.Properties['texto']) { $ctl.Caption = [string]$c.texto }
            $props = Get-EstilosControle $def $c
            foreach ($k in $props.Keys) {
                try { Set-PropControle $ctl $k $props[$k] $tema }
                catch { $falhas++; $log.L('    AVISO ' + $c.nome + '.' + $k + ': ' + $_.Exception.Message.Split([char]13)[0]) }
            }
        } catch {
            $falhas++
            $log.L('    FALHA ' + $c.nome + ': ' + $_.Exception.Message.Split([char]13)[0])
        }
    }
    if ($def.PSObject.Properties['canvas']) {
        # layout v2: canvas = área interna; a moldura (bordas + título) vem do tema
        Set-PropriedadeVBE $comp 'Width' ([double]$def.canvas.largura + [double]$tema.moldura.largura)
        Set-PropriedadeVBE $comp 'Height' ([double]$def.canvas.altura + [double]$tema.moldura.altura)
    }
    if ($def.PSObject.Properties['componentes']) {
        foreach ($k in $def.componentes) {
            try { Add-Componente $d $k $tema }
            catch { $falhas++; $log.L('    FALHA componente ' + $k.nome + ': ' + $_.Exception.Message.Split([char]13)[0]) }
        }
    }
    $codigo = [System.IO.File]::ReadAllText((Join-Path $pastaFonte $def.codigo), [System.Text.Encoding]::GetEncoding(1252))
    $comp.CodeModule.AddFromString($codigo)
    $log.L('  ' + $def.nome.PadRight(10) + @($def.controles).Count + ' controles, ' + $comp.CodeModule.CountOfLines + ' linhas de código' +
           $(if ($falhas) { ', ' + $falhas + ' aviso(s)' } else { '' }))
    return $falhas
}

# ------------------------------------------------------------------ componentes (layout v2)
# Contrato com clsUI/clsBotao/clsGrade (fonte\classes): nomes bg_/ico_ e Tag "ui:...".
function New-Rotulo($d, [string]$nome, [double[]]$pos, [string]$texto) {
    $c = $d.Controls.Add('Forms.Label.1', $nome, $true)
    $c.Left = $pos[0]; $c.Top = $pos[1]; $c.Width = $pos[2]; $c.Height = $pos[3]
    $c.Caption = $texto
    $c.BackStyle = 0; $c.BorderStyle = 0; $c.WordWrap = $false
    return $c
}

function Get-GlifoIcone([string]$nomeIcone) {
    $ic = $script:IconesLayout.PSObject.Properties[$nomeIcone]
    if (-not $ic) { throw ('ícone "' + $nomeIcone + '" não existe em icones.json') }
    return [string][char][Convert]::ToInt32([string]$ic.Value[0], 16)
}

# Label de uma linha centralizado na vertical dentro de pos (MSForms não centraliza).
function Get-PosTexto([double[]]$pos, [double]$tamanho, [double]$esq, [double]$dir) {
    $alt = [Math]::Ceiling($tamanho * 1.45)
    return @(($pos[0] + $esq), ($pos[1] + [Math]::Floor(($pos[3] - $alt) / 2)), ($pos[2] - $esq - $dir), $alt)
}

function Add-Componente($d, $k, $tema) {
    $pos = @([double]$k.pos[0], [double]$k.pos[1], [double]$k.pos[2], [double]$k.pos[3])
    $txt = if ($k.PSObject.Properties['texto']) { [string]$k.texto } else { '' }
    switch ([string]$k.tipo) {
        'botao' {
            $var = [string]$k.variante
            $cores = $tema.botoes.PSObject.Properties[$var]
            if (-not $cores) { throw ('variante "' + $var + '" não existe em tema.botoes') }
            $bg = New-Rotulo $d ('bg_' + $k.nome) $pos ''
            $bg.BackStyle = 1; $bg.BackColor = ConvertTo-CorExcel $cores.Value.fundo
            if ($cores.Value.PSObject.Properties['borda']) { $bg.BorderStyle = 1; $bg.BorderColor = ConvertTo-CorExcel $cores.Value.borda }
            $esq = 0
            if ($k.PSObject.Properties['icone']) {
                $ico = New-Rotulo $d ('ico_' + $k.nome) (Get-PosTexto $pos 11 8 0) (Get-GlifoIcone $k.icone)
                $ico.Width = 16
                Set-FonteControle $ico 'Name' (Resolve-ValorTema $tema '@icones'); Set-FonteControle $ico 'Size' 11
                $ico.ForeColor = ConvertTo-CorExcel $cores.Value.texto; $ico.TextAlign = 2
                # alguns glifos (E76C, E77B, E821...) saem corrompidos do formulário salvo:
                # o código vai na Tag e o clsUI regrava o Caption ao abrir a tela
                $ico.Tag = 'ui:glifo:' + [string]$script:IconesLayout.PSObject.Properties[[string]$k.icone].Value[0]
                $esq = 18
            }
            $t = New-Rotulo $d $k.nome (Get-PosTexto $pos 9 ($esq + 4) 4) $txt
            Set-FonteControle $t 'Name' (Resolve-ValorTema $tema '@semibold'); Set-FonteControle $t 'Size' 9
            $t.ForeColor = ConvertTo-CorExcel $cores.Value.texto; $t.TextAlign = 2
            $t.Tag = 'ui:botao:' + $var
            if ($k.PSObject.Properties['desabilitado'] -and $k.desabilitado) { $t.Enabled = $false }
        }
        'icone' {
            $tam = if ($k.PSObject.Properties['tamanho']) { [double]$k.tamanho } else { 12 }
            $c = New-Rotulo $d $k.nome (Get-PosTexto $pos $tam 0 0) (Get-GlifoIcone $k.icone)
            Set-FonteControle $c 'Name' (Resolve-ValorTema $tema '@icones'); Set-FonteControle $c 'Size' $tam; $c.TextAlign = 2
            $c.Tag = 'ui:glifo:' + [string]$script:IconesLayout.PSObject.Properties[[string]$k.icone].Value[0]
            $c.ForeColor = Resolve-ValorTema $tema $(if ($k.PSObject.Properties['cor']) { $k.cor } else { '@texto' })
        }
        'selo' {
            $sit = $tema.situacoes.PSObject.Properties[[string]$k.situacao]
            if (-not $sit) { throw ('situação "' + $k.situacao + '" não existe em tema.situacoes') }
            $bg = New-Rotulo $d ('bg_' + $k.nome) $pos ''
            $bg.BackStyle = 1; $bg.BackColor = ConvertTo-CorExcel $sit.Value.fundo
            $t = New-Rotulo $d $k.nome (Get-PosTexto $pos 8 2 2) $txt
            Set-FonteControle $t 'Name' (Resolve-ValorTema $tema '@semibold'); Set-FonteControle $t 'Size' 8; $t.TextAlign = 2
            $t.ForeColor = ConvertTo-CorExcel $sit.Value.texto
        }
        'grade' {
            $f = $d.Controls.Add('Forms.Frame.1', $k.nome, $true)
            $f.Left = $pos[0]; $f.Top = $pos[1]; $f.Width = $pos[2]; $f.Height = $pos[3]
            $f.Caption = ''; $f.SpecialEffect = 0; $f.BorderStyle = 1
            $f.BorderColor = Resolve-ValorTema $tema '@borda'; $f.BackColor = Resolve-ValorTema $tema '@branco'
            $f.Tag = 'ui:grade'
        }
        default { throw ('tipo de componente desconhecido: ' + $k.tipo) }
    }
}

# Fundo gerado (PNG da prévia) -> BMP 24 bits a 96 dpi, que o Picture do MSForms aceita.
function ConvertTo-BmpFundo([string]$png, [string]$bmp) {
    Add-Type -AssemblyName System.Drawing
    $img = [System.Drawing.Image]::FromFile($png)
    try {
        $b = New-Object System.Drawing.Bitmap($img.Width, $img.Height, [System.Drawing.Imaging.PixelFormat]::Format24bppRgb)
        try {
            $b.SetResolution(96, 96)
            $g = [System.Drawing.Graphics]::FromImage($b)
            try { $g.Clear([System.Drawing.Color]::White); $g.DrawImage($img, 0, 0, $img.Width, $img.Height) } finally { $g.Dispose() }
            $b.Save($bmp, [System.Drawing.Imaging.ImageFormat]::Bmp)
        } finally { $b.Dispose() }
    } finally { $img.Dispose() }
}

# Embute os fundos: o Designer.Picture só recebe IPictureDisp, que o PowerShell não
# cria; um módulo VBA temporário faz LoadPicture em tempo de desenho e é removido.
# O .xlsm final não lê nada do disco (o arquivo roda de Documentos).
function Set-FundosFormularios($xl, $wb, $vbProj, $fundos, $log) {
    if (@($fundos.Keys).Count -eq 0) { return }
    $m = $vbProj.VBComponents.Add(1)
    $m.Name = 'modBuildVisual'
    $m.CodeModule.AddFromString(@'
Public Function AplicarFundo(ByVal nome As String, ByVal arq As String) As String
    On Error GoTo falha
    With ThisWorkbook.VBProject.VBComponents(nome).Designer
        Set .Picture = LoadPicture(arq)
        .PictureAlignment = 0
        .PictureSizeMode = 0
        .PictureTiling = False
    End With
    AplicarFundo = "OK"
    Exit Function
falha:
    AplicarFundo = Err.Description
End Function
'@)
    try {
        foreach ($nome in @($fundos.Keys)) {
            $r = [string]$xl.Run("'" + $wb.Name + "'!modBuildVisual.AplicarFundo", $nome, $fundos[$nome])
            if ($r -ne 'OK') { throw ('fundo de ' + $nome + ': ' + $r) }
            $log.L('  ' + $nome.PadRight(14) + 'fundo embutido')
        }
    } finally { $vbProj.VBComponents.Remove($m) }
}

# Botão = forma arredondada com a cor da marca e OnAction (vale como controle de formulário).
function Add-BotaoMarca($ws, [string]$texto, [string]$macro, [double[]]$pos, [string]$nome, [int]$cor, $tema) {
    $sh = $ws.Shapes.AddShape(5, $pos[0], $pos[1], $pos[2], $pos[3])   # msoShapeRoundedRectangle
    try { $sh.Adjustments.Item(1) = 0.16 } catch { }
    $sh.Fill.Visible = $true
    $sh.Fill.ForeColor.RGB = $cor
    $sh.Line.Visible = $false
    try { $sh.Shadow.Visible = $false } catch { }
    try {
        $tr = $sh.TextFrame2.TextRange
        $tr.Text = $texto
        $tr.Font.Name = (Resolve-ValorTema $tema '@padrao')
        $tr.Font.Size = 10
        $tr.Font.Bold = $true
        $tr.Font.Fill.ForeColor.RGB = (Resolve-ValorTema $tema '@branco')
        $tr.ParagraphFormat.Alignment = 2                 # msoAlignCenter
        $sh.TextFrame2.VerticalAnchor = 3                 # msoAnchorMiddle
        $sh.TextFrame2.WordWrap = $false
        $sh.TextFrame2.MarginLeft = 2; $sh.TextFrame2.MarginRight = 2
    } catch { try { $sh.TextFrame.Characters().Text = $texto } catch { } }
    $sh.OnAction = $macro
    try { $sh.Placement = 3 } catch { }                  # xlFreeFloating: não estica com a coluna
    if ($nome.Length -gt 31) { $nome = $nome.Substring(0, 31) }   # limite do Excel
    $sh.Name = $nome
    return $sh
}
