<#
  montar-frontend.ps1 - monta o CRM_Zapromaq.xlsm do zero a partir de fonte\ e publica na raiz de BASE.

  É o ÚNICO comando da publicação: faz a sequência inteira, na ordem, e cada
  etapa só começa se a anterior deu certo.

    1. pré-checagens: trava contra dois builds, .xlsm da raiz livre, Excel, banco
    2. BACKUP do banco e do .xlsm atuais (execucao\backup\antes-da-publicacao)
    3. MIGRAÇÕES do banco que estiverem pendentes (banco\migracoes)
    4. versão: VERSAO.txt 1.4 -> 1.5 (só gravada no fim, se tudo der certo)
    5. montagem num arquivo TEMPORÁRIO (execucao\build): abas, módulos, formulários,
       código, botões, Config!B2 com o caminho UNC do banco, Início!B7
    6. compilação do projeto VBA e autoteste (modAutoteste) com VIGIA: se o Excel
       travar numa caixa de erro invisível, o vigia o encerra e o build falha
    7. conferência do pacote pelo zip (sem Excel): projeto VBA, abas, Início!B7
    8. TESTE DE USO no arquivo montado: monta as três grades e o Painel com os
       dados reais (testes\testar-uso.ps1)
    9. publicação: .xlsm anterior -> execucao\historico-front; novo -> raiz;
       config.versao_front no banco; VERSAO.txt; VERSAO-FRONT.txt (raiz, para o
       INICIAR-CRM comparar sem abrir o Excel); CHANGELOG.md
  Build com falha NUNCA mexe na raiz: o .xlsm publicado continua o de antes.

  -Teste          monta e confere, mas não publica nem sobe a versão (resultado em
                  execucao\teste); não faz backup nem migração
  -SemBackup      pula o backup (etapa 2)
  -SemMigracoes   pula as migrações (etapa 3)
  -SemTesteDeUso  pula o teste de uso (etapa 8)
  -Ambiente PRODUCAO | TESTE (padrão: config.ambiente do banco; sem banco, PRODUCAO)

  "Confiar no acesso ao modelo de objeto do projeto do VBA" é ligado só durante a
  montagem (HKCU, sem administrador) e volta ao valor original no fim.
#>
param(
    [switch]$Teste,
    [switch]$SemBackup,
    [switch]$SemMigracoes,
    [switch]$SemTesteDeUso,
    [ValidateSet('', 'PRODUCAO', 'TESTE')][string]$Ambiente = '',
    [int]$TempoLimite = 300
)
$ErrorActionPreference = 'Stop'
$dados = Split-Path -Parent $PSScriptRoot
. (Join-Path $dados 'lib\Comum.ps1')
. (Join-Path $PSScriptRoot 'lib\Versao.ps1')
. (Join-Path $PSScriptRoot 'lib\Pacote.ps1')
. (Join-Path $PSScriptRoot 'lib\Excel.ps1')

$raiz = Get-RaizBase $PSScriptRoot
$fonte = Join-Path $dados 'fonte'
$tema = Read-Json (Join-Path $fonte 'layout\tema.json')
$abas = Read-Json (Join-Path $fonte 'layout\abas.json')
$script:IconesLayout = Read-Json (Join-Path $fonte 'layout\icones.json')
$log = New-Log $raiz 'build'
$pastaBuild = Get-PastaExecucao $raiz 'build'
$destinoRaiz = Get-CaminhoFront $raiz
$bancoArq = Get-CaminhoBanco $raiz

$xl = $null; $idExcel = 0; $trava = $null; $vbomOriginal = 'nao-alterado'; $sucesso = $false
function Parar([string]$msg) { throw $msg }

try {
    $log.L('=== MONTAGEM DO FRONTEND - CRM Zapromaq ===')
    $log.L('Data....: ' + (Get-Date -Format 'dd/MM/yyyy HH:mm:ss') + '   ' + $env:COMPUTERNAME + '\' + $env:USERNAME)
    $log.L('Raiz....: ' + (ConvertTo-CaminhoUNC $raiz))
    $log.L('Modo....: ' + $(if ($Teste) { 'TESTE (não publica)' } else { 'PUBLICAÇÃO' }))

    # ------------------------------------------------------------ 1. pré-checagens
    try { $trava = [System.IO.File]::Open((Join-Path $pastaBuild 'build.lock'), 'OpenOrCreate', 'ReadWrite', 'None') }
    catch { Parar 'Outro build está em andamento (execucao\build\build.lock em uso). Espere ele terminar.' }

    if ($null -eq [type]::GetTypeFromProgID('Excel.Application')) {
        Parar 'Excel não está instalado nesta máquina. O build roda numa estação com Excel 2019 64 bits (veja testes\CHECAR-AMBIENTE.bat).'
    }
    if (-not [Environment]::Is64BitProcess) { Parar 'Rode pelo MONTAR-FRONTEND.bat: o build precisa do PowerShell 64 bits.' }

    if (-not $Teste -and (Test-Path -LiteralPath $destinoRaiz)) {
        try { $f = [System.IO.File]::Open($destinoRaiz, 'Open', 'ReadWrite', 'None'); $f.Close() }
        catch { Parar ('O ' + $script:NomeFront + ' da raiz está aberto por alguém. Feche-o em todas as máquinas e rode de novo. Nada foi alterado.') }
    }

    $temBanco = $false
    if (Test-Path -LiteralPath $bancoArq) {
        try {
            $cn = Open-Banco $bancoArq
            try {
                $temBanco = $true
                if (-not $Ambiente) { $Ambiente = [string](Invoke-Escalar $cn "SELECT valor FROM config WHERE chave='ambiente'" @()) }
                $log.L('Banco...: ' + $bancoArq + '  (esquema ' + (Get-VersaoEsquemaBanco $cn) + ')')
            } finally { $cn.Close(); $cn.Dispose() }
        } catch { $log.L('AVISO: banco não abriu (' + $_.Exception.Message.Split([char]13)[0] + '); a versão não será gravada nele.') }
    } else { $log.L('AVISO: banco não encontrado em ' + $bancoArq + '; a versão não será gravada nele.') }
    if ($Ambiente -notin @('PRODUCAO', 'TESTE')) { $Ambiente = 'PRODUCAO' }

    # ------------------------------------------------------------ 2. backup
    if (-not $Teste -and -not $SemBackup) {
        $log.L('')
        $log.L('--- backup do que está publicado hoje')
        $pastaBkp = New-BackupAntesDePublicar $raiz $bancoArq $destinoRaiz $log
        $log.L('  em ......: ' + $pastaBkp)
    }

    # ------------------------------------------------------------ 3. migrações do banco
    if (-not $Teste -and -not $SemMigracoes -and $temBanco) {
        $log.L('')
        $log.L('--- migrações do banco')
        $saidaMig = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $dados 'banco\aplicar-migracoes.ps1') 2>&1
        foreach ($linha in @($saidaMig)) { if ([string]$linha -ne '') { $log.L('  ' + [string]$linha) } }
        if ($LASTEXITCODE -ne 0) { Parar 'as migrações falharam (veja acima). Nada foi publicado; o backup da etapa 2 está intacto.' }
    }

    # parte visual gerada (fundos, modTema) em dia com fonte\layout? senão, gera agora
    $velhos = @(Get-VisualDesatualizado $dados)
    if ($velhos.Count -gt 0) {
        $log.L('parte visual desatualizada (' + ($velhos -join ', ') + '): rodando gerar-visual')
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'gerar-visual.ps1') | Out-Null
        if ($LASTEXITCODE -ne 0) { Parar 'gerar-visual falhou (veja execucao\logs\visual-*.log). Corrija o layout e rode de novo.' }
        if (@(Get-VisualDesatualizado $dados).Count -gt 0) { Parar 'gerar-visual terminou mas o manifesto continua desatualizado.' }
    }

    # ------------------------------------------------------------ 4. versão
    $versaoAtual = Get-VersaoAtual $dados
    $versao = Get-ProximaVersao $versaoAtual
    $quando = Get-Date
    $textoB7 = Format-TextoB7 $Ambiente $versao $quando
    $log.L('Versão..: ' + $versaoAtual + ' -> ' + $versao + '   ambiente ' + $Ambiente)
    $log.L('')

    # fonte de montagem: cópia com VERSAO_FRONT atualizada (a fonte original não é tocada)
    $palco = Join-Path $pastaBuild 'fonte'
    if (Test-Path -LiteralPath $palco) { Remove-Item -LiteralPath $palco -Recurse -Force }
    Copy-Item -LiteralPath $fonte -Destination $palco -Recurse
    $cp1252 = [System.Text.Encoding]::GetEncoding(1252)
    $arqCfg = Join-Path $palco 'modulos\modConfig.bas'
    [System.IO.File]::WriteAllText($arqCfg, (Set-VersaoNoModConfig ([System.IO.File]::ReadAllText($arqCfg, $cp1252)) $versao), $cp1252)

    # ------------------------------------------------------------ 5. montagem
    $vbomOriginal = Get-AcessoVBOM
    if ($vbomOriginal -ne 1) { Set-AcessoVBOM 1; $log.L('acesso ao modelo de objeto do VBA ligado para a montagem (volta ao original no fim)') }

    $xl = New-Excel ([ref]$idExcel)
    $xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false; $xl.EnableEvents = $false
    $log.L('Excel ' + $xl.Version + ' (build ' + $xl.Build + ', processo ' + $idExcel + ')')
    $wb = $xl.Workbooks.Add()
    while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }

    $vbProj = $null
    try { $vbProj = $wb.VBProject; [void]$vbProj.VBComponents.Count } catch { $vbProj = $null }
    if ($null -eq $vbProj) {
        Parar ('O Excel bloqueou o acesso ao projeto VBA mesmo com a opção ligada (política da empresa?). ' +
               'Arquivo > Opções > Central de Confiabilidade > Configurações de Macro > [x] Confiar no acesso ao modelo de objeto do projeto do VBA.')
    }

    $log.L('--- abas')
    $wb.Worksheets.Item(1).Name = $abas.abas[0].nome
    for ($i = 1; $i -lt @($abas.abas).Count; $i++) {
        $nova = $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count), 1, [Type]::Missing)
        $nova.Name = $abas.abas[$i].nome
    }
    $log.L('  ' + (@($abas.abas | ForEach-Object { $_.nome }) -join ', '))

    $log.L('--- módulos')
    foreach ($arq in @(Get-ChildItem -LiteralPath (Join-Path $palco 'modulos') -Filter '*.bas' | Sort-Object Name)) {
        [void]$vbProj.VBComponents.Import($arq.FullName)
        $log.L('  ' + $arq.BaseName)
    }
    foreach ($arq in @(Get-ChildItem -LiteralPath (Join-Path $palco 'classes') -Filter '*.cls' -ErrorAction SilentlyContinue | Sort-Object Name)) {
        [void]$vbProj.VBComponents.Import($arq.FullName)
        $log.L('  ' + $arq.BaseName + ' (classe)')
    }

    $log.L('--- formulários')
    $avisos = 0
    $fundos = [ordered]@{}
    foreach ($arq in @(Get-ChildItem -LiteralPath (Join-Path $palco 'layout\formularios') -Filter '*.json' | Sort-Object Name)) {
        $avisos += Add-FormularioDeLayout $vbProj $arq.FullName $palco $tema $log
        if ((Read-Json $arq.FullName).PSObject.Properties['fundo']) {
            $png = Join-Path $palco ('assets\gerado\' + $arq.BaseName + '-fundo.png')
            if (-not (Test-Path -LiteralPath $png)) { Parar ('fundo não gerado: ' + $png + '. Rode build\GERAR-VISUAL.bat.') }
            $bmp = Join-Path $pastaBuild ($arq.BaseName + '-fundo.bmp')
            ConvertTo-BmpFundo $png $bmp
            $fundos[$arq.BaseName] = $bmp
        }
    }
    Set-FundosFormularios $xl $wb $vbProj $fundos $log

    $log.L('--- código da pasta de trabalho')
    $codTW = [System.IO.File]::ReadAllText((Join-Path $palco 'pasta-de-trabalho\ThisWorkbook.txt'), $cp1252)
    $codTW = (($codTW -split "`r`n") | Where-Object { $_ -notmatch '^Attribute VB_Name' }) -join "`r`n"
    $compTW = $vbProj.VBComponents.Item($wb.CodeName)     # nome certo em qualquer idioma do Office
    $compTW.CodeModule.AddFromString($codTW)
    $log.L('  ' + $wb.CodeName + ': ' + $compTW.CodeModule.CountOfLines + ' linhas')

    $log.L('--- conteúdo das abas')
    $fundo = Resolve-ValorTema $tema '@fundo'
    $logoArq = Join-Path $palco $abas.inicio.logo.arquivo
    # Config: caminho UNC do banco - o .xlsm roda de Documentos, em qualquer estação
    $cfg = $wb.Worksheets.Item('Config')
    foreach ($p in $abas.config.celulas.PSObject.Properties) { $cfg.Range($p.Name).Value2 = $p.Value }
    $cfg.Range($abas.config.banco).Value2 = (ConvertTo-CaminhoUNC $bancoArq)
    $cfg.Range($abas.config.ambiente).Value2 = $Ambiente
    $cfg.Range('A1').Font.Bold = $true
    $cfg.Columns.Item(1).ColumnWidth = 22; $cfg.Columns.Item(2).ColumnWidth = 90
    $log.L('  Config!B2 = ' + $cfg.Range($abas.config.banco).Value2)

    foreach ($a in @($abas.abas | Where-Object { $_.PSObject.Properties['grade'] })) {
        $ws = $wb.Worksheets.Item($a.nome)
        $b = $abas.grades.botoes
        $x = [double]$b.pos[0]; $n = 0
        foreach ($it in $b.itens) {
            $n++
            $macro = if ($it.macro -eq '<NOVA>') { $a.novaFicha } else { $it.macro }
            [void](Add-BotaoMarca $ws $it.texto $macro @($x, [double]$b.pos[1], [double]$b.pos[2], [double]$b.pos[3]) ('btn' + $a.nome.Substring(0, 3) + $n) (Resolve-ValorTema $tema $it.cor) $tema)
            $x += [double]$b.passo
        }
        if (Test-Path -LiteralPath $logoArq) {
            $l = $abas.grades.logo.pos
            $lg = $ws.Shapes.AddPicture($logoArq, $false, $true, [double]$l[0], [double]$l[1], [double]$l[2], [double]$l[3])
            $lg.Name = 'logoAba'; try { $lg.Placement = 3 } catch { }
        }
        $ws.Columns.Item(1).ColumnWidth = 2.25
        $ws.Rows.Item(1).RowHeight = 4; $ws.Rows.Item(2).RowHeight = 26; $ws.Rows.Item(3).RowHeight = 30
        # título como FORMA, ao lado do logo (em célula ficaria embaixo dele:
        # a largura das colunas muda de aba para aba). modGrade reescreve ao montar.
        $tit = $ws.Shapes.AddTextbox(1, 96, $ws.Rows.Item(2).Top, 460, $ws.Rows.Item(2).Height)
        $tit.Name = 'tituloAba'
        $tit.Line.Visible = $false; $tit.Fill.Visible = $false
        try { $tit.Placement = 3 } catch { }
        $tf = $tit.TextFrame2
        $tf.MarginLeft = 0; $tf.MarginRight = 0; $tf.MarginTop = 0; $tf.MarginBottom = 0
        $tf.VerticalAnchor = 3; $tf.WordWrap = $false
        $tf.TextRange.Text = $a.nome
        $tf.TextRange.Font.Name = (Resolve-ValorTema $tema '@semibold')
        $tf.TextRange.Font.Size = 15
        $tf.TextRange.Font.Bold = $true
        $tf.TextRange.Font.Fill.ForeColor.RGB = (Resolve-ValorTema $tema '@azul')
        $ws.Range($abas.grades.aviso.celula).Value2 = $abas.grades.aviso.texto
        $log.L('  ' + $a.nome.PadRight(14) + @($b.itens).Count + ' botões')
    }

    # Painel: só os botões; o conteúdo é desenhado pelo modPainel ao atualizar
    if ($abas.PSObject.Properties['painel']) {
        $wsP = $wb.Worksheets.Item('Painel')
        $wsP.Cells.Interior.Color = $fundo
        $b = $abas.painel.botoes
        $x = [double]$b.pos[0]; $n = 0
        foreach ($it in $b.itens) {
            $n++
            [void](Add-BotaoMarca $wsP $it.texto $it.macro @($x, [double]$b.pos[1], [double]$b.pos[2], [double]$b.pos[3]) ('btnPai' + $n) (Resolve-ValorTema $tema $it.cor) $tema)
            $x += [double]$b.passo
        }
        $wsP.Range('B2').Value2 = 'Painel comercial'; $wsP.Range('B2').Font.Size = 16; $wsP.Range('B2').Font.Bold = $true
        $wsP.Range('B2').Font.Color = (Resolve-ValorTema $tema '@azul')
        $wsP.Range('B5').Value2 = 'Os indicadores aparecem ao abrir esta aba (ou em Atualizar indicadores).'
        $log.L('  Painel        ' + @($b.itens).Count + ' botões')
    }

    $ini = $wb.Worksheets.Item($abas.abas[0].nome)
    $ini.Range('A1:Z60').Interior.Color = $fundo
    $ini.Columns.Item(1).ColumnWidth = 2.25
    if (Test-Path -LiteralPath $logoArq) {
        $l = $abas.inicio.logo.pos
        $lg = $ini.Shapes.AddPicture($logoArq, $false, $true, [double]$l[0], [double]$l[1], [double]$l[2], [double]$l[3])
        $lg.Name = 'logoZapromaq'; try { $lg.Placement = 3 } catch { }
    } else { $log.L('  AVISO: logo não encontrado em ' + $logoArq) }
    $t = $abas.inicio.titulo
    $c = $ini.Range($t.celula); $c.Value2 = $t.texto; $c.Font.Name = (Resolve-ValorTema $tema '@padrao')
    $c.Font.Size = $t.tamanho; $c.Font.Bold = $true; $c.Font.Color = (Resolve-ValorTema $tema $t.cor)
    $vb7 = $abas.inicio.versao
    $c = $ini.Range($vb7.celula); $c.NumberFormat = '@'; $c.Value2 = $textoB7
    $c.Font.Name = (Resolve-ValorTema $tema '@padrao'); $c.Font.Size = $vb7.tamanho; $c.Font.Color = (Resolve-ValorTema $tema $vb7.cor)
    $ini.Rows.Item(7).RowHeight = 18
    $b = $abas.inicio.botoes
    $y = [double]$b.pos[1]
    foreach ($it in $b.itens) {
        [void](Add-BotaoMarca $ini $it.texto $it.macro @([double]$b.pos[0], $y, [double]$b.pos[2], [double]$b.pos[3]) ('btn' + $it.macro) (Resolve-ValorTema $tema $it.cor) $tema)
        $y += [double]$b.passo
    }
    $log.L('  Inicio!B7 = ' + $textoB7)

    foreach ($a in $abas.abas) { $wb.Worksheets.Item($a.nome).Visible = [int]$a.visivel }
    $ini.Activate(); try { $xl.ActiveWindow.DisplayGridlines = $false } catch { }; [void]$ini.Range('A1').Select()

    $temp = Join-Path $pastaBuild $script:NomeFront
    if (Test-Path -LiteralPath $temp) { Remove-Item -LiteralPath $temp -Force }
    $wb.SaveAs($temp, 52)                                  # xlOpenXMLWorkbookMacroEnabled
    $log.L('')
    $log.L('montado em ' + $temp)

    # ------------------------------------------------------------ 6. compilação e autoteste (com vigia)
    $sinal = Join-Path $pastaBuild ('vigia-' + $PID + '.ok')
    Remove-Item -LiteralPath $sinal -Force -ErrorAction SilentlyContinue
    $vigia = Start-Vigia $idExcel $TempoLimite $sinal
    try {
        $log.L('--- compilação')
        $compilar = $null
        try { $compilar = $xl.VBE.CommandBars.FindControl(1, 578) } catch { }   # Depurar > Compilar VBAProject
        if ($compilar) {
            if ($compilar.Enabled) { $compilar.Execute() }
            if ($xl.VBE.CommandBars.FindControl(1, 578).Enabled) { Parar 'a compilação não terminou (o comando Compilar continua habilitado)' }
            $log.L('  projeto VBA compilado sem erro')
        } else { $log.L('  AVISO: comando Compilar não encontrado; o autoteste compila os módulos que usa') }

        $log.L('--- autoteste')
        $r = [string]$xl.Run("'" + $wb.Name + "'!modAutoteste.Autoteste", $versao)
        foreach ($linha in ($r -split "`n")) { $log.L('  ' + $linha) }
        if (-not $r.StartsWith('OK')) { Parar 'o autoteste falhou (detalhes acima)' }
    } catch {
        $morto = -not (Get-Process -Id $idExcel -ErrorAction SilentlyContinue)
        if ($morto) {
            Parar ('o Excel travou na compilação ou no autoteste e foi encerrado pelo vigia após ' + $TempoLimite + ' s. ' +
                   'Causa provável: erro de compilação. Abra execucao\build\' + $script:NomeFront + ', Alt+F11, Depurar > Compilar VBAProject.')
        }
        throw
    } finally { Stop-Vigia $vigia $sinal }

    $wb.Save()
    Close-Excel $xl $idExcel; $xl = $null

    # ------------------------------------------------------------ 7. conferência do pacote (sem Excel)
    $log.L('--- conferência do pacote')
    $prob = Test-PacoteXlsm $temp @($abas.abas | ForEach-Object { $_.nome })
    $txt = Get-TextoCelulaXlsx $temp $abas.abas[0].nome $abas.inicio.versao.celula
    $pB7 = Test-TextoB7 $txt $versao
    if ($pB7) { [void]$prob.Add($pB7) }
    if ($prob.Count -gt 0) { Parar ('pacote reprovado: ' + ($prob -join '; ')) }
    $log.L('  projeto VBA, ' + @($abas.abas).Count + ' abas e Início!B7 conferidos no arquivo')

    # ------------------------------------------------------------ 8. teste de uso
    if (-not $SemTesteDeUso) {
        $log.L('')
        $log.L('--- teste de uso (grades e Painel com os dados reais)')
        $saidaUso = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $dados 'testes\testar-uso.ps1') -Arquivo $temp -SemRelatorio 2>&1
        foreach ($linha in @($saidaUso)) { if ([string]$linha -ne '') { $log.L('  ' + [string]$linha) } }
        if ($LASTEXITCODE -ne 0) { Parar 'o teste de uso falhou (veja acima). Nada foi publicado.' }
    }

    # ------------------------------------------------------------ 9. publicação
    if ($Teste) {
        $destTeste = Join-Path (Get-PastaExecucao $raiz 'teste') $script:NomeFront
        Copy-Item -LiteralPath $temp -Destination $destTeste -Force
        $log.L(''); $log.L('TESTE CONCLUÍDO: ' + $destTeste + '  (nada publicado, versão continua ' + $versaoAtual + ')')
        $sucesso = $true
    } else {
        $log.L('--- publicação')
        if (Test-Path -LiteralPath $destinoRaiz) {
            try { $f = [System.IO.File]::Open($destinoRaiz, 'Open', 'ReadWrite', 'None'); $f.Close() }
            catch { Parar 'o .xlsm da raiz foi aberto por alguém durante o build. Nada foi publicado; rode de novo.' }
            $hist = Join-Path (Get-PastaExecucao $raiz 'historico-front') ('CRM_Zapromaq_v' + $versaoAtual + '_' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.xlsm')
            Move-Item -LiteralPath $destinoRaiz -Destination $hist
            $log.L('  anterior -> ' + (Split-Path -Leaf $hist))
        }
        Copy-Item -LiteralPath $temp -Destination $destinoRaiz
        $h1 = (Get-FileHash -LiteralPath $temp -Algorithm SHA256).Hash; $h2 = (Get-FileHash -LiteralPath $destinoRaiz -Algorithm SHA256).Hash
        if ($h1 -ne $h2) { Parar 'a cópia publicada não confere com a montada (hash diferente)' }
        $log.L('  publicado: ' + $destinoRaiz)

        if ($temBanco) {
            try { $cn = Open-Banco $bancoArq; try { Set-ConfigBanco $cn 'versao_front' $versao $null } finally { $cn.Close(); $cn.Dispose() }
                  $log.L('  config.versao_front = ' + $versao) }
            catch { $log.L('  AVISO: versão não gravada no banco (' + $_.Exception.Message.Split([char]13)[0] + '): cópias antigas não vão avisar.') }
        }
        Set-VersaoArquivo $dados $versao
        $log.L('  VERSAO.txt = ' + $versao)

        # VERSAO-FRONT.txt na raiz, ao lado do .xlsm: é por ele que o
        # INICIAR-CRM das estações decide se precisa copiar de novo. Ler um
        # texto é instantâneo; ler Início!B7 exigiria abrir o Excel (segundos
        # por estação) ou descompactar o .xlsm. Gravado DEPOIS do .xlsm: a
        # estação nunca vê número novo apontando para arquivo velho.
        # ANSI + CRLF: é o cmd.exe do .bat que lê.
        $arqVersao = Join-Path $raiz 'VERSAO-FRONT.txt'
        $conteudo = $versao + "`r`n" + $textoB7 + "`r`n" +
                    'arquivo=' + $script:NomeFront + "`r`n" +
                    'publicado=' + $quando.ToString('yyyy-MM-dd HH:mm:ss') + "`r`n" +
                    'por=' + $env:COMPUTERNAME + '\' + $env:USERNAME + "`r`n"
        [System.IO.File]::WriteAllText($arqVersao, $conteudo, (New-Object System.Text.ASCIIEncoding))
        $log.L('  VERSAO-FRONT.txt (raiz) = ' + $versao + '   <- o INICIAR-CRM compara por aqui')

        # CHANGELOG: o que mudou em fonte\ desde a publicação anterior
        $manif = Join-Path (Get-PastaExecucao $raiz 'build') 'ultima-publicacao.json'
        $agora = [ordered]@{}
        Get-ChildItem -LiteralPath $fonte -Recurse -File | ForEach-Object {
            $agora[$_.FullName.Substring($fonte.Length + 1)] = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash }
        $antes = @{}
        if (Test-Path -LiteralPath $manif) { (Read-Json $manif).PSObject.Properties | ForEach-Object { $antes[$_.Name] = $_.Value } }
        $mud = @($agora.Keys | Where-Object { -not $antes.ContainsKey($_) -or $antes[$_] -ne $agora[$_] }) + @($antes.Keys | Where-Object { -not $agora.Contains($_) } | ForEach-Object { $_ + ' (removido)' })
        [System.IO.File]::WriteAllText($manif, (ConvertTo-Json -InputObject $agora -Depth 3), (New-Object System.Text.UTF8Encoding($false)))
        $chg = Join-Path $dados 'docs\CHANGELOG.md'
        $cab = "# CHANGELOG - CRM Zapromaq`r`n`r`nCada publicação do front é registrada aqui pelo build (MONTAR-FRONTEND.bat).`r`n"
        $resto = ''
        if (Test-Path -LiteralPath $chg) { $resto = ([System.IO.File]::ReadAllText($chg, [System.Text.Encoding]::UTF8)).Replace($cab, '').TrimStart() }
        $entrada = "`r`n## " + $versao + ' - ' + $quando.ToString('dd/MM/yyyy HH:mm') + ' - ' + $env:USERNAME + '@' + $env:COMPUTERNAME + ' - ' + $Ambiente + "`r`n`r`n"
        if (@($antes.Keys).Count -eq 0) { $entrada += "- primeira publicação registrada por este build`r`n" }
        elseif ($mud.Count -eq 0) { $entrada += "- nenhuma fonte alterada (remontagem)`r`n" }
        else { foreach ($m in ($mud | Sort-Object)) { $entrada += '- ' + $m + "`r`n" } }
        [System.IO.File]::WriteAllText($chg, $cab + $entrada + "`r`n" + $resto, (New-Object System.Text.UTF8Encoding($false)))
        $log.L('  CHANGELOG.md atualizado (' + $mud.Count + ' arquivo(s) de fonte alterado(s))')

        $log.L('')
        $log.L('PUBLICADO: versão ' + $versao + '. As estações pegam a versão nova no próximo INICIAR-CRM.')
        $sucesso = $true
    }
} catch {
    $log.L('')
    $log.L('BUILD FALHOU: ' + $_.Exception.Message)
    # onde falhou (arquivo:linha) - sem isto, erro de conversão não diz de onde veio
    $log.L('Local: ' + (($_.ScriptStackTrace -split "`n" | Select-Object -First 3) -join ' <- '))
    $log.L('A raiz de BASE NÃO foi alterada.')
} finally {
    if ($xl -or $idExcel) { Close-Excel $xl $idExcel }
    if ($vbomOriginal -ne 'nao-alterado' -and $vbomOriginal -ne 1) { Set-AcessoVBOM $vbomOriginal }
    if ($trava) { $trava.Close() }
    $log.Salvar()
}
if ($sucesso) { exit 0 } else { exit 1 }
