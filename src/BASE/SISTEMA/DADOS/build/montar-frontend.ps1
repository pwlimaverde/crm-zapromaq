<#
  montar-frontend.ps1 - monta o CRM_Zapromaq.xlsm e publica o conjunto (banco + front) na raiz de BASE.

  É o ÚNICO comando da publicação. Os 12 passos, nesta ordem; a tela mostra
  "[n/12] <passo>" e, se parar, "PAROU NO PASSO n/12 (<passo>): <motivo>":

    1. pacote de municípios: municipios.csv confere com o SHA-256 do .fonte.json
    2. ambiente: Excel 64 bits, .xlsm da raiz livre, banco na raiz de BASE
    3. trava: crm_zapromaq.laccdb existe -> para e mostra as estações conectadas
    4. esquema: versao_esquema do banco e migrações pendentes
    5. backup: banco e .xlsm publicados -> execucao\backup\antes-da-publicacao (H0 = hash do banco)
    6. cópia de trabalho: execucao\publicacao\crm_zapromaq.accdb
    7. migrações aplicadas NA CÓPIA (dados em transação; estrutura guardada)
    8. conferências na cópia: municípios, papel das etapas, índice único, valores em uso
    9. montagem do front com Config!B2 provisório (apontando para a cópia): compilação e autoteste
    10. teste de uso contra a cópia: na publicação só o de LEITURA (-SoLeitura) e a impressão digital da
        cópia tem de ser igual antes e depois; no -Teste, leitura e depois o completo (banco de TESTE)
    11. troca: Config!B2 definitivo (UNC do banco de produção) no .xlsm; troca do banco de
        produção pela cópia SÓ com a trava livre e o hash igual ao H0
    12. publicação: .xlsm anterior -> execucao\historico-front; novo -> raiz; VERSAO.txt; VERSAO-FRONT.txt
        POR ÚLTIMO (ponto de não retorno); depois CHANGELOG.md (falha vira aviso). config.versao_front já
        foi gravada na cópia (passo 11). Falha antes do VERSAO-FRONT.txt devolve o .xlsm anterior, o
        VERSAO.txt e o banco (backup).

  Falha antes do passo 11 NÃO altera o banco de produção nem a raiz de BASE.

  -Teste          os mesmos passos, só sobre banco com config.ambiente = TESTE (senão recusa);
                  o passo 12 grava em execucao\teste e não sobe a versão
  -Ensaio         ensaio da atualização: passos 1 a 8 numa cópia (execucao\ensaio), sem Excel e sem trocar o
                  banco; gera execucao\ensaio\relatorio-*.txt. Exige a trava livre. Não combina com -Teste.
  -SemTesteDeUso  pula o teste de uso (passo 10)
  -Volta -Motivo "<por que>"   VOLTA (DEV-04): publica o conteúdo do pacote anterior extraído por cima da BASE
                  como nova versão MAIOR; recusa sem motivo, sem pacote, pacote que não é anterior ao publicado ou
                  esquema esperado pela fonte diferente do banco. Ver build\lib\Release.ps1.

  Toda publicação grava execucao\releases\release-<versão>.json (manifesto de release: commit, versões do
  front, conector e esquema, migrações, hashes do .xlsm e dos scripts).
  -Ambiente PRODUCAO | TESTE (padrão: config.ambiente do banco; sem banco, PRODUCAO)

  "Confiar no acesso ao modelo de objeto do projeto do VBA" é ligado só durante a
  montagem (HKCU, sem administrador) e volta ao valor original no fim.
#>
param(
    [switch]$Teste,
    [switch]$SemTesteDeUso,
    [switch]$Ensaio,
    [switch]$Volta,
    [string]$Motivo = '',
    [ValidateSet('', 'PRODUCAO', 'TESTE')][string]$Ambiente = '',
    [int]$TempoLimite = 300
)
$ErrorActionPreference = 'Stop'
$dados = Split-Path -Parent $PSScriptRoot
. (Join-Path $dados 'lib\Comum.ps1')
. (Join-Path $PSScriptRoot 'lib\Versao.ps1')
. (Join-Path $PSScriptRoot 'lib\Pacote.ps1')
. (Join-Path $PSScriptRoot 'lib\Excel.ps1')
. (Join-Path $PSScriptRoot 'lib\Release.ps1')
. (Join-Path $dados 'lib\AtualizacaoBanco.ps1')

$raiz = Get-RaizBase $PSScriptRoot
$fonte = Join-Path $dados 'fonte'
$tema = Read-Json (Join-Path $fonte 'layout\tema.json')
$abas = Read-Json (Join-Path $fonte 'layout\abas.json')
$script:IconesLayout = Read-Json (Join-Path $fonte 'layout\icones.json')
$log = New-Log $raiz 'build'
$pastaBuild = Get-PastaExecucao $raiz 'build'
$destinoRaiz = Get-CaminhoFront $raiz
$bancoArq = Get-CaminhoBanco $raiz

$xl = $null; $idExcel = 0; $trava = $null; $vbomOriginal = 'nao-alterado'; $sucesso = $false; $pontoSemRetorno = $false
$script:passo = 2
$nomesPassos = @{ 1 = 'pacote de municípios'; 2 = 'ambiente'; 3 = 'trava'; 4 = 'esquema'; 5 = 'backup'; 6 = 'cópia de trabalho';
                  7 = 'migrações na cópia'; 8 = 'conferências da cópia'; 9 = 'montagem do front'; 10 = 'teste de uso';
                  11 = 'troca do banco'; 12 = 'publicação' }
function Parar([string]$msg) { throw $msg }
function Parar-NoPasso([int]$n, [string]$msg) { throw ('PASSO:' + $n + '|' + $msg) }

try {
    $log.L('=== MONTAGEM DO FRONTEND - CRM Zapromaq ===')
    $log.L('Data....: ' + (Get-Date -Format 'dd/MM/yyyy HH:mm:ss') + '   ' + $env:COMPUTERNAME + '\' + $env:USERNAME)
    $log.L('Raiz....: ' + (ConvertTo-CaminhoUNC $raiz))
    $log.L('Modo....: ' + $(if ($Ensaio) { 'ENSAIO (passos 1 a 8 numa cópia; não troca o banco, não monta o front)' } elseif ($Teste) { 'TESTE (não publica)' } else { 'PUBLICAÇÃO' }))

    # ------------------------------------------------------------ 1. pré-checagens
    try { $trava = [System.IO.File]::Open((Join-Path $pastaBuild 'build.lock'), 'OpenOrCreate', 'ReadWrite', 'None') }
    catch { Parar 'Outro build está em andamento (execucao\build\build.lock em uso). Espere ele terminar.' }

    if ($Ensaio -and $Teste) { Parar '-Ensaio e -Teste não se combinam: o ensaio nunca troca banco nem monta o front.' }
    if ($Ensaio -and $Volta) { Parar '-Ensaio e -Volta não se combinam.' }
    if ($Volta -and [string]::IsNullOrWhiteSpace($Motivo)) { Parar (Test-PedidoVolta '' $null '' '' '') }

    if (-not $Ensaio -and $null -eq [type]::GetTypeFromProgID('Excel.Application')) {
        Parar 'Excel não está instalado nesta máquina. O build roda numa estação com Excel 2019 64 bits (veja testes\CHECAR-AMBIENTE.bat).'
    }
    if (-not [Environment]::Is64BitProcess) { Parar 'Rode pelo MONTAR-FRONTEND.bat: o build precisa do PowerShell 64 bits.' }

    if (-not $Teste -and -not $Ensaio -and (Test-Path -LiteralPath $destinoRaiz)) {
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
    if ($Teste) { $motivoTeste = Test-ModoTeste $Ambiente; if ($motivoTeste) { Parar-NoPasso 2 $motivoTeste } }
    # publicar sem o banco geraria uma planilha apontando para um arquivo que não existe
    # (por exemplo, depois de trocar a pasta BASE inteira e esquecer o .accdb)
    if (-not $Teste -and -not $temBanco) {   # vale também para o ensaio: sem banco não há o que ensaiar
        Parar ('o banco ' + $script:NomeBanco + ' não está na raiz de BASE ou não abriu (veja o aviso acima). Ele precisa estar em ' + $bancoArq + '. Nada foi publicado.')
    }

    # ------------------------------------------------------------ passo 1: pacote de municípios
    $script:passo = 1
    $log.L('')
    $log.L('[1/12] pacote de municípios...')
    $motivoPacote = Test-PacoteMunicipios $dados
    if ($motivoPacote) { Parar-NoPasso 1 $motivoPacote }
    $motivoManifesto = Test-Manifesto $raiz
    if ($motivoManifesto) { Parar-NoPasso 1 $motivoManifesto }
    $totalMunicipios = [int]((Get-Content -LiteralPath (Join-Path $dados 'banco\referencia\municipios.fonte.json') -Raw -Encoding UTF8 | ConvertFrom-Json).total)

    # ------------------------------------------------------------ -Ensaio: passos 1 a 8 numa cópia, sem Excel e sem trocar nada
    if ($Ensaio) {
        $script:passo = 3
        $log.L('[3/12] a 8/12 (ENSAIO): trava, cópia, migração na cópia e conferências...')
        $ens = Invoke-EnsaioDoBanco $bancoArq (Get-PastaExecucao $raiz 'ensaio') $totalMunicipios $log
        if (-not $ens.ok) { Parar-NoPasso $ens.passo $ens.motivo }
        $log.L('')
        $log.L('ENSAIO APROVADO. Esquema ' + $ens.esquema_antes + ' -> ' + $ens.esquema_depois + ' na cópia; migrações: ' + (@($ens.migracoes | ForEach-Object { ($_ -split '\s+')[0] }) -join ', ') +
               '; hash do banco de produção igual antes e depois (' + $ens.h0.Substring(0, 12) + '); ' + $ens.segundos + ' s.')
        $log.L('Relatório: ' + $ens.relatorio + '   (tem dado real: não vai para o repositório)')
        $log.L('O banco de produção e a raiz de BASE não foram alterados. Para publicar, rode o MONTAR-FRONTEND.bat sem -Ensaio.')
        $sucesso = $true
        exit 0
    }

    # ------------------------------------------------------------ passos 3 a 8: trava, esquema, backup, cópia, migração na cópia, conferências
    $script:passo = 5
    $log.L('[3/12] a 8/12: trava, esquema, backup, cópia, migração na cópia e conferências...')
    $pastaBkp = New-BackupAntesDePublicar $raiz $bancoArq $destinoRaiz $log
    $log.L('  backup do publicado em ' + $pastaBkp)
    $prep = Invoke-PreparacaoDoBanco $bancoArq (Get-PastaExecucao $raiz 'publicacao') $pastaBkp $totalMunicipios $log
    if (-not $prep.ok) { Parar-NoPasso $prep.passo $prep.motivo }
    $log.L('  cópia pronta: ' + $prep.copia)
    # DEV-04: o front destas fontes tem de esperar o esquema do banco que vai ser publicado (vale para a volta e
    # para a publicação normal); e a volta tem motivo, pacote anterior ao publicado e esquema compatível
    $cnE = Open-BancoLeitura $prep.copia
    try {
        $esquemaCopia = Get-VersaoEsquemaBanco $cnE
        $migracoesCopia = New-Object System.Collections.ArrayList
        $cmE = New-ComandoBanco $cnE "SELECT chave FROM config WHERE chave LIKE 'migracao.%' ORDER BY chave" @() $null
        $rdE = $cmE.ExecuteReader(); while ($rdE.Read()) { [void]$migracoesCopia.Add(([string]$rdE.GetValue(0)).Substring(9)) }; $rdE.Close(); $cmE.Dispose()
    } finally { $cnE.Close(); $cnE.Dispose() }
    $esquemaFonte = Get-EsquemaDaFonte $dados
    $pacote = Get-ManifestoPacote $raiz
    if ($Volta) {
        $motivoVolta = Test-PedidoVolta $Motivo $pacote (Get-VersaoPublicada $raiz) $esquemaFonte $esquemaCopia
        if ($motivoVolta) { Parar-NoPasso 8 $motivoVolta }
        $log.L('  VOLTA: conteúdo da versão ' + $pacote.versao + ' (commit ' + $pacote.commit + '), motivo: ' + $Motivo)
    } elseif ($esquemaFonte -ne $esquemaCopia) {
        Parar-NoPasso 8 ('o front destas fontes espera o esquema ' + $esquemaFonte + ' e o banco migrado está no ' + $esquemaCopia + '. Nada foi publicado.')
    }
    # -Teste (só roda com ambiente TESTE, passo 2) declara a homologação NA CÓPIA: os testes que gravam exigem
    # config.homologacao = 1. Bancos de teste antigos, de antes da spec 005, ganham a marca por aqui.
    if ($Teste) {
        $cnH = Open-Banco $prep.copia
        try { Set-ConfigBanco $cnH 'homologacao' '1' $null } finally { $cnH.Close(); $cnH.Dispose() }
        $log.L('  homologação declarada na cópia (config.homologacao = 1)')
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
    $versaoAtual = Get-VersaoDePartida $dados $raiz
    if ($versaoAtual -ne (Get-VersaoAtual $dados)) {
        $log.L('AVISO: VERSAO.txt (' + (Get-VersaoAtual $dados) + ') está atrás da versão publicada (' + $versaoAtual + '); a numeração continua da publicada.')
    }
    $versao = Get-ProximaVersao $versaoAtual
    $quando = Get-Date
    $textoB7 = Format-TextoB7 $Ambiente $versao $quando
    $log.L('Versão..: ' + $versaoAtual + ' -> ' + $versao + '   ambiente ' + $Ambiente)
    $log.L('')

    # ------------------------------------------------------------ passo 9: montagem do front (aponta para a CÓPIA, provisório)
    $script:passo = 9
    if ($Teste -and $env:CRM_FALHA_NO_PASSO -eq '9') { throw 'PASSO:9|falha forçada pelo teste (CRM_FALHA_NO_PASSO)' }
    $log.L('[9/12] montagem do front (Config!B2 provisório = cópia de trabalho)...')
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
    # provisório: aponta para a cópia de trabalho; o passo 11 grava o caminho definitivo do banco de produção
    $cfg.Range($abas.config.banco).Value2 = (ConvertTo-CaminhoUNC $prep.copia)
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

    # Planejamentos (fatia 4.8): logo, título e botões; seletor e grade vêm do modAbaPlanejamentos
    if ($abas.PSObject.Properties['planejamentos']) {
        $wsL = $wb.Worksheets.Item('Planejamentos')
        $b = $abas.planejamentos.botoes
        $x = [double]$b.pos[0]; $n = 0
        foreach ($it in $b.itens) {
            $n++
            [void](Add-BotaoMarca $wsL $it.texto $it.macro @($x, [double]$b.pos[1], [double]$b.pos[2], [double]$b.pos[3]) ('btnPla' + $n) (Resolve-ValorTema $tema $it.cor) $tema)
            $x += [double]$b.passo
        }
        if (Test-Path -LiteralPath $logoArq) {
            $l = $abas.grades.logo.pos
            $lg = $wsL.Shapes.AddPicture($logoArq, $false, $true, [double]$l[0], [double]$l[1], [double]$l[2], [double]$l[3])
            $lg.Name = 'logoAba'; try { $lg.Placement = 3 } catch { }
        }
        $wsL.Columns.Item(1).ColumnWidth = 2.25
        $wsL.Rows.Item(1).RowHeight = 4; $wsL.Rows.Item(2).RowHeight = 28; $wsL.Rows.Item(3).RowHeight = 34
        $wsL.Range('A1:AN6').Interior.Color = $fundo
        $tit = $wsL.Shapes.AddTextbox(1, 96, $wsL.Rows.Item(2).Top, 460, $wsL.Rows.Item(2).Height)
        $tit.Name = 'tituloAba'
        $tit.Line.Visible = $false; $tit.Fill.Visible = $false
        try { $tit.Placement = 3 } catch { }
        $tf = $tit.TextFrame2
        $tf.MarginLeft = 0; $tf.MarginRight = 0; $tf.MarginTop = 0; $tf.MarginBottom = 0
        $tf.VerticalAnchor = 3; $tf.WordWrap = $false
        $tf.TextRange.Text = 'Planejamentos'
        $tf.TextRange.Font.Name = (Resolve-ValorTema $tema '@semibold')
        $tf.TextRange.Font.Size = 15
        $tf.TextRange.Font.Bold = $true
        $tf.TextRange.Font.Fill.ForeColor.RGB = (Resolve-ValorTema $tema '@azul')
        $wsL.Range($abas.planejamentos.aviso.celula).Value2 = $abas.planejamentos.aviso.texto
        $log.L('  Planejamentos ' + @($b.itens).Count + ' botões')
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

    # ------------------------------------------------------------ passo 10: teste de uso (contra a cópia)
    $script:passo = 10
    if ($Teste -and $env:CRM_FALHA_NO_PASSO -eq '10') { throw 'PASSO:10|falha forçada pelo teste (CRM_FALHA_NO_PASSO)' }
    if (-not $SemTesteDeUso) {
        $log.L('')
        $log.L('--- teste de uso (grades e Painel com os dados reais)')
        # Publicação: só o teste de LEITURA, e a impressão digital da cópia tem de ser a mesma antes e depois
        # (o teste que grava só roda no -Teste, sobre banco de TESTE). A cópia vira a produção no passo 11.
        $impAntes = Get-ImpressaoDoBanco $prep.copia
        foreach ($modoUso in @(Get-ExecucoesTesteDeUso ([bool]$Teste))) {
            $argsUso = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', (Join-Path $dados 'testes\testar-uso.ps1'), '-Arquivo', $temp, '-SemRelatorio')
            if ($modoUso -eq 'SoLeitura') { $argsUso += '-SoLeitura' }
            $log.L('  teste de uso: ' + $(if ($modoUso -eq 'SoLeitura') { 'só leitura' } else { 'completo (banco de TESTE)' }))
            $saidaUso = & powershell.exe @argsUso 2>&1
            foreach ($linha in @($saidaUso)) { if ([string]$linha -ne '') { $log.L('  ' + [string]$linha) } }
            if ($LASTEXITCODE -ne 0) { Parar 'o teste de uso falhou (veja acima). Nada foi publicado.' }
            if ($modoUso -eq 'SoLeitura') {
                $impDepois = Get-ImpressaoDoBanco $prep.copia
                if ($impDepois -ne $impAntes) { Parar-NoPasso 10 'o teste de uso de leitura alterou a cópia do banco (impressão digital diferente). Nada foi publicado.' }
                $log.L('  teste de uso só leitura: impressão da cópia igual antes e depois')
            }
        }
    }

    # ------------------------------------------------------------ passo 11: troca do banco (tudo ou nada)
    $script:passo = 11
    if ($Teste -and $env:CRM_FALHA_NO_PASSO -eq '11') { throw 'PASSO:11|falha forçada pelo teste (CRM_FALHA_NO_PASSO)' }
    $log.L('[11/12] troca do banco...')
    $bancoFinal = ConvertTo-CaminhoUNC $bancoArq
    $xl2 = $null; $idExcel2 = 0
    try {
        $xl2 = New-Excel ([ref]$idExcel2)
        $xl2.Visible = $false; $xl2.DisplayAlerts = $false; $xl2.EnableEvents = $false
        $wb2 = $xl2.Workbooks.Open($temp)
        $wb2.Worksheets.Item('Config').Range($abas.config.banco).Value2 = $bancoFinal
        $wb2.Save(); $wb2.Close($false)
    } finally { Close-Excel $xl2 $idExcel2 }
    if ((Get-TextoCelulaXlsx $temp 'Config' $abas.config.banco) -ne $bancoFinal) { Parar-NoPasso 11 'Config!B2 definitivo não conferiu no .xlsm montado' }
    $log.L('  Config!B2 definitivo = ' + $bancoFinal)
    # config.versao_front vai NA CÓPIA, antes da troca: depois dela o hash da produção ficaria diferente do
    # da troca e a reversão (Reverter-Troca) nunca conseguiria voltar. Só na publicação (o -Teste não grava).
    if (-not $Teste -and $temBanco) {
        try { Set-VersaoFrontNaCopia $prep.copia $versao; $log.L('  config.versao_front = ' + $versao + ' (na cópia)') }
        catch { Parar-NoPasso 11 ('não foi possível gravar config.versao_front na cópia: ' + $_.Exception.Message.Split([char]13)[0] + '. Nada foi trocado.') }
    }
    $troca = Trocar-BancoProducao $prep.copia $bancoArq $prep.h0 $pastaBkp
    if (-not $troca.ok) { Parar-NoPasso 11 $troca.motivo }
    $hashNovo = Get-HashArquivo $bancoArq
    $log.L('  ' + $troca.motivo)

    # ------------------------------------------------------------ passo 12: publicação (se falhar, o banco volta ao backup)
    $script:passo = 12
    $log.L('[12/12] publicação...')
    try {
    if ($Teste -and $env:CRM_FALHA_NO_PASSO -eq '12') { throw 'PASSO:12|falha forçada pelo teste (CRM_FALHA_NO_PASSO)' }
    if ($Teste) {
        $destTeste = Join-Path (Get-PastaExecucao $raiz 'teste') $script:NomeFront
        Copy-Item -LiteralPath $temp -Destination $destTeste -Force
        $manT = New-ManifestoRelease -Versao $versao -Tipo $(if ($Volta) { 'volta' } else { 'teste' }) -Motivo $Motivo -Pacote $pacote -VersaoFront $versao -VersaoConector (Get-VersaoConectorDaFonte $raiz) -Esquema $esquemaCopia -Migracoes @($migracoesCopia) -Xlsm $destTeste -Dados $dados -Raiz $raiz
        $log.L('  manifesto de release (teste): ' + (Save-ManifestoRelease $manT (Get-PastaExecucao $raiz 'teste')))
        $log.L(''); $log.L('TESTE CONCLUÍDO: ' + $destTeste + '  (nada publicado, versão continua ' + $versaoAtual + ')')
        $sucesso = $true
    } else {
        $log.L('--- publicação (recuperável até gravar o VERSAO-FRONT.txt)')
        # Ordem: .xlsm anterior -> histórico, .xlsm novo, VERSAO.txt e, por ÚLTIMO, o VERSAO-FRONT.txt (ponto de
        # não retorno: o INICIAR-CRM das estações decide por ele; a estação nunca vê número novo apontando para
        # arquivo velho). Falha antes dele devolve o .xlsm anterior, o VERSAO.txt e o VERSAO-FRONT.txt, e o catch
        # abaixo devolve também o banco (config.versao_front já foi gravada na cópia, no passo 11, então o hash
        # da produção ainda é o da troca e a reversão é possível).
        # falha simulada de subpasso: só com banco de TESTE (testes\testar-publicacao-copia.ps1)
        $falhaSub = ''
        if ($Ambiente -eq 'TESTE' -and $env:CRM_FALHA_SUBPASSO_12) { $falhaSub = [string]$env:CRM_FALHA_SUBPASSO_12 }
        $pub = Publicar-FrontNaRaiz $temp $raiz $dados $versao $textoB7 $quando (Get-PastaExecucao $raiz 'historico-front') $versaoAtual $falhaSub
        if (-not $pub.ok) { Parar-NoPasso 12 ('publicação do front: ' + $pub.motivo) }
        $pontoSemRetorno = $true
        $log.L('  publicado: ' + $destinoRaiz + '   (etapas: ' + ($pub.ordem -join ', ') + ')')
        $log.L('  config.versao_front = ' + $versao + ' (gravada na cópia, antes da troca)')
        $log.L('  VERSAO.txt = ' + $versao)
        $log.L('  VERSAO-FRONT.txt (raiz) = ' + $versao + '   <- o INICIAR-CRM compara por aqui; ponto de não retorno')

        # Depois do ponto de não retorno nada reverte: falha aqui vira aviso.
        try {
            $manR = New-ManifestoRelease -Versao $versao -Tipo $(if ($Volta) { 'volta' } else { 'publicacao' }) -Motivo $Motivo -Pacote $pacote -VersaoFront $versao -VersaoConector (Get-VersaoConectorDaFonte $raiz) -Esquema $esquemaCopia -Migracoes @($migracoesCopia) -Xlsm $destinoRaiz -Dados $dados -Raiz $raiz
            $log.L('  manifesto de release: ' + (Save-ManifestoRelease $manR (Get-PastaExecucao $raiz 'releases')))
        } catch { $log.L('  AVISO: o manifesto de release não foi gravado (' + $_.Exception.Message.Split([char]13)[0] + '). A versão já está publicada.') }
        try {
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
        # fim de linha normalizado antes de tirar o cabeçalho: com LF ele não casava e saía duplicado
        if (Test-Path -LiteralPath $chg) { $resto = (([System.IO.File]::ReadAllText($chg, [System.Text.Encoding]::UTF8)) -replace '\r?\n', "`r`n").Replace($cab, '').TrimStart() }
        $entrada = "`r`n## " + $versao + ' - ' + $quando.ToString('dd/MM/yyyy HH:mm') + ' - ' + $env:USERNAME + '@' + $env:COMPUTERNAME + ' - ' + $Ambiente + "`r`n`r`n"
        if ($Volta) { $entrada += '- VOLTA: conteúdo da versão ' + $pacote.versao + ' (commit ' + $pacote.commit + '). Motivo: ' + $Motivo + "`r`n" }
        if (@($antes.Keys).Count -eq 0) { $entrada += "- primeira publicação registrada por este build`r`n" }
        elseif ($mud.Count -eq 0) { $entrada += "- nenhuma fonte alterada (remontagem)`r`n" }
        else { foreach ($m in ($mud | Sort-Object)) { $entrada += '- ' + $m + "`r`n" } }
        [System.IO.File]::WriteAllText($chg, $cab + $entrada + "`r`n" + $resto, (New-Object System.Text.UTF8Encoding($false)))
        $log.L('  CHANGELOG.md atualizado (' + $mud.Count + ' arquivo(s) de fonte alterado(s))')
        } catch { $log.L('  AVISO: o CHANGELOG não foi atualizado (' + $_.Exception.Message.Split([char]13)[0] + '). A versão já está publicada; registre a mudança à mão.') }

        $unif = 0; $inativos = 0; $munic = 0; $relatorio = ''
        foreach ($linha in @($prep.migracao)) {
            if ($linha -match 'unificações: (\d+)') { $unif = [int]$matches[1] }
            if ($linha -match 'municípios: (\d+) carregado') { $munic = [int]$matches[1] }
            if ($linha -match 'entrou INATIVO') { $inativos++ }
            if ($linha -match '^Log: (.+)$') { $relatorio = $matches[1].Trim() }
        }
        $log.L('')
        $log.L('PRONTO. CRM ' + $versao + ' publicado; banco no esquema 1.1. Municípios: ' + $munic + '. Listas: ' + $unif +
               ' unificação(ões), ' + $inativos + ' valor(es) incluído(s) como inativo(s) (relatório: ' + $relatorio + '). Backup: ' + $pastaBkp + '.')
        $log.L('Usuários do Claude Desktop: fechem e abram o aplicativo e reinstalem a skill IA\instrucoes\crm-zapromaq.zip.')
        $sucesso = $true
    }
    } catch {
        if ($troca -and $troca.ok -and -not $pontoSemRetorno) {
            $reversao = Reverter-Troca $bancoArq $prep.backup $hashNovo
            $log.L('  ' + $reversao.motivo)
        }
        throw
    }
} catch {
    $msg = $_.Exception.Message
    $n = $script:passo
    if ($msg -match '^PASSO:(\d+)\|(.*)$') { $n = [int]$matches[1]; $msg = $matches[2] }
    $log.L('')
    $log.L('PAROU NO PASSO ' + $n + '/12 (' + $nomesPassos[$n] + '): ' + $msg)
    # onde falhou (arquivo:linha) - sem isto, erro de conversão não diz de onde veio
    $log.L('Local: ' + (($_.ScriptStackTrace -split "`n" | Select-Object -First 3) -join ' <- '))
    if ($n -lt 11) { $log.L('Nada foi publicado: a raiz de BASE e o banco de produção não foram alterados.') }
    else { $log.L('Veja o backup em ' + $pastaBkp + ' antes de mexer.') }
} finally {
    if ($xl -or $idExcel) { Close-Excel $xl $idExcel }
    if ($vbomOriginal -ne 'nao-alterado' -and $vbomOriginal -ne 1) { Set-AcessoVBOM $vbomOriginal }
    if ($trava) { $trava.Close() }
    $log.Salvar()
}
if ($sucesso) { exit 0 } else { exit 1 }
