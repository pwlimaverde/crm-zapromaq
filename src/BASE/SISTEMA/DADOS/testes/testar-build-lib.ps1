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
. (Join-Path $dados 'lib\AtualizacaoBanco.ps1')

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

Write-Host '--- banco: trava (C19), H0 na troca (C20) e reversão'
$tmpB = Join-Path ([System.IO.Path]::GetTempPath()) ('crm-banco-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path (Join-Path $tmpB 'trabalho') -Force | Out-Null
try {
    $prod = Join-Path $tmpB 'crm_zapromaq.accdb'
    $copia = Join-Path $tmpB 'trabalho\crm_zapromaq.accdb'
    $back = Join-Path $tmpB 'backup'
    $laccdb = [System.IO.Path]::ChangeExtension($prod, 'laccdb')
    [IO.File]::WriteAllBytes($prod, [Text.Encoding]::ASCII.GetBytes('producao-antiga'))
    [IO.File]::WriteAllBytes($copia, [Text.Encoding]::ASCII.GetBytes('copia-migrada'))

    # C19: trava com uma estação conectada -> não livre, com o nome lido da trava
    $trava = New-Object byte[] 128
    $nomeTrava = [Text.Encoding]::ASCII.GetBytes('ESTACAO-TESTE')
    [Array]::Copy($nomeTrava, 0, $trava, 0, $nomeTrava.Length)
    [IO.File]::WriteAllBytes($laccdb, $trava)
    $livre = Test-BancoLivre $prod
    Conferir 'C19 trava presente: banco não livre' (-not $livre.livre)
    Conferir 'C19 nome da estação lido da trava (registro vazio ignorado)' (@($livre.estacoes).Count -eq 1 -and $livre.estacoes[0] -eq 'ESTACAO-TESTE') ($livre.estacoes -join ',')
    Remove-Item -LiteralPath $laccdb -Force
    Conferir 'sem trava: banco livre' ((Test-BancoLivre $prod).livre)

    # troca com H0 igual: produção vira a cópia; o anterior vai para o backup
    $h0 = Get-HashArquivo $prod
    $r = Trocar-BancoProducao $copia $prod $h0 $back
    Conferir 'troca com H0 igual: ok' $r.ok $r.motivo
    Conferir 'produção passou a ser a cópia' ((Get-HashArquivo $prod) -eq (Get-HashArquivo $copia))
    Conferir 'backup guarda o banco anterior (H0)' ((Get-HashArquivo (Join-Path $back 'crm_zapromaq.accdb')) -eq $h0)

    # C20: alguém gravou depois do H0 -> recusa e nada muda
    $copia2 = Join-Path $tmpB 'trabalho\nova.accdb'
    [IO.File]::WriteAllBytes($copia2, [Text.Encoding]::ASCII.GetBytes('copia-nova'))
    $h0b = Get-HashArquivo $prod
    [IO.File]::WriteAllBytes($prod, [Text.Encoding]::ASCII.GetBytes('gravado-durante-a-montagem'))
    $antesDaRecusa = Get-HashArquivo $prod
    $r2 = Trocar-BancoProducao $copia2 $prod $h0b (Join-Path $tmpB 'backup2')
    Conferir 'C20 hash diferente de H0: recusa' (-not $r2.ok)
    Conferir 'C20 nada foi trocado' ((Get-HashArquivo $prod) -eq $antesDaRecusa)
    Conferir 'C20 mensagem manda rodar de novo' ($r2.motivo -match 'rode de novo')

    # reversão: volta ao backup enquanto a produção ainda é a da troca
    $r3 = Trocar-BancoProducao $copia2 $prod $antesDaRecusa (Join-Path $tmpB 'backup3')
    Conferir 'troca seguinte ok' $r3.ok $r3.motivo
    $hNovo = Get-HashArquivo $prod
    $volta = Reverter-Troca $prod (Join-Path $tmpB 'backup3\crm_zapromaq.accdb') $hNovo
    Conferir 'reversão volta ao banco anterior' ($volta.ok -and (Get-HashArquivo $prod) -eq $antesDaRecusa) $volta.motivo
    Conferir 'reversão recusa se a produção mudou desde a troca' (-not (Reverter-Troca $prod (Join-Path $tmpB 'backup3\crm_zapromaq.accdb') $hNovo).ok)
} finally { Remove-Item -LiteralPath $tmpB -Recurse -Force -ErrorAction SilentlyContinue }

Write-Host '--- montagem: pacote de municípios (passo 1), modo TESTE (passo 2) e passos 3 a 8 com erro forçado'
Conferir 'passo 1: pacote de municípios confere com o .fonte.json' ((Test-PacoteMunicipios $dados) -eq '') (Test-PacoteMunicipios $dados)
Conferir '-Teste recusa banco de PRODUCAO' ((Test-ModoTeste 'PRODUCAO') -ne '')
Conferir '-Teste aceita banco de TESTE' ((Test-ModoTeste 'TESTE') -eq '')
$tmpP = Join-Path ([System.IO.Path]::GetTempPath()) ('crm-prep-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tmpP -Force | Out-Null
try {
    $prodP = Join-Path $tmpP 'crm_zapromaq.accdb'
    [IO.File]::WriteAllBytes($prodP, [Text.Encoding]::ASCII.GetBytes('producao-sem-migrar'))
    $hP = Get-HashArquivo $prodP
    $logP = [pscustomobject]@{ linhas = New-Object System.Collections.ArrayList }
    $logP | Add-Member -MemberType ScriptMethod -Name L -Value { param($x) [void]$this.linhas.Add([string]$x) }
    $pr = Invoke-PreparacaoDoBanco $prodP (Join-Path $tmpP 'publicacao') (Join-Path $tmpP 'backup') 5571 $logP
    Conferir 'passo 5 (backup validado) recusa banco que não abre' ((-not $pr.ok) -and $pr.passo -eq 5) ($pr.motivo)
    Conferir 'erro no passo 5: banco de produção intacto' ((Get-HashArquivo $prodP) -eq $hP)
    Conferir 'erro no passo 5: backup criado com o H0' ((Get-HashArquivo (Join-Path $tmpP 'backup\crm_zapromaq.accdb')) -eq $hP)
    $tmpT = Join-Path $tmpP 'trava'
    New-Item -ItemType Directory -Path $tmpT -Force | Out-Null
    $prodT = Join-Path $tmpT 'crm_zapromaq.accdb'
    [IO.File]::WriteAllBytes($prodT, [Text.Encoding]::ASCII.GetBytes('x'))
    $trT = New-Object byte[] 64
    [Array]::Copy([Text.Encoding]::ASCII.GetBytes('ESTACAO-B'), 0, $trT, 0, 9)
    [IO.File]::WriteAllBytes([IO.Path]::ChangeExtension($prodT, 'laccdb'), $trT)
    $pt = Invoke-PreparacaoDoBanco $prodT (Join-Path $tmpT 'publicacao') (Join-Path $tmpT 'backup') 5571 $logP
    Conferir 'C19 via montagem: para no passo 3 e cita a estação' ((-not $pt.ok) -and $pt.passo -eq 3 -and $pt.motivo -match 'ESTACAO-B') ($pt.motivo)
    Conferir 'C19: nada criado em publicação nem backup' (-not (Test-Path -LiteralPath (Join-Path $tmpT 'publicacao')) -and -not (Test-Path -LiteralPath (Join-Path $tmpT 'backup')))
} finally { Remove-Item -LiteralPath $tmpP -Recurse -Force -ErrorAction SilentlyContinue }

Write-Host '--- pacote: MANIFESTO.json (1.10)'
$tmpM = Join-Path ([System.IO.Path]::GetTempPath()) ('crm-manifesto-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path (Join-Path $tmpM 'SISTEMA\DADOS') -Force | Out-Null
try {
    [IO.File]::WriteAllText((Join-Path $tmpM 'SISTEMA\DADOS\a.txt'), 'um', [Text.Encoding]::ASCII)
    $man = Join-Path $tmpM 'SISTEMA\DADOS\MANIFESTO.json'
    $hashA = Get-HashArquivo (Join-Path $tmpM 'SISTEMA\DADOS\a.txt')
    [IO.File]::WriteAllText($man, ('{"arquivos":[{"caminho":"SISTEMA/DADOS/a.txt","sha256":"' + $hashA + '"}]}'), [Text.Encoding]::UTF8)
    Conferir 'manifesto confere com os arquivos do pacote' ((Test-Manifesto $tmpM) -eq '') (Test-Manifesto $tmpM)
    [IO.File]::WriteAllText((Join-Path $tmpM 'SISTEMA\DADOS\a.txt'), 'dois', [Text.Encoding]::ASCII)
    Conferir 'arquivo alterado depois do pacote é recusado' ((Test-Manifesto $tmpM) -match 'a.txt') (Test-Manifesto $tmpM)
    Remove-Item -LiteralPath (Join-Path $tmpM 'SISTEMA\DADOS\a.txt') -Force
    Conferir 'arquivo listado e ausente é recusado' ((Test-Manifesto $tmpM) -match 'a.txt') (Test-Manifesto $tmpM)
} finally { Remove-Item -LiteralPath $tmpM -Recurse -Force -ErrorAction SilentlyContinue }

Write-Host '--- caminho canônico: UNC x letra x relativo (spec 005)'
Conferir 'mesmo caminho ignora caixa' (Test-MesmoCaminho 'C:\Temp\Dados\crm.accdb' 'c:\temp\DADOS\CRM.ACCDB')
Conferir 'mesmo caminho ignora barra final e ponto' (Test-MesmoCaminho 'C:\Temp\Dados\' 'C:\Temp\.\Dados')
Conferir 'caminhos diferentes' (-not (Test-MesmoCaminho 'C:\Temp\Dados\a.accdb' 'C:\Temp\Dados\b.accdb'))
Push-Location ([System.IO.Path]::GetTempPath())
try { $relT = 'crm-x.accdb'; Conferir 'relativo = absoluto' (Test-MesmoCaminho $relT (Join-Path ([System.IO.Path]::GetTempPath()) 'crm-x.accdb')) } finally { Pop-Location }
Conferir 'compartilhamento administrativo da própria máquina = letra' (Test-MesmoCaminho ('\\localhost\C$\Temp\Dados\crm.accdb') 'C:\Temp\Dados\crm.accdb')
Conferir ('\\' + $env:COMPUTERNAME + '\C$ = letra') (Test-MesmoCaminho ('\\' + $env:COMPUTERNAME + '\C$\Temp\Dados\crm.accdb') 'C:\Temp\Dados\crm.accdb')
Conferir 'UNC de outro servidor não é a letra local' (-not (Test-MesmoCaminho '\\servidor\C$\Temp\Dados\crm.accdb' 'C:\Temp\Dados\crm.accdb'))
Conferir 'prefixo \\?\ é ignorado' (Test-MesmoCaminho '\\?\C:\Temp\Dados\crm.accdb' 'C:\Temp\Dados\crm.accdb')

Write-Host '--- troca sem janela (spec 004, F6): .novo pronto ANTES da última conferência'
$tmpT = Join-Path ([System.IO.Path]::GetTempPath()) ('crm-troca-' + $PID)
New-Item -ItemType Directory -Path $tmpT -Force | Out-Null
try {
    $prodT = Join-Path $tmpT 'crm.accdb'; $copiaT = Join-Path $tmpT 'trabalho\crm.accdb'; $bkpT = Join-Path $tmpT 'backup'
    New-Item -ItemType Directory -Path (Split-Path -Parent $copiaT) -Force | Out-Null
    function Preparar { [IO.File]::WriteAllBytes($prodT, [byte[]](1..200)); [IO.File]::WriteAllBytes($copiaT, [byte[]](201..255)); Remove-Item -LiteralPath ($prodT + '.novo') -Force -ErrorAction SilentlyContinue }
    Preparar
    $h0T = Get-HashArquivo $prodT; $hCopiaT = Get-HashArquivo $copiaT
    $script:GanchoAntesDaTroca = $null
    $rT = Trocar-BancoProducao $copiaT $prodT $h0T $bkpT
    Conferir 'troca sem interferência: ok e produção igual à cópia' ($rT.ok -and (Get-HashArquivo $prodT) -eq $hCopiaT) $rT.motivo
    Conferir 'troca guarda o banco anterior no backup' ((Get-HashArquivo (Join-Path $bkpT 'crm.accdb')) -eq $h0T)

    Preparar
    $script:novoJaPronto = $false
    $script:GanchoAntesDaTroca = { $script:novoJaPronto = (Test-Path -LiteralPath ($prodT + '.novo')); $fs = [IO.File]::Open($prodT, 'Append'); $fs.WriteByte(7); $fs.Close() }
    $rT = Trocar-BancoProducao $copiaT $prodT $h0T $bkpT
    $script:GanchoAntesDaTroca = $null
    Conferir 'o .novo já estava pronto quando a gravação tardia aconteceu' $script:novoJaPronto
    Conferir 'gravação depois de preparar o .novo: troca recusada' (-not $rT.ok -and $rT.motivo -match 'gravou') $rT.motivo
    Conferir 'recusada: produção é a gravada pelo usuário, nada foi trocado' ((Get-HashArquivo $prodT) -ne $hCopiaT -and (Get-Item -LiteralPath $prodT).Length -eq 201)
    Conferir 'recusada: o .novo foi removido' (-not (Test-Path -LiteralPath ($prodT + '.novo')))

    # gravação que escapa no instante do Replace: o que foi parar em .anterior não é o H0, então o anterior volta
    Preparar
    $script:GanchoDepoisDaTroca = { $fs = [IO.File]::Open($prodT + '.anterior', 'Append'); $fs.WriteByte(9); $fs.Close() }
    $rT = Trocar-BancoProducao $copiaT $prodT $h0T $bkpT
    $script:GanchoDepoisDaTroca = $null
    Conferir 'gravação no instante da troca: o banco anterior é devolvido' (-not $rT.ok -and (Get-Item -LiteralPath $prodT).Length -eq 201 -and -not (Test-Path -LiteralPath ($prodT + '.anterior'))) ($rT.motivo + ' tam=' + (Get-Item -LiteralPath $prodT).Length)
} finally { $script:GanchoAntesDaTroca = $null; Remove-Item -LiteralPath $tmpT -Recurse -Force -ErrorAction SilentlyContinue }

Write-Host '--- publicação do front recuperável (spec 004, F7)'
$tmpP = Join-Path ([System.IO.Path]::GetTempPath()) ('crm-publicar-' + $PID)
try {
    $raizP = Join-Path $tmpP 'BASE'; $dadosP = Join-Path $raizP 'SISTEMA\DADOS'; $histP = Join-Path $dadosP 'execucao\historico-front'
    $origemP = Join-Path $tmpP 'novo.xlsm'
    $xlsmP = Join-Path $raizP $script:NomeFront; $frontTxtP = Join-Path $raizP 'VERSAO-FRONT.txt'; $versaoTxtP = Join-Path $dadosP 'VERSAO.txt'
    $b7P = Format-TextoB7 'PRODUCAO' '1.6' ([datetime]::new(2026, 10, 20, 8, 0, 0))
    function PrepararPublicacao([bool]$comAnterior) {
        Remove-Item -LiteralPath $tmpP -Recurse -Force -ErrorAction SilentlyContinue
        New-Item -ItemType Directory -Path $dadosP, $histP -Force | Out-Null
        [IO.File]::WriteAllBytes($origemP, [byte[]](50..120))
        [IO.File]::WriteAllText($versaoTxtP, "1.5`r`n")
        if ($comAnterior) {
            [IO.File]::WriteAllBytes($xlsmP, [byte[]](1..40))
            [IO.File]::WriteAllText($frontTxtP, "1.5`r`nAmbiente: PRODUCAO   |   Versao 1.5   |   Publicada em 01/10/2026 08:00`r`n")
        }
    }
    function Estado { return ((@($xlsmP, $frontTxtP, $versaoTxtP) | ForEach-Object { if (Test-Path -LiteralPath $_) { Get-HashArquivo $_ } else { '-' } }) -join '|') }

    PrepararPublicacao $true
    $hOrigemP = Get-HashArquivo $origemP; $hAntigoP = Get-HashArquivo $xlsmP
    $rP = Publicar-FrontNaRaiz $origemP $raizP $dadosP '1.6' $b7P ([datetime]::new(2026, 10, 20, 8, 0, 0)) $histP '1.5' ''
    Conferir 'publicação: ok e na ordem histórico, cópia, VERSAO.txt, VERSAO-FRONT.txt (por último)' ($rP.ok -and ($rP.ordem -join ',') -eq 'historico,copiar,versao,versao-front') (($rP.ordem -join ',') + ' ' + $rP.motivo)
    Conferir 'publicação: .xlsm da raiz é o novo e o anterior foi para o histórico' ((Get-HashArquivo $xlsmP) -eq $hOrigemP -and @(Get-ChildItem -LiteralPath $histP -Filter '*.xlsm' | Where-Object { (Get-HashArquivo $_.FullName) -eq $hAntigoP }).Count -eq 1)
    $linhasP = @([IO.File]::ReadAllLines($frontTxtP))
    Conferir 'publicação: VERSAO-FRONT.txt tem a versão e o texto de Início!B7' ($linhasP[0] -eq '1.6' -and $linhasP[1] -eq $b7P)
    Conferir 'publicação: VERSAO.txt = 1.6' (([IO.File]::ReadAllText($versaoTxtP)).Trim() -eq '1.6')

    foreach ($falhaP in 'historico', 'copiar', 'versao', 'versao-front') {
        PrepararPublicacao $true
        $antesP = Estado
        $rP = Publicar-FrontNaRaiz $origemP $raizP $dadosP '1.6' $b7P ([datetime]::new(2026, 10, 20, 8, 0, 0)) $histP '1.5' $falhaP
        Conferir ('falha forçada em "' + $falhaP + '": recusa e devolve .xlsm, VERSAO.txt e VERSAO-FRONT.txt (hash igual)') (-not $rP.ok -and (Estado) -eq $antesP -and @(Get-ChildItem -LiteralPath $histP -File).Count -eq 0) ($rP.motivo + ' | antes=' + $antesP + ' depois=' + (Estado))
    }
    PrepararPublicacao $false
    $rP = Publicar-FrontNaRaiz $origemP $raizP $dadosP '1.6' $b7P ([datetime]::new(2026, 10, 20, 8, 0, 0)) $histP '1.5' 'versao'
    Conferir 'primeira publicação com falha: não sobra .xlsm nem VERSAO-FRONT.txt na raiz' (-not $rP.ok -and -not (Test-Path -LiteralPath $xlsmP) -and -not (Test-Path -LiteralPath $frontTxtP) -and (([IO.File]::ReadAllText($versaoTxtP)).Trim() -eq '1.5'))
    PrepararPublicacao $true
    $trancado = [IO.File]::Open($xlsmP, 'Open', 'ReadWrite', 'None')
    try {
        $rP = Publicar-FrontNaRaiz $origemP $raizP $dadosP '1.6' $b7P ([datetime]::new(2026, 10, 20, 8, 0, 0)) $histP '1.5' ''
    } finally { $trancado.Close() }
    Conferir '.xlsm da raiz aberto por alguém: recusa sem alterar nada' (-not $rP.ok -and $rP.motivo -match 'aberto' -and ([IO.File]::ReadAllText($versaoTxtP)).Trim() -eq '1.5' -and -not (Test-Path -LiteralPath $histP -PathType Leaf)) $rP.motivo
} finally { Remove-Item -LiteralPath $tmpP -Recurse -Force -ErrorAction SilentlyContinue }

Write-Host '--- teste de uso no passo 10 (spec 004, F5)'
Conferir 'publicação roda só o teste de leitura' ((@(Get-ExecucoesTesteDeUso $false) -join ',') -eq 'SoLeitura') (@(Get-ExecucoesTesteDeUso $false) -join ',')
Conferir '-Teste roda a leitura e depois o completo' ((@(Get-ExecucoesTesteDeUso $true) -join ',') -eq 'SoLeitura,Completo') (@(Get-ExecucoesTesteDeUso $true) -join ',')

Write-Host '--- DEV-04: volta como nova versão maior e manifesto de release (ciclo 007)'
. (Join-Path $dados 'build\lib\Release.ps1')
$pac = [pscustomobject]@{ versao = '1.5'; commit = 'abc1234' }
Conferir 'volta sem motivo é recusada' ((Test-PedidoVolta '' $pac '1.6' '1.1' '1.1') -match 'motivo')
Conferir 'volta sem pacote (MANIFESTO.json) é recusada' ((Test-PedidoVolta 'regressão na ficha' $null '1.6' '1.1' '1.1') -match 'pacote')
Conferir 'volta para versão igual ou maior que a publicada é recusada' ((Test-PedidoVolta 'regressão na ficha' ([pscustomobject]@{ versao = '1.6'; commit = 'x' }) '1.6' '1.1' '1.1') -match 'anterior')
Conferir 'volta com esquema da fonte diferente do banco é recusada (incompatível)' ((Test-PedidoVolta 'regressão na ficha' ([pscustomobject]@{ versao = '1.4'; commit = 'x' }) '1.6' '1.0' '1.1') -match 'esquema')
Conferir 'volta válida passa' ((Test-PedidoVolta 'regressão na ficha' $pac '1.6' '1.1' '1.1') -eq '')
Conferir 'esquema esperado pela fonte (modConfig.VERSAO_ESQUEMA)' ((Get-EsquemaDaFonte $dados) -match '^\d+\.\d+$') (Get-EsquemaDaFonte $dados)
$raizReal = Split-Path -Parent (Split-Path -Parent $dados)
Conferir 'versão do conector lida de IA\mcp\servidor.ps1' ((Get-VersaoConectorDaFonte $raizReal) -match '^\d+\.\d+\.\d+$') (Get-VersaoConectorDaFonte $raizReal)
$tmpR = Join-Path ([System.IO.Path]::GetTempPath()) ('crm-release-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tmpR -Force | Out-Null
try {
    $xl = Join-Path $tmpR 'CRM_Zapromaq.xlsm'; [IO.File]::WriteAllText($xl, 'xlsm-falso')
    $m = New-ManifestoRelease -Versao '1.7' -Tipo 'volta' -Motivo 'regressão na ficha' -Pacote $pac -VersaoFront '1.7' -VersaoConector '1.6.0' -Esquema '1.1' -Migracoes @('001', '002', '003') -Xlsm $xl -Dados $dados -Raiz $raizReal
    Conferir 'manifesto de release: versão, tipo, motivo e conteudo_de' ($m.versao -eq '1.7' -and $m.tipo -eq 'volta' -and $m.motivo -eq 'regressão na ficha' -and $m.conteudo_de -eq '1.5') ($m | ConvertTo-Json -Compress -Depth 4)
    Conferir 'manifesto de release: commit das fontes, front, conector e esquema' ($m.commit -eq 'abc1234' -and $m.versao_front -eq '1.7' -and $m.versao_conector -eq '1.6.0' -and $m.esquema -eq '1.1')
    Conferir 'manifesto de release: migrações e hash do .xlsm' ((@($m.migracoes) -join ',') -eq '001,002,003' -and $m.xlsm_sha256 -eq (Get-HashArquivo $xl))
    Conferir 'manifesto de release: hashes dos scripts de build e operação' (@($m.scripts | Where-Object { $_.caminho -match '^SISTEMA\\DADOS\\(build|operacao)\\' }).Count -ge 5 -and @($m.scripts | Where-Object { $_.sha256 -notmatch '^[0-9a-f]{64}$' }).Count -eq 0)
    $semPacote = New-ManifestoRelease -Versao '1.7' -Tipo 'publicacao' -Motivo '' -Pacote $null -VersaoFront '1.7' -VersaoConector '1.6.0' -Esquema '1.1' -Migracoes @() -Xlsm $xl -Dados $dados -Raiz $raizReal
    Conferir 'sem pacote (desenvolvimento) o commit é declarado como desconhecido' ($semPacote.commit -match 'sem pacote' -and $null -eq $semPacote.conteudo_de)
    $arq = Save-ManifestoRelease $m $tmpR
    Conferir 'manifesto salvo em releases\release-<versão>.json' ((Split-Path -Leaf $arq) -eq 'release-1.7.json' -and ((Get-Content -LiteralPath $arq -Raw -Encoding UTF8 | ConvertFrom-Json).tipo -eq 'volta'))
} finally { Remove-Item -LiteralPath $tmpR -Recurse -Force -ErrorAction SilentlyContinue }

Write-Host ''
if ($script:falhas -eq 0) { Write-Host 'TUDO CERTO.' -ForegroundColor Green; exit 0 }
Write-Host ($script:falhas.ToString() + ' falha(s).') -ForegroundColor Red; exit 1
