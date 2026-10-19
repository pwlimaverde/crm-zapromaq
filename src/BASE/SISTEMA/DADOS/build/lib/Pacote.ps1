<#
  Pacote.ps1 - lê o .xlsm pronto SEM abrir o Excel (é um zip Open XML).

  É a conferência independente antes de publicar: o arquivo que vai para a
  raiz tem de ter o projeto VBA, ser do tipo habilitado para macro e trazer
  Início!B7 no padrão do contrato. Se o Excel mentiu, o zip não mente.
#>
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Read-EntradaZip($zip, [string]$nome) {
    $e = $zip.GetEntry($nome)
    if ($null -eq $e) { return $null }
    $r = New-Object System.IO.StreamReader($e.Open(), [System.Text.Encoding]::UTF8)
    try { return $r.ReadToEnd() } finally { $r.Dispose() }
}

# Texto de uma célula de uma aba, pelo nome da aba (ex.: 'Inicio', 'B7').
function Get-TextoCelulaXlsx([string]$arquivo, [string]$aba, [string]$celula) {
    $zip = [System.IO.Compression.ZipFile]::OpenRead($arquivo)
    try {
        [xml]$wb = Read-EntradaZip $zip 'xl/workbook.xml'
        [xml]$rels = Read-EntradaZip $zip 'xl/_rels/workbook.xml.rels'
        $ns = New-Object System.Xml.XmlNamespaceManager($wb.NameTable)
        $ns.AddNamespace('m', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
        $ns.AddNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        $sheet = $wb.SelectSingleNode("//m:sheets/m:sheet[@name='" + $aba + "']", $ns)
        if ($null -eq $sheet) { throw ('aba "' + $aba + '" não existe no pacote') }
        $rid = $sheet.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        $alvo = ($rels.Relationships.Relationship | Where-Object { $_.Id -eq $rid }).Target
        $caminho = if ($alvo.StartsWith('/')) { $alvo.TrimStart('/') } else { 'xl/' + $alvo }
        [xml]$ws = Read-EntradaZip $zip $caminho
        $nsw = New-Object System.Xml.XmlNamespaceManager($ws.NameTable)
        $nsw.AddNamespace('m', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
        $c = $ws.SelectSingleNode("//m:c[@r='" + $celula + "']", $nsw)
        if ($null -eq $c) { return $null }
        $t = $c.GetAttribute('t')
        if ($t -eq 'inlineStr') { return $c.SelectSingleNode('m:is', $nsw).InnerText }
        $v = $c.SelectSingleNode('m:v', $nsw)
        if ($null -eq $v) { return $null }
        if ($t -eq 's') {
            [xml]$ss = Read-EntradaZip $zip 'xl/sharedStrings.xml'
            $nss = New-Object System.Xml.XmlNamespaceManager($ss.NameTable)
            $nss.AddNamespace('m', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
            return $ss.SelectNodes('//m:si', $nss)[[int]$v.InnerText].InnerText
        }
        return $v.InnerText
    } finally { $zip.Dispose() }
}

# Lista de problemas do pacote (vazia = ok).
function Test-PacoteXlsm([string]$arquivo, [string[]]$abasEsperadas) {
    $p = New-Object System.Collections.ArrayList
    if (-not (Test-Path -LiteralPath $arquivo)) { [void]$p.Add('arquivo não existe: ' + $arquivo); return ,$p }
    $zip = [System.IO.Compression.ZipFile]::OpenRead($arquivo)
    try {
        if ($null -eq $zip.GetEntry('xl/vbaProject.bin')) { [void]$p.Add('sem projeto VBA (xl/vbaProject.bin)') }
        $ct = Read-EntradaZip $zip '[Content_Types].xml'
        if ($ct -notmatch 'sheet\.macroEnabled\.main\+xml') { [void]$p.Add('tipo de conteúdo não é pasta habilitada para macro') }
        [xml]$wb = Read-EntradaZip $zip 'xl/workbook.xml'
        $nomes = @($wb.workbook.sheets.sheet | ForEach-Object { $_.name })
        foreach ($a in $abasEsperadas) { if ($nomes -notcontains $a) { [void]$p.Add('falta a aba ' + $a) } }
    } finally { $zip.Dispose() }
    return ,$p
}

# Entradas do gerador visual cujo hash difere do manifesto (vazio = em dia).
# Arquivo de layout novo, ainda fora do manifesto, também conta como desatualizado.
function Get-VisualDesatualizado([string]$dados) {
    $man = Join-Path $dados 'fonte\assets\gerado\manifesto.json'
    if (-not (Test-Path -LiteralPath $man)) { return @('manifesto ausente') }
    $m = Get-Content -LiteralPath $man -Raw -Encoding UTF8 | ConvertFrom-Json
    $saida = New-Object System.Collections.ArrayList
    foreach ($e in $m.entradas.PSObject.Properties) {
        $arq = Join-Path $dados $e.Name
        if (-not (Test-Path -LiteralPath $arq) -or (Get-FileHash -LiteralPath $arq -Algorithm SHA256).Hash -ne $e.Value) { [void]$saida.Add($e.Name) }
    }
    foreach ($f in @(Get-ChildItem -LiteralPath (Join-Path $dados 'fonte\layout') -Filter '*.json' -Recurse)) {
        $rel = $f.FullName.Substring($dados.TrimEnd('\').Length + 1)
        if (-not $m.entradas.PSObject.Properties[$rel]) { [void]$saida.Add($rel) }
    }
    return $saida.ToArray()
}

# ------------------------------------------------------------------ backup antes de publicar
# Cópia do banco e do .xlsm ANTES de qualquer alteração, em
# execucao\backup\antes-da-publicacao\<data-hora>. É o que permite voltar
# atrás se a versão nova tiver problema. Devolve a pasta criada.
function New-BackupAntesDePublicar([string]$raiz, [string]$banco, [string]$front, $log) {
    $pasta = Join-Path (Get-PastaExecucao $raiz 'backup') ('antes-da-publicacao\' + (Get-Date -Format 'yyyyMMdd-HHmmss'))
    New-Item -ItemType Directory -Path $pasta -Force | Out-Null

    if (Test-Path -LiteralPath $front) {
        Copy-Item -LiteralPath $front -Destination $pasta -Force
        $log.L('  front ..: ' + (Split-Path -Leaf $front) + '  (' + [math]::Round((Get-Item -LiteralPath $front).Length / 1MB, 1) + ' MB)')
    } else {
        $log.L('  front ..: não existe ainda (primeira publicação)')
    }

    if (Test-Path -LiteralPath $banco) {
        if (Test-Path -LiteralPath ([System.IO.Path]::ChangeExtension($banco, 'laccdb'))) {
            $log.L('  AVISO: há usuário conectado ao banco (.laccdb). A cópia sai, mas pode não estar consistente.')
        }
        $destino = Join-Path $pasta (Split-Path -Leaf $banco)
        Copy-Item -LiteralPath $banco -Destination $destino -Force
        $tam = (Get-Item -LiteralPath $destino).Length
        if ($tam -ne (Get-Item -LiteralPath $banco).Length) {
            Remove-Item -LiteralPath $destino -Force
            throw 'a cópia do banco saiu com tamanho diferente do original; nada foi alterado'
        }
        $log.L('  banco ..: ' + (Split-Path -Leaf $banco) + '  (' + [math]::Round($tam / 1MB, 1) + ' MB)')
    } else {
        $log.L('  banco ..: não encontrado')
    }
    return $pasta
}

# Passo 12 recuperável (spec 004, F7). Publica o front na raiz de BASE na ordem:
#   1 historico  : .xlsm anterior -> pasta do histórico (só se a raiz o libera: ninguém com ele aberto)
#   2 copiar     : .xlsm novo -> raiz, com conferência de hash
#   3 versao     : VERSAO.txt
#   4 versao-front: VERSAO-FRONT.txt, SEMPRE por último - é o ponto de não retorno: o INICIAR-CRM das
#                  estações decide por ele, então nunca vê número novo apontando para arquivo velho.
# Falha em qualquer subpasso devolve tudo (o .xlsm anterior, VERSAO.txt e VERSAO-FRONT.txt) e o resultado
# diz o que aconteceu. $falharEm só serve aos testes: simula a falha nesse subpasso.
function Publicar-FrontNaRaiz([string]$origem, [string]$raiz, [string]$dados, [string]$versao, [string]$textoB7,
                              [datetime]$quando, [string]$pastaHistorico, [string]$versaoAtual, [string]$falharEm = '') {
    $destino = Join-Path $raiz $script:NomeFront
    $arqFront = Join-Path $raiz 'VERSAO-FRONT.txt'
    $arqTxt = Join-Path $dados 'VERSAO.txt'
    $bytesTxt = if (Test-Path -LiteralPath $arqTxt) { [System.IO.File]::ReadAllBytes($arqTxt) } else { $null }
    $bytesFront = if (Test-Path -LiteralPath $arqFront) { [System.IO.File]::ReadAllBytes($arqFront) } else { $null }
    $ordem = New-Object System.Collections.ArrayList
    $historico = ''; $copiou = $false
    try {
        if (Test-Path -LiteralPath $destino) {
            try { $f = [System.IO.File]::Open($destino, 'Open', 'ReadWrite', 'None'); $f.Close() }
            catch { throw 'o .xlsm da raiz foi aberto por alguém durante o build. Nada foi publicado; rode de novo.' }
            if ($falharEm -eq 'historico') { throw 'falha simulada: historico' }
            New-Item -ItemType Directory -Path $pastaHistorico -Force | Out-Null
            $historico = Join-Path $pastaHistorico ('CRM_Zapromaq_v' + $versaoAtual + '_' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.xlsm')
            Move-Item -LiteralPath $destino -Destination $historico
            [void]$ordem.Add('historico')
        } elseif ($falharEm -eq 'historico') { throw 'falha simulada: historico' }
        else { [void]$ordem.Add('historico') }

        if ($falharEm -eq 'copiar') { throw 'falha simulada: copiar' }
        Copy-Item -LiteralPath $origem -Destination $destino
        $copiou = $true
        if ((Get-FileHash -LiteralPath $origem -Algorithm SHA256).Hash -ne (Get-FileHash -LiteralPath $destino -Algorithm SHA256).Hash) {
            throw 'a cópia publicada não confere com a montada (hash diferente)'
        }
        [void]$ordem.Add('copiar')

        if ($falharEm -eq 'versao') { throw 'falha simulada: versao' }
        Set-VersaoArquivo $dados $versao
        [void]$ordem.Add('versao')

        # ANSI + CRLF: é o cmd.exe do .bat que lê.
        $conteudo = $versao + "`r`n" + $textoB7 + "`r`n" +
                    'arquivo=' + $script:NomeFront + "`r`n" +
                    'publicado=' + $quando.ToString('yyyy-MM-dd HH:mm:ss') + "`r`n" +
                    'por=' + $env:COMPUTERNAME + '\' + $env:USERNAME + "`r`n"
        [System.IO.File]::WriteAllText($arqFront, $conteudo, (New-Object System.Text.ASCIIEncoding))
        if ($falharEm -eq 'versao-front') { throw 'falha simulada: versao-front' }
        [void]$ordem.Add('versao-front')
    } catch {
        $motivo = $_.Exception.Message
        $erros = New-Object System.Collections.ArrayList
        # devolve na ordem inversa; cada passo da devolução é independente dos outros
        try { if ($null -ne $bytesFront) { [System.IO.File]::WriteAllBytes($arqFront, $bytesFront) } elseif (Test-Path -LiteralPath $arqFront) { Remove-Item -LiteralPath $arqFront -Force } } catch { [void]$erros.Add('VERSAO-FRONT.txt: ' + $_.Exception.Message) }
        try { if ($null -ne $bytesTxt) { [System.IO.File]::WriteAllBytes($arqTxt, $bytesTxt) } } catch { [void]$erros.Add('VERSAO.txt: ' + $_.Exception.Message) }
        try { if ($copiou -and (Test-Path -LiteralPath $destino)) { Remove-Item -LiteralPath $destino -Force } } catch { [void]$erros.Add('.xlsm novo: ' + $_.Exception.Message) }
        try { if ($historico -and (Test-Path -LiteralPath $historico) -and -not (Test-Path -LiteralPath $destino)) { Move-Item -LiteralPath $historico -Destination $destino } } catch { [void]$erros.Add('.xlsm anterior: ' + $_.Exception.Message) }
        $extra = if ($erros.Count) { ' ATENÇÃO, a devolução falhou em: ' + ($erros -join '; ') } else { ' O .xlsm anterior, VERSAO.txt e VERSAO-FRONT.txt foram devolvidos.' }
        return [pscustomobject]@{ ok = $false; motivo = ($motivo + '.' + $extra); ordem = @($ordem); devolvido = ($erros.Count -eq 0) }
    }
    return [pscustomobject]@{ ok = $true; motivo = 'front publicado'; ordem = @($ordem); devolvido = $false }
}
