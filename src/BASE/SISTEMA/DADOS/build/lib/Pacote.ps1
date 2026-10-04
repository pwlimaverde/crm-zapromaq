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
