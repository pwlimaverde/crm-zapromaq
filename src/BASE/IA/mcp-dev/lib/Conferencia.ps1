<#
  Conferencia.ps1 - verificar_projeto do MCP de desenvolvimento, só em PowerShell.

  Porta para PowerShell as checagens de ferramentas\verificacao\verificar.py
  (Python não é instalado nas estações):
    VBA ........ Windows-1252, sem BOM, CRLF, sem UTF-8 trocado ("Ã?"), .bas com Attribute VB_Name
    Formulários  todo controle citado no código existe no layout JSON
    Esquema .... campos de modSchema x DDL (criar-banco.ps1) x SELECT de modCRM
    PowerShell . BOM quando há acento + sintaxe pelo analisador do próprio PowerShell
    .bat ....... só ASCII e CRLF
    JSON ....... sintaxe
  NÃO porta: sintaxe VBA por parser formal (ANTLR) e a estrutura de blocos do
  conferir-vba.py - isso fica com o montar_teste, que compila no Excel e roda o autoteste.
  Só lê arquivos; não grava nada.
#>

$script:ConfProblemas = $null
function Add-Problema([string]$grupo, [string]$msg) { [void]$script:ConfProblemas.Add([ordered]@{ grupo = $grupo; msg = $msg }) }

function Get-RaizBase { return [System.IO.Path]::GetFullPath((Join-Path $script:PastaDados (Join-Path '..' '..'))) }
function Get-Rel([string]$p) { $r = (Get-RaizBase).TrimEnd('\', '/'); if ($p.StartsWith($r)) { return $p.Substring($r.Length + 1) }; return $p }
function Join-Varios([string]$base) { $p = $base; foreach ($x in $args) { $p = Join-Path $p $x }; return $p }

function Read-Cp1252([string]$p) { return ([System.Text.Encoding]::GetEncoding(1252).GetString([System.IO.File]::ReadAllBytes($p))).Replace("`r`n", "`n") }
function Read-Utf8([string]$p) { return ([System.IO.File]::ReadAllText($p, (New-Object System.Text.UTF8Encoding($false)))).TrimStart([char]0xFEFF) }

function Get-ArquivosProjeto([string]$ext) {
    $sep = [System.IO.Path]::DirectorySeparatorChar
    return @(Get-ChildItem -LiteralPath (Get-RaizBase) -Recurse -File -ErrorAction SilentlyContinue | Where-Object {
        $_.Extension -ieq $ext -and $_.FullName -notlike ('*' + $sep + 'execucao' + $sep + '*') -and $_.FullName -notlike ('*' + $sep + '.git' + $sep + '*')
    } | Sort-Object FullName | ForEach-Object { $_.FullName })
}

function Test-TemBom([byte[]]$b) { return ($b.Length -ge 3 -and $b[0] -eq 0xEF -and $b[1] -eq 0xBB -and $b[2] -eq 0xBF) }
# Bytes lidos como Latin-1: 1 byte = 1 caractere, sem perder nada.
function Get-Latin1([byte[]]$b) { return [System.Text.Encoding]::GetEncoding(28591).GetString($b) }
function Test-SoCrlf([byte[]]$b) { return -not ((Get-Latin1 $b).Replace("`r`n", '') -match "[`r`n]") }
function Test-TemNaoAscii([byte[]]$b, [int]$inicio) { return ((Get-Latin1 $b).Substring([Math]::Min($inicio, $b.Length)) -match '[^\x00-\x7F]') }

# ------------------------------------------------------------------ VBA
function Get-FontesVba {
    $saida = @()
    foreach ($par in @(@('modulos', '.bas'), @('classes', '.cls'), @('formularios', '.txt'), @('pasta-de-trabalho', '.txt'))) {
        $pasta = Join-Path $script:PastaFonte $par[0]
        if (Test-Path -LiteralPath $pasta) {
            $saida += @(Get-ChildItem -LiteralPath $pasta -File | Where-Object { $_.Extension -ieq $par[1] } | Sort-Object Name | ForEach-Object { $_.FullName })
        }
    }
    return $saida
}

function Test-CodificacaoVba([string[]]$arqs) {
    $rxUtf8 = New-Object System.Text.RegularExpressions.Regex ('Ã[' + [char]0x80 + '-' + [char]0xFF + ']')
    foreach ($p in $arqs) {
        $b = [System.IO.File]::ReadAllBytes($p)
        if (Test-TemBom $b) { Add-Problema 'VBA' ((Get-Rel $p) + ': tem BOM UTF-8 - o VBE importa em Windows-1252 e o BOM vira lixo') }
        if (-not (Test-SoCrlf $b)) { Add-Problema 'VBA' ((Get-Rel $p) + ': quebra de linha que não é CRLF') }
        $txt = [System.Text.Encoding]::GetEncoding(1252).GetString($b)
        if ($rxUtf8.IsMatch($txt)) { Add-Problema 'VBA' ((Get-Rel $p) + ': parece UTF-8 lido como Windows-1252 (sequência "Ã?")') }
        if ($p.EndsWith('.bas') -and -not $txt.StartsWith('Attribute VB_Name = "')) { Add-Problema 'VBA' ((Get-Rel $p) + ': .bas sem "Attribute VB_Name" na primeira linha') }
    }
}

# ------------------------------------------------------------------ formulários: código x layout
function Test-Formularios {
    $pasta = Join-Varios $script:PastaFonte 'layout' 'formularios'
    if (-not (Test-Path -LiteralPath $pasta)) { return }
    $rx = New-Object System.Text.RegularExpressions.Regex '\b((?:bg_|ico_)?(?:cmd|lbl|txt|cbo|chk|fra|lst|grd|ico|selo|opt)[A-Z][A-Za-z0-9]*)'
    foreach ($arq in @(Get-ChildItem -LiteralPath $pasta -File -Filter '*.json' | Sort-Object Name)) {
        try { $d = (Read-Utf8 $arq.FullName) | ConvertFrom-Json } catch { Add-Problema 'Formulários' ($arq.Name + ': JSON inválido'); continue }
        $partes = ([string]$d.codigo) -split '[\\/]'
        $cod = $script:PastaFonte; foreach ($x in $partes) { $cod = Join-Path $cod $x }
        if (-not (Test-Path -LiteralPath $cod)) { Add-Problema 'Formulários' ($arq.Name + ': código ' + $d.codigo + ' não existe'); continue }
        $nomes = New-Object 'System.Collections.Generic.HashSet[string]'
        foreach ($c in @($d.controles)) { if ($c) { [void]$nomes.Add([string]$c.nome) } }
        foreach ($k in @($d.componentes)) {
            if (-not $k) { continue }
            [void]$nomes.Add([string]$k.nome)
            if ($k.tipo -eq 'botao' -or $k.tipo -eq 'selo') { [void]$nomes.Add('bg_' + $k.nome) }
            if ($k.tipo -eq 'botao' -and $k.icone) { [void]$nomes.Add('ico_' + $k.nome) }
        }
        $vistos = New-Object 'System.Collections.Generic.HashSet[string]'
        foreach ($linha in ((Read-Cp1252 $cod) -split "`n")) {
            $l = [regex]::Replace($linha, '"[^"]*"', '""')
            $l = $l.Split([char]39)[0]
            foreach ($m in $rx.Matches($l)) {
                $n = $m.Groups[1].Value
                if (-not $nomes.Contains($n) -and $vistos.Add($n)) { Add-Problema 'Formulários' ($d.codigo + ': "' + $n + '" usado no código e ausente do layout') }
            }
        }
    }
}

# ------------------------------------------------------------------ esquema x SQL x modSchema
function Get-DdlColunas {
    $arq = Join-Varios $script:PastaDados 'banco' 'esquema' 'criar-banco.ps1'
    $tabelas = @{}
    if (-not (Test-Path -LiteralPath $arq)) { return $tabelas }
    $junto = [regex]::Replace(((Read-Utf8 $arq) -replace "`r`n", "`n"), "'\s*\+\s*\n\s*'", '')
    foreach ($m in [regex]::Matches($junto, "CREATE TABLE (\w+) \((.*?)\)'", 'Singleline')) {
        $cols = New-Object System.Collections.ArrayList; $prof = 0; $atual = New-Object System.Text.StringBuilder
        foreach ($ch in $m.Groups[2].Value.ToCharArray()) {
            if ($ch -eq '(') { $prof++ } elseif ($ch -eq ')') { $prof-- }
            if ($ch -eq ',' -and $prof -eq 0) { [void]$cols.Add($atual.ToString()); [void]$atual.Clear() } else { [void]$atual.Append($ch) }
        }
        [void]$cols.Add($atual.ToString())
        $set = New-Object 'System.Collections.Generic.HashSet[string]'
        foreach ($c in $cols) { if ($c.Trim()) { [void]$set.Add(($c.Trim() -split '\s+')[0].ToLowerInvariant()) } }
        $tabelas[$m.Groups[1].Value.ToLowerInvariant()] = $set
    }
    return $tabelas
}

function Get-BlocoFuncao([string]$txt, [string]$nome) {
    $m = [regex]::Match($txt, 'Public Function ' + $nome + '\b(.*?)\nEnd Function', 'Singleline')
    if ($m.Success) { return $m.Groups[1].Value }; return ''
}

function Get-CamposSchema([string]$txt, [string]$funcao, [string]$tabela) {
    $corpo = Get-BlocoFuncao $txt $funcao
    $m = [regex]::Match($corpo, 'Case "' + $tabela + '"(.*?)(?:\n\s*Case "|\n\s*End Select)', 'Singleline')
    if (-not $m.Success) { return @() }
    return @([regex]::Matches($m.Groups[1].Value, 'Add c, "([^"]+)"') | ForEach-Object { , ($_.Groups[1].Value -split '\|') })
}

function Get-ColunasSelect([string]$txt, [string]$funcao) {
    $corpo = Get-BlocoFuncao $txt $funcao
    $junto = (@([regex]::Matches($corpo, '"([^"]*)"') | ForEach-Object { $_.Groups[1].Value }) -join '')
    $set = New-Object 'System.Collections.Generic.HashSet[string]'
    $m = [regex]::Match($junto, 'SELECT (.*?) FROM', 'Singleline, IgnoreCase')
    if (-not $m.Success) { return $set }
    foreach ($c in ($m.Groups[1].Value -split ',')) {
        $c = $c.Trim()
        $am = [regex]::Match($c, '\bAS\s+(\w+)$', 'IgnoreCase')
        if ($am.Success) { [void]$set.Add($am.Groups[1].Value.ToLowerInvariant()) } else { [void]$set.Add(($c -split '\.')[-1].ToLowerInvariant()) }
    }
    return , $set
}

function Test-Esquema {
    $ddl = Get-DdlColunas
    if ($ddl.Count -eq 0) { Add-Problema 'Esquema' 'não consegui ler o DDL de banco\esquema\criar-banco.ps1'; return }
    $schema = Read-Cp1252 (Join-Varios $script:PastaFonte 'modulos' 'modSchema.bas')
    $crm = Read-Cp1252 (Join-Varios $script:PastaFonte 'modulos' 'modCRM.bas')
    $consultas = [ordered]@{ clientes = 'SQLClientes'; contatos = 'SQLContatos'; oportunidades = 'SQLOportunidades' }
    $calculados = @('codigo_visual', 'situacao')
    foreach ($tabela in $consultas.Keys) {
        $funcao = $consultas[$tabela]
        $sel = Get-ColunasSelect $crm $funcao
        if ($sel.Count -eq 0) { Add-Problema 'Esquema' ('modCRM.' + $funcao + ': SELECT não encontrado'); continue }
        $colsDdl = $ddl[$tabela]
        foreach ($p in (Get-CamposSchema $schema 'CamposDe' $tabela)) {
            if ($p.Count -ne 7) { Add-Problema 'Esquema' ('modSchema.CamposDe(' + $tabela + '): especificação com ' + $p.Count + ' partes: ' + ($p -join '|')); continue }
            $campo = $p[0].ToLowerInvariant()
            if ($p[6] -eq '1' -and -not ($colsDdl -and $colsDdl.Contains($campo))) { Add-Problema 'Esquema' ($tabela + '.' + $campo + ': campo gravável da ficha não existe no DDL') }
            if (-not $sel.Contains($campo)) { Add-Problema 'Esquema' ($tabela + '.' + $campo + ': campo da ficha não vem na consulta modCRM.' + $funcao) }
        }
        foreach ($fc in 'ColunasDe', 'ColunasPlanilha') {
            foreach ($p in (Get-CamposSchema $schema $fc $tabela)) {
                $campo = $p[0].ToLowerInvariant()
                if (-not $sel.Contains($campo) -and $calculados -notcontains $campo) { Add-Problema 'Esquema' ('modSchema.' + $fc + '(' + $tabela + '): coluna "' + $campo + '" não vem em modCRM.' + $funcao) }
            }
        }
    }
}

# ------------------------------------------------------------------ PowerShell, .bat, JSON
function Test-PowerShell {
    $utf8Estrito = New-Object System.Text.UTF8Encoding($false, $true)
    foreach ($p in (Get-ArquivosProjeto '.ps1')) {
        $b = [System.IO.File]::ReadAllBytes($p)
        try { [void]$utf8Estrito.GetString($b) } catch { Add-Problema 'PowerShell' ((Get-Rel $p) + ': não é UTF-8'); continue }
        $bom = Test-TemBom $b
        if (-not $bom -and (Test-TemNaoAscii $b 0)) { Add-Problema 'PowerShell' ((Get-Rel $p) + ': tem acento e não tem BOM (o Windows PowerShell 5.1 lê como ANSI)') }
        $erros = $null; $tokens = $null
        [void][System.Management.Automation.Language.Parser]::ParseInput(($utf8Estrito.GetString($b)).TrimStart([char]0xFEFF), [ref]$tokens, [ref]$erros)
        foreach ($e in @($erros)) { if ($e) { Add-Problema 'PowerShell' ((Get-Rel $p) + ' linha ' + $e.Extent.StartLineNumber + ': ' + $e.Message) } }
    }
}

function Test-Bat {
    $sep = [System.IO.Path]::DirectorySeparatorChar
    foreach ($p in (Get-ArquivosProjeto '.bat')) {
        if ($p -like ('*' + $sep + 'migracao-3.6' + $sep + '*')) { continue }   # registro histórico, não é executado
        $b = [System.IO.File]::ReadAllBytes($p)
        if (Test-TemNaoAscii $b 0) { Add-Problema '.bat' ((Get-Rel $p) + ': tem caractere fora do ASCII (o cmd.exe usa a página OEM)') }
        if (-not (Test-SoCrlf $b)) { Add-Problema '.bat' ((Get-Rel $p) + ': quebra de linha que não é CRLF (o cmd.exe erra rótulos)') }
    }
}

function Test-Json {
    foreach ($p in (Get-ArquivosProjeto '.json')) {
        try { [void]((Read-Utf8 $p) | ConvertFrom-Json) } catch { Add-Problema 'JSON' ((Get-Rel $p) + ': ' + $_.Exception.Message) }
    }
}

# ------------------------------------------------------------------ ferramenta
function Invoke-VerificarProjeto($a) {
    $script:ConfProblemas = New-Object System.Collections.ArrayList
    $vba = @(Get-FontesVba)
    foreach ($etapa in 'Test-Formularios', 'Test-Esquema', 'Test-PowerShell', 'Test-Bat', 'Test-Json') {
        try { & $etapa } catch { Add-Problema 'Conferência' ($etapa + ' falhou: ' + $_.Exception.Message) }
    }
    try { Test-CodificacaoVba $vba } catch { Add-Problema 'Conferência' ('Test-CodificacaoVba falhou: ' + $_.Exception.Message) }
    return [ordered]@{
        ferramenta = 'conferência em PowerShell'
        fontes_vba = $vba.Count
        sem_problemas = ($script:ConfProblemas.Count -eq 0)
        problemas = @($script:ConfProblemas | ForEach-Object { '[' + $_.grupo + '] ' + $_.msg })
        observacao = 'Não cobre a sintaxe VBA por parser formal nem a estrutura de blocos: o montar_teste compila no Excel e roda o autoteste.'
    }
}
