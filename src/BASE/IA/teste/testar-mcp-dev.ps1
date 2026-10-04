<#
  testar-mcp-dev.ps1 - faz o papel do Claude Desktop com o MCP de desenvolvimento.

  Trabalha numa CÓPIA de SISTEMA\DADOS (pasta temporária): nada do projeto é
  alterado. Confere protocolo (legado e moderno), leitura, gravação com
  controle de versão, recusa de caminho fora de fonte\ e de arquivo gerado,
  codificação Windows-1252 do VBA e a prévia devolvendo imagem (precisa do Edge).

  -SemPrevia : pula a prévia (máquina sem Edge)
#>
param([switch]$SemPrevia)
$ErrorActionPreference = 'Stop'

$raizIA = Split-Path -Parent $PSScriptRoot
$raizBase = Split-Path -Parent $raizIA
$dadosReais = Join-Path $raizBase 'SISTEMA\DADOS'

# cópia: BASE\SISTEMA\DADOS\{VERSAO.txt, lib, build, fonte, banco, ferramentas} - o gerador acha a raiz por VERSAO.txt
$tmp = Join-Path $env:TEMP ('crm-mcp-dev-' + $PID)
$dados = Join-Path $tmp 'BASE\SISTEMA\DADOS'
New-Item -ItemType Directory -Path $dados -Force | Out-Null
Copy-Item -LiteralPath (Join-Path $dadosReais 'VERSAO.txt') -Destination $dados
foreach ($d in 'lib', 'build', 'fonte', 'banco', 'ferramentas') { Copy-Item -LiteralPath (Join-Path $dadosReais $d) -Destination $dados -Recurse }

$servidor = Join-Path $raizIA 'mcp-dev\servidor-dev.ps1'
$psi = New-Object System.Diagnostics.ProcessStartInfo
$psi.FileName = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$psi.Arguments = '-NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "' + $servidor + '" -Dados "' + $dados + '"'
$psi.UseShellExecute = $false
$psi.RedirectStandardInput = $true
$psi.RedirectStandardOutput = $true
$psi.RedirectStandardError = $true
$psi.StandardOutputEncoding = New-Object System.Text.UTF8Encoding($false)
$psi.CreateNoWindow = $true
$proc = [System.Diagnostics.Process]::Start($psi)
$leituraErro = $proc.StandardError.ReadToEndAsync()
$utf8 = New-Object System.Text.UTF8Encoding($false)

$script:falhas = 0
$script:id = 0
function Enviar($metodo, $params) {
    $script:id++
    $msg = [ordered]@{ jsonrpc = '2.0'; id = $script:id; method = $metodo }
    if ($null -ne $params) { $msg['params'] = $params }
    $bytes = $utf8.GetBytes((ConvertTo-Json -InputObject $msg -Depth 20 -Compress) + "`n")
    $proc.StandardInput.BaseStream.Write($bytes, 0, $bytes.Length)
    $proc.StandardInput.BaseStream.Flush()
    $linha = $proc.StandardOutput.ReadLine()
    if ($null -eq $linha) { throw 'O servidor fechou a saída.' }
    return ($linha | ConvertFrom-Json)
}
function Conferir([string]$nome, [bool]$ok, [string]$detalhe = '') {
    if ($ok) { Write-Host ('  OK    ' + $nome) -ForegroundColor Green }
    else { $script:falhas++; Write-Host ('  FALHA ' + $nome + '  ' + $detalhe) -ForegroundColor Red }
}
function Chamar($ferramenta, $argumentos) {
    $p = [ordered]@{ name = $ferramenta; arguments = $argumentos; _meta = [ordered]@{ 'io.modelcontextprotocol/protocolVersion' = '2026-07-28' } }
    return (Enviar 'tools/call' $p).result
}
function Dados($r) { return ($r.content[0].text | ConvertFrom-Json) }

try {
    Write-Host '--- protocolo'
    $r = Enviar 'initialize' ([ordered]@{ protocolVersion = '2025-11-25'; capabilities = @{}; clientInfo = @{ name = 'teste'; version = '1' } })
    Conferir 'initialize legado' ($r.result.serverInfo.name -eq 'crm-zapromaq-dev')
    $r = Enviar 'server/discover' ([ordered]@{ _meta = [ordered]@{ 'io.modelcontextprotocol/protocolVersion' = '2026-07-28' } })
    Conferir 'server/discover moderno' ($r.result.resultType -eq 'complete')
    $r = Enviar 'tools/list' $null
    $nomes = @($r.result.tools | ForEach-Object { $_.name })
    Conferir 'catálogo com 15 ferramentas' ($nomes.Count -eq 15) ($nomes -join ',')

    Write-Host '--- leitura'
    $r = Chamar 'listar_fontes' ([ordered]@{ pasta = 'layout\formularios' })
    $l = Dados $r
    Conferir 'listar_fontes filtra a pasta' (@($l.arquivos | Where-Object { $_.caminho -eq 'layout\formularios\frmCRM.json' }).Count -eq 1)
    Conferir 'listar_fontes esconde gerados' (@((Dados (Chamar 'listar_fontes' ([ordered]@{}))).arquivos | Where-Object { $_.caminho -like 'assets\gerado\*' }).Count -eq 0)
    $f = Dados (Chamar 'ler_fonte' ([ordered]@{ caminho = 'modulos\modConfig.bas' }))
    Conferir 'ler_fonte devolve versao' ($f.versao.Length -eq 16)

    Write-Host '--- segurança'
    $r = Chamar 'ler_fonte' ([ordered]@{ caminho = '..\build\montar-frontend.ps1' })
    Conferir 'recusa caminho com ..' ($r.isError -and $r.content[0].text -like '*inválido*')
    $r = Chamar 'ler_fonte' ([ordered]@{ caminho = 'C:\Windows\win.ini' })
    Conferir 'recusa caminho absoluto' ($r.isError)
    $r = Chamar 'gravar_fonte' ([ordered]@{ caminho = 'modulos\modTema.bas'; conteudo = 'x'; versao = 'x' })
    Conferir 'recusa arquivo gerado' ($r.isError -and $r.content[0].text -like '*GERADO*')
    $r = Chamar 'ler_fonte' ([ordered]@{ caminho = 'assets\logo_zapromaq.png' })
    Conferir 'recusa tipo não editável' ($r.isError)

    Write-Host '--- gravação'
    $novo = $f.conteudo.Replace("Option Explicit", "Option Explicit`r`n' ajuste de teste: Função")
    $r = Chamar 'gravar_fonte' ([ordered]@{ caminho = 'modulos\modConfig.bas'; conteudo = $novo; versao = 'AAAAAAAAAAAAAAAA' })
    Conferir 'conflito de versão recusado' ($r.isError -and $r.content[0].text -like 'Conflito*')
    $r = Chamar 'gravar_fonte' ([ordered]@{ caminho = 'modulos\modConfig.bas'; conteudo = $novo; versao = $f.versao; motivo = 'teste' })
    Conferir 'grava com a versão lida' (-not $r.isError) $r.content[0].text
    $bytes = [IO.File]::ReadAllBytes((Join-Path $dados 'fonte\modulos\modConfig.bas'))
    $txt = [Text.Encoding]::GetEncoding(1252).GetString($bytes)
    Conferir 'VBA gravado em Windows-1252' ($txt.Contains('Função') -and -not ($bytes -contains 0xC3))
    Conferir 'VBA gravado com CRLF' ($txt.Contains("Explicit`r`n' ajuste") -and -not ($txt -match "[^`r]`n"))
    $r = Chamar 'gravar_fonte' ([ordered]@{ caminho = 'modulos\modConfig.bas'; conteudo = 'x ' + [char]0x2192; versao = (Dados (Chamar 'ler_fonte' ([ordered]@{ caminho = 'modulos\modConfig.bas' }))).versao })
    Conferir 'recusa caractere fora do Windows-1252' ($r.isError -and $r.content[0].text -like '*ChrW*')
    $r = Chamar 'gravar_fonte' ([ordered]@{ caminho = 'layout\novo.json'; conteudo = '{ quebrado' })
    Conferir 'recusa JSON inválido' ($r.isError)
    $a = Dados (Chamar 'listar_alteracoes' ([ordered]@{}))
    Conferir 'alteração registrada' (@($a.alteracoes).Count -eq 1 -and $a.alteracoes[0].motivo -eq 'teste')
    Conferir 'cópia do anterior guardada' (@(Get-ChildItem (Join-Path $dados 'execucao\dev\copias') -Recurse -File).Count -eq 1)
    $r = Chamar 'publicar' ([ordered]@{ confirmacao = 'sim' })
    Conferir 'publicar exige PUBLICAR' ($r.isError)

    Write-Host '--- banco e conferência'
    # ConvertTo-Json desempacota array de um item só; o Desktop manda array de verdade
    function Lista([string[]]$itens) { $a = New-Object System.Collections.ArrayList; foreach ($i in $itens) { [void]$a.Add($i) }; return ,$a }
    # sem banco na cópia: o estado ainda responde, só sem as versões do banco
    $e = Dados (Chamar 'estado_sistema' ([ordered]@{}))
    Conferir 'estado_sistema traz a versão da próxima publicação' ($e.versao_proxima_publicacao -match '^\d+\.\d+$') ($e | ConvertTo-Json -Compress)
    $m = Dados (Chamar 'listar_migracoes' ([ordered]@{}))
    Conferir 'listar_migracoes vê as migrações existentes' (@($m.migracoes).Count -ge 2) (@($m.migracoes | ForEach-Object { $_.arquivo }) -join ',')
    $r = Chamar 'criar_migracao' ([ordered]@{ nome = 'teste-perigo'; descricao = 'teste'; sql = (Lista 'DROP TABLE clientes') })
    Conferir 'criar_migracao recusa SQL destrutivo sem confirmação' ($r.isError -and $r.content[0].text -like '*CONFIRMO*')
    $r = Chamar 'criar_migracao' ([ordered]@{ nome = 'campo-teste'; descricao = 'campo de teste'; tipo = 'estrutura'; sql = (Lista 'ALTER TABLE clientes ADD COLUMN teste_mcp TEXT(10)') })
    Conferir 'criar_migracao de estrutura exige esquema_para' ($r.isError -and $r.content[0].text -like '*esquema_para*')
    $d = Dados (Chamar 'criar_migracao' ([ordered]@{ nome = 'campo-teste'; descricao = "campo de teste d'angelo"; tipo = 'dados'; sql = (Lista "UPDATE clientes SET uf = 'SP' WHERE uf IS NULL") }))
    $arqNovo = Join-Path $dados ('banco\migracoes\' + $d.arquivo)
    Conferir 'criar_migracao grava o próximo número' ((Test-Path -LiteralPath $arqNovo) -and $d.arquivo -match '^\d{3}-campo-teste\.ps1$') $d.arquivo
    $mig = & $arqNovo
    Conferir 'migração criada é um script válido' ($mig.Id -eq $d.arquivo.Substring(0, 3) -and $mig.Descricao -eq "campo de teste d'angelo" -and $mig.Aplicar -is [scriptblock])
    Conferir 'migração de dados não mexe no esquema' ($null -eq $mig.EsquemaDe -and $null -eq $mig.EsquemaPara)
    Conferir 'listar_migracoes mostra a nova como pendente' (@((Dados (Chamar 'listar_migracoes' ([ordered]@{}))).migracoes | Where-Object { $_.arquivo -eq $d.arquivo -and -not $_.aplicada }).Count -eq 1)
    $c = Dados (Chamar 'ler_migracao' ([ordered]@{ arquivo = $d.arquivo }))
    Conferir 'ler_migracao devolve o conteúdo' ($c.conteudo -like "*UPDATE clientes SET uf = ''SP''*")
    $v = Dados (Chamar 'verificar_projeto' ([ordered]@{}))
    Conferir 'verificar_projeto sem problemas' ($v.sem_problemas) (($v | ConvertTo-Json -Compress -Depth 5))

    if (-not $SemPrevia) {
        Write-Host '--- prévia (Edge headless)'
        $r = Chamar 'gerar_previa' ([ordered]@{ tela = 'frmLote'; escala = 100 })
        $img = @($r.content | Where-Object { $_.type -eq 'image' })
        Conferir 'prévia devolve imagem PNG' ((-not $r.isError) -and $img.Count -eq 1 -and $img[0].mimeType -eq 'image/png' -and $img[0].data.Length -gt 1000) $r.content[0].text
        $v = Dados (Chamar 'verificar_layout' ([ordered]@{ tela = 'frmLote' }))
        Conferir 'verificar_layout sem problemas' ($v.sem_problemas) $v.relatorio
    }
} catch {
    $script:falhas++
    Write-Host ('ERRO: ' + $_.Exception.Message) -ForegroundColor Red
} finally {
    try { $proc.StandardInput.Close() } catch { }
    if (-not $proc.WaitForExit(10000)) { $proc.Kill() }
    Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host ''
if ($script:falhas -eq 0) { Write-Host 'TUDO CERTO.' -ForegroundColor Green; exit 0 }
Write-Host ($script:falhas.ToString() + ' falha(s). stderr do servidor:') -ForegroundColor Red
Write-Host $leituraErro.Result
exit 1
