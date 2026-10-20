<#
  Release.ps1 - volta como nova versão e manifesto de release (DEV-04, ciclo 007).

  Volta = PUBLICAR DE NOVO o conteúdo de uma versão anterior, com número MAIOR e motivo registrado: o
  INICIAR-CRM das estações só troca para versão maior, então "copiar o .xlsm antigo de volta" nunca chegava
  às estações. Roteiro: extrair POR CIMA da BASE o pacote da versão boa (dist\BASE-vX.Y-*.zip, que traz o
  MANIFESTO.json com a versão e o commit) e rodar  MONTAR-FRONTEND.bat -Volta -Motivo "<por que>".

  Manifesto de release: um arquivo por publicação em execucao\releases\release-<versão>.json com o commit
  das fontes, as versões do front, do conector e do esquema, as migrações aplicadas, o hash do .xlsm e dos
  scripts de build\, operacao\ e banco\. Identifica o código que produziu o .xlsm publicado.
#>

# Esquema que o front montado destas fontes espera (modConfig.VERSAO_ESQUEMA).
function Get-EsquemaDaFonte([string]$dados) {
    $txt = [System.IO.File]::ReadAllText((Join-Path $dados 'fonte\modulos\modConfig.bas'), [System.Text.Encoding]::GetEncoding(1252))
    if ($txt -match '(?m)^Public Const VERSAO_ESQUEMA As String = "([^"]*)"') { return $Matches[1] }
    throw 'modConfig.bas: constante VERSAO_ESQUEMA não encontrada'
}

# Versão do conector nas fontes (IA\mcp\servidor.ps1, $script:VersaoServidor).
function Get-VersaoConectorDaFonte([string]$raiz) {
    $arq = Join-Path $raiz 'IA\mcp\servidor.ps1'
    if (-not (Test-Path -LiteralPath $arq)) { return '' }
    $txt = [System.IO.File]::ReadAllText($arq)
    if ($txt -match "(?m)^\`$script:VersaoServidor = '([^']+)'") { return $Matches[1] }
    return ''
}

# MANIFESTO.json do pacote extraído na BASE (empacotar-base.ps1), ou $null (pasta de desenvolvimento).
function Get-ManifestoPacote([string]$raiz) {
    $man = Join-Path $raiz 'SISTEMA\DADOS\MANIFESTO.json'
    if (-not (Test-Path -LiteralPath $man)) { return $null }
    return (Get-Content -LiteralPath $man -Raw -Encoding UTF8 | ConvertFrom-Json)
}

# Motivo para recusar a volta ('' = aceita).
function Test-PedidoVolta([string]$motivo, $pacote, [string]$versaoPublicada, [string]$esquemaFonte, [string]$esquemaBanco) {
    if ([string]::IsNullOrWhiteSpace($motivo)) { return 'a volta exige -Motivo "<por que>" (fica registrado no manifesto e no CHANGELOG). Nada foi publicado.' }
    if ($null -eq $pacote -or -not $pacote.versao) { return 'a volta exige o pacote da versão boa extraído por cima da BASE (SISTEMA\DADOS\MANIFESTO.json ausente). Nada foi publicado.' }
    if ($versaoPublicada -and (Compare-Versao ([string]$pacote.versao) $versaoPublicada) -ge 0) {
        return ('o pacote extraído é da versão ' + $pacote.versao + ', que não é anterior à publicada (' + $versaoPublicada + '). Nada foi publicado.')
    }
    if ($esquemaFonte -ne $esquemaBanco) {
        return ('volta incompatível: o front da versão ' + $pacote.versao + ' espera o esquema ' + $esquemaFonte + ' e o banco está no ' + $esquemaBanco +
                ' (banco não volta de esquema sobre gravações novas). Nada foi publicado.')
    }
    return ''
}

function New-ManifestoRelease([string]$Versao, [string]$Tipo, [string]$Motivo, $Pacote, [string]$VersaoFront,
                              [string]$VersaoConector, [string]$Esquema, [string[]]$Migracoes, [string]$Xlsm,
                              [string]$Dados, [string]$Raiz) {
    $scripts = New-Object System.Collections.ArrayList
    foreach ($sub in @('build', 'operacao', 'banco', 'lib')) {
        $p = Join-Path $Dados $sub
        if (-not (Test-Path -LiteralPath $p)) { continue }
        foreach ($f in @(Get-ChildItem -LiteralPath $p -Recurse -File -Include '*.ps1', '*.bat' | Sort-Object FullName)) {
            [void]$scripts.Add([ordered]@{ caminho = $f.FullName.Substring($Raiz.TrimEnd('\').Length + 1); sha256 = (Get-HashArquivo $f.FullName) })
        }
    }
    $pointer = Join-Path $Raiz 'IA\conector\ATUAL.txt'
    $conPublicado = if (Test-Path -LiteralPath $pointer) { ([System.IO.File]::ReadAllText($pointer)).Trim() } else { '' }
    return [ordered]@{
        versao_manifesto = 1
        versao           = $Versao
        tipo             = $Tipo
        motivo           = $Motivo
        conteudo_de      = $(if ($Tipo -eq 'volta' -and $Pacote) { [string]$Pacote.versao } else { $null })
        instante         = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss zzz')
        por              = ($env:COMPUTERNAME + '\' + $env:USERNAME)
        commit           = $(if ($Pacote -and $Pacote.commit) { [string]$Pacote.commit } else { 'sem pacote (montado da pasta de desenvolvimento)' })
        versao_pacote    = $(if ($Pacote) { [string]$Pacote.versao } else { $null })
        versao_front     = $VersaoFront
        versao_conector  = $VersaoConector
        conector_publicado = $conPublicado
        esquema          = $Esquema
        migracoes        = @($Migracoes)
        xlsm_sha256      = $(if ($Xlsm -and (Test-Path -LiteralPath $Xlsm)) { Get-HashArquivo $Xlsm } else { $null })
        scripts          = @($scripts)
    }
}

function Save-ManifestoRelease($manifesto, [string]$pasta) {
    New-Item -ItemType Directory -Path $pasta -Force | Out-Null
    $arq = Join-Path $pasta ('release-' + $manifesto.versao + '.json')
    [System.IO.File]::WriteAllText($arq, (ConvertTo-Json -InputObject $manifesto -Depth 5), (New-Object System.Text.UTF8Encoding($false)))
    return $arq
}
