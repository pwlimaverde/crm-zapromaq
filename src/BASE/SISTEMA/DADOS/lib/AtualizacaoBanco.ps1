<#
  AtualizacaoBanco.ps1 - atualização do banco de produção pela montagem (spec 003, passos 3 a 8 e 11).

  Tudo ou nada: a migração e as conferências rodam numa CÓPIA; o banco de produção só é trocado no
  fim, e só se a trava estiver livre e o hash dele ainda for o H0 (o do backup). Qualquer falha antes
  da troca deixa o banco de produção exatamente como estava.

  Depende de Comum.ps1 (Open-Banco, Invoke-Escalar, Get-ColunasDaLista, Get-ChaveLista) carregado antes.
  Não publica o front: isso é o passo 12 da montagem.
#>

function Get-HashArquivo([string]$arquivo) {
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $fs = [System.IO.File]::OpenRead($arquivo)
    try { return ([BitConverter]::ToString($sha.ComputeHash($fs))).Replace('-', '').ToLowerInvariant() }
    finally { $fs.Dispose(); $sha.Dispose() }
}

# Access grava um registro de 64 bytes por usuário conectado (32 de estação, 32 de usuário).
# O nome da estação só orienta quem vai fechar o CRM; a trava é o que bloqueia.
function Get-EstacoesDaTrava([string]$laccdb) {
    $estacoes = New-Object System.Collections.ArrayList
    if (-not (Test-Path -LiteralPath $laccdb)) { return , $estacoes }
    $bytes = [System.IO.File]::ReadAllBytes($laccdb)
    for ($i = 0; $i + 64 -le $bytes.Length; $i += 64) {
        $nome = [Text.Encoding]::ASCII.GetString($bytes, $i, 32).Trim([char]0, ' ')
        if ($nome) { [void]$estacoes.Add($nome) }
    }
    return , $estacoes
}

function Test-BancoLivre([string]$banco) {
    $laccdb = [System.IO.Path]::ChangeExtension($banco, 'laccdb')
    $estacoes = Get-EstacoesDaTrava $laccdb
    return [pscustomobject]@{ livre = -not (Test-Path -LiteralPath $laccdb); estacoes = @($estacoes); laccdb = $laccdb }
}

# Passo 1: o CSV de municípios bate com o SHA-256 registrado no .fonte.json (o da migração 003 é conferido lá).
# Devolve o motivo; vazio = pacote aprovado.
function Test-PacoteMunicipios([string]$dados) {
    $ref = Join-Path $dados 'banco\referencia'
    $csv = Join-Path $ref 'municipios.csv'
    $fonte = Join-Path $ref 'municipios.fonte.json'
    if (-not (Test-Path -LiteralPath $csv) -or -not (Test-Path -LiteralPath $fonte)) { return 'municipios.csv ou municipios.fonte.json ausente' }
    $registrado = [string]((Get-Content -LiteralPath $fonte -Raw -Encoding UTF8 | ConvertFrom-Json).csv_sha256)
    if ((Get-HashArquivo $csv) -ne $registrado) { return 'municipios.csv não confere com o SHA-256 do municipios.fonte.json' }
    return ''
}

# Passo 1 (pacote): cada arquivo do MANIFESTO.json tem de existir com o SHA-256 gravado. O manifesto
# só existe no pacote da empresa (empacotar-base.ps1); na pasta de desenvolvimento não há, e não se confere.
# Devolve o motivo; vazio = pacote aprovado.
function Test-Manifesto([string]$pastaBase) {
    $man = Join-Path $pastaBase 'SISTEMA\DADOS\MANIFESTO.json'
    if (-not (Test-Path -LiteralPath $man)) { return '' }
    $lista = Get-Content -LiteralPath $man -Raw -Encoding UTF8 | ConvertFrom-Json
    foreach ($a in @($lista.arquivos)) {
        $arq = Join-Path $pastaBase ([string]$a.caminho -replace '/', '\')
        if (-not (Test-Path -LiteralPath $arq)) { return 'arquivo do pacote ausente: ' + $a.caminho }
        if ((Get-HashArquivo $arq) -ne [string]$a.sha256) { return 'arquivo do pacote alterado depois de gerado: ' + $a.caminho }
    }
    return ''
}

# -Teste só roda sobre banco de TESTE: nunca sobre o de produção. Devolve o motivo; vazio = aceito.
function Test-ModoTeste([string]$ambiente) {
    if ($ambiente -eq 'TESTE') { return '' }
    return ('o -Teste só roda sobre banco com config.ambiente = TESTE (este banco está em ' + $ambiente + '). Nada foi alterado.')
}

function New-ResultadoBanco([bool]$ok, [int]$passo, [string]$motivo, $extra = @{}) {
    $r = [ordered]@{ ok = $ok; passo = $passo; motivo = $motivo }
    foreach ($k in $extra.Keys) { $r[$k] = $extra[$k] }
    return [pscustomobject]$r
}

# Conferências da CÓPIA migrada (passo 8). Devolve a lista de problemas; vazia = aprovada.
function Get-ProblemasDoBanco($cn, [int]$totalMunicipios) {
    $problemas = New-Object System.Collections.ArrayList
    $versao = [string](Invoke-Escalar $cn "SELECT valor FROM config WHERE chave='versao_esquema'" @() $null)
    if ($versao -ne '1.1') { [void]$problemas.Add('esquema ' + $versao + ' (esperado 1.1)') }
    $municipios = [int](Invoke-Escalar $cn 'SELECT COUNT(*) FROM municipios' @() $null)
    if ($municipios -ne $totalMunicipios) { [void]$problemas.Add('municípios: ' + $municipios + ' (esperado ' + $totalMunicipios + ')') }
    if ([int](Invoke-Escalar $cn 'SELECT COUNT(*) FROM listas WHERE chave IS NULL' @() $null) -ne 0) { [void]$problemas.Add('há linhas de lista sem chave') }
    if ([int](Invoke-Escalar $cn "SELECT COUNT(*) FROM listas WHERE tipo='Etapa' AND quadro IS NULL" @() $null) -ne 0) { [void]$problemas.Add('há etapas sem papel') }
    $schema = $cn.GetSchema('Indexes')
    $temIndice = $false
    foreach ($r in $schema.Rows) { if ([string]$r['INDEX_NAME'] -eq 'ix_lis_tipo_chave') { $temIndice = $true } }
    if (-not $temIndice) { [void]$problemas.Add('índice único ix_lis_tipo_chave ausente') }

    # nenhum valor em uso pode ficar sem linha na lista (pela chave)
    $linhas = New-ComandoBanco $cn 'SELECT tipo, chave FROM listas' @() $null
    $rd = $linhas.ExecuteReader()
    $existe = @{}
    while ($rd.Read()) { $existe[[string]$rd.GetValue(0) + '|' + [string]$rd.GetValue(1)] = $true }
    $rd.Close(); $linhas.Dispose()
    foreach ($tipo in (Get-ColunasDaLista).Keys) {
        foreach ($tc in (Get-ColunasDaLista)[$tipo]) {
            $partes = $tc.Split(':')
            $c = New-ComandoBanco $cn ('SELECT DISTINCT ' + $partes[1] + ' FROM ' + $partes[0] + ' WHERE ' + $partes[1] + ' IS NOT NULL') @() $null
            $rd = $c.ExecuteReader()
            while ($rd.Read()) {
                $v = [string]$rd.GetValue(0)
                if ($v -ne '' -and -not $existe.ContainsKey($tipo + '|' + (Get-ChaveLista $v))) {
                    [void]$problemas.Add('valor em uso sem linha na lista: ' + $tipo + ' "' + $v + '"')
                }
            }
            $rd.Close(); $c.Dispose()
        }
    }
    return , $problemas
}

# Passos 3 a 8 da montagem: trava, backup com H0, cópia, migração na cópia, conferências.
# Não toca no banco de produção além de ler e copiar.
function Invoke-PreparacaoDoBanco([string]$producao, [string]$pastaTrabalho, [string]$pastaBackup, [int]$totalMunicipios, $log) {
    $dados = Split-Path -Parent $PSScriptRoot

    $livre = Test-BancoLivre $producao
    if (-not $livre.livre) {
        $quem = if ($livre.estacoes.Count) { ' Estações: ' + ($livre.estacoes -join ', ') + '.' } else { '' }
        return New-ResultadoBanco $false 3 ('há usuário conectado ao banco (.laccdb).' + $quem + ' Peça para fecharem o CRM e rode de novo. Nada foi alterado.')
    }
    $log.L('  trava: livre')

    New-Item -ItemType Directory -Path $pastaBackup -Force | Out-Null
    $backup = Join-Path $pastaBackup ([IO.Path]::GetFileName($producao))
    Copy-Item -LiteralPath $producao -Destination $backup -Force
    $h0 = Get-HashArquivo $producao
    if ((Get-HashArquivo $backup) -ne $h0) { return New-ResultadoBanco $false 5 'o backup não confere com o banco de produção. Nada foi alterado.' }
    $log.L('  backup: ' + $backup + '  H0 ' + $h0.Substring(0, 12))

    New-Item -ItemType Directory -Path $pastaTrabalho -Force | Out-Null
    $copia = Join-Path $pastaTrabalho ([IO.Path]::GetFileName($producao))
    Copy-Item -LiteralPath $producao -Destination $copia -Force
    if ((Get-HashArquivo $copia) -ne $h0) { return New-ResultadoBanco $false 6 'a cópia de trabalho não confere com o banco de produção. Nada foi alterado.' }
    $log.L('  cópia de trabalho: ' + $copia)

    $migrar = Join-Path $dados 'banco\aplicar-migracoes.ps1'
    # stderr do processo filho não pode virar exceção aqui: a falha é tratada pelo código de saída
    $preferencia = $ErrorActionPreference; $ErrorActionPreference = 'Continue'
    try { $saidaMig = @(& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $migrar -Banco $copia 2>&1 | ForEach-Object { [string]$_ }) }
    finally { $ErrorActionPreference = $preferencia }
    foreach ($linha in $saidaMig) { if ($linha -ne '') { $log.L('  ' + $linha) } }
    if ($LASTEXITCODE -ne 0) {
        return New-ResultadoBanco $false 7 ('a migração falhou na cópia (veja o log de migrações). Banco de produção intacto.') @{ copia = $copia; h0 = $h0; backup = $backup }
    }
    $log.L('  migrações aplicadas na cópia')

    $cn = Open-Banco $copia
    try { $problemas = Get-ProblemasDoBanco $cn $totalMunicipios } finally { $cn.Close(); $cn.Dispose() }
    if ($problemas.Count -gt 0) {
        return New-ResultadoBanco $false 8 ('conferência da cópia: ' + ($problemas -join '; ') + '. Banco de produção intacto.') @{ copia = $copia; h0 = $h0; backup = $backup }
    }
    $log.L('  conferências da cópia: aprovadas')
    return New-ResultadoBanco $true 8 'cópia pronta' @{ copia = $copia; h0 = $h0; backup = $backup; migracao = $saidaMig }
}

# Passo 11: troca o banco de produção pela cópia, só com a trava livre e o hash igual ao H0 (C19, C20).
# Antes de trocar, guarda o banco anterior em $pastaBackup (para reverter).
function Trocar-BancoProducao([string]$copia, [string]$producao, [string]$h0, [string]$pastaBackup) {
    $livre = Test-BancoLivre $producao
    if (-not $livre.livre) { return New-ResultadoBanco $false 11 'há usuário conectado ao banco (.laccdb); nada foi trocado. Rode de novo quando fecharem o CRM.' }
    $atual = Get-HashArquivo $producao
    if ($atual -ne $h0) { return New-ResultadoBanco $false 11 'alguém gravou no banco durante a montagem (o hash mudou); nada foi trocado. Rode de novo.' }

    New-Item -ItemType Directory -Path $pastaBackup -Force | Out-Null
    Copy-Item -LiteralPath $producao -Destination (Join-Path $pastaBackup ([IO.Path]::GetFileName($producao))) -Force

    # a cópia novo fica na MESMA pasta do banco de produção: a troca é um renomear dentro do volume
    $novo = $producao + '.novo'
    Copy-Item -LiteralPath $copia -Destination $novo -Force
    if ((Get-HashArquivo $novo) -ne (Get-HashArquivo $copia)) {
        Remove-Item -LiteralPath $novo -Force -ErrorAction SilentlyContinue
        return New-ResultadoBanco $false 11 'a cópia do novo banco não conferiu na pasta de produção; nada foi trocado.'
    }
    $anterior = $producao + '.anterior'   # o Replace exige um nome de backup real, na mesma pasta
    try { [System.IO.File]::Replace($novo, $producao, $anterior) }
    catch {
        Remove-Item -LiteralPath $novo -Force -ErrorAction SilentlyContinue
        return New-ResultadoBanco $false 11 ('não foi possível trocar o arquivo de produção (' + $_.Exception.Message + '); nada foi trocado.')
    }
    Remove-Item -LiteralPath $anterior -Force -ErrorAction SilentlyContinue
    if ((Get-HashArquivo $producao) -ne (Get-HashArquivo $copia)) {
        return New-ResultadoBanco $false 11 ('a produção não ficou igual à cópia depois da troca; use o backup em ' + $pastaBackup)
    }
    return New-ResultadoBanco $true 11 'banco de produção trocado pela cópia migrada'
}

# Volta ao banco anterior, mas só se a produção ainda é a que a troca deixou (hash $hashNovo).
function Reverter-Troca([string]$producao, [string]$backup, [string]$hashNovo) {
    if ((Get-HashArquivo $producao) -ne $hashNovo) {
        return New-ResultadoBanco $false 12 'o banco de produção mudou desde a troca; não revertido. Avise antes de mexer.'
    }
    $novo = $producao + '.novo'
    Copy-Item -LiteralPath $backup -Destination $novo -Force
    $anterior = $producao + '.anterior'
    try { [System.IO.File]::Replace($novo, $producao, $anterior) }
    catch {
        Remove-Item -LiteralPath $novo -Force -ErrorAction SilentlyContinue
        return New-ResultadoBanco $false 12 ('não foi possível voltar ao backup: ' + $_.Exception.Message)
    }
    Remove-Item -LiteralPath $anterior -Force -ErrorAction SilentlyContinue
    if ((Get-HashArquivo $producao) -ne (Get-HashArquivo $backup)) { return New-ResultadoBanco $false 12 'a reversão não conferiu com o backup' }
    return New-ResultadoBanco $true 12 'banco voltou ao backup'
}
