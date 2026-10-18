<#
  Backup.ps1 - manifesto e validação de backup do banco (spec 004, F8).

  Backup que nunca foi restaurado não é backup. Aqui "validado" significa: o arquivo bate com o SHA-256 do
  manifesto E uma CÓPIA dele, restaurada numa pasta descartável e aberta só para leitura, tem o esquema e as
  contagens do manifesto. O arquivo de backup em si NUNCA é aberto pelo ACE (abrir um .accdb pode alterar bytes).

  Depende de Comum.ps1 e de AtualizacaoBanco.ps1 (Get-HashArquivo, Get-ContagensDoBanco) carregados antes.
#>

# Política (DA-2): uma cópia validada por dia útil (RPO 24 h), restauração em até 4 h (RTO), alerta se não
# houver backup validado há mais de 36 h. Mudar o número é mudar uma linha aqui.
$script:BackupRpoHoras = 24
$script:BackupRtoHoras = 4
$script:BackupAlertaHoras = 36

# Escreve <backup>.manifesto.json: instante, tipo, SHA-256 e tamanho do arquivo, esquema e contagens do banco
# de ORIGEM (lido em modo somente leitura, com o banco fechado). Não abre o arquivo de backup.
function New-ManifestoBackup([string]$arquivoBackup, [string]$origemBanco, [string]$tipo) {
    $cn = Open-BancoLeitura $origemBanco
    try { $esquema = Get-VersaoEsquemaBanco $cn; $contagens = Get-ContagensDoBanco $cn } finally { $cn.Close(); $cn.Dispose() }
    $agora = Get-Date
    $m = [ordered]@{
        versao_manifesto = 1
        instante_utc     = $agora.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
        instante_local   = $agora.ToString('yyyy-MM-dd HH:mm:ss')
        tipo             = $tipo
        arquivo          = [IO.Path]::GetFileName($arquivoBackup)
        sha256           = (Get-HashArquivo $arquivoBackup)
        tamanho          = (Get-Item -LiteralPath $arquivoBackup).Length
        estacao          = $env:COMPUTERNAME
        esquema          = $esquema
        contagens        = $contagens
    }
    $json = ConvertTo-Json -InputObject $m -Depth 4
    [System.IO.File]::WriteAllText(($arquivoBackup + '.manifesto.json'), $json, (New-Object System.Text.UTF8Encoding($false)))
    return [pscustomobject]$m
}

# Valida o backup: hash == manifesto; cópia restaurada numa pasta descartável abre e tem esquema e contagens
# do manifesto. Devolve { ok; problemas; esquema; contagens }.
function Test-BackupRestauravel([string]$arquivoBackup, [string]$manifesto = '') {
    $problemas = New-Object System.Collections.ArrayList
    if (-not $manifesto) { $manifesto = $arquivoBackup + '.manifesto.json' }
    if (-not (Test-Path -LiteralPath $arquivoBackup)) { [void]$problemas.Add('backup inexistente: ' + $arquivoBackup); return [pscustomobject]@{ ok = $false; problemas = $problemas; esquema = $null; contagens = $null } }
    if (-not (Test-Path -LiteralPath $manifesto)) { [void]$problemas.Add('sem manifesto: ' + $manifesto); return [pscustomobject]@{ ok = $false; problemas = $problemas; esquema = $null; contagens = $null } }
    $m = Get-Content -LiteralPath $manifesto -Raw -Encoding UTF8 | ConvertFrom-Json
    if ((Get-HashArquivo $arquivoBackup) -ne [string]$m.sha256) { [void]$problemas.Add('o SHA-256 do backup não confere com o do manifesto') }
    $pasta = Join-Path ([System.IO.Path]::GetTempPath()) ('crm-restaurar-' + [guid]::NewGuid().ToString('N'))
    $esquema = $null; $contagens = $null
    try {
        New-Item -ItemType Directory -Path $pasta -Force | Out-Null
        $copia = Join-Path $pasta 'restaurado.accdb'
        Copy-Item -LiteralPath $arquivoBackup -Destination $copia
        try {
            $cn = Open-BancoLeitura $copia
            try { $esquema = Get-VersaoEsquemaBanco $cn; $contagens = Get-ContagensDoBanco $cn } finally { $cn.Close(); $cn.Dispose() }
        } catch { [void]$problemas.Add('a cópia restaurada NÃO abriu: ' + $_.Exception.Message.Split([char]13)[0]) }
        if ($null -ne $contagens) {
            if ($esquema -ne [string]$m.esquema) { [void]$problemas.Add('esquema ' + $esquema + ' (manifesto: ' + $m.esquema + ')') }
            foreach ($p in $m.contagens.PSObject.Properties) {
                $obtido = if ($contagens.Contains($p.Name)) { [int]$contagens[$p.Name] } else { -1 }
                if ($obtido -ne [int]$p.Value) { [void]$problemas.Add('contagem de ' + $p.Name + ': ' + $obtido + ' (manifesto: ' + $p.Value + ')') }
            }
        }
    } finally { Remove-Item -LiteralPath $pasta -Recurse -Force -ErrorAction SilentlyContinue }
    return [pscustomobject]@{ ok = ($problemas.Count -eq 0); problemas = $problemas; esquema = $esquema; contagens = $contagens }
}

# Instante de um backup: o do manifesto, se houver; senão o do NOME (crm_<tipo>_AAAAMMDD-HHMM_...); só em
# último caso o LastWriteTime (Copy-Item preserva a data do banco original, então ela não é a hora do backup).
function Get-InstanteDoBackup([System.IO.FileInfo]$arquivo) {
    $man = $arquivo.FullName + '.manifesto.json'
    if (Test-Path -LiteralPath $man) {
        try {
            $m = Get-Content -LiteralPath $man -Raw -Encoding UTF8 | ConvertFrom-Json
            $d = [datetime]::MinValue
            if ([datetime]::TryParse([string]$m.instante_local, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$d)) { return $d }
        } catch { }
    }
    if ($arquivo.Name -match '^crm_[a-z]+_(\d{8})-(\d{4})_') {
        $d = [datetime]::MinValue
        if ([datetime]::TryParseExact($Matches[1] + $Matches[2], 'yyyyMMddHHmm', [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$d)) { return $d }
    }
    return $arquivo.LastWriteTime
}

# O backup mais recente de uma pasta (crm_*.accdb), pelo instante acima.
function Get-BackupMaisRecente([string]$pasta) {
    $lista = @(Get-ChildItem -LiteralPath $pasta -Filter 'crm_*.accdb' -File -ErrorAction SilentlyContinue)
    if ($lista.Count -eq 0) { return $null }
    return ($lista | Sort-Object { Get-InstanteDoBackup $_ } -Descending | Select-Object -First 1)
}
