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

# ------------------------------------------------------------------ trava de backup (spec 006)
# Uma trava por pasta de backup (backup.lock), criada de forma atômica (CreateNew), com dono (estação, usuário,
# PID, sessão) e instante. Duas execuções ao mesmo tempo (tarefa agendada + botão, ou duas estações) não se
# atropelam. Trava ABANDONADA é recuperada: na mesma estação quando o PID morreu; em qualquer estação quando
# passou de 60 minutos (um backup não leva isso).
$script:TravaBackupMinutos = 60

function Enter-TravaBackup([string]$pasta, [string]$tipo) {
    New-Item -ItemType Directory -Path $pasta -Force | Out-Null
    $arq = Join-Path $pasta 'backup.lock'
    $dono = [ordered]@{ estacao = $env:COMPUTERNAME; usuario = $env:USERNAME; pid = $PID; sessao = [guid]::NewGuid().ToString('N'); tipo = $tipo; inicio = (Get-Date).ToString('o') }
    $recuperada = $false
    for ($tentativa = 0; $tentativa -lt 2; $tentativa++) {
        try {
            $fs = New-Object System.IO.FileStream($arq, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::Read)
            try { $b = [System.Text.Encoding]::UTF8.GetBytes((ConvertTo-Json -InputObject $dono)); $fs.Write($b, 0, $b.Length) } finally { $fs.Dispose() }
            return [pscustomobject]@{ ok = $true; arquivo = $arq; sessao = $dono.sessao; recuperada = $recuperada; motivo = '' }
        } catch [System.IO.IOException] {
            $atual = $null
            try {
                $fl = New-Object System.IO.FileStream($arq, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
                try { $sr = New-Object System.IO.StreamReader($fl); $atual = $sr.ReadToEnd() | ConvertFrom-Json } finally { $fl.Dispose() }
            } catch { }
            $abandonada = $false
            if ($null -eq $atual) { $abandonada = $true }   # ilegível: lixo
            else {
                $inicio = [datetime]::MinValue
                [void][datetime]::TryParse([string]$atual.inicio, [ref]$inicio)
                if (((Get-Date) - $inicio).TotalMinutes -gt $script:TravaBackupMinutos) { $abandonada = $true }
                elseif ([string]$atual.estacao -eq $env:COMPUTERNAME -and -not (Get-Process -Id ([int]$atual.pid) -ErrorAction SilentlyContinue)) { $abandonada = $true }
            }
            if (-not $abandonada) {
                return [pscustomobject]@{ ok = $false; arquivo = $arq; sessao = ''; recuperada = $false
                    motivo = ('há outro backup em andamento: ' + $atual.tipo + ' iniciado às ' + $atual.inicio + ' por ' + $atual.estacao + '\' + $atual.usuario + ' (PID ' + $atual.pid + ').') }
            }
            try { Remove-Item -LiteralPath $arq -Force; $recuperada = $true } catch { }
        }
    }
    return [pscustomobject]@{ ok = $false; arquivo = $arq; sessao = ''; recuperada = $false; motivo = 'não foi possível obter a trava de backup' }
}

# Só remove a trava se a sessão ainda é a nossa (não apaga a trava de quem a recuperou depois).
function Exit-TravaBackup($trava) {
    if ($null -eq $trava -or -not $trava.ok) { return }
    try {
        $atual = Get-Content -LiteralPath $trava.arquivo -Raw -Encoding UTF8 | ConvertFrom-Json
        if ([string]$atual.sessao -eq [string]$trava.sessao) { Remove-Item -LiteralPath $trava.arquivo -Force }
    } catch { }
}

# ------------------------------------------------------------------ backup validado e alerta (spec 006)
# <backup>.validado = JSON { validado_utc, sha256 }: gravado só depois de Test-BackupRestauravel aprovar.
function Set-BackupValidado([string]$arquivoBackup) {
    $m = [ordered]@{ validado_utc = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ'); sha256 = (Get-HashArquivo $arquivoBackup) }
    [System.IO.File]::WriteAllText(($arquivoBackup + '.validado'), (ConvertTo-Json -InputObject $m), (New-Object System.Text.UTF8Encoding($false)))
}

# O backup validado mais recente de uma pasta: { arquivo; instante (local) } ou $null.
function Get-UltimoBackupValidado([string]$pasta) {
    if (-not (Test-Path -LiteralPath $pasta)) { return $null }
    $melhor = $null
    foreach ($f in @(Get-ChildItem -LiteralPath $pasta -Filter 'crm_*.validado' -File -ErrorAction SilentlyContinue)) {
        try {
            $j = Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
            $d = [datetime]::Parse([string]$j.validado_utc, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AdjustToUniversal -bor [Globalization.DateTimeStyles]::AssumeUniversal).ToLocalTime()
            if ($null -eq $melhor -or $d -gt $melhor.instante) { $melhor = [pscustomobject]@{ arquivo = ($f.Name -replace '\.validado$', ''); instante = $d } }
        } catch { }
    }
    return $melhor
}

# Alerta se não há backup validado ou se o último passou de BackupAlertaHoras (36 h).
function Get-AlertaBackup([string]$pasta) {
    $u = Get-UltimoBackupValidado $pasta
    if ($null -eq $u) { return [pscustomobject]@{ alerta = $true; ultimo = $null; horas = $null; mensagem = ('Nenhum backup validado encontrado em ' + $pasta + '. Rode FAZER-BACKUP.bat e TESTAR-RESTAURACAO.bat.') } }
    $h = [math]::Round(((Get-Date) - $u.instante).TotalHours, 1)
    $alerta = ($h -gt $script:BackupAlertaHoras)
    $msg = if ($alerta) { 'O último backup validado tem ' + $h + ' h (limite ' + $script:BackupAlertaHoras + ' h): ' + $u.arquivo } else { 'Último backup validado há ' + $h + ' h.' }
    return [pscustomobject]@{ alerta = $alerta; ultimo = $u; horas = $h; mensagem = $msg }
}

# Nova geração da base: config.geracao_base = GUID. Todo substituto do arquivo do banco (publicação, restauração)
# renova a geração, para o front perceber que a base foi trocada por baixo dele (ciclo 011).
function Set-GeracaoBase($cn, $tx = $null) {
    $g = [guid]::NewGuid().ToString('N')
    Set-ConfigBanco $cn 'geracao_base' $g $tx
    return $g
}
