<#
  Backup.ps1 - manifesto e validação de backup do banco (spec 004, F8).

  Backup que nunca foi restaurado não é backup. Aqui "validado" significa: o arquivo bate com o SHA-256 do
  manifesto E uma CÓPIA dele, restaurada numa pasta descartável e aberta só para leitura, tem o esquema e as
  contagens do manifesto. O arquivo de backup em si NUNCA é aberto pelo ACE (abrir um .accdb pode alterar bytes).

  Depende de Comum.ps1 e de AtualizacaoBanco.ps1 (Get-HashArquivo, Get-ContagensDoBanco) carregados antes.
#>

# Política (seção 5.3 do plano de 19/10/2026, substitui o alerta de 36 h do DA-2): janelas 9h, 12h30 e 16h nos
# dias úteis; aviso por janela sem cópia e por dia útil sem cópia completa validada (Get-AvisoBackupJanela).
# RPO e RTO são definidos pela gestão (manual de restauração, docs\implantacao.md).

# Escreve <backup>.manifesto.json: instante, tipo, SHA-256 e tamanho do arquivo, esquema e contagens do banco
# de ORIGEM (lido em modo somente leitura, com o banco fechado). Não abre o arquivo de backup.
function New-ManifestoBackup([string]$arquivoBackup, [string]$origemBanco, [string]$tipo, $extra = @{}) {
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
    foreach ($k in $extra.Keys) { $m[$k] = $extra[$k] }
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
            if (-not (Test-Path -LiteralPath $arq)) { continue }   # o dono acabou de soltar: tenta de novo
            $atual = $null
            try {
                $fl = New-Object System.IO.FileStream($arq, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
                try { $sr = New-Object System.IO.StreamReader($fl); $atual = $sr.ReadToEnd() | ConvertFrom-Json } finally { $fl.Dispose() }
            } catch { }
            $abandonada = $false
            if ($null -eq $atual) {
                # ilegível: ou é lixo antigo, ou o dono ACABOU de criar e ainda está escrevendo (disparos simultâneos,
                # ciclo 007). Só é lixo se o arquivo tem mais de 2 minutos.
                $idade = 0
                try { $idade = ((Get-Date) - (Get-Item -LiteralPath $arq).LastWriteTime).TotalSeconds } catch { }
                if ($idade -gt 120) { $abandonada = $true }
                else { return [pscustomobject]@{ ok = $false; arquivo = $arq; sessao = ''; recuperada = $false; motivo = 'há outro backup começando agora (trava sendo criada).' } }
            }
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

# ------------------------------------------------------------------ exclusividade durante a cópia (DEV-03, ciclo 007)
# O backup completo segura o ARQUIVO do banco com FileShare.Read durante TODA a captura (cópia, hash e contagens do
# manifesto) e copia pelo próprio fluxo. Enquanto isso o ACE de quem abrir o CRM entra SÓ LEITURA (consultar
# continua funcionando) e qualquer gravação falha com "A operação deve usar uma consulta atualizável" - antes de
# gravar, sem efeito parcial. A marca execucao\backup\copiando.txt existe enquanto o arquivo está seguro: é por
# ela que o front e o conector trocam esse erro pela mensagem controlada ("cópia de segurança em andamento").
# Quem já está conectado (gravação aberta) impede a captura: o FileStream não abre e a tentativa repete.
function Get-MarcaCopiando([string]$pastaBkp) { return (Join-Path $pastaBkp 'copiando.txt') }

function Enter-CopiaExclusiva([string]$banco, [string]$pastaBkp, [int]$esperarSeg = 60) {
    $marca = Get-MarcaCopiando $pastaBkp
    New-Item -ItemType Directory -Path $pastaBkp -Force | Out-Null
    [System.IO.File]::WriteAllText($marca, ($env:COMPUTERNAME + '|' + $PID + '|' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')), (New-Object System.Text.UTF8Encoding($false)))
    $limite = (Get-Date).AddSeconds($esperarSeg)
    $ultimoErro = ''
    do {
        try {
            $fs = New-Object System.IO.FileStream($banco, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
            return [pscustomobject]@{ ok = $true; fluxo = $fs; marca = $marca; motivo = '' }
        } catch [System.IO.IOException] { $ultimoErro = $_.Exception.Message.Split([char]13)[0] }
        Start-Sleep -Milliseconds 500
    } while ((Get-Date) -lt $limite)
    Remove-Item -LiteralPath $marca -Force -ErrorAction SilentlyContinue
    return [pscustomobject]@{ ok = $false; fluxo = $null; marca = $marca; motivo = ('o banco está em uso por quem grava (' + $ultimoErro + ')') }
}

function Exit-CopiaExclusiva($copia) {
    if ($null -eq $copia) { return }
    if ($copia.fluxo) { try { $copia.fluxo.Dispose() } catch { } }
    Remove-Item -LiteralPath $copia.marca -Force -ErrorAction SilentlyContinue
}

# Copia um fluxo aberto para um arquivo (o .tmp do backup).
function Copy-FluxoParaArquivo($fluxo, [string]$destino) {
    $fluxo.Position = 0
    $saida = New-Object System.IO.FileStream($destino, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
    try { $fluxo.CopyTo($saida, 1048576) } finally { $saida.Dispose() }
}

# Cópia simples (CRM aberto): lê com FileShare.ReadWrite, para NÃO bloquear quem grava (o Copy-Item abre o
# arquivo só com compartilhamento de leitura e fazia o ACE recusar gravações durante a cópia).
function Copy-BancoSemBloquear([string]$banco, [string]$destino) {
    $fs = New-Object System.IO.FileStream($banco, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    try { Copy-FluxoParaArquivo $fs $destino } finally { $fs.Dispose() }
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

# DEV-03 (ciclo 018): a cópia para RESTAURAR quando ninguém escolhe uma - a mais recente (pelo instante do
# backup) COMPLETA (manifesto com consistencia = exclusiva) que tem <backup>.validado e cujo SHA-256 ainda é o
# que a validação registrou (cópia trocada depois da validação não vale). Cópia simples (nao-verificada) e cópia
# sem manifesto nunca entram. Devolve o FileInfo ou $null.
function Get-BackupValidadoMaisRecente([string]$pasta) {
    if (-not (Test-Path -LiteralPath $pasta)) { return $null }
    $lista = @(Get-ChildItem -LiteralPath $pasta -Filter 'crm_*.accdb' -File -ErrorAction SilentlyContinue |
               Where-Object { (Test-Path -LiteralPath ($_.FullName + '.validado')) -and (Test-Path -LiteralPath ($_.FullName + '.manifesto.json')) })
    foreach ($f in @($lista | Sort-Object { Get-InstanteDoBackup $_ } -Descending)) {
        try {
            $man = Get-Content -LiteralPath ($f.FullName + '.manifesto.json') -Raw -Encoding UTF8 | ConvertFrom-Json
            if ([string]$man.consistencia -ne 'exclusiva') { continue }
            $j = Get-Content -LiteralPath ($f.FullName + '.validado') -Raw -Encoding UTF8 | ConvertFrom-Json
            if ([string]$j.sha256 -and [string]$j.sha256 -eq (Get-HashArquivo $f.FullName)) { return $f }
        } catch { }
    }
    return $null
}

# ------------------------------------------------------------------ janelas e aviso (seção 5.3, ciclo 007)
# Janelas: 9h, 12h30 e 16h, de segunda a sexta. Tentativas no horário, +10 e +20 min; a janela ENCERRA 30 min
# depois do início (20 min de tentativas + 10 de cópia e validação). Uma cópia conta para a janela se o instante
# dela estiver em [início, início + 60 min). Mesma regra em IA\mcp\lib\Ferramentas.ps1 (crm_status) e
# modBackupAviso (front): mudou aqui, muda lá - os três testes usam os mesmos casos.
$script:JanelasBackup = @('09:00', '12:30', '16:00')
$script:JanelaEncerraMin = 30
$script:JanelaContaMin = 60
$script:FimDoDiaUtil = '16:30'   # depois disto, o dia útil tem de ter cópia completa validada

function Test-DiaUtil([datetime]$d) { return ($d.DayOfWeek -ne 'Saturday' -and $d.DayOfWeek -ne 'Sunday') }
function Get-InicioJanela([datetime]$dia, [string]$hhmm) { $p = $hhmm.Split(':'); return $dia.Date.AddHours([int]$p[0]).AddMinutes([int]$p[1]) }

# Janela que uma execução agendada cumpre agora: 'AAAAMMDD-HHMM' ou '' (fora de janela).
function Get-JanelaBackup([datetime]$agora) {
    if (-not (Test-DiaUtil $agora)) { return '' }
    foreach ($h in ($script:JanelasBackup | Sort-Object -Descending)) {
        $ini = Get-InicioJanela $agora $h
        if ($agora -ge $ini -and ($agora - $ini).TotalMinutes -lt $script:JanelaContaMin) { return $ini.ToString('yyyyMMdd-HHmm') }
    }
    return ''
}

# Arquivo de cópia cujo manifesto diz que atendeu a janela ('' = nenhum).
function Get-CopiaDaJanela([string]$pasta, [string]$janela) {
    foreach ($m in @(Get-ChildItem -LiteralPath $pasta -Filter 'crm_*.accdb.manifesto.json' -File -ErrorAction SilentlyContinue)) {
        try {
            $j = Get-Content -LiteralPath $m.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
            if ([string]$j.janela -eq $janela -and (Test-Path -LiteralPath ($m.FullName -replace '\.manifesto\.json$', ''))) { return [string]$j.arquivo }
        } catch { }
    }
    return ''
}

# Cópias da pasta, pelo NOME: { instante; tipo (diario|periodico); validado }.
function Get-CopiasBackup([string]$pasta) {
    $lista = New-Object System.Collections.ArrayList
    foreach ($f in @(Get-ChildItem -LiteralPath $pasta -Filter 'crm_*.accdb' -File -ErrorAction SilentlyContinue)) {
        if ($f.Name -match '^crm_(diario|periodico)_(\d{8})-(\d{4}|\d{6})(-\d+)?_') {
            $hora = $Matches[3]; if ($hora.Length -eq 4) { $hora += '00' }
            $d = [datetime]::MinValue
            if ([datetime]::TryParseExact($Matches[2] + $hora, 'yyyyMMddHHmmss', [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$d)) {
                [void]$lista.Add([pscustomobject]@{ instante = $d; tipo = $Matches[1]; validado = (Test-Path -LiteralPath ($f.FullName + '.validado')) })
            }
        }
    }
    return , $lista
}

# Regra pura do aviso (seção 5.3): (a) a última janela encerrada não tem cópia nenhuma; (b) o último dia útil
# cujas 16h30 já passaram não tem cópia COMPLETA VALIDADA (a simples não apaga este aviso).
function Get-AvisoBackupJanela([datetime]$agora, $copias) {
    $motivos = New-Object System.Collections.ArrayList
    $copias = @($copias)
    # (a)
    $ultima = $null
    for ($k = 0; $k -le 10 -and $null -eq $ultima; $k++) {
        $dia = $agora.Date.AddDays(-$k)
        if (-not (Test-DiaUtil $dia)) { continue }
        foreach ($h in ($script:JanelasBackup | Sort-Object -Descending)) {
            $ini = Get-InicioJanela $dia $h
            if ($agora -ge $ini.AddMinutes($script:JanelaEncerraMin)) { $ultima = $ini; break }
        }
    }
    if ($null -ne $ultima) {
        $tem = @($copias | Where-Object { $_.instante -ge $ultima -and $_.instante -lt $ultima.AddMinutes($script:JanelaContaMin) }).Count -gt 0
        if (-not $tem) { [void]$motivos.Add('a janela das ' + $ultima.ToString('HH:mm') + ' de ' + $ultima.ToString('dd/MM/yyyy') + ' terminou sem cópia') }
    }
    # (b)
    $diaB = $null
    for ($k = 0; $k -le 10 -and $null -eq $diaB; $k++) {
        $dia = $agora.Date.AddDays(-$k)
        if ((Test-DiaUtil $dia) -and $agora -ge (Get-InicioJanela $dia $script:FimDoDiaUtil)) { $diaB = $dia }
    }
    if ($null -ne $diaB) {
        $tem = @($copias | Where-Object { $_.tipo -eq 'diario' -and $_.validado -and $_.instante.Date -eq $diaB }).Count -gt 0
        if (-not $tem) { [void]$motivos.Add('o dia ' + $diaB.ToString('dd/MM/yyyy') + ' terminou sem cópia completa validada') }
    }
    $msg = if ($motivos.Count) { 'Backup atrasado: ' + ($motivos -join '; ') + '. Avise quem cuida dos backups (FAZER-BACKUP.bat).' } else { 'Backups em dia.' }
    return [pscustomobject]@{ alerta = ($motivos.Count -gt 0); motivos = @($motivos); mensagem = $msg }
}

# Aviso da pasta de backup: regra por janela + o último backup validado (para mostrar).
function Get-AlertaBackup([string]$pasta, $agora = $null) {
    if ($null -eq $agora) { $agora = Get-Date }
    $u = Get-UltimoBackupValidado $pasta
    $aviso = Get-AvisoBackupJanela $agora (Get-CopiasBackup $pasta)
    $h = if ($u) { [math]::Round(($agora - $u.instante).TotalHours, 1) } else { $null }
    return [pscustomobject]@{ alerta = $aviso.alerta; motivos = $aviso.motivos; ultimo = $u; horas = $h; mensagem = $aviso.mensagem }
}

# Nova geração da base: config.geracao_base = GUID. Todo substituto do arquivo do banco (publicação, restauração)
# renova a geração, para o front perceber que a base foi trocada por baixo dele (ciclo 011).
function Set-GeracaoBase($cn, $tx = $null) {
    $g = [guid]::NewGuid().ToString('N')
    Set-ConfigBanco $cn 'geracao_base' $g $tx
    return $g
}
