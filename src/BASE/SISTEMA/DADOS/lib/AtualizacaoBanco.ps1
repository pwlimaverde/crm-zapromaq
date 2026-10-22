<#
  AtualizacaoBanco.ps1 - atualização do banco de produção pela montagem (spec 003, passos 3 a 8 e 11).

  Tudo ou nada: a migração e as conferências rodam numa CÓPIA; o banco de produção só é trocado no
  fim, e só se a trava estiver livre e o hash dele ainda for o H0 (o do backup). Qualquer falha antes
  da troca deixa o banco de produção exatamente como estava.

  Depende de Comum.ps1 (Open-Banco, Invoke-Escalar, Get-ColunasDaLista, Get-ChaveLista) carregado antes.
  Não publica o front: isso é o passo 12 da montagem.
#>

. (Join-Path $PSScriptRoot 'Backup.ps1')   # manifesto e validação do backup (passo 5)

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
    # o .laccdb está aberto pelo próprio Access enquanto há alguém conectado: ler com compartilhamento total
    try {
        $fs = New-Object System.IO.FileStream($laccdb, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        try { $bytes = New-Object byte[] ([int]$fs.Length); [void]$fs.Read($bytes, 0, $bytes.Length) } finally { $fs.Dispose() }
    } catch { return , $estacoes }   # sem conseguir ler os nomes, a trava continua valendo (livre = não existe o arquivo)
    for ($i = 0; $i + 64 -le $bytes.Length; $i += 64) {
        $nome = [Text.Encoding]::ASCII.GetString($bytes, $i, 32).Trim([char]0, ' ')
        if ($nome) { [void]$estacoes.Add($nome) }
    }
    return , $estacoes
}

function Test-BancoLivre([string]$banco) {
    $laccdb = [System.IO.Path]::ChangeExtension($banco, 'laccdb')
    $estacoes = Get-EstacoesDaTrava $laccdb
    return [pscustomobject]@{ livre = -not (Test-BancoEmUso $banco); estacoes = @($estacoes); laccdb = $laccdb }
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
    if ($versao -ne $script:EsquemaAtual) { [void]$problemas.Add('esquema ' + $versao + ' (esperado ' + $script:EsquemaAtual + ')') }
    $municipios = [int](Invoke-Escalar $cn 'SELECT COUNT(*) FROM municipios' @() $null)
    if ($municipios -ne $totalMunicipios) { [void]$problemas.Add('municípios: ' + $municipios + ' (esperado ' + $totalMunicipios + ')') }
    if ([int](Invoke-Escalar $cn 'SELECT COUNT(*) FROM listas WHERE chave IS NULL' @() $null) -ne 0) { [void]$problemas.Add('há linhas de lista sem chave') }
    if ([int](Invoke-Escalar $cn "SELECT COUNT(*) FROM listas WHERE tipo='Etapa' AND quadro IS NULL" @() $null) -ne 0) { [void]$problemas.Add('há etapas sem papel') }
    $schema = $cn.GetSchema('Indexes')
    $indices = @{}
    foreach ($r in $schema.Rows) { $indices[[string]$r['INDEX_NAME']] = $true }
    # 1.1: chave de lista; 1.2 (ciclo 010): controle único e identificador de operação
    foreach ($ix in @('ix_lis_tipo_chave', 'ix_cto_controle', 'ix_opo_controle', 'ix_ope_id')) {
        if (-not $indices.ContainsKey($ix)) { [void]$problemas.Add('índice único ' + $ix + ' ausente') }
    }

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

    # backup VALIDADO (DA-8): manifesto e restauração numa pasta descartável; o arquivo de backup não é aberto
    try { $manifestoBkp = New-ManifestoBackup $backup $producao 'antes-da-publicacao' }
    catch { return New-ResultadoBanco $false 5 ('não foi possível ler o banco de produção para validar o backup (' + $_.Exception.Message.Split([char]13)[0] + '). Nada foi alterado.') }
    $valBkp = Test-BackupRestauravel $backup
    if (-not $valBkp.ok) { return New-ResultadoBanco $false 5 ('o backup não é restaurável: ' + ($valBkp.problemas -join '; ') + '. Nada foi alterado.') }
    $log.L('  backup validado: restaurado numa pasta descartável, esquema ' + $valBkp.esquema + ' e contagens conferidas com o manifesto')

    New-Item -ItemType Directory -Path $pastaTrabalho -Force | Out-Null
    $copia = Join-Path $pastaTrabalho ([IO.Path]::GetFileName($producao))
    Copy-Item -LiteralPath $producao -Destination $copia -Force
    if ((Get-HashArquivo $copia) -ne $h0) { return New-ResultadoBanco $false 6 'a cópia de trabalho não confere com o banco de produção. Nada foi alterado.' }
    $log.L('  cópia de trabalho: ' + $copia)

    $migrar = Join-Path $dados 'banco\aplicar-migracoes.ps1'
    # stderr do processo filho não pode virar exceção aqui: a falha é tratada pelo código de saída
    $preferencia = $ErrorActionPreference; $ErrorActionPreference = 'Continue'
    try { $saidaMig = @(& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $migrar -Banco $copia -CopiaDeTrabalho 2>&1 | ForEach-Object { [string]$_ }) }
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

# Contagem de linhas de cada tabela de dados (nome -> total), na ordem do nome. Só leitura.
function Get-ContagensDoBanco($cn) {
    $nomes = New-Object System.Collections.ArrayList
    foreach ($r in $cn.GetSchema('Tables').Rows) {
        $n = [string]$r['TABLE_NAME']
        if ([string]$r['TABLE_TYPE'] -eq 'TABLE' -and -not $n.StartsWith('MSys') -and -not $n.StartsWith('~')) { [void]$nomes.Add($n) }
    }
    $c = [ordered]@{}
    foreach ($n in ($nomes | Sort-Object)) { $c[$n] = [int](Invoke-Escalar $cn ('SELECT COUNT(*) FROM [' + $n + ']') @() $null) }
    return $c
}

# Impressão digital do conteúdo do banco (texto): para cada tabela, contagem, MAX(id) e SUM(versao) quando
# existem as colunas. Qualquer INSERT, DELETE ou UPDATE de versão muda o texto. Só leitura (Mode=Read).
# É o que a publicação usa para provar que o teste de uso (passo 10) não gravou nada na cópia.
function Get-ImpressaoDoBanco([string]$caminho) {
    $cn = Open-BancoLeitura $caminho
    try {
        $colunas = @{}
        foreach ($r in $cn.GetSchema('Columns').Rows) {
            $t = [string]$r['TABLE_NAME']
            if (-not $colunas.ContainsKey($t)) { $colunas[$t] = New-Object System.Collections.ArrayList }
            [void]$colunas[$t].Add(([string]$r['COLUMN_NAME']).ToLowerInvariant())
        }
        $linhas = New-Object System.Collections.ArrayList
        foreach ($t in ($colunas.Keys | Where-Object { -not $_.StartsWith('MSys') -and -not $_.StartsWith('~') } | Sort-Object)) {
            $partes = @('COUNT(*)')
            if ($colunas[$t].Contains('id')) { $partes += 'MAX(id)' }
            if ($colunas[$t].Contains('versao')) { $partes += 'SUM(versao)' }
            $c = New-ComandoBanco $cn ('SELECT ' + ($partes -join ', ') + ' FROM [' + $t + ']') @() $null
            $rd = $c.ExecuteReader()
            $v = @()
            if ($rd.Read()) { for ($i = 0; $i -lt $rd.FieldCount; $i++) { $x = $rd.GetValue($i); $v += $(if ($x -is [DBNull]) { '-' } else { [string]$x }) } }
            $rd.Close(); $c.Dispose()
            [void]$linhas.Add($t + '=' + ($v -join '|'))
        }
        return ($linhas -join "`n")
    } finally { $cn.Close(); $cn.Dispose() }
}

# Quais execuções do teste de uso o passo 10 faz. Publicação real: só a de leitura (dados reais não podem ser
# gravados nem apagados por teste). -Teste: a de leitura e, depois, a completa (banco de TESTE).
function Get-ExecucoesTesteDeUso([bool]$teste) {
    if ($teste) { return @('SoLeitura', 'Completo') }
    return @('SoLeitura')
}

# Ensaio da atualização (spec 004, F4): faz os passos 3 a 8 sobre uma cópia, em $pastaEnsaio, e escreve o
# relatório. NUNCA troca o banco de produção, não cria backup em execucao\backup e não precisa de Excel.
# Exige a trava livre (a cópia só é consistente com o banco fechado). Confere o hash da produção antes e depois.
function Invoke-EnsaioDoBanco([string]$producao, [string]$pastaEnsaio, [int]$totalMunicipios, $log) {
    $inicio = Get-Date
    $livre = Test-BancoLivre $producao
    if (-not $livre.livre) {
        $quem = if ($livre.estacoes.Count) { ' Estações: ' + ($livre.estacoes -join ', ') + '.' } else { '' }
        return New-ResultadoBanco $false 3 ('há usuário conectado ao banco (.laccdb).' + $quem + ' O ensaio precisa do banco fechado. Nada foi alterado.')
    }
    $h0 = Get-HashArquivo $producao
    $cnL = Open-BancoLeitura $producao
    try { $esquemaAntes = Get-VersaoEsquemaBanco $cnL; $ambiente = [string](Invoke-Escalar $cnL "SELECT valor FROM config WHERE chave='ambiente'" @() $null); $contAntes = Get-ContagensDoBanco $cnL }
    finally { $cnL.Close(); $cnL.Dispose() }
    $log.L('  produção: esquema ' + $esquemaAntes + ', ambiente ' + $ambiente + ', H0 ' + $h0.Substring(0, 12))

    # pasta do ensaio sempre começa vazia (cópia e backup do ensaio anterior saem)
    foreach ($sub in @('copia', 'backup')) {
        $p = Join-Path $pastaEnsaio $sub
        if (Test-Path -LiteralPath $p) { Remove-Item -LiteralPath $p -Recurse -Force }
    }
    $prep = Invoke-PreparacaoDoBanco $producao (Join-Path $pastaEnsaio 'copia') (Join-Path $pastaEnsaio 'backup') $totalMunicipios $log
    if (-not $prep.ok) { return $prep }

    $hDepois = Get-HashArquivo $producao
    if ($hDepois -ne $h0) { return New-ResultadoBanco $false 8 'o banco de produção mudou durante o ensaio (hash diferente); alguém gravou nele. Refaça o ensaio.' }

    $cn = Open-Banco $prep.copia
    try {
        $esquemaDepois = Get-VersaoEsquemaBanco $cn
        $contDepois = Get-ContagensDoBanco $cn
        $aplicadas = New-Object System.Collections.ArrayList
        $c = New-ComandoBanco $cn "SELECT chave, valor FROM config WHERE chave LIKE 'migracao.%' ORDER BY chave" @() $null
        $rd = $c.ExecuteReader()
        while ($rd.Read()) { [void]$aplicadas.Add([string]$rd.GetValue(0) + '  ' + [string]$rd.GetValue(1)) }
        $rd.Close(); $c.Dispose()
    } finally { $cn.Close(); $cn.Dispose() }

    $unif = 0; $inativos = 0
    foreach ($linha in @($prep.migracao)) {
        if ($linha -match 'unificações: (\d+)') { $unif = [int]$matches[1] }
        if ($linha -match 'entrou INATIVO') { $inativos++ }
    }
    $segundos = [int]((Get-Date) - $inicio).TotalSeconds
    $rel = New-Object System.Collections.ArrayList
    [void]$rel.Add('RELATÓRIO DO ENSAIO - CRM Zapromaq   ' + (Get-Date -Format 'dd/MM/yyyy HH:mm:ss'))
    [void]$rel.Add('Este relatório pode conter dado real da produção: NÃO vai para o repositório.')
    [void]$rel.Add('')
    [void]$rel.Add('Produção....: ' + $producao + '  (ambiente ' + $ambiente + ')')
    [void]$rel.Add('Esquema.....: ' + $esquemaAntes + ' -> ' + $esquemaDepois + '  (na cópia)')
    [void]$rel.Add('H0 antes....: ' + $h0)
    [void]$rel.Add('H0 depois...: ' + $hDepois + '  (igual: a produção não foi tocada)')
    [void]$rel.Add('Tempo.......: ' + $segundos + ' s (cópia, migração e conferências)')
    [void]$rel.Add('Migrações registradas na cópia:')
    foreach ($a in $aplicadas) { [void]$rel.Add('  ' + $a) }
    [void]$rel.Add('Listas: ' + $unif + ' unificação(ões); ' + $inativos + ' valor(es) incluído(s) como inativo(s)')
    [void]$rel.Add('')
    [void]$rel.Add('Contagens (antes -> depois, na cópia):')
    $todas = @($contAntes.Keys) + @($contDepois.Keys | Where-Object { -not $contAntes.Contains($_) })
    foreach ($t in $todas) {
        $a = if ($contAntes.Contains($t)) { [string]$contAntes[$t] } else { '(não existia)' }
        $d = if ($contDepois.Contains($t)) { [string]$contDepois[$t] } else { '-' }
        [void]$rel.Add(('  {0,-20} {1,10} -> {2}' -f $t, $a, $d))
    }
    [void]$rel.Add('')
    [void]$rel.Add('Saída da migração:')
    foreach ($linha in @($prep.migracao)) { if ($linha -ne '') { [void]$rel.Add('  ' + $linha) } }
    $arq = Join-Path $pastaEnsaio ('relatorio-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.txt')
    [System.IO.File]::WriteAllLines($arq, [string[]]$rel.ToArray(), (New-Object System.Text.UTF8Encoding($true)))
    return New-ResultadoBanco $true 8 'ensaio aprovado' @{ relatorio = $arq; h0 = $h0; esquema_antes = $esquemaAntes; esquema_depois = $esquemaDepois; migracoes = @($aplicadas); copia = $prep.copia; segundos = $segundos }
}

# config.versao_front é gravada NA CÓPIA, antes da troca (passo 11). Antes ela era gravada na produção
# depois da troca, o que mudava o hash e impedia a reversão (Reverter-Troca exige o hash da troca).
function Set-VersaoFrontNaCopia([string]$copia, [string]$versao) {
    $cn = Open-Banco $copia
    # a troca do arquivo do banco também renova a geração da base (config.geracao_base, spec 006)
    try { Set-ConfigBanco $cn 'versao_front' $versao $null; [void](Set-GeracaoBase $cn) } finally { $cn.Close(); $cn.Dispose() }
}

# Passo 11: troca o banco de produção pela cópia, só com a trava livre e o hash igual ao H0 (C19, C20).
# Antes de trocar, guarda o banco anterior em $pastaBackup (para reverter).
function Trocar-BancoProducao([string]$copia, [string]$producao, [string]$h0, [string]$pastaBackup) {
    $livre = Test-BancoLivre $producao
    if (-not $livre.livre) { return New-ResultadoBanco $false 11 'há usuário conectado ao banco (.laccdb); nada foi trocado. Rode de novo quando fecharem o CRM.' }
    $atual = Get-HashArquivo $producao
    if ($atual -ne $h0) { return New-ResultadoBanco $false 11 'alguém gravou no banco durante a montagem (o hash mudou); nada foi trocado. Rode de novo.' }

    # Tudo o que demora (backup da produção, cópia do .novo para a pasta de produção e a conferência do
    # hash dele) acontece ANTES da última conferência. Entre ela e o Replace sobra só uma checagem barata
    # (trava, tamanho e data do arquivo), para não deixar uma janela de segundos em que uma gravação se perde.
    New-Item -ItemType Directory -Path $pastaBackup -Force | Out-Null
    Copy-Item -LiteralPath $producao -Destination (Join-Path $pastaBackup ([IO.Path]::GetFileName($producao))) -Force

    # a cópia novo fica na MESMA pasta do banco de produção: a troca é um renomear dentro do volume
    $novo = $producao + '.novo'
    Copy-Item -LiteralPath $copia -Destination $novo -Force
    if ((Get-HashArquivo $novo) -ne (Get-HashArquivo $copia)) {
        Remove-Item -LiteralPath $novo -Force -ErrorAction SilentlyContinue
        return New-ResultadoBanco $false 11 'a cópia do novo banco não conferiu na pasta de produção; nada foi trocado.'
    }
    $arq = Get-Item -LiteralPath $producao
    $tamanho = $arq.Length; $quandoGravou = $arq.LastWriteTimeUtc
    if ($script:GanchoAntesDaTroca) { & $script:GanchoAntesDaTroca }   # só os testes usam: simula uma gravação tardia

    # última conferência, logo antes do Replace
    $arq = Get-Item -LiteralPath $producao
    if (-not (Test-BancoLivre $producao).livre) {
        Remove-Item -LiteralPath $novo -Force -ErrorAction SilentlyContinue
        return New-ResultadoBanco $false 11 'alguém abriu o banco (.laccdb) logo antes da troca; nada foi trocado. Rode de novo quando fecharem o CRM.'
    }
    if ($arq.Length -ne $tamanho -or $arq.LastWriteTimeUtc -ne $quandoGravou) {
        Remove-Item -LiteralPath $novo -Force -ErrorAction SilentlyContinue
        return New-ResultadoBanco $false 11 'alguém gravou no banco logo antes da troca (o arquivo mudou depois da preparação); nada foi trocado. Rode de novo.'
    }
    $anterior = $producao + '.anterior'   # o Replace exige um nome de backup real, na mesma pasta
    try { [System.IO.File]::Replace($novo, $producao, $anterior) }
    catch {
        Remove-Item -LiteralPath $novo -Force -ErrorAction SilentlyContinue
        return New-ResultadoBanco $false 11 ('não foi possível trocar o arquivo de produção (' + $_.Exception.Message + '); nada foi trocado.')
    }
    if ($script:GanchoDepoisDaTroca) { & $script:GanchoDepoisDaTroca }   # só os testes usam
    # o que o Replace pôs em .anterior é o arquivo que existia no instante da troca: tem de ser o H0
    if ((Get-HashArquivo $anterior) -ne $h0) {
        $rejeitado = $producao + '.rejeitado'
        try {
            [System.IO.File]::Replace($anterior, $producao, $rejeitado)
            Remove-Item -LiteralPath $rejeitado -Force -ErrorAction SilentlyContinue
            return New-ResultadoBanco $false 11 'alguém gravou no banco no instante da troca; o banco anterior (com essa gravação) foi devolvido. Rode de novo.'
        } catch {
            return New-ResultadoBanco $false 11 ('alguém gravou no banco no instante da troca e não foi possível devolver o anterior (' + $_.Exception.Message + '); ele está em ' + $anterior + '. Avise antes de mexer.')
        }
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
