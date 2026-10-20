<#
  testar-recuperacao.ps1 - testes da Etapa 1 do plano de 19/10/2026 (ciclo 007), com ACE e sem Excel:
    DEV-08 sinal de manutenção (montagem e migrações), DEV-03 exclusividade durante a cópia, janela
    atendida e aviso por janela, DEV-06A diagnóstico e retenção.

  Tudo em pasta temporária e com bancos FICTÍCIOS (criar-banco-teste.ps1): nada de BASE é tocado.
#>
$ErrorActionPreference = 'Stop'
$dados = Split-Path -Parent $PSScriptRoot
$raizBase = Split-Path -Parent (Split-Path -Parent $dados)
. (Join-Path $dados 'lib\Comum.ps1')
. (Join-Path $dados 'lib\AtualizacaoBanco.ps1')

$ps = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('crm-trc-' + $PID)
New-Item -ItemType Directory -Path $tmp -Force | Out-Null

$script:falhas = 0
function Conferir([string]$nome, [bool]$ok, [string]$detalhe = '') {
    if ($ok) { Write-Host ('  OK    ' + $nome) -ForegroundColor Green }
    else { $script:falhas++; Write-Host ('  FALHA ' + $nome + '  ' + $detalhe) -ForegroundColor Red }
}
function Executar([string]$script, [string[]]$argumentos) {
    $pref = $ErrorActionPreference; $ErrorActionPreference = 'Continue'
    try {
        $saida = @(& $ps -NoProfile -ExecutionPolicy Bypass -File $script @argumentos 2>&1 | ForEach-Object { [string]$_ })
        return [pscustomobject]@{ codigo = $LASTEXITCODE; saida = ($saida -join "`n") }
    } finally { $ErrorActionPreference = $pref }
}
# BASE mínima em pasta temporária, com banco de homologação fictício na raiz.
function NovaBase([string]$nome) {
    $b = Join-Path $tmp $nome
    foreach ($rel in @('IA', 'SISTEMA\DADOS\lib', 'SISTEMA\DADOS\banco', 'SISTEMA\DADOS\build', 'SISTEMA\DADOS\testes', 'SISTEMA\DADOS\operacao')) {
        $dest = Join-Path $b $rel
        New-Item -ItemType Directory -Path (Split-Path -Parent $dest) -Force | Out-Null
        Copy-Item -LiteralPath (Join-Path $raizBase $rel) -Destination $dest -Recurse
    }
    Copy-Item -LiteralPath (Join-Path $dados 'VERSAO.txt') -Destination (Join-Path $b 'SISTEMA\DADOS\VERSAO.txt')
    $r = Executar (Join-Path $b 'IA\teste\criar-banco-teste.ps1') @('-Banco', (Join-Path $b 'crm_zapromaq.accdb'), '-Recriar')
    if ($r.codigo -ne 0) { throw ('não criou o banco de teste: ' + $r.saida) }
    # a criação aplica as migrações, e o aplicar-migracoes faz um backup do banco da raiz: cada caso começa sem cópias
    $bk = Join-Path $b 'SISTEMA\DADOS\execucao\backup'
    if (Test-Path -LiteralPath $bk) { Remove-Item -LiteralPath $bk -Recurse -Force }
    return $b
}
function ValorConfig([string]$banco, [string]$chave) {
    $cn = Open-BancoLeitura $banco
    try { return [string](Invoke-Escalar $cn 'SELECT valor FROM config WHERE chave=?' @((New-Parametro 'texto' $chave)) $null) } finally { $cn.Close(); $cn.Dispose() }
}

try {
    Write-Host '--- DEV-08: sinal de manutenção (config.manutencao)'
    $b1 = NovaBase 'base-manutencao'
    $banco1 = Join-Path $b1 'crm_zapromaq.accdb'
    $cn = Open-Banco $banco1
    try {
        $s1 = Set-Manutencao $cn 'teste de manutenção'
        Conferir 'Set-Manutencao grava estação|usuário|PID|início|motivo' ((Get-Manutencao $cn) -eq $s1 -and $s1.Split('|').Count -eq 5 -and $s1.StartsWith($env:COMPUTERNAME + '|')) $s1
        Conferir 'Clear-Manutencao com outro dono não desliga' (-not (Clear-Manutencao $cn 'OUTRA|x|1|y|z') -and (Get-Manutencao $cn) -eq $s1)
        Conferir 'Clear-Manutencao do próprio dono desliga' ((Clear-Manutencao $cn $s1) -and (Get-Manutencao $cn) -eq '')
        [void](Set-Manutencao $cn 'dono que caiu')
    } finally { $cn.Close(); $cn.Dispose() }
    $r = Executar (Join-Path $b1 'SISTEMA\DADOS\banco\aplicar-migracoes.ps1') @('-Banco', $banco1, '-LiberarManutencao')
    Conferir 'aplicar-migracoes -LiberarManutencao desliga o sinal deixado por dono que caiu' ($r.codigo -eq 0 -and (ValorConfig $banco1 'manutencao') -eq '') $r.saida

    Write-Host '--- DEV-03: aviso por janela (seção 5.3), casos com hora fixa (os mesmos do testar-mcp e do modAutoteste)'
    function C([string]$quando, [string]$tipo, [bool]$val) { return [pscustomobject]@{ instante = [datetime]::ParseExact($quando, 'yyyy-MM-dd HH:mm', $null); tipo = $tipo; validado = $val } }
    function Agora([string]$quando) { return [datetime]::ParseExact($quando, 'yyyy-MM-dd HH:mm', $null) }
    $casos = @(
        @{ n = 'J1 tudo em dia (seg 10h, sexta completa e 9h de seg)'; agora = '2026-10-19 10:00'; copias = @((C '2026-10-16 16:05' 'diario' $true), (C '2026-10-19 09:05' 'diario' $true)); alerta = $false; motivos = 0 }
        @{ n = 'J2 janela das 9h de seg encerrada sem cópia'; agora = '2026-10-19 10:00'; copias = @((C '2026-10-16 16:05' 'diario' $true)); alerta = $true; motivos = 1 }
        @{ n = 'J3 janela das 9h ainda aberta (9h20): não avisa'; agora = '2026-10-19 09:20'; copias = @((C '2026-10-16 16:05' 'diario' $true)); alerta = $false; motivos = 0 }
        @{ n = 'J4 só cópias simples no dia: aviso (b) depois das 16h30'; agora = '2026-10-19 17:00'; copias = @((C '2026-10-16 16:05' 'diario' $true), (C '2026-10-19 09:05' 'periodico' $true), (C '2026-10-19 12:35' 'periodico' $true), (C '2026-10-19 16:10' 'periodico' $true)); alerta = $true; motivos = 1 }
        @{ n = 'J5 completa validada às 16h10: em dia'; agora = '2026-10-19 17:00'; copias = @((C '2026-10-19 16:10' 'diario' $true)); alerta = $false; motivos = 0 }
        @{ n = 'J6 estação desligada na sexta: os dois avisos na segunda cedo'; agora = '2026-10-19 08:00'; copias = @((C '2026-10-15 16:05' 'diario' $true)); alerta = $true; motivos = 2 }
        @{ n = 'J7 sábado não tem janela'; agora = '2026-10-24 11:00'; copias = @((C '2026-10-23 16:05' 'diario' $true)); alerta = $false; motivos = 0 }
        @{ n = 'J8 completa sem validação não apaga o aviso (b)'; agora = '2026-10-19 17:00'; copias = @((C '2026-10-19 16:10' 'diario' $false)); alerta = $true; motivos = 1 }
    )
    foreach ($k in $casos) {
        $av = Get-AvisoBackupJanela (Agora $k.agora) $k.copias
        Conferir $k.n ($av.alerta -eq $k.alerta -and @($av.motivos).Count -eq $k.motivos) ($av | ConvertTo-Json -Compress)
    }
    Conferir 'janela de execução: 9h05 de segunda é 20261019-0900' ((Get-JanelaBackup (Agora '2026-10-19 09:05')) -eq '20261019-0900')
    Conferir 'janela de execução: 12h40 é a de 12h30' ((Get-JanelaBackup (Agora '2026-10-19 12:40')) -eq '20261019-1230')
    Conferir 'janela de execução: 11h não é janela' ((Get-JanelaBackup (Agora '2026-10-19 11:00')) -eq '')
    Conferir 'janela de execução: sábado não é janela' ((Get-JanelaBackup (Agora '2026-10-24 09:05')) -eq '')

    Write-Host '--- DEV-03: exclusividade durante TODA a cópia (ACE real)'
    $b3 = NovaBase 'base-exclusiva'
    $banco3 = Join-Path $b3 'crm_zapromaq.accdb'
    $pastaBkp3 = Join-Path $b3 'SISTEMA\DADOS\execucao\backup'
    function TentaGravar([string]$banco) {
        $cn = New-Object System.Data.OleDb.OleDbConnection('Provider=' + (Get-ProvedorAce) + ';Data Source=' + $banco + ';OLE DB Services=-4;')
        try {
            $cn.Open()
            $c = $cn.CreateCommand(); $c.CommandText = "UPDATE config SET valor='x' WHERE chave='nao-existe'"
            try { [void]$c.ExecuteNonQuery(); return '' } catch { return $_.Exception.Message } finally { $c.Dispose() }
        } catch { return ('abrir: ' + $_.Exception.Message) } finally { $cn.Dispose() }
    }
    function ChamarMcp([string]$banco, [string]$ferramenta, $argumentos) {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $ps
        $psi.Arguments = '-NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "' + (Join-Path (Split-Path -Parent $banco) 'IA\mcp\servidor.ps1') + '" -Banco "' + $banco + '"'
        $psi.UseShellExecute = $false; $psi.RedirectStandardInput = $true; $psi.RedirectStandardOutput = $true; $psi.RedirectStandardError = $true
        $psi.StandardOutputEncoding = New-Object System.Text.UTF8Encoding($false); $psi.CreateNoWindow = $true
        $p = [System.Diagnostics.Process]::Start($psi)
        $err = $p.StandardError.ReadToEndAsync()
        $utf8 = New-Object System.Text.UTF8Encoding($false)
        $msg = [ordered]@{ jsonrpc = '2.0'; id = 1; method = 'tools/call'; params = [ordered]@{ name = $ferramenta; arguments = $argumentos } }
        $bt = $utf8.GetBytes((ConvertTo-Json -InputObject $msg -Depth 20 -Compress) + "`n")
        $p.StandardInput.BaseStream.Write($bt, 0, $bt.Length); $p.StandardInput.BaseStream.Flush()
        $linha = $p.StandardOutput.ReadLine()
        $p.StandardInput.Close(); if (-not $p.WaitForExit(8000)) { $p.Kill() }
        [void]$err.Result
        return ($linha | ConvertFrom-Json).result
    }
    $cli3 = (ChamarMcp $banco3 'obter_cliente' ([ordered]@{ codigo_cliente = 2 })).content[0].text | ConvertFrom-Json
    $h3 = Get-HashArquivo $banco3
    $ex = Enter-CopiaExclusiva $banco3 $pastaBkp3 5
    try {
        Conferir 'com o banco livre, a cópia exclusiva é obtida e deixa a marca copiando.txt' ($ex.ok -and (Test-Path -LiteralPath (Get-MarcaCopiando $pastaBkp3))) $ex.motivo
        $erro = TentaGravar $banco3
        Conferir 'durante a cópia, gravação pelo ACE falha (abre só leitura)' ($erro -ne '' -and $erro -notmatch '^abrir:') $erro
        $res = ChamarMcp $banco3 'atualizar_cadastro_cliente' ([ordered]@{ codigo_cliente = 2; versao = [int]$cli3.versao; cidade = 'CIDADE DURANTE BACKUP'; autor = 'teste' })
        Conferir 'durante a cópia, o conector devolve a mensagem controlada do backup' ($res.isError -and $res.content[0].text -match 'Cópia de segurança') $res.content[0].text
        $lido = ChamarMcp $banco3 'obter_cliente' ([ordered]@{ codigo_cliente = 2 })
        Conferir 'durante a cópia, a leitura continua' (-not $lido.isError) $lido.content[0].text
        $tmpCopia = Join-Path $tmp 'copia-fluxo.accdb'
        Copy-FluxoParaArquivo $ex.fluxo $tmpCopia
        Conferir 'a cópia sai do próprio fluxo, igual ao original' ((Get-HashArquivo $tmpCopia) -eq $h3)
    } finally { Exit-CopiaExclusiva $ex }
    Conferir 'ao sair, a marca some' (-not (Test-Path -LiteralPath (Get-MarcaCopiando $pastaBkp3)))
    Conferir 'nada foi gravado durante a cópia (hash igual)' ((Get-HashArquivo $banco3) -eq $h3)
    Conferir 'depois da cópia a gravação volta a funcionar' ((TentaGravar $banco3) -eq '')
    # quem está conectado impede a captura
    $aberta = New-Object System.Data.OleDb.OleDbConnection('Provider=' + (Get-ProvedorAce) + ';Data Source=' + $banco3 + ';OLE DB Services=-4;')
    $aberta.Open()
    try {
        $ex2 = Enter-CopiaExclusiva $banco3 $pastaBkp3 2
        Conferir 'com alguém conectado, a cópia exclusiva não é obtida e a marca não fica' (-not $ex2.ok -and -not (Test-Path -LiteralPath (Get-MarcaCopiando $pastaBkp3))) $ex2.motivo
        Exit-CopiaExclusiva $ex2
        $r = Executar (Join-Path $b3 'SISTEMA\DADOS\operacao\scripts\backup-banco.ps1') @('-Modo', 'Diario', '-EsperaExclusivaSeg', '2')
        Conferir 'backup Diario com alguém conectado sai com 2 (sem cópia)' ($r.codigo -eq 2) $r.saida
    } finally { $aberta.Close(); $aberta.Dispose() }
    $r = Executar (Join-Path $b3 'SISTEMA\DADOS\operacao\scripts\backup-banco.ps1') @('-Modo', 'Diario')
    $copias3 = @(Get-ChildItem -LiteralPath $pastaBkp3 -Filter 'crm_diario_*.accdb' -File)
    $man3 = if ($copias3.Count -eq 1) { Get-Content -LiteralPath ($copias3[0].FullName + '.manifesto.json') -Raw -Encoding UTF8 | ConvertFrom-Json } else { $null }
    Conferir 'backup Diario livre: cópia completa com consistência exclusiva e validada' ($r.codigo -eq 0 -and $null -ne $man3 -and $man3.consistencia -eq 'exclusiva' -and (Test-Path -LiteralPath ($copias3[0].FullName + '.validado'))) ($r.saida + ' | pasta: ' + ((Get-ChildItem -LiteralPath $pastaBkp3 -File | ForEach-Object { $_.Name }) -join ', '))
    $r = Executar (Join-Path $b3 'SISTEMA\DADOS\operacao\scripts\backup-banco.ps1') @('-Modo', 'Periodico')
    $simples = @(Get-ChildItem -LiteralPath $pastaBkp3 -Filter 'crm_periodico_*.accdb' -File)
    $manS = if ($simples.Count -eq 1) { Get-Content -LiteralPath ($simples[0].FullName + '.manifesto.json') -Raw -Encoding UTF8 | ConvertFrom-Json } else { $null }
    Conferir 'cópia simples marcada como não verificada' ($r.codigo -eq 0 -and $null -ne $manS -and $manS.consistencia -eq 'nao-verificada') $r.saida

    Write-Host '--- DEV-03: quatro disparos simultâneos para a mesma janela = uma captura'
    $b4 = NovaBase 'base-janela'
    $pastaBkp4 = Join-Path $b4 'SISTEMA\DADOS\execucao\backup'
    $fazer4 = Join-Path $b4 'SISTEMA\DADOS\operacao\scripts\backup-banco.ps1'
    $procs = @()
    for ($i = 1; $i -le 4; $i++) {
        $procs += Start-Process -FilePath $ps -PassThru -WindowStyle Hidden -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"' + $fazer4 + '"'), '-Modo', 'Agendado', '-Janela', '20261019-0900', '-EsperaTravaSeg', '180', '-EsperaExclusivaSeg', '30')
    }
    foreach ($p in $procs) { if (-not $p.WaitForExit(240000)) { try { $p.Kill() } catch { } } }
    $codigos = @($procs | ForEach-Object { $_.ExitCode })
    $copias4 = @(Get-ChildItem -LiteralPath $pastaBkp4 -Filter 'crm_*.accdb' -File)
    $logs4 = @(Get-ChildItem -LiteralPath (Join-Path $pastaBkp4 'logs') -Filter 'backup-*.log' -File -ErrorAction SilentlyContinue)
    $atendidas = @($logs4 | Where-Object { (Get-Content -LiteralPath $_.FullName -Raw -Encoding UTF8) -match 'janela 20261019-0900 atendida' }).Count
    Conferir 'os quatro disparos terminam com 0' (@($codigos | Where-Object { $_ -ne 0 }).Count -eq 0) ($codigos -join ',')
    Conferir 'uma única captura na janela' ($copias4.Count -eq 1 -and (Get-CopiaDaJanela $pastaBkp4 '20261019-0900') -ne '') (($copias4 | ForEach-Object { $_.Name }) -join ',')
    Conferir 'os outros três registram "janela atendida", cada um no seu log' ($atendidas -eq 3 -and $logs4.Count -eq 4) ('atendidas=' + $atendidas + ' logs=' + $logs4.Count)
    Conferir 'nenhum .tmp, trava ou marca sobrou' (@(Get-ChildItem -LiteralPath $pastaBkp4 -File | Where-Object { $_.Name -match '\.tmp$|^backup\.lock$|^copiando\.txt$' }).Count -eq 0)

    Write-Host '--- DEV-06 parte A: diagnóstico do conector (retenção pelo nome, falha de escrita no crm_status)'
    $b5 = NovaBase 'base-diag'
    $banco5 = Join-Path $b5 'crm_zapromaq.accdb'
    $logs5 = Join-Path $b5 'IA\logs'
    if (Test-Path -LiteralPath $logs5) { Remove-Item -LiteralPath $logs5 -Recurse -Force }
    New-Item -ItemType Directory -Path $logs5 -Force | Out-Null
    $est = ([Environment]::MachineName -replace '[^A-Za-z0-9_-]', '_')
    $mesAtual = (Get-Date -Day 1).Date
    $vencido = Join-Path $logs5 ('mcp-' + $est + '-' + $mesAtual.AddMonths(-2).ToString('yyyyMMdd') + '.log')
    $anterior = Join-Path $logs5 ('mcp-' + $est + '-' + $mesAtual.AddMonths(-1).ToString('yyyyMMdd') + '.log')
    foreach ($f in $vencido, $anterior) { [IO.File]::WriteAllText($f, 'x') }
    (Get-Item -LiteralPath $vencido).LastWriteTime = Get-Date   # a data que vale é a do NOME
    $st = (ChamarMcp $banco5 'crm_status' ([ordered]@{})).content[0].text | ConvertFrom-Json
    Conferir 'retenção: arquivo de dois meses atrás sai (pela data do nome)' (-not (Test-Path -LiteralPath $vencido))
    Conferir 'retenção: o do mês anterior fica' (Test-Path -LiteralPath $anterior)
    Conferir 'crm_status traz diagnostico sem falhas' ($null -ne $st.diagnostico -and [int]$st.diagnostico.falhas -eq 0 -and [string]$st.diagnostico.arquivo -match ('mcp-' + $est)) ($st.diagnostico | ConvertTo-Json -Compress)
    $hoje5 = Join-Path $logs5 ('mcp-' + $est + '-' + (Get-Date -Format 'yyyyMMdd') + '.log')
    $linhas5 = @(Get-Content -LiteralPath $hoje5 -Encoding UTF8)
    Conferir 'linha do diagnóstico: instante com fuso, estação\usuário, versão e ferramenta' (@($linhas5 | Where-Object { $_ -match '^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\.\d{3} [+-]\d{2}:\d{2} \[\d+\] \[.+\\.+ v\d+\.\d+\.\d+ crm_status\]' }).Count -ge 1) ($linhas5 -join ' | ')
    $res = ChamarMcp $banco5 'obter_atendimento' ([ordered]@{ codigo = 'AT-0001-0001' })
    $linhas5 = @(Get-Content -LiteralPath $hoje5 -Encoding UTF8)
    Conferir 'linha do diagnóstico traz o código do registro' (@($linhas5 | Where-Object { $_ -match 'obter_atendimento codigo=AT-0001-0001' }).Count -ge 1) ($linhas5 -join ' | ')
    Set-ItemProperty -LiteralPath $hoje5 -Name IsReadOnly -Value $true
    try {
        $res = ChamarMcp $banco5 'crm_status' ([ordered]@{})
        $st = $res.content[0].text | ConvertFrom-Json
        Conferir 'diagnóstico sem poder gravar: a operação responde e o crm_status mostra a falha' ((-not $res.isError) -and [int]$st.diagnostico.falhas -ge 1 -and [string]$st.diagnostico.ultima_falha -ne '') ($st.diagnostico | ConvertTo-Json -Compress)
    } finally { Set-ItemProperty -LiteralPath $hoje5 -Name IsReadOnly -Value $false }

    # FIM-DOS-CASOS
} finally { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue }

Write-Host ''
if ($script:falhas -gt 0) { Write-Host ('RESULTADO: ' + $script:falhas + ' falha(s)'); exit 1 }
Write-Host 'TUDO CERTO.'
exit 0
