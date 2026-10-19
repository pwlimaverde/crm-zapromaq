<#
  testar-preparacao-rede.ps1 - testes da preparação da rede (spec 004), com ACE e sem Excel.

  Tudo acontece em pasta temporária e com bancos FICTÍCIOS: nada de BASE é tocado.
  Cada caso que recusa uma operação confere também que o banco ficou byte a byte igual (SHA-256).
#>
$ErrorActionPreference = 'Stop'
$dados = Split-Path -Parent $PSScriptRoot
$raizBase = Split-Path -Parent (Split-Path -Parent $dados)
. (Join-Path $dados 'lib\Comum.ps1')
. (Join-Path $dados 'lib\AtualizacaoBanco.ps1')
. (Join-Path $dados 'build\lib\Pacote.ps1')

$ps = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$esquema = Join-Path $dados 'banco\esquema\criar-banco.ps1'
$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('crm-tpr-' + $PID)
New-Item -ItemType Directory -Path $tmp -Force | Out-Null

$script:falhas = 0
function Conferir([string]$nome, [bool]$ok, [string]$detalhe = '') {
    if ($ok) { Write-Host ('  OK    ' + $nome) -ForegroundColor Green }
    else { $script:falhas++; Write-Host ('  FALHA ' + $nome + '  ' + $detalhe) -ForegroundColor Red }
}
# Roda um script em outro PowerShell 5.1 e devolve código e saída (stderr incluído).
function Executar([string]$script, [string[]]$argumentos) {
    $pref = $ErrorActionPreference; $ErrorActionPreference = 'Continue'
    try {
        $saida = @(& $ps -NoProfile -ExecutionPolicy Bypass -File $script @argumentos 2>&1 | ForEach-Object { [string]$_ })
        return [pscustomobject]@{ codigo = $LASTEXITCODE; saida = ($saida -join "`n") }
    } finally { $ErrorActionPreference = $pref }
}
function NovoBanco([string]$caminho, [string]$ambiente) {
    $r = Executar $esquema @('-Banco', $caminho, '-Ambiente', $ambiente)
    if ($r.codigo -ne 0) { throw ('não criou o banco ' + $caminho + ': ' + $r.saida) }
}

# BASE mínima em pasta temporária: o suficiente para rodar os scripts de IA\teste e SISTEMA\DADOS.
# Devolve a raiz da BASE temporária (sem banco: cada caso cria o seu).
function NovaBaseTemporaria([string]$nome) {
    $b = Join-Path $tmp $nome
    foreach ($rel in @('IA\mcp', 'IA\teste', 'SISTEMA\DADOS\lib', 'SISTEMA\DADOS\banco', 'SISTEMA\DADOS\build\lib', 'SISTEMA\DADOS\testes', 'SISTEMA\DADOS\operacao')) {
        $dest = Join-Path $b $rel
        New-Item -ItemType Directory -Path (Split-Path -Parent $dest) -Force | Out-Null
        Copy-Item -LiteralPath (Join-Path $raizBase $rel) -Destination $dest -Recurse
    }
    Copy-Item -LiteralPath (Join-Path $dados 'VERSAO.txt') -Destination (Join-Path $b 'SISTEMA\DADOS\VERSAO.txt')
    return $b
}
# .xlsm falso (só o suficiente para Get-TextoCelulaXlsx): aba Config com $texto em B2.
function NovoXlsmFalso([string]$caminho, [string]$texto) {
    if (Test-Path -LiteralPath $caminho) { Remove-Item -LiteralPath $caminho -Force }
    $zip = [System.IO.Compression.ZipFile]::Open($caminho, 'Create')
    function Parte($nome, $conteudo) {
        $e = $zip.CreateEntry($nome); $w = New-Object System.IO.StreamWriter($e.Open(), (New-Object System.Text.UTF8Encoding($false)))
        $w.Write($conteudo); $w.Dispose()
    }
    Parte 'xl/workbook.xml' '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Config" sheetId="1" r:id="rId1"/></sheets></workbook>'
    Parte 'xl/_rels/workbook.xml.rels' '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>'
    Parte 'xl/worksheets/sheet1.xml' ('<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="2"><c r="B2" t="inlineStr"><is><t>' + [Security.SecurityElement]::Escape($texto) + '</t></is></c></row></sheetData></worksheet>')
    $zip.Dispose()
}

try {
    Write-Host '--- F1: -Recriar só apaga banco de TESTE'
    $prod = Join-Path $tmp 'producao.accdb'
    NovoBanco $prod 'PRODUCAO'
    $h = Get-HashArquivo $prod
    $r = Executar $esquema @('-Banco', $prod, '-Recriar')
    Conferir 'criar-banco -Recriar recusa banco PRODUCAO' ($r.codigo -ne 0 -and $r.saida -match 'PRODUCAO') $r.saida
    Conferir 'banco PRODUCAO ficou igual (hash)' ((Get-HashArquivo $prod) -eq $h)

    $lixo = Join-Path $tmp 'lixo.accdb'
    [IO.File]::WriteAllText($lixo, 'isto não é um banco')
    $h2 = Get-HashArquivo $lixo
    $r = Executar $esquema @('-Banco', $lixo, '-Recriar')
    Conferir 'arquivo que não abre é recusado' ($r.codigo -ne 0) $r.saida
    Conferir 'arquivo ilegível ficou igual (hash)' ((Get-HashArquivo $lixo) -eq $h2)

    $teste = Join-Path $tmp 'teste.accdb'
    NovoBanco $teste 'TESTE'
    $r = Executar $esquema @('-Banco', $teste, '-Recriar')
    Conferir 'banco TESTE é recriado' ($r.codigo -eq 0) $r.saida

    $novo = Join-Path $tmp 'inexistente.accdb'
    $r = Executar $esquema @('-Banco', $novo, '-Recriar')
    Conferir 'arquivo inexistente é criado' ($r.codigo -eq 0 -and (Test-Path -LiteralPath $novo)) $r.saida

    $msg = Assert-PodeSubstituirBanco $prod
    Conferir 'Assert-PodeSubstituirBanco devolve o motivo para PRODUCAO' ($msg -match 'PRODUCAO')
    Conferir 'Assert-PodeSubstituirBanco aceita TESTE' ((Assert-PodeSubstituirBanco $teste) -eq '')
    Conferir 'Assert-PodeSubstituirBanco aceita inexistente' ((Assert-PodeSubstituirBanco (Join-Path $tmp 'nao-existe.accdb')) -eq '')

    Write-Host '--- F2: testes que gravam recusam banco que não seja TESTE'
    $b2 = NovaBaseTemporaria 'base-f2'
    $bancoB2 = Join-Path $b2 'crm_zapromaq.accdb'
    NovoBanco $bancoB2 'PRODUCAO'
    $h = Get-HashArquivo $bancoB2
    $r = Executar (Join-Path $b2 'IA\teste\testar-mcp.ps1') @()
    Conferir 'testar-mcp recusa banco PRODUCAO' ($r.codigo -ne 0 -and $r.saida -match 'PRODUCAO') $r.saida
    Conferir 'banco PRODUCAO ficou igual depois do testar-mcp (hash)' ((Get-HashArquivo $bancoB2) -eq $h)

    $xlsm = Join-Path $tmp 'falso.xlsm'
    NovoXlsmFalso $xlsm $bancoB2
    $r = Executar (Join-Path $dados 'testes\testar-uso.ps1') @('-Arquivo', $xlsm)
    Conferir 'testar-uso (completo) recusa front apontando para banco PRODUCAO' ($r.codigo -ne 0 -and $r.saida -match 'PRODUCAO') $r.saida
    Conferir 'banco PRODUCAO ficou igual depois do testar-uso (hash)' ((Get-HashArquivo $bancoB2) -eq $h)

    Write-Host '--- F3: migração estrutural só pela montagem (cópia de trabalho)'
    $b3 = NovaBaseTemporaria 'base-f3'
    $mig3 = Join-Path $b3 'SISTEMA\DADOS\banco\aplicar-migracoes.ps1'
    $p3 = Join-Path $b3 'crm_zapromaq.accdb'
    NovoBanco $p3 'PRODUCAO'
    $h = Get-HashArquivo $p3
    $r = Executar $mig3 @('-Banco', $p3)
    Conferir 'PRODUCAO 1.0 sem a opção: migração estrutural é recusada' ($r.codigo -ne 0 -and $r.saida -match 'estrutural') $r.saida
    Conferir 'PRODUCAO recusada ficou igual, sem nem a 001/002 (hash)' ((Get-HashArquivo $p3) -eq $h)

    $fora = Join-Path $tmp 'fora.accdb'
    Copy-Item -LiteralPath $p3 -Destination $fora
    $hf = Get-HashArquivo $fora
    $r = Executar $mig3 @('-Banco', $fora, '-CopiaDeTrabalho')
    Conferir '-CopiaDeTrabalho fora de execucao\publicacao e execucao\ensaio é recusado' ($r.codigo -ne 0 -and $r.saida -match 'execucao') $r.saida
    Conferir 'cópia fora de execucao ficou igual (hash)' ((Get-HashArquivo $fora) -eq $hf)
    $r = Executar $mig3 @('-Banco', $p3, '-CopiaDeTrabalho')
    Conferir '-CopiaDeTrabalho apontando para o banco de produção é recusado' ($r.codigo -ne 0) $r.saida
    Conferir 'banco de produção ficou igual (hash)' ((Get-HashArquivo $p3) -eq $h)

    $pastaEns = Join-Path $b3 'SISTEMA\DADOS\execucao\ensaio'
    New-Item -ItemType Directory -Path $pastaEns -Force | Out-Null
    $copia3 = Join-Path $pastaEns 'crm_zapromaq.accdb'
    Copy-Item -LiteralPath $p3 -Destination $copia3
    $r = Executar $mig3 @('-Banco', $copia3, '-CopiaDeTrabalho')
    $cn3 = Open-Banco $copia3
    try { $esq3 = Get-VersaoEsquemaBanco $cn3 } finally { $cn3.Close(); $cn3.Dispose() }
    Conferir 'cópia de trabalho em execucao\ensaio chega ao esquema 1.1' ($r.codigo -eq 0 -and $esq3 -eq '1.1') ($r.saida + ' esquema=' + $esq3)
    Conferir 'produção continua igual depois da migração na cópia (hash)' ((Get-HashArquivo $p3) -eq $h)

    $t3 = Join-Path $tmp 'teste-1-0.accdb'
    NovoBanco $t3 'TESTE'
    $r = Executar $mig3 @('-Banco', $t3)
    $cn3 = Open-Banco $t3
    try { $esq3 = Get-VersaoEsquemaBanco $cn3 } finally { $cn3.Close(); $cn3.Dispose() }
    Conferir 'banco TESTE 1.0 migra no lugar até o 1.1' ($r.codigo -eq 0 -and $esq3 -eq '1.1') ($r.saida + ' esquema=' + $esq3)

    Write-Host '--- F4: ensaio (montar-frontend -Ensaio, passos 1 a 8, sem Excel e sem trocar o banco)'
    # banco "como a rede": esquema 1.0, ambiente PRODUCAO, sem nenhum registro migracao.*, com os casos da 003
    $b4 = NovaBaseTemporaria 'base-f4'
    New-Item -ItemType Directory -Path (Join-Path $b4 'SISTEMA\DADOS\fonte') -Force | Out-Null
    Copy-Item -LiteralPath (Join-Path $dados 'fonte\layout') -Destination (Join-Path $b4 'SISTEMA\DADOS\fonte\layout') -Recurse
    Copy-Item -LiteralPath (Join-Path $dados 'build\montar-frontend.ps1') -Destination (Join-Path $b4 'SISTEMA\DADOS\build\montar-frontend.ps1')
    $p4 = Join-Path $b4 'crm_zapromaq.accdb'
    & $ps -NoProfile -ExecutionPolicy Bypass -File (Join-Path $raizBase 'IA\teste\criar-banco-teste.ps1') -Banco $p4 -Recriar -CasosDaMigracao | Out-Null
    $cn4 = Open-Banco $p4
    try {
        Set-ConfigBanco $cn4 'ambiente' 'PRODUCAO' $null
        $migrAntes = [int](Invoke-Escalar $cn4 "SELECT COUNT(*) FROM config WHERE chave LIKE 'migracao.%'" @() $null)
        $esqAntes = Get-VersaoEsquemaBanco $cn4
    } finally { $cn4.Close(); $cn4.Dispose() }
    Conferir 'banco "como a rede": 1.0, PRODUCAO, sem migracao.*' ($esqAntes -eq '1.0' -and $migrAntes -eq 0 -and (Get-AmbienteDoBanco $p4) -eq 'PRODUCAO')
    $h = Get-HashArquivo $p4
    $r = Executar (Join-Path $b4 'SISTEMA\DADOS\build\montar-frontend.ps1') @('-Ensaio')
    Conferir 'montar-frontend -Ensaio termina com 0 e diz ENSAIO APROVADO' ($r.codigo -eq 0 -and $r.saida -match 'ENSAIO APROVADO') $r.saida
    Conferir 'ensaio: banco de produção ficou igual (hash)' ((Get-HashArquivo $p4) -eq $h)
    $ensaio4 = Join-Path $b4 'SISTEMA\DADOS\execucao\ensaio'
    $cop4 = Join-Path $ensaio4 'copia\crm_zapromaq.accdb'
    Conferir 'ensaio: a cópia existe em execucao\ensaio' (Test-Path -LiteralPath $cop4)
    if (Test-Path -LiteralPath $cop4) {
        $cn4 = Open-Banco $cop4
        try {
            $nm = [int](Invoke-Escalar $cn4 "SELECT COUNT(*) FROM config WHERE chave IN ('migracao.001','migracao.002','migracao.003')" @() $null)
            $esq4 = Get-VersaoEsquemaBanco $cn4
        } finally { $cn4.Close(); $cn4.Dispose() }
        Conferir 'ensaio: 001, 002 e 003 registradas NA CÓPIA e esquema 1.1' ($nm -eq 3 -and $esq4 -eq '1.1') ('migracoes=' + $nm + ' esquema=' + $esq4)
    }
    $rel4 = @(Get-ChildItem -LiteralPath $ensaio4 -Filter 'relatorio-*.txt' -ErrorAction SilentlyContinue)
    Conferir 'ensaio: relatório gerado' ($rel4.Count -eq 1)
    if ($rel4.Count -eq 1) {
        $txt4 = Get-Content -LiteralPath $rel4[0].FullName -Raw -Encoding UTF8
        Conferir 'relatório traz esquema antes/depois, hash e contagens' ($txt4 -match '1\.0' -and $txt4 -match '1\.1' -and $txt4 -match $h.Substring(0, 12) -and $txt4 -match 'clientes')
    }
    Conferir 'ensaio: execucao\backup\antes-da-publicacao não foi criada' (-not (Test-Path -LiteralPath (Join-Path $b4 'SISTEMA\DADOS\execucao\backup\antes-da-publicacao')))
    # trava ocupada: o ensaio recusa (a cópia não seria consistente)
    $conexaoAberta = Open-Banco $p4   # uma conexão de verdade segura o .laccdb (o ACE apaga trava falsa)
    try {
        $r = Executar (Join-Path $b4 'SISTEMA\DADOS\build\montar-frontend.ps1') @('-Ensaio')
        Conferir 'ensaio com banco em uso (.laccdb) é recusado, sem alterar nada' ($r.codigo -ne 0 -and $r.saida -match 'conectado') $r.saida
    } finally { $conexaoAberta.Close(); $conexaoAberta.Dispose() }
    Conferir 'banco em uso: produção ficou igual (hash)' ((Get-HashArquivo $p4) -eq $h)

    Write-Host '--- F5: impressão digital do banco (o teste de uso da publicação não pode gravar)'
    $i1 = Get-ImpressaoDoBanco $cop4
    $i2 = Get-ImpressaoDoBanco $cop4
    Conferir 'impressão é igual no mesmo banco' ($i1 -eq $i2 -and $i1 -match 'clientes' -and $i1 -match 'log_alteracoes') $i1
    $cnI = Open-Banco $cop4
    try { [void](Invoke-Comando $cnI 'UPDATE clientes SET versao = versao + 1 WHERE id = (SELECT MIN(id) FROM clientes)' @() $null) } finally { $cnI.Close(); $cnI.Dispose() }
    $i3 = Get-ImpressaoDoBanco $cop4
    Conferir 'impressão muda depois de UPDATE versao+1' ($i3 -ne $i1)
    $cnI = Open-Banco $cop4
    try { [void](Invoke-Comando $cnI 'UPDATE clientes SET versao = versao - 1 WHERE id = (SELECT MIN(id) FROM clientes)' @() $null) } finally { $cnI.Close(); $cnI.Dispose() }
    Conferir 'impressão volta ao valor original ao desfazer' ((Get-ImpressaoDoBanco $cop4) -eq $i1)
    $cnI = Open-Banco $cop4
    try { [void](Invoke-Comando $cnI "INSERT INTO log_alteracoes (tabela, id_registro, campo, acao, usuario, quando, origem) VALUES ('clientes', 1, 'x', 'TESTE', 'tpr', Now(), 'FRONT')" @() $null) } finally { $cnI.Close(); $cnI.Dispose() }
    Conferir 'impressão muda depois de linha nova no log' ((Get-ImpressaoDoBanco $cop4) -ne $i1)

    Write-Host '--- F7: config.versao_front é gravada na CÓPIA, antes da troca'
    $vfProd = { $c = Open-BancoLeitura $p4; try { return [string](Invoke-Escalar $c "SELECT valor FROM config WHERE chave='versao_front'" @() $null) } finally { $c.Close(); $c.Dispose() } }
    $vfAntes = & $vfProd
    Set-VersaoFrontNaCopia $cop4 '9.9'
    $cnV = Open-BancoLeitura $cop4
    try { $vfCopia = [string](Invoke-Escalar $cnV "SELECT valor FROM config WHERE chave='versao_front'" @() $null) } finally { $cnV.Close(); $cnV.Dispose() }
    Conferir 'versao_front gravada na cópia' ($vfCopia -eq '9.9') $vfCopia
    Conferir 'produção intacta: hash igual e versao_front antiga' ((Get-HashArquivo $p4) -eq $h -and (& $vfProd) -eq $vfAntes)

    Write-Host '--- F8: backup validado e restauração verificável'
    . (Join-Path $dados 'lib\Backup.ps1')
    Conferir 'ensaio: o backup da janela tem manifesto' (Test-Path -LiteralPath (Join-Path $ensaio4 'backup\crm_zapromaq.accdb.manifesto.json'))
    $b8 = NovaBaseTemporaria 'base-f8'
    $ref8 = Join-Path $b8 'crm_zapromaq.accdb'
    & $ps -NoProfile -ExecutionPolicy Bypass -File (Join-Path $raizBase 'IA\teste\criar-banco-teste.ps1') -Banco $ref8 -Recriar | Out-Null
    $pastaBk8 = Join-Path $b8 'SISTEMA\DADOS\execucao\backup'
    New-Item -ItemType Directory -Path $pastaBk8 -Force | Out-Null
    $bk8 = Join-Path $pastaBk8 'crm_diario_20261005-0800_EST1.accdb'
    Copy-Item -LiteralPath $ref8 -Destination $bk8
    $hBk8 = Get-HashArquivo $bk8
    $man8 = New-ManifestoBackup $bk8 $ref8 'diario'
    $mj = Get-Content -LiteralPath ($bk8 + '.manifesto.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    Conferir 'manifesto: instante, sha256, esquema e contagens' ($mj.sha256 -eq $hBk8 -and $mj.esquema -eq '1.1' -and $mj.contagens.clientes -ge 1 -and [string]$mj.instante_utc -ne '')
    Conferir 'o arquivo de backup não foi aberto nem alterado (hash)' ((Get-HashArquivo $bk8) -eq $hBk8)
    $val8 = Test-BackupRestauravel $bk8
    Conferir 'backup restaurável é aprovado' ($val8.ok -and $val8.problemas.Count -eq 0) ($val8.problemas -join '; ')
    Conferir 'validar não alterou o arquivo de backup (hash)' ((Get-HashArquivo $bk8) -eq $hBk8)

    $rest8 = Join-Path $b8 'SISTEMA\DADOS\operacao\scripts\restaurar-banco.ps1'
    $r = Executar $rest8 @('-Backup', $bk8, '-Referencia', $ref8)
    Conferir 'restauração aprovada termina com 0' ($r.codigo -eq 0) $r.saida
    # contagem da referência adulterada (uma linha de cliente a mais): diverge
    $cn8 = Open-Banco $ref8
    try { [void](Invoke-Comando $cn8 "INSERT INTO clientes (estagio, empresa, ativo, versao) VALUES ('Pré-cliente', 'EXTRA FICTICIA', True, 1)" @() $null) } finally { $cn8.Close(); $cn8.Dispose() }
    $r = Executar $rest8 @('-Backup', $bk8, '-Referencia', $ref8)
    Conferir 'contagem divergente da referência: termina com 1' ($r.codigo -eq 1) $r.saida
    # manifesto adulterado: reprovado
    $txt = (Get-Content -LiteralPath ($bk8 + '.manifesto.json') -Raw -Encoding UTF8) -replace '"clientes":\s*\d+', '"clientes": 999'
    [IO.File]::WriteAllText(($bk8 + '.manifesto.json'), $txt, (New-Object System.Text.UTF8Encoding($false)))
    $val8 = Test-BackupRestauravel $bk8
    Conferir 'manifesto adulterado: backup reprovado' (-not $val8.ok -and $val8.problemas.Count -ge 1)
    $r = Executar $rest8 @('-Backup', $bk8)
    Conferir 'manifesto adulterado: restauração termina com 1' ($r.codigo -eq 1) $r.saida

    # "mais recente": pelo nome (ou manifesto), não pelo LastWriteTime
    Remove-Item -LiteralPath ($bk8 + '.manifesto.json') -Force
    $velho8 = Join-Path $pastaBk8 'crm_diario_20261001-0800_EST1.accdb'
    Copy-Item -LiteralPath $bk8 -Destination $velho8
    (Get-Item -LiteralPath $velho8).LastWriteTime = (Get-Date)
    (Get-Item -LiteralPath $bk8).LastWriteTime = (Get-Date).AddDays(-30)
    $esc = Get-BackupMaisRecente $pastaBk8
    Conferir 'mais recente pelo nome, mesmo com LastWriteTime invertido' ($esc.Name -eq 'crm_diario_20261005-0800_EST1.accdb') $esc.Name
    $r = Executar $rest8 @()
    Conferir 'restaurar sem -Backup usa o mais recente pelo nome' ($r.saida -match 'crm_diario_20261005') $r.saida

    Write-Host '--- F9: conector exige o esquema 1.1 e recusa com mensagem clara no 1.0'
    # cliente MCP mínimo (protocolo legado): sobe o servidor sobre um banco, faz initialize e uma chamada
    function ChamarMcp([string]$banco, [string]$ferramenta, $argumentos) {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $ps
        $psi.Arguments = '-NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "' + (Join-Path $raizBase 'IA\mcp\servidor.ps1') + '" -Banco "' + $banco + '"'
        $psi.UseShellExecute = $false; $psi.RedirectStandardInput = $true; $psi.RedirectStandardOutput = $true; $psi.RedirectStandardError = $true
        $psi.StandardOutputEncoding = New-Object System.Text.UTF8Encoding($false); $psi.CreateNoWindow = $true
        $p = [System.Diagnostics.Process]::Start($psi)
        $err = $p.StandardError.ReadToEndAsync()
        $utf8 = New-Object System.Text.UTF8Encoding($false)
        function Linha($o) { $b = $utf8.GetBytes((ConvertTo-Json -InputObject $o -Depth 20 -Compress) + "`n"); $p.StandardInput.BaseStream.Write($b, 0, $b.Length); $p.StandardInput.BaseStream.Flush() }
        try {
            Linha ([ordered]@{ jsonrpc = '2.0'; id = 1; method = 'initialize'; params = [ordered]@{ protocolVersion = '2025-06-18'; capabilities = [ordered]@{}; clientInfo = [ordered]@{ name = 'tpr'; version = '1.0' } } })
            [void]$p.StandardOutput.ReadLine()
            Linha ([ordered]@{ jsonrpc = '2.0'; method = 'notifications/initialized' })
            Linha ([ordered]@{ jsonrpc = '2.0'; id = 2; method = 'tools/call'; params = [ordered]@{ name = $ferramenta; arguments = $argumentos } })
            return ($p.StandardOutput.ReadLine() | ConvertFrom-Json).result
        } finally { try { $p.StandardInput.Close() } catch { }; if (-not $p.WaitForExit(5000)) { $p.Kill() } }
    }
    $v10 = Join-Path $tmp 'conector-1-0.accdb'
    & $ps -NoProfile -ExecutionPolicy Bypass -File (Join-Path $raizBase 'IA\teste\criar-banco-teste.ps1') -Banco $v10 -Recriar -CasosDaMigracao | Out-Null
    $hv = Get-HashArquivo $v10
    $res = ChamarMcp $v10 'crm_status' ([ordered]@{})
    $st = $res.content[0].text | ConvertFrom-Json
    Conferir 'banco 1.0: crm_status responde, com atualizacao_pendente=true e gravacao_liberada=false' (-not $res.isError -and $st.atualizacao_pendente -eq $true -and $st.gravacao_liberada -eq $false) ($res.content[0].text)
    Conferir 'crm_status informa esquemas_suportados = 1.1' (@($st.esquemas_suportados) -join ',' -eq '1.1') (@($st.esquemas_suportados) -join ',')
    Conferir 'crm_status informa versao_servidor_mcp 1.5.0' ($st.versao_servidor_mcp -eq '1.5.0') $st.versao_servidor_mcp
    $res = ChamarMcp $v10 'obter_atendimento' ([ordered]@{ codigo = 'AT-0001-0001' })
    Conferir 'banco 1.0: leitura recusada com a mensagem do esquema 1.0' ($res.isError -eq $true -and $res.content[0].text -match 'esquema 1\.0') ($res.content[0].text)
    $res = ChamarMcp $v10 'atualizar_atendimento' ([ordered]@{ codigo = 'AT-0001-0001'; versao = 1; tentativas = 2; autor = 'tpr' })
    Conferir 'banco 1.0: gravação recusada com a mensagem do esquema 1.0' ($res.isError -eq $true -and $res.content[0].text -match 'esquema 1\.0') ($res.content[0].text)
    Conferir 'banco 1.0: nada foi gravado (hash)' ((Get-HashArquivo $v10) -eq $hv)
    $res = ChamarMcp $ref8 'crm_status' ([ordered]@{})
    $st = $res.content[0].text | ConvertFrom-Json
    Conferir 'banco 1.1: crm_status normal, sem atualização pendente e gravação liberada' ($st.atualizacao_pendente -eq $false -and $st.gravacao_liberada -eq $true -and $st.versao_esquema -eq '1.1') ($res.content[0].text)

    Write-Host '--- 005: homologação declarada (config.homologacao) e pasta execucao\homologacao'
    $cnH = Open-BancoLeitura $teste
    try { $hom = [string](Invoke-Escalar $cnH "SELECT valor FROM config WHERE chave='homologacao'" @() $null) } finally { $cnH.Close(); $cnH.Dispose() }
    Conferir 'banco TESTE criado agora nasce com config.homologacao = 1' ($hom -eq '1') $hom
    Conferir 'banco de homologação é aceito pela trava' ((Assert-BancoDeHomologacao $teste) -eq '')
    Conferir 'banco PRODUCAO não é de homologação' ((Assert-BancoDeHomologacao $prod) -match 'PRODUCAO')
    $semHom = Join-Path $tmp 'teste-sem-homologacao.accdb'
    Copy-Item -LiteralPath $teste -Destination $semHom
    $cnH = Open-Banco $semHom
    try { [void](Invoke-Comando $cnH "DELETE FROM config WHERE chave='homologacao'" @() $null) } finally { $cnH.Close(); $cnH.Dispose() }
    Conferir 'TESTE sem config.homologacao não é de homologação' ((Assert-BancoDeHomologacao $semHom) -match 'homologacao')
    $b5 = NovaBaseTemporaria 'base-005'
    $bancoB5 = Join-Path $b5 'crm_zapromaq.accdb'
    Copy-Item -LiteralPath $semHom -Destination $bancoB5
    $hb5 = Get-HashArquivo $bancoB5
    $r = Executar (Join-Path $b5 'IA\teste\testar-mcp.ps1') @()
    Conferir 'testar-mcp recusa TESTE sem homologação declarada' ($r.codigo -ne 0 -and $r.saida -match 'homologacao') $r.saida
    Conferir 'banco sem homologação ficou igual (hash)' ((Get-HashArquivo $bancoB5) -eq $hb5)
    Conferir 'Get-PastaHomologacao cria execucao\homologacao' ((Get-PastaHomologacao $b5) -and (Test-Path -LiteralPath (Join-Path $b5 'SISTEMA\DADOS\execucao\homologacao')))

    Write-Host '--- 005: diagnóstico do conector (estação, usuário técnico, ferramenta; um arquivo por estação)'
    $res = ChamarMcp $ref8 'crm_status' ([ordered]@{})
    $logs = Join-Path $raizBase 'IA\logs'
    $arqDiag = Join-Path $logs ('mcp-' + $env:COMPUTERNAME + '-' + (Get-Date -Format 'yyyyMMdd') + '.log')
    Conferir 'existe um arquivo de diagnóstico por estação (mcp-<ESTACAO>-AAAAMMDD.log)' (Test-Path -LiteralPath $arqDiag) $arqDiag
    if (Test-Path -LiteralPath $arqDiag) {
        $ult = @(Get-Content -LiteralPath $arqDiag -Tail 30 -Encoding UTF8)
        Conferir 'a linha do diagnóstico traz estação\usuário e a ferramenta' (@($ult | Where-Object { $_ -match [regex]::Escape($env:COMPUTERNAME + '\' + $env:USERNAME) -and $_ -match 'crm_status' }).Count -ge 1) ($ult -join "`n")
    }

    Write-Host '--- 005: checar-ambiente (relatório, conector, IA, OneDrive)'
    $b6 = NovaBaseTemporaria 'base-005b'
    $r = Executar (Join-Path $b6 'SISTEMA\DADOS\testes\checar-ambiente.ps1') @()
    $relAmb = @(Get-ChildItem -LiteralPath (Join-Path $b6 'SISTEMA\DADOS\execucao\logs') -Filter 'ambiente-*.txt' -ErrorAction SilentlyContinue)
    Conferir 'checar-ambiente grava o relatório em execucao\logs' ($relAmb.Count -eq 1) $r.saida
    if ($relAmb.Count -eq 1) {
        $tx = Get-Content -LiteralPath $relAmb[0].FullName -Raw -Encoding UTF8
        Conferir 'relatório confere conector, pasta IA e OneDrive' ($tx -match 'Conector MCP' -and $tx -match 'Pasta IA' -and $tx -match 'OneDrive') $tx
    }

    Write-Host '--- 006: trava de backup com dono e sessão'
    $b7 = NovaBaseTemporaria 'base-006'
    $pastaT = Join-Path $b7 'trava-teste'
    New-Item -ItemType Directory -Path $pastaT -Force | Out-Null
    $libComum = Join-Path $dados 'lib\Comum.ps1'; $libAtu = Join-Path $dados 'lib\AtualizacaoBanco.ps1'
    $filho = Start-Process -FilePath $ps -PassThru -WindowStyle Hidden -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command',
        ('. "' + $libComum + '"; . "' + $libAtu + '"; $t = Enter-TravaBackup "' + $pastaT + '" "teste"; if ($t.ok) { Start-Sleep -Seconds 12 }; Exit-TravaBackup $t'))
    $esperou = 0; while (-not (Test-Path -LiteralPath (Join-Path $pastaT 'backup.lock')) -and $esperou -lt 40) { Start-Sleep -Milliseconds 500; $esperou++ }
    $t2 = Enter-TravaBackup $pastaT 'teste'
    Conferir 'segunda sessão com a trava ocupada é recusada, citando o dono' (-not $t2.ok -and $t2.motivo -match [string]$filho.Id) $t2.motivo
    try { Stop-Process -Id $filho.Id -Force } catch { }
    Start-Sleep -Milliseconds 800
    $t3 = Enter-TravaBackup $pastaT 'teste'
    Conferir 'trava abandonada (processo morto na mesma estação) é recuperada' ($t3.ok -and $t3.recuperada -eq $true) $t3.motivo
    Exit-TravaBackup $t3
    Conferir 'Exit-TravaBackup remove a trava' (-not (Test-Path -LiteralPath (Join-Path $pastaT 'backup.lock')))
    [IO.File]::WriteAllText((Join-Path $pastaT 'backup.lock'), (ConvertTo-Json @{ estacao = 'OUTRA-ESTACAO'; usuario = 'x'; pid = 4; sessao = 'a'; tipo = 'teste'; inicio = (Get-Date).ToString('o') }))
    $t4 = Enter-TravaBackup $pastaT 'teste'
    Conferir 'trava recente de outra estação é respeitada' (-not $t4.ok -and $t4.motivo -match 'OUTRA-ESTACAO') $t4.motivo
    [IO.File]::WriteAllText((Join-Path $pastaT 'backup.lock'), (ConvertTo-Json @{ estacao = 'OUTRA-ESTACAO'; usuario = 'x'; pid = 4; sessao = 'a'; tipo = 'teste'; inicio = (Get-Date).AddHours(-3).ToString('o') }))
    $t5 = Enter-TravaBackup $pastaT 'teste'
    Conferir 'trava de outra estação com mais de 60 min é recuperada' ($t5.ok -and $t5.recuperada) $t5.motivo
    Exit-TravaBackup $t5

    Write-Host '--- 006: backup com nome único, arquivo temporário promovido, manifesto, validado e log por execução'
    $banco7 = Join-Path $b7 'crm_zapromaq.accdb'
    & $ps -NoProfile -ExecutionPolicy Bypass -File (Join-Path $raizBase 'IA\teste\criar-banco-teste.ps1') -Banco $banco7 -Recriar | Out-Null
    $hProd7 = Get-HashArquivo $banco7
    $bkp7 = Join-Path $b7 'SISTEMA\DADOS\execucao\backup'
    $fazer = Join-Path $b7 'SISTEMA\DADOS\operacao\scripts\backup-banco.ps1'
    $r1 = Executar $fazer @('-Modo', 'Diario')
    $r2 = Executar $fazer @('-Modo', 'Diario')
    $copias7 = @(Get-ChildItem -LiteralPath $bkp7 -Filter 'crm_diario_*.accdb' -File)
    Conferir 'dois backups seguidos terminam com 0' ($r1.codigo -eq 0 -and $r2.codigo -eq 0) ($r1.saida + $r2.saida)
    Conferir 'nomes únicos (AAAAMMDD-HHMMSS) mesmo no mesmo minuto' ($copias7.Count -eq 2 -and $copias7[0].Name -match '^crm_diario_\d{8}-\d{6}_' ) (($copias7 | ForEach-Object { $_.Name }) -join ',')
    Conferir 'nenhum .tmp sobrou' (@(Get-ChildItem -LiteralPath $bkp7 -Filter '*.tmp' -File).Count -eq 0)
    Conferir 'cada cópia tem manifesto e marca de validado' (@($copias7 | Where-Object { (Test-Path -LiteralPath ($_.FullName + '.manifesto.json')) -and (Test-Path -LiteralPath ($_.FullName + '.validado')) }).Count -eq 2)
    Conferir 'um log por execução em execucao\backup\logs' (@(Get-ChildItem -LiteralPath (Join-Path $bkp7 'logs') -Filter 'backup-*.log' -File -ErrorAction SilentlyContinue).Count -ge 2)
    Conferir 'o backup.log consolidado existe' (Test-Path -LiteralPath (Join-Path $bkp7 'backup.log'))
    Conferir 'o banco ficou igual depois dos backups (hash)' ((Get-HashArquivo $banco7) -eq $hProd7)

    Write-Host '--- 006: retenção nunca apaga o último backup validado'
    $pr = Join-Path $tmp 'retencao'
    New-Item -ItemType Directory -Path $pr -Force | Out-Null
    $velhoV = Join-Path $pr 'crm_diario_20260101-080000_EST1.accdb'
    $velhoN = Join-Path $pr 'crm_diario_20260102-080000_EST1.accdb'
    $novoN = Join-Path $pr ('crm_periodico_' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '_EST1.accdb')
    foreach ($f in $velhoV, $velhoN, $novoN) { [IO.File]::WriteAllText($f, 'x'); [IO.File]::WriteAllText(($f + '.manifesto.json'), '{}') }
    [IO.File]::WriteAllText(($velhoV + '.validado'), (ConvertTo-Json @{ validado_utc = '2026-01-01T11:00:00Z' }))
    $null = Executar $fazer @('-PastaTeste', $pr)
    Conferir 'o único validado fica, mesmo velho, enquanto o novo não foi validado' (Test-Path -LiteralPath $velhoV)
    Conferir 'o velho não validado sai, com o manifesto dele' (-not (Test-Path -LiteralPath $velhoN) -and -not (Test-Path -LiteralPath ($velhoN + '.manifesto.json')))
    [IO.File]::WriteAllText(($novoN + '.validado'), (ConvertTo-Json @{ validado_utc = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ') }))
    $null = Executar $fazer @('-PastaTeste', $pr)
    Conferir 'com um validado novo, o velho validado sai (e a marca dele)' (-not (Test-Path -LiteralPath $velhoV) -and -not (Test-Path -LiteralPath ($velhoV + '.validado')) -and (Test-Path -LiteralPath $novoN))

    Write-Host '--- 006: último backup validado e alerta (RPO 24 h, alerta 36 h)'
    $pa = Join-Path $tmp 'alerta'
    New-Item -ItemType Directory -Path $pa -Force | Out-Null
    $al = Get-AlertaBackup $pa
    Conferir 'sem nenhum backup validado: alerta' ($al.alerta -eq $true -and $null -eq $al.ultimo)
    $fa = Join-Path $pa 'crm_diario_20261010-080000_EST1.accdb'
    [IO.File]::WriteAllText(($fa + '.validado'), (ConvertTo-Json @{ validado_utc = (Get-Date).ToUniversalTime().AddHours(-10).ToString('yyyy-MM-ddTHH:mm:ssZ') }))
    $al = Get-AlertaBackup $pa
    Conferir 'validado há 10 h: sem alerta' ($al.alerta -eq $false -and $al.horas -ge 9 -and $al.horas -lt 11) ($al | ConvertTo-Json -Compress)
    [IO.File]::WriteAllText(($fa + '.validado'), (ConvertTo-Json @{ validado_utc = (Get-Date).ToUniversalTime().AddHours(-40).ToString('yyyy-MM-ddTHH:mm:ssZ') }))
    $al = Get-AlertaBackup $pa
    Conferir 'validado há 40 h: alerta' ($al.alerta -eq $true -and $al.mensagem -match '36')
    $res = ChamarMcp $banco7 'crm_status' ([ordered]@{})
    $st = $res.content[0].text | ConvertFrom-Json
    Conferir 'crm_status traz ultimo_backup_validado e alerta_backup=false depois do backup' ($st.alerta_backup -eq $false -and [string]$st.ultimo_backup_validado.arquivo -match '^crm_diario_') ($res.content[0].text)
    $res = ChamarMcp $ref8 'crm_status' ([ordered]@{})
    $st = $res.content[0].text | ConvertFrom-Json
    Conferir 'crm_status sem backup validado: alerta_backup=true' ($st.alerta_backup -eq $true)

    Write-Host '--- 006: geração da base (config.geracao_base)'
    $cnG = Open-BancoLeitura $banco7
    try { $g1 = [string](Invoke-Escalar $cnG "SELECT valor FROM config WHERE chave='geracao_base'" @() $null) } finally { $cnG.Close(); $cnG.Dispose() }
    Conferir 'banco novo nasce com geracao_base' ($g1 -match '^[0-9a-f]{32}$') $g1
    $ren = Join-Path $b7 'SISTEMA\DADOS\operacao\scripts\renovar-geracao.ps1'
    $r = Executar $ren @('-Banco', $banco7)
    $cnG = Open-BancoLeitura $banco7
    try { $g2 = [string](Invoke-Escalar $cnG "SELECT valor FROM config WHERE chave='geracao_base'" @() $null) } finally { $cnG.Close(); $cnG.Dispose() }
    Conferir 'renovar-geracao.ps1 troca a geração da base' ($r.codigo -eq 0 -and $g2 -match '^[0-9a-f]{32}$' -and $g2 -ne $g1) $r.saida
    $abertoG = Open-Banco $banco7
    try { $r = Executar $ren @('-Banco', $banco7) } finally { $abertoG.Close(); $abertoG.Dispose() }
    Conferir 'renovar-geracao.ps1 recusa banco em uso' ($r.codigo -ne 0 -and $r.saida -match 'conectado') $r.saida

    # FIM-DOS-CASOS
} finally { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue }

Write-Host ''
if ($script:falhas -gt 0) { Write-Host ('RESULTADO: ' + $script:falhas + ' falha(s)'); exit 1 }
Write-Host 'TUDO CERTO.'
exit 0
