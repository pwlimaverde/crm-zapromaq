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

    # FIM-DOS-CASOS
} finally { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue }

Write-Host ''
if ($script:falhas -gt 0) { Write-Host ('RESULTADO: ' + $script:falhas + ' falha(s)'); exit 1 }
Write-Host 'TUDO CERTO.'
exit 0
