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

    # FIM-DOS-CASOS
} finally { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue }

Write-Host ''
if ($script:falhas -gt 0) { Write-Host ('RESULTADO: ' + $script:falhas + ' falha(s)'); exit 1 }
Write-Host 'TUDO CERTO.'
exit 0
