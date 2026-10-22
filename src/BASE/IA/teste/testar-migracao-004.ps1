<#
  testar-migracao-004.ps1 - migração 004 (esquema 1.1 -> 1.2, ciclo 010) em bancos fictícios temporários:
  inventário que PARA com controle repetido, anterior que aponta para si, de outra empresa ou em ciclo; aplicação
  limpa com índices únicos, ctx_versao, ctx_familia 255, colunas de autor no log e tabela operacoes; idempotência;
  conferências do passo 8. Nada da pasta é alterado.
#>
param()
$ErrorActionPreference = 'Stop'
$raizIA = Split-Path -Parent $PSScriptRoot
$raizBase = Split-Path -Parent $raizIA
$pastaMig = Join-Path $raizBase 'SISTEMA\DADOS\banco\migracoes'
. (Join-Path $raizBase 'SISTEMA\DADOS\lib\Comum.ps1')
. (Join-Path $raizBase 'SISTEMA\DADOS\lib\AtualizacaoBanco.ps1')

$script:falhas = 0
function Conferir([string]$nome, [bool]$ok, [string]$detalhe = '') {
    if ($ok) { Write-Host ('  OK    ' + $nome) -ForegroundColor Green }
    else { $script:falhas++; Write-Host ('  FALHA ' + $nome + '  ' + $detalhe) -ForegroundColor Red }
}
function NovoLog {
    $l = [pscustomobject]@{ linhas = New-Object System.Collections.ArrayList }
    $l | Add-Member -MemberType ScriptMethod -Name L -Value { param($t) [void]$this.linhas.Add([string]$t) }
    return $l
}
$m3 = & (Join-Path $pastaMig '003-municipios-visitas-papel-listas.ps1')
$m4 = & (Join-Path $pastaMig '004-esquema-1-2.ps1')
# banco 1.1 fictício: o de casos da migração (1.0) com a 003 aplicada
$modelo = Join-Path ([System.IO.Path]::GetTempPath()) ('migracao-004-modelo-' + [guid]::NewGuid().ToString('N') + '.accdb')
& (Join-Path $PSScriptRoot 'criar-banco-teste.ps1') -Banco $modelo -Recriar -CasosDaMigracao | Out-Null
$cn = Open-Banco $modelo
try { & $m3.Aplicar $cn $null (NovoLog) $pastaMig; Set-ConfigBanco $cn 'versao_esquema' '1.1' $null } finally { $cn.Close(); $cn.Dispose() }
$temporarios = New-Object System.Collections.ArrayList
function Copia {
    $c = Join-Path ([System.IO.Path]::GetTempPath()) ('migracao-004-' + [guid]::NewGuid().ToString('N') + '.accdb')
    Copy-Item -LiteralPath $modelo -Destination $c; [void]$temporarios.Add($c); return $c
}
function Tentar([string]$banco, [scriptblock]$estrago) {
    $cn = Open-Banco $banco
    try {
        & $estrago $cn
        $log = NovoLog; $erro = ''
        try { & $m4.Aplicar $cn $null $log $pastaMig } catch { $erro = $_.Exception.Message }
        $indice = $false; foreach ($r in $cn.GetSchema('Indexes').Rows) { if ([string]$r['INDEX_NAME'] -eq 'ix_opo_controle') { $indice = $true } }
        return [pscustomobject]@{ erro = $erro; log = ($log.linhas -join ' | '); criouEstrutura = $indice }
    } finally { $cn.Close(); $cn.Dispose() }
}
function Ex($cn, [string]$sql) { [void](Invoke-Comando $cn $sql @() $null) }

try {
    Write-Host '--- inventário: com achado, a 004 para antes de mexer na estrutura'
    $r = Tentar (Copia) { param($cn) $max = [int](Invoke-Escalar $cn 'SELECT MAX(controle) FROM oportunidades' @() $null); Ex $cn ('UPDATE oportunidades SET controle=' + $max + ' WHERE id=(SELECT MIN(id) FROM oportunidades)') }
    Conferir 'controle repetido: para, lista o achado e não cria estrutura' ($r.erro -match 'inventário' -and $r.log -match 'controle \d+ repetido' -and -not $r.criouEstrutura) ($r.erro + ' / ' + $r.log)
    $r = Tentar (Copia) { param($cn) Ex $cn 'UPDATE oportunidades SET id_atendimento_anterior=id WHERE id=(SELECT MIN(id) FROM oportunidades)' }
    Conferir 'anterior que aponta para si: para' ($r.erro -match 'inventário' -and $r.log -match 'ele mesmo') $r.log
    $r = Tentar (Copia) { param($cn) Ex $cn 'UPDATE oportunidades SET id_atendimento_anterior=999999 WHERE id=(SELECT MIN(id) FROM oportunidades)' }
    Conferir 'anterior que não existe: para' ($r.erro -match 'inventário' -and $r.log -match 'não existe') $r.log
    $r = Tentar (Copia) {
        param($cn)
        $c = New-ComandoBanco $cn 'SELECT id, id_cliente FROM oportunidades ORDER BY id_cliente, id' @() $null
        $rd = $c.ExecuteReader(); $l = @(); while ($rd.Read()) { $l += , @([int]$rd.GetValue(0), [int]$rd.GetValue(1)) }; $rd.Close(); $c.Dispose()
        $par = $null; for ($i = 1; $i -lt $l.Count; $i++) { if ($l[$i][1] -eq $l[$i - 1][1]) { $par = @($l[$i - 1][0], $l[$i][0]); break } }
        Ex $cn ('UPDATE oportunidades SET id_atendimento_anterior=' + $par[1] + ' WHERE id=' + $par[0])
        Ex $cn ('UPDATE oportunidades SET id_atendimento_anterior=' + $par[0] + ' WHERE id=' + $par[1])
    }
    Conferir 'anteriores em ciclo (A -> B -> A): para' ($r.erro -match 'inventário' -and $r.log -match 'ciclo') $r.log

    Write-Host '--- aplicação limpa'
    $limpo = Copia
    $cn = Open-Banco $limpo
    try {
        $log = NovoLog
        & $m4.Aplicar $cn $null $log $pastaMig
        Set-ConfigBanco $cn 'versao_esquema' $script:EsquemaAtual $null
        $ix = @{}; foreach ($x in $cn.GetSchema('Indexes').Rows) { $ix[[string]$x['INDEX_NAME']] = $true }
        Conferir 'índices únicos de controle e de operação' ($ix['ix_cto_controle'] -and $ix['ix_opo_controle'] -and $ix['ix_ope_id']) ''
        Conferir 'ctx_versao começa em 0' ([int](Invoke-Escalar $cn 'SELECT COUNT(*) FROM clientes WHERE ctx_versao IS NULL OR ctx_versao<>0' @() $null) -eq 0)
        $longo = 'F' * 200
        [void](Invoke-Comando $cn 'UPDATE clientes SET ctx_familia=? WHERE id=(SELECT MIN(id) FROM clientes)' @((New-Parametro 'texto' $longo)) $null)
        Conferir 'ctx_familia aceita 200 caracteres (255)' ([string](Invoke-Escalar $cn 'SELECT ctx_familia FROM clientes WHERE id=(SELECT MIN(id) FROM clientes)' @() $null) -eq $longo)
        [void](Invoke-Comando $cn 'INSERT INTO log_alteracoes (tabela, id_registro, campo, acao, usuario, quando, origem, usuario_tecnico, estacao, versao_cliente, autor_declarado) VALUES (?,?,?,?,?,?,?,?,?,?,?)' @(
            (New-Parametro 'texto' 'clientes'), (New-Parametro 'inteiro' 1), (New-Parametro 'texto' 'x'), (New-Parametro 'texto' 'TESTE'), (New-Parametro 'texto' 'IA'),
            (New-Parametro 'data' (Get-Date)), (New-Parametro 'texto' 'IA'), (New-Parametro 'texto' 'usuario'), (New-Parametro 'texto' 'EST'), (New-Parametro 'texto' '1.8.0'), (New-Parametro 'texto' 'Ana')) $null)
        Conferir 'log aceita usuário técnico, estação, versão do cliente e autor declarado' ($true)
        $erroDup = ''
        try { Ex $cn 'UPDATE contatos SET controle=(SELECT MAX(controle) FROM contatos) WHERE id=(SELECT MIN(id) FROM contatos)' } catch { $erroDup = $_.Exception.Message }
        Conferir 'controle repetido passa a ser recusado pelo banco' ($erroDup -ne '') $erroDup
        $erro2 = ''
        try { & $m4.Aplicar $cn $null (NovoLog) $pastaMig } catch { $erro2 = $_.Exception.Message }
        Conferir 'idempotente: aplicar de novo não dá erro' ($erro2 -eq '') $erro2
        $problemas = Get-ProblemasDoBanco $cn 5571
        Conferir 'cópia migrada passa nas conferências do passo 8' ($problemas.Count -eq 0) ($problemas -join '; ')
    } finally { $cn.Close(); $cn.Dispose() }
} finally {
    foreach ($t in $temporarios) { Remove-Item -LiteralPath $t -Force -ErrorAction SilentlyContinue }
    Remove-Item -LiteralPath $modelo -Force -ErrorAction SilentlyContinue
}
Write-Host ''
if ($script:falhas -eq 0) { Write-Host 'TUDO CERTO.' -ForegroundColor Green; exit 0 }
Write-Host ($script:falhas.ToString() + ' falha(s).') -ForegroundColor Red; exit 1
