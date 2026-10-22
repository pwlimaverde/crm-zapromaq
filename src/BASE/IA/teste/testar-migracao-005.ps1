<#
  testar-migracao-005.ps1 - aplica a migração 005 (normalização da seção 5.1) num banco de teste temporário com
  textos "legados" (minúsculas, espaço duplo, dois e-mails, URL nas observações) e confere o resultado, o log
  com origem MIGRACAO e a idempotência. Nada da pasta é alterado; dados fictícios.
#>
param()
$ErrorActionPreference = 'Stop'
$raizIA = Split-Path -Parent $PSScriptRoot
$raizBase = Split-Path -Parent $raizIA
$pastaMig = Join-Path $raizBase 'SISTEMA\DADOS\banco\migracoes'
. (Join-Path $raizBase 'SISTEMA\DADOS\lib\Comum.ps1')

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
function Aplicar($cn) {
    $m = & (Join-Path $pastaMig '005-normalizacao-secao-5-1.ps1')
    $log = NovoLog
    $tx = $cn.BeginTransaction()
    try { & $m.Aplicar $cn $tx $log $pastaMig; $tx.Commit() } catch { $tx.Rollback(); throw }
    return $log
}

$banco = Join-Path ([System.IO.Path]::GetTempPath()) ('migracao-005-' + [guid]::NewGuid().ToString('N') + '.accdb')
& (Join-Path $PSScriptRoot 'criar-banco-teste.ps1') -Banco $banco -Recriar | Out-Null
$cn = Open-Banco $banco
try {
    # casos legados (gravados por fora da regra, como estavam antes do ciclo 009)
    [void](Invoke-Comando $cn "UPDATE clientes SET empresa=?, email=?, observacoes=? WHERE codigo_cliente=1" @(
        (New-Parametro 'texto' '  metalúrgica   exemplo  ltda '), (New-Parametro 'texto' 'Compras@Exemplo.com.br; vendas@exemplo.com.br'),
        (New-Parametro 'memo' ("ligar  segunda`r`nver http://site.com/Ab e C:\Pasta\Arq.txt"))) $null)
    [void](Invoke-Comando $cn "UPDATE contatos SET nome=? WHERE codigo='CT-0001-0001'" @((New-Parametro 'texto' 'joão  da silva')) $null)
    $maxLog = [int](Invoke-Escalar $cn 'SELECT MAX(id) FROM log_alteracoes' @() $null)
    $v1 = [int](Invoke-Escalar $cn 'SELECT versao FROM clientes WHERE codigo_cliente=1' @() $null)

    $log = Aplicar $cn
    $cli = @{}
    $c = New-ComandoBanco $cn 'SELECT empresa, email, observacoes, versao FROM clientes WHERE codigo_cliente=1' @() $null
    $rd = $c.ExecuteReader(); [void]$rd.Read(); $cli.empresa = [string]$rd.GetValue(0); $cli.email = [string]$rd.GetValue(1); $cli.obs = [string]$rd.GetValue(2); $cli.versao = [int]$rd.GetValue(3); $rd.Close(); $c.Dispose()
    $nome = [string](Invoke-Escalar $cn "SELECT nome FROM contatos WHERE codigo='CT-0001-0001'" @() $null)
    Conferir 'empresa em maiúsculas, acento preservado, espaço colapsado' ($cli.empresa -ceq ('METAL' + [char]0x00DA + 'RGICA EXEMPLO LTDA')) $cli.empresa
    Conferir 'e-mail: só o primeiro, minúsculo' ($cli.email -ceq 'compras@exemplo.com.br') $cli.email
    Conferir 'observações: maiúsculas, quebra de linha e URL/caminho preservados, e-mail adicional acrescentado' ($cli.obs -ceq ("LIGAR SEGUNDA`r`nVER http://site.com/Ab E C:\Pasta\Arq.txt`r`nE-MAIL ADICIONAL: vendas@exemplo.com.br")) $cli.obs
    Conferir 'contato: nome em maiúsculas' ($nome -ceq ('JO' + [char]0x00C3 + 'O DA SILVA')) $nome
    Conferir 'a versão do registro subiu uma vez' ($cli.versao -eq $v1 + 1) ([string]$cli.versao)
    $nLog = [int](Invoke-Escalar $cn "SELECT COUNT(*) FROM log_alteracoes WHERE id > ? AND origem='MIGRACAO' AND acao='NORMALIZACAO'" @((New-Parametro 'inteiro' $maxLog)) $null)
    Conferir 'cada campo alterado tem log com origem MIGRACAO (valor antigo e novo)' ($nLog -ge 4) ([string]$nLog)
    Conferir 'a saída traz a prévia por campo' (@($log.linhas | Where-Object { $_ -match 'clientes\.empresa: \d+ valor' }).Count -eq 1) ($log.linhas -join ' | ')
    $maxLog2 = [int](Invoke-Escalar $cn 'SELECT MAX(id) FROM log_alteracoes' @() $null)
    [void](Aplicar $cn)
    $nLog2 = [int](Invoke-Escalar $cn 'SELECT COUNT(*) FROM log_alteracoes WHERE id > ?' @((New-Parametro 'inteiro' $maxLog2)) $null)
    $quais = ''
    if ($nLog2 -gt 0) {
        $c = New-ComandoBanco $cn 'SELECT tabela, id_registro, campo, valor_antigo, valor_novo FROM log_alteracoes WHERE id > ?' @((New-Parametro 'inteiro' $maxLog2)) $null
        $rd = $c.ExecuteReader(); while ($rd.Read()) { $quais += ([string]$rd.GetValue(0) + '.' + [string]$rd.GetValue(2) + ' [' + [string]$rd.GetValue(3) + '] -> [' + [string]$rd.GetValue(4) + ']; ') }; $rd.Close(); $c.Dispose()
    }
    Conferir 'idempotente: rodar de novo não muda nada' ($nLog2 -eq 0) $quais
} finally { $cn.Close(); $cn.Dispose(); Remove-Item -LiteralPath $banco -Force -ErrorAction SilentlyContinue }

Write-Host ''
if ($script:falhas -gt 0) { Write-Host ('RESULTADO: ' + $script:falhas + ' falha(s)'); exit 1 }
Write-Host 'TUDO CERTO.'
exit 0
