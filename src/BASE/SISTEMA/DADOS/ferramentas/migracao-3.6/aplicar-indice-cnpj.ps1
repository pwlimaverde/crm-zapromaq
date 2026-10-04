<#
  aplicar-indice-cnpj.ps1 - indice UNICO de CNPJ em clientes.

  Conferido em 17/09/2026 na producao, por leitura mdbtools:
    853 registros - 392 com CNPJ - 461 com CNPJ NULO - 0 duplicado.
  Por isso o indice entra sem tratamento previo.

  O QUE ESTE INDICE NAO FAZ: o Access aceita varios NULOS num
  indice unico. 54% da base nao tem CNPJ, entao o indice protege
  392 registros e NAO protege os 461 sem CNPJ. A barreira dos
  pre-clientes sem CNPJ e a conferencia por nome normalizado no
  cadastro em lote, nao este indice.

  Roda de novo sem estragar nada: se o indice ja existe, avisa e sai.
#>
param([string]$Banco = "")
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Data

$base = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($Banco)) {
    $Banco = Join-Path $base '..\..\..\..\..\01 - CRM\01 - CONTROLE\BASE\crm_zapromaq.accdb'
}
$Banco = [System.IO.Path]::GetFullPath($Banco)
$log = Join-Path $base ('log-indice-cnpj-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.txt')
$saida = New-Object System.Collections.ArrayList
function L($t) { [void]$saida.Add([string]$t); Write-Host $t }
function Fim { $saida | Set-Content -Path $log -Encoding UTF8; Write-Host ""; Write-Host ("Log: " + $log) }

L "=== INDICE UNICO DE CNPJ ==="
L ("Banco: " + $Banco)
if (-not (Test-Path $Banco)) { L "ERRO: banco nao encontrado."; Fim; exit 1 }

$provider = $null
foreach ($prov in @('Microsoft.ACE.OLEDB.16.0','Microsoft.ACE.OLEDB.12.0')) {
    try { $t = New-Object System.Data.OleDb.OleDbConnection("Provider=$prov;Data Source=$Banco;"); $t.Open(); $t.Close(); $t.Dispose(); $provider = $prov; break } catch { }
}
if (-not $provider) { L "ERRO: nenhum provedor ACE."; Fim; exit 1 }

$cn = New-Object System.Data.OleDb.OleDbConnection("Provider=$provider;Data Source=$Banco;")
$cn.Open()
function Exec($sql) { $cmd = $cn.CreateCommand(); $cmd.CommandText = $sql; $null = $cmd.ExecuteNonQuery(); $cmd.Dispose() }
function Escalar($sql) { $cmd = $cn.CreateCommand(); $cmd.CommandText = $sql; $v = $cmd.ExecuteScalar(); $cmd.Dispose(); if ($null -eq $v -or $v -is [System.DBNull]) { return 0 }; return $v }

L ""
L "--- estado antes ---"
$tot = [int](Escalar "SELECT COUNT(*) FROM clientes")
$com = [int](Escalar "SELECT COUNT(*) FROM clientes WHERE cnpj IS NOT NULL AND Len(Trim(cnpj))>0")
$vaz = [int](Escalar "SELECT COUNT(*) FROM clientes WHERE cnpj IS NOT NULL AND Len(Trim(cnpj))=0")
$dup = [int](Escalar "SELECT COUNT(*) FROM (SELECT cnpj FROM clientes WHERE cnpj IS NOT NULL AND Len(Trim(cnpj))>0 GROUP BY cnpj HAVING COUNT(*)>1)")
L ("  registros ............: " + $tot)
L ("  com CNPJ .............: " + $com)
L ("  CNPJ vazio (nao nulo) : " + $vaz)
L ("  CNPJ duplicado .......: " + $dup)

# String vazia nao e NULL para o indice unico: a segunda vazia seria recusada.
# Se aparecer alguma, vira NULL antes do indice.
if ($vaz -gt 0) {
    L ""
    L ("--- convertendo " + $vaz + " CNPJ vazio em NULO ---")
    Exec "UPDATE clientes SET cnpj=NULL WHERE cnpj IS NOT NULL AND Len(Trim(cnpj))=0"
    L "  OK"
}

if ($dup -gt 0) {
    L ""
    L "ERRO: existe CNPJ duplicado. O Access recusa o indice unico."
    L "      Decida o que fazer com as duplicatas antes - este script nao apaga registro."
    $cmd = $cn.CreateCommand()
    $cmd.CommandText = "SELECT cnpj, COUNT(*) AS n FROM clientes WHERE cnpj IS NOT NULL GROUP BY cnpj HAVING COUNT(*)>1"
    $rd = $cmd.ExecuteReader()
    while ($rd.Read()) { L ("      " + $rd.GetValue(0) + "  ->  " + $rd.GetValue(1) + " registros") }
    $rd.Close(); $cmd.Dispose()
    $cn.Close(); Fim; exit 1
}

L ""
L "--- aplicando ---"
$jaExistia = $false
try {
    Exec "CREATE UNIQUE INDEX idx_cli_cnpj ON clientes (cnpj)"
    L "  OK    idx_cli_cnpj"
} catch {
    $msg = $_.Exception.Message.Split([char]13)[0]
    if ($msg -match 'exist' -or $msg -match 'existe') { $jaExistia = $true; L "  JA EXISTIA idx_cli_cnpj - nada a fazer." }
    else { L ("  FALHA idx_cli_cnpj -> " + $msg); $cn.Close(); Fim; exit 1 }
}

L ""
L "--- verificando que protege de verdade ---"
# O teste roda dentro de uma transacao e SEMPRE volta atras: nao deixa
# registro de teste na producao nem depende de um DELETE dar certo depois.
$amostra = Escalar "SELECT TOP 1 cnpj FROM clientes WHERE cnpj IS NOT NULL AND Len(Trim(cnpj))>0"
$protege = $false
$tx = $cn.BeginTransaction()
try {
    $cmd = $cn.CreateCommand()
    $cmd.Transaction = $tx
    $cmd.CommandText = "INSERT INTO clientes (empresa, estagio, cnpj, ativo, versao) VALUES ('ZZ TESTE INDICE','Pre-cliente','" + $amostra + "',False,1)"
    $null = $cmd.ExecuteNonQuery()
    $cmd.Dispose()
    L "  ATENCAO: CNPJ repetido foi ACEITO. O indice NAO esta protegendo."
} catch {
    $protege = $true
    L "  OK: insercao de CNPJ repetido recusada pelo banco."
}
try { $tx.Rollback() } catch { }
$sobrou = [int](Escalar "SELECT COUNT(*) FROM clientes WHERE empresa='ZZ TESTE INDICE'")
if ($sobrou -gt 0) {
    Exec "DELETE FROM clientes WHERE empresa='ZZ TESTE INDICE'"
    L ("  " + $sobrou + " registro(s) de teste removido(s) apos o rollback.")
} else {
    L "  Nada de teste ficou na base."
}

L ""
L "--- o que este indice NAO cobre ---"
$semCnpj = [int](Escalar "SELECT COUNT(*) FROM clientes WHERE cnpj IS NULL")
L ("  " + $semCnpj + " registro(s) sem CNPJ seguem sem barreira de unicidade no banco.")
L "  Duplicata neles se evita na conferencia por nome do cadastro em lote."

$cn.Close(); $cn.Dispose()
L ""
if (-not $protege -and -not $jaExistia) { L "TERMINOU COM FALHA - conferir acima."; Fim; exit 1 }
L "Indice de CNPJ aplicado."
Fim
