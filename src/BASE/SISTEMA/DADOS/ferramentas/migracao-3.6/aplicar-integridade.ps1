<#
  aplicar-integridade.ps1 - chaves estrangeiras do CRM Zapromaq.
  Rodar DEPOIS da carga: aplicar antes obrigaria ordem de insercao no mesmo passo.
#>
param([string]$Banco = "")
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Data

$base = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($Banco)) { $Banco = Join-Path $base 'crm_zapromaq_teste.accdb' }
$log = Join-Path $base ('log-integridade-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.txt')
$saida = New-Object System.Collections.ArrayList
function L($t) { [void]$saida.Add([string]$t); Write-Host $t }
function Fim { $saida | Set-Content -Path $log -Encoding UTF8; Write-Host ""; Write-Host ("Log: " + $log) }

L "=== INTEGRIDADE REFERENCIAL ==="
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
L "--- orfaos antes de aplicar ---"
$orfCto = [int](Escalar "SELECT COUNT(*) FROM contatos WHERE id_cliente NOT IN (SELECT id FROM clientes)")
$orfOpo = [int](Escalar "SELECT COUNT(*) FROM oportunidades WHERE id_contato NOT IN (SELECT id FROM contatos)")
L ("  contatos orfaos ......: " + $orfCto)
L ("  oportunidades orfas ..: " + $orfOpo)
if ($orfCto -gt 0 -or $orfOpo -gt 0) {
    L "ERRO: existe registro orfao. A constraint nao pode ser aplicada. Corrija a carga."
    $cn.Close(); Fim; exit 1
}

L ""
L "--- aplicando ---"
$constraints = @(
 @{ n='fk_cto_cli'; s="ALTER TABLE contatos ADD CONSTRAINT fk_cto_cli FOREIGN KEY (id_cliente) REFERENCES clientes (id)" },
 @{ n='fk_opo_cto'; s="ALTER TABLE oportunidades ADD CONSTRAINT fk_opo_cto FOREIGN KEY (id_contato) REFERENCES contatos (id)" },
 @{ n='fk_opo_cli'; s="ALTER TABLE oportunidades ADD CONSTRAINT fk_opo_cli FOREIGN KEY (id_cliente) REFERENCES clientes (id)" }
)
$falhou = $false
foreach ($c in $constraints) {
    try { Exec $c.s; L ("  OK    " + $c.n) }
    catch { $falhou = $true; L ("  FALHA " + $c.n + " -> " + $_.Exception.Message.Split([char]13)[0]) }
}

L ""
L "--- verificando que protege de verdade ---"
$idAlvo = [int](Escalar "SELECT TOP 1 id_cliente FROM contatos")
try {
    Exec ("DELETE FROM clientes WHERE id=" + $idAlvo)
    L "  ATENCAO: cliente com contato foi EXCLUIDO. A integridade NAO esta protegendo."
    $falhou = $true
} catch { L "  OK: exclusao de cliente com contato recusada pelo banco." }

$cn.Close(); $cn.Dispose()
L ""
if ($falhou) { L "TERMINOU COM FALHA - conferir acima."; Fim; exit 1 }
L "Integridade aplicada. Proximo passo: conferir-base.ps1"
Fim
