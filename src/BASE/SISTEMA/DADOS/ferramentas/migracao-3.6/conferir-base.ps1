<#
  conferir-base.ps1 - compara o banco carregado com esperado.json.
  Divergencia de um registro interrompe a virada.
#>
param([string]$Banco = "", [string]$Dados = "")
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Data

$base = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($Banco)) { $Banco = Join-Path $base 'crm_zapromaq_teste.accdb' }
if ([string]::IsNullOrWhiteSpace($Dados)) { $Dados = Join-Path $base 'carga' }
$log = Join-Path $base ('log-conferencia-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.txt')
$saida = New-Object System.Collections.ArrayList
function L($t) { [void]$saida.Add([string]$t); Write-Host $t }
function Fim { $saida | Set-Content -Path $log -Encoding UTF8; Write-Host ""; Write-Host ("Log: " + $log) }

L "=== CONFERENCIA DA CARGA ==="
L ("Data..: " + (Get-Date -Format 'dd/MM/yyyy HH:mm:ss'))
L ("Banco.: " + $Banco)

$esperado = Get-Content -Path (Join-Path $Dados 'esperado.json') -Raw -Encoding UTF8 | ConvertFrom-Json

$provider = $null
foreach ($prov in @('Microsoft.ACE.OLEDB.16.0','Microsoft.ACE.OLEDB.12.0')) {
    try { $t = New-Object System.Data.OleDb.OleDbConnection("Provider=$prov;Data Source=$Banco;"); $t.Open(); $t.Close(); $t.Dispose(); $provider = $prov; break } catch { }
}
if (-not $provider) { L "ERRO: nenhum provedor ACE."; Fim; exit 1 }
$cn = New-Object System.Data.OleDb.OleDbConnection("Provider=$provider;Data Source=$Banco;")
$cn.Open()
function Escalar($sql) { $cmd = $cn.CreateCommand(); $cmd.CommandText = $sql; $v = $cmd.ExecuteScalar(); $cmd.Dispose(); if ($null -eq $v -or $v -is [System.DBNull]) { return 0 }; return $v }

$divergencias = 0
function Comparar($rotulo, $esp, $obt) {
    $ok = ([string]$esp -eq [string]$obt)
    if (-not $ok) { $script:divergencias = $script:divergencias + 1 }
    L (("  " + $rotulo).PadRight(38) + ("esperado: " + $esp).PadRight(26) + "obtido: " + $obt + $(if ($ok) { "   OK" } else { "   <<< DIVERGENTE" }))
}

L ""
L "--- contagens ---"
Comparar "clientes"              $esperado.clientes             (Escalar "SELECT COUNT(*) FROM clientes")
Comparar "clientes com codigo"   $esperado.clientes_com_codigo  (Escalar "SELECT COUNT(*) FROM clientes WHERE codigo_cliente IS NOT NULL")
Comparar "pre-clientes"          $esperado.clientes_pre         (Escalar "SELECT COUNT(*) FROM clientes WHERE codigo_cliente IS NULL")
Comparar "estagio = Cliente"     $esperado.clientes_com_codigo  (Escalar "SELECT COUNT(*) FROM clientes WHERE estagio='Cliente'")
Comparar "estagio = Pre-cliente" $esperado.clientes_pre         (Escalar "SELECT COUNT(*) FROM clientes WHERE estagio='Pré-cliente'")
Comparar "contatos"              $esperado.contatos             (Escalar "SELECT COUNT(*) FROM contatos")
Comparar "oportunidades"         $esperado.oportunidades        (Escalar "SELECT COUNT(*) FROM oportunidades")
Comparar "listas"                $esperado.listas               (Escalar "SELECT COUNT(*) FROM listas")

L ""
L "--- valores e codigos ---"
Comparar "soma de valor"         $esperado.soma_valor           ([math]::Round([double](Escalar "SELECT SUM(valor) FROM oportunidades"),2))
Comparar "maior codigo_cliente"  $esperado.max_codigo_cliente   (Escalar "SELECT MAX(codigo_cliente) FROM clientes")
Comparar "maior controle contato" $esperado.max_controle_contato (Escalar "SELECT MAX(controle) FROM contatos")
Comparar "maior controle atend."  $esperado.max_controle_oportunidade (Escalar "SELECT MAX(controle) FROM oportunidades")

L ""
L "--- distribuicao por etapa ---"
foreach ($prop in $esperado.por_etapa.PSObject.Properties) {
    $cmd = $cn.CreateCommand()
    $cmd.CommandText = "SELECT COUNT(*) FROM oportunidades WHERE etapa=?"
    $par = $cmd.CreateParameter(); $par.OleDbType = [System.Data.OleDb.OleDbType]::VarWChar; $par.Value = $prop.Name
    [void]$cmd.Parameters.Add($par)
    $obtido = $cmd.ExecuteScalar(); $cmd.Dispose()
    Comparar $prop.Name $prop.Value $obtido
}

L ""
L "--- integridade ---"
Comparar "contatos orfaos"       0 (Escalar "SELECT COUNT(*) FROM contatos WHERE id_cliente NOT IN (SELECT id FROM clientes)")
Comparar "oportunidades orfas"   0 (Escalar "SELECT COUNT(*) FROM oportunidades WHERE id_contato NOT IN (SELECT id FROM contatos)")
Comparar "codigos CT duplicados" 0 (Escalar "SELECT COUNT(*) FROM (SELECT codigo FROM contatos GROUP BY codigo HAVING COUNT(*)>1)")
Comparar "codigos AT duplicados" 0 (Escalar "SELECT COUNT(*) FROM (SELECT codigo FROM oportunidades GROUP BY codigo HAVING COUNT(*)>1)")
Comparar "oportunidade sem etapa" 0 (Escalar "SELECT COUNT(*) FROM oportunidades WHERE etapa IS NULL OR etapa=''")
Comparar "cliente sem empresa"   0 (Escalar "SELECT COUNT(*) FROM clientes WHERE empresa IS NULL OR empresa=''")

$cn.Close(); $cn.Dispose()
L ""
if ($divergencias -eq 0) {
    L "CONFERENCIA OK - nenhuma divergencia. Pode seguir."
} else {
    L ("CONFERENCIA REPROVADA - " + $divergencias + " divergencia(s). NAO seguir para a virada.")
    Fim; exit 1
}
Fim
