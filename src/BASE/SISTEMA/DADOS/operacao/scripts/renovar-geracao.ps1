<#
  renovar-geracao.ps1 - dá ao banco uma NOVA geração (config.geracao_base).

  A geração identifica "esta base". Todo evento que substitui o arquivo do banco por outro (publicação do
  MONTAR-FRONTEND, restauração de um backup sobre a produção) tem de renovar a geração; assim as estações
  percebem que a base foi trocada por baixo delas e recarregam tudo em vez de seguir um feed que não vale mais.

  Depois de RESTAURAR um backup sobre o banco da produção, rode este script (banco fechado).
  Recusa se houver alguém conectado (.laccdb) ou se o banco for de PRODUCAO sem -Confirmo.

  Uso: renovar-geracao.ps1 [-Banco <accdb>] [-Confirmo]
#>
param([string]$Banco = '', [switch]$Confirmo)
$ErrorActionPreference = 'Stop'
$dadosLib = Join-Path (Split-Path -Parent $PSScriptRoot) '..\lib'
. (Join-Path $dadosLib 'Comum.ps1')
. (Join-Path $dadosLib 'AtualizacaoBanco.ps1')

$raiz = Get-RaizBase $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($Banco)) { $Banco = Get-CaminhoBanco $raiz }
if (-not (Test-Path -LiteralPath $Banco)) { Write-Host ('Banco não encontrado: ' + $Banco); exit 1 }

$livre = Test-BancoLivre $Banco
if (-not $livre.livre) {
    $quem = if ($livre.estacoes.Count) { ' Estações: ' + ($livre.estacoes -join ', ') + '.' } else { '' }
    Write-Host ('RECUSADO: há usuário conectado ao banco (.laccdb).' + $quem + ' Peça para fecharem o CRM. Nada foi alterado.'); exit 2
}
$amb = Get-AmbienteDoBanco $Banco
if ($amb -ne 'TESTE' -and -not $Confirmo) {
    Write-Host ('RECUSADO: o banco está com ambiente ' + $amb + '. Confirme com -Confirmo depois de restaurar o backup e com todos fora do CRM. Nada foi alterado.'); exit 3
}
$cn = Open-Banco $Banco
try { $g = Set-GeracaoBase $cn } finally { $cn.Close(); $cn.Dispose() }
Write-Host ('Nova geração da base: ' + $g)
exit 0
