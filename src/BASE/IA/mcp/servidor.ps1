<#
  servidor.ps1 - servidor MCP (Model Context Protocol) do CRM Zapromaq.

  Transporte stdio: o Claude Desktop inicia este script e troca mensagens
  JSON-RPC, uma por linha, em UTF-8. Regra de ouro do stdio: NADA além de
  mensagem MCP pode sair no stdout - diagnóstico vai para o stderr e para
  IA\logs. Por isso não se usa Write-Host/Write-Output aqui.

  Roda no Windows PowerShell 5.1 64 bits, que já vem no Windows. O banco
  é acessado pelo motor ACE, o mesmo do front: é o que garante o controle
  de bloqueio do .accdb com os vendedores usando o sistema ao mesmo tempo.

  Banco: por padrão, BASE\crm_zapromaq.accdb (dois níveis acima desta
  pasta). Pode ser trocado por -Banco ou pela variável CRM_BANCO.
#>
param([string]$Banco = $env:CRM_BANCO)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Data

$script:VersaoServidor = '1.0.0'
# Servidor "dual-era" (referencias/12-mcp-especificacao/2026-07-28_basic_versioning.md):
# protocolo moderno 2026-07-28 (sem estado) e legado (initialize). O núcleo do
# protocolo fica em lib\Protocolo.ps1, compartilhado com o MCP de desenvolvimento.
$script:InfoServidor = [ordered]@{ name = 'crm-zapromaq'; title = 'CRM Zapromaq'; version = $script:VersaoServidor }

$raizIA   = Split-Path -Parent $PSScriptRoot
$raizBase = Split-Path -Parent $raizIA
if ([string]::IsNullOrWhiteSpace($Banco)) { $Banco = Join-Path $raizBase 'crm_zapromaq.accdb' }
$script:CaminhoBanco = $Banco
$script:UsuarioLog = 'IA'

. (Join-Path $PSScriptRoot 'lib\Banco.ps1')
. (Join-Path $PSScriptRoot 'lib\Regras.ps1')
. (Join-Path $PSScriptRoot 'lib\Ferramentas.ps1')
. (Join-Path $PSScriptRoot 'lib\Protocolo.ps1')

# ------------------------------------------------------------------ instruções
$script:Instrucoes = @'
Servidor do CRM Comercial Zapromaq (banco Access compartilhado em rede).
- Leia antes de gravar: obter_atendimento / obter_cliente devolvem a "versao" exigida nas ferramentas de atualização.
- Se a gravação voltar "Conflito", alguém alterou o registro depois da leitura: leia de novo, confira com o usuário e só então repita.
- Campos de lista (etapa, responsável, família, origem, prioridade, motivo, segmento, UF) só aceitam valores de listar_opcoes.
- Datas sempre AAAA-MM-DD. Valores em reais como número (ex.: 1250000.50).
- Nada é apagado por estas ferramentas; observações só recebem linhas novas.
- Toda gravação fica no log de auditoria com origem IA; confirme com o usuário antes de alterar dados de negócio.
'@


# quem gravou: 'IA' ou 'IA (<autor>)' no log de auditoria
$script:AntesDaFerramenta = {
    param($a)
    $script:UsuarioLog = 'IA'
    $autor = Get-Arg $a 'autor'
    if ($autor) { $script:UsuarioLog = ('IA (' + ([string]$autor).Trim() + ')') }
    if ($script:UsuarioLog.Length -gt 50) { $script:UsuarioLog = $script:UsuarioLog.Substring(0, 50) }
}

Initialize-LogMcp $raizIA 'mcp'
Write-Diag ('banco: ' + $script:CaminhoBanco)
Start-ServidorMcp
