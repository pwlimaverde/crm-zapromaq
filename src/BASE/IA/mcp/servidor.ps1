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

# 1.6.0 (ciclo 007): crm_status com versao_conector, capacidades, diagnostico, motivos_backup e pasta_conector;
# recusa de gravação por manutenção (config.manutencao) e mensagem controlada durante o backup.
# 1.6.1 (ciclo 008): moeda/inteiro sem perda (C17-C21), repetição de bloqueio antes do commit, resultado indeterminado.
# 1.7.0 (ciclo 009): seção 5.1/5.2 nos três canais; atualizar_contato sem codigo_cliente_destino; reativar só a empresa.
# 1.8.0 (ciclo 010): esquema 1.2; ctx_versao exigida no contexto; codigo_atendimento_anterior no atualizar_atendimento
# (sem ciclo); id_operacao nas criações; log com estação, usuário técnico, versão e autor declarado.
# 1.9.0 (ciclo 012): situação "Retomada a avaliar"; retomadas_a_avaliar e cursor no planejamento; exportar_base
# sem destino (raiz execucao\extratos, manifesto); descartar_extratos; roteiro na raiz de extratos.
# 1.10.0 (ciclo 016): listar_candidatos ganha sem_atendimento (aditivo).
# 1.11.0 (ciclo 017): duplicar_atendimento (ferramenta nova, aditiva).
# 1.12.0 (ciclo 018): próxima ação só hoje ou futura (C34) e demais pendências finais (nota no IA\LEIA-ME.md).
$script:VersaoServidor = '1.12.0'
# Servidor "dual-era" (referencias/12-mcp-especificacao/2026-07-28_basic_versioning.md):
# protocolo moderno 2026-07-28 (sem estado) e legado (initialize). O núcleo do
# protocolo fica em lib\Protocolo.ps1, compartilhado com o MCP de desenvolvimento.
$script:InfoServidor = [ordered]@{ name = 'crm-zapromaq'; title = 'CRM Zapromaq'; version = $script:VersaoServidor }

# Carregado pelo lançador de uma pasta de versão (IA\conector\versoes\<v>, DEV-05), a pasta IA vem do
# lançador; carregado direto (IA\mcp), é a pasta acima desta.
$raizIA   = if ($env:CRM_RAIZ_IA -and (Test-Path -LiteralPath $env:CRM_RAIZ_IA)) { $env:CRM_RAIZ_IA } else { Split-Path -Parent $PSScriptRoot }
$raizBase = Split-Path -Parent $raizIA
$script:PastaConector = $PSScriptRoot
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
- Leia antes de gravar: obter_atendimento / obter_cliente / obter_contato devolvem a "versao" exigida nas ferramentas de atualização.
- Se a gravação voltar "Conflito", alguém alterou o registro depois da leitura: leia de novo, confira com o usuário e só então repita.
- Campos de lista (etapa, responsável, família, origem, prioridade, motivo, segmento, UF) só aceitam valores de listar_opcoes.
- Datas sempre AAAA-MM-DD. Valores em reais como número (ex.: 1250000.50).
- Só pré-cliente sem vínculo pode ser excluído (alterar_situacao_cliente); cliente e contato se desativam; observações só recebem linhas novas.
- Cadastrar contato ou abrir atendimento em pré-cliente promove a empresa a cliente: só com confirmação do usuário (promover_pre_cliente=true).
- Toda gravação fica no log de auditoria com origem IA; confirme com o usuário antes de alterar dados de negócio.
'@


# quem gravou: 'IA' ou 'IA (<autor>)' no log de auditoria
$script:AntesDaFerramenta = {
    param($a, $nomeFerramenta)
    # banco em esquema não suportado (ex.: 1.0, aguardando a montagem): só o crm_status responde
    if ($nomeFerramenta -ne 'crm_status') {
        $motivoEsquema = Get-MotivoEsquemaNaoSuportado
        if ($motivoEsquema) { throw $motivoEsquema }
    }
    $script:UsuarioLog = 'IA'
    $script:EmailsAdicionais = ''
    $autor = Get-Arg $a 'autor'
    $script:AutorDeclarado = $null
    if ($autor) { $script:UsuarioLog = ('IA (' + ([string]$autor).Trim() + ')'); $script:AutorDeclarado = ([string]$autor).Trim() }
    if ($script:AutorDeclarado -and $script:AutorDeclarado.Length -gt 60) { $script:AutorDeclarado = $script:AutorDeclarado.Substring(0, 60) }
    if ($script:UsuarioLog.Length -gt 50) { $script:UsuarioLog = $script:UsuarioLog.Substring(0, 50) }
}

Initialize-LogMcp $raizIA 'mcp'
Write-Diag ('banco: ' + $script:CaminhoBanco)
Start-ServidorMcp
