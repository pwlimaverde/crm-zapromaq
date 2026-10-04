<#
  servidor-dev.ps1 - MCP de DESENVOLVIMENTO do CRM Zapromaq (D12).

  Separado do MCP de dados (IA\mcp\servidor.ps1): não toca no banco. Serve
  para o Claude Desktop da empresa fazer ajustes pequenos no front - ler e
  alterar SISTEMA\DADOS\fonte\, ver a prévia da tela como imagem, montar um
  .xlsm de teste e, só a pedido, publicar.

  Mesmo núcleo de protocolo do servidor de dados (IA\mcp\lib\Protocolo.ps1):
  stdio, dual-era, nada além de JSON-RPC no stdout.

  -Dados: pasta SISTEMA\DADOS a usar (padrão: a desta BASE). O teste
  automático aponta para uma cópia.
#>
param([string]$Dados = $env:CRM_DEV_DADOS)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$script:VersaoServidor = '1.0.0'
$script:InfoServidor = [ordered]@{ name = 'crm-zapromaq-dev'; title = 'CRM Zapromaq - desenvolvimento do front'; version = $script:VersaoServidor }

$raizIA   = Split-Path -Parent $PSScriptRoot
$raizBase = Split-Path -Parent $raizIA
if ([string]::IsNullOrWhiteSpace($Dados)) { $Dados = Join-Path $raizBase 'SISTEMA\DADOS' }
$script:PastaDados = [System.IO.Path]::GetFullPath($Dados)
$script:PastaFonte = Join-Path $script:PastaDados 'fonte'

. (Join-Path $PSScriptRoot 'lib\FerramentasDev.ps1')
. (Join-Path $raizIA 'mcp\lib\Protocolo.ps1')

$script:Instrucoes = @'
Servidor de DESENVOLVIMENTO do front do CRM Zapromaq (Excel 2019 + VBA, montado a partir de fontes em texto).
- Só se altera SISTEMA\DADOS\fonte: modulos\ e classes\ (VBA), formularios\ (código dos UserForms), layout\ (tema.json, icones.json e formularios\<tela>.json com a posição de cada controle, em pontos).
- Fluxo: ler_fonte -> mostrar a alteração ao usuário -> gravar_fonte (com a versao lida) -> gerar_previa da tela (devolve a imagem) -> montar_teste -> publicar só quando o usuário pedir.
- Nomes de controles em layout\formularios\*.json são contrato com o código do formulário: renomear exige alterar os dois.
- VBA: Windows-1252; acento fora dessa página vira ChrW$(&H....). Excel 2019: nada de funções do Microsoft 365.
- Nunca editar arquivos gerados (assets\gerado, modulos\modTema.bas): mudam por layout\ + gerar_previa.
- Toda gravação fica em listar_alteracoes; lembre o usuário de que o ajuste precisa voltar ao repositório de desenvolvimento.
'@

Initialize-LogMcp $raizIA 'mcp-dev'
Write-Diag ('fonte: ' + $script:PastaFonte)
Start-ServidorMcp
