<#
  lancador.ps1 - ponto de entrada FIXO do conector de dados registrado no Claude Desktop e no Codex (DEV-05).

  Lê UMA vez, na abertura da sessão, o ponteiro IA\conector\ATUAL.txt (nome da versão) e carrega o servidor
  da pasta imutável IA\conector\versoes\<versão>\ (servidor.ps1, lib\ e a skill). Publicar e voltar é trocar o
  ponteiro (IA\mcp\publicar-conector.ps1): ninguém edita arquivo em uso, e a sessão aberta continua com os
  módulos que carregou. Sem ponteiro (instalação antiga, antes da primeira publicação), carrega IA\mcp\servidor.ps1.

  NADA pode sair no stdout além do protocolo: aviso vai para o stderr.
#>
param([string]$Banco = $env:CRM_BANCO)
$ErrorActionPreference = 'Stop'
$raizIA = Split-Path -Parent $PSScriptRoot
$servidor = Join-Path $PSScriptRoot 'servidor.ps1'
$ponteiro = Join-Path $raizIA 'conector\ATUAL.txt'
if (Test-Path -LiteralPath $ponteiro) {
    $versao = ([System.IO.File]::ReadAllText($ponteiro)).Trim()
    $candidato = Join-Path $raizIA ('conector\versoes\' + $versao + '\servidor.ps1')
    if ($versao -match '^\d+\.\d+\.\d+$' -and (Test-Path -LiteralPath $candidato)) { $servidor = $candidato }
    else { [Console]::Error.WriteLine('lancador: ponteiro ' + $ponteiro + ' aponta para "' + $versao + '", que não existe; usando IA\mcp\servidor.ps1') }
}
$env:CRM_RAIZ_IA = $raizIA
if ([string]::IsNullOrWhiteSpace($Banco)) { & $servidor } else { & $servidor -Banco $Banco }
