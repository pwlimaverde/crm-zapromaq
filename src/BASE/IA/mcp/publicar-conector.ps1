<#
  publicar-conector.ps1 - publica o conector de dados numa pasta de versão imutável e troca o ponteiro (DEV-05).

    IA\conector\versoes\<versão>\   servidor.ps1, lib\*.ps1, skill\SKILL.md e manifesto.json (SHA-256 de cada arquivo)
    IA\conector\ATUAL.txt           a versão que o lançador (IA\mcp\lancador.ps1) carrega na próxima sessão
    IA\conector\historico.log       quem publicou/voltou, quando, de qual para qual versão e por quê

  Publicar (sem parâmetros): a versão é a de IA\mcp\servidor.ps1 ($script:VersaoServidor). A pasta nasce como
  temporária, é conferida arquivo a arquivo e só então renomeada; uma pasta de versão que já existe com conteúdo
  DIFERENTE é recusada (versão publicada não muda - suba a versão). Antes de trocar o ponteiro roda o
  testar-mcp contra a versão nova, num banco de homologação temporário (-SemTeste pula; só para testes).
  Promoção interrompida não deixa versão mista: o ponteiro só muda no fim, por troca atômica.

  Voltar (-Voltar <versão>): reaponta para uma versão já publicada, conferindo o manifesto dela e se ela
  suporta o esquema do banco (volta incompatível é recusada). Nada é copiado nem apagado.

  Uso: publicar-conector.ps1 [-Motivo "..."] [-SemTeste] | -Voltar <versão> -Motivo "..."
  Saída: 0 ok, 1 recusado/falhou. Não precisa de administrador.
#>
param([string]$Voltar = '', [string]$Motivo = '', [switch]$SemTeste, [string]$Banco = '')
$ErrorActionPreference = 'Stop'
$raizIA = Split-Path -Parent $PSScriptRoot
$raizBase = Split-Path -Parent $raizIA
$pastaCon = Join-Path $raizIA 'conector'
$versoes = Join-Path $pastaCon 'versoes'
$ponteiro = Join-Path $pastaCon 'ATUAL.txt'
$historico = Join-Path $pastaCon 'historico.log'
$ps = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
if (-not $Banco) { $Banco = Join-Path $raizBase 'crm_zapromaq.accdb' }

function Saida([string]$t) { Write-Host $t }
function Recusar([string]$t) { Write-Host ('RECUSADO: ' + $t) -ForegroundColor Red; exit 1 }
function Hash([string]$a) {
    $sha = [System.Security.Cryptography.SHA256]::Create(); $fs = [System.IO.File]::OpenRead($a)
    try { return ([BitConverter]::ToString($sha.ComputeHash($fs))).Replace('-', '').ToLowerInvariant() } finally { $fs.Dispose(); $sha.Dispose() }
}
function VersaoDe([string]$servidorPs1) {
    if ([System.IO.File]::ReadAllText($servidorPs1) -match "(?m)^\`$script:VersaoServidor = '([^']+)'") { return $Matches[1] }
    return ''
}
# Esquemas que uma pasta de conector suporta (lib\Banco.ps1, $script:EsquemasSuportados).
function EsquemasDe([string]$pasta) {
    $t = [System.IO.File]::ReadAllText((Join-Path $pasta 'lib\Banco.ps1'))
    if ($t -match "(?m)^\`$script:EsquemasSuportados = @\(([^)]*)\)") { return @([regex]::Matches($Matches[1], "'([^']+)'") | ForEach-Object { $_.Groups[1].Value }) }
    return @()
}
# Arquivos de uma versão: caminho relativo na pasta da versão -> arquivo de origem.
function Get-ArquivosDaVersao {
    $m = [ordered]@{ 'servidor.ps1' = (Join-Path $PSScriptRoot 'servidor.ps1') }
    foreach ($f in @(Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot 'lib') -Filter '*.ps1' -File | Sort-Object Name)) { $m['lib\' + $f.Name] = $f.FullName }
    $skill = Join-Path $raizIA 'instrucoes\crm-zapromaq\SKILL.md'
    if (Test-Path -LiteralPath $skill) { $m['skill\SKILL.md'] = $skill }
    return $m
}
# Problemas de uma pasta de versão contra o manifesto dela ('' = íntegra).
function Test-PastaVersao([string]$pasta) {
    $man = Join-Path $pasta 'manifesto.json'
    if (-not (Test-Path -LiteralPath $man)) { return 'sem manifesto.json' }
    $m = Get-Content -LiteralPath $man -Raw -Encoding UTF8 | ConvertFrom-Json
    foreach ($a in @($m.arquivos)) {
        $arq = Join-Path $pasta ([string]$a.caminho)
        if (-not (Test-Path -LiteralPath $arq)) { return ('falta ' + $a.caminho) }
        if ((Hash $arq) -ne [string]$a.sha256) { return ($a.caminho + ' alterado depois de publicado') }
    }
    return ''
}
function Set-Ponteiro([string]$versao, [string]$acao) {
    $anterior = if (Test-Path -LiteralPath $ponteiro) { ([System.IO.File]::ReadAllText($ponteiro)).Trim() } else { '' }
    $tmp = $ponteiro + '.tmp'
    [System.IO.File]::WriteAllText($tmp, $versao, (New-Object System.Text.UTF8Encoding($false)))
    if (Test-Path -LiteralPath $ponteiro) { [System.IO.File]::Replace($tmp, $ponteiro, [System.Management.Automation.Language.NullString]::Value) } else { [System.IO.File]::Move($tmp, $ponteiro) }
    $linha = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + '  ' + $acao + '  ' + $(if ($anterior) { $anterior } else { '(nenhuma)' }) + ' -> ' + $versao + '  por ' + $env:COMPUTERNAME + '\' + $env:USERNAME + $(if ($Motivo) { '  motivo: ' + $Motivo } else { '' })
    Add-Content -LiteralPath $historico -Value $linha -Encoding UTF8
    Saida ('ponteiro: ' + $(if ($anterior) { $anterior } else { '(nenhum)' }) + ' -> ' + $versao)
}
function Get-EsquemaDoBanco([string]$banco) {
    $prov = @('Microsoft.ACE.OLEDB.16.0', 'Microsoft.ACE.OLEDB.12.0') | Where-Object { @((New-Object System.Data.OleDb.OleDbEnumerator).GetElements() | ForEach-Object { $_.SOURCES_NAME }) -contains $_ } | Select-Object -First 1
    $cn = New-Object System.Data.OleDb.OleDbConnection('Provider=' + $prov + ';Data Source=' + $banco + ';Mode=Read;OLE DB Services=-4;')
    $cn.Open()
    try { $c = $cn.CreateCommand(); $c.CommandText = "SELECT valor FROM config WHERE chave='versao_esquema'"; return [string]$c.ExecuteScalar() } finally { $cn.Dispose() }
}

New-Item -ItemType Directory -Path $versoes -Force | Out-Null
Add-Type -AssemblyName System.Data

# ------------------------------------------------------------------ voltar
if ($Voltar) {
    if ($Voltar -notmatch '^\d+\.\d+\.\d+$') { Recusar ('versão inválida: ' + $Voltar) }
    $pasta = Join-Path $versoes $Voltar
    if (-not (Test-Path -LiteralPath $pasta)) { Recusar ('a versão ' + $Voltar + ' não foi publicada nesta pasta (IA\conector\versoes).') }
    $p = Test-PastaVersao $pasta
    if ($p) { Recusar ('a versão ' + $Voltar + ' não está íntegra: ' + $p) }
    if (Test-Path -LiteralPath $Banco) {
        $esq = Get-EsquemaDoBanco $Banco
        $sup = EsquemasDe $pasta
        if ($sup -notcontains $esq) { Recusar ('volta incompatível: a versão ' + $Voltar + ' suporta o esquema ' + ($sup -join ', ') + ' e o banco está no ' + $esq + '.') }
    }
    Set-Ponteiro $Voltar 'VOLTA'
    Saida 'Feche e abra o Claude Desktop (e o Codex): a sessão nova carrega a versão do ponteiro.'
    exit 0
}

# ------------------------------------------------------------------ publicar
$versao = VersaoDe (Join-Path $PSScriptRoot 'servidor.ps1')
if ($versao -notmatch '^\d+\.\d+\.\d+$') { Recusar 'não achei $script:VersaoServidor em IA\mcp\servidor.ps1' }
$arquivos = Get-ArquivosDaVersao
$destino = Join-Path $versoes $versao
# restos de promoção interrompida não são versão: saem
Get-ChildItem -LiteralPath $versoes -Directory -Filter '*.tmp-*' -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
if (Test-Path -LiteralPath $destino) {
    foreach ($rel in $arquivos.Keys) {
        $alvo = Join-Path $destino $rel
        if (-not (Test-Path -LiteralPath $alvo) -or (Hash $alvo) -ne (Hash $arquivos[$rel])) {
            Recusar ('a versão ' + $versao + ' já foi publicada com outro conteúdo (' + $rel + '). Versão publicada não muda: suba $script:VersaoServidor.')
        }
    }
    Saida ('a versão ' + $versao + ' já está publicada com este mesmo conteúdo')
} else {
    $tmp = Join-Path $versoes ($versao + '.tmp-' + [guid]::NewGuid().ToString('N'))
    $lista = New-Object System.Collections.ArrayList
    foreach ($rel in $arquivos.Keys) {
        $alvo = Join-Path $tmp $rel
        New-Item -ItemType Directory -Path (Split-Path -Parent $alvo) -Force | Out-Null
        Copy-Item -LiteralPath $arquivos[$rel] -Destination $alvo
        if ((Hash $alvo) -ne (Hash $arquivos[$rel])) { Remove-Item -LiteralPath $tmp -Recurse -Force; Recusar ('a cópia de ' + $rel + ' não conferiu') }
        [void]$lista.Add([ordered]@{ caminho = $rel; sha256 = (Hash $alvo) })
    }
    $man = [ordered]@{ versao = $versao; publicado = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'); por = ($env:COMPUTERNAME + '\' + $env:USERNAME); arquivos = @($lista) }
    [System.IO.File]::WriteAllText((Join-Path $tmp 'manifesto.json'), (ConvertTo-Json -InputObject $man -Depth 4), (New-Object System.Text.UTF8Encoding($false)))
    if ($env:CRM_FALHA_PUBLICAR_CONECTOR -eq 'antes-de-renomear') { throw 'falha simulada antes de renomear (teste)' }
    Rename-Item -LiteralPath $tmp -NewName $versao
    Saida ('versão ' + $versao + ' copiada e conferida em ' + $destino)
}

if (-not $SemTeste) {
    # testar-mcp contra a versão nova, num banco de homologação temporário (nunca o da rede)
    $tmpB = Join-Path ([System.IO.Path]::GetTempPath()) ('crm-pubcon-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $tmpB -Force | Out-Null
    try {
        $bancoT = Join-Path $tmpB 'crm_zapromaq.accdb'
        $pref = $ErrorActionPreference; $ErrorActionPreference = 'Continue'
        try {
            & $ps -NoProfile -ExecutionPolicy Bypass -File (Join-Path $raizIA 'teste\criar-banco-teste.ps1') -Banco $bancoT -Recriar 2>&1 | Out-Null
            $saida = @(& $ps -NoProfile -ExecutionPolicy Bypass -File (Join-Path $raizIA 'teste\testar-mcp.ps1') -Servidor (Join-Path $destino 'servidor.ps1') -Banco $bancoT 2>&1 | ForEach-Object { [string]$_ })
            $codigo = $LASTEXITCODE
        } finally { $ErrorActionPreference = $pref }
        if ($codigo -ne 0) {
            $saida | Where-Object { $_ -match 'FALHA|ERRO' } | Select-Object -First 20 | ForEach-Object { Write-Host ('  ' + $_) }
            Recusar ('o testar-mcp falhou contra a versão ' + $versao + '; o ponteiro NÃO foi trocado.')
        }
        Saida ('testar-mcp aprovado contra a versão ' + $versao + ' (banco de homologação temporário)')
    } finally { Remove-Item -LiteralPath $tmpB -Recurse -Force -ErrorAction SilentlyContinue }
}
Set-Ponteiro $versao 'PUBLICACAO'
Saida 'Feche e abra o Claude Desktop (e o Codex): a sessão nova carrega a versão do ponteiro.'
exit 0
