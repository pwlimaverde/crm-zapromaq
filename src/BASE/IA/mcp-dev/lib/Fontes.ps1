<#
  Fontes.ps1 - base do MCP de desenvolvimento (D12, Fase 7): caminhos, leitura e listagem das fontes.

  Para ajustes pequenos pelo Claude Desktop da empresa, sem Node e sem editor:
    - lê e altera SÓ SISTEMA\DADOS\fonte\ (nunca build\, banco, IA\ ou a raiz);
    - arquivos gerados (assets\gerado\, modTema.bas) não se editam: são regerados;
    - gravar exige o hash lido (controle otimista, como no banco) e guarda cópia
      do anterior em execucao\dev\copias\ + registro em execucao\dev\alteracoes.log;
    - codificação certa por tipo: VBA em Windows-1252 + CRLF; JSON/MD em UTF-8;
    - prévia devolve a IMAGEM da tela; montar só gera o .xlsm de TESTE;
      publicar só com confirmação explícita.
  Arquivos: Fontes.ps1 (listar/ler fontes), Gravacao.ps1 (gravar_fonte),
  Previa.ps1 (prévia, montagem, publicação, logs), Processos.ps1
  (scripts auxiliares), Migracoes.ps1 (banco e conferência), Catalogo.ps1.
  O registro de alterações é o que se usa para trazer os ajustes de volta ao
  repositório de desenvolvimento (docs\sincronizar-rede.md).
#>

$script:SoLeitura = [ordered]@{ readOnlyHint = $true; destructiveHint = $false; openWorldHint = $false }
$script:Grava = [ordered]@{ readOnlyHint = $false; destructiveHint = $false; idempotentHint = $false; openWorldHint = $false }
$script:Publica = [ordered]@{ readOnlyHint = $false; destructiveHint = $true; idempotentHint = $false; openWorldHint = $false }

$script:ExtVba = @('.bas', '.cls', '.txt')
$script:ExtTexto = @('.json', '.md')
$script:Cp1252 = [System.Text.Encoding]::GetEncoding(1252)
$script:Utf8 = New-Object System.Text.UTF8Encoding($false)

function Test-TemArg($a, [string]$nome) { return ($null -ne $a -and ($a.PSObject.Properties.Name -contains $nome)) }
function Get-Arg($a, [string]$nome, $padrao = $null) { if (Test-TemArg $a $nome) { return $a.$nome }; return $padrao }

# ------------------------------------------------------------------ caminhos
# Caminho relativo a fonte\ -> caminho completo, recusando qualquer saída da pasta
# (.., caminho absoluto, unidade, UNC).
function Resolve-CaminhoFonte([string]$relativo, [switch]$ParaGravar) {
    $r = ([string]$relativo).Trim().Replace('/', '\')
    if ($r -eq '' -or $r.StartsWith('\') -or $r -match '^[A-Za-z]:' -or $r -match '(^|\\)\.\.(\\|$)') {
        throw ("Caminho inválido: '" + $relativo + "'. Use um caminho relativo a fonte\, ex.: formularios\frmCRM.txt")
    }
    $raiz = [System.IO.Path]::GetFullPath($script:PastaFonte).TrimEnd('\') + '\'
    $cheio = [System.IO.Path]::GetFullPath((Join-Path $raiz $r))
    if (-not $cheio.StartsWith($raiz, [StringComparison]::OrdinalIgnoreCase)) { throw ("Caminho fora de fonte\: '" + $relativo + "'.") }
    $ext = [System.IO.Path]::GetExtension($cheio).ToLowerInvariant()
    if (($script:ExtVba + $script:ExtTexto) -notcontains $ext) {
        throw ("Tipo de arquivo não editável por aqui: '" + $ext + "'. Aceitos: " + (($script:ExtVba + $script:ExtTexto) -join ' '))
    }
    if ($ParaGravar) {
        $rel = $cheio.Substring($raiz.Length)
        if ($rel -like 'assets\gerado\*' -or $rel -ieq 'modulos\modTema.bas') {
            throw ("'" + $rel + "' é GERADO por build\gerar-visual.ps1 a partir de layout\: altere layout\tema.json ou layout\icones.json e rode gerar_previa.")
        }
    }
    return $cheio
}

function Get-RelativoFonte([string]$cheio) {
    $raiz = [System.IO.Path]::GetFullPath($script:PastaFonte).TrimEnd('\') + '\'
    return $cheio.Substring($raiz.Length)
}

function Test-EhVba([string]$cheio) {
    $ext = [System.IO.Path]::GetExtension($cheio).ToLowerInvariant()
    return ($script:ExtVba -contains $ext)
}

function Get-HashTexto([byte[]]$bytes) {
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($sha.ComputeHash($bytes))).Replace('-', '').Substring(0, 16) } finally { $sha.Dispose() }
}

# ------------------------------------------------------------------ ferramentas
function Invoke-ListarFontes($a) {
    $filtro = [string](Get-Arg $a 'pasta' '')
    $raiz = [System.IO.Path]::GetFullPath($script:PastaFonte)
    $itens = Get-ChildItem -LiteralPath $raiz -Recurse -File | Where-Object {
        (($script:ExtVba + $script:ExtTexto) -contains $_.Extension.ToLowerInvariant()) -and ($_.FullName -notlike '*\assets\gerado\*')
    } | ForEach-Object {
        [ordered]@{ caminho = (Get-RelativoFonte $_.FullName); tamanho = $_.Length; alterado_em = $_.LastWriteTime.ToString('yyyy-MM-dd HH:mm') }
    }
    if ($filtro) { $itens = @($itens | Where-Object { $_.caminho -like ($filtro.TrimEnd('\') + '\*') }) }
    return [ordered]@{ pasta = 'SISTEMA\DADOS\fonte'; arquivos = @($itens) }
}

function Invoke-LerFonte($a) {
    $cheio = Resolve-CaminhoFonte ([string]$a.caminho)
    if (-not (Test-Path -LiteralPath $cheio)) { throw ('Não existe: ' + (Get-RelativoFonte $cheio)) }
    $bytes = [System.IO.File]::ReadAllBytes($cheio)
    $texto = if (Test-EhVba $cheio) { $script:Cp1252.GetString($bytes) } else { $script:Utf8.GetString($bytes).TrimStart([char]0xFEFF) }
    return [ordered]@{
        caminho = (Get-RelativoFonte $cheio)
        versao = (Get-HashTexto $bytes)
        codificacao = $(if (Test-EhVba $cheio) { 'Windows-1252 + CRLF' } else { 'UTF-8' })
        conteudo = $texto
    }
}
