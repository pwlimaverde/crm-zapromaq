<#
  FerramentasDev.ps1 - catálogo FECHADO do MCP de desenvolvimento (D12, Fase 7).

  Para ajustes pequenos pelo Claude Desktop da empresa, sem Node e sem editor:
    - lê e altera SÓ SISTEMA\DADOS\fonte\ (nunca build\, banco, IA\ ou a raiz);
    - arquivos gerados (assets\gerado\, modTema.bas) não se editam: são regerados;
    - gravar exige o hash lido (controle otimista, como no banco) e guarda cópia
      do anterior em execucao\dev\copias\ + registro em execucao\dev\alteracoes.log;
    - codificação certa por tipo: VBA em Windows-1252 + CRLF; JSON/MD em UTF-8;
    - prévia devolve a IMAGEM da tela; montar só gera o .xlsm de TESTE;
      publicar só com confirmação explícita.
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

# ------------------------------------------------------------------ processos (nada no stdout do MCP)
# Roda um script PowerShell 64 bits e devolve código de saída e texto. A saída do
# filho é capturada: se vazasse para o stdout deste servidor, quebraria o protocolo.
function Invoke-ScriptFilho([string]$script, [string]$argumentos, [int]$limiteSeg) {
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $psi.Arguments = '-NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "' + $script + '" ' + $argumentos
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.RedirectStandardInput = $true
    $psi.CreateNoWindow = $true
    $p = [System.Diagnostics.Process]::Start($psi)
    $p.StandardInput.Close()
    $out = $p.StandardOutput.ReadToEndAsync()
    $err = $p.StandardError.ReadToEndAsync()
    if (-not $p.WaitForExit($limiteSeg * 1000)) {
        try { $p.Kill() } catch { }
        throw ('Passou de ' + $limiteSeg + ' s e foi interrompido: ' + [System.IO.Path]::GetFileName($script))
    }
    $p.WaitForExit()
    return [ordered]@{ codigo = $p.ExitCode; saida = ($out.Result + $err.Result) }
}

# Últimas n linhas de um texto (relatório curto para o modelo).
function Get-Final([string]$texto, [int]$n) {
    $l = @($texto -split "`r?`n" | Where-Object { $_.Trim() -ne '' })
    if ($l.Count -le $n) { return ($l -join "`n") }
    return ($l[($l.Count - $n)..($l.Count - 1)] -join "`n")
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

function Invoke-GravarFonte($a) {
    $cheio = Resolve-CaminhoFonte ([string]$a.caminho) -ParaGravar
    $rel = Get-RelativoFonte $cheio
    $esperada = [string](Get-Arg $a 'versao' '')
    $existe = Test-Path -LiteralPath $cheio
    $antes = $null
    if ($existe) {
        $antes = [System.IO.File]::ReadAllBytes($cheio)
        $atual = Get-HashTexto $antes
        if ($esperada -eq '') { throw ("'" + $rel + "' já existe: leia com ler_fonte e passe a 'versao' devolvida.") }
        if ($esperada -ne $atual) {
            throw ("Conflito: '" + $rel + "' mudou depois da leitura (versão lida " + $esperada + ', atual ' + $atual + '). Leia de novo e refaça a alteração.')
        }
    } elseif ($esperada -ne '') {
        throw ("'" + $rel + "' não existe (para criar arquivo novo, não passe 'versao').")
    }

    $texto = [string]$a.conteudo
    if (Test-EhVba $cheio) {
        # VBA: Windows-1252 + CRLF; caractere fora da página de código viraria '?'
        $texto = ($texto -replace "`r`n", "`n") -replace "`n", "`r`n"
        $volta = $script:Cp1252.GetString($script:Cp1252.GetBytes($texto))
        if ($volta -ne $texto) {
            for ($i = 0; $i -lt $texto.Length; $i++) { if ($volta[$i] -ne $texto[$i]) { break } }
            throw ('Caractere fora do Windows-1252 na posição ' + $i + " ('" + $texto[$i] + '''): o VBA não o grava. Use ChrW$(&H' + ('{0:X4}' -f [int]$texto[$i]) + ').')
        }
        $novo = $script:Cp1252.GetBytes($texto)
    } else {
        if ([System.IO.Path]::GetExtension($cheio) -ieq '.json') {
            try { [void]($texto | ConvertFrom-Json) } catch { throw ('JSON inválido: ' + $_.Exception.Message) }
        }
        $texto = $texto -replace "`r`n", "`n"
        $novo = $script:Utf8.GetBytes($texto)
    }

    $quando = Get-Date
    $pastaDev = Join-Path $script:PastaDados 'execucao\dev'
    if ($existe) {
        $copia = Join-Path $pastaDev ('copias\' + $quando.ToString('yyyyMMdd-HHmmss') + '\' + $rel)
        New-Item -ItemType Directory -Path (Split-Path -Parent $copia) -Force | Out-Null
        [System.IO.File]::WriteAllBytes($copia, $antes)
    } else {
        New-Item -ItemType Directory -Path (Split-Path -Parent $cheio) -Force | Out-Null
    }
    [System.IO.File]::WriteAllBytes($cheio, $novo)
    $nova = Get-HashTexto $novo
    $registro = [ordered]@{ quando = $quando.ToString('yyyy-MM-dd HH:mm:ss'); usuario = $env:USERNAME; maquina = $env:COMPUTERNAME
                            caminho = $rel; versao_antes = $(if ($existe) { Get-HashTexto $antes } else { '' }); versao_depois = $nova
                            motivo = [string](Get-Arg $a 'motivo' '') }
    New-Item -ItemType Directory -Path $pastaDev -Force | Out-Null
    Add-Content -LiteralPath (Join-Path $pastaDev 'alteracoes.log') -Value (ConvertTo-Json -InputObject $registro -Compress) -Encoding UTF8
    return [ordered]@{ caminho = $rel; versao = $nova; criado = (-not $existe)
                       proximo_passo = 'Rode verificar_layout ou gerar_previa (tela) para conferir; montar_teste gera o .xlsm de teste.' }
}

function Invoke-GerarPrevia($a) {
    $tela = [string]$a.tela
    $escala = [int](Get-Arg $a 'escala' 100)
    $r = Invoke-ScriptFilho (Join-Path $script:PastaDados 'build\gerar-visual.ps1') ('-Tela ' + $tela + ' -Tolerante') 240
    $png = Join-Path $script:PastaDados ('execucao\previa\' + $tela + '-' + $escala + '.png')
    $relatorio = Get-Final $r.saida 14
    if (-not (Test-Path -LiteralPath $png)) { throw ('A prévia não foi gerada:' + "`n" + $relatorio) }
    $b64 = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($png))
    return [ordered]@{ __conteudo = @(
        [ordered]@{ type = 'text'; text = ('Prévia de ' + $tela + ' a ' + $escala + '% (escala do Windows). Relatório do gerador:' + "`n" + $relatorio) },
        [ordered]@{ type = 'image'; data = $b64; mimeType = 'image/png' }
    ) }
}

function Invoke-VerificarLayout($a) {
    $arg = ''
    if (Test-TemArg $a 'tela') { $arg = '-Tela ' + [string]$a.tela }
    $r = Invoke-ScriptFilho (Join-Path $script:PastaDados 'build\gerar-visual.ps1') $arg 300
    $linhas = @($r.saida -split "`r?`n" | Where-Object { $_ -match 'PROBLEMA|conferência|RESULTADO|^---' })
    return [ordered]@{ sem_problemas = ($r.codigo -eq 0); relatorio = ($linhas -join "`n") }
}

function Invoke-MontarTeste($a) {
    $r = Invoke-ScriptFilho (Join-Path $script:PastaDados 'build\montar-frontend.ps1') '-Teste' 600
    return [ordered]@{ sucesso = ($r.codigo -eq 0); arquivo = 'SISTEMA\DADOS\execucao\teste\CRM_Zapromaq.xlsm'
                       relatorio = (Get-Final $r.saida 30) }
}

function Invoke-Publicar($a) {
    if ([string]$a.confirmacao -cne 'PUBLICAR') { throw "Para publicar, passe confirmacao = 'PUBLICAR' (depois de o usuário confirmar)." }
    $r = Invoke-ScriptFilho (Join-Path $script:PastaDados 'build\montar-frontend.ps1') '' 600
    return [ordered]@{ sucesso = ($r.codigo -eq 0); relatorio = (Get-Final $r.saida 30)
                       lembrete = 'Ajuste feito na rede precisa voltar ao repositório de desenvolvimento: veja listar_alteracoes.' }
}

function Invoke-ListarAlteracoes($a) {
    $arq = Join-Path $script:PastaDados 'execucao\dev\alteracoes.log'
    if (-not (Test-Path -LiteralPath $arq)) { return [ordered]@{ alteracoes = @() } }
    $itens = @(Get-Content -LiteralPath $arq -Encoding UTF8 | Where-Object { $_.Trim() } | ForEach-Object { $_ | ConvertFrom-Json })
    $n = [int](Get-Arg $a 'limite' 50)
    if ($itens.Count -gt $n) { $itens = $itens[($itens.Count - $n)..($itens.Count - 1)] }
    return [ordered]@{ alteracoes = $itens }
}

function Invoke-VerLog($a) {
    $tipo = [string]$a.tipo
    $arq = Get-ChildItem -LiteralPath (Join-Path $script:PastaDados 'execucao\logs') -Filter ($tipo + '-*.txt') -ErrorAction SilentlyContinue |
           Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $arq) { throw ('Nenhum log de ' + $tipo + ' ainda.') }
    return [ordered]@{ arquivo = $arq.Name; conteudo = (Get-Final ([System.IO.File]::ReadAllText($arq.FullName)) 80) }
}


# ------------------------------------------------------------------ banco: migrações
# Alterar a estrutura do banco (campo novo, tabela nova) NUNCA se faz à mão:
# vira um arquivo numerado em banco\migracoes, aplicado uma vez por banco e
# registrado em config. O mesmo arquivo roda aqui, no teste e na produção.
function Get-PastaMigracoes { return (Join-Path $script:PastaDados 'banco\migracoes') }

function Get-EstadoBanco {
    $banco = Join-Path (Split-Path -Parent (Split-Path -Parent $script:PastaDados)) 'crm_zapromaq.accdb'
    $r = [ordered]@{ caminho = $banco; existe = (Test-Path -LiteralPath $banco); em_uso = $false
                     versao_esquema = ''; versao_front = ''; ambiente = ''; migracoes_aplicadas = @() }
    if (-not $r.existe) { return $r }
    $r.em_uso = Test-Path -LiteralPath ([System.IO.Path]::ChangeExtension($banco, 'laccdb'))
    . (Join-Path $script:PastaDados 'lib\Comum.ps1')
    $cn = Open-Banco $banco
    try {
        $leitor = (New-ComandoBanco $cn 'SELECT chave, valor FROM config' @() $null).ExecuteReader()
        $ap = New-Object System.Collections.ArrayList
        while ($leitor.Read()) {
            $k = [string]$leitor.GetValue(0); $v = [string]$leitor.GetValue(1)
            switch -Wildcard ($k) {
                'versao_esquema' { $r.versao_esquema = $v }
                'versao_front'   { $r.versao_front = $v }
                'ambiente'       { $r.ambiente = $v }
                'migracao.*'     { [void]$ap.Add($k.Substring(9) + '  (' + $v + ')') }
            }
        }
        $leitor.Close()
        $r.migracoes_aplicadas = @($ap | Sort-Object)
    } finally { $cn.Close(); $cn.Dispose() }
    return $r
}

function Invoke-EstadoSistema($a) {
    $b = Get-EstadoBanco
    $versaoArq = ''
    $vt = Join-Path $script:PastaDados 'VERSAO.txt'
    if (Test-Path -LiteralPath $vt) { $versaoArq = (Get-Content -LiteralPath $vt -Raw).Trim() }
    $publicada = ''
    $vf = Join-Path (Split-Path -Parent (Split-Path -Parent $script:PastaDados)) 'VERSAO-FRONT.txt'
    if (Test-Path -LiteralPath $vf) { $publicada = (Get-Content -LiteralPath $vf -TotalCount 1).Trim() }
    $pend = @(Get-Migracoes | Where-Object { -not $_.aplicada } | ForEach-Object { $_.arquivo })
    return [ordered]@{
        versao_proxima_publicacao = $versaoArq
        versao_publicada_na_raiz = $publicada
        versao_front_no_banco = $b.versao_front
        versao_esquema = $b.versao_esquema
        ambiente = $b.ambiente
        banco = $b.caminho
        banco_em_uso = $b.em_uso
        migracoes_pendentes = $pend
        migracoes_aplicadas = $b.migracoes_aplicadas
        excel_instalado = ($null -ne [type]::GetTypeFromProgID('Excel.Application'))
        python_instalado = ($null -ne (Get-Command python -ErrorAction SilentlyContinue))
    }
}

function Get-Migracoes {
    $b = Get-EstadoBanco
    $ids = @($b.migracoes_aplicadas | ForEach-Object { ($_ -split '  ')[0] })
    return @(Get-ChildItem -LiteralPath (Get-PastaMigracoes) -Filter '*.ps1' -File |
             Where-Object { $_.Name -match '^\d{3}-' } | Sort-Object Name | ForEach-Object {
        $id = $_.Name.Substring(0, 3)
        $desc = ''
        try { $m = & $_.FullName; $desc = [string]$m.Descricao } catch { $desc = '(arquivo com erro: ' + $_.Exception.Message + ')' }
        [ordered]@{ id = $id; arquivo = $_.Name; descricao = $desc; aplicada = ($ids -contains $id) }
    })
}

function Invoke-ListarMigracoes($a) {
    return [ordered]@{ pasta = 'SISTEMA\DADOS\banco\migracoes'; migracoes = @(Get-Migracoes) }
}

function Invoke-LerMigracao($a) {
    $arq = Join-Path (Get-PastaMigracoes) ([System.IO.Path]::GetFileName([string]$a.arquivo))
    if (-not (Test-Path -LiteralPath $arq)) { throw ('não existe: ' + $a.arquivo) }
    $bytes = [System.IO.File]::ReadAllBytes($arq)
    return [ordered]@{ arquivo = (Split-Path -Leaf $arq); versao = (Get-HashTexto $bytes)
                       conteudo = $script:Utf8.GetString($bytes).TrimStart([char]0xFEFF) }
}

# SQL perigoso exige confirmação explícita do usuário.
function Assert-SqlSeguro([string]$sql, [string]$confirmacao) {
    $s = ' ' + ($sql -replace '\s+', ' ').ToUpperInvariant() + ' '
    $perigo = ''
    if ($s -match ' DROP TABLE | DROP COLUMN | DROP INDEX ') { $perigo = 'apaga estrutura (DROP)' }
    elseif ($s -match ' DELETE FROM ' -and $s -notmatch ' WHERE ') { $perigo = 'apaga TODAS as linhas (DELETE sem WHERE)' }
    elseif ($s -match ' UPDATE ' -and $s -notmatch ' WHERE ') { $perigo = 'altera TODAS as linhas (UPDATE sem WHERE)' }
    if ($perigo -and $confirmacao -cne 'CONFIRMO') {
        throw ('Este comando ' + $perigo + ". Mostre ao usuário o que vai acontecer e, se ele confirmar, repita com confirmacao = 'CONFIRMO'.")
    }
}

function Invoke-CriarMigracao($a) {
    $nome = ([string]$a.nome).Trim().ToLowerInvariant() -replace '[^a-z0-9]+', '-'
    $nome = $nome.Trim('-')
    if ($nome -eq '') { throw 'nome vazio' }
    $descricao = ([string]$a.descricao).Trim()
    if ($descricao -eq '') { throw 'descricao vazia (ela aparece no log da migração)' }
    $estrutural = ([string](Get-Arg $a 'tipo' 'dados') -eq 'estrutura')
    $sqls = @(Get-Arg $a 'sql' @())
    if ($sqls.Count -eq 0) { throw 'informe ao menos um comando SQL' }
    foreach ($q in $sqls) { Assert-SqlSeguro ([string]$q) ([string](Get-Arg $a 'confirmacao' '')) }

    $estado = Get-EstadoBanco
    $de = ''; $para = ''
    if ($estrutural) {
        $de = $estado.versao_esquema
        $para = [string](Get-Arg $a 'esquema_para' '')
        if ($para -eq '') { throw "migração de estrutura precisa de 'esquema_para' (ex.: 1.1): mudar a estrutura obriga a publicar front e IA na mesma versão" }
    }

    $existentes = @(Get-ChildItem -LiteralPath (Get-PastaMigracoes) -Filter '*.ps1' -File | Where-Object { $_.Name -match '^\d{3}-' })
    $proximo = 1
    if ($existentes.Count -gt 0) { $proximo = 1 + [int](@($existentes | Sort-Object Name)[-1].Name.Substring(0, 3)) }
    $id = '{0:000}' -f $proximo
    $arq = Join-Path (Get-PastaMigracoes) ($id + '-' + $nome + '.ps1')

    $corpo = ''
    foreach ($q in $sqls) {
        $texto = ([string]$q).Trim() -replace '\s+', ' '
        $corpo += "        [void](Invoke-Comando `$cn '" + $texto.Replace("'", "''") + "' @() `$tx)`r`n"
        $corpo += "        `$log.L('  ok: " + $texto.Replace("'", "''").Substring(0, [Math]::Min(70, $texto.Length)) + "')`r`n"
    }
    $cab = if ($estrutural) {
        "  ESTRUTURA: exige o banco exclusivo (ninguém usando) e muda versao_esquema`r`n" +
        "  de $de para $para. Depois desta migração, o front e a pasta IA precisam ser`r`n" +
        "  publicados na mesma versão - o sistema recusa gravar com esquema diferente."
    } else {
        "  Só DADOS: roda numa transação e pode ser aplicada com gente usando."
    }
    $conteudo = "<#`r`n  $id - $descricao`r`n`r`n$cab`r`n`r`n" +
                "  Criada em " + (Get-Date -Format 'dd/MM/yyyy HH:mm') + " por " + $env:USERNAME +
                " pelo MCP de desenvolvimento.`r`n#>`r`n@{`r`n" +
                "    Id = '$id'`r`n" +
                "    Descricao = '" + $descricao.Replace("'", "''") + "'`r`n" +
                "    EsquemaDe = " + $(if ($estrutural) { "'$de'" } else { '$null' }) + "`r`n" +
                "    EsquemaPara = " + $(if ($estrutural) { "'$para'" } else { '$null' }) + "`r`n" +
                "    Aplicar = {`r`n        param(`$cn, `$tx, `$log)`r`n" + $corpo + "    }`r`n}`r`n"
    [System.IO.File]::WriteAllText($arq, $conteudo, (New-Object System.Text.UTF8Encoding($true)))

    $proximo_passo = 'Confira com ler_migracao, aplique com aplicar_migracoes (use simular=true antes).'
    if ($estrutural) {
        $proximo_passo += ' Campo novo também precisa entrar em fonte\modulos\modSchema.bas (CamposDe/ColunasPlanilha),' +
                          ' nas consultas de modCRM.bas que listam as colunas, e no esquema de banco\esquema\criar-banco.ps1;' +
                          ' depois verificar_projeto, montar_teste e publicar.'
    }
    return [ordered]@{ arquivo = (Split-Path -Leaf $arq); tipo = $(if ($estrutural) { 'estrutura' } else { 'dados' })
                       esquema = $(if ($estrutural) { "$de -> $para" } else { '(não muda)' })
                       proximo_passo = $proximo_passo }
}

function Invoke-AplicarMigracoes($a) {
    $simular = [bool](Get-Arg $a 'simular' $false)
    $arg = if ($simular) { '-Simular' } else { '' }
    $r = Invoke-ScriptFilho (Join-Path $script:PastaDados 'banco\aplicar-migracoes.ps1') $arg 300
    return [ordered]@{ sucesso = ($r.codigo -eq 0); simulacao = $simular; relatorio = (Get-Final $r.saida 40) }
}

# ------------------------------------------------------------------ conferência
# verificar.py cobre tudo (codificação, sintaxe VBA por parser, controles x
# layout, modSchema x DDL). Sem Python na máquina, sobra a conferência básica
# em PowerShell - melhor que nada, e o build ainda compila e roda o autoteste.
function Invoke-VerificarProjeto($a) {
    $py = Get-Command python -ErrorAction SilentlyContinue
    $arqPy = Join-Path $script:PastaDados 'ferramentas\verificacao\verificar.py'
    if ($py -and (Test-Path -LiteralPath $arqPy)) {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $py.Source
        $psi.Arguments = '"' + (Join-Path $script:PastaDados 'ferramentas\verificacao\verificar.py') + '"'
        $psi.UseShellExecute = $false; $psi.RedirectStandardOutput = $true; $psi.RedirectStandardError = $true
        $psi.StandardOutputEncoding = New-Object System.Text.UTF8Encoding($false)
        $psi.CreateNoWindow = $true
        $p = [System.Diagnostics.Process]::Start($psi)
        $out = $p.StandardOutput.ReadToEnd() + $p.StandardError.ReadToEnd()
        $p.WaitForExit()
        return [ordered]@{ ferramenta = 'verificar.py (completa)'; sem_problemas = ($p.ExitCode -eq 0); relatorio = (Get-Final $out 60) }
    }
    $problemas = New-Object System.Collections.ArrayList
    $cp1252 = [System.Text.Encoding]::GetEncoding(1252)
    # -Include é ignorado junto com -LiteralPath: filtrar pela extensão na mão.
    $vba = @(Get-ChildItem -LiteralPath $script:PastaFonte -Recurse -File | Where-Object { $script:ExtVba -contains $_.Extension.ToLowerInvariant() })
    foreach ($f in $vba) {
        $bytes = [System.IO.File]::ReadAllBytes($f.FullName)
        if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB) { [void]$problemas.Add($f.Name + ': está em UTF-8; o VBA precisa de Windows-1252') }
        $texto = $cp1252.GetString($bytes)
        if ($texto -match "[^`r]`n") { [void]$problemas.Add($f.Name + ': linha sem CRLF') }
    }
    foreach ($f in @(Get-ChildItem -LiteralPath (Join-Path $script:PastaFonte 'layout') -Recurse -File | Where-Object { $_.Extension -ieq '.json' })) {
        try { [void]((Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8) | ConvertFrom-Json) }
        catch { [void]$problemas.Add($f.Name + ': JSON inválido - ' + $_.Exception.Message) }
    }
    return [ordered]@{ ferramenta = 'conferência básica (sem Python nesta máquina)'
                       sem_problemas = ($problemas.Count -eq 0); problemas = @($problemas)
                       observacao = 'A conferência completa (sintaxe VBA, controles x layout, modSchema x DDL) precisa do Python. O montar_teste ainda compila o projeto e roda o autoteste.' }
}

# ------------------------------------------------------------------ catálogo
$script:Ferramentas = @(
    [ordered]@{
        name = 'listar_fontes'; title = 'Listar fontes do front'; handler = 'Invoke-ListarFontes'; annotations = $script:SoLeitura
        description = 'Lista os arquivos editáveis de SISTEMA\DADOS\fonte (VBA .bas/.cls/.txt, layout .json). Use "pasta" para filtrar (ex.: layout\formularios).'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; properties = [ordered]@{
            pasta = [ordered]@{ type = 'string'; maxLength = 120; description = 'Subpasta de fonte\, ex.: modulos' } } }
    },
    [ordered]@{
        name = 'ler_fonte'; title = 'Ler um arquivo de fonte'; handler = 'Invoke-LerFonte'; annotations = $script:SoLeitura
        description = 'Devolve o conteúdo e a "versao" (hash) de um arquivo de fonte\. A versao é obrigatória para gravar_fonte.'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; required = @('caminho'); properties = [ordered]@{
            caminho = [ordered]@{ type = 'string'; maxLength = 200; description = 'Relativo a fonte\, ex.: layout\formularios\frmCRM.json' } } }
    },
    [ordered]@{
        name = 'gravar_fonte'; title = 'Gravar um arquivo de fonte'; handler = 'Invoke-GravarFonte'; annotations = $script:Grava
        description = 'Substitui o conteúdo INTEIRO de um arquivo de fonte\ (ou cria um novo, sem "versao"). Exige a versao lida em ler_fonte; guarda cópia do anterior. VBA é gravado em Windows-1252 + CRLF. Arquivos gerados (assets\gerado, modTema.bas) são recusados. Mostre a alteração ao usuário antes.'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; required = @('caminho', 'conteudo'); properties = [ordered]@{
            caminho = [ordered]@{ type = 'string'; maxLength = 200 }
            conteudo = [ordered]@{ type = 'string'; maxLength = 400000 }
            versao = [ordered]@{ type = 'string'; maxLength = 16; description = 'Versão devolvida por ler_fonte (omitir só para arquivo novo)' }
            motivo = [ordered]@{ type = 'string'; maxLength = 300; description = 'O que foi ajustado e por quê (vai para o registro de alterações)' } } }
    },
    [ordered]@{
        name = 'gerar_previa'; title = 'Prévia visual de uma tela'; handler = 'Invoke-GerarPrevia'; annotations = $script:SoLeitura
        description = 'Gera a prévia da tela a partir de layout\ (Edge headless) e devolve a IMAGEM, mais o relatório de problemas (texto que não cabe, sobreposição, fora da tela). Também regera fundos e modTema.'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; required = @('tela'); properties = [ordered]@{
            tela = [ordered]@{ type = 'string'; pattern = '^frm[A-Za-z0-9]+$'; description = 'Nome do formulário, ex.: frmCRM' }
            escala = [ordered]@{ type = 'integer'; enum = @(100, 125, 150); default = 100 } } }
    },
    [ordered]@{
        name = 'verificar_layout'; title = 'Conferir layout'; handler = 'Invoke-VerificarLayout'; annotations = $script:SoLeitura
        description = 'Roda o gerador visual em todas as telas (ou em uma) e devolve só a conferência: sem_problemas e a lista do que não cabe ou se sobrepõe.'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; properties = [ordered]@{
            tela = [ordered]@{ type = 'string'; pattern = '^frm[A-Za-z0-9]+$' } } }
    },
    [ordered]@{
        name = 'montar_teste'; title = 'Montar .xlsm de teste'; handler = 'Invoke-MontarTeste'; annotations = $script:Grava
        description = 'Monta o front em SISTEMA\DADOS\execucao\teste\ (compila e roda o autoteste) SEM publicar nem subir a versão. Exige Excel nesta máquina. Leva de 1 a 3 minutos.'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; properties = [ordered]@{} }
    },
    [ordered]@{
        name = 'publicar'; title = 'Publicar o front para todos'; handler = 'Invoke-Publicar'; annotations = $script:Publica
        description = 'Monta e PUBLICA o CRM_Zapromaq.xlsm na raiz (sobe a versão; as estações copiam na próxima abertura). Só com pedido explícito do usuário, depois de montar_teste dar certo. Exige confirmacao = "PUBLICAR".'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; required = @('confirmacao'); properties = [ordered]@{
            confirmacao = [ordered]@{ type = 'string'; enum = @('PUBLICAR') } } }
    },
    [ordered]@{
        name = 'listar_alteracoes'; title = 'Alterações feitas por aqui'; handler = 'Invoke-ListarAlteracoes'; annotations = $script:SoLeitura
        description = 'Registro das gravações feitas por este servidor (quando, quem, arquivo, motivo). É a lista do que precisa voltar ao repositório de desenvolvimento.'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; properties = [ordered]@{
            limite = [ordered]@{ type = 'integer'; minimum = 1; maximum = 500; default = 50 } } }
    },
    [ordered]@{
        name = 'estado_sistema'; title = 'Estado do sistema'; handler = 'Invoke-EstadoSistema'; annotations = $script:SoLeitura
        description = 'Versões (próxima publicação, publicada na raiz, no banco), versão do esquema, ambiente, migrações pendentes e aplicadas, se o banco está em uso e o que existe na máquina (Excel, Python). Use antes de qualquer alteração.'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; properties = [ordered]@{} }
    },
    [ordered]@{
        name = 'listar_migracoes'; title = 'Migrações do banco'; handler = 'Invoke-ListarMigracoes'; annotations = $script:SoLeitura
        description = 'Arquivos de banco\migracoes com id, descrição e se já foram aplicados neste banco.'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; properties = [ordered]@{} }
    },
    [ordered]@{
        name = 'ler_migracao'; title = 'Ler uma migração'; handler = 'Invoke-LerMigracao'; annotations = $script:SoLeitura
        description = 'Conteúdo de um arquivo de migração (para conferir antes de aplicar).'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; required = @('arquivo'); properties = [ordered]@{
            arquivo = [ordered]@{ type = 'string'; maxLength = 120; description = 'Nome do arquivo, ex.: 003-campo-novo.ps1' } } }
    },
    [ordered]@{
        name = 'criar_migracao'; title = 'Criar migração do banco'; handler = 'Invoke-CriarMigracao'; annotations = $script:Grava
        description = 'Cria o próximo arquivo numerado em banco\migracoes a partir dos comandos SQL informados. tipo=dados (roda em transação, pode ser aplicada com gente usando) ou tipo=estrutura (campo/tabela nova: exige banco exclusivo, esquema_para e publicação de front e IA na mesma versão). Alterar a estrutura do banco à mão é proibido: é sempre por migração. DROP, DELETE ou UPDATE sem WHERE exigem confirmacao = "CONFIRMO".'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; required = @('nome', 'descricao', 'sql'); properties = [ordered]@{
            nome = [ordered]@{ type = 'string'; maxLength = 60; description = 'Parte do nome do arquivo, ex.: campo-cnae-em-clientes' }
            descricao = [ordered]@{ type = 'string'; maxLength = 200; description = 'O que a migração faz (aparece no log)' }
            tipo = [ordered]@{ type = 'string'; enum = @('dados', 'estrutura'); default = 'dados' }
            esquema_para = [ordered]@{ type = 'string'; maxLength = 10; description = 'Nova versão do esquema (só em tipo=estrutura), ex.: 1.1' }
            sql = [ordered]@{ type = 'array'; description = 'Comandos SQL na ordem (ACE/Access)'; items = [ordered]@{ type = 'string'; maxLength = 2000 } }
            confirmacao = [ordered]@{ type = 'string'; enum = @('CONFIRMO'); description = 'Obrigatório para comandos destrutivos' } } }
    },
    [ordered]@{
        name = 'aplicar_migracoes'; title = 'Aplicar migrações pendentes'; handler = 'Invoke-AplicarMigracoes'; annotations = $script:Grava
        description = 'Aplica as migrações ainda não aplicadas (faz backup do banco antes). simular=true só mostra o que seria feito. A publicação (publicar) também aplica sozinha o que estiver pendente.'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; properties = [ordered]@{
            simular = [ordered]@{ type = 'boolean'; default = $false } } }
    },
    [ordered]@{
        name = 'verificar_projeto'; title = 'Conferir o projeto'; handler = 'Invoke-VerificarProjeto'; annotations = $script:SoLeitura
        description = 'Conferência estática antes de montar: codificação e sintaxe do VBA, controles citados no código x layout, campos de modSchema x banco. Se a máquina não tiver Python, faz a conferência básica e avisa.'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; properties = [ordered]@{} }
    },
    [ordered]@{
        name = 'ver_log'; title = 'Último log'; handler = 'Invoke-VerLog'; annotations = $script:SoLeitura
        description = 'Final do último log: build (publicação), visual (prévias), migracoes, backup ou ambiente (execucao\logs).'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; required = @('tipo'); properties = [ordered]@{
            tipo = [ordered]@{ type = 'string'; enum = @('build', 'visual', 'migracoes', 'backup', 'ambiente') } } }
    }
)
