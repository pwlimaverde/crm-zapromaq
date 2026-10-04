<#
  Migracoes.ps1 - estado do sistema e migrações do banco
  (MCP de desenvolvimento). Usa Get-Arg/Get-HashTexto de Fontes.ps1 e
  Invoke-ScriptFilho/Invoke-ProcessoFilho de Processos.ps1.
#>

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
