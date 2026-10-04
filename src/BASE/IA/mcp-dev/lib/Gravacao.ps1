<#
  Gravacao.ps1 - gravar_fonte do MCP de desenvolvimento (separado de propósito:
  o antivírus da rede bloqueava o arquivo único que listava, lia e gravava).
  Usa Resolve-CaminhoFonte/Get-HashTexto de Fontes.ps1.
#>

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
