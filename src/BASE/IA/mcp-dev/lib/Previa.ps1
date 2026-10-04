<#
  Previa.ps1 - prévia, conferência de layout, montagem, publicação, registro de
  alterações e logs (MCP de desenvolvimento).
#>

# ------------------------------------------------------------------ prévia e montagem
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
