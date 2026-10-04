<#
  finalizar-branch.ps1 - substitui o "git flow <tipo> finish" do GitFlow .NET 2.3.0
  (winget Kubis1982.GitFlow), que falha em todo merge --no-ff com
  "Value cannot be null. (Parameter 'message')" e deixa o merge pela metade.
  O "git flow <tipo> start" da ferramenta funciona; só o finish passa por aqui.

  Lê os nomes e prefixos do próprio git-flow (git config gitflow.*):
    feature/ e bugfix/  -> merge --no-ff em develop
    release/ e hotfix/  -> merge --no-ff na main + tag vX.Y + merge --no-ff em develop
  Depois apaga a branch local (e a remota, com -Enviar).

  Se encontrar o merge que o finish da ferramenta deixou preparado (MERGE_HEAD
  igual à ponta da branch), só conclui esse merge.

  Uso (na raiz do repositório ou em qualquer pasta dele):
    powershell -File agent-config\ferramentas\finalizar-branch.ps1 [-Branch feature/x] [-Enviar]
  Sem -Branch, usa a branch atual. -Enviar faz o push de develop (e main + tag) e
  apaga a branch remota, se existir.
#>
param([string]$Branch = '', [switch]$Enviar)
# Continue: no PowerShell 5.1, stderr de comando nativo (git push escreve progresso nele)
# vira erro fatal com Stop. Os erros reais saem por throw, que sempre interrompe.
$ErrorActionPreference = 'Continue'

function G {
    $saida = & git @args 2>&1
    if ($LASTEXITCODE -ne 0) { throw ('git ' + ($args -join ' ') + "`n" + ($saida | Out-String)) }
    return $saida
}
function Cfg([string]$chave, [string]$padrao) {
    $v = & git config --get $chave 2>$null
    if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($v)) { return $padrao }
    return $v.Trim()
}

Set-Location ((G rev-parse --show-toplevel) | Select-Object -First 1)
$main    = Cfg 'gitflow.production' (Cfg 'gitflow.branch.master' 'main')
$develop = Cfg 'gitflow.development' (Cfg 'gitflow.branch.develop' 'develop')
$prefTag = Cfg 'gitflow.prefix.versiontag' (Cfg 'gitflow.prefix.version' 'v')
$tipos = [ordered]@{
    feature = Cfg 'gitflow.prefix.feature' 'feature/'
    bugfix  = Cfg 'gitflow.prefix.bugfix' 'bugfix/'
    release = Cfg 'gitflow.prefix.release' 'release/'
    hotfix  = Cfg 'gitflow.prefix.hotfix' 'hotfix/'
}

$mergeHead = Join-Path (G rev-parse --git-dir | Select-Object -First 1) 'MERGE_HEAD'
if ([string]::IsNullOrWhiteSpace($Branch)) {
    $Branch = (G rev-parse --abbrev-ref HEAD | Select-Object -First 1).Trim()
    if (($Branch -eq $develop -or $Branch -eq $main) -and (Test-Path $mergeHead)) {
        # merge deixado pela metade pelo finish da ferramenta: descobre a branch pela ponta
        $ponta = (Get-Content $mergeHead -TotalCount 1).Trim()
        $Branch = @(G for-each-ref --format='%(refname:short)' --points-at $ponta refs/heads) |
                  Where-Object { $n = $_; @($tipos.Values | Where-Object { $n.StartsWith($_) }).Count -gt 0 } | Select-Object -First 1
        if (-not $Branch) { throw 'Há um merge em andamento que não é de uma branch do git-flow. Resolva com git merge --abort ou git commit.' }
    }
}
$tipo = @($tipos.Keys | Where-Object { $Branch.StartsWith($tipos[$_]) }) | Select-Object -First 1
if (-not $tipo) { throw ("'" + $Branch + "' não é feature/, bugfix/, release/ nem hotfix/.") }
$nome = $Branch.Substring($tipos[$tipo].Length)
& git rev-parse --verify --quiet ('refs/heads/' + $Branch) > $null
if ($LASTEXITCODE -ne 0) { throw ('Branch não encontrada: ' + $Branch) }
$ponta = (G rev-parse $Branch | Select-Object -First 1).Trim()

function Mesclar([string]$destino, [string]$origem, [string]$mensagem) {
    if (Test-Path $mergeHead) {
        $atual = (G rev-parse --abbrev-ref HEAD | Select-Object -First 1).Trim()
        $pendente = (Get-Content $mergeHead -TotalCount 1).Trim()
        if ($atual -eq $destino -and $pendente -eq (G rev-parse $origem | Select-Object -First 1).Trim()) {
            Write-Host ('  concluindo o merge que ficou pela metade em ' + $destino)
            G commit --no-edit -m $mensagem | Out-Null
            return
        }
        throw 'Há outro merge em andamento. Resolva antes (git merge --abort ou git commit).'
    }
    if (@(G status --porcelain --untracked-files=no).Count -gt 0) { throw 'Há alterações não commitadas. Faça commit ou descarte antes de finalizar.' }
    G checkout -q $destino | Out-Null
    & git merge --no-ff --no-edit -m $mensagem $origem
    if ($LASTEXITCODE -ne 0) { throw ('Conflito ao mesclar ' + $origem + ' em ' + $destino + '. Resolva, faça o commit e rode este script de novo.') }
}

Write-Host ('Finalizando ' + $Branch + ' (' + $tipo + ')')
if ($tipo -in @('feature', 'bugfix')) {
    Mesclar $develop $Branch ("Merge branch '" + $Branch + "' into " + $develop)
} else {
    $tag = $prefTag + $nome
    Mesclar $main $Branch ("Merge branch '" + $Branch + "'")
    & git rev-parse --verify --quiet ('refs/tags/' + $tag) > $null
    if ($LASTEXITCODE -eq 0) { throw ('A tag ' + $tag + ' já existe.') }
    G tag -a $tag -m ('Versão ' + $nome) | Out-Null
    Write-Host ('  tag ' + $tag)
    Mesclar $develop $Branch ("Merge branch '" + $Branch + "' into " + $develop)
}
G branch -d $Branch | Out-Null
Write-Host ('  branch ' + $Branch + ' apagada; ' + $ponta.Substring(0, 7) + ' mesclada')

if ($Enviar) {
    G push origin $develop | Out-Null
    if ($tipo -in @('release', 'hotfix')) { G push origin $main ('refs/tags/' + $tag) | Out-Null }
    & git ls-remote --exit-code --heads origin $Branch > $null 2>&1
    if ($LASTEXITCODE -eq 0) { G push origin --delete $Branch | Out-Null }
    Write-Host '  enviado ao GitHub'
}
G checkout -q $develop | Out-Null
Write-Host 'Pronto.'
