<#
  aplicar-migracoes.ps1 - aplica, em ordem, as migrações de banco\migracoes ainda não aplicadas.

  Cada migração é um arquivo NNN-descricao.ps1 que devolve um hashtable:
      @{ Id = '001'; Descricao = '...'
         EsquemaDe = $null ou '1.0'   # exige esta versao_esquema antes (migração estrutural)
         EsquemaPara = $null ou '1.1' # grava esta versao_esquema depois
         Aplicar = { param($cn, $tx) ... } }
  - Migração só de DADOS (EsquemaPara vazio) roda numa transação e pode rodar com gente usando.
  - Migração ESTRUTURAL (EsquemaPara preenchido) exige banco exclusivo (sem .laccdb): DDL no ACE
    não é transacional. Mudar versao_esquema obriga a publicar front e IA na mesma versão.
  - Antes de qualquer migração é feito um backup (execucao\backup). Sem backup, não migra.
  - O que foi aplicado fica em config: chave 'migracao.NNN' = data e usuário.

  Uso: aplicar-migracoes.ps1 [-Banco <accdb>] [-Simular]
#>
param([string]$Banco = '', [switch]$Simular)
$ErrorActionPreference = 'Stop'
. (Join-Path (Split-Path -Parent $PSScriptRoot) 'lib\Comum.ps1')

$raiz = Get-RaizBase $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($Banco)) { $Banco = Get-CaminhoBanco $raiz }
$log = New-Log $raiz 'migracoes'
$log.L('=== MIGRAÇÕES DO BANCO - CRM Zapromaq ===')
$log.L('Banco...: ' + $Banco)

$pasta = Join-Path $PSScriptRoot 'migracoes'
$arquivos = @(Get-ChildItem -LiteralPath $pasta -Filter '*.ps1' -File | Where-Object { $_.Name -match '^\d{3}-' } | Sort-Object Name)
$cn = Open-Banco $Banco
try {
    $aplicadas = @{}
    $leitor = (New-ComandoBanco $cn "SELECT chave, valor FROM config WHERE chave LIKE 'migracao.%'" @() $null).ExecuteReader()
    while ($leitor.Read()) { $aplicadas[[string]$leitor.GetValue(0)] = [string]$leitor.GetValue(1) }
    $leitor.Close()
    $esquema = Get-VersaoEsquemaBanco $cn
    $log.L('Esquema.: ' + $esquema + '   migrações já aplicadas: ' + $aplicadas.Count)
    $log.L('')

    $pendentes = @()
    foreach ($a in $arquivos) {
        $m = & $a.FullName
        if (-not $m.Id -or -not $m.Aplicar) { throw ($a.Name + ': migração sem Id ou sem Aplicar') }
        if ($aplicadas.ContainsKey('migracao.' + $m.Id)) { $log.L('  já aplicada  ' + $a.Name); continue }
        $m['Arquivo'] = $a.Name
        $pendentes += , $m
        $log.L('  PENDENTE     ' + $a.Name + '  -  ' + $m.Descricao)
    }
    if ($pendentes.Count -eq 0) { $log.L(''); $log.L('Nada a aplicar.'); $log.Salvar(); exit 0 }
    if ($Simular) { $log.L(''); $log.L('Simulação: nada foi alterado.'); $log.Salvar(); exit 0 }
} finally { $cn.Close(); $cn.Dispose() }

# ---- backup antes de mexer (e exclusividade, se houver migração estrutural)
$estrutural = @($pendentes | Where-Object { $_.EsquemaPara }).Count -gt 0
$ocupado = Test-Path -LiteralPath ([System.IO.Path]::ChangeExtension($Banco, 'laccdb'))
if ($estrutural -and $ocupado) {
    $log.L(''); $log.L('ABORTADO: há migração estrutural e o banco está em uso (.laccdb). Peça para todos fecharem o CRM.')
    $log.Salvar(); exit 2
}
if ($Banco -eq (Get-CaminhoBanco $raiz)) {
    & (Join-Path (Get-PastaDados $raiz) 'operacao\scripts\backup-banco.ps1') -Modo $(if ($ocupado) { 'Periodico' } else { 'Diario' })
    if ($LASTEXITCODE -ne 0) { $log.L('ABORTADO: o backup falhou; nenhuma migração aplicada.'); $log.Salvar(); exit 1 }
}

$cn = Open-Banco $Banco
try {
    foreach ($m in $pendentes) {
        $atual = Get-VersaoEsquemaBanco $cn
        if ($m.EsquemaDe -and $atual -ne $m.EsquemaDe) {
            throw ($m.Arquivo + ': exige esquema ' + $m.EsquemaDe + ', o banco está em ' + $atual)
        }
        $log.L('')
        $log.L('--- ' + $m.Arquivo)
        if ($m.EsquemaPara) {
            & $m.Aplicar $cn $null $log
            Set-ConfigBanco $cn ('migracao.' + $m.Id) ((Get-Date -Format 'yyyy-MM-dd HH:mm') + ' ' + $env:USERNAME) $null
            Set-ConfigBanco $cn 'versao_esquema' $m.EsquemaPara $null
            $log.L('  esquema -> ' + $m.EsquemaPara)
        } else {
            $tx = $cn.BeginTransaction()
            try {
                & $m.Aplicar $cn $tx $log
                Set-ConfigBanco $cn ('migracao.' + $m.Id) ((Get-Date -Format 'yyyy-MM-dd HH:mm') + ' ' + $env:USERNAME) $tx
                $tx.Commit()
            } catch { $tx.Rollback(); throw }
        }
        $log.L('  OK')
    }
} catch {
    $log.L(''); $log.L('ERRO: ' + $_.Exception.Message); $log.L('A migração com erro foi desfeita (dados) ou parou no meio (estrutura): restaure o backup se preciso.')
    $log.Salvar(); exit 1
} finally { $cn.Close(); $cn.Dispose() }

$log.L('')
$log.L('Migrações aplicadas: ' + $pendentes.Count)
$log.Salvar()
exit 0
