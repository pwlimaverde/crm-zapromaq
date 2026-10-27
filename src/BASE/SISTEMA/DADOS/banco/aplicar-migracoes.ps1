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
  - O que foi aplicado fica em config: chave 'migracao.NNN' = data e usuário, e 'sha256.migracao.NNN' = checksum
    do arquivo (ciclo 018, DEV-04): migração aplicada que mudou depois é recusada (saída 4).

  Migração estrutural só pela montagem: no banco de produção ela é recusada (exit 3). A montagem
  aplica numa cópia de trabalho (-CopiaDeTrabalho, em execucao\publicacao ou execucao\ensaio) e só
  depois troca o banco. Em banco com config.ambiente = TESTE, aplica no lugar.

  Manutenção (DEV-08): aplicando no lugar (sem -CopiaDeTrabalho), liga config.manutencao antes e desliga no
  fim; front e conector recusam gravação enquanto isso. Migração estrutural que parou no meio deixa o sinal
  ligado (falha segura): restaure o backup e rode -LiberarManutencao.

  Uso: aplicar-migracoes.ps1 [-Banco <accdb>] [-Simular] [-CopiaDeTrabalho] [-LiberarManutencao]
#>
param([string]$Banco = '', [switch]$Simular, [switch]$CopiaDeTrabalho, [switch]$LiberarManutencao)
$ErrorActionPreference = 'Stop'
. (Join-Path (Split-Path -Parent $PSScriptRoot) 'lib\Comum.ps1')

$raiz = Get-RaizBase $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($Banco)) { $Banco = Get-CaminhoBanco $raiz }
$log = New-Log $raiz 'migracoes'
$log.L('=== MIGRAÇÕES DO BANCO - CRM Zapromaq ===')
$log.L('Banco...: ' + $Banco)

if ($LiberarManutencao) {
    $cn = Open-Banco $Banco
    try { $antes = Get-Manutencao $cn; [void](Clear-Manutencao $cn) } finally { $cn.Close(); $cn.Dispose() }
    $log.L($(if ($antes) { 'Manutenção desligada (estava: ' + $antes + ').' } else { 'Não havia manutenção ligada.' }))
    $log.Salvar(); exit 0
}

# Migração ESTRUTURAL (DDL no ACE não é transacional) só é aplicada:
#   - numa cópia de trabalho da montagem (-CopiaDeTrabalho), que tem de estar em execucao\publicacao
#     ou execucao\ensaio e não pode ser o banco de produção da raiz; ou
#   - em banco com config.ambiente = TESTE.
# Qualquer outro caso (banco de produção no lugar) é recusado ANTES de aplicar qualquer migração,
# inclusive as de dados: a atualização do banco real acontece só pelo MONTAR-FRONTEND.bat.
if ($CopiaDeTrabalho) {
    $cheio = [System.IO.Path]::GetFullPath($Banco)
    $pastaExec = [System.IO.Path]::GetFullPath((Join-Path (Get-PastaDados $raiz) 'execucao')) + '\'
    $dentro = $false
    foreach ($sub in @('publicacao', 'ensaio')) {
        if ($cheio.StartsWith($pastaExec + $sub + '\', [StringComparison]::OrdinalIgnoreCase)) { $dentro = $true }
    }
    $ehProducao = (Test-MesmoCaminho $cheio (Get-CaminhoBanco $raiz))   # UNC, letra e relativo contam como o mesmo arquivo
    if (-not $dentro -or $ehProducao) {
        $log.L(''); $log.L('RECUSADO: -CopiaDeTrabalho exige um banco dentro de execucao\publicacao ou execucao\ensaio e diferente do banco de produção. Nada foi alterado.')
        $log.Salvar(); exit 3
    }
}

$pasta = Join-Path $PSScriptRoot 'migracoes'
$arquivos = @(Get-ChildItem -LiteralPath $pasta -Filter '*.ps1' -File | Where-Object { $_.Name -match '^\d{3}-' } | Sort-Object Name)
# DEV-04 (ciclo 018): o checksum das aplicadas antes do checksum existir só é registrado onde a montagem pode
# escrever (cópia de trabalho ou banco de TESTE); na produção, fora da montagem, só se informa
$podeRegistrar = $CopiaDeTrabalho -or ((Get-AmbienteDoBanco $Banco) -eq 'TESTE')
$porId = @{}
$cn = Open-Banco $Banco
try {
    $aplicadas = @{}
    $leitor = (New-ComandoBanco $cn "SELECT chave, valor FROM config WHERE chave LIKE 'migracao.%' OR chave LIKE 'sha256.migracao.%'" @() $null).ExecuteReader()
    while ($leitor.Read()) { $aplicadas[[string]$leitor.GetValue(0)] = [string]$leitor.GetValue(1) }
    $leitor.Close()
    $esquema = Get-VersaoEsquemaBanco $cn
    $log.L('Esquema.: ' + $esquema + '   migrações já aplicadas: ' + $aplicadas.Count)
    $log.L('')

    $pendentes = @()
    foreach ($a in $arquivos) {
        $m = & $a.FullName
        if (-not $m.Id -or -not $m.Aplicar) { throw ($a.Name + ': migração sem Id ou sem Aplicar') }
        $porId[[string]$m.Id] = $a.FullName
        if ($aplicadas.ContainsKey('migracao.' + $m.Id)) { $log.L('  já aplicada  ' + $a.Name); continue }
        $m['Arquivo'] = $a.Name
        $pendentes += , $m
        $log.L('  PENDENTE     ' + $a.Name + '  -  ' + $m.Descricao)
    }
    # DEV-04: migração já aplicada que mudou depois é recusada antes de qualquer alteração
    $conf = Test-ChecksumsMigracoes $aplicadas $porId
    if (@($conf.divergentes).Count -gt 0) {
        $log.L(''); $log.L('RECUSADO: migração já aplicada mudou depois de aplicada (checksum diferente): ' + (@($conf.divergentes) -join ', ') + '.')
        $log.L('Volte o arquivo da migração ao conteúdo aplicado (repositório) - correção vai numa migração NOVA. Nada foi alterado.')
        $log.Salvar(); exit 4
    }
    $semRegistro = @($conf.sem_registro)
    if ($semRegistro.Count -gt 0) {
        $log.L('  checksum ainda não registrado (aplicadas antes do ciclo 018): ' + ($semRegistro -join ', ') +
               $(if ($podeRegistrar -and -not $Simular) { ' - registrando agora' } else { ' - registrado pela montagem' }))
    }
    if ($pendentes.Count -eq 0) {
        if ($semRegistro.Count -gt 0 -and $podeRegistrar -and -not $Simular) {
            foreach ($id in $semRegistro) { Set-ConfigBanco $cn (Get-ChaveChecksumMigracao $id) (Get-ChecksumMigracao $porId[$id]) $null }
        }
        $log.L(''); $log.L('Nada a aplicar.'); $log.Salvar(); exit 0
    }
    if ($Simular) { $log.L(''); $log.L('Simulação: nada foi alterado.'); $log.Salvar(); exit 0 }
} finally { $cn.Close(); $cn.Dispose() }

# ---- backup antes de mexer (e exclusividade, se houver migração estrutural)
$estrutural = @($pendentes | Where-Object { $_.EsquemaPara }).Count -gt 0
if ($estrutural -and -not $CopiaDeTrabalho) {
    $amb = Get-AmbienteDoBanco $Banco
    if ($amb -ne 'TESTE') {
        $log.L(''); $log.L('RECUSADO: há migração estrutural pendente e este banco não é uma cópia de trabalho da montagem nem está com config.ambiente = TESTE (ambiente: ' + $(if ($amb) { $amb } else { 'ilegível' }) + ').')
        $log.L('O banco de produção só é atualizado pelo MONTAR-FRONTEND.bat (cópia, conferência e troca). Nada foi alterado.')
        $log.Salvar(); exit 3
    }
}
$ocupado = Test-BancoEmUso $Banco
if ($estrutural -and $ocupado) {
    $log.L(''); $log.L('ABORTADO: há migração estrutural e o banco está em uso (.laccdb). Peça para todos fecharem o CRM.')
    $log.Salvar(); exit 2
}
if ($Banco -eq (Get-CaminhoBanco $raiz)) {
    & (Join-Path (Get-PastaDados $raiz) 'operacao\scripts\backup-banco.ps1') -Modo $(if ($ocupado) { 'Periodico' } else { 'Diario' })
    if ($LASTEXITCODE -ne 0) { $log.L('ABORTADO: o backup falhou; nenhuma migração aplicada.'); $log.Salvar(); exit 1 }
}

$cn = Open-Banco $Banco
$sinal = ''
$deixarSinal = $false
try {
    if (-not $CopiaDeTrabalho) { $sinal = Set-Manutencao $cn ('migrações ' + (@($pendentes | ForEach-Object { $_.Id }) -join ',')); $log.L('manutenção ligada: ' + $sinal) }
    if ($podeRegistrar) { foreach ($id in $semRegistro) { Set-ConfigBanco $cn (Get-ChaveChecksumMigracao $id) (Get-ChecksumMigracao $porId[$id]) $null } }
    foreach ($m in $pendentes) {
        $atual = Get-VersaoEsquemaBanco $cn
        if ($m.EsquemaDe -and $atual -ne $m.EsquemaDe) {
            throw ($m.Arquivo + ': exige esquema ' + $m.EsquemaDe + ', o banco está em ' + $atual)
        }
        $log.L('')
        $log.L('--- ' + $m.Arquivo)
        if ($m.EsquemaPara) {
            & $m.Aplicar $cn $null $log $pasta
            Set-ConfigBanco $cn ('migracao.' + $m.Id) ((Get-Date -Format 'yyyy-MM-dd HH:mm') + ' ' + $env:USERNAME) $null
            Set-ConfigBanco $cn (Get-ChaveChecksumMigracao $m.Id) (Get-ChecksumMigracao (Join-Path $pasta $m.Arquivo)) $null
            Set-ConfigBanco $cn 'versao_esquema' $m.EsquemaPara $null
            $log.L('  esquema -> ' + $m.EsquemaPara)
        } else {
            $tx = $cn.BeginTransaction()
            try {
                & $m.Aplicar $cn $tx $log $pasta
                Set-ConfigBanco $cn ('migracao.' + $m.Id) ((Get-Date -Format 'yyyy-MM-dd HH:mm') + ' ' + $env:USERNAME) $tx
                Set-ConfigBanco $cn (Get-ChaveChecksumMigracao $m.Id) (Get-ChecksumMigracao (Join-Path $pasta $m.Arquivo)) $tx
                $tx.Commit()
            } catch { $tx.Rollback(); throw }
        }
        $log.L('  OK')
    }
} catch {
    $log.L(''); $log.L('ERRO: ' + $_.Exception.Message); $log.L('A migração com erro foi desfeita (dados) ou parou no meio (estrutura): restaure o backup se preciso.')
    # estrutura pela metade: o sinal fica ligado (ninguém grava num esquema misto) até alguém restaurar e liberar
    if ($sinal -and @($pendentes | Where-Object { $_.EsquemaPara }).Count -gt 0) {
        $deixarSinal = $true
        $log.L('A manutenção CONTINUA LIGADA. Depois de restaurar o backup: aplicar-migracoes.ps1 -LiberarManutencao')
    }
    $log.Salvar(); exit 1
} finally {
    if ($sinal -and -not $deixarSinal) { try { [void](Clear-Manutencao $cn $sinal); $log.L('manutenção desligada') } catch { $log.L('AVISO: não desligou a manutenção: ' + $_.Exception.Message) } }
    $cn.Close(); $cn.Dispose()
}

$log.L('')
$log.L('Migrações aplicadas: ' + $pendentes.Count)
$log.Salvar()
exit 0
