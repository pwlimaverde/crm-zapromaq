<#
  backup-banco.ps1 - cópia de segurança do banco do CRM.

  Modos:
    -Modo Diario     cópia COMPLETA: exige exclusividade (ninguém gravando). Sai com 2 se ocupado.
    -Modo Periodico  cópia simples, roda com gente usando (pode sair inconsistente; não bloqueia quem grava).
    -Modo Agendado   o das tarefas agendadas (9h, 12h30, 16h): tenta o completo até -Tentativas vezes,
                     esperando -IntervaloMin entre elas (no horário, +10 e +20 min); se não conseguir a
                     exclusividade, faz a cópia simples marcada como NÃO verificada. Nunca termina sem
                     cópia, salvo falha de gravação.

  Exclusividade (DEV-03, ciclo 007): a cópia completa segura o arquivo do banco (FileShare.Read) durante TODA
  a captura - cópia, hash e contagens do manifesto - e copia pelo próprio fluxo. Quem abrir o CRM nesse
  intervalo lê normalmente; quem tentar gravar recebe a mensagem controlada "cópia de segurança em andamento"
  (marca execucao\backup\copiando.txt). Ver lib\Backup.ps1.

  Janela (ciclo 007): disparos do modo Agendado para a mesma janela (ex.: duas estações com a tarefa das 9h)
  geram UMA captura: quem encontra a trava ocupada espera (até -EsperaTravaSeg) e, se a janela já tem cópia
  (manifesto com "janela"), registra "janela atendida" e sai com 0.

  Nome do arquivo: crm_<diario|periodico>_AAAAMMDD-HHMMSS_<ESTACAO>.accdb
    diario    = cópia completa, feita com exclusividade
    periodico = cópia simples, feita com o CRM em uso
  Se o nome já existir (duas cópias no mesmo segundo), ganha -2, -3...

  Coordenação (spec 006):
    - trava com dono (backup.lock): duas execuções ao mesmo tempo não se atropelam (sai com 4);
      trava abandonada é recuperada (ver lib\Backup.ps1);
    - a cópia nasce como .tmp e só vira .accdb depois de conferida (tamanho e hash);
    - manifesto em TODA cópia (<arquivo>.manifesto.json) e validação por restauração numa pasta
      descartável: passando, grava <arquivo>.validado. Cópia simples feita com o CRM aberto pode não
      validar (inconsistente): fica sem a marca e o log avisa;
    - um log por execução em execucao\backup\logs e o consolidado backup.log.

  Retenção (todos os modos): mês corrente e o anterior (-Meses 2).
    A data vem do NOME do arquivo, não da data de modificação. O expurgo só roda depois de uma
    cópia bem-sucedida, nunca deixa a pasta sem cópia, NUNCA apaga o último backup validado (se
    ele é o único validado, fica, mesmo velho) e leva junto o manifesto e a marca de quem apaga.
    Não mexe em subpastas (antes-da-publicacao, logs) nem em arquivo fora do padrão.

  Destino: SISTEMA\DADOS\execucao\backup. Código de saída: 0 ok, 1 falha, 2 ocupado (modo diário),
  4 outra execução de backup em andamento.
#>
param(
    [ValidateSet('Periodico', 'Diario', 'Agendado')][string]$Modo = 'Periodico',
    [ValidateRange(1, 12)][int]$Tentativas = 3,
    [ValidateRange(1, 60)][int]$IntervaloMin = 10,
    [ValidateRange(1, 24)][int]$Meses = 2,
    [ValidateRange(1, 600)][int]$EsperaExclusivaSeg = 60,   # quanto cada tentativa insiste na exclusividade
    [ValidateRange(0, 3600)][int]$EsperaTravaSeg = 1500,    # Agendado: espera a outra execução da mesma janela
    [string]$Janela = '',      # Agendado: janela AAAAMMDD-HHMM (padrão: a do horário atual; os testes informam)
    [string]$PastaTeste = ''   # só para teste do expurgo: -PastaTeste <pasta> roda apenas o expurgo nela
)
$ErrorActionPreference = 'Stop'
$dadosLib = Join-Path (Split-Path -Parent $PSScriptRoot) '..\lib'
. (Join-Path $dadosLib 'Comum.ps1')
. (Join-Path $dadosLib 'AtualizacaoBanco.ps1')   # Backup.ps1 vem junto

$script:linhasExecucao = New-Object System.Collections.ArrayList
function Anotar([string]$t) {
    $linha = (Get-Date -Format 'dd/MM/yyyy HH:mm:ss') + '  [' + $Modo + ']  [' + $env:COMPUTERNAME + ']  ' + $t
    Write-Host $linha
    [void]$script:linhasExecucao.Add($linha)
    # o backup.log é compartilhado por todas as execuções (várias estações na mesma janela): escrita com
    # compartilhamento e nova tentativa; se não der, a linha continua no log desta execução (logs\backup-*.log)
    # (nomes próprios: este script também roda dentro de outros - aplicar-migracoes - e o PowerShell 5.1 não
    # compila "$x++" quando um $x de mesmo nome e outro tipo existe no chamador)
    $bytesLinhaLog = (New-Object System.Text.UTF8Encoding($false)).GetBytes($linha + "`r`n")
    foreach ($esperaLogMs in @(0, 50, 100, 150, 200, 300, 400, 500, 700, 1000)) {
        if ($esperaLogMs) { Start-Sleep -Milliseconds $esperaLogMs }
        try {
            $fluxoLog = New-Object System.IO.FileStream($log, [System.IO.FileMode]::Append, [System.IO.FileAccess]::Write, [System.IO.FileShare]::ReadWrite)
            try { $fluxoLog.Write($bytesLinhaLog, 0, $bytesLinhaLog.Length) } finally { $fluxoLog.Dispose() }
            return
        } catch { }
    }
}

function Remove-Copia([string]$arquivo) {
    foreach ($ext in @('', '.manifesto.json', '.validado')) {
        $p = $arquivo + $ext
        if (Test-Path -LiteralPath $p) { Remove-Item -LiteralPath $p -Force }
    }
}

function Invoke-Expurgo([string]$pasta) {
    $corte = (Get-Date -Day 1).Date.AddMonths(-($Meses - 1))
    $amanha = (Get-Date).Date.AddDays(1)
    $rx = '^crm_(diario|periodico)_(\d{8})-\d{4,6}(-\d+)?_.+\.accdb$'
    $copias = @(Get-ChildItem -LiteralPath $pasta -Filter 'crm_*.accdb' -File | ForEach-Object {
        if ($_.Name -match $rx) {
            $d = [datetime]::MinValue
            if ([datetime]::TryParseExact($Matches[2], 'yyyyMMdd', [Globalization.CultureInfo]::InvariantCulture,
                                          [Globalization.DateTimeStyles]::None, [ref]$d)) {
                New-Object PSObject -Property @{ Arquivo = $_; Data = $d; Validado = (Test-Path -LiteralPath ($_.FullName + '.validado')) }
            }
        }
    })
    $futuras = @($copias | Where-Object { $_.Data -ge $amanha })
    if ($futuras.Count -gt 0) {
        # relógio da estação errado: apagar agora poderia levar cópias boas
        Anotar ('expurgo NÃO feito: ' + $futuras.Count + ' cópia(s) com data futura no nome. Confira o relógio da estação.')
        return
    }
    $velhos = @($copias | Where-Object { $_.Data -lt $corte })
    if ($velhos.Count -eq 0) { return }
    if (($copias.Count - $velhos.Count) -lt 1) {
        Anotar 'expurgo não feito: removeria a última cópia'
        return
    }
    # nunca apagar o último backup validado: se todos os validados estão entre os velhos, o mais novo deles fica
    $validadosFora = @($copias | Where-Object { $_.Validado -and $_.Data -ge $corte })
    if ($validadosFora.Count -eq 0) {
        $poupado = @($velhos | Where-Object { $_.Validado } | Sort-Object { Get-InstanteDoBackup $_.Arquivo } -Descending | Select-Object -First 1)
        if ($poupado.Count -eq 1) {
            Anotar ('expurgo: ' + $poupado[0].Arquivo.Name + ' fica, mesmo velho: é o último backup validado')
            $velhos = @($velhos | Where-Object { $_.Arquivo.FullName -ne $poupado[0].Arquivo.FullName })
        }
    }
    if ($velhos.Count -eq 0) { return }
    $falhas = 0
    foreach ($v in $velhos) {
        try { Remove-Copia $v.Arquivo.FullName }
        catch { $falhas++; Anotar ('expurgo: não removeu ' + $v.Arquivo.Name + ' -> ' + $_.Exception.Message.Split([char]13)[0]) }
    }
    Anotar ('expurgo: ' + ($velhos.Count - $falhas) + ' cópia(s) anterior(es) a ' + $corte.ToString('dd/MM/yyyy') + ' removida(s)')
}

# ---- modo de teste do expurgo: não toca no banco
if ($PastaTeste) {
    $log = Join-Path $PastaTeste 'backup.log'
    Invoke-Expurgo $PastaTeste
    exit 0
}

$raiz = Get-RaizBase $PSScriptRoot
$banco = Get-CaminhoBanco $raiz
$pastaBkp = Get-PastaExecucao $raiz 'backup'
$pastaLogs = Join-Path $pastaBkp 'logs'
New-Item -ItemType Directory -Path $pastaLogs -Force | Out-Null
$log = Join-Path $pastaBkp 'backup.log'

if (-not (Test-Path -LiteralPath $banco)) {
    # tarefa agendada que falha em silêncio é pior que tarefa que não existe
    Anotar ('FALHOU: banco não encontrado -> ' + $banco)
    exit 1
}

# ---- janela (só o modo Agendado tem janela: é ela que a tarefa das 9h, 12h30 ou 16h cumpre)
if ($Modo -eq 'Agendado' -and -not $Janela) { $Janela = Get-JanelaBackup (Get-Date) }

# ---- trava de backup (uma execução por vez nesta pasta)
$travaBkp = Enter-TravaBackup $pastaBkp $Modo
if (-not $travaBkp.ok -and $Modo -eq 'Agendado' -and $Janela) {
    # outra execução (outra estação, mesma janela): espera ela terminar e confere se a janela ficou atendida
    Anotar ('janela ' + $Janela + ': ' + $travaBkp.motivo + ' Aguardando até ' + $EsperaTravaSeg + ' s.')
    $limiteTrava = (Get-Date).AddSeconds($EsperaTravaSeg)
    while (-not $travaBkp.ok -and (Get-Date) -lt $limiteTrava) {
        Start-Sleep -Seconds 2
        $travaBkp = Enter-TravaBackup $pastaBkp $Modo
    }
}
if (-not $travaBkp.ok) { Anotar ('ABORTADO: ' + $travaBkp.motivo); exit 4 }
if ($travaBkp.recuperada) { Anotar 'AVISO: havia uma trava de backup abandonada; foi recuperada.' }

$codigo = 0
$copiaExcl = $null
try {
    if ($Janela) {
        $atendida = Get-CopiaDaJanela $pastaBkp $Janela
        if ($atendida) {
            Anotar ('janela ' + $Janela + ' atendida: ' + $atendida + '. Nada a copiar.')
            exit 0
        }
    }
    # ---- decide o tipo de cópia: completa só com a exclusividade em mãos
    if ($Modo -eq 'Periodico') { $tipo = 'periodico' }
    else {
        $copiaExcl = Enter-CopiaExclusiva $banco $pastaBkp $EsperaExclusivaSeg
        if ($Modo -eq 'Agendado') {
            for ($i = 1; -not $copiaExcl.ok -and $i -lt $Tentativas; $i++) {
                Anotar ('sem exclusividade (tentativa ' + $i + ' de ' + $Tentativas + ', ' + $copiaExcl.motivo + '): nova tentativa em ' + $IntervaloMin + ' min')
                Start-Sleep -Seconds ($IntervaloMin * 60)
                $copiaExcl = Enter-CopiaExclusiva $banco $pastaBkp $EsperaExclusivaSeg
            }
        } elseif (-not $copiaExcl.ok) {
            Anotar ('ABORTADO: ' + $copiaExcl.motivo + '. O backup diário exige exclusividade.')
            $codigo = 2
            exit 2
        }
        $tipo = if ($copiaExcl.ok) { 'diario' } else { 'periodico' }
        if (-not $copiaExcl.ok) { Anotar ('AVISO: sem exclusividade depois de ' + $Tentativas + ' tentativa(s). Cópia simples, marcada como NÃO verificada.') }
    }

    # ---- cópia: .tmp, conferida, promovida
    $base = 'crm_' + $tipo + '_' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '_' + ($env:COMPUTERNAME -replace '[^A-Za-z0-9_-]', '_')
    $nome = $base + '.accdb'; $n = 1
    while (Test-Path -LiteralPath (Join-Path $pastaBkp $nome)) { $n++; $nome = $base.Replace('_' + $env:COMPUTERNAME, '-' + $n + '_' + $env:COMPUTERNAME) + '.accdb' }
    $destino = Join-Path $pastaBkp $nome
    $tmp = $destino + '.tmp'
    try {
        if ($tipo -eq 'diario') { Copy-FluxoParaArquivo $copiaExcl.fluxo $tmp; $tamOrigem = $copiaExcl.fluxo.Length }
        else { Copy-BancoSemBloquear $banco $tmp; $tamOrigem = (Get-Item -LiteralPath $banco).Length }
        $tamCopia = (Get-Item -LiteralPath $tmp).Length
        if ($tamCopia -ne $tamOrigem) {
            Anotar ('FALHA: cópia com ' + $tamCopia + ' bytes contra ' + $tamOrigem + ' do original. Removida.')
            Remove-Item -LiteralPath $tmp -Force
            $codigo = 1
            exit 1
        }
        Move-Item -LiteralPath $tmp -Destination $destino
        Anotar ('OK: ' + (Split-Path -Leaf $destino) + '  (' + [math]::Round($tamCopia / 1MB, 2) + ' MB)' + $(if ($tipo -eq 'diario') { ', com exclusividade' } else { '' }))
    } catch {
        Anotar ('FALHA ao copiar: ' + $_.Exception.Message.Split([char]13)[0])
        Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
        $codigo = 1
        exit 1
    }

    # ---- manifesto e validação (restauração numa pasta descartável; o arquivo de backup não é aberto)
    # Na cópia completa o manifesto (contagens da origem) sai AINDA com o arquivo seguro: mesmo corte da cópia.
    try {
        $extra = [ordered]@{ consistencia = $(if ($tipo -eq 'diario') { 'exclusiva' } else { 'nao-verificada' }) }
        if ($Janela) { $extra['janela'] = $Janela }
        [void](New-ManifestoBackup $destino $banco $tipo $extra)
        Exit-CopiaExclusiva $copiaExcl; $copiaExcl = $null
        $val = Test-BackupRestauravel $destino
        if ($val.ok) { Set-BackupValidado $destino; Anotar 'validado: manifesto, restauração numa pasta descartável e contagens conferem' }
        else { Anotar ('AVISO: a cópia NÃO foi validada (' + ($val.problemas -join '; ') + '). Fica sem a marca de validado.') }
    } catch {
        Anotar ('AVISO: não foi possível validar a cópia (' + $_.Exception.Message.Split([char]13)[0] + '). Fica sem a marca de validado.')
    }

    # ---- expurgo: só depois de cópia bem-sucedida
    Invoke-Expurgo $pastaBkp

    $total = @(Get-ChildItem -LiteralPath $pastaBkp -Filter '*.accdb' -File)
    Anotar ('na pasta: ' + $total.Count + ' cópia(s), ' + [math]::Round((($total | Measure-Object Length -Sum).Sum) / 1MB, 1) + ' MB')
} finally {
    Exit-CopiaExclusiva $copiaExcl
    Exit-TravaBackup $travaBkp
    try {
        $arqExec = Join-Path $pastaLogs ('backup-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '-' + $PID + '-' + ($env:COMPUTERNAME -replace '[^A-Za-z0-9_-]', '_') + '.log')
        [System.IO.File]::WriteAllLines($arqExec, [string[]]$script:linhasExecucao.ToArray(), (New-Object System.Text.UTF8Encoding($true)))
    } catch { }
}
exit $codigo
