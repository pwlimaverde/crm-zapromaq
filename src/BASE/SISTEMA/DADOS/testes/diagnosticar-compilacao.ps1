<#
  diagnosticar-compilacao.ps1 - diz ONDE a compilação do VBA falha.

  O build detecta que "a compilação não terminou", mas a mensagem do VBA aparece
  numa caixa de diálogo do Excel oculto. Este script abre o .xlsm montado,
  manda compilar e, em paralelo, um vigia lê o texto da caixa de erro e a fecha;
  depois pergunta ao VBE qual módulo e linha ficaram selecionados.
  Repete até compilar ou até -Maximo erros (corrige-se um por vez).

  Uso: diagnosticar-compilacao.ps1 [-Arquivo <xlsm>]   (padrão: execucao\build\CRM_Zapromaq.xlsm)
#>
param([string]$Arquivo = '')
$ErrorActionPreference = 'Stop'
$dados = Split-Path -Parent $PSScriptRoot
. (Join-Path $dados 'lib\Comum.ps1')
. (Join-Path $dados 'build\lib\Excel.ps1')
if (-not $Arquivo) { $Arquivo = Join-Path $dados 'execucao\build\CRM_Zapromaq.xlsm' }
if (-not (Test-Path -LiteralPath $Arquivo)) { throw ('não existe: ' + $Arquivo + ' (rode o build antes)') }

$saidaVigia = Join-Path $env:TEMP ('crm-compila-' + $PID + '.txt')
$vbomOriginal = Get-AcessoVBOM
Set-AcessoVBOM 1
$xl = $null; $idExcel = 0
try {
    $xl = New-Excel ([ref]$idExcel)
    $xl.DisplayAlerts = $false; $xl.EnableEvents = $false
    $wb = $xl.Workbooks.Open($Arquivo)
    Remove-Item -LiteralPath $saidaVigia -ErrorAction SilentlyContinue

    # vigia: acha a caixa de diálogo (#32770) do processo do Excel, grava os textos e clica OK
    $vigia = @'
Add-Type @"
using System; using System.Text; using System.Runtime.InteropServices; using System.Collections.Generic;
public static class W {
  public delegate bool P(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(P f, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr h, P f, IntPtr l);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint p);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
  public static string Classe(IntPtr h){ var s=new StringBuilder(256); GetClassName(h,s,256); return s.ToString(); }
  public static string Texto(IntPtr h){ var s=new StringBuilder(2048); GetWindowText(h,s,2048); return s.ToString(); }
}
"@
$alvo = [uint32]__PID__
$fim = (Get-Date).AddSeconds(90)
while ((Get-Date) -lt $fim) {
  $achou = $null
  [W]::EnumWindows({ param($h,$l) $p=[uint32]0; [void][W]::GetWindowThreadProcessId($h,[ref]$p)
      if ($p -eq $alvo -and [W]::Classe($h) -eq '#32770' -and [W]::IsWindowVisible($h)) { $script:achou = $h; return $false }; return $true }, [IntPtr]::Zero) | Out-Null
  if ($script:achou) {
    $textos = New-Object System.Collections.ArrayList; $botoes = New-Object System.Collections.ArrayList
    [void][W]::EnumChildWindows($script:achou, { param($h,$l) $c=[W]::Classe($h); $t=[W]::Texto($h)
        if ($c -eq 'Button') { [void]$botoes.Add($h) } elseif ($t) { [void]$textos.Add($t) }; return $true }, [IntPtr]::Zero)
    Set-Content -LiteralPath '__SAIDA__' -Value (([W]::Texto($script:achou)) + ' | ' + ($textos -join ' | ')) -Encoding UTF8
    if ($botoes.Count -gt 0) { [void][W]::SendMessage($botoes[0], 0x00F5, [IntPtr]::Zero, [IntPtr]::Zero) }  # BM_CLICK
    exit 0
  }
  Start-Sleep -Milliseconds 300
}
'@
    $vigia = $vigia.Replace('__PID__', [string]$idExcel).Replace('__SAIDA__', $saidaVigia)
    $arqVigia = Join-Path $env:TEMP ('crm-vigia-' + $PID + '.ps1')
    [IO.File]::WriteAllText($arqVigia, $vigia, (New-Object Text.UTF8Encoding($true)))
    $ps = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $proc = Start-Process $ps -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $arqVigia) -PassThru -WindowStyle Hidden

    $cmd = $xl.VBE.CommandBars.FindControl(1, 578)                 # Depurar > Compilar VBAProject
    $cmd.Execute()
    [void]$proc.WaitForExit(5000); if (-not $proc.HasExited) { $proc.Kill() }

    if ($xl.VBE.CommandBars.FindControl(1, 578).Enabled -eq $false) {
        Write-Host 'COMPILOU SEM ERRO.'
    } else {
        $msg = if (Test-Path -LiteralPath $saidaVigia) { Get-Content -LiteralPath $saidaVigia -Raw -Encoding UTF8 } else { '(caixa de erro não capturada)' }
        $pane = $xl.VBE.ActiveCodePane
        $l1 = 0; $c1 = 0; $l2 = 0; $c2 = 0
        $modulo = '?'; $linha = ''
        if ($pane) {
            $pane.GetSelection([ref]$l1, [ref]$c1, [ref]$l2, [ref]$c2)
            $modulo = $pane.CodeModule.Parent.Name
            if ($l1 -gt 0) { $linha = $pane.CodeModule.Lines($l1, 1) }
        }
        Write-Host ('ERRO DE COMPILAÇÃO: ' + $msg.Trim())
        Write-Host ('  em ' + $modulo + ', linha ' + $l1 + ' (colunas ' + $c1 + '-' + $c2 + '):')
        Write-Host ('  ' + $linha.Trim())
        exit 1
    }
} finally {
    if ($xl) { try { $wb.Close($false) } catch { } }
    Close-Excel $xl $idExcel
    Set-AcessoVBOM $vbomOriginal
    Remove-Item -LiteralPath $saidaVigia -ErrorAction SilentlyContinue
}
