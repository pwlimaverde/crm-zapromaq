<#
  capturar-tela.ps1 - mostra um formulário do .xlsm montado e salva um print dele.

  Serve para conferir a tela REAL (MSForms), não a prévia do Edge: abre uma cópia do
  .xlsm de teste com o Excel visível, mostra o formulário SEM modo modal por um
  módulo temporário, captura a janela e fecha tudo sem salvar.

  Uso: capturar-tela.ps1 -Tela frmVinculo [-Arquivo <xlsm>]
  Saída: execucao\teste\tela-<Tela>.png
#>
param([Parameter(Mandatory = $true)][string]$Tela, [string]$Arquivo = '')
$ErrorActionPreference = 'Stop'
$dados = Split-Path -Parent $PSScriptRoot
. (Join-Path $dados 'lib\Comum.ps1')
. (Join-Path $dados 'build\lib\Excel.ps1')
Add-Type -AssemblyName System.Drawing
Add-Type @"
using System; using System.Runtime.InteropServices;
public static class Jan {
  [StructLayout(LayoutKind.Sequential)] public struct R { public int L, T, Ri, B; }
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern IntPtr FindWindow(string c, string t);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out R r);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
  [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr depois, int x, int y, int cx, int cy, uint f);
}
"@
[void][Jan]::SetProcessDPIAware()
$raiz = Get-RaizBase $PSScriptRoot
if (-not $Arquivo) { $Arquivo = Join-Path $dados 'execucao\teste\CRM_Zapromaq.xlsm' }
$copia = Join-Path $env:TEMP ('crm-tela-' + $PID + '.xlsm')
Copy-Item -LiteralPath $Arquivo -Destination $copia -Force
$saida = Join-Path (Get-PastaExecucao $raiz 'teste') ('tela-' + $Tela + '.png')

$vbom = Get-AcessoVBOM; Set-AcessoVBOM 1
$xl = $null; $id = 0
try {
    $xl = New-Excel ([ref]$id)
    $xl.EnableEvents = $false
    $wb = $xl.Workbooks.Open($copia)
    $m = $wb.VBProject.VBComponents.Add(1)
    $m.CodeModule.AddFromString(@"
Public Function MostrarTela() As String
    Dim f As Object
    Set f = VBA.UserForms.Add("$Tela")
    f.Show vbModeless
    MostrarTela = f.Caption
End Function
"@)
    $xl.Visible = $true
    $titulo = [string]$xl.Run("'" + $wb.Name + "'!MostrarTela")
    Start-Sleep -Milliseconds 1500
    $h = [Jan]::FindWindow('ThunderDFrame', $titulo)
    if ($h -eq [IntPtr]::Zero) { throw ('janela do formulário não encontrada: ' + $titulo) }
    # por cima de tudo durante a captura (HWND_TOPMOST; sem mover nem redimensionar)
    [void][Jan]::SetWindowPos($h, [IntPtr](-1), 0, 0, 0, 0, 0x43)
    [void][Jan]::SetForegroundWindow($h)
    Start-Sleep -Milliseconds 800
    $r = New-Object Jan+R; [void][Jan]::GetWindowRect($h, [ref]$r)
    $bmp = New-Object Drawing.Bitmap(($r.Ri - $r.L), ($r.B - $r.T))
    $g = [Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($r.L, $r.T, 0, 0, $bmp.Size)
    $bmp.Save($saida, [Drawing.Imaging.ImageFormat]::Png); $g.Dispose(); $bmp.Dispose()
    Write-Host ('print: ' + $saida)
} finally {
    if ($xl) { try { $xl.Visible = $false } catch { } }
    Close-Excel $xl $id
    Set-AcessoVBOM $vbom
    Remove-Item -LiteralPath $copia -Force -ErrorAction SilentlyContinue
}
