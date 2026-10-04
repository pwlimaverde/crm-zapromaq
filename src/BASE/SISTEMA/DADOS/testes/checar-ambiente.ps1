<#
  checar-ambiente.ps1 - confere se ESTA máquina tem o que o CRM precisa.

  Não altera nada. Relatório em SISTEMA\DADOS\execucao\logs\ambiente-*.txt.
  Rode na estação antes do primeiro build e sempre que algo "não funcionar
  só nesta máquina".

  Confere: Windows e PowerShell 64 bits; Excel (versão, edição, 64 bits);
  "Confiar no acesso ao modelo de objeto do projeto do VBA"; motor ACE;
  Edge (gerador visual); fontes Segoe UI e Segoe MDL2 Assets; resolução e
  escala da tela (layout de 1000 x 505 pt com Zoom automático); caminho UNC
  e permissão de escrita na pasta BASE.
#>
$ErrorActionPreference = 'Continue'
. (Join-Path (Split-Path -Parent $PSScriptRoot) 'lib\Comum.ps1')
$raiz = Get-RaizBase $PSScriptRoot
$log = New-Log $raiz 'ambiente'
$pendencias = 0
function Item([string]$nome, [bool]$ok, [string]$detalhe) {
    $marca = if ($ok) { 'OK      ' } else { 'ATENÇÃO ' }
    if (-not $ok) { $script:pendencias++ }
    $log.L('  ' + $marca + $nome.PadRight(36) + $detalhe)
}

$log.L('=== AMBIENTE DA ESTAÇÃO - CRM Zapromaq ===')
$log.L('Data....: ' + (Get-Date -Format 'dd/MM/yyyy HH:mm:ss') + '   ' + $env:COMPUTERNAME + '\' + $env:USERNAME)
$os = Get-CimInstance Win32_OperatingSystem
$log.L('Windows.: ' + $os.Caption + ' (' + $os.Version + ')')
$log.L('')

Item 'PowerShell 64 bits' ([Environment]::Is64BitProcess) ('versão ' + $PSVersionTable.PSVersion)

# ---------------------------------------------------------------- Excel
$c2r = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue
$edicao = if ($c2r) { [string]$c2r.ProductReleaseIds } else { '(instalação MSI ou não encontrada)' }
$bits = if ($c2r) { [string]$c2r.Platform } else { '' }
$xlVer = ''
try {
    $xl = [Activator]::CreateInstance([Type]::GetTypeFromProgID('Excel.Application'))   # ligação tardia: o COM puro não depende da biblioteca de tipos do Office
    $xlVer = [string]$xl.Version + ' build ' + [string]$xl.Build
    $xl.Quit(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    Item 'Excel instalado' $true ($xlVer + '   edição: ' + $edicao)
} catch { Item 'Excel instalado' $false 'não abriu via COM - build e front exigem Excel 2019 64 bits' }
if ($bits) { Item 'Office 64 bits' ($bits -eq 'x64') ('plataforma: ' + $bits) }
if ($edicao -match '365|O365') { Item 'Edição do Excel = 2019' $false ('é ' + $edicao + ': funciona, mas a validação final é no Excel 2019') }

$vbom = $null
foreach ($v in @('16.0')) {
    $k = Get-ItemProperty ('HKCU:\Software\Microsoft\Office\' + $v + '\Excel\Security') -ErrorAction SilentlyContinue
    if ($k -and ($k.PSObject.Properties.Name -contains 'AccessVBOM')) { $vbom = [int]$k.AccessVBOM }
}
Item 'Acesso ao modelo de objeto do VBA' ($vbom -eq 1) ('necessário SÓ na máquina que roda o build (MONTAR-FRONTEND.bat)')

# ---------------------------------------------------------------- banco
try { Item 'Motor ACE (Access Database Engine)' $true (Get-ProvedorAce) }
catch { Item 'Motor ACE (Access Database Engine)' $false $_.Exception.Message }
$banco = Get-CaminhoBanco $raiz
Item 'Banco na raiz de BASE' (Test-Path -LiteralPath $banco) $banco

# ---------------------------------------------------------------- visual
$edge = @("${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe", "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe") |
        Where-Object { Test-Path $_ } | Select-Object -First 1
Item 'Microsoft Edge (gerador visual)' ([bool]$edge) ([string]$edge)
Add-Type -AssemblyName System.Drawing
$fontes = (New-Object System.Drawing.Text.InstalledFontCollection).Families | ForEach-Object { $_.Name }
Item 'Fonte Segoe UI' ($fontes -contains 'Segoe UI') ''
Item 'Fonte Segoe MDL2 Assets (ícones)' ($fontes -contains 'Segoe MDL2 Assets') 'vem no Windows 10 e 11'

Add-Type -AssemblyName System.Windows.Forms
$tela = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
Add-Type -Namespace CrmDpi -Name Api -MemberDefinition @'
[DllImport("user32.dll")] public static extern System.IntPtr GetDC(System.IntPtr h);
[DllImport("user32.dll")] public static extern int ReleaseDC(System.IntPtr h, System.IntPtr dc);
[DllImport("gdi32.dll")] public static extern int GetDeviceCaps(System.IntPtr dc, int i);
'@
$dc = [CrmDpi.Api]::GetDC([IntPtr]::Zero); $dpi = [CrmDpi.Api]::GetDeviceCaps($dc, 88); [void][CrmDpi.Api]::ReleaseDC([IntPtr]::Zero, $dc)
# Este processo não declara suporte a DPI: o Windows lhe entrega tela e DPI
# "lógicos" (DPI sempre 96). A escala real vem do registro. A conta do zoom
# (área de trabalho em pontos = pixels x 72 / DPI) dá o mesmo resultado nos
# dois modos - é a mesma conta que o formulário faz no VBA (modTela).
$dpiReal = $dpi
$wm = Get-ItemProperty 'HKCU:\Control Panel\Desktop\WindowMetrics' -ErrorAction SilentlyContinue
if ($wm -and ($wm.PSObject.Properties.Name -contains 'AppliedDPI')) { $dpiReal = [int]$wm.AppliedDPI }
$area = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
$ptW = $area.Width * 72 / $dpi; $ptH = $area.Height * 72 / $dpi
$zoom = [math]::Min(100, [math]::Floor([math]::Min($ptW / 1000, $ptH / 540) * 100))
$fisW = [math]::Round($tela.Width * $dpiReal / $dpi); $fisH = [math]::Round($tela.Height * $dpiReal / $dpi)
Item 'Resolução da tela' ($fisW -ge 1366 -and $fisH -ge 768) ('' + $fisW + ' x ' + $fisH + ' (mínimo 1366 x 768)')
$log.L('  INFO    ' + 'Escala do Windows'.PadRight(36) + [math]::Round($dpiReal / 96 * 100) + '%  (' + $dpiReal + ' dpi); área útil ' +
       [math]::Round($ptW) + ' x ' + [math]::Round($ptH) + ' pt -> o formulário abre com Zoom ' + $zoom)

# ---------------------------------------------------------------- pasta
$unc = ConvertTo-CaminhoUNC $raiz
Item 'Caminho de BASE' $true $unc
$teste = Join-Path (Get-PastaExecucao $raiz 'logs') ('.teste-escrita-' + $PID)
try { Set-Content -LiteralPath $teste -Value 'ok'; Remove-Item -LiteralPath $teste -Force; Item 'Permissão de escrita em execucao\' $true '' }
catch { Item 'Permissão de escrita em execucao\' $false $_.Exception.Message }

$log.L('')
$log.L($(if ($pendencias -eq 0) { 'RESULTADO: ambiente pronto.' } else { 'RESULTADO: ' + $pendencias + ' ponto(s) de atenção acima.' }))
$log.Salvar()
