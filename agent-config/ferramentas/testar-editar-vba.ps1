<#
  testar-editar-vba.ps1 - testes do editar-vba.ps1 (troca de trecho em fonte VBA
  cp1252 + CRLF sem perder acento). Roda numa pasta temporaria; nao toca no repositorio.
  Sai 0 se tudo passou, 1 se algo falhou. Chamado pelo verificar-tudo (-Rapido).
#>
# Continue: no PowerShell 5.1 o stderr do editor (recusas esperadas) viraria erro fatal
$ErrorActionPreference = 'Continue'
$editor = Join-Path $PSScriptRoot 'editar-vba.ps1'
$ps = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$cp = [System.Text.Encoding]::GetEncoding(1252)
$pasta = Join-Path ([System.IO.Path]::GetTempPath()) ('crm-editar-vba-' + $PID)
New-Item -ItemType Directory -Force $pasta | Out-Null
$falhas = 0
function Conferir([string]$nome, [bool]$ok, [string]$det = '') {
    if ($ok) { Write-Output ('  OK    ' + $nome) } else { $script:falhas++; Write-Output ('  FALHA ' + $nome + '  ' + $det) }
}
# acentos montados por codigo: este arquivo fica em ASCII puro
$a_til = [string][char]0xE3; $c_ced = [string][char]0xE7; $e_agudo = [string][char]0xE9
$base = 'Attribute VB_Name = "modX"' + "`r`n" + "' linha um`r`n" +
        '    MsgBox "O banco n' + $a_til + 'o respondeu"' + "`r`n" + "' linha dois`r`n"
function Novo([string]$nome, [byte[]]$bytes) { $p = Join-Path $pasta $nome; [System.IO.File]::WriteAllBytes($p, $bytes); return $p }
function Editar([string[]]$argumentos) {
    $saida = & $ps -NoProfile -ExecutionPolicy Bypass -File $editor @argumentos 2>&1 | ForEach-Object { "$_" }
    return [pscustomobject]@{ Codigo = $LASTEXITCODE; Saida = ($saida -join ' | ') }
}
function Igual([byte[]]$x, [byte[]]$y) { if ($x.Length -ne $y.Length) { return $false }; for ($i = 0; $i -lt $x.Length; $i++) { if ($x[$i] -ne $y[$i]) { return $false } }; return $true }
function SoCrlf([byte[]]$b) { for ($i = 0; $i -lt $b.Length; $i++) { if ($b[$i] -eq 10 -and ($i -eq 0 -or $b[$i - 1] -ne 13)) { return $false } }; return $true }

try {
    # 1. linha ASCII trocada num arquivo com acento em outra linha: so ela muda
    $p = Novo 'um.bas' $cp.GetBytes($base)
    $r = Editar @('-Arquivo', $p, '-Antigo', "' linha dois", '-Novo', "' linha DOIS")
    $esperado = $cp.GetBytes($base.Replace("' linha dois", "' linha DOIS"))
    Conferir 'troca ASCII preserva o acento das outras linhas (byte a byte)' ($r.Codigo -eq 0 -and (Igual ([System.IO.File]::ReadAllBytes($p)) $esperado)) $r.Saida

    # 2. texto novo com acento vira byte cp1252 (E7), nunca UTF-8 (C3 A7)
    $p = Novo 'dois.bas' $cp.GetBytes($base)
    $fa = Join-Path $pasta 'a.txt'; $fn = Join-Path $pasta 'n.txt'
    [System.IO.File]::WriteAllText($fa, ('O banco n' + $a_til + 'o respondeu'), (New-Object System.Text.UTF8Encoding($false)))
    [System.IO.File]::WriteAllText($fn, ('A conex' + $a_til + 'o ' + $c_ced + $e_agudo + ' nova'), (New-Object System.Text.UTF8Encoding($false)))
    $r = Editar @('-Arquivo', $p, '-AntigoArquivo', $fa, '-NovoArquivo', $fn)
    $b = [System.IO.File]::ReadAllBytes($p)
    $esperado = $cp.GetBytes($base.Replace('O banco n' + $a_til + 'o respondeu', 'A conex' + $a_til + 'o ' + $c_ced + $e_agudo + ' nova'))
    Conferir 'acento no texto novo e gravado em cp1252 (arquivos UTF-8 de entrada)' ($r.Codigo -eq 0 -and (Igual $b $esperado)) $r.Saida

    # 3. texto novo com LF: no arquivo vira CRLF
    $p = Novo 'tres.bas' $cp.GetBytes($base)
    $r = Editar @('-Arquivo', $p, '-Antigo', "' linha um", '-Novo', ("' linha um`n' linha um e meio"))
    $b = [System.IO.File]::ReadAllBytes($p)
    Conferir 'quebra de linha do texto novo vira CRLF' ($r.Codigo -eq 0 -and (SoCrlf $b) -and $cp.GetString($b).Contains("' linha um`r`n' linha um e meio`r`n")) $r.Saida

    # 4-8. recusas: o arquivo fica intacto
    $casos = @(
        @{ n = 'trecho inexistente e recusado'; b = $cp.GetBytes($base); a = @('-Antigo', 'nao existe', '-Novo', 'x') },
        @{ n = 'trecho repetido e recusado'; b = $cp.GetBytes($base + "' linha dois`r`n"); a = @('-Antigo', "' linha dois", '-Novo', 'x') },
        @{ n = 'arquivo ja corrompido (U+FFFD) e recusado'; b = [byte[]]($cp.GetBytes("' a`r`n") + [byte[]](0xEF, 0xBF, 0xBD) + $cp.GetBytes("`r`n")); a = @('-Antigo', "' a", '-Novo', "' b") },
        @{ n = 'arquivo com BOM e recusado'; b = [byte[]]([byte[]](0xEF, 0xBB, 0xBF) + $cp.GetBytes("' a`r`n")); a = @('-Antigo', "' a", '-Novo', "' b") },
        @{ n = 'caractere fora do cp1252 e recusado'; b = $cp.GetBytes($base); a = @('-Antigo', "' linha um", '-Novo', ("' seta " + [char]0x2192)) },
        @{ n = 'arquivo com LF sem CR e recusado'; b = $cp.GetBytes("' a`n' b`n"); a = @('-Antigo', "' a", '-Novo', "' c") }
    )
    $k = 0
    foreach ($c in $casos) {
        $k++
        $p = Novo ('recusa' + $k + '.bas') $c.b
        $r = Editar (@('-Arquivo', $p) + $c.a)
        Conferir $c.n ($r.Codigo -ne 0 -and (Igual ([System.IO.File]::ReadAllBytes($p)) $c.b)) ('codigo=' + $r.Codigo + ' ' + $r.Saida)
    }

    # 10. -Simular nao grava
    $p = Novo 'simular.bas' $cp.GetBytes($base)
    $r = Editar @('-Arquivo', $p, '-Antigo', "' linha um", '-Novo', "' linha UM", '-Simular')
    Conferir '-Simular nao grava' ($r.Codigo -eq 0 -and (Igual ([System.IO.File]::ReadAllBytes($p)) $cp.GetBytes($base))) $r.Saida
}
finally { Remove-Item -Recurse -Force $pasta -ErrorAction SilentlyContinue }

if ($falhas -gt 0) { Write-Output ('  ' + $falhas + ' FALHA(s)'); exit 1 }
Write-Output '  TUDO CERTO.'
exit 0
