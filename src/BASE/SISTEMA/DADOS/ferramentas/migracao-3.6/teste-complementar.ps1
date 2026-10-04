$ErrorActionPreference = 'Stop'
$base = Split-Path -Parent $MyInvocation.MyCommand.Path
$db   = Join-Path $base 'teste_complementar.accdb'
$log  = Join-Path $base ('resultado-complementar-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.txt')
$saida = New-Object System.Collections.ArrayList
function L($t) { [void]$saida.Add([string]$t); Write-Host $t }
function Fim { $saida | Set-Content -Path $log -Encoding UTF8; Write-Host ""; Write-Host ("Relatorio salvo em: " + $log) }

Add-Type -AssemblyName System.Data

L "================================================================"
L " TESTE COMPLEMENTAR - decisoes de modelagem - Comercial Zapromaq"
L "================================================================"
L ("Data.....: " + (Get-Date -Format 'dd/MM/yyyy HH:mm:ss'))
L ("Maquina..: " + $env:COMPUTERNAME + " / " + $env:USERNAME)
L ""

foreach ($ext in @('accdb','laccdb')) {
    $arq = [System.IO.Path]::ChangeExtension($db, $ext)
    if (Test-Path $arq) { Remove-Item $arq -Force -ErrorAction SilentlyContinue }
}
$provider = $null
foreach ($prov in @('Microsoft.ACE.OLEDB.16.0','Microsoft.ACE.OLEDB.12.0')) {
    try {
        $cat = New-Object -ComObject ADOX.Catalog
        $null = $cat.Create("Provider=$prov;Data Source=$db;")
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($cat) | Out-Null
        $cat = $null; $provider = $prov; break
    } catch { }
}
if (-not $provider) { L "Nenhum ACE disponivel."; Fim; exit 1 }
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
L ("Provedor: " + $provider)

$cs = "Provider=$provider;Data Source=$db;Persist Security Info=False;"
$cn = New-Object System.Data.OleDb.OleDbConnection($cs)
$cn.Open()

$tpTexto  = [System.Data.OleDb.OleDbType]::VarWChar
$tpMoeda  = [System.Data.OleDb.OleDbType]::Currency

function Exec($sql) { $cmd = $cn.CreateCommand(); $cmd.CommandText = $sql; $qt = $cmd.ExecuteNonQuery(); $cmd.Dispose(); return $qt }
function Escalar($sql) {
    $cmd = $cn.CreateCommand(); $cmd.CommandText = $sql
    $v = $cmd.ExecuteScalar(); $cmd.Dispose()
    if ($null -eq $v -or $v -is [System.DBNull]) { return $null }
    return $v
}

# ---------------------------------------------------------------
L ""
L "--- 1. INDICE UNICO COM VARIOS NULOS (as 372 pre-clientes) ---"
try {
    $null = Exec ("CREATE TABLE tclientes (" +
        " id AUTOINCREMENT PRIMARY KEY," +
        " codigo_cliente LONG," +
        " estagio TEXT(12) NOT NULL," +
        " empresa TEXT(120) NOT NULL," +
        " ativo YESNO, versao LONG)")
    $null = Exec "CREATE UNIQUE INDEX ix_codcli ON tclientes (codigo_cliente)"
    L "  tabela e indice unico criados."

    for ($i=1; $i -le 5; $i++) {
        $null = Exec ("INSERT INTO tclientes (codigo_cliente, estagio, empresa, ativo, versao) VALUES (NULL,'Pre-cliente','Pre-cliente " + $i + "',True,1)")
    }
    L ("  5 registros com codigo NULL inseridos: " + (Escalar "SELECT COUNT(*) FROM tclientes WHERE codigo_cliente IS NULL") + " na tabela")
    L "  >>> indice unico ACEITA varios nulos."
} catch {
    L ("  >>> indice unico RECUSA multiplos nulos: " + $_.Exception.Message.Split([char]13)[0])
    L "      SAIDA: indice NAO unico + validacao de duplicata no front e no executor."
}

try {
    $null = Exec "INSERT INTO tclientes (codigo_cliente, estagio, empresa, ativo, versao) VALUES (11,'Cliente','Cliente A',True,1)"
    $null = Exec "INSERT INTO tclientes (codigo_cliente, estagio, empresa, ativo, versao) VALUES (11,'Cliente','Cliente duplicado',True,1)"
    L "  ATENCAO: codigo preenchido duplicado foi ACEITO."
} catch { L "  OK: codigo preenchido duplicado recusado." }

# proximo codigo SEM usar NZ - o tratamento de nulo e do lado do cliente
try {
    $maxCod = Escalar "SELECT MAX(codigo_cliente) FROM tclientes"
    $proximo = if ($null -eq $maxCod) { 1 } else { [int]$maxCod + 1 }
    L ("  proximo codigo (MAX+1, nulo tratado no cliente): " + $proximo)
} catch { L ("  FALHA ao calcular proximo codigo: " + $_.Exception.Message.Split([char]13)[0]) }

# ---------------------------------------------------------------
L ""
L "--- 2. INTEGRIDADE REFERENCIAL ---"
try {
    $null = Exec ("CREATE TABLE tcontatos (" +
        " id AUTOINCREMENT PRIMARY KEY," +
        " codigo TEXT(15) NOT NULL," +
        " controle LONG NOT NULL," +
        " id_cliente LONG NOT NULL," +
        " nome TEXT(120) NOT NULL, versao LONG)")
    try {
        $null = Exec ("ALTER TABLE tcontatos ADD CONSTRAINT fk_cto_cli FOREIGN KEY (id_cliente) REFERENCES tclientes (id)")
        L "  constraint de chave estrangeira aplicada."
    } catch { L ("  FALHOU ao aplicar FK: " + $_.Exception.Message.Split([char]13)[0]) }

    $idCli = [int](Escalar "SELECT id FROM tclientes WHERE codigo_cliente=11")
    $null = Exec ("INSERT INTO tcontatos (codigo, controle, id_cliente, nome, versao) VALUES ('CT-0011-0001',1," + $idCli + ",'Contato de teste',1)")
    L "  contato vinculado criado."

    try {
        $null = Exec ("DELETE FROM tclientes WHERE id=" + $idCli)
        L "  ATENCAO: cliente com contato foi EXCLUIDO. A integridade nao esta protegendo."
    } catch { L "  OK: exclusao de cliente com contato recusada pelo banco." }

    try {
        $null = Exec "INSERT INTO tcontatos (codigo, controle, id_cliente, nome, versao) VALUES ('CT-9999-9999',99,999999,'Orfao',1)"
        L "  ATENCAO: contato apontando para cliente inexistente foi ACEITO."
    } catch { L "  OK: contato orfao recusado pelo banco." }
} catch { L ("  bloco 2 interrompido: " + $_.Exception.Message.Split([char]13)[0]) }

# ---------------------------------------------------------------
L ""
L "--- 3. ORDENACAO E BUSCA COM ACENTO ---"
try {
    foreach ($nomeEmp in @('ACOS LEVES','ACUCAR MODELO','ACO NOBRE','ABETO')) {
        $cmd = $cn.CreateCommand()
        $cmd.CommandText = "INSERT INTO tclientes (estagio, empresa, ativo, versao) VALUES ('Pre-cliente',?,True,1)"
        $par = $cmd.CreateParameter(); $par.OleDbType = $tpTexto; $par.Value = $nomeEmp
        [void]$cmd.Parameters.Add($par)
        [void]$cmd.ExecuteNonQuery(); $cmd.Dispose()
    }
    $cmd = $cn.CreateCommand()
    $cmd.CommandText = "SELECT empresa FROM tclientes WHERE empresa LIKE 'A%' ORDER BY empresa"
    $rd = $cmd.ExecuteReader()
    $ordemLista = @()
    while ($rd.Read()) { $ordemLista += [string]$rd['empresa'] }
    $rd.Close(); $cmd.Dispose()
    L ("  ordem devolvida: " + ($ordemLista -join ' | '))
} catch { L ("  bloco 3 interrompido: " + $_.Exception.Message.Split([char]13)[0]) }

# ---------------------------------------------------------------
L ""
L "--- 4. CONSULTAS AGREGADAS (o que vai alimentar o Painel) ---"
try {
    $null = Exec ("CREATE TABLE toportunidades (" +
        " id AUTOINCREMENT PRIMARY KEY, codigo TEXT(15), id_contato LONG," +
        " etapa TEXT(40), responsavel TEXT(60), valor CURRENCY, dt_prox_acao DATETIME, versao LONG)")
    $etapasTeste = @('Contato Inicial','Proposta em Stand By','Negociacao','Perdido','Pedido Fechado')
    $trans = $cn.BeginTransaction()
    for ($i=1; $i -le 500; $i++) {
        $cmd = $cn.CreateCommand(); $cmd.Transaction = $trans
        $cmd.CommandText = "INSERT INTO toportunidades (codigo, id_contato, etapa, responsavel, valor, versao) VALUES (?,1,?,?,?,1)"
        $p1 = $cmd.CreateParameter(); $p1.OleDbType = $tpTexto; $p1.Value = ('AT-0011-' + $i.ToString('0000')); [void]$cmd.Parameters.Add($p1)
        $p2 = $cmd.CreateParameter(); $p2.OleDbType = $tpTexto; $p2.Value = $etapasTeste[$i % 5]; [void]$cmd.Parameters.Add($p2)
        $p3 = $cmd.CreateParameter(); $p3.OleDbType = $tpTexto; $p3.Value = ('Vendedor ' + ($i % 4)); [void]$cmd.Parameters.Add($p3)
        $p4 = $cmd.CreateParameter(); $p4.OleDbType = $tpMoeda; $p4.Value = ($i * 137.5); [void]$cmd.Parameters.Add($p4)
        [void]$cmd.ExecuteNonQuery(); $cmd.Dispose()
    }
    $trans.Commit()
    L "  500 oportunidades de teste criadas."

    $inicio = Get-Date
    $cmd = $cn.CreateCommand()
    $cmd.CommandText = "SELECT etapa, COUNT(*) AS qt, SUM(valor) AS total FROM toportunidades GROUP BY etapa ORDER BY etapa"
    $rd = $cmd.ExecuteReader()
    while ($rd.Read()) { L ("    " + $rd['etapa'].ToString().PadRight(24) + " qt=" + $rd['qt'] + "  total=" + [math]::Round([double]$rd['total'],2)) }
    $rd.Close(); $cmd.Dispose()
    L ("  GROUP BY por etapa: " + [math]::Round(((Get-Date) - $inicio).TotalMilliseconds) + " ms")

    $inicio = Get-Date
    $cmd = $cn.CreateCommand()
    $cmd.CommandText = "SELECT responsavel, COUNT(*) AS qt, SUM(valor) AS total FROM toportunidades WHERE etapa NOT IN ('Perdido','Pedido Fechado') GROUP BY responsavel"
    $rd = $cmd.ExecuteReader()
    $qtLinhas = 0
    while ($rd.Read()) { $qtLinhas++ }
    $rd.Close(); $cmd.Dispose()
    L ("  GROUP BY por responsavel, com filtro: " + [math]::Round(((Get-Date) - $inicio).TotalMilliseconds) + " ms (" + $qtLinhas + " linhas)")
} catch { L ("  bloco 4 interrompido: " + $_.Exception.Message.Split([char]13)[0]) }

# ---------------------------------------------------------------
L ""
L "--- 5. VOCABULARIO DE FUNCOES DISPONIVEL VIA OLE DB ---"
L "  (o que passar aqui pode ser usado no Painel e nos campos calculados;"
L "   o que falhar tem de ser feito em VBA ou no PowerShell)"
$testes = @(
    @{ n='IIF';           s="SELECT IIF(1=1,'sim','nao')" },
    @{ n='IS NULL';       s="SELECT COUNT(*) FROM tclientes WHERE codigo_cliente IS NULL" },
    @{ n='NZ';            s="SELECT NZ(MAX(codigo_cliente),0) FROM tclientes" },
    @{ n='NOW';           s="SELECT NOW()" },
    @{ n='DATE';          s="SELECT DATE()" },
    @{ n='DATEDIFF';      s="SELECT DATEDIFF('d',#01/01/2026#,NOW())" },
    @{ n='DATEADD';       s="SELECT DATEADD('d',7,#01/01/2026#)" },
    @{ n='FORMAT';        s="SELECT FORMAT(#01/03/2026#,'yyyymm')" },
    @{ n='YEAR / MONTH';  s="SELECT YEAR(#01/03/2026#)*100+MONTH(#01/03/2026#)" },
    @{ n='LEN';           s="SELECT LEN('abcdef')" },
    @{ n='LEFT / MID';    s="SELECT LEFT('AT-0011-0148',2) & MID('AT-0011-0148',4,4)" },
    @{ n='SWITCH';        s="SELECT SWITCH(1=1,'a',1=2,'b')" },
    @{ n='VAL';           s="SELECT VAL('0148')" },
    @{ n='CONCAT &';      s="SELECT 'AT-' & '0011' & '-' & '0148'" },
    @{ n='TOP n';         s="SELECT TOP 3 empresa FROM tclientes ORDER BY empresa" },
    @{ n='SUBCONSULTA';   s="SELECT COUNT(*) FROM tclientes WHERE id IN (SELECT id_cliente FROM tcontatos)" },
    @{ n='LEFT JOIN';     s="SELECT COUNT(*) FROM tclientes cl LEFT JOIN tcontatos ct ON cl.id=ct.id_cliente" },
    @{ n='HAVING';        s="SELECT etapa FROM toportunidades GROUP BY etapa HAVING COUNT(*) > 50" }
)
foreach ($t in $testes) {
    try {
        $v = Escalar $t.s
        if ($null -eq $v) { $v = '(nulo)' }
        L ("  OK     " + $t.n.PadRight(14) + " -> " + $v)
    } catch {
        L ("  FALHA  " + $t.n.PadRight(14) + " -> " + $_.Exception.Message.Split([char]13)[0])
    }
}

$cn.Close(); $cn.Dispose()
L ""
L "================================================================"
L " FIM"
L "================================================================"
Fim
