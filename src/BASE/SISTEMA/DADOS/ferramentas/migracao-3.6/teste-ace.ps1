$ErrorActionPreference = 'Stop'
$base = Split-Path -Parent $MyInvocation.MyCommand.Path
$db   = Join-Path $base 'teste_zapromaq.accdb'
$log  = Join-Path $base ('resultado-teste-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.txt')
$saida = New-Object System.Collections.ArrayList
function L($t) { [void]$saida.Add([string]$t); Write-Host $t }
function Fim { $saida | Set-Content -Path $log -Encoding UTF8; Write-Host ""; Write-Host ("Relatorio salvo em: " + $log) }

Add-Type -AssemblyName System.Data

L "================================================================"
L " TESTE DE VIABILIDADE - ACE / ACCDB - Comercial Zapromaq  (v4)"
L "================================================================"
L ("Data..............: " + (Get-Date -Format 'dd/MM/yyyy HH:mm:ss'))
L ("Maquina / usuario.: " + $env:COMPUTERNAME + " / " + $env:USERNAME)
L ("PowerShell........: " + $PSVersionTable.PSVersion + "   64-bit: " + [Environment]::Is64BitProcess)
L ("Pasta do teste....: " + $base)
L ""

if (-not [Environment]::Is64BitProcess) { L "Rode em 64 bits (RODAR-TESTE.bat)."; Fim; exit 1 }

# ---------------------------------------------------------------
# 1. Provedor
# ---------------------------------------------------------------
L "--- 1. PROVEDOR OLE DB ---"
foreach ($ext in @('accdb','laccdb','ldb')) {
    $arq = [System.IO.Path]::ChangeExtension($db, $ext)
    if (Test-Path $arq) { Remove-Item $arq -Force -ErrorAction SilentlyContinue }
}
$provider = $null
foreach ($prov in @('Microsoft.ACE.OLEDB.16.0','Microsoft.ACE.OLEDB.12.0')) {
    try {
        $cat = New-Object -ComObject ADOX.Catalog
        $null = $cat.Create("Provider=$prov;Data Source=$db;")
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($cat) | Out-Null
        $cat = $null
        $provider = $prov
        L ("  OK ..... " + $prov)
        break
    } catch { L ("  FALHOU . " + $prov + "   (" + $_.Exception.Message.Split([char]13)[0] + ")") }
}
if (-not $provider) { L "NENHUM ACE DISPONIVEL."; Fim; exit 1 }
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
L ("  Arquivo criado: " + (Get-Item $db).Length + " bytes")

$cs = "Provider=$provider;Data Source=$db;Persist Security Info=False;"
$cn = New-Object System.Data.OleDb.OleDbConnection($cs)
$cn.Open()
L "  Conexao aberta."

# --- tipos (nomes com prefixo tp: PowerShell nao diferencia maiusculas) ---
$tpTexto = [System.Data.OleDb.OleDbType]::VarWChar
$tpMemo  = [System.Data.OleDb.OleDbType]::LongVarWChar
$tpData  = [System.Data.OleDb.OleDbType]::Date
$tpInt   = [System.Data.OleDb.OleDbType]::Integer

function NovoCmd($transacao) {
    $cmd = $cn.CreateCommand()
    if ($null -ne $transacao) { $cmd.Transaction = $transacao }
    return $cmd
}
function AddPar($cmd, $tipo, $valor) {
    $par = $cmd.CreateParameter()
    $par.OleDbType = $tipo
    if ($null -eq $valor) { $par.Value = [DBNull]::Value } else { $par.Value = $valor }
    [void]$cmd.Parameters.Add($par)
}
function Exec($sql) { $cmd = NovoCmd $null; $cmd.CommandText = $sql; $qt = $cmd.ExecuteNonQuery(); $cmd.Dispose(); return $qt }
function Escalar($sql) { $cmd = NovoCmd $null; $cmd.CommandText = $sql; $v = $cmd.ExecuteScalar(); $cmd.Dispose(); return $v }

# ---------------------------------------------------------------
# 2. Tabela
# ---------------------------------------------------------------
L ""
L "--- 2. CRIACAO DA TABELA ---"
$ddl = "CREATE TABLE teste_clientes (" +
       " id AUTOINCREMENT PRIMARY KEY," +
       " codigo_cliente TEXT(20)," +
       " empresa TEXT(120)," +
       " endereco TEXT(180)," +
       " contexto_geral MEMO," +
       " versao LONG," +
       " alterado_em DATETIME," +
       " alterado_por TEXT(50))"
$null = Exec $ddl
$null = Exec "CREATE UNIQUE INDEX ix_cod ON teste_clientes (codigo_cliente)"
L "  Tabela teste_clientes criada, com indice unico em codigo_cliente."
if (Test-Path ([System.IO.Path]::ChangeExtension($db,'laccdb'))) { L "  Arquivo de bloqueio .laccdb presente na pasta de rede.  OK" }
else { L "  ATENCAO: .laccdb ausente." }

# ---------------------------------------------------------------
# 3. INSERT parametrizado
# ---------------------------------------------------------------
L ""
L "--- 3. INCLUSAO (INSERT parametrizado) ---"
$textoCtx = ('Contexto geral de teste com acentuacao: producao, inox, secao, cabecalho. ' * 120)
$nomeApos = "INDUSTRIA D" + [char]39 + "ANGELO LTDA"
$sqlIns = "INSERT INTO teste_clientes (codigo_cliente, empresa, endereco, contexto_geral, versao, alterado_em, alterado_por) VALUES (?,?,?,?,1,?,?)"

$dados = @()
$dados += [pscustomobject]@{ cod='0011'; emp='METALURGICA EXEMPLO LTDA'; end='RUA DOS TESTES, 100 - CENTRO, CIDADE MODELO/RS'; ctx=$textoCtx }
$dados += [pscustomobject]@{ cod='0002'; emp='TUBULARES EXEMPLO'; end='Av. Joao Pessoa, 1.234 - acentuacao: producao, secao, inox'; ctx='memo curto' }
$dados += [pscustomobject]@{ cod='0099'; emp=$nomeApos;           end='Rua do Apostrofo, 1'; ctx=("contexto com " + [char]39 + "aspas" + [char]39 + " internas") }

foreach ($reg in $dados) {
    $cmd = NovoCmd $null
    $cmd.CommandText = $sqlIns
    AddPar $cmd $tpTexto $reg.cod
    AddPar $cmd $tpTexto $reg.emp
    AddPar $cmd $tpTexto $reg.end
    AddPar $cmd $tpMemo  $reg.ctx
    AddPar $cmd $tpData  (Get-Date)
    AddPar $cmd $tpTexto $env:USERNAME
    $qt = $cmd.ExecuteNonQuery()
    $cmd.Dispose()
    L ("  incluido (" + $qt + " linha): " + $reg.cod + "  " + $reg.emp)
}
L ("  @@IDENTITY na mesma conexao: " + (Escalar "SELECT @@IDENTITY") + "   <- e assim que o id novo deve ser lido")

# ---------------------------------------------------------------
# 4. Consulta
# ---------------------------------------------------------------
L ""
L "--- 4. CONSULTA (SELECT) ---"
$cmd = NovoCmd $null
$cmd.CommandText = "SELECT id, codigo_cliente, empresa, endereco, LEN(contexto_geral) AS tam FROM teste_clientes ORDER BY codigo_cliente"
$rd = $cmd.ExecuteReader()
while ($rd.Read()) {
    L ("  id=" + $rd['id'] + "  cod=" + $rd['codigo_cliente'] + "  " + $rd['empresa'] + "  | memo: " + $rd['tam'] + " car.")
    L ("      endereco: " + $rd['endereco'])
}
$rd.Close(); $cmd.Dispose()

# ---------------------------------------------------------------
# 5. EDICAO pelo codigo, com controle de versao
# ---------------------------------------------------------------
L ""
L "--- 5. EDICAO (UPDATE pelo codigo, com conferencia) ---"
$alvo = '0011'
$enderecoNovo = 'AVENIDA NOVA DE TESTE, 4321 - DISTRITO INDUSTRIAL, CAXIAS DO SUL/RS'
$cmd = NovoCmd $null
$cmd.CommandText = "SELECT endereco, versao FROM teste_clientes WHERE codigo_cliente='$alvo'"
$rd = $cmd.ExecuteReader(); [void]$rd.Read()
$enderecoAntes = [string]$rd['endereco']; $versaoLida = [int]$rd['versao']
$rd.Close(); $cmd.Dispose()
L ("  valor atual .....: " + $enderecoAntes)
L ("  versao lida .....: " + $versaoLida)

$sqlUpd = "UPDATE teste_clientes SET endereco=?, versao=versao+1, alterado_em=?, alterado_por=? WHERE codigo_cliente=? AND versao=?"
$cmd = NovoCmd $null
$cmd.CommandText = $sqlUpd
AddPar $cmd $tpTexto $enderecoNovo
AddPar $cmd $tpData  (Get-Date)
AddPar $cmd $tpTexto $env:USERNAME
AddPar $cmd $tpTexto $alvo
AddPar $cmd $tpInt   $versaoLida
$afetadas = $cmd.ExecuteNonQuery(); $cmd.Dispose()
L ("  linhas afetadas .: " + $afetadas + "   (tem que ser exatamente 1)")
L ("  valor depois ....: " + (Escalar "SELECT endereco FROM teste_clientes WHERE codigo_cliente='$alvo'"))
L ("  versao depois ...: " + (Escalar "SELECT versao FROM teste_clientes WHERE codigo_cliente='$alvo'"))

$cmd = NovoCmd $null
$cmd.CommandText = $sqlUpd
AddPar $cmd $tpTexto 'tentativa com versao velha'
AddPar $cmd $tpData  (Get-Date)
AddPar $cmd $tpTexto $env:USERNAME
AddPar $cmd $tpTexto $alvo
AddPar $cmd $tpInt   $versaoLida
$afetadas2 = $cmd.ExecuteNonQuery(); $cmd.Dispose()
L ("  repetindo com a versao antiga -> linhas afetadas: " + $afetadas2 + "   (0 = alguem alterou desde a leitura; e assim que se detecta edicao simultanea)")

# ---------------------------------------------------------------
# 6. Memo, acento e apostrofo
# ---------------------------------------------------------------
L ""
L "--- 6. MEMO E CARACTERES ---"
$ctxVolta = [string](Escalar "SELECT contexto_geral FROM teste_clientes WHERE codigo_cliente='0011'")
L ("  memo gravado: " + $textoCtx.Length + " car. | lido: " + $ctxVolta.Length + " car. | identico: " + ($ctxVolta -eq $textoCtx))
L ("  acentuacao ..: " + (Escalar "SELECT endereco FROM teste_clientes WHERE codigo_cliente='0002'"))
L ("  apostrofo ...: " + (Escalar "SELECT empresa  FROM teste_clientes WHERE codigo_cliente='0099'"))

# ---------------------------------------------------------------
# 7. Transacao
# ---------------------------------------------------------------
L ""
L "--- 7. TRANSACAO (dado + auditoria na mesma unidade) ---"
$null = Exec "CREATE TABLE teste_log (id AUTOINCREMENT PRIMARY KEY, tabela TEXT(30), id_registro TEXT(30), acao TEXT(20), usuario TEXT(50), quando DATETIME, detalhe MEMO)"
$trans = $cn.BeginTransaction()
try {
    $cmd = NovoCmd $trans
    $cmd.CommandText = "INSERT INTO teste_clientes (codigo_cliente, empresa, versao) VALUES ('0500','EMPRESA EM TRANSACAO',1)"
    [void]$cmd.ExecuteNonQuery()
    $cmd.CommandText = "INSERT INTO teste_log (tabela, id_registro, acao, usuario, quando) VALUES ('teste_clientes','0500','INSERIR','" + $env:USERNAME + "',Now())"
    [void]$cmd.ExecuteNonQuery()
    $cmd.Dispose()
    $trans.Commit()
    L "  commit: dado e log gravados juntos.  OK"
} catch { $trans.Rollback(); L ("  ROLLBACK inesperado: " + $_.Exception.Message) }

$trans = $cn.BeginTransaction()
try {
    $cmd = NovoCmd $trans
    $cmd.CommandText = "INSERT INTO teste_clientes (codigo_cliente, empresa, versao) VALUES ('0501','SERA DESFEITA',1)"
    [void]$cmd.ExecuteNonQuery()
    $cmd.Dispose()
    $trans.Rollback()
    L ("  rollback testado: registro 0501 existe? " + (Escalar "SELECT COUNT(*) FROM teste_clientes WHERE codigo_cliente='0501'") + "   (0 = a transacao desfez, como deve)")
} catch { L ("  erro no teste de rollback: " + $_.Exception.Message) }

# ---------------------------------------------------------------
# 8. Desempenho na rede
# ---------------------------------------------------------------
L ""
L "--- 8. DESEMPENHO NA PASTA DE REDE ---"
$inicio = Get-Date
$trans = $cn.BeginTransaction()
for ($i=1; $i -le 500; $i++) {
    $cmd = NovoCmd $trans
    $cmd.CommandText = "INSERT INTO teste_clientes (codigo_cliente, empresa, versao) VALUES (?,?,1)"
    AddPar $cmd $tpTexto ('T' + $i.ToString('0000'))
    AddPar $cmd $tpTexto ('Empresa de carga ' + $i)
    [void]$cmd.ExecuteNonQuery()
    $cmd.Dispose()
}
$trans.Commit()
$msLote = ((Get-Date) - $inicio).TotalMilliseconds
L ("  500 insercoes em transacao unica ...: " + [math]::Round($msLote) + " ms  (" + [math]::Round($msLote/500,2) + " ms/registro)")

$inicio = Get-Date
for ($i=501; $i -le 550; $i++) {
    $cmd = NovoCmd $null
    $cmd.CommandText = "INSERT INTO teste_clientes (codigo_cliente, empresa, versao) VALUES (?,?,1)"
    AddPar $cmd $tpTexto ('T' + $i.ToString('0000'))
    AddPar $cmd $tpTexto ('Avulso ' + $i)
    [void]$cmd.ExecuteNonQuery()
    $cmd.Dispose()
}
$msAvulso = ((Get-Date) - $inicio).TotalMilliseconds
L ("  50 insercoes avulsas ...............: " + [math]::Round($msAvulso) + " ms  (" + [math]::Round($msAvulso/50,2) + " ms/registro)")

$totalReg = Escalar "SELECT COUNT(*) FROM teste_clientes"
$inicio = Get-Date
$achados = Escalar "SELECT COUNT(*) FROM teste_clientes WHERE empresa LIKE '%carga%'"
$msLike = ((Get-Date) - $inicio).TotalMilliseconds
L ("  busca LIKE sobre " + $totalReg + " registros .: " + [math]::Round($msLike) + " ms  (" + $achados + " encontrados)")

$inicio = Get-Date
$null = Escalar "SELECT COUNT(*) FROM teste_clientes WHERE codigo_cliente='T0250'"
$msIndice = ((Get-Date) - $inicio).TotalMilliseconds
L ("  busca por indice unico .............: " + [math]::Round($msIndice,1) + " ms")

# ---------------------------------------------------------------
# 9. Integridade
# ---------------------------------------------------------------
L ""
L "--- 9. INTEGRIDADE ---"
try { $null = Exec "INSERT INTO teste_clientes (codigo_cliente, empresa) VALUES ('0011','Duplicata proibida')"
      L "  FALHA DO TESTE: o banco aceitou codigo duplicado." }
catch { L "  OK: duplicata recusada pelo indice unico." }
try { $null = Exec "INSERT INTO teste_clientes (codigo_cliente, empresa, alterado_em) VALUES ('0600','Data invalida',#02/31/2026#)"
      L "  ATENCAO: data invalida aceita." }
catch { L "  OK: data invalida recusada." }

# ---------------------------------------------------------------
# 10. Conexao curta
# ---------------------------------------------------------------
$cn.Close(); $cn.Dispose()
L ""
L "--- 10. CONEXAO CURTA (abre, executa, fecha) ---"
$inicio = Get-Date
for ($i=1; $i -le 20; $i++) {
    $cnCurta = New-Object System.Data.OleDb.OleDbConnection($cs)
    $cnCurta.Open()
    $cmdCurto = $cnCurta.CreateCommand(); $cmdCurto.CommandText = "SELECT COUNT(*) FROM teste_clientes"
    [void]$cmdCurto.ExecuteScalar(); $cmdCurto.Dispose()
    $cnCurta.Close(); $cnCurta.Dispose()
}
$msCiclo = ((Get-Date) - $inicio).TotalMilliseconds
L ("  20 ciclos abre/consulta/fecha: " + [math]::Round($msCiclo) + " ms  (" + [math]::Round($msCiclo/20,1) + " ms por ciclo)")
L ("  Tamanho final do .accdb: " + [math]::Round((Get-Item $db).Length/1KB) + " KB")

L ""
L "================================================================"
L " FIM"
L "================================================================"
Fim
