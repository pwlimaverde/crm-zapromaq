<#
  criar-banco.ps1 - cria o .accdb e toda a estrutura do CRM Zapromaq.
  Roda na estacao Windows, PowerShell 64 bits. Nada a instalar.
  Nao carrega dado: isso e o carregar.ps1.
#>
param(
    [string]$Banco = "",
    [ValidateSet('TESTE','PRODUCAO')][string]$Ambiente = 'TESTE',
    [switch]$Recriar
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Data

$base = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($Banco)) { $Banco = Join-Path $base 'crm_zapromaq_teste.accdb' }
$log = Join-Path $base ('log-criar-banco-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.txt')
$saida = New-Object System.Collections.ArrayList
function L($t) { [void]$saida.Add([string]$t); Write-Host $t }
function Fim { $saida | Set-Content -Path $log -Encoding UTF8; Write-Host ""; Write-Host ("Log: " + $log) }

L "=== CRIACAO DO BANCO - CRM Comercial Zapromaq ==="
L ("Data..: " + (Get-Date -Format 'dd/MM/yyyy HH:mm:ss') + "   " + $env:COMPUTERNAME + "\" + $env:USERNAME)
L ("Banco.: " + $Banco)
if (-not [Environment]::Is64BitProcess) { L "ERRO: rode em PowerShell 64 bits."; Fim; exit 1 }

if (Test-Path $Banco) {
    if ($Recriar) {
        Remove-Item $Banco -Force
        L "banco anterior removido (-Recriar)"
    } else {
        L "ERRO: o banco ja existe. Use -Recriar para substituir, ou aponte outro caminho."
        Fim; exit 1
    }
}

$provider = $null
foreach ($prov in @('Microsoft.ACE.OLEDB.16.0','Microsoft.ACE.OLEDB.12.0')) {
    try {
        $cat = New-Object -ComObject ADOX.Catalog
        $null = $cat.Create("Provider=$prov;Data Source=$Banco;")
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($cat) | Out-Null
        $cat = $null; $provider = $prov; break
    } catch { }
}
if (-not $provider) { L "ERRO: nenhum provedor ACE disponivel."; Fim; exit 1 }
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
L ("provedor: " + $provider)

$cn = New-Object System.Data.OleDb.OleDbConnection("Provider=$provider;Data Source=$Banco;Persist Security Info=False;")
$cn.Open()
function Exec($sql) { $cmd = $cn.CreateCommand(); $cmd.CommandText = $sql; $null = $cmd.ExecuteNonQuery(); $cmd.Dispose() }

L ""
L "--- tabelas ---"

Exec ("CREATE TABLE clientes (" +
 " id AUTOINCREMENT PRIMARY KEY," +
 " codigo_cliente LONG," +
 " estagio TEXT(12) NOT NULL," +
 " empresa TEXT(120) NOT NULL," +
 " cnpj TEXT(20)," +
 " cidade TEXT(60)," +
 " uf TEXT(2)," +
 " segmento TEXT(60)," +
 " telefone TEXT(30)," +
 " email TEXT(120)," +
 " responsavel TEXT(60)," +
 " aderencia INTEGER," +
 " porte INTEGER," +
 " qualificacao TEXT(40)," +
 " observacoes MEMO," +
 " ctx_resumo TEXT(255)," +
 " ctx_familia TEXT(60)," +
 " ctx_arquivo TEXT(255)," +
 " ctx_atualizado_em DATETIME," +
 " ctx_atualizado_por TEXT(50)," +
 " ativo YESNO," +
 " versao LONG," +
 " criado_em DATETIME," +
 " alterado_em DATETIME," +
 " alterado_por TEXT(50))")
L "  clientes"

Exec ("CREATE TABLE contatos (" +
 " id AUTOINCREMENT PRIMARY KEY," +
 " codigo TEXT(15) NOT NULL," +
 " controle LONG NOT NULL," +
 " id_cliente LONG NOT NULL," +
 " nome TEXT(120) NOT NULL," +
 " cargo TEXT(60)," +
 " telefone TEXT(30)," +
 " email TEXT(120)," +
 " observacoes MEMO," +
 " ativo YESNO," +
 " versao LONG," +
 " criado_em DATETIME," +
 " alterado_em DATETIME," +
 " alterado_por TEXT(50))")
L "  contatos"

Exec ("CREATE TABLE oportunidades (" +
 " id AUTOINCREMENT PRIMARY KEY," +
 " codigo TEXT(15) NOT NULL," +
 " controle LONG NOT NULL," +
 " id_contato LONG NOT NULL," +
 " id_cliente LONG NOT NULL," +
 " id_atendimento_anterior LONG," +
 " orcamento TEXT(30)," +
 " etapa TEXT(40) NOT NULL," +
 " responsavel TEXT(60)," +
 " maquina TEXT(255)," +
 " familia TEXT(60)," +
 " categoria TEXT(60)," +
 " tipo_venda TEXT(40)," +
 " valor CURRENCY," +
 " origem TEXT(40)," +
 " prioridade TEXT(20)," +
 " dt_entrada DATETIME," +
 " dt_proposta DATETIME," +
 " ultima_interacao DATETIME," +
 " tentativas INTEGER," +
 " prox_acao MEMO," +
 " dt_prox_acao DATETIME," +
 " retomar_em DATETIME," +
 " motivo_desfecho TEXT(60)," +
 " dt_desfecho DATETIME," +
 " observacoes MEMO," +
 " ctx_resumo TEXT(255)," +
 " ctx_arquivo TEXT(255)," +
 " ctx_atualizado_em DATETIME," +
 " ctx_atualizado_por TEXT(50)," +
 " versao LONG," +
 " criado_em DATETIME," +
 " alterado_em DATETIME," +
 " alterado_por TEXT(50))")
L "  oportunidades"

Exec ("CREATE TABLE listas (" +
 " id AUTOINCREMENT PRIMARY KEY," +
 " tipo TEXT(30) NOT NULL," +
 " valor TEXT(120) NOT NULL," +
 " ordem INTEGER," +
 " ativo YESNO)")
L "  listas"

Exec ("CREATE TABLE metas (" +
 " id AUTOINCREMENT PRIMARY KEY," +
 " ano INTEGER NOT NULL," +
 " mes INTEGER NOT NULL," +
 " meta_faturamento CURRENCY," +
 " meta_pedidos INTEGER," +
 " meta_propostas INTEGER," +
 " meta_contatos INTEGER," +
 " alterado_em DATETIME," +
 " alterado_por TEXT(50))")
L "  metas"

Exec ("CREATE TABLE log_alteracoes (" +
 " id AUTOINCREMENT PRIMARY KEY," +
 " tabela TEXT(30)," +
 " id_registro LONG," +
 " campo TEXT(40)," +
 " valor_antigo MEMO," +
 " valor_novo MEMO," +
 " acao TEXT(20)," +
 " usuario TEXT(50)," +
 " quando DATETIME," +
 " origem TEXT(20))")
L "  log_alteracoes"

Exec ("CREATE TABLE config (" +
 " id AUTOINCREMENT PRIMARY KEY," +
 " chave TEXT(40) NOT NULL," +
 " valor TEXT(255))")
L "  config"

L ""
L "--- indices ---"
$indices = @(
 "CREATE UNIQUE INDEX ix_cli_codigo ON clientes (codigo_cliente)",
 "CREATE INDEX ix_cli_empresa ON clientes (empresa)",
 "CREATE INDEX ix_cli_cnpj ON clientes (cnpj)",
 "CREATE INDEX ix_cli_estagio ON clientes (estagio)",
 "CREATE UNIQUE INDEX ix_cto_codigo ON contatos (codigo)",
 "CREATE INDEX ix_cto_cliente ON contatos (id_cliente)",
 "CREATE INDEX ix_cto_nome ON contatos (nome)",
 "CREATE UNIQUE INDEX ix_opo_codigo ON oportunidades (codigo)",
 "CREATE INDEX ix_opo_contato ON oportunidades (id_contato)",
 "CREATE INDEX ix_opo_cliente ON oportunidades (id_cliente)",
 "CREATE INDEX ix_opo_etapa ON oportunidades (etapa)",
 "CREATE INDEX ix_opo_proxacao ON oportunidades (dt_prox_acao)",
 "CREATE INDEX ix_opo_orcamento ON oportunidades (orcamento)",
 "CREATE INDEX ix_lis_tipo ON listas (tipo)",
 "CREATE UNIQUE INDEX ix_met_anomes ON metas (ano, mes)",
 "CREATE UNIQUE INDEX ix_cfg_chave ON config (chave)"
)
foreach ($ix in $indices) {
    try { Exec $ix; L ("  OK   " + ($ix -replace '^CREATE (UNIQUE )?INDEX (\w+) ON .*$', '$2')) }
    catch { L ("  FALHA " + $ix + " -> " + $_.Exception.Message.Split([char]13)[0]) }
}

L ""
L "--- configuracao ---"
Exec ("INSERT INTO config (chave, valor) VALUES ('versao_esquema','1.0')")
Exec ("INSERT INTO config (chave, valor) VALUES ('criado_em','" + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + "')")
Exec ("INSERT INTO config (chave, valor) VALUES ('ambiente','" + $Ambiente + "')")
L ("  versao_esquema = 1.0   ambiente = " + $Ambiente)

$cn.Close(); $cn.Dispose()
L ""
L ("Banco criado: " + [math]::Round((Get-Item $Banco).Length/1KB) + " KB")
L "Proximo passo: carregar.ps1"
Fim
