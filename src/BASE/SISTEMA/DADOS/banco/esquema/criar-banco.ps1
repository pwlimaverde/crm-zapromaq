<#
  criar-banco.ps1 - cria o .accdb com toda a estrutura do CRM Zapromaq (esquema 1.0).

  Fonte única da estrutura do banco. Não carrega dados de negócio.
  Esquema 1.0 = o banco de produção de 17/09/2026: criar-banco + aplicar-integridade
  + aplicar-indice-cnpj do legado (ferramentas\migracao-3.6), com os MESMOS nomes de
  índice e de chave estrangeira - migrações futuras se referem a eles.
  Alteração de estrutura depois de criado NÃO se faz aqui: vira uma migração
  numerada em banco\migracoes e sobe config.versao_esquema.

  Requisitos: Windows PowerShell 64 bits + motor ACE (Office 64 bits ou
  Access Database Engine 2016 Redistributable x64).

  Uso: criar-banco.ps1 -Banco <caminho.accdb> [-Ambiente TESTE|PRODUCAO] [-Recriar]
#>
param(
    [Parameter(Mandatory = $true)][string]$Banco,
    [ValidateSet('TESTE', 'PRODUCAO')][string]$Ambiente = 'TESTE',
    [switch]$Recriar
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Data

function L([string]$t) { Write-Host $t }

if (-not [Environment]::Is64BitProcess) { throw 'Rode no PowerShell 64 bits.' }

if (Test-Path -LiteralPath $Banco) {
    if (-not $Recriar) { throw ('O banco já existe: ' + $Banco + '. Use -Recriar para substituir.') }
    Remove-Item -LiteralPath $Banco -Force
    L 'banco anterior removido (-Recriar)'
}

# Jet.OLEDB.4.0 é PROIBIDO: cria arquivo JET4 com extensão .accdb, sem erro.
$provider = $null
foreach ($prov in @('Microsoft.ACE.OLEDB.16.0', 'Microsoft.ACE.OLEDB.12.0')) {
    try {
        $cat = New-Object -ComObject ADOX.Catalog
        $null = $cat.Create("Provider=$prov;Data Source=$Banco;")
        $cat.ActiveConnection.Close()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($cat)
        $cat = $null; $provider = $prov; break
    } catch { }
}
if (-not $provider) { throw 'Nenhum provedor ACE disponível (Microsoft.ACE.OLEDB.16.0 / 12.0).' }
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
L ('provedor: ' + $provider)

$cn = New-Object System.Data.OleDb.OleDbConnection("Provider=$provider;Data Source=$Banco;OLE DB Services=-4;")
$cn.Open()
try {
    function Exec([string]$sql) { $cmd = $cn.CreateCommand(); $cmd.CommandText = $sql; [void]$cmd.ExecuteNonQuery(); $cmd.Dispose() }

    Exec ('CREATE TABLE clientes (id AUTOINCREMENT PRIMARY KEY, codigo_cliente LONG, estagio TEXT(12) NOT NULL,' +
          ' empresa TEXT(120) NOT NULL, cnpj TEXT(20), cidade TEXT(60), uf TEXT(2), segmento TEXT(60), telefone TEXT(30),' +
          ' email TEXT(120), responsavel TEXT(60), aderencia INTEGER, porte INTEGER, qualificacao TEXT(40), observacoes MEMO,' +
          ' ctx_resumo TEXT(255), ctx_familia TEXT(60), ctx_arquivo TEXT(255), ctx_atualizado_em DATETIME, ctx_atualizado_por TEXT(50),' +
          ' ativo YESNO, versao LONG, criado_em DATETIME, alterado_em DATETIME, alterado_por TEXT(50))')
    Exec ('CREATE TABLE contatos (id AUTOINCREMENT PRIMARY KEY, codigo TEXT(15) NOT NULL, controle LONG NOT NULL,' +
          ' id_cliente LONG NOT NULL, nome TEXT(120) NOT NULL, cargo TEXT(60), telefone TEXT(30), email TEXT(120), observacoes MEMO,' +
          ' ativo YESNO, versao LONG, criado_em DATETIME, alterado_em DATETIME, alterado_por TEXT(50))')
    Exec ('CREATE TABLE oportunidades (id AUTOINCREMENT PRIMARY KEY, codigo TEXT(15) NOT NULL, controle LONG NOT NULL,' +
          ' id_contato LONG NOT NULL, id_cliente LONG NOT NULL, id_atendimento_anterior LONG, orcamento TEXT(30),' +
          ' etapa TEXT(40) NOT NULL, responsavel TEXT(60), maquina TEXT(255), familia TEXT(60), categoria TEXT(60),' +
          ' tipo_venda TEXT(40), valor CURRENCY, origem TEXT(40), prioridade TEXT(20), dt_entrada DATETIME, dt_proposta DATETIME,' +
          ' ultima_interacao DATETIME, tentativas INTEGER, prox_acao MEMO, dt_prox_acao DATETIME, retomar_em DATETIME,' +
          ' motivo_desfecho TEXT(60), dt_desfecho DATETIME, observacoes MEMO, ctx_resumo TEXT(255), ctx_arquivo TEXT(255),' +
          ' ctx_atualizado_em DATETIME, ctx_atualizado_por TEXT(50), versao LONG, criado_em DATETIME, alterado_em DATETIME,' +
          ' alterado_por TEXT(50))')
    Exec 'CREATE TABLE listas (id AUTOINCREMENT PRIMARY KEY, tipo TEXT(30) NOT NULL, valor TEXT(120) NOT NULL, ordem INTEGER, ativo YESNO)'
    Exec ('CREATE TABLE metas (id AUTOINCREMENT PRIMARY KEY, ano INTEGER NOT NULL, mes INTEGER NOT NULL, meta_faturamento CURRENCY,' +
          ' meta_pedidos INTEGER, meta_propostas INTEGER, meta_contatos INTEGER, alterado_em DATETIME, alterado_por TEXT(50))')
    Exec ('CREATE TABLE log_alteracoes (id AUTOINCREMENT PRIMARY KEY, tabela TEXT(30), id_registro LONG, campo TEXT(40),' +
          ' valor_antigo MEMO, valor_novo MEMO, acao TEXT(20), usuario TEXT(50), quando DATETIME, origem TEXT(20))')
    Exec 'CREATE TABLE config (id AUTOINCREMENT PRIMARY KEY, chave TEXT(40) NOT NULL, valor TEXT(255))'
    L 'tabelas: clientes, contatos, oportunidades, listas, metas, log_alteracoes, config'

    foreach ($ix in @(
        'CREATE UNIQUE INDEX ix_cli_codigo ON clientes (codigo_cliente)',
        'CREATE INDEX ix_cli_empresa ON clientes (empresa)',
        'CREATE INDEX ix_cli_cnpj ON clientes (cnpj)',
        'CREATE UNIQUE INDEX idx_cli_cnpj ON clientes (cnpj)',
        'CREATE INDEX ix_cli_estagio ON clientes (estagio)',
        'CREATE UNIQUE INDEX ix_cto_codigo ON contatos (codigo)',
        'CREATE INDEX ix_cto_cliente ON contatos (id_cliente)',
        'CREATE INDEX ix_cto_nome ON contatos (nome)',
        'CREATE UNIQUE INDEX ix_opo_codigo ON oportunidades (codigo)',
        'CREATE INDEX ix_opo_contato ON oportunidades (id_contato)',
        'CREATE INDEX ix_opo_cliente ON oportunidades (id_cliente)',
        'CREATE INDEX ix_opo_etapa ON oportunidades (etapa)',
        'CREATE INDEX ix_opo_proxacao ON oportunidades (dt_prox_acao)',
        'CREATE INDEX ix_opo_orcamento ON oportunidades (orcamento)',
        'CREATE INDEX ix_lis_tipo ON listas (tipo)',
        'CREATE UNIQUE INDEX ix_met_anomes ON metas (ano, mes)',
        'CREATE UNIQUE INDEX ix_cfg_chave ON config (chave)')) { Exec $ix }
    L 'índices: 17 (inclui idx_cli_cnpj único, aplicado em produção em 17/09/2026)'

    # integridade referencial: não se apaga cliente com contato, nem contato com atendimento
    Exec 'ALTER TABLE contatos ADD CONSTRAINT fk_cto_cli FOREIGN KEY (id_cliente) REFERENCES clientes (id)'
    Exec 'ALTER TABLE oportunidades ADD CONSTRAINT fk_opo_cto FOREIGN KEY (id_contato) REFERENCES contatos (id)'
    Exec 'ALTER TABLE oportunidades ADD CONSTRAINT fk_opo_cli FOREIGN KEY (id_cliente) REFERENCES clientes (id)'
    L 'integridade referencial: 3 chaves estrangeiras'

    $cmd = $cn.CreateCommand()
    $cmd.CommandText = 'INSERT INTO config (chave, valor) VALUES (?, ?)'
    $p1 = $cmd.Parameters.Add('p1', [System.Data.OleDb.OleDbType]::VarWChar, 40)
    $p2 = $cmd.Parameters.Add('p2', [System.Data.OleDb.OleDbType]::VarWChar, 255)
    foreach ($par in @(@('versao_esquema', '1.0'), @('ambiente', $Ambiente), @('criado_em', (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')))) {
        $p1.Value = $par[0]; $p2.Value = $par[1]; [void]$cmd.ExecuteNonQuery()
    }
    $cmd.Dispose()
    L ('config: versao_esquema = 1.0, ambiente = ' + $Ambiente)
} finally { $cn.Close(); $cn.Dispose() }

L ('banco criado: ' + $Banco)
