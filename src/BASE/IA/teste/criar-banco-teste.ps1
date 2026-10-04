<#
  criar-banco-teste.ps1 - cria BASE\crm_zapromaq.accdb com dados FICTÍCIOS
  para testar o servidor MCP no Claude Desktop.

  A estrutura vem do dono dela (SISTEMA\DADOS\banco\esquema\criar-banco.ps1);
  este script só acrescenta listas e registros de exemplo. Nenhum dado real.
#>
param([switch]$Recriar)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Data

$raizIA   = Split-Path -Parent $PSScriptRoot
$raizBase = Split-Path -Parent $raizIA
$banco    = Join-Path $raizBase 'crm_zapromaq.accdb'
$esquema  = Join-Path $raizBase 'SISTEMA\DADOS\banco\esquema\criar-banco.ps1'

if ((Test-Path $banco) -and -not $Recriar) {
    throw ('Já existe um banco em ' + $banco + '. Rode com -Recriar para apagar e gerar de novo.')
}
& $esquema -Banco $banco -Ambiente TESTE -Recriar:$Recriar

$prov = @('Microsoft.ACE.OLEDB.16.0', 'Microsoft.ACE.OLEDB.12.0') |
        Where-Object { @((New-Object System.Data.OleDb.OleDbEnumerator).GetElements() | ForEach-Object { $_.SOURCES_NAME }) -contains $_ } |
        Select-Object -First 1
$cn = New-Object System.Data.OleDb.OleDbConnection("Provider=$prov;Data Source=$banco;OLE DB Services=-4;")
$cn.Open()

function Ins([string]$tabela, $campos) {
    $cmd = $cn.CreateCommand()
    $nomes = @($campos.Keys)
    $cmd.CommandText = 'INSERT INTO ' + $tabela + ' (' + ($nomes -join ', ') + ') VALUES (' + ((@('?') * $nomes.Count) -join ', ') + ')'
    foreach ($n in $nomes) {
        $v = $campos[$n]; $par = $cmd.CreateParameter()
        if ($v -is [datetime])   { $par.OleDbType = [System.Data.OleDb.OleDbType]::Date }
        elseif ($v -is [bool])   { $par.OleDbType = [System.Data.OleDb.OleDbType]::Boolean }
        elseif ($v -is [decimal]) { $par.OleDbType = [System.Data.OleDb.OleDbType]::Currency }
        elseif ($v -is [int])    { $par.OleDbType = [System.Data.OleDb.OleDbType]::Integer }
        else                     { $par.OleDbType = [System.Data.OleDb.OleDbType]::LongVarWChar }
        if ($null -eq $v) { $par.Value = [DBNull]::Value } else { $par.Value = $v }
        [void]$cmd.Parameters.Add($par)
    }
    [void]$cmd.ExecuteNonQuery()
    $cmd.CommandText = 'SELECT @@IDENTITY'; $cmd.Parameters.Clear()
    $id = [int]$cmd.ExecuteScalar(); $cmd.Dispose()
    return $id
}

$hoje = (Get-Date).Date
$agora = Get-Date

try {
    # ------------------------------------------------------------ listas (valores do CRM 3.6)
    $listas = [ordered]@{
        Etapa = @('Contato Inicial', 'Sem Retorno', 'Retorno Agendado', 'Descartado', 'Levantamento Técnico',
                  'Elaboração da Proposta', 'Negociação', 'Proposta em Análise', 'Proposta em Stand By', 'Pedido Fechado', 'Perdido')
        Familia = @('A Definir', 'Calandra de Chapas', 'Calandra de Tubos / Perfis', 'Conformação de Tubos',
                    'Corte de Chapas à Laser', 'Corte de Tubos à Laser e Serra', 'Curvadora de Tubos / Perfis', 'Dobra de Chapas', 'Solda Laser')
        Segmento = @('Moveleiro', 'Automotivo', 'Agrícola', 'Metalurgia', 'Refrigeração', 'Construção / Estrutural')
        UF = @('PR', 'RS', 'SC', 'SP', 'MG')
        Origem = @('Telefone', 'WhatsApp', 'E-mail', 'Site', 'Feira', 'Indicação', 'Prospecção')
        Responsavel = @('Vendedor A', 'Vendedor B', 'Vendedor C')
        Prioridade = @('Alta', 'Média', 'Baixa')
        Motivo = @('Preço', 'Prazo de entrega', 'Projeto adiado', 'Comprou concorrente nacional', 'Sem verba', 'Ganho por preço')
        Categoria = @('Linha própria – nacionalização acima de 85%', 'Linha própria – nacionalização abaixo de 85%')
    }
    foreach ($tipo in $listas.Keys) {
        $i = 0
        foreach ($v in $listas[$tipo]) { $i++; [void](Ins 'listas' ([ordered]@{ tipo = $tipo; valor = $v; ordem = $i; ativo = $true })) }
    }

    # ------------------------------------------------------------ clientes fictícios
    $base = [ordered]@{ versao = 1; criado_em = $agora; alterado_em = $agora; alterado_por = 'teste'; ativo = $true }
    function NovoCliente($cod, $empresa, $cidade, $uf, $seg, $resp) {
        $c = [ordered]@{ codigo_cliente = $cod; estagio = $(if ($cod) { 'Cliente' } else { 'Pré-cliente' }); empresa = $empresa
                         cidade = $cidade; uf = $uf; segmento = $seg; responsavel = $resp; aderencia = 2; porte = 2 }
        foreach ($k in $base.Keys) { $c[$k] = $base[$k] }
        return (Ins 'clientes' $c)
    }
    $c1 = NovoCliente 1 'METALURGICA EXEMPLO LTDA' 'CURITIBA' 'PR' 'Metalurgia' 'Vendedor A'
    $c2 = NovoCliente 2 "INDUSTRIA D'ANGELO DE MOVEIS LTDA" 'BENTO GONCALVES' 'RS' 'Moveleiro' 'Vendedor B'
    $c3 = NovoCliente 3 'AGRO IMPLEMENTOS FICTICIA SA' 'CASCAVEL' 'PR' 'Agrícola' 'Vendedor A'
    $c4 = NovoCliente 4 'REFRIGERACAO MODELO EIRELI' 'JOINVILLE' 'SC' 'Refrigeração' 'Vendedor C'
    $p1 = NovoCliente $null 'ESTRUTURAS PROSPECTO LTDA' 'SAO PAULO' 'SP' 'Construção / Estrutural' 'Vendedor B'
    [void](NovoCliente $null 'AUTOPECAS FILA DE QUALIFICACAO LTDA' 'BELO HORIZONTE' 'MG' 'Automotivo' 'Vendedor C')

    # ------------------------------------------------------------ contatos
    function Cto($ctrl, $idCli, $codCli, $nome, $cargo) {
        $c = [ordered]@{ codigo = ('CT-{0:0000}-{1:0000}' -f $codCli, $ctrl); controle = $ctrl; id_cliente = $idCli
                         nome = $nome; cargo = $cargo; telefone = '41999990000'; email = 'contato@exemplo.com.br' }
        foreach ($k in $base.Keys) { $c[$k] = $base[$k] }
        return (Ins 'contatos' $c)
    }
    $t1 = Cto 1 $c1 1 'João da Silva' 'Gerente industrial'
    $t2 = Cto 2 $c2 2 'Maria Souza' 'Compradora'
    $t3 = Cto 3 $c3 3 'Pedro Almeida' 'Diretor'
    $t4 = Cto 4 $c4 4 'Ana Lima' 'Engenharia'

    # ------------------------------------------------------------ atendimentos (datas relativas a hoje)
    function Opo($ctrl, $idCto, $idCli, $codCli, $etapa, $resp, $familia, $valor, $prox, $proxAcao, $extra) {
        $o = [ordered]@{ codigo = ('AT-{0:0000}-{1:0000}' -f $codCli, $ctrl); controle = $ctrl; id_contato = $idCto; id_cliente = $idCli
                         etapa = $etapa; responsavel = $resp; familia = $familia; valor = $valor; origem = 'Telefone'; prioridade = 'Média'
                         dt_entrada = $hoje.AddDays(-60); ultima_interacao = $hoje.AddDays(-10); tentativas = 1
                         dt_prox_acao = $prox; prox_acao = $proxAcao }
        if ($extra) { foreach ($k in $extra.Keys) { $o[$k] = $extra[$k] } }
        foreach ($k in @('versao', 'criado_em', 'alterado_em', 'alterado_por')) { $o[$k] = $base[$k] }
        [void](Ins 'oportunidades' $o)
    }
    Opo 1 $t1 $c1 1 'Negociação' 'Vendedor A' 'Dobra de Chapas' ([decimal]850000) $hoje.AddDays(-3) 'Ligar para fechar condição de pagamento' @{ maquina = 'DOBRADEIRA CNC 3M'; dt_proposta = $hoje.AddDays(-30) }
    Opo 2 $t1 $c1 1 'Contato Inicial' 'Vendedor A' 'A Definir' $null $hoje 'Enviar catálogo' $null
    Opo 3 $t2 $c2 2 'Proposta em Análise' 'Vendedor B' 'Curvadora de Tubos / Perfis' ([decimal]420000.50) $hoje.AddDays(20) 'Retornar sobre a proposta' @{ maquina = 'CURVADORA CNC 50MM'; dt_proposta = $hoje.AddDays(-15) }
    Opo 4 $t3 $c3 3 'Proposta em Stand By' 'Vendedor A' 'Corte de Tubos à Laser e Serra' ([decimal]1250000) $null $null @{ retomar_em = $hoje.AddDays(90); dt_proposta = $hoje.AddDays(-45) }
    Opo 5 $t4 $c4 4 'Perdido' 'Vendedor C' 'Solda Laser' ([decimal]300000) $null $null @{ motivo_desfecho = 'Preço'; dt_desfecho = $hoje.AddDays(-5); dt_proposta = $hoje.AddDays(-40) }

    # config esperada pelo front
    $cmd = $cn.CreateCommand(); $cmd.CommandText = "INSERT INTO config (chave, valor) VALUES ('versao_front', '1.4')"
    [void]$cmd.ExecuteNonQuery(); $cmd.Dispose()
} finally { $cn.Close(); $cn.Dispose() }

Write-Host ''
Write-Host 'Banco de teste pronto:' $banco
Write-Host '  4 clientes, 2 pré-clientes, 4 contatos, 5 atendimentos, listas preenchidas. Tudo fictício.'
