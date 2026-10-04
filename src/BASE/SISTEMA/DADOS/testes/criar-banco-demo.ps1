<#
  criar-banco-demo.ps1 - cria BASE\crm_zapromaq.accdb com dados FICTÍCIOS para
  demonstração e conferência visual do front: 20 empresas (16 clientes e 4
  pré-clientes), 20 contatos, 20 atendimentos em etapas e prazos variados e as
  metas do ano corrente.

  A estrutura vem do dono dela (banco\esquema\criar-banco.ps1); as migrações
  pendentes são aplicadas no fim, como no banco de produção. Nenhum dado real:
  nomes, CNPJs (dígito verificador válido, raízes inventadas), telefones e
  e-mails são inventados.

  Para os testes do MCP use IA\teste\criar-banco-teste.ps1 (conjunto menor e
  fixo, do qual os testes dependem).

  Uso: criar-banco-demo.ps1 [-Banco <accdb>] [-Recriar]
#>
param([string]$Banco = '', [switch]$Recriar)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Data

$dados = Split-Path -Parent $PSScriptRoot
$raizBase = Split-Path -Parent (Split-Path -Parent $dados)
if ([string]::IsNullOrWhiteSpace($Banco)) { $Banco = Join-Path $raizBase 'crm_zapromaq.accdb' }
$esquema = Join-Path $dados 'banco\esquema\criar-banco.ps1'
$migracoes = Join-Path $dados 'banco\aplicar-migracoes.ps1'

if ((Test-Path $Banco) -and -not $Recriar) {
    throw ('Já existe um banco em ' + $Banco + '. Rode com -Recriar para apagar e gerar de novo.')
}
& $esquema -Banco $Banco -Ambiente TESTE -Recriar:$Recriar

$prov = @('Microsoft.ACE.OLEDB.16.0', 'Microsoft.ACE.OLEDB.12.0') |
        Where-Object { @((New-Object System.Data.OleDb.OleDbEnumerator).GetElements() | ForEach-Object { $_.SOURCES_NAME }) -contains $_ } |
        Select-Object -First 1
$cn = New-Object System.Data.OleDb.OleDbConnection("Provider=$prov;Data Source=$Banco;OLE DB Services=-4;")
$cn.Open()
$tx = $cn.BeginTransaction()

function Ins([string]$tabela, $campos) {
    $cmd = $cn.CreateCommand(); $cmd.Transaction = $tx
    $nomes = @($campos.Keys)
    $cmd.CommandText = 'INSERT INTO ' + $tabela + ' (' + ($nomes -join ', ') + ') VALUES (' + ((@('?') * $nomes.Count) -join ', ') + ')'
    foreach ($n in $nomes) {
        $v = $campos[$n]; $par = $cmd.CreateParameter()
        if ($v -is [datetime])    { $par.OleDbType = [System.Data.OleDb.OleDbType]::Date }
        elseif ($v -is [bool])    { $par.OleDbType = [System.Data.OleDb.OleDbType]::Boolean }
        elseif ($v -is [decimal]) { $par.OleDbType = [System.Data.OleDb.OleDbType]::Currency }
        elseif ($v -is [int])     { $par.OleDbType = [System.Data.OleDb.OleDbType]::Integer }
        else                      { $par.OleDbType = [System.Data.OleDb.OleDbType]::LongVarWChar }
        if ($null -eq $v) { $par.Value = [DBNull]::Value } else { $par.Value = $v }
        [void]$cmd.Parameters.Add($par)
    }
    [void]$cmd.ExecuteNonQuery()
    $cmd.CommandText = 'SELECT @@IDENTITY'; $cmd.Parameters.Clear()
    $id = [int]$cmd.ExecuteScalar(); $cmd.Dispose()
    return $id
}

# CNPJ com dígitos verificadores válidos a partir de uma raiz de 12 dígitos
function Cnpj([string]$base12) {
    $p1 = 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2; $p2 = 6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2
    $s = 0; for ($i = 0; $i -lt 12; $i++) { $s += [int][string]$base12[$i] * $p1[$i] }
    $d1 = 11 - ($s % 11); if ($d1 -ge 10) { $d1 = 0 }
    $b13 = $base12 + $d1
    $s = 0; for ($i = 0; $i -lt 13; $i++) { $s += [int][string]$b13[$i] * $p2[$i] }
    $d2 = 11 - ($s % 11); if ($d2 -ge 10) { $d2 = 0 }
    return $b13 + $d2
}

$hoje = (Get-Date).Date
$agora = Get-Date
$base = [ordered]@{ ativo = $true; versao = 1; criado_em = $agora; alterado_em = $agora; alterado_por = 'demo' }

try {
    # ------------------------------------------------------------ listas
    $listas = [ordered]@{
        Etapa = @('Contato Inicial', 'Sem Retorno', 'Retorno Agendado', 'Descartado', 'Levantamento Técnico',
                  'Elaboração da Proposta', 'Negociação', 'Proposta em Análise', 'Proposta em Stand By', 'Pedido Fechado', 'Perdido')
        Familia = @('A Definir', 'Calandra de Chapas', 'Calandra de Tubos / Perfis', 'Conformação de Tubos',
                    'Corte de Chapas à Laser', 'Corte de Tubos à Laser e Serra', 'Curvadora de Tubos / Perfis', 'Dobra de Chapas', 'Solda Laser')
        Segmento = @('Moveleiro', 'Automotivo', 'Agrícola', 'Metalurgia', 'Refrigeração', 'Construção / Estrutural', 'Fitness', 'Hospitalar')
        UF = @('PR', 'RS', 'SC', 'SP', 'MG', 'GO', 'MS')
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

    # ------------------------------------------------------------ 20 empresas: 16 clientes + 4 pré-clientes
    # empresa | cidade | uf | segmento | responsável | aderência | porte | qualificação
    $empresas = @(
        @('METALURGICA EXEMPLO LTDA',            'CURITIBA',         'PR', 'Metalurgia',              'Vendedor A', 3, 3, 'Prioridade alta'),
        @("INDUSTRIA D'ANGELO DE MOVEIS LTDA",   'BENTO GONCALVES',  'RS', 'Moveleiro',               'Vendedor B', 3, 2, 'Prioridade alta'),
        @('AGRO IMPLEMENTOS FICTICIA SA',        'CASCAVEL',         'PR', 'Agrícola',                'Vendedor A', 2, 3, 'Fila padrão'),
        @('REFRIGERACAO MODELO EIRELI',          'JOINVILLE',        'SC', 'Refrigeração',            'Vendedor C', 2, 2, 'Fila padrão'),
        @('AUTOPECAS EXEMPLO LTDA',              'SAO JOSE DOS PINHAIS', 'PR', 'Automotivo',          'Vendedor B', 3, 3, 'Prioridade alta'),
        @('CALDEIRARIA FICTICIA LTDA',           'CAXIAS DO SUL',    'RS', 'Metalurgia',              'Vendedor C', 2, 2, 'Reserva'),
        @('SERRALHERIA MODELO ME',               'BLUMENAU',         'SC', 'Construção / Estrutural', 'Vendedor A', 1, 1, 'Reserva'),
        @('MOVEIS TESTE INDUSTRIA LTDA',         'ARAPONGAS',        'PR', 'Moveleiro',               'Vendedor B', 2, 2, 'Fila padrão'),
        @('TUBOS DEMONSTRACAO SA',               'SOROCABA',         'SP', 'Metalurgia',              'Vendedor C', 3, 3, 'Prioridade alta'),
        @('ESTAMPARIA EXEMPLO LTDA',             'CONTAGEM',         'MG', 'Automotivo',              'Vendedor A', 2, 2, 'Fila padrão'),
        @('EQUIPAMENTOS FICTICIOS LTDA',         'CHAPECO',          'SC', 'Agrícola',                'Vendedor B', 2, 1, 'Fila padrão'),
        @('FITNESS MODELO LTDA',                 'LONDRINA',         'PR', 'Fitness',                 'Vendedor C', 2, 2, 'Reserva'),
        @('HOSPITALAR TESTE SA',                 'RIBEIRAO PRETO',   'SP', 'Hospitalar',              'Vendedor A', 1, 2, 'Reserva'),
        @('IMPLEMENTOS EXEMPLO LTDA',            'RIO VERDE',        'GO', 'Agrícola',                'Vendedor B', 3, 2, 'Prioridade alta'),
        @('DEMONSTRACAO METAL LTDA',             'CAMPO GRANDE',     'MS', 'Metalurgia',              'Vendedor C', 2, 2, 'Fila padrão'),
        @('ESTRUTURAS METALICAS MODELO LTDA',    'MARINGA',          'PR', 'Construção / Estrutural', 'Vendedor A', 3, 3, 'Prioridade alta'),
        @('ESTRUTURAS PROSPECTO LTDA',           'SAO PAULO',        'SP', 'Construção / Estrutural', 'Vendedor B', 2, 2, 'Fila padrão'),
        @('AUTOPECAS FILA DE QUALIFICACAO LTDA', 'BELO HORIZONTE',   'MG', 'Automotivo',              'Vendedor C', 1, 1, 'Insuficiente para qualificar'),
        @('INOX FICTICIO COMERCIO LTDA',         'NOVO HAMBURGO',    'RS', 'Metalurgia',              'Vendedor A', 2, 1, 'Reserva'),
        @('CARROCERIAS PROSPECTO LTDA',          'ERECHIM',          'RS', 'Automotivo',              'Vendedor B', 3, 2, 'Prioridade alta')
    )
    $ids = @{}; $cods = @{}
    for ($i = 0; $i -lt $empresas.Count; $i++) {
        $e = $empresas[$i]; $cod = if ($i -lt 16) { $i + 1 } else { $null }
        $c = [ordered]@{
            codigo_cliente = $cod; estagio = $(if ($cod) { 'Cliente' } else { 'Pré-cliente' })
            empresa = $e[0]; cnpj = (Cnpj ('{0:00}{1:000}{2:000}0001' -f @((11 + $i), ((100 + 37 * $i) % 1000), ((200 + 53 * $i) % 1000))))
            cidade = $e[1]; uf = $e[2]; segmento = $e[3]
            telefone = ('4133{0:000000}' -f (100000 + 4321 * $i)); email = ('comercial{0:00}@exemplo.com.br' -f ($i + 1))
            responsavel = $e[4]; aderencia = [int]$e[5]; porte = [int]$e[6]; qualificacao = $e[7]
            observacoes = $(if ($i % 4 -eq 0) { 'Cadastro fictício para demonstração do CRM.' } else { $null })
        }
        foreach ($k in $base.Keys) { $c[$k] = $base[$k] }
        $ids[$i] = Ins 'clientes' $c; $cods[$i] = $cod
    }

    # ------------------------------------------------------------ 20 contatos (só clientes têm código CT)
    $pessoas = @(
        @('João da Silva', 'Gerente industrial'), @('Maria Souza', 'Compradora'), @('Pedro Almeida', 'Diretor'),
        @('Ana Lima', 'Engenharia'), @('Carlos Pereira', 'Supervisor de produção'), @('Fernanda Costa', 'Compras'),
        @('Ricardo Gomes', 'Sócio'), @('Juliana Rocha', 'Engenheira de processos'), @('Marcos Ribeiro', 'Gerente de manutenção'),
        @('Patrícia Martins', 'Diretora'), @('Lucas Carvalho', 'Comprador'), @('Camila Teixeira', 'Planejamento'),
        @('Roberto Araújo', 'Proprietário'), @('Beatriz Fernandes', 'Engenharia'), @('Eduardo Barbosa', 'Gerente geral'),
        @('Gabriela Melo', 'Compras'), @('Felipe Cardoso', 'Encarregado'), @('Larissa Dias', 'Engenharia de produto'),
        @('Thiago Nunes', 'Diretor industrial'), @('Renata Moreira', 'Compradora')
    )
    # cliente (índice 0..15) de cada contato: os 16 primeiros um por cliente, mais 4 segundos contatos
    $donos = @(0..15) + @(0, 1, 4, 8)
    $ctos = @{}; $ctrlPorCli = @{}
    for ($i = 0; $i -lt 20; $i++) {
        $ci = $donos[$i]; $ctrlPorCli[$ci] = 1 + [int]$ctrlPorCli[$ci]
        $ctrl = $i + 1
        $c = [ordered]@{
            codigo = ('CT-{0:0000}-{1:0000}' -f $cods[$ci], $ctrl); controle = $ctrl; id_cliente = $ids[$ci]
            nome = $pessoas[$i][0]; cargo = $pessoas[$i][1]
            telefone = ('41999{0:000000}' -f (10000 + 777 * $i)); email = ('contato{0:00}@exemplo.com.br' -f ($i + 1))
        }
        foreach ($k in $base.Keys) { $c[$k] = $base[$k] }
        $ctos[$i] = @{ id = (Ins 'contatos' $c); cli = $ci }
    }

    # ------------------------------------------------------------ 20 atendimentos com situações variadas
    # contato | etapa | família | valor | próx. ação (dias) | texto | extras
    $atend = @(
        @(0,  'Negociação',             'Dobra de Chapas',                850000,   -3, 'Ligar para fechar condição de pagamento', @{ maquina = 'DOBRADEIRA CNC 3M'; dt_proposta = -30 }),
        @(0,  'Contato Inicial',        'A Definir',                      $null,     0, 'Enviar catálogo', @{}),
        @(1,  'Proposta em Análise',    'Curvadora de Tubos / Perfis',    420000.5, 20, 'Retornar sobre a proposta', @{ maquina = 'CURVADORA CNC 50MM'; dt_proposta = -15 }),
        @(2,  'Proposta em Stand By',   'Corte de Tubos à Laser e Serra', 1250000, $null, $null, @{ retomar_em = 90; dt_proposta = -45 }),
        @(3,  'Perdido',                'Solda Laser',                    300000,  $null, $null, @{ motivo_desfecho = 'Preço'; dt_desfecho = -5; dt_proposta = -40 }),
        @(4,  'Pedido Fechado',         'Corte de Chapas à Laser',        980000,  $null, $null, @{ maquina = 'LASER FIBRA 6KW'; dt_proposta = -50; dt_desfecho = -8; motivo_desfecho = 'Ganho por preço' }),
        @(5,  'Levantamento Técnico',   'Calandra de Chapas',             $null,     2, 'Visita técnica na fábrica', @{}),
        @(6,  'Sem Retorno',            'A Definir',                      $null,    -12, 'Nova tentativa por WhatsApp', @{ tentativas = 3 }),
        @(7,  'Elaboração da Proposta', 'Dobra de Chapas',                510000,    5, 'Enviar proposta revisada', @{ maquina = 'DOBRADEIRA 100T' }),
        @(8,  'Retorno Agendado',       'Conformação de Tubos',           $null,     1, 'Ligação agendada com a engenharia', @{}),
        @(9,  'Negociação',             'Calandra de Tubos / Perfis',     265000,   -1, 'Aguardar contraproposta', @{ maquina = 'CALANDRA 3 ROLOS'; dt_proposta = -20 }),
        @(10, 'Descartado',             'A Definir',                      $null,   $null, $null, @{ motivo_desfecho = 'Sem verba'; dt_desfecho = -30 }),
        @(11, 'Proposta em Análise',    'Solda Laser',                    189000,   12, 'Confirmar recebimento da proposta', @{ maquina = 'SOLDA LASER PORTATIL'; dt_proposta = -7 }),
        @(12, 'Contato Inicial',        'A Definir',                      $null,     7, 'Apresentar a linha de corte', @{}),
        @(13, 'Pedido Fechado',         'Curvadora de Tubos / Perfis',    345000,  $null, $null, @{ maquina = 'CURVADORA CNC 76MM'; dt_proposta = -70; dt_desfecho = -20; motivo_desfecho = 'Ganho por preço' }),
        @(14, 'Perdido',                'Corte de Chapas à Laser',        720000,  $null, $null, @{ motivo_desfecho = 'Comprou concorrente nacional'; dt_desfecho = -15; dt_proposta = -60 }),
        @(15, 'Negociação',             'Corte de Tubos à Laser e Serra', 1480000,   3, 'Reunião com a diretoria', @{ maquina = 'LASER TUBOS 3KW'; dt_proposta = -25 }),
        @(16, 'Proposta em Stand By',   'Dobra de Chapas',                390000,  $null, $null, @{ retomar_em = 3; dt_proposta = -100 }),
        @(17, 'Levantamento Técnico',   'Conformação de Tubos',           $null,    -6, 'Receber desenhos das peças', @{}),
        @(18, 'Elaboração da Proposta', 'Calandra de Chapas',             610000,   10, 'Fechar configuração da máquina', @{ maquina = 'CALANDRA 4 ROLOS 3M' }),
        @(19, 'Contato Inicial',        'A Definir',                      $null,    -2, 'Retornar contato do site', @{})
    )
    $origens = $listas.Origem; $prioridades = $listas.Prioridade; $responsaveis = $listas.Responsavel
    for ($i = 0; $i -lt 20; $i++) {
        $a = $atend[$i]; $ct = $ctos[$a[0]]; $ci = $ct.cli; $ctrl = $i + 1
        $o = [ordered]@{
            codigo = ('AT-{0:0000}-{1:0000}' -f $cods[$ci], $ctrl); controle = $ctrl; id_contato = $ct.id; id_cliente = $ids[$ci]
            orcamento = $(if ($null -ne $a[3]) { 'ORC-{0:000}/26' -f (100 + $i) } else { $null })
            etapa = $a[1]; responsavel = $responsaveis[$i % 3]; familia = $a[2]
            categoria = $listas.Categoria[$i % 2]; tipo_venda = $(if ($i % 3 -eq 0) { 'Revenda' } else { 'Venda direta' })
            valor = $(if ($null -ne $a[3]) { [decimal]$a[3] } else { $null })
            origem = $origens[$i % $origens.Count]; prioridade = $prioridades[$i % 3]
            dt_entrada = $hoje.AddDays(-(15 + 7 * $i)); ultima_interacao = $hoje.AddDays(-($i % 9)); tentativas = 1 + ($i % 3)
            dt_prox_acao = $(if ($null -ne $a[4]) { $hoje.AddDays($a[4]) } else { $null }); prox_acao = $a[5]
        }
        foreach ($k in $a[6].Keys) {
            $v = $a[6][$k]
            if ($k -like 'dt_*' -or $k -eq 'retomar_em') { $v = $hoje.AddDays($v) }
            $o[$k] = $v
        }
        foreach ($k in @('versao', 'criado_em', 'alterado_em', 'alterado_por')) { $o[$k] = $base[$k] }
        [void](Ins 'oportunidades' $o)
    }

    # ------------------------------------------------------------ metas do ano corrente
    for ($m = 1; $m -le 12; $m++) {
        [void](Ins 'metas' ([ordered]@{ ano = $hoje.Year; mes = $m; meta_faturamento = [decimal]1500000; meta_pedidos = 2
                                       meta_propostas = 6; meta_contatos = 20; alterado_em = $agora; alterado_por = 'demo' }))
    }

    $cmd = $cn.CreateCommand(); $cmd.Transaction = $tx
    $cmd.CommandText = "INSERT INTO config (chave, valor) VALUES ('versao_front', ?)"
    $par = $cmd.CreateParameter(); $par.OleDbType = [System.Data.OleDb.OleDbType]::VarWChar
    $par.Value = (Get-Content -LiteralPath (Join-Path $dados 'VERSAO.txt') -TotalCount 1).Trim(); [void]$cmd.Parameters.Add($par)
    [void]$cmd.ExecuteNonQuery(); $cmd.Dispose()
    $tx.Commit()
} catch { $tx.Rollback(); throw }
finally { $cn.Close(); $cn.Dispose() }

& $migracoes -Banco $Banco

Write-Host ''
Write-Host 'Banco de demonstração pronto:' $Banco
Write-Host '  16 clientes, 4 pré-clientes, 20 contatos, 20 atendimentos, metas do ano. Tudo fictício.'
