<#
  Ferramentas.ps1 - catálogo FECHADO de operações que a IA pode executar.

  A IA não manda SQL: ela só enxerga estas ferramentas. Toda gravação
    - confere a versão do esquema do banco;
    - confere a versão do registro (controle otimista), igual ao front;
    - normaliza com as mesmas regras do front;
    - grava dado + log_alteracoes (origem = 'IA') na mesma transação.
  Acrescentar operação = acrescentar uma entrada em $script:Ferramentas.
#>

$script:TamanhoTexto = @{
    empresa = 120; cidade = 60; uf = 2; segmento = 60; telefone = 30; email = 120; cnpj = 20
    responsavel = 60; qualificacao = 40; maquina = 255; familia = 60; categoria = 60; tipo_venda = 40
    origem = 40; prioridade = 20; orcamento = 30; motivo_desfecho = 60; etapa = 40
    ctx_resumo = 255; ctx_familia = 60; ctx_arquivo = 255; nome = 120; cargo = 60
}

# ------------------------------------------------------------------ utilidades
function Test-TemArg($a, [string]$nome) {
    return ($null -ne $a -and ($a.PSObject.Properties.Name -contains $nome))
}
function Get-Arg($a, [string]$nome, $padrao = $null) {
    if (Test-TemArg $a $nome) { return $a.$nome }
    return $padrao
}
function Get-Limite($a, [int]$padrao = 20, [int]$maximo = 200) {
    $n = [int](Get-Arg $a 'limite' $padrao)
    if ($n -lt 1) { $n = 1 }
    if ($n -gt $maximo) { $n = $maximo }
    return $n
}

function Get-ValoresLista($ctx, [string]$tipo) {
    $sql = 'SELECT valor FROM listas WHERE tipo=? AND ativo=True ORDER BY ordem, valor'
    $linhas = if ($ctx) { $ctx.Consultar($sql, @((P 'texto' $tipo))) } else { Invoke-Consulta $sql @((P 'texto' $tipo)) }
    return @($linhas | ForEach-Object { [string]$_.valor })
}

# Devolve o valor exatamente como está na lista (acerta maiúsculas); lança se não existir.
function Resolve-ValorLista($ctx, [string]$tipo, [string]$campo, $valor) {
    if ($null -eq $valor -or [string]$valor -eq '') { return $null }
    $validos = Get-ValoresLista $ctx $tipo
    foreach ($v in $validos) { if ($v -ieq ([string]$valor).Trim()) { return $v } }
    throw ("Campo '" + $campo + "': '" + $valor + "' não está na lista " + $tipo + ". Valores aceitos: " + ($validos -join ' | '))
}

function Assert-Tamanho([string]$campo, $valor) {
    if ($null -eq $valor) { return }
    $max = $script:TamanhoTexto[$campo]
    if ($max -and ([string]$valor).Length -gt $max) {
        throw ("Campo '" + $campo + "' aceita no máximo " + $max + ' caracteres (recebido: ' + ([string]$valor).Length + ').')
    }
}

function Add-Observacao($atual, [string]$texto) {
    $linha = '[' + (Get-Date).ToString('dd/MM/yyyy') + ' ' + $script:UsuarioLog + '] ' + $texto.Trim()
    if ($null -eq $atual -or [string]$atual -eq '') { return $linha }
    return ([string]$atual).TrimEnd() + "`r`n" + $linha
}

function Test-Igual($a, $b) {
    if (($null -eq $a -or [string]$a -eq '') -and ($null -eq $b -or [string]$b -eq '')) { return $true }
    if ($null -eq $a -or $null -eq $b) { return $false }
    if ($a -is [datetime] -or $b -is [datetime]) {
        return ((ConvertTo-DataOuNulo $a) -eq (ConvertTo-DataOuNulo $b))
    }
    if (($a -is [double] -or $a -is [int] -or $a -is [decimal]) -or ($b -is [double] -or $b -is [int] -or $b -is [decimal])) {
        return ([double]$a -eq [double]$b)
    }
    return ([string]$a -ceq [string]$b)
}

<#
  UPDATE com controle de versão, dentro da transação $ctx.
  $mudancas: ordered @{ campo = @{ Tipo; Novo; Antigo } } - só o que mudou.
#>
function Update-ComVersao($ctx, [string]$tabela, [int]$id, [int]$versaoLida, $mudancas, [string]$acao) {
    $sets = New-Object System.Collections.ArrayList
    $params = New-Object System.Collections.ArrayList
    foreach ($campo in $mudancas.Keys) {
        [void]$sets.Add($campo + '=?')
        [void]$params.Add((P $mudancas[$campo].Tipo $mudancas[$campo].Novo))
    }
    $sql = 'UPDATE ' + $tabela + ' SET ' + ($sets -join ', ') +
           ', versao=versao+1, alterado_em=?, alterado_por=? WHERE id=? AND versao=?'
    [void]$params.Add((P 'data' (Get-Date)))
    [void]$params.Add((P 'texto' $script:UsuarioLog))
    [void]$params.Add((P 'inteiro' $id))
    [void]$params.Add((P 'inteiro' $versaoLida))
    $n = $ctx.Executar($sql, $params.ToArray())
    if ($n -ne 1) {
        throw ('Conflito: o registro foi alterado por outra pessoa enquanto a IA preparava a mudança. ' +
               'Nada foi gravado. Leia o registro de novo e refaça.')
    }
    foreach ($campo in $mudancas.Keys) {
        $ctx.Log($tabela, $id, $campo, $mudancas[$campo].Antigo, $mudancas[$campo].Novo, $acao)
    }
}

function Assert-Versao($atual, $a) {
    $lida = [int](Get-Arg $a 'versao')
    if ([int]$atual.versao -ne $lida) {
        throw ('Conflito de versão: a IA leu a versão ' + $lida + ', mas o registro já está na versão ' + $atual.versao +
               ' (alguém alterou depois da leitura). Nada foi gravado. Leia de novo e confira antes de repetir.')
    }
}

function Get-ClienteNaTransacao($ctx, $a) {
    if (Test-TemArg $a 'codigo_cliente') {
        $l = $ctx.Consultar('SELECT * FROM clientes WHERE codigo_cliente=?', @((P 'inteiro' ([int]$a.codigo_cliente))))
        $ref = 'código ' + $a.codigo_cliente
    } elseif (Test-TemArg $a 'id_cliente') {
        $l = $ctx.Consultar('SELECT * FROM clientes WHERE id=?', @((P 'inteiro' ([int]$a.id_cliente))))
        $ref = 'id ' + $a.id_cliente
    } else { throw 'Informe codigo_cliente (cliente) ou id_cliente (pré-cliente, que não tem código).' }
    if ($l.Count -eq 0) { throw ('Cliente não encontrado: ' + $ref + '.') }
    return $l[0]
}

function Get-AtendimentoNaTransacao($ctx, [string]$codigo) {
    if ($codigo -notmatch '^AT-\d{4}-\d{4}$') { throw ("Código de atendimento inválido: '" + $codigo + "'. Formato: AT-0000-0000.") }
    $l = $ctx.Consultar('SELECT * FROM oportunidades WHERE codigo=?', @((P 'texto' $codigo)))
    if ($l.Count -eq 0) { throw ('Atendimento não encontrado: ' + $codigo + '.') }
    return $l[0]
}

function Add-Calculados($o) {
    $o['situacao']    = Get-Situacao ($null -ne $o.id_contato) ([string]$o.etapa) $o.dt_prox_acao $o.retomar_em
    $o['quadro']      = Get-Quadro ([string]$o.etapa)
    $o['dias_parado'] = Get-DiasParado $o.ultima_interacao
    $o['ciclo']       = Get-Ciclo $o.dt_proposta $o.dt_desfecho
    return $o
}

# ================================================================= LEITURA
function Invoke-CrmStatus($a) {
    $cont = [ordered]@{}
    foreach ($t in @('clientes', 'contatos', 'oportunidades', 'listas', 'metas', 'log_alteracoes')) {
        $cont[$t] = [int](Invoke-Valor ('SELECT COUNT(*) FROM ' + $t) @())
    }
    $esq = Get-VersaoEsquema
    return [ordered]@{
        banco               = $script:CaminhoBanco
        provedor            = Get-ProvedorAce
        versao_esquema      = $esq
        gravacao_liberada   = ($script:EsquemasSuportados -contains $esq)
        versao_front        = Invoke-Valor "SELECT valor FROM config WHERE chave='versao_front'" @()
        ambiente            = Invoke-Valor "SELECT valor FROM config WHERE chave='ambiente'" @()
        registros           = $cont
        ultimo_log_id       = Invoke-Valor 'SELECT MAX(id) FROM log_alteracoes' @()
        versao_servidor_mcp = $script:VersaoServidor
    }
}

function Invoke-ListarOpcoes($a) {
    $tipo = [string](Get-Arg $a 'tipo' '')
    if ($tipo -eq '') {
        $tipos = Invoke-Consulta 'SELECT DISTINCT tipo FROM listas ORDER BY tipo' @()
        return [ordered]@{ tipos = @($tipos | ForEach-Object { $_.tipo }) }
    }
    return [ordered]@{ tipo = $tipo; valores = @(Get-ValoresLista $null $tipo) }
}

function Invoke-BuscarClientes($a) {
    $n = Get-Limite $a
    $w = New-Object System.Collections.ArrayList; $p = New-Object System.Collections.ArrayList
    $texto = [string](Get-Arg $a 'texto' '')
    if ($texto.Trim() -ne '') {
        $like = '%' + $texto.Trim() + '%'
        $dig = Get-SoDigitos $texto
        if ($texto.Trim() -match '^\d{1,4}$') {
            [void]$w.Add('(codigo_cliente=? OR empresa LIKE ?)')
            [void]$p.Add((P 'inteiro' ([int]$texto))); [void]$p.Add((P 'texto' $like))
        } else {
            $cond = '(empresa LIKE ? OR cidade LIKE ? OR segmento LIKE ?'
            [void]$p.Add((P 'texto' $like)); [void]$p.Add((P 'texto' $like)); [void]$p.Add((P 'texto' $like))
            if ($dig.Length -ge 4) { $cond += ' OR cnpj LIKE ?'; [void]$p.Add((P 'texto' ('%' + $dig + '%'))) }
            [void]$w.Add($cond + ')')
        }
    }
    if (Test-TemArg $a 'estagio')     { [void]$w.Add('estagio=?');     [void]$p.Add((P 'texto' $a.estagio)) }
    if (Test-TemArg $a 'responsavel') { [void]$w.Add('responsavel=?'); [void]$p.Add((P 'texto' $a.responsavel)) }
    if (Test-TemArg $a 'uf')          { [void]$w.Add('uf=?');          [void]$p.Add((P 'texto' ([string]$a.uf).ToUpper())) }
    if (-not [bool](Get-Arg $a 'incluir_inativos' $false)) { [void]$w.Add('ativo=True') }
    $sql = 'SELECT TOP ' + $n + ' id, codigo_cliente, estagio, empresa, cnpj, cidade, uf, segmento, responsavel,' +
           ' aderencia, porte, qualificacao, ativo, ctx_resumo, versao FROM clientes'
    if ($w.Count -gt 0) { $sql += ' WHERE ' + ($w -join ' AND ') }
    $sql += ' ORDER BY IIF(codigo_cliente IS NULL, 1, 0), codigo_cliente, empresa, id'
    $linhas = Invoke-Consulta $sql $p.ToArray()
    foreach ($l in $linhas) { $l['codigo'] = Format-CodigoCliente $l.codigo_cliente }
    return [ordered]@{ quantidade = $linhas.Count; limite = $n; clientes = $linhas }
}

function Invoke-ObterCliente($a) {
    if (Test-TemArg $a 'codigo_cliente') {
        $l = Invoke-Consulta 'SELECT * FROM clientes WHERE codigo_cliente=?' @((P 'inteiro' ([int]$a.codigo_cliente)))
    } elseif (Test-TemArg $a 'id_cliente') {
        $l = Invoke-Consulta 'SELECT * FROM clientes WHERE id=?' @((P 'inteiro' ([int]$a.id_cliente)))
    } else { throw 'Informe codigo_cliente ou id_cliente.' }
    if ($l.Count -eq 0) { throw 'Cliente não encontrado.' }
    $c = $l[0]
    $c['codigo'] = Format-CodigoCliente $c.codigo_cliente
    $c['contatos'] = Invoke-Consulta ('SELECT id, codigo, nome, cargo, telefone, email, ativo FROM contatos' +
                                      ' WHERE id_cliente=? ORDER BY codigo') @((P 'inteiro' ([int]$c.id)))
    $ops = Invoke-Consulta ('SELECT o.codigo, o.etapa, o.responsavel, o.valor, o.maquina, o.dt_prox_acao, o.retomar_em,' +
                            ' o.id_contato, ct.nome AS contato FROM oportunidades o INNER JOIN contatos ct ON o.id_contato = ct.id' +
                            ' WHERE o.id_cliente=? ORDER BY o.codigo DESC') @((P 'inteiro' ([int]$c.id)))
    foreach ($o in $ops) { $o['situacao'] = Get-Situacao $true ([string]$o.etapa) $o.dt_prox_acao $o.retomar_em }
    $c['atendimentos'] = $ops
    return $c
}

function Invoke-BuscarAtendimentos($a) {
    $n = Get-Limite $a
    $w = New-Object System.Collections.ArrayList; $p = New-Object System.Collections.ArrayList
    $texto = [string](Get-Arg $a 'texto' '')
    if ($texto.Trim() -ne '') {
        $like = '%' + $texto.Trim() + '%'
        [void]$w.Add('(o.codigo LIKE ? OR c.empresa LIKE ? OR ct.nome LIKE ? OR o.maquina LIKE ? OR o.orcamento LIKE ?)')
        1..5 | ForEach-Object { [void]$p.Add((P 'texto' $like)) }
    }
    if (Test-TemArg $a 'etapa')       { [void]$w.Add('o.etapa=?');       [void]$p.Add((P 'texto' $a.etapa)) }
    if (Test-TemArg $a 'responsavel') { [void]$w.Add('o.responsavel=?'); [void]$p.Add((P 'texto' $a.responsavel)) }
    if (Test-TemArg $a 'codigo_cliente') { [void]$w.Add('c.codigo_cliente=?'); [void]$p.Add((P 'inteiro' ([int]$a.codigo_cliente))) }
    if ([bool](Get-Arg $a 'somente_abertos' $true)) {
        [void]$w.Add("o.etapa NOT IN ('Pedido Fechado', 'Perdido', 'Descartado')")
    }
    $filtroSituacao = [string](Get-Arg $a 'situacao' '')
    # Com filtro de situação (calculada fora do banco) lê o conjunto e corta depois.
    $top = if ($filtroSituacao -ne '') { '' } else { 'TOP ' + $n + ' ' }
    $sql = 'SELECT ' + $top + 'o.id, o.codigo, o.etapa, o.responsavel, o.valor, o.maquina, o.familia, o.prioridade,' +
           ' o.dt_prox_acao, o.retomar_em, o.ultima_interacao, o.prox_acao, o.id_contato, o.versao,' +
           ' c.codigo_cliente, c.empresa, ct.nome AS contato' +
           ' FROM (oportunidades o INNER JOIN clientes c ON o.id_cliente = c.id)' +
           ' INNER JOIN contatos ct ON o.id_contato = ct.id'
    if ($w.Count -gt 0) { $sql += ' WHERE ' + ($w -join ' AND ') }
    # data nula vai por último: no Access o ORDER BY põe nulo na frente
    $sql += ' ORDER BY IIF(o.dt_prox_acao IS NULL, 1, 0), o.dt_prox_acao, o.codigo DESC'
    $linhas = Invoke-Consulta $sql $p.ToArray()
    $saida = New-Object System.Collections.ArrayList
    foreach ($l in $linhas) {
        $l['situacao'] = Get-Situacao $true ([string]$l.etapa) $l.dt_prox_acao $l.retomar_em
        if ($filtroSituacao -ne '' -and $l.situacao -ne $filtroSituacao) { continue }
        [void]$saida.Add($l)
        if ($saida.Count -ge $n) { break }
    }
    return [ordered]@{ quantidade = $saida.Count; limite = $n; atendimentos = $saida }
}

function Invoke-ObterAtendimento($a) {
    $codigo = [string](Get-Arg $a 'codigo' '')
    if ($codigo -notmatch '^AT-\d{4}-\d{4}$') { throw 'Código inválido. Formato: AT-0000-0000.' }
    $l = Invoke-Consulta ('SELECT o.*, c.codigo_cliente, c.empresa, c.cidade, c.uf, ct.codigo AS codigo_contato, ct.nome AS contato,' +
                          ' ct.cargo AS contato_cargo, ct.telefone AS contato_telefone, ct.email AS contato_email' +
                          ' FROM (oportunidades o INNER JOIN clientes c ON o.id_cliente = c.id)' +
                          ' INNER JOIN contatos ct ON o.id_contato = ct.id WHERE o.codigo=?') @((P 'texto' $codigo))
    if ($l.Count -eq 0) { throw ('Atendimento não encontrado: ' + $codigo) }
    return (Add-Calculados $l[0])
}

function Invoke-UltimasAlteracoes($a) {
    $n = Get-Limite $a 30 200
    $w = New-Object System.Collections.ArrayList; $p = New-Object System.Collections.ArrayList
    if (Test-TemArg $a 'tabela')      { [void]$w.Add('tabela=?');      [void]$p.Add((P 'texto' $a.tabela)) }
    if (Test-TemArg $a 'id_registro') { [void]$w.Add('id_registro=?'); [void]$p.Add((P 'inteiro' ([int]$a.id_registro))) }
    if (Test-TemArg $a 'desde_id')    { [void]$w.Add('id>?');          [void]$p.Add((P 'inteiro' ([int]$a.desde_id))) }
    $sql = 'SELECT TOP ' + $n + ' id, tabela, id_registro, campo, valor_antigo, valor_novo, acao, usuario, quando, origem FROM log_alteracoes'
    if ($w.Count -gt 0) { $sql += ' WHERE ' + ($w -join ' AND ') }
    $sql += ' ORDER BY id DESC'
    $linhas = Invoke-Consulta $sql $p.ToArray()
    return [ordered]@{ quantidade = $linhas.Count; alteracoes = $linhas }
}

# ================================================================= RELATÓRIO SEMANAL
# atendimentos_por_interacao: atendimentos (abertos e encerrados) cuja
# ultima_interacao cai num período, numa única chamada, para o relatório
# semanal. Devolve JSON compacto e sem campos vazios (monta o próprio
# '__conteudo'; as demais ferramentas continuam com o JSON indentado).

# Teto do texto devolvido, em caracteres do JSON compacto. Acima dele a lista
# é cortada com truncado=true e a quantidade real: o chamador refaz por
# responsável. Valor de projeto, ajustável aqui.
$script:LimiteCaracteresInteracao = 60000

$script:CamposInteracaoPadrao = @(
    'codigo', 'codigo_cliente', 'empresa', 'cidade', 'uf', 'contato', 'contato_cargo', 'contato_telefone',
    'contato_email', 'responsavel', 'etapa', 'situacao', 'origem', 'prioridade', 'familia', 'maquina',
    'tipo_venda', 'valor', 'orcamento', 'dt_entrada', 'dt_proposta', 'ultima_interacao', 'tentativas',
    'prox_acao', 'dt_prox_acao', 'retomar_em', 'motivo_desfecho', 'dt_desfecho', 'observacoes',
    'ctx_resumo', 'ctx_atualizado_em', 'ctx_atualizado_por', 'alterado_em', 'alterado_por')

# Segunda-feira da semana ISO 8601 (a semana 1 é a que contém 4 de janeiro).
function Get-SegundaIso([int]$semana, [int]$ano) {
    $jan4 = New-Object DateTime($ano, 1, 4)
    $dia = [int]$jan4.DayOfWeek; if ($dia -eq 0) { $dia = 7 }
    $segunda = $jan4.AddDays(1 - $dia).AddDays(7 * ($semana - 1))
    # a quinta-feira define o ano da semana: sem ela no ano, a semana não existe
    if ($segunda.AddDays(3).Year -ne $ano) {
        throw ('O ano ' + $ano + ' não tem a semana ISO ' + $semana + ' (tem 52 semanas). Nada foi consultado.')
    }
    return $segunda
}

function Get-SemanaIso([datetime]$data) {
    $dia = [int]$data.DayOfWeek; if ($dia -eq 0) { $dia = 7 }
    $quinta = $data.Date.AddDays(4 - $dia)
    return @{ semana = ([int][math]::Floor(($quinta.DayOfYear - 1) / 7) + 1); ano = $quinta.Year }
}

function Get-PeriodoInteracao($a) {
    $temSemana = (Test-TemArg $a 'semana') -or (Test-TemArg $a 'ano')
    $temDatas  = (Test-TemArg $a 'data_inicio') -or (Test-TemArg $a 'data_fim')
    if ($temSemana -and $temDatas) {
        throw 'Informe semana/ano OU data_inicio/data_fim, não os dois. Nada foi consultado.'
    }
    if ($temSemana) {
        if (-not ((Test-TemArg $a 'semana') -and (Test-TemArg $a 'ano'))) {
            throw 'semana e ano vão juntos: informe os dois. Nada foi consultado.'
        }
        $sem = [int]$a.semana; $ano = [int]$a.ano
        if ($sem -lt 1 -or $sem -gt 53) { throw ('Semana ' + $sem + ' fora de 1 a 53. Nada foi consultado.') }
        $ini = Get-SegundaIso $sem $ano
        $fim = $ini.AddDays(6)
    } elseif ($temDatas) {
        if (-not ((Test-TemArg $a 'data_inicio') -and (Test-TemArg $a 'data_fim'))) {
            throw 'data_inicio e data_fim vão juntas: informe as duas. Nada foi consultado.'
        }
        $ini = ConvertTo-Data 'data_inicio' $a.data_inicio
        $fim = ConvertTo-Data 'data_fim' $a.data_fim
        if ($fim -lt $ini) {
            throw ('data_fim (' + $a.data_fim + ') é anterior a data_inicio (' + $a.data_inicio + '). Nada foi consultado.')
        }
        $sem = $null; $ano = $null
        # período que coincide com uma semana ISO inteira leva o número dela
        $s = Get-SemanaIso $ini
        if ([int]$ini.DayOfWeek -eq 1 -and ($fim - $ini).TotalDays -eq 6) { $sem = $s.semana; $ano = $s.ano }
    } else {
        # semana ISO imediatamente anterior à data atual
        $hoje = (Get-Date).Date
        $dia = [int]$hoje.DayOfWeek; if ($dia -eq 0) { $dia = 7 }
        $ini = $hoje.AddDays(1 - $dia - 7)
        $fim = $ini.AddDays(6)
        $s = Get-SemanaIso $ini; $sem = $s.semana; $ano = $s.ano
    }
    $p = [ordered]@{}
    if ($null -ne $sem) { $p['semana'] = $sem; $p['ano'] = $ano }
    $p['data_inicio'] = $ini.ToString('yyyy-MM-dd')
    $p['data_fim'] = $fim.ToString('yyyy-MM-dd')
    return @{ Periodo = $p; Inicio = $ini; Fim = $fim }
}

function Test-ValorVazio($v) {
    if ($null -eq $v -or $v -is [System.DBNull]) { return $true }
    if ($v -is [string] -and $v.Trim() -eq '') { return $true }
    return $false
}

function Add-Contagem($dic, [string]$chave) {
    if ($dic.Contains($chave)) { $dic[$chave] = $dic[$chave] + 1 } else { $dic[$chave] = 1 }
}

function Invoke-AtendimentosPorInteracao($a) {
    # validação completa antes de qualquer acesso ao banco
    $per = Get-PeriodoInteracao $a
    $completo = [bool](Get-Arg $a 'completo' $false)
    $resp = [string](Get-Arg $a 'responsavel' '')

    $p = New-Object System.Collections.ArrayList
    [void]$p.Add((P 'data' $per.Inicio))
    [void]$p.Add((P 'data' $per.Fim.AddDays(1)))   # < dia seguinte: inclui o último dia inteiro, com ou sem hora
    # LEFT JOIN: atendimento sem contato ou cliente vinculado também entra (situação vazia, como em obter_atendimento)
    $sql = 'SELECT o.*, c.codigo_cliente, c.empresa, c.cidade, c.uf, ct.codigo AS codigo_contato, ct.nome AS contato,' +
           ' ct.cargo AS contato_cargo, ct.telefone AS contato_telefone, ct.email AS contato_email' +
           ' FROM (oportunidades o LEFT JOIN clientes c ON o.id_cliente = c.id)' +
           ' LEFT JOIN contatos ct ON o.id_contato = ct.id' +
           ' WHERE o.ultima_interacao >= ? AND o.ultima_interacao < ?'
    if ($resp.Trim() -ne '') { $sql += ' AND o.responsavel=?'; [void]$p.Add((P 'texto' $resp.Trim())) }
    $sql += ' ORDER BY o.responsavel, o.ultima_interacao, o.codigo'
    $linhas = Invoke-Consulta $sql $p.ToArray()

    $registros = New-Object System.Collections.ArrayList
    $grupos = [ordered]@{}
    foreach ($l in $linhas) {
        [void](Add-Calculados $l)   # situação: mesma regra do front e de buscar_atendimentos
        $r = [ordered]@{}
        $campos = if ($completo) { @($l.Keys) } else { $script:CamposInteracaoPadrao }
        foreach ($c in $campos) {
            if ($l.Contains($c) -and -not (Test-ValorVazio $l[$c])) { $r[$c] = $l[$c] }
        }
        [void]$registros.Add($r)

        $nomeResp = if (Test-ValorVazio $l['responsavel']) { '(sem responsável)' } else { [string]$l['responsavel'] }
        if (-not $grupos.Contains($nomeResp)) {
            $grupos[$nomeResp] = [ordered]@{ responsavel = $nomeResp; total = 0; por_etapa = [ordered]@{}; por_situacao = [ordered]@{} }
        }
        $g = $grupos[$nomeResp]
        $g['total'] = $g['total'] + 1
        Add-Contagem $g['por_etapa'] $(if (Test-ValorVazio $l['etapa']) { '(sem etapa)' } else { [string]$l['etapa'] })
        Add-Contagem $g['por_situacao'] $(if (Test-ValorVazio $l['situacao']) { '(sem situação)' } else { [string]$l['situacao'] })
    }

    $saida = [ordered]@{ periodo = $per.Periodo; quantidade = $registros.Count }
    $resumo = @($grupos.Values)
    if ($resp.Trim() -ne '' -and $registros.Count -eq 0) {
        $saida['aviso'] = ("Nenhum atendimento de '" + $resp.Trim() + "' no período. Confira a grafia do responsável em listar_opcoes.")
    }

    # corte por tamanho: nunca em silêncio
    $vazio = [ordered]@{}; foreach ($k in $saida.Keys) { $vazio[$k] = $saida[$k] }
    $vazio['truncado'] = $true; $vazio['devolvidos'] = 0; $vazio['resumo'] = $resumo
    $vazio['atendimentos'] = @()
    $disponivel = $script:LimiteCaracteresInteracao - (ConvertTo-Json -InputObject $vazio -Depth 10 -Compress).Length
    $lista = New-Object System.Collections.ArrayList
    $usado = 0
    foreach ($r in $registros) {
        $t = (ConvertTo-Json -InputObject $r -Depth 5 -Compress).Length + 1
        if ($usado + $t -gt $disponivel) { break }
        [void]$lista.Add($r); $usado += $t
    }
    if ($lista.Count -lt $registros.Count) {
        $saida['truncado'] = $true
        $saida['devolvidos'] = $lista.Count
        $saida['aviso'] = ('Resultado maior que o limite da ferramenta: devolvidos ' + $lista.Count + ' de ' + $registros.Count +
                         '. O resumo cobre todos. Refaça a chamada por responsável para obter os demais.')
    }
    $saida['resumo'] = $resumo
    $saida['atendimentos'] = $lista

    $json = ConvertTo-Json -InputObject $saida -Depth 10 -Compress
    return [ordered]@{ __conteudo = [ordered]@{ type = 'text'; text = $json } }
}

# ================================================================= PLANEJAMENTO
# atendimentos_por_proxima_acao: atendimentos EM ABERTO com dt_prox_acao num
# período de hoje em diante, para o relatório semanal de planejamento. O CRM
# sobrescreve a próxima ação quando ela é executada: período passado não tem
# dado confiável e é recusado. Mesmo formato de atendimentos_por_interacao
# (JSON compacto, sem campos vazios, '__conteudo'), mas o corte por tamanho
# pagina (a_partir_de / proximo) em vez de mandar refazer por responsável:
# um mês de um único responsável já passa do teto. Teto próprio, menor que o
# de atendimentos_por_interacao: 60.000 caracteres passam do que o cliente MCP
# mostra numa resposta (testado em 28/09/2026). O resumo vai só na primeira
# página (a_partir_de = 0), para as seguintes levarem mais registros.

$script:LimiteDiasPlanejamento = 92
$script:LimiteCaracteresPlanejamento = 35000
$script:MsgPlanejamentoPassado = ('O planejamento só vale de hoje em diante: o CRM sobrescreve a próxima ação quando ela é executada. ' +
                                 'Para período passado use atendimentos_por_interacao. Nada foi consultado.')
$script:CamposPlanejamentoPadrao = @($script:CamposInteracaoPadrao + @('quadro', 'categoria', 'codigo_contato', 'vencida'))
# em aberto = etapa não encerrada; etapa vazia também é atendimento em aberto
$script:SqlAtendimentoAberto = "(o.etapa IS NULL OR o.etapa NOT IN ('Pedido Fechado', 'Perdido', 'Descartado'))"
$script:QuadroFunil = '2. Funil comercial'

function Get-PeriodoPlanejamento($a) {
    $hoje = (Get-Date).Date
    $temSemana = (Test-TemArg $a 'semana') -or (Test-TemArg $a 'ano')
    $temDatas  = (Test-TemArg $a 'data_inicio') -or (Test-TemArg $a 'data_fim')
    if ($temSemana -and $temDatas) {
        throw 'Informe semana/ano OU data_inicio/data_fim, não os dois. Nada foi consultado.'
    }
    $sem = $null; $ano = $null
    if ($temSemana) {
        if (-not ((Test-TemArg $a 'semana') -and (Test-TemArg $a 'ano'))) {
            throw 'semana e ano vão juntos: informe os dois. Nada foi consultado.'
        }
        $sem = [int]$a.semana; $ano = [int]$a.ano
        if ($sem -lt 1 -or $sem -gt 53) { throw ('Semana ' + $sem + ' fora de 1 a 53. Nada foi consultado.') }
        $ini = Get-SegundaIso $sem $ano
        $fim = $ini.AddDays(6)
    } elseif ($temDatas) {
        if (-not ((Test-TemArg $a 'data_inicio') -and (Test-TemArg $a 'data_fim'))) {
            throw 'data_inicio e data_fim vão juntas: informe as duas. Nada foi consultado.'
        }
        $ini = ConvertTo-Data 'data_inicio' $a.data_inicio
        $fim = ConvertTo-Data 'data_fim' $a.data_fim
        if ($fim -lt $ini) {
            throw ('data_fim (' + $a.data_fim + ') é anterior a data_inicio (' + $a.data_inicio + '). Nada foi consultado.')
        }
    } else {
        # de hoje até o domingo da semana ISO corrente
        $dia = [int]$hoje.DayOfWeek; if ($dia -eq 0) { $dia = 7 }
        $ini = $hoje
        $fim = $hoje.AddDays(7 - $dia)
        $s = Get-SemanaIso $hoje; $sem = $s.semana; $ano = $s.ano
    }
    if ($fim -lt $hoje) { throw $script:MsgPlanejamentoPassado }
    $pedido = $null
    if ($ini -lt $hoje) { $pedido = $ini; $ini = $hoje }
    $dias = [int]($fim - $ini).TotalDays + 1
    if ($dias -gt $script:LimiteDiasPlanejamento) {
        throw ('Período de ' + $dias + ' dias (' + $ini.ToString('yyyy-MM-dd') + ' a ' + $fim.ToString('yyyy-MM-dd') +
               '): o limite é ' + $script:LimiteDiasPlanejamento + ' dias. Divida em períodos menores. Nada foi consultado.')
    }
    # período por datas que coincide com uma semana ISO inteira leva o número dela
    if ($temDatas -and [int]$ini.DayOfWeek -eq 1 -and $dias -eq 7) {
        $s = Get-SemanaIso $ini; $sem = $s.semana; $ano = $s.ano
    }
    $p = [ordered]@{}
    if ($null -ne $sem) { $p['semana'] = $sem; $p['ano'] = $ano }
    $p['data_inicio'] = $ini.ToString('yyyy-MM-dd')
    $p['data_fim'] = $fim.ToString('yyyy-MM-dd')
    $p['hoje'] = $hoje.ToString('yyyy-MM-dd')
    if ($null -ne $pedido) {
        $p['ajustado'] = $true
        $p['data_inicio_pedida'] = $pedido.ToString('yyyy-MM-dd')
    }
    return @{ Periodo = $p; Inicio = $ini; Fim = $fim; Hoje = $hoje }
}

function Get-OrdemPrioridade($v) {
    if (Test-ValorVazio $v) { return 3 }
    $s = ([string]$v).Trim().ToLowerInvariant()
    if ($s.StartsWith('a')) { return 0 }   # Alta
    if ($s.StartsWith('m')) { return 1 }   # Média
    if ($s.StartsWith('b')) { return 2 }   # Baixa
    return 3
}

function New-ResumoPlanejamento([string]$nome) {
    $g = [ordered]@{}
    if ($nome -ne '') { $g['responsavel'] = $nome }
    $g['total'] = 0; $g['vencidas'] = 0; $g['sem_proxima_acao'] = 0; $g['valor_funil'] = 0.0
    $g['por_dia'] = [ordered]@{}; $g['por_etapa'] = [ordered]@{}; $g['por_quadro'] = [ordered]@{}
    return $g
}

# total, por_dia, por_etapa, por_quadro e valor_funil cobrem só o período;
# as vencidas contam à parte, em 'vencidas'
function Add-ResumoPlanejamento($g, $l, [bool]$vencida, [string]$dia) {
    if ($vencida) { $g['vencidas'] = $g['vencidas'] + 1; return }
    $g['total'] = $g['total'] + 1
    Add-Contagem $g['por_dia'] $dia
    Add-Contagem $g['por_etapa'] $(if (Test-ValorVazio $l['etapa']) { '(sem etapa)' } else { [string]$l['etapa'] })
    Add-Contagem $g['por_quadro'] $(if (Test-ValorVazio $l['quadro']) { '(sem quadro)' } else { [string]$l['quadro'] })
    if ($l['quadro'] -eq $script:QuadroFunil -and -not (Test-ValorVazio $l['valor'])) {
        $g['valor_funil'] = $g['valor_funil'] + [double]$l['valor']
    }
}

function Complete-ResumoPlanejamento($g) {
    $dias = [ordered]@{}
    foreach ($k in (@($g['por_dia'].Keys) | Sort-Object)) { $dias[$k] = $g['por_dia'][$k] }
    $g['por_dia'] = $dias
    $g['valor_funil'] = [math]::Round([double]$g['valor_funil'], 2)
}

function Invoke-AtendimentosPorProximaAcao($a) {
    # validação completa antes de qualquer acesso ao banco
    $per = Get-PeriodoPlanejamento $a
    $completo = [bool](Get-Arg $a 'completo' $false)
    $comVencidas = [bool](Get-Arg $a 'incluir_vencidas' $true)
    $resp = ([string](Get-Arg $a 'responsavel' '')).Trim()
    $aPartirDe = [int](Get-Arg $a 'a_partir_de' 0)
    if ($aPartirDe -lt 0) { throw 'a_partir_de não pode ser negativo. Nada foi consultado.' }
    $hoje = $per.Hoje

    # LEFT JOIN: atendimento sem contato ou cliente vinculado também entra
    $p = New-Object System.Collections.ArrayList
    $sql = 'SELECT o.*, c.codigo_cliente, c.empresa, c.cidade, c.uf, ct.codigo AS codigo_contato, ct.nome AS contato,' +
           ' ct.cargo AS contato_cargo, ct.telefone AS contato_telefone, ct.email AS contato_email' +
           ' FROM (oportunidades o LEFT JOIN clientes c ON o.id_cliente = c.id)' +
           ' LEFT JOIN contatos ct ON o.id_contato = ct.id' +
           ' WHERE ' + $script:SqlAtendimentoAberto + ' AND o.dt_prox_acao < ?'
    [void]$p.Add((P 'data' $per.Fim.AddDays(1)))   # < dia seguinte: inclui o último dia inteiro, com ou sem hora
    if ($comVencidas) {
        # o início nunca é anterior a hoje: vencida = antes de hoje; o intervalo entre hoje e o início fica fora
        $sql += ' AND (o.dt_prox_acao < ? OR o.dt_prox_acao >= ?)'
        [void]$p.Add((P 'data' $hoje)); [void]$p.Add((P 'data' $per.Inicio))
    } else {
        $sql += ' AND o.dt_prox_acao >= ?'
        [void]$p.Add((P 'data' $per.Inicio))
    }
    if ($resp -ne '') { $sql += ' AND o.responsavel=?'; [void]$p.Add((P 'texto' $resp)) }
    $linhas = Invoke-Consulta $sql $p.ToArray()

    # atendimentos abertos sem próxima ação: carteira toda, fora do filtro de período
    $p2 = New-Object System.Collections.ArrayList
    $sql2 = 'SELECT o.responsavel, COUNT(*) AS n FROM oportunidades o WHERE ' + $script:SqlAtendimentoAberto +
            ' AND o.dt_prox_acao IS NULL'
    if ($resp -ne '') { $sql2 += ' AND o.responsavel=?'; [void]$p2.Add((P 'texto' $resp)) }
    $sql2 += ' GROUP BY o.responsavel'
    $semProxima = Invoke-Consulta $sql2 $p2.ToArray()

    # ordenação: vencidas, responsável, data da próxima ação, prioridade (Alta, Média, Baixa, vazio), empresa
    $itens = New-Object System.Collections.ArrayList
    foreach ($l in $linhas) {
        [void](Add-Calculados $l)   # situação e quadro: mesma regra do front
        $dt = ConvertTo-DataOuNulo $l['dt_prox_acao']
        $venc = ($dt -lt $hoje)
        if ($venc) { $l['vencida'] = $true }
        $nomeResp = if (Test-ValorVazio $l['responsavel']) { '(sem responsável)' } else { [string]$l['responsavel'] }
        [void]$itens.Add([pscustomobject]@{
            V = $(if ($venc) { 0 } else { 1 }); R = $nomeResp; D = $dt; Pr = (Get-OrdemPrioridade $l['prioridade'])
            E = [string]$l['empresa']; C = [string]$l['codigo']; Venc = $venc; L = $l })
    }
    $ordenados = @($itens | Sort-Object V, R, D, Pr, E, C)

    $grupos = @{}
    $equipe = New-ResumoPlanejamento ''
    $registros = New-Object System.Collections.ArrayList
    foreach ($it in $ordenados) {
        $l = $it.L
        if (-not $grupos.ContainsKey($it.R)) { $grupos[$it.R] = New-ResumoPlanejamento $it.R }
        $dia = $it.D.ToString('yyyy-MM-dd')
        Add-ResumoPlanejamento $grupos[$it.R] $l $it.Venc $dia
        Add-ResumoPlanejamento $equipe $l $it.Venc $dia
        $r = [ordered]@{}
        $campos = if ($completo) { @($l.Keys) } else { $script:CamposPlanejamentoPadrao }
        foreach ($c in $campos) {
            if ($l.Contains($c) -and -not (Test-ValorVazio $l[$c])) { $r[$c] = $l[$c] }
        }
        [void]$registros.Add($r)
    }
    foreach ($s in $semProxima) {
        $nomeResp = if (Test-ValorVazio $s['responsavel']) { '(sem responsável)' } else { [string]$s['responsavel'] }
        if (-not $grupos.ContainsKey($nomeResp)) { $grupos[$nomeResp] = New-ResumoPlanejamento $nomeResp }
        $grupos[$nomeResp]['sem_proxima_acao'] = [int]$s['n']
        $equipe['sem_proxima_acao'] = $equipe['sem_proxima_acao'] + [int]$s['n']
    }
    $resumo = @()
    foreach ($k in (@($grupos.Keys) | Sort-Object)) { Complete-ResumoPlanejamento $grupos[$k]; $resumo += $grupos[$k] }
    Complete-ResumoPlanejamento $equipe

    $total = $registros.Count
    $saida = [ordered]@{ periodo = $per.Periodo; quantidade = $total }
    $avisos = New-Object System.Collections.ArrayList
    if ($resp -ne '' -and $total -eq 0) {
        if ($equipe['sem_proxima_acao'] -gt 0) {
            [void]$avisos.Add("Nenhum atendimento de '" + $resp + "' com próxima ação no período; há " + $equipe['sem_proxima_acao'] + ' em aberto sem próxima ação.')
        } else {
            [void]$avisos.Add("Nenhum atendimento em aberto de '" + $resp + "'. Confira a grafia do responsável em listar_opcoes.")
        }
    }
    if ($aPartirDe -gt 0 -and $aPartirDe -ge $total) {
        [void]$avisos.Add('a_partir_de (' + $aPartirDe + ') está além do último registro (' + $total + '): nada a devolver.')
    }

    # corte por tamanho: nunca em silêncio e sem perder registro - pagina
    $molde = [ordered]@{}; foreach ($k in $saida.Keys) { $molde[$k] = $saida[$k] }
    $molde['a_partir_de'] = $aPartirDe; $molde['devolvidos'] = $total; $molde['truncado'] = $true; $molde['proximo'] = $total
    $molde['aviso'] = (($avisos + @('Resultado maior que o limite da ferramenta: devolvidos 0000 de 0000 a partir de 0000. Repita com a_partir_de=0000 até não vir mais proximo; o resumo de todos vem na primeira página.')) -join ' ')
    $comResumo = ($aPartirDe -eq 0)
    if ($comResumo) { $molde['resumo'] = $resumo; $molde['equipe'] = $equipe }
    $molde['atendimentos'] = @()
    $disponivel = $script:LimiteCaracteresPlanejamento - (ConvertTo-Json -InputObject $molde -Depth 10 -Compress).Length
    $lista = New-Object System.Collections.ArrayList
    $usado = 0
    $i = $aPartirDe
    while ($i -lt $total) {
        $t = (ConvertTo-Json -InputObject $registros[$i] -Depth 5 -Compress).Length + 1
        # um registro sozinho maior que o espaço ainda sai, para a paginação sempre avançar
        if ($lista.Count -gt 0 -and $usado + $t -gt $disponivel) { break }
        [void]$lista.Add($registros[$i]); $usado += $t; $i++
    }
    if ($aPartirDe -gt 0) { $saida['a_partir_de'] = $aPartirDe }
    $saida['devolvidos'] = $lista.Count
    if ($i -lt $total) {
        $saida['truncado'] = $true
        $saida['proximo'] = $i
        [void]$avisos.Add('Resultado maior que o limite da ferramenta: devolvidos ' + $lista.Count + ' de ' + $total +
                          ' a partir de ' + $aPartirDe + '. Repita com a_partir_de=' + $i + ' até não vir mais proximo; o resumo de todos vem na primeira página.')
    }
    if ($avisos.Count -gt 0) { $saida['aviso'] = ($avisos -join ' ') }
    if ($comResumo) {
        $saida['resumo'] = $resumo
        $saida['equipe'] = $equipe
    }
    $saida['atendimentos'] = $lista

    $json = ConvertTo-Json -InputObject $saida -Depth 10 -Compress
    return [ordered]@{ __conteudo = [ordered]@{ type = 'text'; text = $json } }
}

# ================================================================= EXTRATO
# exportar_base: grava a foto das tabelas em CSV numa pasta de trabalho dentro
# da pasta do setor e devolve só o envelope (arquivos, linhas, colunas). É a
# entrada dos scripts de análise: o dado não passa pela conversa. Não grava no
# banco. CSV em UTF-8 sem BOM, separador ',', todo campo entre aspas, CRLF - o
# formato que os scripts do padrão (00 - SCRIPTS\_comum.py) leem sem conversão.
$script:TabelasExtrato = @('clientes', 'contatos', 'oportunidades', 'log_alteracoes', 'listas', 'metas')
$script:TabelasExtratoPadrao = @('clientes', 'contatos', 'oportunidades')
$script:NomeRaizSetor = '1 - COMERCIAL ZAPROMAQ'
$script:PastaLibFerramentas = $PSScriptRoot
$script:SqlExtrato = @{
    clientes       = 'SELECT * FROM clientes ORDER BY id'
    contatos       = ('SELECT ct.*, c.codigo_cliente, c.empresa FROM contatos ct LEFT JOIN clientes c ON ct.id_cliente = c.id' +
                      ' ORDER BY ct.id')
    oportunidades  = ('SELECT o.*, c.codigo_cliente, c.empresa, ct.codigo AS codigo_contato' +
                      ' FROM (oportunidades o LEFT JOIN clientes c ON o.id_cliente = c.id)' +
                      ' LEFT JOIN contatos ct ON o.id_contato = ct.id ORDER BY o.id')
    log_alteracoes = 'SELECT * FROM log_alteracoes ORDER BY id'
    listas         = 'SELECT * FROM listas ORDER BY id'
    metas          = 'SELECT * FROM metas ORDER BY id'
}
$script:CalculadosExtrato = @('situacao', 'quadro', 'dias_parado', 'ciclo')

# Raiz do setor pelo NOME da pasta, subindo a partir do conector: nunca letra de unidade.
function Get-RaizSetor {
    $d = Get-Item -LiteralPath $script:PastaLibFerramentas
    while ($null -ne $d) {
        if ($d.Name -eq $script:NomeRaizSetor) { return $d.FullName }
        $d = $d.Parent
    }
    throw ('Pasta "' + $script:NomeRaizSetor + '" não encontrada acima do conector. Nada foi exportado.')
}

function Resolve-DestinoExtrato([string]$destino) {
    $bruto = $destino.Trim()
    if ($bruto -match '^[A-Za-z]:' -or $bruto.StartsWith('\\') -or $bruto.StartsWith('//')) {
        throw 'destino é relativo à pasta 1 - COMERCIAL ZAPROMAQ: sem letra de unidade e sem \\servidor. Nada foi exportado.'
    }
    $d = $bruto.Replace('/', '\').Trim('\')
    if ($d -eq '') { throw 'destino vazio. Nada foi exportado.' }
    foreach ($parte in ($d -split '\\')) {
        if ($parte -eq '..' -or $parte -eq '.') { throw 'destino não aceita "." nem "..". Nada foi exportado.' }
    }
    if ($d -like '01 - CRM\01 - CONTROLE*') {
        throw 'destino dentro de 01 - CRM\01 - CONTROLE (pasta do sistema) é recusado. Nada foi exportado.'
    }
    $alvo = Join-Path (Get-RaizSetor) $d
    if (-not (Test-Path -LiteralPath $alvo -PathType Container)) {
        throw ('A pasta de destino não existe: ' + $d + '. Crie a pasta antes. Nada foi exportado.')
    }
    return @{ Relativo = $d; Absoluto = $alvo }
}

function Get-ColunasConsulta([string]$sql) {
    $cn = Open-Conexao
    try {
        $cmd = New-Comando $cn $sql @() $null
        try {
            $rd = $cmd.ExecuteReader([System.Data.CommandBehavior]::SchemaOnly)
            try {
                $nomes = New-Object System.Collections.ArrayList
                for ($i = 0; $i -lt $rd.FieldCount; $i++) { [void]$nomes.Add($rd.GetName($i)) }
                return ,$nomes
            } finally { $rd.Close() }
        } finally { $cmd.Dispose() }
    } finally { $cn.Close(); $cn.Dispose() }
}

function ConvertTo-CampoCsv($v) {
    if ($null -eq $v -or $v -is [System.DBNull]) { return '""' }
    if ($v -is [bool]) { $s = $(if ($v) { '1' } else { '0' }) }
    elseif ($v -is [double] -or $v -is [single] -or $v -is [decimal]) {
        $s = ([double]$v).ToString('R', [System.Globalization.CultureInfo]::InvariantCulture)
    }
    else { $s = [string]$v }
    return ('"' + $s.Replace('"', '""') + '"')
}

function Write-ExtratoCsv([string]$arquivo, $linhas, [string[]]$colunas) {
    $tmp = $arquivo + '.tmp'
    $w = New-Object System.IO.StreamWriter($tmp, $false, (New-Object System.Text.UTF8Encoding($false)))
    try {
        $w.NewLine = "`r`n"
        $w.WriteLine((@($colunas | ForEach-Object { ConvertTo-CampoCsv $_ }) -join ','))
        foreach ($l in $linhas) {
            $vals = New-Object System.Collections.ArrayList
            foreach ($c in $colunas) {
                $v = $null; if ($l.Contains($c)) { $v = $l[$c] }
                [void]$vals.Add((ConvertTo-CampoCsv $v))
            }
            $w.WriteLine(($vals -join ','))
        }
    } finally { $w.Close() }
    # troca só depois de escrito inteiro: quem lê nunca pega arquivo pela metade
    if (Test-Path -LiteralPath $arquivo) { Remove-Item -LiteralPath $arquivo -Force }
    Move-Item -LiteralPath $tmp -Destination $arquivo
}

function Invoke-ExportarBase($a) {
    # validação completa antes de qualquer acesso ao banco
    $dest = Resolve-DestinoExtrato ([string]$a.destino)
    $tabelas = @($script:TabelasExtratoPadrao)
    if (Test-TemArg $a 'tabelas') {
        $tabelas = @()
        foreach ($t in @($a.tabelas)) {
            $t = ([string]$t).Trim().ToLower()
            if ($script:TabelasExtrato -notcontains $t) {
                throw ("Tabela desconhecida: '" + $t + "'. Aceitas: " + ($script:TabelasExtrato -join ', ') + '. Nada foi exportado.')
            }
            if ($tabelas -notcontains $t) { $tabelas += $t }
        }
    }

    $gerado = Get-Date
    $saida = [ordered]@{
        gerado_em = $gerado.ToString('yyyy-MM-ddTHH:mm:ss')
        versao_esquema = [string](Get-VersaoEsquema)
        versao_servidor = $script:VersaoServidor
        destino = $dest.Relativo
        separador = ','
        codificacao = 'UTF-8'
        tabelas = [ordered]@{}
    }
    foreach ($t in $tabelas) {
        $linhas = Invoke-Consulta $script:SqlExtrato[$t] @()
        if ($t -eq 'oportunidades') { foreach ($l in $linhas) { [void](Add-Calculados $l) } }
        if ($linhas.Count -gt 0) { $colunas = @($linhas[0].Keys) }
        else {
            $colunas = @(); foreach ($c in (Get-ColunasConsulta $script:SqlExtrato[$t])) { $colunas += [string]$c }
            if ($t -eq 'oportunidades') { $colunas += $script:CalculadosExtrato }
        }
        $nome = 'crm-' + $t + '.csv'
        Write-ExtratoCsv (Join-Path $dest.Absoluto $nome) $linhas ([string[]]$colunas)
        $saida.tabelas[$t] = [ordered]@{ arquivo = $nome; linhas = $linhas.Count; colunas = $colunas.Count }
    }
    $manifesto = Join-Path $dest.Absoluto 'crm-extrato.json'
    [System.IO.File]::WriteAllText($manifesto, (ConvertTo-Json -InputObject $saida -Depth 5), (New-Object System.Text.UTF8Encoding($false)))
    $saida['manifesto'] = 'crm-extrato.json'
    return $saida
}

# ================================================================= GRAVAÇÃO
function Invoke-AtualizarAtendimento($a) {
    Assert-EsquemaGravavel
    $codigo = [string]$a.codigo
    return Invoke-Transacao {
        param($ctx)
        $atual = Get-AtendimentoNaTransacao $ctx $codigo
        Assert-Versao $atual $a
        $novo = [ordered]@{}; foreach ($k in $atual.Keys) { $novo[$k] = $atual[$k] }
        $mud = [ordered]@{}
        foreach ($campo in $script:DefsAtendimento.Keys) {
            if (-not (Test-TemArg $a $campo)) { continue }
            $v = ConvertTo-CampoAtendimento $ctx $campo $a.$campo
            $novo[$campo] = $v
            if (-not (Test-Igual $atual[$campo] $v)) { $mud[$campo] = @{ Tipo = $script:DefsAtendimento[$campo][0]; Novo = $v; Antigo = $atual[$campo] } }
        }
        # troca do contato: só dentro da mesma empresa (o código AT é congelado e nomeia a pasta do cliente)
        if ((Test-TemArg $a 'codigo_contato') -or (Test-TemArg $a 'id_contato')) {
            $ct = Get-ContatoNaTransacao $ctx $a
            if ([int]$ct.id_cliente -ne [int]$atual.id_cliente) { throw ('O contato ' + $ct.codigo + ' é de outra empresa. O atendimento só troca de contato dentro da mesma empresa.') }
            if (-not [bool]$ct.ativo) { throw ('O contato ' + $ct.codigo + ' está inativo.') }
            if ([int]$ct.id -ne [int]$atual.id_contato) { $mud['id_contato'] = @{ Tipo = 'inteiro'; Novo = [int]$ct.id; Antigo = $atual.id_contato } }
        }
        if (Test-TemArg $a 'observacoes_acrescentar') {
            $obs = Add-Observacao $atual.observacoes ([string]$a.observacoes_acrescentar)
            $mud['observacoes'] = @{ Tipo = 'memo'; Novo = $obs; Antigo = $atual.observacoes }
        }
        $avisos = Get-CriticaAtendimento $novo
        if ($mud.Count -eq 0) {
            return [ordered]@{ codigo = $codigo; gravado = $false; mensagem = 'Nada a alterar: os valores já estão iguais.'; versao = $atual.versao }
        }
        Update-ComVersao $ctx 'oportunidades' ([int]$atual.id) ([int]$atual.versao) $mud 'ALTERAR'
        return [ordered]@{
            codigo = $codigo; gravado = $true; versao = ([int]$atual.versao + 1)
            alterados = @($mud.Keys | ForEach-Object { [ordered]@{ campo = $_; antes = (ConvertFrom-ValorBanco $mud[$_].Antigo); depois = (ConvertFrom-ValorBanco $mud[$_].Novo) } })
            situacao = Get-Situacao $true ([string]$novo.etapa) $novo.dt_prox_acao $novo.retomar_em
            avisos = @($avisos)
        }
    }
}

function Invoke-AtualizarCadastroCliente($a) {
    Assert-EsquemaGravavel
    return Invoke-Transacao {
        param($ctx)
        $atual = Get-ClienteNaTransacao $ctx $a
        Assert-Versao $atual $a
        $defs = [ordered]@{
            empresa = @('texto', $null); cnpj = @('texto', $null); cidade = @('texto', $null); uf = @('texto', 'UF')
            segmento = @('texto', 'Segmento'); telefone = @('texto', $null); email = @('texto', $null)
            responsavel = @('texto', 'Responsavel'); qualificacao = @('texto', $null)
            aderencia = @('inteiro', $null); porte = @('inteiro', $null)
        }
        $mud = [ordered]@{}
        foreach ($campo in $defs.Keys) {
            if (-not (Test-TemArg $a $campo)) { continue }
            $tipo = $defs[$campo][0]; $lista = $defs[$campo][1]
            $v = $a.$campo
            if ($tipo -eq 'inteiro') {
                if ($null -ne $v) { $v = [int]$v; if ($v -lt 1 -or $v -gt 3) { throw "Campo '$campo' vai de 1 a 3." } }
            } elseif ($lista) {
                $v = Resolve-ValorLista $ctx $lista $campo $v
            } else {
                $v = ConvertTo-Normalizado $campo $v
                Assert-Tamanho $campo $v
            }
            switch ($campo) {
                'empresa'  { if ($null -eq $v) { throw 'A empresa não pode ficar vazia.' } }
                'telefone' { $m = Get-CriticaTelefone $v; if ($m) { throw $m } }
                'email'    { if ($v -and $v -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') { throw ("E-mail inválido: '" + $v + "'.") } }
                'cnpj'     {
                    if ($v) {
                        if (-not (Test-CnpjValido $v)) { throw 'CNPJ inválido (dígito verificador não confere). Deixe sem alterar se não tiver o número correto.' }
                        $dono = $ctx.Valor('SELECT empresa FROM clientes WHERE cnpj=? AND id<>?', @((P 'texto' $v), (P 'inteiro' ([int]$atual.id))))
                        if ($dono) { throw ('CNPJ já cadastrado em outra empresa: ' + $dono + '.') }
                    }
                }
            }
            if (-not (Test-Igual $atual[$campo] $v)) { $mud[$campo] = @{ Tipo = $tipo; Novo = $v; Antigo = $atual[$campo] } }
        }
        if (Test-TemArg $a 'observacoes_acrescentar') {
            $obs = Add-Observacao $atual.observacoes ([string]$a.observacoes_acrescentar)
            $mud['observacoes'] = @{ Tipo = 'memo'; Novo = $obs; Antigo = $atual.observacoes }
        }
        if ($mud.Count -eq 0) {
            return [ordered]@{ id_cliente = $atual.id; gravado = $false; mensagem = 'Nada a alterar: os valores já estão iguais.'; versao = $atual.versao }
        }
        Update-ComVersao $ctx 'clientes' ([int]$atual.id) ([int]$atual.versao) $mud 'ALTERAR'
        return [ordered]@{
            id_cliente = $atual.id; codigo = Format-CodigoCliente $atual.codigo_cliente; gravado = $true
            versao = ([int]$atual.versao + 1)
            alterados = @($mud.Keys | ForEach-Object { [ordered]@{ campo = $_; antes = (ConvertFrom-ValorBanco $mud[$_].Antigo); depois = (ConvertFrom-ValorBanco $mud[$_].Novo) } })
        }
    }
}

<#
  Contexto (ctx_*): campos que SÓ a IA escreve - o formulário não os edita.
  Por isso não entram no controle de versão: não há conflito possível com o
  vendedor, e a IA pode gravar com a ficha aberta na tela de alguém.
  O log é gravado igual, e é por ele que as estações percebem a mudança.
#>
function Set-Contexto($ctx, [string]$tabela, $atual, $a, [string[]]$campos) {
    $mud = [ordered]@{}
    foreach ($campo in $campos) {
        if (-not (Test-TemArg $a $campo)) { continue }
        $v = [string]$a.$campo
        if ($null -ne $a.$campo) { $v = $v.Trim() }
        if ($v -eq '') { $v = $null }
        Assert-Tamanho $campo $v
        if (-not (Test-Igual $atual[$campo] $v)) { $mud[$campo] = @{ Novo = $v; Antigo = $atual[$campo] } }
    }
    if ($mud.Count -eq 0) { return @() }
    $sets = @(); $params = @()
    foreach ($c in $mud.Keys) { $sets += ($c + '=?'); $params += (P 'texto' $mud[$c].Novo) }
    $sql = 'UPDATE ' + $tabela + ' SET ' + ($sets -join ', ') + ', ctx_atualizado_em=?, ctx_atualizado_por=? WHERE id=?'
    $params += (P 'data' (Get-Date)); $params += (P 'texto' $script:UsuarioLog); $params += (P 'inteiro' ([int]$atual.id))
    $n = $ctx.Executar($sql, $params)
    if ($n -ne 1) { throw ('UPDATE afetou ' + $n + ' linha(s); nada foi gravado.') }
    foreach ($c in $mud.Keys) { $ctx.Log($tabela, [int]$atual.id, $c, $mud[$c].Antigo, $mud[$c].Novo, 'CONTEXTO') }
    return @($mud.Keys)
}

function Invoke-AtualizarContextoCliente($a) {
    Assert-EsquemaGravavel
    return Invoke-Transacao {
        param($ctx)
        $atual = Get-ClienteNaTransacao $ctx $a
        $alt = Set-Contexto $ctx 'clientes' $atual $a @('ctx_resumo', 'ctx_familia', 'ctx_arquivo')
        return [ordered]@{ id_cliente = $atual.id; empresa = $atual.empresa; gravado = ($alt.Count -gt 0); campos = @($alt) }
    }
}

function Invoke-AtualizarContextoAtendimento($a) {
    Assert-EsquemaGravavel
    $codigo = [string]$a.codigo
    return Invoke-Transacao {
        param($ctx)
        $atual = Get-AtendimentoNaTransacao $ctx $codigo
        $alt = Set-Contexto $ctx 'oportunidades' $atual $a @('ctx_resumo', 'ctx_arquivo')
        return [ordered]@{ codigo = $codigo; gravado = ($alt.Count -gt 0); campos = @($alt) }
    }
}

# ================================================================= CONTATOS
<#
  Contato: mesmas regras do front (modCRM.CriarContato, modValidacao.CriticarContato,
  modAcoes.AlternarSimples).
    - código CT-CCCC-NNNN é CONGELADO na criação: nunca recalculado, nem ao trocar de empresa;
    - contato não se exclui: só desativa/reativa;
    - cadastrar contato em pré-cliente PROMOVE a empresa (código novo) - só com promover_pre_cliente=true,
      que é o equivalente da pergunta que o front faz ao vendedor.
#>
function Get-ContatoNaTransacao($ctx, $a) {
    if (Test-TemArg $a 'codigo_contato') {
        $cod = ([string]$a.codigo_contato).Trim().ToUpper()
        if ($cod -notmatch '^CT-\d{4}-\d{4}$') { throw ("Código de contato inválido: '" + $a.codigo_contato + "'. Formato: CT-0000-0000.") }
        $sql = 'SELECT * FROM contatos WHERE codigo=?'; $par = @((P 'texto' $cod)); $ref = $cod
    } elseif (Test-TemArg $a 'id_contato') {
        $sql = 'SELECT * FROM contatos WHERE id=?'; $par = @((P 'inteiro' ([int]$a.id_contato))); $ref = 'id ' + $a.id_contato
    } else { throw 'Informe codigo_contato (CT-0000-0000) ou id_contato.' }
    # @(...): uma linha só não pode ser desembrulhada (viraria o próprio dicionário e [0] leria o 1º campo)
    $l = @(if ($ctx) { $ctx.Consultar($sql, $par) } else { Invoke-Consulta $sql $par })
    if ($l.Count -eq 0) { throw ('Contato não encontrado: ' + $ref + '.') }
    return $l[0]
}

# nome, cargo, telefone, email: normaliza e critica igual ao front. Devolve o valor final.
function ConvertTo-CampoContato([string]$campo, $valor) {
    $v = ConvertTo-Normalizado $campo $valor
    Assert-Tamanho $campo $v
    switch ($campo) {
        'nome'     { if ($null -eq $v) { throw 'Informe o nome do contato (use GERAL se for o contato geral da empresa).' } }
        'telefone' { $m = Get-CriticaTelefone $v; if ($m) { throw $m } }
        'email'    { if ($v -and $v -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') { throw ("E-mail inválido: '" + $v + "'.") } }
    }
    return $v
}

function Test-ErroDuplicidade([System.Exception]$ex) {
    $m = ''
    for ($e = $ex; $e; $e = $e.InnerException) { $m += ' ' + $e.Message }
    return ($m -match '(?i)duplica|duplicate')
}

function Invoke-BuscarContatos($a) {
    $n = Get-Limite $a
    $w = New-Object System.Collections.ArrayList; $p = New-Object System.Collections.ArrayList
    $texto = [string](Get-Arg $a 'texto' '')
    if ($texto.Trim() -ne '') {
        $like = '%' + $texto.Trim() + '%'
        $cond = '(ct.nome LIKE ? OR ct.cargo LIKE ? OR ct.email LIKE ? OR ct.codigo LIKE ? OR cl.empresa LIKE ?'
        1..5 | ForEach-Object { [void]$p.Add((P 'texto' $like)) }
        $dig = Get-SoDigitos $texto
        if ($dig.Length -ge 4 -and $texto.Trim() -notmatch '^CT-') { $cond += ' OR ct.telefone LIKE ?'; [void]$p.Add((P 'texto' ('%' + $dig + '%'))) }
        [void]$w.Add($cond + ')')
    }
    if (Test-TemArg $a 'codigo_cliente') { [void]$w.Add('cl.codigo_cliente=?'); [void]$p.Add((P 'inteiro' ([int]$a.codigo_cliente))) }
    if (Test-TemArg $a 'id_cliente')     { [void]$w.Add('ct.id_cliente=?');     [void]$p.Add((P 'inteiro' ([int]$a.id_cliente))) }
    if (-not [bool](Get-Arg $a 'incluir_inativos' $false)) { [void]$w.Add('ct.ativo=True') }
    $sql = 'SELECT TOP ' + $n + ' ct.id, ct.codigo, ct.nome, ct.cargo, ct.telefone, ct.email, ct.ativo, ct.versao,' +
           ' ct.id_cliente, cl.codigo_cliente, cl.empresa FROM contatos ct LEFT JOIN clientes cl ON ct.id_cliente = cl.id'
    if ($w.Count -gt 0) { $sql += ' WHERE ' + ($w -join ' AND ') }
    $sql += ' ORDER BY ct.codigo, ct.id'
    $linhas = Invoke-Consulta $sql $p.ToArray()
    return [ordered]@{ quantidade = $linhas.Count; limite = $n; contatos = $linhas }
}

function Invoke-ObterContato($a) {
    $c = Get-ContatoNaTransacao $null $a
    $cl = Invoke-Consulta 'SELECT id, codigo_cliente, estagio, empresa, cidade, uf, ativo FROM clientes WHERE id=?' @((P 'inteiro' ([int]$c.id_cliente)))
    if ($cl.Count -gt 0) { $cl[0]['codigo'] = Format-CodigoCliente $cl[0].codigo_cliente; $c['empresa'] = $cl[0] }
    $ops = Invoke-Consulta ('SELECT o.codigo, o.etapa, o.responsavel, o.valor, o.maquina, o.dt_prox_acao, o.retomar_em, o.id_cliente' +
                            ' FROM oportunidades o WHERE o.id_contato=? ORDER BY o.codigo DESC') @((P 'inteiro' ([int]$c.id)))
    foreach ($o in $ops) { $o['situacao'] = Get-Situacao $true ([string]$o.etapa) $o.dt_prox_acao $o.retomar_em }
    $c['atendimentos'] = $ops
    return $c
}

function Invoke-CriarContato($a) {
    Assert-EsquemaGravavel
    # críticas antes de abrir a transação: erro de digitação não gasta tentativa
    $d = [ordered]@{}
    foreach ($campo in @('nome', 'cargo', 'telefone', 'email')) {
        if ($campo -eq 'nome' -or (Test-TemArg $a $campo)) { $d[$campo] = ConvertTo-CampoContato $campo (Get-Arg $a $campo) }
    }
    $obs = $null
    if ((Test-TemArg $a 'observacoes') -and [string]$a.observacoes -ne '') { $obs = Add-Observacao $null ([string]$a.observacoes) }
    $promover = [bool](Get-Arg $a 'promover_pre_cliente' $false)

    for ($tentativa = 1; $tentativa -le 5; $tentativa++) {
        try {
            return Invoke-Transacao {
                param($ctx)
                $cli = Get-ClienteNaTransacao $ctx $a
                if (-not [bool]$cli.ativo) { throw ('A empresa ' + $cli.empresa + ' está inativa. Reative-a antes de cadastrar contato.') }
                $codCli = $cli.codigo_cliente
                $promovida = $false
                if ($null -eq $codCli) {
                    if (-not $promover) {
                        throw ('A empresa ' + $cli.empresa + ' é pré-cliente. Cadastrar contato a promove a cliente e atribui o próximo código. ' +
                               'Confirme com o usuário e repita com promover_pre_cliente=true.')
                    }
                    $codCli = [int]$ctx.Valor('SELECT MAX(codigo_cliente) FROM clientes', @()) + 1
                    $n = $ctx.Executar('UPDATE clientes SET codigo_cliente=?, estagio=?, versao=versao+1, alterado_em=?, alterado_por=? WHERE id=? AND codigo_cliente IS NULL',
                        @((P 'inteiro' $codCli), (P 'texto' 'Cliente'), (P 'data' (Get-Date)), (P 'texto' $script:UsuarioLog), (P 'inteiro' ([int]$cli.id))))
                    if ($n -ne 1) { throw 'Conflito: a empresa foi promovida por outra pessoa agora. Nada foi gravado; repita.' }
                    $ctx.Log('clientes', [int]$cli.id, 'codigo_cliente', '', [string]$codCli, 'PROMOVER')
                    $ctx.Log('clientes', [int]$cli.id, 'estagio', ('Pr' + [char]0x00E9 + '-cliente'), 'Cliente', 'PROMOVER')
                    $promovida = $true
                }
                $controle = [int]$ctx.Valor('SELECT MAX(controle) FROM contatos', @()) + 1
                $codigo = 'CT-' + ([int]$codCli).ToString('0000') + '-' + $controle.ToString('0000')
                $agora = Get-Date
                $ctx.Executar(('INSERT INTO contatos (codigo, controle, id_cliente, nome, cargo, telefone, email, observacoes,' +
                               ' ativo, versao, criado_em, alterado_em, alterado_por) VALUES (?,?,?,?,?,?,?,?,True,1,?,?,?)'),
                    @((P 'texto' $codigo), (P 'inteiro' $controle), (P 'inteiro' ([int]$cli.id)), (P 'texto' $d['nome']),
                      (P 'texto' $d['cargo']), (P 'texto' $d['telefone']), (P 'texto' $d['email']), (P 'memo' $obs),
                      (P 'data' $agora), (P 'data' $agora), (P 'texto' $script:UsuarioLog))) | Out-Null
                $id = [int]$ctx.Valor('SELECT id FROM contatos WHERE codigo=?', @((P 'texto' $codigo)))
                $ctx.Log('contatos', $id, '', '', 'cadastro novo pela IA', 'INSERIR')
                return [ordered]@{
                    gravado = $true; id_contato = $id; codigo_contato = $codigo; versao = 1
                    empresa = $cli.empresa; codigo_cliente = Format-CodigoCliente $codCli; empresa_promovida = $promovida
                    nome = $d['nome']; cargo = $d['cargo']; telefone = $d['telefone']; email = $d['email']
                }
            }
        } catch {
            if ($tentativa -lt 5 -and (Test-ErroDuplicidade $_.Exception)) { continue }
            throw
        }
    }
}

function Invoke-AtualizarContato($a) {
    Assert-EsquemaGravavel
    return Invoke-Transacao {
        param($ctx)
        $atual = Get-ContatoNaTransacao $ctx $a
        Assert-Versao $atual $a
        $mud = [ordered]@{}
        foreach ($campo in @('nome', 'cargo', 'telefone', 'email')) {
            if (-not (Test-TemArg $a $campo)) { continue }
            $v = ConvertTo-CampoContato $campo $a.$campo
            if (-not (Test-Igual $atual[$campo] $v)) { $mud[$campo] = @{ Tipo = 'texto'; Novo = $v; Antigo = $atual[$campo] } }
        }
        if (Test-TemArg $a 'observacoes_acrescentar') {
            $obs = Add-Observacao $atual.observacoes ([string]$a.observacoes_acrescentar)
            $mud['observacoes'] = @{ Tipo = 'memo'; Novo = $obs; Antigo = $atual.observacoes }
        }
        $avisos = New-Object System.Collections.ArrayList
        if (Test-TemArg $a 'codigo_cliente_destino') {
            $dest = $ctx.Consultar('SELECT id, empresa, ativo FROM clientes WHERE codigo_cliente=?', @((P 'inteiro' ([int]$a.codigo_cliente_destino))))
            if ($dest.Count -eq 0) { throw ('Empresa de destino não encontrada: código ' + $a.codigo_cliente_destino + '. Pré-cliente não recebe contato transferido: promova antes.') }
            if (-not [bool]$dest[0].ativo) { throw ('A empresa de destino ' + $dest[0].empresa + ' está inativa.') }
            if ([int]$dest[0].id -ne [int]$atual.id_cliente) {
                $mud['id_cliente'] = @{ Tipo = 'inteiro'; Novo = [int]$dest[0].id; Antigo = $atual.id_cliente }
                [void]$avisos.Add('Código ' + $atual.codigo + ' mantido (é congelado). Atendimentos já abertos continuam na empresa em que foram abertos.')
            }
        }
        $acao = 'ALTERAR'
        if (Test-TemArg $a 'ativo') {
            $v = [bool]$a.ativo
            if ([bool]$atual.ativo -ne $v) {
                $mud['ativo'] = @{ Tipo = 'logico'; Novo = $v; Antigo = [bool]$atual.ativo }
                if ($mud.Count -eq 1) { $acao = if ($v) { 'REATIVACAO' } else { 'DESATIVACAO' } }
                if (-not $v) {
                    $abertos = [int]$ctx.Valor(('SELECT COUNT(*) FROM oportunidades WHERE id_contato=? AND etapa NOT IN (' +
                                               (($script:EtapasEncerradas | ForEach-Object { '?' }) -join ',') + ')'),
                                               (@((P 'inteiro' ([int]$atual.id))) + @($script:EtapasEncerradas | ForEach-Object { P 'texto' $_ })))
                    if ($abertos -gt 0) { [void]$avisos.Add('Contato desativado com ' + $abertos + ' atendimento(s) em aberto; eles continuam como estão.') }
                }
            }
        }
        if ($mud.Count -eq 0) {
            return [ordered]@{ codigo_contato = $atual.codigo; gravado = $false; mensagem = 'Nada a alterar: os valores já estão iguais.'; versao = $atual.versao }
        }
        Update-ComVersao $ctx 'contatos' ([int]$atual.id) ([int]$atual.versao) $mud $acao
        return [ordered]@{
            codigo_contato = $atual.codigo; id_contato = $atual.id; gravado = $true; versao = ([int]$atual.versao + 1)
            alterados = @($mud.Keys | ForEach-Object { [ordered]@{ campo = $_; antes = (ConvertFrom-ValorBanco $mud[$_].Antigo); depois = (ConvertFrom-ValorBanco $mud[$_].Novo) } })
            avisos = @($avisos)
        }
    }
}

# ================================================================= CLIENTES, ATENDIMENTOS, LISTAS, METAS
# Críticas do atendimento (modValidacao.CriticarOportunidade + AvisarOportunidade). Lança no erro; devolve avisos.
function Get-CriticaAtendimento($novo) {
    $etapa = [string]$novo.etapa
    if ($etapa -eq '') { throw 'Informe a etapa.' }
    if ([string]$novo.responsavel -eq '') { throw 'Informe o responsável.' }
    $dEnt = ConvertTo-DataOuNulo $novo.dt_entrada
    $dPro = ConvertTo-DataOuNulo $novo.dt_proposta
    $dDes = ConvertTo-DataOuNulo $novo.dt_desfecho
    if ($dEnt -and $dPro -and $dPro -lt $dEnt) { throw 'A data da proposta é anterior à data de entrada.' }
    if ($dEnt -and $dDes -and $dDes -lt $dEnt) { throw 'A data do desfecho é anterior à data de entrada.' }
    if ($script:EtapasEncerradas -contains $etapa) {
        if ([string]$novo.motivo_desfecho -eq '') { throw ('Etapa "' + $etapa + '" exige motivo_desfecho.') }
        if (-not $dDes) { throw ('Etapa "' + $etapa + '" exige dt_desfecho.') }
    }
    if ($script:EtapasComFamilia -contains $etapa -and [string]$novo.familia -eq '') {
        throw ('Etapa "' + $etapa + '" exige a família de máquina.')
    }
    $avisos = New-Object System.Collections.ArrayList
    if ($etapa -eq 'Pedido Fechado' -and $null -eq $novo.valor) { [void]$avisos.Add('Pedido fechado sem valor informado.') }
    if (($script:EtapasEncerradas -notcontains $etapa) -and -not $novo.dt_prox_acao -and -not $novo.retomar_em) {
        [void]$avisos.Add('Sem próxima ação e sem data de retomada: o atendimento some da fila da manhã.')
    }
    return ,$avisos
}

# Campos editáveis do atendimento: campo -> @(tipo, lista)
$script:DefsAtendimento = [ordered]@{
    etapa = @('texto', 'Etapa'); responsavel = @('texto', 'Responsavel'); prioridade = @('texto', 'Prioridade')
    origem = @('texto', 'Origem'); familia = @('texto', 'Familia'); categoria = @('texto', 'Categoria')
    motivo_desfecho = @('texto', 'Motivo'); tipo_venda = @('texto', $null); maquina = @('texto', $null)
    orcamento = @('texto', $null); valor = @('moeda', $null); tentativas = @('inteiro', $null)
    prox_acao = @('memo', $null); dt_entrada = @('data', $null); dt_proposta = @('data', $null); ultima_interacao = @('data', $null)
    dt_prox_acao = @('data', $null); retomar_em = @('data', $null); dt_desfecho = @('data', $null)
}
function ConvertTo-CampoAtendimento($ctx, [string]$campo, $v) {
    $tipo = $script:DefsAtendimento[$campo][0]; $lista = $script:DefsAtendimento[$campo][1]
    switch ($tipo) {
        'data'    { return (ConvertTo-Data $campo $v) }
        'moeda'   { if ($null -ne $v) { $v = [decimal]$v; if ($v -lt 0) { throw 'O valor não pode ser negativo.' } }; return $v }
        'inteiro' { if ($null -ne $v) { $v = [int]$v; if ($v -lt 0) { throw "Campo '$campo' não pode ser negativo." } }; return $v }
        default   {
            if ($lista) { return (Resolve-ValorLista $ctx $lista $campo $v) }
            $v = ConvertTo-Normalizado $campo $v; Assert-Tamanho $campo $v; return $v
        }
    }
}

# Repete o bloco quando o número sequencial colidiu no índice único (dois gravando ao mesmo tempo).
function Invoke-ComRetentativa([scriptblock]$bloco) {
    for ($t = 1; $t -le 5; $t++) {
        try { return (Invoke-Transacao $bloco) }
        catch { if ($t -lt 5 -and (Test-ErroDuplicidade $_.Exception)) { continue }; throw }
    }
}

# Promove pré-cliente dentro da transação. Devolve o código atribuído.
function Invoke-PromoverNaTransacao($ctx, $cli) {
    $cod = [int]$ctx.Valor('SELECT MAX(codigo_cliente) FROM clientes', @()) + 1
    $n = $ctx.Executar('UPDATE clientes SET codigo_cliente=?, estagio=?, versao=versao+1, alterado_em=?, alterado_por=? WHERE id=? AND codigo_cliente IS NULL',
        @((P 'inteiro' $cod), (P 'texto' 'Cliente'), (P 'data' (Get-Date)), (P 'texto' $script:UsuarioLog), (P 'inteiro' ([int]$cli.id))))
    if ($n -ne 1) { throw 'Conflito: a empresa foi promovida por outra pessoa agora. Nada foi gravado; leia de novo.' }
    $ctx.Log('clientes', [int]$cli.id, 'codigo_cliente', '', [string]$cod, 'PROMOVER')
    $ctx.Log('clientes', [int]$cli.id, 'estagio', ('Pr' + [char]0x00E9 + '-cliente'), 'Cliente', 'PROMOVER')
    return $cod
}

# Normaliza e critica um campo de cliente (mesmas regras de atualizar_cadastro_cliente).
function ConvertTo-CampoCliente($ctx, [string]$campo, $v, [int]$idProprio) {
    $defs = @{ uf = 'UF'; segmento = 'Segmento'; responsavel = 'Responsavel' }
    if ($campo -in @('aderencia', 'porte')) {
        if ($null -ne $v) { $v = [int]$v; if ($v -lt 1 -or $v -gt 3) { throw "Campo '$campo' vai de 1 a 3." } }
        return $v
    }
    if ($defs.ContainsKey($campo)) { return (Resolve-ValorLista $ctx $defs[$campo] $campo $v) }
    if ($campo -eq 'qualificacao' -and @(Get-ValoresLista $ctx 'Qualificacao').Count -gt 0) { return (Resolve-ValorLista $ctx 'Qualificacao' $campo $v) }
    $v = ConvertTo-Normalizado $campo $v
    Assert-Tamanho $campo $v
    switch ($campo) {
        'empresa'  { if ($null -eq $v) { throw 'A empresa não pode ficar vazia.' } }
        'telefone' { $m = Get-CriticaTelefone $v; if ($m) { throw $m } }
        'email'    { if ($v -and $v -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') { throw ("E-mail inválido: '" + $v + "'.") } }
        'cnpj'     {
            if ($v) {
                if (-not (Test-CnpjValido $v)) { throw 'CNPJ inválido (dígito verificador não confere). Deixe em branco se não tiver o número correto.' }
                $dono = $ctx.Valor('SELECT empresa FROM clientes WHERE cnpj=? AND id<>?', @((P 'texto' $v), (P 'inteiro' $idProprio)))
                if ($dono) { throw ('CNPJ já cadastrado em outra empresa: ' + $dono + '.') }
            }
        }
    }
    return $v
}

function Invoke-CriarCliente($a) {
    Assert-EsquemaGravavel
    $comoCliente = [bool](Get-Arg $a 'como_cliente' $false)
    return Invoke-ComRetentativa {
        param($ctx)
        $campos = @('empresa', 'cnpj', 'cidade', 'uf', 'segmento', 'telefone', 'email', 'responsavel', 'aderencia', 'porte', 'qualificacao')
        $tipos = @{ aderencia = 'inteiro'; porte = 'inteiro' }
        $d = [ordered]@{}
        foreach ($c in $campos) { if ($c -eq 'empresa' -or (Test-TemArg $a $c)) { $d[$c] = ConvertTo-CampoCliente $ctx $c (Get-Arg $a $c) 0 } }
        $parecidas = $ctx.Consultar('SELECT TOP 5 id, codigo_cliente, empresa, cidade FROM clientes WHERE empresa=?', @((P 'texto' $d['empresa'])))
        if ($parecidas.Count -gt 0 -and -not [bool](Get-Arg $a 'ignorar_homonimo' $false)) {
            throw ('Já existe empresa com o nome ' + $d['empresa'] + ' (id ' + (($parecidas | ForEach-Object { $_.id }) -join ', ') +
                   '). Confira com o usuário; se for outra empresa, repita com ignorar_homonimo=true.')
        }
        $cod = $null; $estagio = ('Pr' + [char]0x00E9 + '-cliente')
        if ($comoCliente) { $cod = [int]$ctx.Valor('SELECT MAX(codigo_cliente) FROM clientes', @()) + 1; $estagio = 'Cliente' }
        $obs = $null
        if ((Test-TemArg $a 'observacoes') -and [string]$a.observacoes -ne '') { $obs = Add-Observacao $null ([string]$a.observacoes) }
        $cols = @('codigo_cliente', 'estagio'); $par = @((P 'inteiro' $cod), (P 'texto' $estagio))
        foreach ($c in $d.Keys) {
            $t = 'texto'; if ($tipos.ContainsKey($c)) { $t = $tipos[$c] }
            $cols += $c; $par += (P $t $d[$c])
        }
        $cols += 'observacoes'; $par += (P 'memo' $obs)
        $agora = Get-Date
        $sql = 'INSERT INTO clientes (' + ($cols -join ', ') + ', ativo, versao, criado_em, alterado_em, alterado_por) VALUES (' +
               (($cols | ForEach-Object { '?' }) -join ',') + ',True,1,?,?,?)'
        $par += (P 'data' $agora); $par += (P 'data' $agora); $par += (P 'texto' $script:UsuarioLog)
        [void]$ctx.Executar($sql, $par)
        $id = [int]$ctx.Valor('SELECT @@IDENTITY', @())
        $ctx.Log('clientes', $id, '', '', 'cadastro novo pela IA', 'INSERIR')
        return [ordered]@{ gravado = $true; id_cliente = $id; codigo_cliente = Format-CodigoCliente $cod; estagio = $estagio; versao = 1; empresa = $d['empresa'] }
    }
}

function Invoke-PromoverCliente($a) {
    Assert-EsquemaGravavel
    return Invoke-ComRetentativa {
        param($ctx)
        $cli = Get-ClienteNaTransacao $ctx $a
        if ($null -ne $cli.codigo_cliente) { throw ('A empresa já é cliente, código ' + (Format-CodigoCliente $cli.codigo_cliente) + '.') }
        $cod = Invoke-PromoverNaTransacao $ctx $cli
        return [ordered]@{ gravado = $true; id_cliente = $cli.id; empresa = $cli.empresa; codigo_cliente = Format-CodigoCliente $cod; estagio = 'Cliente' }
    }
}

# Muda ativo dos registros do filtro que ainda não estão no estado novo; sobe versão e loga cada um (modAcoes.MudarAtivo).
function Set-AtivoEm($ctx, [string]$tabela, [string]$filtro, [int]$valor, [bool]$novo, [string]$acao, [string]$detalhe) {
    $ids = $ctx.Consultar(('SELECT id FROM ' + $tabela + ' WHERE ' + $filtro + ' AND ativo=?'), @((P 'inteiro' $valor), (P 'logico' (-not $novo))))
    $n = 0
    foreach ($r in $ids) {
        $n += $ctx.Executar(('UPDATE ' + $tabela + ' SET ativo=?, versao=versao+1, alterado_em=?, alterado_por=? WHERE id=? AND ativo=?'),
            @((P 'logico' $novo), (P 'data' (Get-Date)), (P 'texto' $script:UsuarioLog), (P 'inteiro' ([int]$r.id)), (P 'logico' (-not $novo))))
        $txt = [string]$novo; if ($detalhe) { $txt += ' (' + $detalhe + ')' }
        $ctx.Log($tabela, [int]$r.id, 'ativo', [string](-not $novo), $txt, $acao)
    }
    return $n
}

function Invoke-SituacaoCliente($a) {
    Assert-EsquemaGravavel
    $acao = [string]$a.acao
    return Invoke-Transacao {
        param($ctx)
        $cli = Get-ClienteNaTransacao $ctx $a
        Assert-Versao $cli $a
        $id = [int]$cli.id
        $encerr = (($script:EtapasEncerradas | ForEach-Object { '?' }) -join ',')
        $pEnc = @($script:EtapasEncerradas | ForEach-Object { P 'texto' $_ })
        switch ($acao) {
            'desativar' {
                if (-not [bool]$cli.ativo) { throw 'A empresa já está inativa.' }
                $nOp = [int]$ctx.Valor(('SELECT COUNT(*) FROM oportunidades WHERE id_cliente=? AND etapa NOT IN (' + $encerr + ')'), (@((P 'inteiro' $id)) + $pEnc))
                [void](Set-AtivoEm $ctx 'clientes' 'id=?' $id $false 'DESATIVACAO' ([string]$nOp + ' oportunidade(s) em aberto mantida(s)'))
                $nCt = Set-AtivoEm $ctx 'contatos' 'id_cliente=?' $id $false 'DESATIVACAO' 'cascata da desativação do cliente'
                $aviso = $null
                if ($nOp -gt 0) { $aviso = [string]$nOp + ' atendimento(s) em aberto continuam no funil; encerre um a um se não forem mais trabalhados.' }
                return [ordered]@{ gravado = $true; empresa = $cli.empresa; ativo = $false; contatos_desativados = $nCt; aviso = $aviso }
            }
            'reativar' {
                if ([bool]$cli.ativo) { throw 'A empresa já está ativa.' }
                [void](Set-AtivoEm $ctx 'clientes' 'id=?' $id $true 'REATIVACAO' '')
                $nCt = Set-AtivoEm $ctx 'contatos' 'id_cliente=?' $id $true 'REATIVACAO' 'cascata da reativação do cliente'
                return [ordered]@{ gravado = $true; empresa = $cli.empresa; ativo = $true; contatos_reativados = $nCt }
            }
            'excluir_pre_cliente' {
                if ($null -ne $cli.codigo_cliente) { throw 'Cliente com código nunca é excluído: use acao=desativar.' }
                $nCt = [int]$ctx.Valor('SELECT COUNT(*) FROM contatos WHERE id_cliente=?', @((P 'inteiro' $id)))
                $nOp = [int]$ctx.Valor('SELECT COUNT(*) FROM oportunidades WHERE id_cliente=?', @((P 'inteiro' $id)))
                if ($nCt -gt 0 -or $nOp -gt 0) { throw ('Pré-cliente com vínculo não é excluído (contatos: ' + $nCt + ', atendimentos: ' + $nOp + '). Confira o cadastro.') }
                $n = $ctx.Executar('DELETE FROM clientes WHERE id=? AND codigo_cliente IS NULL', @((P 'inteiro' $id)))
                if ($n -ne 1) { throw 'O registro não foi excluído: não existe mais ou já virou cliente.' }
                $texto = [string]$cli.empresa
                if ($cli.cidade) { $texto += ' - ' + $cli.cidade + '/' + $cli.uf }
                if ($cli.cnpj) { $texto += ' - CNPJ ' + $cli.cnpj }
                $ctx.Log('clientes', $id, '(registro)', $texto, '', 'EXCLUSAO')
                return [ordered]@{ gravado = $true; excluido = $true; empresa = $cli.empresa; id_cliente = $id }
            }
            default { throw 'acao deve ser desativar, reativar ou excluir_pre_cliente.' }
        }
    }
}

function Invoke-AbrirAtendimento($a) {
    Assert-EsquemaGravavel
    $promover = [bool](Get-Arg $a 'promover_pre_cliente' $false)
    return Invoke-ComRetentativa {
        param($ctx)
        $ct = Get-ContatoNaTransacao $ctx $a
        if (-not [bool]$ct.ativo) { throw ('O contato ' + $ct.codigo + ' está inativo. Reative-o antes de abrir atendimento.') }
        $cli = $ctx.Consultar('SELECT * FROM clientes WHERE id=?', @((P 'inteiro' ([int]$ct.id_cliente))))[0]
        if (-not [bool]$cli.ativo) { throw ('A empresa ' + $cli.empresa + ' está inativa.') }
        $novo = [ordered]@{}
        foreach ($c in $script:DefsAtendimento.Keys) { $novo[$c] = $null }
        $novo['dt_entrada'] = (Get-Date).Date
        foreach ($c in $script:DefsAtendimento.Keys) { if (Test-TemArg $a $c) { $novo[$c] = ConvertTo-CampoAtendimento $ctx $c $a.$c } }
        $avisos = Get-CriticaAtendimento $novo
        $idAnt = $null
        if (Test-TemArg $a 'codigo_atendimento_anterior') {
            # retomada: o anterior é da MESMA empresa (mesma regra do front: modCRM.CriticarAnterior)
            $ant = Get-AtendimentoNaTransacao $ctx ([string]$a.codigo_atendimento_anterior)
            if ([int]$ant.id_cliente -ne [int]$cli.id) { throw ('O atendimento ' + $ant.codigo + ' é de outra empresa. O atendimento anterior tem de ser da mesma empresa do contato.') }
            $idAnt = [int]$ant.id
        }
        $codCli = $cli.codigo_cliente; $promovida = $false
        if ($null -eq $codCli) {
            if (-not $promover) { throw ('A empresa ' + $cli.empresa + ' é pré-cliente. Abrir atendimento a promove a cliente; confirme com o usuário e repita com promover_pre_cliente=true.') }
            $codCli = Invoke-PromoverNaTransacao $ctx $cli; $promovida = $true
        }
        $controle = [int]$ctx.Valor('SELECT MAX(controle) FROM oportunidades', @()) + 1
        $codigo = 'AT-' + ([int]$codCli).ToString('0000') + '-' + $controle.ToString('0000')
        $obs = $null
        if ((Test-TemArg $a 'observacoes') -and [string]$a.observacoes -ne '') { $obs = Add-Observacao $null ([string]$a.observacoes) }
        $cols = @('codigo', 'controle', 'id_cliente', 'id_contato', 'id_atendimento_anterior')
        $par = @((P 'texto' $codigo), (P 'inteiro' $controle), (P 'inteiro' ([int]$cli.id)), (P 'inteiro' ([int]$ct.id)), (P 'inteiro' $idAnt))
        foreach ($c in $script:DefsAtendimento.Keys) { $cols += $c; $par += (P $script:DefsAtendimento[$c][0] $novo[$c]) }
        $cols += 'observacoes'; $par += (P 'memo' $obs)
        $agora = Get-Date
        $par += (P 'data' $agora); $par += (P 'data' $agora); $par += (P 'texto' $script:UsuarioLog)
        [void]$ctx.Executar(('INSERT INTO oportunidades (' + ($cols -join ', ') + ', versao, criado_em, alterado_em, alterado_por) VALUES (' +
                            (($cols | ForEach-Object { '?' }) -join ',') + ',1,?,?,?)'), $par)
        $id = [int]$ctx.Valor('SELECT id FROM oportunidades WHERE codigo=?', @((P 'texto' $codigo)))
        $ctx.Log('oportunidades', $id, '', '', 'cadastro novo pela IA', 'INSERIR')
        return [ordered]@{ gravado = $true; codigo = $codigo; id = $id; versao = 1; empresa = $cli.empresa; contato = $ct.nome
                           empresa_promovida = $promovida; situacao = (Get-Situacao $true ([string]$novo.etapa) $novo.dt_prox_acao $novo.retomar_em)
                           avisos = @($avisos) }
    }
}

function Invoke-GerenciarLista($a) {
    Assert-EsquemaGravavel
    $tipo = ([string]$a.tipo).Trim(); $valor = ([string]$a.valor).Trim(); $acao = [string]$a.acao
    if ($tipo -eq '' -or $valor -eq '') { throw 'Informe tipo e valor.' }
    if ($valor.Length -gt 120) { throw 'Valor aceita no máximo 120 caracteres.' }
    return Invoke-Transacao {
        param($ctx)
        $tipos = @($ctx.Consultar('SELECT DISTINCT tipo FROM listas', @()) | ForEach-Object { [string]$_.tipo })
        $tipoOk = $tipos | Where-Object { $_ -ieq $tipo } | Select-Object -First 1
        if (-not $tipoOk) { throw ("Lista '" + $tipo + "' não existe. Listas: " + ($tipos -join ' | ')) }
        $ex = $ctx.Consultar('SELECT * FROM listas WHERE tipo=? AND valor=?', @((P 'texto' $tipoOk), (P 'texto' $valor)))
        $ordem = $null; if (Test-TemArg $a 'ordem') { $ordem = [int]$a.ordem }
        switch ($acao) {
            'incluir' {
                if ($ex.Count -gt 0) {
                    $sufixo = ''; if (-not [bool]$ex[0].ativo) { $sufixo = ' (inativo: use acao=ativar)' }
                    throw ("'" + $valor + "' já existe na lista " + $tipoOk + $sufixo + '.')
                }
                if ($null -eq $ordem) { $ordem = [int]$ctx.Valor('SELECT MAX(ordem) FROM listas WHERE tipo=?', @((P 'texto' $tipoOk))) + 1 }
                [void]$ctx.Executar('INSERT INTO listas (tipo, valor, ordem, ativo) VALUES (?,?,?,True)', @((P 'texto' $tipoOk), (P 'texto' $valor), (P 'inteiro' $ordem)))
                $id = [int]$ctx.Valor('SELECT @@IDENTITY', @())
                $ctx.Log('listas', $id, 'valor', '', ($tipoOk + ': ' + $valor), 'INSERIR')
            }
            { $_ -in @('ativar', 'desativar', 'ordenar') } {
                if ($ex.Count -eq 0) { throw ("'" + $valor + "' não existe na lista " + $tipoOk + '.') }
                $r = $ex[0]; $id = [int]$r.id
                if ($acao -eq 'ordenar') {
                    if ($null -eq $ordem) { throw 'Informe ordem.' }
                    [void]$ctx.Executar('UPDATE listas SET ordem=? WHERE id=?', @((P 'inteiro' $ordem), (P 'inteiro' $id)))
                    $ctx.Log('listas', $id, 'ordem', $r.ordem, $ordem, 'ALTERAR')
                } else {
                    $novo = ($acao -eq 'ativar')
                    if ([bool]$r.ativo -eq $novo) { throw ("'" + $valor + "' já está no estado pedido.") }
                    [void]$ctx.Executar('UPDATE listas SET ativo=? WHERE id=?', @((P 'logico' $novo), (P 'inteiro' $id)))
                    $ac = 'DESATIVACAO'; if ($novo) { $ac = 'REATIVACAO' }
                    $ctx.Log('listas', $id, 'ativo', [string](-not $novo), [string]$novo, $ac)
                }
            }
            default { throw 'acao deve ser incluir, ativar, desativar ou ordenar.' }
        }
        return [ordered]@{ gravado = $true; tipo = $tipoOk; valor = $valor; acao = $acao; valores_ativos = @(Get-ValoresLista $ctx $tipoOk) }
    }
}

function Invoke-ListarMetas($a) {
    $w = ''; $p = @()
    if (Test-TemArg $a 'ano') { $w = ' WHERE ano=?'; $p = @((P 'inteiro' ([int]$a.ano))) }
    $l = Invoke-Consulta ('SELECT id, ano, mes, meta_faturamento, meta_pedidos, meta_propostas, meta_contatos, alterado_em, alterado_por FROM metas' + $w + ' ORDER BY ano, mes') $p
    return [ordered]@{ quantidade = $l.Count; metas = $l }
}

function Invoke-GravarMeta($a) {
    Assert-EsquemaGravavel
    $ano = [int]$a.ano; $mes = [int]$a.mes
    if ($ano -lt 2000 -or $ano -gt 2100) { throw 'Ano fora do intervalo.' }
    if ($mes -lt 1 -or $mes -gt 12) { throw 'Mês vai de 1 a 12.' }
    $defs = [ordered]@{ meta_faturamento = 'moeda'; meta_pedidos = 'inteiro'; meta_propostas = 'inteiro'; meta_contatos = 'inteiro' }
    return Invoke-Transacao {
        param($ctx)
        $ex = $ctx.Consultar('SELECT * FROM metas WHERE ano=? AND mes=?', @((P 'inteiro' $ano), (P 'inteiro' $mes)))
        $mud = [ordered]@{}
        foreach ($c in $defs.Keys) {
            if (-not (Test-TemArg $a $c)) { continue }
            $v = $a.$c
            if ($null -ne $v) {
                if ($defs[$c] -eq 'moeda') { $v = [decimal]$v } else { $v = [int]$v }
                if ($v -lt 0) { throw "Campo '$c' não pode ser negativo." }
            }
            $antigo = $null; if ($ex.Count -gt 0) { $antigo = $ex[0][$c] }
            if (-not (Test-Igual $antigo $v)) { $mud[$c] = @{ Tipo = $defs[$c]; Novo = $v; Antigo = $antigo } }
        }
        if ($mud.Count -eq 0) { return [ordered]@{ gravado = $false; mensagem = 'Nada a alterar.' } }
        $agora = Get-Date
        $valores = @($mud.Keys | ForEach-Object { P $mud[$_].Tipo $mud[$_].Novo })
        if ($ex.Count -eq 0) {
            $cols = @('ano', 'mes') + @($mud.Keys)
            $par = @((P 'inteiro' $ano), (P 'inteiro' $mes)) + $valores + @((P 'data' $agora), (P 'texto' $script:UsuarioLog))
            [void]$ctx.Executar(('INSERT INTO metas (' + ($cols -join ', ') + ', alterado_em, alterado_por) VALUES (' + (($cols | ForEach-Object { '?' }) -join ',') + ',?,?)'), $par)
            $id = [int]$ctx.Valor('SELECT @@IDENTITY', @())
            $acao = 'INSERIR'
        } else {
            $id = [int]$ex[0].id
            $sets = @($mud.Keys | ForEach-Object { $_ + '=?' })
            $par = $valores + @((P 'data' $agora), (P 'texto' $script:UsuarioLog), (P 'inteiro' $id))
            [void]$ctx.Executar(('UPDATE metas SET ' + ($sets -join ', ') + ', alterado_em=?, alterado_por=? WHERE id=?'), $par)
            $acao = 'ALTERAR'
        }
        foreach ($c in $mud.Keys) { $ctx.Log('metas', $id, $c, $mud[$c].Antigo, $mud[$c].Novo, $acao) }
        return [ordered]@{ gravado = $true; ano = $ano; mes = $mes; alterados = @($mud.Keys | ForEach-Object { [ordered]@{ campo = $_; antes = (ConvertFrom-ValorBanco $mud[$_].Antigo); depois = (ConvertFrom-ValorBanco $mud[$_].Novo) } }) }
    }
}

# ================================================================= CATÁLOGO
$script:SoLeitura = [ordered]@{ readOnlyHint = $true; destructiveHint = $false; idempotentHint = $true; openWorldHint = $false }
$script:Gravacao  = [ordered]@{ readOnlyHint = $false; destructiveHint = $true; idempotentHint = $false; openWorldHint = $false }

$script:PropAutor = [ordered]@{ type = 'string'; maxLength = 40; description = 'Quem pediu a alteração (ex.: "Paulo"). Vai para o log junto com "IA".' }
$script:PropData  = [ordered]@{ type = @('string', 'null'); pattern = '^\d{4}-\d{2}-\d{2}$'; description = 'Data AAAA-MM-DD; null limpa o campo.' }

$script:Ferramentas = @(
    [ordered]@{
        name = 'criar_cliente'; title = 'Cadastrar empresa'; handler = 'Invoke-CriarCliente'; annotations = $script:Gravacao
        description = ('Cadastra uma empresa como pré-cliente (padrão, sem código) ou já como cliente (como_cliente=true, recebe o próximo código). ' +
                       'Mesmas críticas de atualizar_cadastro_cliente (CNPJ, telefone, listas). Recusa nome já existente, a menos que ignorar_homonimo=true. Procure antes com buscar_clientes.')
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; required = @('empresa'); properties = [ordered]@{
            empresa = [ordered]@{ type = 'string'; maxLength = 120 }; como_cliente = [ordered]@{ type = 'boolean'; default = $false }
            cnpj = [ordered]@{ type = @('string', 'null') }; cidade = [ordered]@{ type = @('string', 'null'); maxLength = 60 }
            uf = [ordered]@{ type = @('string', 'null') }; segmento = [ordered]@{ type = @('string', 'null') }
            telefone = [ordered]@{ type = @('string', 'null') }; email = [ordered]@{ type = @('string', 'null'); maxLength = 120 }
            responsavel = [ordered]@{ type = @('string', 'null') }; qualificacao = [ordered]@{ type = @('string', 'null'); maxLength = 40 }
            aderencia = [ordered]@{ type = @('integer', 'null'); minimum = 1; maximum = 3 }; porte = [ordered]@{ type = @('integer', 'null'); minimum = 1; maximum = 3 }
            observacoes = [ordered]@{ type = 'string' }; ignorar_homonimo = [ordered]@{ type = 'boolean'; default = $false }
            autor = $script:PropAutor } }
    },
    [ordered]@{
        name = 'promover_cliente'; title = 'Promover pré-cliente'; handler = 'Invoke-PromoverCliente'; annotations = $script:Gravacao
        description = 'Promove um pré-cliente a cliente e atribui o próximo código. Irreversível pelo sistema: confirme com o usuário antes.'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; required = @('id_cliente'); properties = [ordered]@{
            id_cliente = [ordered]@{ type = 'integer' }; autor = $script:PropAutor } }
    },
    [ordered]@{
        name = 'alterar_situacao_cliente'; title = 'Desativar, reativar ou excluir empresa'; handler = 'Invoke-SituacaoCliente'; annotations = $script:Gravacao
        description = ('desativar: a empresa e seus contatos ficam inativos (atendimentos em aberto continuam no funil). reativar: volta com os contatos inativos. ' +
                       'excluir_pre_cliente: apaga de vez um pré-cliente sem contato nem atendimento (cliente com código nunca é excluído). Leia antes com obter_cliente e passe a "versao"; confirme com o usuário.')
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; required = @('acao', 'versao'); properties = [ordered]@{
            codigo_cliente = [ordered]@{ type = 'integer' }; id_cliente = [ordered]@{ type = 'integer' }
            versao = [ordered]@{ type = 'integer' }
            acao = [ordered]@{ type = 'string'; enum = @('desativar', 'reativar', 'excluir_pre_cliente') }
            autor = $script:PropAutor } }
    },
    [ordered]@{
        name = 'abrir_atendimento'; title = 'Abrir atendimento'; handler = 'Invoke-AbrirAtendimento'; annotations = $script:Gravacao
        description = ('Abre um atendimento para um contato ativo e gera o código AT-CCCC-NNNN. Etapa e responsável obrigatórios; demais campos e críticas iguais a atualizar_atendimento. ' +
                       'dt_entrada padrão = hoje. Em pré-cliente exige promover_pre_cliente=true (promove a empresa). codigo_atendimento_anterior liga a retomada ao atendimento antigo.')
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; required = @('etapa', 'responsavel'); properties = [ordered]@{
            codigo_contato = [ordered]@{ type = 'string'; pattern = '^[Cc][Tt]-\d{4}-\d{4}$' }; id_contato = [ordered]@{ type = 'integer' }
            etapa = [ordered]@{ type = 'string' }; responsavel = [ordered]@{ type = 'string' }
            prioridade = [ordered]@{ type = 'string' }; origem = [ordered]@{ type = 'string' }
            familia = [ordered]@{ type = 'string' }; categoria = [ordered]@{ type = 'string' }
            motivo_desfecho = [ordered]@{ type = @('string', 'null') }; tipo_venda = [ordered]@{ type = @('string', 'null'); maxLength = 40 }
            maquina = [ordered]@{ type = @('string', 'null'); maxLength = 255 }; orcamento = [ordered]@{ type = @('string', 'null'); maxLength = 30 }
            valor = [ordered]@{ type = @('number', 'null'); minimum = 0 }; tentativas = [ordered]@{ type = @('integer', 'null'); minimum = 0 }
            prox_acao = [ordered]@{ type = @('string', 'null') }
            dt_entrada = $script:PropData; dt_prox_acao = $script:PropData; retomar_em = $script:PropData; ultima_interacao = $script:PropData
            dt_proposta = $script:PropData; dt_desfecho = $script:PropData
            codigo_atendimento_anterior = [ordered]@{ type = 'string'; pattern = '^AT-\d{4}-\d{4}$' }
            observacoes = [ordered]@{ type = 'string' }; promover_pre_cliente = [ordered]@{ type = 'boolean'; default = $false }
            autor = $script:PropAutor } }
    },
    [ordered]@{
        name = 'gerenciar_lista'; title = 'Manter listas suspensas'; handler = 'Invoke-GerenciarLista'; annotations = $script:Gravacao
        description = ('Inclui, ativa, desativa ou reordena um valor de lista suspensa (Etapa, Responsavel, Familia...). Valor nunca é apagado nem renomeado: registros antigos continuam com ele. ' +
                       'Desativar Etapa ou Responsavel afeta o funil e a fila de todos: confirme com o usuário.')
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; required = @('tipo', 'valor', 'acao'); properties = [ordered]@{
            tipo = [ordered]@{ type = 'string' }; valor = [ordered]@{ type = 'string'; maxLength = 120 }
            acao = [ordered]@{ type = 'string'; enum = @('incluir', 'ativar', 'desativar', 'ordenar') }
            ordem = [ordered]@{ type = 'integer' }; autor = $script:PropAutor } }
    },
    [ordered]@{
        name = 'listar_metas'; title = 'Metas mensais'; handler = 'Invoke-ListarMetas'; annotations = $script:SoLeitura
        description = 'Metas mensais do comercial (faturamento, pedidos, propostas, contatos).'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; properties = [ordered]@{ ano = [ordered]@{ type = 'integer' } } }
    },
    [ordered]@{
        name = 'gravar_meta'; title = 'Gravar meta mensal'; handler = 'Invoke-GravarMeta'; annotations = $script:Gravacao
        description = 'Cria ou altera a meta de um mês. Envie só os campos que mudam; null limpa.'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; required = @('ano', 'mes'); properties = [ordered]@{
            ano = [ordered]@{ type = 'integer' }; mes = [ordered]@{ type = 'integer'; minimum = 1; maximum = 12 }
            meta_faturamento = [ordered]@{ type = @('number', 'null'); minimum = 0 }
            meta_pedidos = [ordered]@{ type = @('integer', 'null'); minimum = 0 }
            meta_propostas = [ordered]@{ type = @('integer', 'null'); minimum = 0 }
            meta_contatos = [ordered]@{ type = @('integer', 'null'); minimum = 0 }
            autor = $script:PropAutor } }
    },
    [ordered]@{
        name = 'buscar_contatos'; title = 'Buscar contatos'; handler = 'Invoke-BuscarContatos'; annotations = $script:SoLeitura
        description = 'Procura pessoas (contatos) por nome, cargo, e-mail, telefone, código CT ou empresa. Por padrão só ativos.'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; properties = [ordered]@{
            texto = [ordered]@{ type = 'string'; description = 'Trecho do nome, cargo, e-mail, telefone, código CT ou empresa' }
            codigo_cliente = [ordered]@{ type = 'integer' }; id_cliente = [ordered]@{ type = 'integer' }
            incluir_inativos = [ordered]@{ type = 'boolean'; default = $false }
            limite = [ordered]@{ type = 'integer'; minimum = 1; maximum = 200; default = 20 } } }
    },
    [ordered]@{
        name = 'obter_contato'; title = 'Ficha do contato'; handler = 'Invoke-ObterContato'; annotations = $script:SoLeitura
        description = 'Ficha completa de um contato (CT-0000-0000 ou id), com a empresa e os atendimentos dele. Devolve "versao", necessária para atualizar_contato.'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; properties = [ordered]@{
            codigo_contato = [ordered]@{ type = 'string'; pattern = '^[Cc][Tt]-\d{4}-\d{4}$' }
            id_contato = [ordered]@{ type = 'integer' } } }
    },
    [ordered]@{
        name = 'criar_contato'; title = 'Cadastrar contato'; handler = 'Invoke-CriarContato'; annotations = $script:Gravacao
        description = ('Cadastra um contato numa empresa e gera o código CT-CCCC-NNNN (congelado). Nome obrigatório (GERAL para o contato geral da empresa); telefone só com DDD + número (10 ou 11 dígitos). ' +
                       'Em pré-cliente, a gravação é recusada a menos que promover_pre_cliente=true: cadastrar contato promove a empresa a cliente com o próximo código - confirme com o usuário antes. ' +
                       'Confira com buscar_contatos se a pessoa já não existe.')
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; required = @('nome'); properties = [ordered]@{
            codigo_cliente = [ordered]@{ type = 'integer' }; id_cliente = [ordered]@{ type = 'integer'; description = 'Use para pré-cliente' }
            nome = [ordered]@{ type = 'string'; maxLength = 120 }; cargo = [ordered]@{ type = @('string', 'null'); maxLength = 60 }
            telefone = [ordered]@{ type = @('string', 'null') }; email = [ordered]@{ type = @('string', 'null'); maxLength = 120 }
            observacoes = [ordered]@{ type = 'string'; description = 'Primeira linha datada das observações' }
            promover_pre_cliente = [ordered]@{ type = 'boolean'; default = $false }
            autor = $script:PropAutor } }
    },
    [ordered]@{
        name = 'atualizar_contato'; title = 'Alterar contato'; handler = 'Invoke-AtualizarContato'; annotations = $script:Gravacao
        description = ('Altera nome, cargo, telefone, e-mail, observações, empresa (codigo_cliente_destino) ou situação (ativo) de um contato. Leia antes com obter_contato e passe a "versao". ' +
                       'Contato não se exclui: ativo=false desativa. O código CT não muda ao trocar de empresa; atendimentos já abertos ficam onde estão. Envie só os campos que mudam.')
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; required = @('versao'); properties = [ordered]@{
            codigo_contato = [ordered]@{ type = 'string'; pattern = '^[Cc][Tt]-\d{4}-\d{4}$' }; id_contato = [ordered]@{ type = 'integer' }
            versao = [ordered]@{ type = 'integer'; description = 'Versão lida em obter_contato' }
            nome = [ordered]@{ type = 'string'; maxLength = 120 }; cargo = [ordered]@{ type = @('string', 'null'); maxLength = 60 }
            telefone = [ordered]@{ type = @('string', 'null') }; email = [ordered]@{ type = @('string', 'null'); maxLength = 120 }
            codigo_cliente_destino = [ordered]@{ type = 'integer'; description = 'Transfere o contato para esta empresa (precisa ser cliente ativo)' }
            ativo = [ordered]@{ type = 'boolean' }
            observacoes_acrescentar = [ordered]@{ type = 'string'; description = 'Acrescenta uma linha datada (nunca apaga o que existe)' }
            autor = $script:PropAutor } }
    },
    [ordered]@{
        name = 'crm_status'; title = 'Situação do banco do CRM'; handler = 'Invoke-CrmStatus'; annotations = $script:SoLeitura
        description = 'Mostra onde está o banco, a versão do esquema, se a gravação está liberada e a quantidade de registros. Use primeiro para confirmar a conexão.'
        inputSchema = [ordered]@{ type = 'object'; properties = [ordered]@{}; additionalProperties = $false }
    },
    [ordered]@{
        name = 'listar_opcoes'; title = 'Opções das listas suspensas'; handler = 'Invoke-ListarOpcoes'; annotations = $script:SoLeitura
        description = 'Valores aceitos em campos de lista (Etapa, Responsavel, Familia, Origem, Prioridade, Motivo, Segmento, UF, Categoria). Sem "tipo", devolve os tipos existentes. Consulte antes de gravar um campo de lista.'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; properties = [ordered]@{
            tipo = [ordered]@{ type = 'string'; description = 'Nome da lista, ex.: Etapa' } } }
    },
    [ordered]@{
        name = 'buscar_clientes'; title = 'Buscar clientes'; handler = 'Invoke-BuscarClientes'; annotations = $script:SoLeitura
        description = 'Procura empresas por nome, cidade, segmento, CNPJ ou código (até 4 dígitos). Pré-clientes não têm código: use id_cliente para eles.'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; properties = [ordered]@{
            texto = [ordered]@{ type = 'string'; description = 'Trecho do nome, cidade, segmento, CNPJ ou o código do cliente' }
            estagio = [ordered]@{ type = 'string'; enum = @('Cliente', ('Pr' + [char]0x00E9 + '-cliente')) }
            responsavel = [ordered]@{ type = 'string' }
            uf = [ordered]@{ type = 'string'; maxLength = 2 }
            incluir_inativos = [ordered]@{ type = 'boolean'; default = $false }
            limite = [ordered]@{ type = 'integer'; minimum = 1; maximum = 200; default = 20 } } }
    },
    [ordered]@{
        name = 'obter_cliente'; title = 'Ficha do cliente'; handler = 'Invoke-ObterCliente'; annotations = $script:SoLeitura
        description = 'Ficha completa de uma empresa, com contatos e atendimentos. Devolve "versao", necessária para atualizar_cadastro_cliente.'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; properties = [ordered]@{
            codigo_cliente = [ordered]@{ type = 'integer'; description = 'Código do cliente (ex.: 11 para 0011)' }
            id_cliente = [ordered]@{ type = 'integer'; description = 'Id interno (use para pré-clientes)' } } }
    },
    [ordered]@{
        name = 'buscar_atendimentos'; title = 'Buscar atendimentos'; handler = 'Invoke-BuscarAtendimentos'; annotations = $script:SoLeitura
        description = 'Lista atendimentos (oportunidades) com filtros. Por padrão só os em aberto, ordenados pela próxima ação. Traz a situação calculada igual à do front.'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; properties = [ordered]@{
            texto = [ordered]@{ type = 'string'; description = 'Trecho do código AT, empresa, contato, máquina ou orçamento' }
            etapa = [ordered]@{ type = 'string' }
            responsavel = [ordered]@{ type = 'string' }
            codigo_cliente = [ordered]@{ type = 'integer' }
            situacao = [ordered]@{ type = 'string'; enum = @(('A' + [char]0x00E7 + [char]0x00E3 + 'o atrasada'), 'Retomar hoje',
                                   ('A' + [char]0x00E7 + [char]0x00E3 + 'o nesta semana'), 'Em dia',
                                   ('Sem pr' + [char]0x00F3 + 'xima a' + [char]0x00E7 + [char]0x00E3 + 'o'), 'Encerrado', 'Sem etapa definida') }
            somente_abertos = [ordered]@{ type = 'boolean'; default = $true }
            limite = [ordered]@{ type = 'integer'; minimum = 1; maximum = 200; default = 20 } } }
    },
    [ordered]@{
        name = 'obter_atendimento'; title = 'Ficha do atendimento'; handler = 'Invoke-ObterAtendimento'; annotations = $script:SoLeitura
        description = 'Ficha completa de um atendimento pelo código AT-0000-0000, com empresa, contato e campos calculados. Devolve "versao", necessária para atualizar_atendimento.'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; required = @('codigo'); properties = [ordered]@{
            codigo = [ordered]@{ type = 'string'; pattern = '^AT-\d{4}-\d{4}$' } } }
    },
    [ordered]@{
        name = 'ultimas_alteracoes'; title = 'Histórico de alterações'; handler = 'Invoke-UltimasAlteracoes'; annotations = $script:SoLeitura
        description = 'Últimas linhas do log de auditoria (quem alterou o quê, de onde: FRONT ou IA).'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; properties = [ordered]@{
            tabela = [ordered]@{ type = 'string'; enum = @('clientes', 'contatos', 'oportunidades', 'listas', 'metas') }
            id_registro = [ordered]@{ type = 'integer' }
            desde_id = [ordered]@{ type = 'integer'; description = 'Só alterações com id maior que este' }
            limite = [ordered]@{ type = 'integer'; minimum = 1; maximum = 200; default = 30 } } }
    },
    [ordered]@{
        name = 'atendimentos_por_interacao'; title = 'Atendimentos por última interação'; handler = 'Invoke-AtendimentosPorInteracao'; annotations = $script:SoLeitura
        description = ('Atendimentos abertos e encerrados com última interação dentro de um período, numa única chamada, para o relatório semanal. ' +
                       'Período: semana ISO (semana + ano, segunda a domingo) OU data_inicio + data_fim, inclusive as pontas; sem nenhum dos dois, a semana ISO anterior à data atual. ' +
                       'Ordena por responsável e última interação; traz resumo por responsável (total, por etapa e por situação). ' +
                       'JSON compacto sem campos vazios; completo=true inclui os campos técnicos. Se passar do limite de tamanho, devolve truncado=true com a quantidade real: refaça por responsável.')
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; properties = [ordered]@{
            semana = [ordered]@{ type = 'integer'; minimum = 1; maximum = 53; description = 'Semana ISO 8601 (1 a 53). Exige ano.' }
            ano = [ordered]@{ type = 'integer'; minimum = 2000; maximum = 2100; description = 'Ano da semana ISO. Exige semana.' }
            data_inicio = [ordered]@{ type = 'string'; pattern = '^\d{4}-\d{2}-\d{2}$'; description = 'AAAA-MM-DD. Alternativa a semana/ano; exige data_fim.' }
            data_fim = [ordered]@{ type = 'string'; pattern = '^\d{4}-\d{2}-\d{2}$'; description = 'AAAA-MM-DD, inclusive. Exige data_inicio.' }
            responsavel = [ordered]@{ type = 'string'; maxLength = 60; description = 'Filtra um responsável (ex.: "Paulo").' }
            completo = [ordered]@{ type = 'boolean'; default = $false; description = 'true devolve todos os campos do atendimento, inclusive os técnicos.' } } }
    },
    [ordered]@{
        name = 'atendimentos_por_proxima_acao'; title = 'Atendimentos por próxima ação (planejamento)'; handler = 'Invoke-AtendimentosPorProximaAcao'; annotations = $script:SoLeitura
        description = ('Atendimentos em aberto com próxima ação num período de hoje em diante, para o relatório de planejamento. ' +
                       'Período: data_inicio + data_fim OU semana + ano; sem período, de hoje até domingo. ' +
                       'Período encerrado é recusado; período iniciado antes de hoje começa hoje (ajustado). ' +
                       'incluir_vencidas traz as ações anteriores a hoje marcadas como vencida. ' +
                       'Resumo por responsável e da equipe (só na primeira página), com contagem por dia, vencidas, valor no funil e atendimentos sem próxima ação. ' +
                       'Se passar do limite de tamanho, devolve truncado=true e proximo: repita com a_partir_de=proximo.')
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; properties = [ordered]@{
            data_inicio = [ordered]@{ type = 'string'; pattern = '^\d{4}-\d{2}-\d{2}$'; description = 'AAAA-MM-DD, inclusive. Exige data_fim. Anterior a hoje vira hoje (ajustado).' }
            data_fim = [ordered]@{ type = 'string'; pattern = '^\d{4}-\d{2}-\d{2}$'; description = 'AAAA-MM-DD, inclusive. Exige data_inicio. Anterior a hoje é recusada. Período de até 92 dias.' }
            semana = [ordered]@{ type = 'integer'; minimum = 1; maximum = 53; description = 'Semana ISO 8601 (1 a 53). Exige ano. Semana encerrada é recusada; a corrente começa hoje.' }
            ano = [ordered]@{ type = 'integer'; minimum = 2000; maximum = 2100; description = 'Ano da semana ISO. Exige semana.' }
            responsavel = [ordered]@{ type = 'string'; maxLength = 60; description = 'Filtra um responsável (ex.: "Paulo").' }
            incluir_vencidas = [ordered]@{ type = 'boolean'; default = $true; description = 'Inclui os abertos com próxima ação anterior a hoje, com vencida=true.' }
            completo = [ordered]@{ type = 'boolean'; default = $false; description = 'true devolve todos os campos do atendimento, inclusive os técnicos.' }
            a_partir_de = [ordered]@{ type = 'integer'; minimum = 0; default = 0; description = 'Índice do primeiro registro devolvido (paginação): use o proximo da chamada anterior.' } } }
    },
    [ordered]@{
        name = 'exportar_base'; title = 'Extrato da base em CSV'; handler = 'Invoke-ExportarBase'
        annotations = [ordered]@{ readOnlyHint = $false; destructiveHint = $false; idempotentHint = $true; openWorldHint = $false }
        description = ('Grava a foto das tabelas do CRM em CSV (crm-<tabela>.csv + crm-extrato.json) numa pasta da pasta do setor, para análise por script. ' +
                       'Não altera o banco. Devolve só o envelope: data do extrato, arquivos, linhas e colunas por tabela - nunca os dados. ' +
                       'Atendimentos saem com codigo_cliente, empresa, codigo_contato e os calculados situacao, quadro, dias_parado e ciclo (mesma regra do front, na data do extrato). ' +
                       'CSV UTF-8, separador vírgula, campos entre aspas; datas AAAA-MM-DD ou AAAA-MM-DDTHH:MM:SS. Regrava os arquivos a cada chamada.')
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; required = @('destino'); properties = [ordered]@{
            destino = [ordered]@{ type = 'string'; maxLength = 255; description = 'Pasta relativa a 1 - COMERCIAL ZAPROMAQ, sem letra de unidade (ex.: 00 - GESTÃO COMERCIAL\01 - PADROES\00 - IA\PROJETOS\PRJ-COM-006-analise-de-dados-crm\20-dados). Precisa existir; 01 - CRM\01 - CONTROLE é recusado.' }
            tabelas = [ordered]@{ type = 'array'; items = [ordered]@{ type = 'string' }; maxItems = 6; description = 'clientes, contatos, oportunidades, log_alteracoes, listas, metas. Padrão: clientes, contatos, oportunidades.' } } }
    },
    [ordered]@{
        name = 'atualizar_atendimento'; title = 'Atualizar atendimento'; handler = 'Invoke-AtualizarAtendimento'; annotations = $script:Gravacao
        description = ('Altera campos de um atendimento. Leia antes com obter_atendimento e passe a "versao" lida: se alguém alterou depois, a gravação é recusada em vez de sobrescrever. ' +
                       'Envie só os campos que mudam. Campos de lista precisam de valor existente (listar_opcoes). Datas em AAAA-MM-DD. ' +
                       'Etapa Pedido Fechado/Perdido/Descartado exige motivo_desfecho e dt_desfecho.')
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; required = @('codigo', 'versao'); properties = [ordered]@{
            codigo = [ordered]@{ type = 'string'; pattern = '^AT-\d{4}-\d{4}$' }
            versao = [ordered]@{ type = 'integer'; description = 'Versão lida em obter_atendimento' }
            etapa = [ordered]@{ type = 'string' }; responsavel = [ordered]@{ type = 'string' }
            prioridade = [ordered]@{ type = 'string' }; origem = [ordered]@{ type = 'string' }
            familia = [ordered]@{ type = 'string' }; categoria = [ordered]@{ type = 'string' }
            motivo_desfecho = [ordered]@{ type = @('string', 'null') }; tipo_venda = [ordered]@{ type = @('string', 'null'); maxLength = 40 }
            maquina = [ordered]@{ type = @('string', 'null'); maxLength = 255 }; orcamento = [ordered]@{ type = @('string', 'null'); maxLength = 30 }
            valor = [ordered]@{ type = @('number', 'null'); minimum = 0 }; tentativas = [ordered]@{ type = @('integer', 'null'); minimum = 0 }
            prox_acao = [ordered]@{ type = @('string', 'null'); description = 'Texto da próxima ação' }
            dt_prox_acao = $script:PropData; retomar_em = $script:PropData; ultima_interacao = $script:PropData
            dt_proposta = $script:PropData; dt_desfecho = $script:PropData; dt_entrada = $script:PropData
            codigo_contato = [ordered]@{ type = 'string'; pattern = '^[Cc][Tt]-\d{4}-\d{4}$'; description = 'Troca o contato (só da mesma empresa)' }
            observacoes_acrescentar = [ordered]@{ type = 'string'; description = 'Acrescenta uma linha datada às observações (nunca apaga o que existe)' }
            autor = $script:PropAutor } }
    },
    [ordered]@{
        name = 'atualizar_cadastro_cliente'; title = 'Corrigir cadastro de cliente'; handler = 'Invoke-AtualizarCadastroCliente'; annotations = $script:Gravacao
        description = ('Corrige dados cadastrais de uma empresa. Leia antes com obter_cliente e passe a "versao". CNPJ é conferido (dígito e duplicidade); telefone só com DDD + número (10 ou 11 dígitos), ' +
                       'um por campo. Código e estágio do cliente não são alterados por aqui.')
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; required = @('versao'); properties = [ordered]@{
            codigo_cliente = [ordered]@{ type = 'integer' }; id_cliente = [ordered]@{ type = 'integer' }
            versao = [ordered]@{ type = 'integer'; description = 'Versão lida em obter_cliente' }
            empresa = [ordered]@{ type = 'string'; maxLength = 120 }; cnpj = [ordered]@{ type = @('string', 'null') }
            cidade = [ordered]@{ type = @('string', 'null'); maxLength = 60 }; uf = [ordered]@{ type = @('string', 'null') }
            segmento = [ordered]@{ type = @('string', 'null') }; telefone = [ordered]@{ type = @('string', 'null') }
            email = [ordered]@{ type = @('string', 'null'); maxLength = 120 }; responsavel = [ordered]@{ type = @('string', 'null') }
            qualificacao = [ordered]@{ type = @('string', 'null'); maxLength = 40 }
            aderencia = [ordered]@{ type = @('integer', 'null'); minimum = 1; maximum = 3 }; porte = [ordered]@{ type = @('integer', 'null'); minimum = 1; maximum = 3 }
            observacoes_acrescentar = [ordered]@{ type = 'string' }
            autor = $script:PropAutor } }
    },
    [ordered]@{
        name = 'atualizar_contexto_cliente'; title = 'Atualizar contexto do cliente'; handler = 'Invoke-AtualizarContextoCliente'; annotations = $script:Gravacao
        description = 'Grava o resumo de contexto da empresa (ctx_*): síntese de uma linha, família aderente e caminho do CONTEXTO-GERAL.md. Campos exclusivos da IA; não conflitam com o vendedor.'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; properties = [ordered]@{
            codigo_cliente = [ordered]@{ type = 'integer' }; id_cliente = [ordered]@{ type = 'integer' }
            ctx_resumo = [ordered]@{ type = @('string', 'null'); maxLength = 255 }
            ctx_familia = [ordered]@{ type = @('string', 'null'); maxLength = 60 }
            ctx_arquivo = [ordered]@{ type = @('string', 'null'); maxLength = 255; description = 'Caminho relativo do CONTEXTO-GERAL.md' }
            autor = $script:PropAutor } }
    },
    [ordered]@{
        name = 'atualizar_contexto_atendimento'; title = 'Atualizar contexto do atendimento'; handler = 'Invoke-AtualizarContextoAtendimento'; annotations = $script:Gravacao
        description = 'Grava o resumo de contexto de um atendimento (ctx_resumo e caminho do CONTEXTO-ATENDIMENTO.md). Campos exclusivos da IA.'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; required = @('codigo'); properties = [ordered]@{
            codigo = [ordered]@{ type = 'string'; pattern = '^AT-\d{4}-\d{4}$' }
            ctx_resumo = [ordered]@{ type = @('string', 'null'); maxLength = 255 }
            ctx_arquivo = [ordered]@{ type = @('string', 'null'); maxLength = 255 }
            autor = $script:PropAutor } }
    }
)
