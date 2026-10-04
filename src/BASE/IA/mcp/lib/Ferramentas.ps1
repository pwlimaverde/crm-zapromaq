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
    ctx_resumo = 255; ctx_familia = 60; ctx_arquivo = 255
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
    foreach ($t in @('clientes', 'contatos', 'oportunidades', 'log_alteracoes')) {
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

# ================================================================= GRAVAÇÃO
function Invoke-AtualizarAtendimento($a) {
    Assert-EsquemaGravavel
    $codigo = [string]$a.codigo
    return Invoke-Transacao {
        param($ctx)
        $atual = Get-AtendimentoNaTransacao $ctx $codigo
        Assert-Versao $atual $a

        # campo -> tipo de parâmetro e lista (quando for lista suspensa)
        $defs = [ordered]@{
            etapa = @('texto', 'Etapa'); responsavel = @('texto', 'Responsavel'); prioridade = @('texto', 'Prioridade')
            origem = @('texto', 'Origem'); familia = @('texto', 'Familia'); categoria = @('texto', 'Categoria')
            motivo_desfecho = @('texto', 'Motivo'); tipo_venda = @('texto', $null); maquina = @('texto', $null)
            orcamento = @('texto', $null); valor = @('moeda', $null); tentativas = @('inteiro', $null)
            prox_acao = @('memo', $null); dt_proposta = @('data', $null); ultima_interacao = @('data', $null)
            dt_prox_acao = @('data', $null); retomar_em = @('data', $null); dt_desfecho = @('data', $null)
        }
        $novo = [ordered]@{}; foreach ($k in $atual.Keys) { $novo[$k] = $atual[$k] }
        $mud = [ordered]@{}
        foreach ($campo in $defs.Keys) {
            if (-not (Test-TemArg $a $campo)) { continue }
            $tipo = $defs[$campo][0]; $lista = $defs[$campo][1]
            $v = $a.$campo
            switch ($tipo) {
                'data'    { $v = ConvertTo-Data $campo $v }
                'moeda'   { if ($null -ne $v) { $v = [decimal]$v; if ($v -lt 0) { throw 'O valor não pode ser negativo.' } } }
                'inteiro' { if ($null -ne $v) { $v = [int]$v; if ($v -lt 0) { throw "Campo '$campo' não pode ser negativo." } } }
                default   {
                    if ($lista) { $v = Resolve-ValorLista $ctx $lista $campo $v }
                    else { $v = ConvertTo-Normalizado $campo $v; Assert-Tamanho $campo $v }
                }
            }
            $novo[$campo] = $v
            if (-not (Test-Igual $atual[$campo] $v)) { $mud[$campo] = @{ Tipo = $tipo; Novo = $v; Antigo = $atual[$campo] } }
        }
        if (Test-TemArg $a 'observacoes_acrescentar') {
            $obs = Add-Observacao $atual.observacoes ([string]$a.observacoes_acrescentar)
            $novo['observacoes'] = $obs
            $mud['observacoes'] = @{ Tipo = 'memo'; Novo = $obs; Antigo = $atual.observacoes }
        }

        # --- as mesmas críticas do front (modValidacao.CriticarOportunidade)
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

        if ($mud.Count -eq 0) {
            return [ordered]@{ codigo = $codigo; gravado = $false; mensagem = 'Nada a alterar: os valores já estão iguais.'; versao = $atual.versao }
        }
        Update-ComVersao $ctx 'oportunidades' ([int]$atual.id) ([int]$atual.versao) $mud 'ALTERAR'
        return [ordered]@{
            codigo = $codigo; gravado = $true; versao = ([int]$atual.versao + 1)
            alterados = @($mud.Keys | ForEach-Object { [ordered]@{ campo = $_; antes = (ConvertFrom-ValorBanco $mud[$_].Antigo); depois = (ConvertFrom-ValorBanco $mud[$_].Novo) } })
            situacao = Get-Situacao $true $etapa $novo.dt_prox_acao $novo.retomar_em
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

# ================================================================= CATÁLOGO
$script:SoLeitura = [ordered]@{ readOnlyHint = $true; destructiveHint = $false; idempotentHint = $true; openWorldHint = $false }
$script:Gravacao  = [ordered]@{ readOnlyHint = $false; destructiveHint = $true; idempotentHint = $false; openWorldHint = $false }

$script:PropAutor = [ordered]@{ type = 'string'; maxLength = 40; description = 'Quem pediu a alteração (ex.: "Paulo"). Vai para o log junto com "IA".' }
$script:PropData  = [ordered]@{ type = @('string', 'null'); pattern = '^\d{4}-\d{2}-\d{2}$'; description = 'Data AAAA-MM-DD; null limpa o campo.' }

$script:Ferramentas = @(
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
            tabela = [ordered]@{ type = 'string'; enum = @('clientes', 'contatos', 'oportunidades') }
            id_registro = [ordered]@{ type = 'integer' }
            desde_id = [ordered]@{ type = 'integer'; description = 'Só alterações com id maior que este' }
            limite = [ordered]@{ type = 'integer'; minimum = 1; maximum = 200; default = 30 } } }
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
            dt_proposta = $script:PropData; dt_desfecho = $script:PropData
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
