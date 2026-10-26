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
    ctx_resumo = 255; ctx_familia = 255; ctx_arquivo = 255; nome = 120; cargo = 60
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
# Valor INATIVO só passa se for o próprio valor já gravado no registro ($atual): D-j, C06.
function Resolve-ValorLista($ctx, [string]$tipo, [string]$campo, $valor, $atual = $null) {
    if ($null -eq $valor -or [string]$valor -eq '') { return $null }
    $validos = Get-ValoresLista $ctx $tipo
    $texto = ([string]$valor).Trim()
    foreach ($v in $validos) { if ($v -ieq $texto) { return $v } }
    if ($null -ne $atual -and [string]$atual -ieq $texto) { return [string]$atual }
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
    # seção 5.1: o texto comercial acrescentado entra em maiúsculas (URL, caminho e e-mail preservados)
    $linha = '[' + (Get-Date).ToString('dd/MM/yyyy') + ' ' + $script:UsuarioLog + '] ' + (ConvertTo-TextoComercial $texto.Trim() $true)
    if ($null -eq $atual -or [string]$atual -eq '') { return $linha }
    return ([string]$atual).TrimEnd() + "`r`n" + $linha
}

# DEV-13 (esquema 1.2): identificador de operação. A mesma criação repetida com o MESMO id_operacao e o mesmo
# conteúdo devolve o resultado da primeira (sem gravar de novo nem gerar log); o mesmo id com outro conteúdo é
# recusado. Gravado na tabela operacoes na MESMA transação do dado. Guardado 60 dias.
$script:DiasOperacao = 60
function Get-HashOperacao($a) {
    $partes = @(); foreach ($p in ($a.PSObject.Properties | Where-Object { $_.Name -notin @('id_operacao', 'autor') } | Sort-Object Name)) { $partes += ($p.Name + '=' + (ConvertTo-Json -InputObject $p.Value -Compress -Depth 5)) }
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($partes -join "`n")))).Replace('-', '').ToLowerInvariant() } finally { $sha.Dispose() }
}
function Get-OperacaoAnterior($ctx, $a, [string]$tabela) {
    if (-not (Test-TemArg $a 'id_operacao')) { return $null }
    $l = $ctx.Consultar('SELECT tabela, id_registro, codigo, versao, hash_conteudo FROM operacoes WHERE id_operacao=?', @((P 'texto' ([string]$a.id_operacao))))
    if ($l.Count -eq 0) { return $null }
    if ([string]$l[0].hash_conteudo -ne (Get-HashOperacao $a) -or [string]$l[0].tabela -ne $tabela) {
        throw ('O id_operacao ' + $a.id_operacao + ' já foi usado com outro conteúdo. Nada foi gravado: use um identificador novo para uma operação nova.')
    }
    $r = [ordered]@{ gravado = $false; repetida = $true; id_operacao = [string]$a.id_operacao; id = [int]$l[0].id_registro; codigo = $l[0].codigo; versao = $l[0].versao
                     mensagem = 'Operação já aplicada antes com este id_operacao: nada foi gravado de novo (resultado da primeira).' }
    switch ($tabela) {
        'contatos' { $r['id_contato'] = $r.id; $r['codigo_contato'] = $r.codigo }
        'clientes' { $r['id_cliente'] = $r.id; $r['codigo_cliente'] = $r.codigo }
    }
    return $r
}
function Save-Operacao($ctx, $a, [string]$tabela, [int]$id, $codigo, [int]$versao) {
    if (-not (Test-TemArg $a 'id_operacao')) { return }
    [void]$ctx.Executar('DELETE FROM operacoes WHERE criado_em < ?', @((P 'data' ((Get-Date).AddDays(-$script:DiasOperacao)))))
    [void]$ctx.Executar('INSERT INTO operacoes (id_operacao, tabela, id_registro, codigo, versao, hash_conteudo, criado_em, origem, estacao) VALUES (?,?,?,?,?,?,?,?,?)',
        @((P 'texto' ([string]$a.id_operacao)), (P 'texto' $tabela), (P 'inteiro' $id), (P 'texto' ([string]$codigo)), (P 'inteiro' $versao),
          (P 'texto' (Get-HashOperacao $a)), (P 'data' (Get-Date)), (P 'texto' 'IA'), (P 'texto' ([Environment]::MachineName))))
}

# Seção 5.1: e-mails além do primeiro, recebidos em "email", vão para as observações (na mesma gravação).
function Add-EmailsNasObservacoes($a, $atual, $mud) {
    if (-not (Test-TemArg $a 'email')) { return }
    [void](ConvertTo-Normalizado 'email' $a.email)
    if (-not $script:EmailsAdicionais) { return }
    $base = if ($mud.Contains('observacoes')) { $mud['observacoes'].Novo } else { $atual.observacoes }
    $novo = Add-EmailsAdicionais $base $script:EmailsAdicionais
    if (-not (Test-Igual $atual.observacoes $novo)) { $mud['observacoes'] = @{ Tipo = 'memo'; Novo = $novo; Antigo = $atual.observacoes } }
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

# Aviso de backup por janela (seção 5.3, ciclo 007). CÓPIA FIEL das funções de SISTEMA\DADOS\lib\Backup.ps1
# (o conector publicado numa pasta de versão não depende de SISTEMA em tempo de execução); o verificar.py
# confere que os textos são iguais.
$script:JanelasBackup = @('09:00', '12:30', '16:00')
$script:JanelaEncerraMin = 30
$script:JanelaContaMin = 60
$script:FimDoDiaUtil = '16:30'   # depois disto, o dia útil tem de ter cópia completa validada

function Test-DiaUtil([datetime]$d) { return ($d.DayOfWeek -ne 'Saturday' -and $d.DayOfWeek -ne 'Sunday') }
function Get-InicioJanela([datetime]$dia, [string]$hhmm) { $p = $hhmm.Split(':'); return $dia.Date.AddHours([int]$p[0]).AddMinutes([int]$p[1]) }

# Cópias da pasta, pelo NOME: { instante; tipo (diario|periodico); validado }.
function Get-CopiasBackup([string]$pasta) {
    $lista = New-Object System.Collections.ArrayList
    foreach ($f in @(Get-ChildItem -LiteralPath $pasta -Filter 'crm_*.accdb' -File -ErrorAction SilentlyContinue)) {
        if ($f.Name -match '^crm_(diario|periodico)_(\d{8})-(\d{4}|\d{6})(-\d+)?_') {
            $hora = $Matches[3]; if ($hora.Length -eq 4) { $hora += '00' }
            $d = [datetime]::MinValue
            if ([datetime]::TryParseExact($Matches[2] + $hora, 'yyyyMMddHHmmss', [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$d)) {
                [void]$lista.Add([pscustomobject]@{ instante = $d; tipo = $Matches[1]; validado = (Test-Path -LiteralPath ($f.FullName + '.validado')) })
            }
        }
    }
    return , $lista
}

# Regra pura do aviso (seção 5.3): (a) a última janela encerrada não tem cópia nenhuma; (b) o último dia útil
# cujas 16h30 já passaram não tem cópia COMPLETA VALIDADA (a simples não apaga este aviso).
function Get-AvisoBackupJanela([datetime]$agora, $copias) {
    $motivos = New-Object System.Collections.ArrayList
    $copias = @($copias)
    # (a)
    $ultima = $null
    for ($k = 0; $k -le 10 -and $null -eq $ultima; $k++) {
        $dia = $agora.Date.AddDays(-$k)
        if (-not (Test-DiaUtil $dia)) { continue }
        foreach ($h in ($script:JanelasBackup | Sort-Object -Descending)) {
            $ini = Get-InicioJanela $dia $h
            if ($agora -ge $ini.AddMinutes($script:JanelaEncerraMin)) { $ultima = $ini; break }
        }
    }
    if ($null -ne $ultima) {
        $tem = @($copias | Where-Object { $_.instante -ge $ultima -and $_.instante -lt $ultima.AddMinutes($script:JanelaContaMin) }).Count -gt 0
        if (-not $tem) { [void]$motivos.Add('a janela das ' + $ultima.ToString('HH:mm') + ' de ' + $ultima.ToString('dd/MM/yyyy') + ' terminou sem cópia') }
    }
    # (b)
    $diaB = $null
    for ($k = 0; $k -le 10 -and $null -eq $diaB; $k++) {
        $dia = $agora.Date.AddDays(-$k)
        if ((Test-DiaUtil $dia) -and $agora -ge (Get-InicioJanela $dia $script:FimDoDiaUtil)) { $diaB = $dia }
    }
    if ($null -ne $diaB) {
        $tem = @($copias | Where-Object { $_.tipo -eq 'diario' -and $_.validado -and $_.instante.Date -eq $diaB }).Count -gt 0
        if (-not $tem) { [void]$motivos.Add('o dia ' + $diaB.ToString('dd/MM/yyyy') + ' terminou sem cópia completa validada') }
    }
    $msg = if ($motivos.Count) { 'Backup atrasado: ' + ($motivos -join '; ') + '. Avise quem cuida dos backups (FAZER-BACKUP.bat).' } else { 'Backups em dia.' }
    return [pscustomobject]@{ alerta = ($motivos.Count -gt 0); motivos = @($motivos); mensagem = $msg }
}

# Último backup VALIDADO (marca <backup>.validado, gravada só depois de restaurar uma cópia numa pasta
# descartável e conferir) e o aviso por janela.
function Get-StatusBackup {
    $pasta = Join-Path (Split-Path -Parent $script:CaminhoBanco) 'SISTEMA\DADOS\execucao\backup'
    $melhor = $null
    if (Test-Path -LiteralPath $pasta) {
        foreach ($f in @(Get-ChildItem -LiteralPath $pasta -Filter 'crm_*.validado' -File -ErrorAction SilentlyContinue)) {
            try {
                $j = Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
                $d = [datetime]::Parse([string]$j.validado_utc, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AdjustToUniversal -bor [Globalization.DateTimeStyles]::AssumeUniversal).ToLocalTime()
                if ($null -eq $melhor -or $d -gt $melhor.instante) { $melhor = [pscustomobject]@{ arquivo = ($f.Name -replace '\.validado$', ''); instante = $d } }
            } catch { }
        }
    }
    $aviso = Get-AvisoBackupJanela (Get-Date) (Get-CopiasBackup $pasta)
    $ultimo = $null
    if ($melhor) { $ultimo = [ordered]@{ arquivo = $melhor.arquivo; validado_em = $melhor.instante.ToString('yyyy-MM-dd HH:mm:ss'); horas = [math]::Round(((Get-Date) - $melhor.instante).TotalHours, 1) } }
    return [ordered]@{ ultimo = $ultimo; alerta = $aviso.alerta; motivos = @($aviso.motivos); mensagem = $aviso.mensagem }
}

# ================================================================= LEITURA
function Invoke-CrmStatus($a) {
    $cont = [ordered]@{}
    foreach ($t in @('clientes', 'contatos', 'oportunidades', 'listas', 'metas', 'log_alteracoes')) {
        $cont[$t] = [int](Invoke-Valor ('SELECT COUNT(*) FROM ' + $t) @())
    }
    $esq = Get-VersaoEsquema
    $bkp = Get-StatusBackup
    return [ordered]@{
        ultimo_backup_validado = $bkp.ultimo
        alerta_backup       = $bkp.alerta
        mensagem_backup     = $bkp.mensagem
        motivos_backup      = @($bkp.motivos)
        banco               = $script:CaminhoBanco
        provedor            = Get-ProvedorAce
        versao_esquema      = $esq
        gravacao_liberada   = ($script:EsquemasSuportados -contains $esq)
        atualizacao_pendente = (-not ($script:EsquemasSuportados -contains $esq))
        esquemas_suportados = @($script:EsquemasSuportados)
        versao_front        = Invoke-Valor "SELECT valor FROM config WHERE chave='versao_front'" @()
        ambiente            = Invoke-Valor "SELECT valor FROM config WHERE chave='ambiente'" @()
        registros           = $cont
        ultimo_log_id       = Invoke-Valor 'SELECT MAX(id) FROM log_alteracoes' @()
        versao_servidor_mcp = $script:VersaoServidor
        # DEV-07: o consumidor (skill) confere versão e capacidades antes de usar uma ferramenta nova
        versao_conector     = $script:VersaoServidor
        capacidades         = @($script:Ferramentas | ForEach-Object { [string]$_.name } | Sort-Object)
        quantidade_ferramentas = @($script:Ferramentas).Count
        pasta_conector      = $script:PastaConector   # DEV-05: de onde esta sessão carregou o conector
        # DEV-06 parte A: falha ao gravar o diagnóstico aparece aqui, sem bloquear nada
        diagnostico         = [ordered]@{ arquivo = $script:ArquivoLog; falhas = $script:DiagFalhas; ultima_falha = $script:DiagUltimaFalha }
    }
}

# Espelho de Get-ColunasDaLista (SISTEMA\DADOS\lib\Comum.ps1): onde cada lista é usada. Mudou lá, muda aqui.
$script:ColunasDaLista = @{
    UF = @('clientes:uf'); Segmento = @('clientes:segmento'); Qualificacao = @('clientes:qualificacao')
    Responsavel = @('clientes:responsavel', 'oportunidades:responsavel')
    Etapa = @('oportunidades:etapa'); Familia = @('oportunidades:familia'); Categoria = @('oportunidades:categoria')
    Origem = @('oportunidades:origem'); Prioridade = @('oportunidades:prioridade'); Motivo = @('oportunidades:motivo_desfecho')
}

function Invoke-ListarOpcoes($a) {
    $tipo = [string](Get-Arg $a 'tipo' '')
    if ($tipo -eq '') {
        $tipos = Invoke-Consulta 'SELECT DISTINCT tipo FROM listas ORDER BY tipo' @()
        return [ordered]@{ tipos = @($tipos | ForEach-Object { $_.tipo }) }
    }
    if (-not ((Test-TemArg $a 'detalhar') -and $a.detalhar -eq $true)) {
        return [ordered]@{ tipo = $tipo; valores = @(Get-ValoresLista $null $tipo) }
    }
    # detalhar: inativos também, com o uso (atendimentos e clientes) e, na etapa, o papel
    $linhas = Invoke-Consulta 'SELECT valor, ordem, ativo, encerra, resultado, quadro, exige_familia FROM listas WHERE tipo=? ORDER BY ordem, valor' @((P 'texto' $tipo))
    # uso: uma leitura agrupada por coluna (nunca uma conexão por valor). Tabela e coluna vêm do mapa fixo acima.
    $usoPorValor = @{}
    foreach ($tc in @($script:ColunasDaLista[$tipo])) {
        $parte = $tc.Split(':')
        foreach ($g in (Invoke-Consulta ('SELECT ' + $parte[1] + ' AS v, COUNT(*) AS n FROM ' + $parte[0] + ' WHERE ' + $parte[1] + ' IS NOT NULL GROUP BY ' + $parte[1]) @())) {
            $k = ([string]$g.v).ToUpperInvariant()
            $usoPorValor[$k] = [int]$usoPorValor[$k] + [int]$g.n
        }
    }
    $itens = @($linhas | ForEach-Object {
        $linha = $_
        $uso = [int]$usoPorValor[([string]$linha.valor).ToUpperInvariant()]
        $item = [ordered]@{ valor = [string]$linha.valor; ordem = $linha.ordem; ativo = [bool]$linha.ativo; uso = $uso }
        if ($tipo -eq 'Etapa') {
            $item['encerra'] = [bool]$linha.encerra; $item['resultado'] = [string]$linha.resultado
            $item['quadro'] = [string]$linha.quadro; $item['exige_familia'] = [bool]$linha.exige_familia
        }
        $item
    })
    return [ordered]@{ tipo = $tipo; itens = $itens }
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
        [void]$w.Add('o.etapa NOT IN (' + $script:SqlEtapasEncerradas + ')')
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
    $o = Add-Calculados $l[0]
    # resumo das visitas do atendimento: total e a lista (pendentes primeiro)
    $vis = Invoke-Consulta 'SELECT id, situacao, data_agendada, hora_agendada, tipo, resultado, versao FROM visitas WHERE id_oportunidade=? ORDER BY IIf(situacao=''PENDENTE'', 0, 1), data_agendada, id' @((P 'inteiro' ([int]$o.id)))
    $o['visitas_total'] = $vis.Count
    $o['visitas'] = $vis
    return $o
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
# etapas que encerram (tabela listas, Encerra = Sim): subconsulta, sem nome de etapa no codigo
$script:SqlEtapasEncerradas = "SELECT l.valor FROM listas AS l WHERE l.tipo='Etapa' AND l.encerra=True"
$script:SqlAtendimentoAberto = "(o.etapa IS NULL OR o.etapa NOT IN (" + $script:SqlEtapasEncerradas + "))"
$script:QuadroFunil = '2. Funil comercial'
# DEV-21: encerrado com resultado PERDA (Perdido, Descartado) cuja retomar_em chegou: "retomadas a avaliar"
$script:SqlEtapasPerda = "SELECT l.valor FROM listas AS l WHERE l.tipo='Etapa' AND l.encerra=True AND l.resultado='PERDA'"
$script:CamposRetomada = @($script:CamposPlanejamentoPadrao + @('retomar_em', 'motivo_desfecho', 'dt_desfecho'))
$script:LimitacaoPaginacao = ('Leitura dinâmica: cada página lê o banco de novo. A página seguinte continua depois do último item ' +
    'devolvido, então inclusão e exclusão entre páginas não deslocam os demais; um atendimento alterado entre páginas pode ' +
    'mudar de posição (aparecer de novo ou não aparecer).')

# DEV-22: chave de ordenação completa - vencidas, responsável, data, prioridade, empresa e o código (único, desempate).
# Texto sem acento e em maiúsculas, comparado em ordem ordinal: o cursor continua exatamente depois do último item.
function Get-ChavePlanejamento([bool]$venc, [string]$resp, [datetime]$dt, [int]$pr, [string]$empresa, [string]$codigo) {
    $partes = @($(if ($venc) { '0' } else { '1' }), (ConvertTo-ChaveLista $resp), $dt.ToString('yyyyMMddHHmmss'), [string]$pr,
                (ConvertTo-ChaveLista $empresa), $codigo)
    return ($partes -join [string][char]1)
}

# Impressão digital dos filtros: o cursor só vale para a mesma consulta.
function Get-FiltroPlanejamento($per, [string]$resp, [bool]$completo, [bool]$comVencidas) {
    $texto = ($per.Inicio.ToString('yyyyMMdd') + '|' + $per.Fim.ToString('yyyyMMdd') + '|' + $resp + '|' + $completo + '|' + $comVencidas)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { return (([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($texto)))) -replace '-', '').Substring(0, 16).ToLowerInvariant() }
    finally { $sha.Dispose() }
}

function ConvertTo-CursorPlanejamento([string]$chave, [string]$filtro, [string]$geracao) {
    $j = ConvertTo-Json -InputObject ([ordered]@{ k = $chave; f = $filtro; g = $geracao }) -Compress
    return [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($j))
}

function ConvertFrom-CursorPlanejamento([string]$cursor) {
    $o = $null
    try { $o = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($cursor)) | ConvertFrom-Json } catch { $o = $null }
    if ($null -eq $o -or $null -eq $o.k -or $null -eq $o.f) {
        throw 'cursor ilegível: use o proximo_cursor da chamada anterior, sem alterar. Nada foi consultado.'
    }
    return $o
}

function Get-GeracaoBase { return [string](Invoke-Valor "SELECT valor FROM config WHERE chave='geracao_base'" @()) }

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
    $cursorArg = ([string](Get-Arg $a 'cursor' '')).Trim()
    if ($cursorArg -ne '' -and $aPartirDe -gt 0) { throw 'Use cursor OU a_partir_de, não os dois. Nada foi consultado.' }
    $cursor = $null
    if ($cursorArg -ne '') { $cursor = ConvertFrom-CursorPlanejamento $cursorArg }
    $filtro = Get-FiltroPlanejamento $per $resp $completo $comVencidas
    if ($null -ne $cursor -and [string]$cursor.f -ne $filtro) {
        throw 'O cursor é de outra consulta (período, responsável ou opções diferentes): recomece sem cursor. Nada foi consultado.'
    }
    $geracao = Get-GeracaoBase
    if ($null -ne $cursor -and [string]$cursor.g -ne $geracao) {
        throw 'A base foi trocada (restaurada ou substituída) entre as páginas: recomece sem cursor. Nada foi consultado.'
    }
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

    # DEV-21: retomadas a avaliar - Perdido/Descartado com retomar_em até o fim do período (as que já chegaram e
    # ainda não foram avaliadas também), cada um uma vez, sem reabrir o encerrado
    $p3 = New-Object System.Collections.ArrayList
    $sql3 = 'SELECT o.*, c.codigo_cliente, c.empresa, c.cidade, c.uf, ct.codigo AS codigo_contato, ct.nome AS contato,' +
            ' ct.cargo AS contato_cargo, ct.telefone AS contato_telefone, ct.email AS contato_email' +
            ' FROM (oportunidades o LEFT JOIN clientes c ON o.id_cliente = c.id)' +
            ' LEFT JOIN contatos ct ON o.id_contato = ct.id' +
            ' WHERE o.etapa IN (' + $script:SqlEtapasPerda + ') AND o.retomar_em < ?'
    [void]$p3.Add((P 'data' $per.Fim.AddDays(1)))
    if ($resp -ne '') { $sql3 += ' AND o.responsavel=?'; [void]$p3.Add((P 'texto' $resp)) }
    $sql3 += ' ORDER BY o.retomar_em, o.codigo'
    $retomadas = New-Object System.Collections.ArrayList
    foreach ($l in (Invoke-Consulta $sql3 $p3.ToArray())) {
        [void](Add-Calculados $l)
        $r = [ordered]@{}
        $campos = if ($completo) { @($l.Keys) } else { $script:CamposRetomada }
        foreach ($c in $campos) { if ($l.Contains($c) -and -not (Test-ValorVazio $l[$c])) { $r[$c] = $l[$c] } }
        [void]$retomadas.Add($r)
    }

    # ordenação: vencidas, responsável, data da próxima ação, prioridade (Alta, Média, Baixa, vazio), empresa e
    # código - pela chave completa em ordem ordinal (DEV-22: a mesma que o cursor usa)
    $itens = New-Object System.Collections.ArrayList
    foreach ($l in $linhas) {
        [void](Add-Calculados $l)   # situação e quadro: mesma regra do front
        $dt = ConvertTo-DataOuNulo $l['dt_prox_acao']
        $venc = ($dt -lt $hoje)
        if ($venc) { $l['vencida'] = $true }
        $nomeResp = if (Test-ValorVazio $l['responsavel']) { '(sem responsável)' } else { [string]$l['responsavel'] }
        [void]$itens.Add([pscustomobject]@{
            K = (Get-ChavePlanejamento $venc $nomeResp $dt (Get-OrdemPrioridade $l['prioridade']) ([string]$l['empresa']) ([string]$l['codigo']))
            R = $nomeResp; D = $dt; Venc = $venc; L = $l })
    }
    # Array.Sort(chaves, itens) pelo PowerShell ordena só as chaves (o vetor de itens chega como cópia):
    # ordena as chaves (únicas: o código desempata) e busca o item de cada uma
    $porChave = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([StringComparer]::Ordinal)
    foreach ($it in $itens) { $porChave[$it.K] = $it }
    $chaves = [string[]]@($porChave.Keys)
    [Array]::Sort($chaves, [StringComparer]::Ordinal)
    $ordenados = @(foreach ($k in $chaves) { $porChave[$k] })

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
    $saida = [ordered]@{ periodo = $per.Periodo; quantidade = $total; consistencia = 'dinamica' }
    # início da página: depois da chave do cursor (DEV-22) ou pelo deslocamento antigo (a_partir_de)
    if ($null -ne $cursor) {
        $aPartirDe = $total
        for ($j = 0; $j -lt $total; $j++) { if ([string]::CompareOrdinal($chaves[$j], [string]$cursor.k) -gt 0) { $aPartirDe = $j; break } }
    }
    $avisos = New-Object System.Collections.ArrayList
    if ($resp -ne '' -and $total -eq 0) {
        if ($equipe['sem_proxima_acao'] -gt 0) {
            [void]$avisos.Add("Nenhum atendimento de '" + $resp + "' com próxima ação no período; há " + $equipe['sem_proxima_acao'] + ' em aberto sem próxima ação.')
        } else {
            [void]$avisos.Add("Nenhum atendimento em aberto de '" + $resp + "'. Confira a grafia do responsável em listar_opcoes.")
        }
    }
    if ($null -eq $cursor -and $aPartirDe -gt 0 -and $aPartirDe -ge $total) {
        [void]$avisos.Add('a_partir_de (' + $aPartirDe + ') está além do último registro (' + $total + '): nada a devolver.')
    }
    if ($null -eq $cursor -and $aPartirDe -gt 0) {
        [void]$avisos.Add('a_partir_de é deslocamento: se a base mudar entre as páginas, itens podem repetir ou faltar; prefira cursor = proximo_cursor.')
    }

    # corte por tamanho: nunca em silêncio e sem perder registro - pagina
    $molde = [ordered]@{}; foreach ($k in $saida.Keys) { $molde[$k] = $saida[$k] }
    $molde['a_partir_de'] = $aPartirDe; $molde['devolvidos'] = $total; $molde['truncado'] = $true; $molde['proximo'] = $total
    $molde['proximo_cursor'] = $(if ($total -gt 0) { ConvertTo-CursorPlanejamento $chaves[$total - 1] $filtro $geracao } else { '' })
    $molde['limitacao'] = $script:LimitacaoPaginacao
    $molde['aviso'] = (($avisos + @('Resultado maior que o limite da ferramenta: devolvidos 0000 de 0000 a partir de 0000. Repita com cursor=proximo_cursor até não vir mais proximo_cursor; o resumo de todos vem na primeira página.')) -join ' ')
    $comResumo = ($aPartirDe -eq 0 -and $null -eq $cursor)
    if ($comResumo) { $molde['resumo'] = $resumo; $molde['equipe'] = $equipe; $molde['retomadas_a_avaliar'] = $retomadas; $molde['quantidade_retomadas'] = $retomadas.Count }
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
        $saida['proximo_cursor'] = ConvertTo-CursorPlanejamento $chaves[$i - 1] $filtro $geracao
        [void]$avisos.Add('Resultado maior que o limite da ferramenta: devolvidos ' + $lista.Count + ' de ' + $total +
                          ' a partir de ' + $aPartirDe + '. Repita com cursor=proximo_cursor até não vir mais proximo_cursor; o resumo de todos vem na primeira página.')
    }
    if ($i -lt $total -or $null -ne $cursor -or $aPartirDe -gt 0) { $saida['limitacao'] = $script:LimitacaoPaginacao }
    if ($avisos.Count -gt 0) { $saida['aviso'] = ($avisos -join ' ') }
    if ($comResumo) {
        $saida['resumo'] = $resumo
        $saida['equipe'] = $equipe
        $saida['quantidade_retomadas'] = $retomadas.Count
        $saida['retomadas_a_avaliar'] = $retomadas
    }
    $saida['atendimentos'] = $lista

    $json = ConvertTo-Json -InputObject $saida -Depth 10 -Compress
    return [ordered]@{ __conteudo = [ordered]@{ type = 'text'; text = $json } }
}

# ================================================================= EXTRATO
# exportar_base: grava a foto das tabelas em CSV e devolve só o envelope (arquivos, linhas, colunas, hashes).
# É a entrada dos scripts de análise: o dado não passa pela conversa. Não grava no banco. CSV em UTF-8 sem
# BOM, separador ',', todo campo entre aspas, CRLF - o formato que os scripts do padrão leem sem conversão.
# DEV-23 (ciclo 012, seção 5.3): destino fixo SISTEMA\DADOS\execucao\extratos, uma subpasta por extrato, com
# manifesto (solicitante, estação, finalidade, instante, corte e SHA-256); descarte por descartar_extratos.
# DEV-22: cada tabela é lida numa consulta separada - o manifesto declara consistencia = nao_comprovada.
$script:TabelasExtrato = @('clientes', 'contatos', 'oportunidades', 'log_alteracoes', 'listas', 'metas', 'planejamentos', 'visitas', 'municipios')
$script:TabelasExtratoPadrao = @('clientes', 'contatos', 'oportunidades')
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
    planejamentos  = 'SELECT * FROM planejamentos ORDER BY id'
    visitas        = 'SELECT * FROM visitas ORDER BY id'
    municipios     = 'SELECT * FROM municipios ORDER BY cod_ibge'
}
$script:CalculadosExtrato = @('situacao', 'quadro', 'dias_parado', 'ciclo')
$script:NomeManifestoExtrato = 'crm-extrato.json'
$script:ReleiturasExtrato = 2

# Raiz dos extratos: ao lado do banco, em SISTEMA\DADOS\execucao\extratos (nunca letra de unidade nem Documentos).
function Get-RaizExtratos {
    return (Join-Path (Split-Path -Parent $script:CaminhoBanco) 'SISTEMA\DADOS\execucao\extratos')
}

# Subpasta nova e única por extrato: AAAAMMDD-HHMMSS-<6 hex>[-sufixo].
function New-PastaExtrato([string]$sufixo) {
    $raiz = Get-RaizExtratos
    if (-not (Test-Path -LiteralPath $raiz)) { New-Item -ItemType Directory -Path $raiz -Force | Out-Null }
    for ($t = 0; $t -lt 5; $t++) {
        $nome = (Get-Date).ToString('yyyyMMdd-HHmmss') + '-' + [guid]::NewGuid().ToString('N').Substring(0, 6)
        if ($sufixo) { $nome += '-' + ($sufixo -replace '[^A-Za-z0-9-]', '') }
        $pasta = Join-Path $raiz $nome
        if (-not (Test-Path -LiteralPath $pasta)) { New-Item -ItemType Directory -Path $pasta | Out-Null; return $pasta }
    }
    throw 'Não consegui criar a pasta do extrato. Nada foi exportado.'
}

function New-ManifestoExtrato([string]$conteudo, $a, [string]$finalidadePadrao) {
    $autor = ([string](Get-Arg $a 'autor' '')).Trim()
    $fin = ([string](Get-Arg $a 'finalidade' '')).Trim()
    return [ordered]@{
        tipo = 'extrato-crm'; versao_manifesto = 1; conteudo = $conteudo
        criado_em = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss'); fuso = (Get-Date).ToString('zzz')
        solicitante = $(if ($autor) { $autor } else { 'não informado' })
        estacao = [Environment]::MachineName; usuario_tecnico = [Environment]::UserName
        finalidade = $(if ($fin) { $fin } else { $finalidadePadrao })
        versao_servidor = $script:VersaoServidor; versao_esquema = [string](Get-VersaoEsquema)
        retencao = 'mês corrente e anterior (seção 5.3); descarte por descartar_extratos'
        em_processamento = $true
    }
}

# Grava o manifesto por arquivo temporário: quem lê nunca pega o JSON pela metade.
function Write-ManifestoExtrato([string]$pasta, $manifesto) {
    $arq = Join-Path $pasta $script:NomeManifestoExtrato
    $tmp = $arq + '.tmp'
    [System.IO.File]::WriteAllText($tmp, (ConvertTo-Json -InputObject $manifesto -Depth 6), (New-Object System.Text.UTF8Encoding($false)))
    if (Test-Path -LiteralPath $arq) { Remove-Item -LiteralPath $arq -Force }
    Move-Item -LiteralPath $tmp -Destination $arq
}

function Get-HashArquivo([string]$arquivo) {
    return (Get-FileHash -LiteralPath $arquivo -Algorithm SHA256).Hash.ToLowerInvariant()
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

function Get-MaiorIdLog { $v = Invoke-Valor 'SELECT MAX(id) FROM log_alteracoes' @(); if ($null -eq $v) { return 0 }; return [int]$v }

function Invoke-ExportarBase($a) {
    # validação completa antes de qualquer acesso ao banco ou ao disco
    if (Test-TemArg $a 'destino') {
        throw ('destino saiu do contrato (conector 1.9.0): o extrato vai sempre para SISTEMA\DADOS\execucao\extratos, ' +
               'uma subpasta por extrato, e a resposta traz o caminho. Nada foi exportado.')
    }
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
    if (([string](Get-Arg $a 'finalidade' '')).Length -gt 200) { throw 'finalidade passa de 200 caracteres. Nada foi exportado.' }

    $pasta = New-PastaExtrato ''
    $man = New-ManifestoExtrato 'base' $a 'análise'
    $man['tabelas'] = [ordered]@{}
    Write-ManifestoExtrato $pasta $man
    try {
        # Sem corte comprovado: refaz a leitura (até 2 vezes) se o log mudou durante ela. Mesmo log igual antes e
        # depois não prova foto num instante (commit fora de ordem), por isso a consistência fica nao_comprovada.
        $releituras = 0
        while ($true) {
            $antes = Get-MaiorIdLog
            $info = [ordered]@{}
            foreach ($t in $tabelas) {
                $linhas = Invoke-Consulta $script:SqlExtrato[$t] @()
                if ($t -eq 'oportunidades') { foreach ($l in $linhas) { [void](Add-Calculados $l) } }
                if ($linhas.Count -gt 0) { $colunas = @($linhas[0].Keys) }
                else {
                    $colunas = @(); foreach ($c in (Get-ColunasConsulta $script:SqlExtrato[$t])) { $colunas += [string]$c }
                    if ($t -eq 'oportunidades') { $colunas += $script:CalculadosExtrato }
                }
                $nome = 'crm-' + $t + '.csv'
                $arq = Join-Path $pasta $nome
                Write-ExtratoCsv $arq $linhas ([string[]]$colunas)
                $info[$t] = [ordered]@{ arquivo = $nome; linhas = $linhas.Count; colunas = $colunas.Count; sha256 = (Get-HashArquivo $arq) }
            }
            $depois = Get-MaiorIdLog
            if ($depois -eq $antes -or $releituras -ge $script:ReleiturasExtrato) { break }
            $releituras++
        }
        $man['tabelas'] = $info
        $man['consistencia'] = 'nao_comprovada'
        $man['corte'] = [ordered]@{
            geracao_base = (Get-GeracaoBase); max_log_antes = $antes; max_log_depois = $depois; releituras = $releituras
            log_mudou_durante_a_leitura = ($antes -ne $depois)
            nota = 'Cada tabela é lida numa consulta separada: o extrato não é uma foto de um único instante. Log igual antes e depois só reduz a chance de mistura, não prova o corte.'
        }
        $man['em_processamento'] = $false
        Write-ManifestoExtrato $pasta $man
    } catch {
        $man['em_processamento'] = $false; $man['falhou'] = [string]$_.Exception.Message
        try { Write-ManifestoExtrato $pasta $man } catch { }
        throw
    }
    $saida = [ordered]@{
        gerado_em = $man.criado_em
        pasta = $pasta
        pasta_relativa = ('SISTEMA\DADOS\execucao\extratos\' + (Split-Path -Leaf $pasta))
        manifesto = $script:NomeManifestoExtrato
        consistencia = 'nao_comprovada'
        corte = $man.corte
        versao_esquema = $man.versao_esquema
        versao_servidor = $script:VersaoServidor
        separador = ','
        codificacao = 'UTF-8'
        tabelas = $man.tabelas
    }
    if ($man.corte.log_mudou_durante_a_leitura) {
        $saida['aviso'] = 'Houve gravação durante a leitura mesmo depois de ' + $releituras + ' releitura(s): as tabelas podem não combinar entre si. Repita mais tarde se precisar de um extrato coerente.'
    }
    return $saida
}

# descartar_extratos (DEV-23): só subpastas diretas da raiz de extratos, com manifesto válido, fora de
# processamento e criadas antes do mês anterior. Link (ponto de reanálise) nunca é seguido.
function Invoke-DescartarExtratos($a) {
    $simular = [bool](Get-Arg $a 'simular' $true)
    $autor = ([string](Get-Arg $a 'autor' '')).Trim()
    $raiz = Get-RaizExtratos
    $limite = (Get-Date -Day 1).Date.AddMonths(-1)
    $res = [ordered]@{ simulado = $simular; raiz = $raiz; criadas_antes_de = $limite.ToString('yyyy-MM-dd')
                       descartar = @(); descartadas = @(); mantidas = 0; ignoradas = @() }
    if (-not (Test-Path -LiteralPath $raiz -PathType Container)) { return $res }
    $raizCanon = [System.IO.Path]::GetFullPath($raiz).TrimEnd('\')
    foreach ($d in @(Get-ChildItem -LiteralPath $raiz -Directory -Force)) {
        if ($d.Attributes -band [System.IO.FileAttributes]::ReparsePoint) { $res.ignoradas += ($d.Name + ': link, nunca seguido'); continue }
        if ([System.IO.Path]::GetFullPath($d.FullName).TrimEnd('\') -ne ($raizCanon + '\' + $d.Name)) { $res.ignoradas += ($d.Name + ': fora da raiz'); continue }
        $m = $null
        try { $m = Get-Content -LiteralPath (Join-Path $d.FullName $script:NomeManifestoExtrato) -Raw -Encoding UTF8 -ErrorAction Stop | ConvertFrom-Json } catch { $m = $null }
        if ($null -eq $m -or [string]$m.tipo -ne 'extrato-crm') { $res.ignoradas += ($d.Name + ': sem manifesto válido'); continue }
        $quando = [datetime]::MinValue
        if (-not [datetime]::TryParseExact([string]$m.criado_em, 'yyyy-MM-ddTHH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$quando)) {
            $res.ignoradas += ($d.Name + ': manifesto sem data'); continue
        }
        if ($m.em_processamento -eq $true) { $res.ignoradas += ($d.Name + ': em processamento'); continue }
        if ($quando -ge $limite) { $res.mantidas++; continue }
        $res.descartar += $d.Name
        if (-not $simular) {
            Remove-Item -LiteralPath $d.FullName -Recurse -Force
            $res.descartadas += $d.Name
        }
    }
    if (-not $simular -and $res.descartadas.Count -gt 0) {
        $linha = (Get-Date).ToString('yyyy-MM-ddTHH:mm:sszzz') + ' | ' + [Environment]::MachineName + ' | ' + [Environment]::UserName + ' | ' +
                 $(if ($autor) { $autor } else { 'não informado' }) + ' | ' + ($res.descartadas -join ', ')
        Add-Content -LiteralPath (Join-Path $raiz 'descartes.log') -Value $linha -Encoding UTF8
    }
    return $res
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
            $v = ConvertTo-CampoAtendimento $ctx $campo $a.$campo $atual[$campo]
            $novo[$campo] = $v
            if (-not (Test-Igual $atual[$campo] $v)) { $mud[$campo] = @{ Tipo = $script:DefsAtendimento[$campo][0]; Novo = $v; Antigo = $atual[$campo] } }
        }
        # troca do contato: só dentro da mesma empresa (o código AT é congelado e nomeia a pasta do cliente)
        if ((Test-TemArg $a 'codigo_contato') -or (Test-TemArg $a 'id_contato')) {
            $ct = Get-ContatoNaTransacao $ctx $a
            # DEV-11: conferido dentro da transação (C29)
            $mCt = Test-ContatoDoAtendimento $true ([bool]$ct.ativo) ([int]$ct.id_cliente) ([int]$atual.id_cliente)
            if ($mCt) { throw ($ct.codigo + ': ' + $mCt) }
            if ([int]$ct.id -ne [int]$atual.id_contato) { $mud['id_contato'] = @{ Tipo = 'inteiro'; Novo = [int]$ct.id; Antigo = $atual.id_contato } }
        }
        if (Test-TemArg $a 'observacoes_acrescentar') {
            $obs = Add-Observacao $atual.observacoes ([string]$a.observacoes_acrescentar)
            $mud['observacoes'] = @{ Tipo = 'memo'; Novo = $obs; Antigo = $atual.observacoes }
        }
        # DEV-16: troca do atendimento anterior (null remove): mesma empresa, não o próprio, sem ciclo (C33)
        if (Test-TemArg $a 'codigo_atendimento_anterior') {
            $idAntNovo = $null
            if ($null -ne $a.codigo_atendimento_anterior) {
                $ant = Get-AtendimentoNaTransacao $ctx ([string]$a.codigo_atendimento_anterior)
                if ([int]$ant.id -eq [int]$atual.id) { throw 'O atendimento anterior não pode ser o próprio atendimento.' }
                if ([int]$ant.id_cliente -ne [int]$atual.id_cliente) { throw ('O atendimento ' + $ant.codigo + ' é de outra empresa. O atendimento anterior tem de ser da mesma empresa.') }
                $mapa = @{}
                foreach ($o in $ctx.Consultar('SELECT id, id_atendimento_anterior FROM oportunidades WHERE id_cliente=?', @((P 'inteiro' ([int]$atual.id_cliente))))) { $mapa[[int]$o.id] = [int]$(if ($null -eq $o.id_atendimento_anterior) { 0 } else { $o.id_atendimento_anterior }) }
                if (Test-AnteriorFormaCiclo ([int]$atual.id) ([int]$ant.id) $mapa) { throw ('O atendimento ' + $ant.codigo + ' já vem deste atendimento: a corrente de anteriores fecharia um ciclo (C33).') }
                $idAntNovo = [int]$ant.id
            }
            if (-not (Test-Igual $atual.id_atendimento_anterior $idAntNovo)) { $mud['id_atendimento_anterior'] = @{ Tipo = 'inteiro'; Novo = $idAntNovo; Antigo = $atual.id_atendimento_anterior } }
        }
        # DEV-10: transições de etapa e retomada (seção 5.2), as mesmas do front (modRegras)
        $papelNovo = Get-PapelEtapa ([string]$novo.etapa)
        $perdaNova = ($papelNovo.encerra -and $papelNovo.resultado -eq 'PERDA')
        $mRet = Test-RetomadaPermitida $papelNovo.encerra $perdaNova $novo.retomar_em $atual.retomar_em
        if ($mRet) { throw $mRet }
        $tr = Get-TransicaoEtapa (Get-PapelEtapa ([string]$atual.etapa)).encerra $papelNovo.encerra $perdaNova
        foreach ($c in $tr.limpar) {
            $novo[$c] = $null
            if (Test-Igual $atual[$c] $null) { $mud.Remove($c) } else { $mud[$c] = @{ Tipo = $script:DefsAtendimento[$c][0]; Novo = $null; Antigo = $atual[$c] } }
        }
        # C34 (seção 5.2): próxima ação nova ou alterada só hoje ou futura; a vencida já gravada fica
        $mPx = Test-ProxAcaoPermitida $novo.dt_prox_acao $atual.dt_prox_acao
        if ($mPx) { throw $mPx }
        $avisos = Get-CriticaAtendimento $novo
        if ($tr.orientacao) { [void]$avisos.Add($tr.orientacao) }
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
                if ($null -ne $v) { $v = ConvertTo-Inteiro $campo $v; if ($v -lt 1 -or $v -gt 3) { throw "Campo '$campo' vai de 1 a 3." } }
            } elseif ($lista) {
                $v = Resolve-ValorLista $ctx $lista $campo $v $atual[$campo]
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
        Add-EmailsNasObservacoes $a $atual $mud
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
function Get-MensagemConflitoContexto($atual, [int]$lida) {
    $agora = [ordered]@{ ctx_versao = $atual.ctx_versao; ctx_resumo = $atual.ctx_resumo; ctx_arquivo = $atual.ctx_arquivo; ctx_atualizado_por = $atual.ctx_atualizado_por; ctx_atualizado_em = (ConvertFrom-ValorBanco $atual.ctx_atualizado_em) }
    if ($atual.Contains('ctx_familia')) { $agora['ctx_familia'] = $atual.ctx_familia }
    return ('Conflito de contexto: você leu a ctx_versao ' + $lida + ', mas o contexto já está na ' + $atual.ctx_versao +
            ' (outra sessão gravou antes). Nada foi gravado. Contexto atual: ' + (ConvertTo-Json -InputObject $agora -Compress) + '. Releia, junte o que for preciso e repita com a ctx_versao atual.')
}

function Set-Contexto($ctx, [string]$tabela, $atual, $a, [string[]]$campos) {
    $mud = [ordered]@{}
    foreach ($campo in $campos) {
        if (-not (Test-TemArg $a $campo)) { continue }
        # seção 5.1 (ciclo 009): resumo e família em maiúsculas; ctx_arquivo é caminho, como gravado
        $v = ConvertTo-Normalizado $campo $a.$campo
        Assert-Tamanho $campo $v
        if (-not (Test-Igual $atual[$campo] $v)) { $mud[$campo] = @{ Novo = $v; Antigo = $atual[$campo] } }
    }
    # DEV-17 (esquema 1.2): versão PRÓPRIA do contexto, independente da comercial. Quem leu uma versão e outra
    # sessão gravou antes recebe conflito com o conteúdo atual (nada se sobrescreve sem aviso).
    $lida = [int](Get-Arg $a 'ctx_versao')
    $atualCtx = [int]$(if ($null -eq $atual.ctx_versao) { 0 } else { $atual.ctx_versao })
    if ($lida -ne $atualCtx) { throw (Get-MensagemConflitoContexto $atual $lida) }
    if ($mud.Count -eq 0) { return @() }
    $sets = @(); $params = @()
    foreach ($c in $mud.Keys) { $sets += ($c + '=?'); $params += (P 'texto' $mud[$c].Novo) }
    $sql = 'UPDATE ' + $tabela + ' SET ' + ($sets -join ', ') + ', ctx_versao=ctx_versao+1, ctx_atualizado_em=?, ctx_atualizado_por=? WHERE id=? AND ctx_versao=?'
    $params += (P 'data' (Get-Date)); $params += (P 'texto' $script:UsuarioLog); $params += (P 'inteiro' ([int]$atual.id)); $params += (P 'inteiro' $lida)
    $n = $ctx.Executar($sql, $params)
    if ($n -ne 1) { throw (Get-MensagemConflitoContexto $atual $lida) }
    foreach ($c in $mud.Keys) { $ctx.Log($tabela, [int]$atual.id, $c, $mud[$c].Antigo, $mud[$c].Novo, 'CONTEXTO') }
    return @($mud.Keys)
}

function Invoke-AtualizarContextoCliente($a) {
    Assert-EsquemaGravavel
    return Invoke-Transacao {
        param($ctx)
        $atual = Get-ClienteNaTransacao $ctx $a
        $alt = Set-Contexto $ctx 'clientes' $atual $a @('ctx_resumo', 'ctx_familia', 'ctx_arquivo')
        $v = [int]$(if ($null -eq $atual.ctx_versao) { 0 } else { $atual.ctx_versao })
        return [ordered]@{ id_cliente = $atual.id; empresa = $atual.empresa; gravado = ($alt.Count -gt 0); campos = @($alt); ctx_versao = $(if ($alt.Count -gt 0) { $v + 1 } else { $v }) }
    }
}

function Invoke-AtualizarContextoAtendimento($a) {
    Assert-EsquemaGravavel
    $codigo = [string]$a.codigo
    return Invoke-Transacao {
        param($ctx)
        $atual = Get-AtendimentoNaTransacao $ctx $codigo
        $alt = Set-Contexto $ctx 'oportunidades' $atual $a @('ctx_resumo', 'ctx_arquivo')
        $v = [int]$(if ($null -eq $atual.ctx_versao) { 0 } else { $atual.ctx_versao })
        return [ordered]@{ codigo = $codigo; gravado = ($alt.Count -gt 0); campos = @($alt); ctx_versao = $(if ($alt.Count -gt 0) { $v + 1 } else { $v }) }
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
    $extrasEmail = $(if (Test-TemArg $a 'email') { $script:EmailsAdicionais } else { '' })
    $obs = $null
    if ((Test-TemArg $a 'observacoes') -and [string]$a.observacoes -ne '') { $obs = Add-Observacao $null ([string]$a.observacoes) }
    $obs = Add-EmailsAdicionais $obs $extrasEmail
    $promover = [bool](Get-Arg $a 'promover_pre_cliente' $false)

    for ($tentativa = 1; $tentativa -le 5; $tentativa++) {
        try {
            return Invoke-Transacao {
                param($ctx)
                $repetida = Get-OperacaoAnterior $ctx $a 'contatos'
                if ($repetida) { return $repetida }
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
                # DEV-10: contato com o mesmo nome na mesma empresa é aviso, não bloqueio
                $homonimos = @($ctx.Consultar('SELECT codigo FROM contatos WHERE id_cliente=? AND nome=?', @((P 'inteiro' ([int]$cli.id)), (P 'texto' $d['nome']))) | ForEach-Object { [string]$_.codigo })
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
                Save-Operacao $ctx $a 'contatos' $id $codigo 1
                return [ordered]@{
                    gravado = $true; id_contato = $id; codigo_contato = $codigo; versao = 1
                    empresa = $cli.empresa; codigo_cliente = Format-CodigoCliente $codCli; empresa_promovida = $promovida
                    nome = $d['nome']; cargo = $d['cargo']; telefone = $d['telefone']; email = $d['email']
                    avisos = @($(if ($homonimos.Count -gt 0) { 'Já existe contato com o mesmo nome nesta empresa (' + ($homonimos -join ', ') + '). Confira com o usuário se não é o mesmo.' }))
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
        Add-EmailsNasObservacoes $a $atual $mud
        $avisos = New-Object System.Collections.ArrayList
        $acao = 'ALTERAR'
        if (Test-TemArg $a 'ativo') {
            $v = [bool]$a.ativo
            if ([bool]$atual.ativo -ne $v) {
                $mud['ativo'] = @{ Tipo = 'logico'; Novo = $v; Antigo = [bool]$atual.ativo }
                if ($mud.Count -eq 1) { $acao = if ($v) { 'REATIVACAO' } else { 'DESATIVACAO' } }
                if (-not $v) {
                    $abertos = [int]$ctx.Valor(('SELECT COUNT(*) FROM oportunidades WHERE id_contato=? AND etapa NOT IN (' +
                                               $script:SqlEtapasEncerradas + ')'),
                                               (@((P 'inteiro' ([int]$atual.id)))))
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
    if ((Get-PapelEtapa $etapa).encerra) {
        if ([string]$novo.motivo_desfecho -eq '') { throw ('Etapa "' + $etapa + '" exige motivo_desfecho.') }
        if (-not $dDes) { throw ('Etapa "' + $etapa + '" exige dt_desfecho.') }
    }
    if ((Get-PapelEtapa $etapa).exige_familia -and [string]$novo.familia -eq '') {
        throw ('Etapa "' + $etapa + '" exige a família de máquina.')
    }
    $avisos = New-Object System.Collections.ArrayList
    # C28 (seção 5.2): datas fora de ordem são aviso, não recusa
    if ($dEnt -and $dPro -and $dPro -lt $dEnt) { [void]$avisos.Add('A data da proposta é anterior à data de entrada.') }
    if ($dEnt -and $dDes -and $dDes -lt $dEnt) { [void]$avisos.Add('A data do desfecho é anterior à data de entrada.') }
    if ((Get-PapelEtapa $etapa).resultado -eq 'GANHO' -and $null -eq $novo.valor) { [void]$avisos.Add('Pedido fechado sem valor informado.') }
    if ((-not (Get-PapelEtapa $etapa).encerra) -and -not $novo.dt_prox_acao -and -not $novo.retomar_em) {
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
function ConvertTo-CampoAtendimento($ctx, [string]$campo, $v, $atual = $null) {
    $tipo = $script:DefsAtendimento[$campo][0]; $lista = $script:DefsAtendimento[$campo][1]
    switch ($tipo) {
        'data'    { return (ConvertTo-Data $campo $v) }
        'moeda'   { $v = ConvertTo-Moeda $campo $v; if ($null -ne $v) { if ($v -lt 0) { throw 'O valor não pode ser negativo.' } }; return $v }
        'inteiro' { $v = ConvertTo-Inteiro $campo $v; if ($null -ne $v) { if ($v -lt 0) { throw "Campo '$campo' não pode ser negativo." } }; return $v }
        default   {
            if ($lista) { return (Resolve-ValorLista $ctx $lista $campo $v $atual) }
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
function ConvertTo-CampoCliente($ctx, [string]$campo, $v, [int]$idProprio, $atual = $null) {
    $defs = @{ uf = 'UF'; segmento = 'Segmento'; responsavel = 'Responsavel' }
    if ($campo -in @('aderencia', 'porte')) {
        if ($null -ne $v) { $v = ConvertTo-Inteiro $campo $v; if ($v -lt 1 -or $v -gt 3) { throw "Campo '$campo' vai de 1 a 3." } }
        return $v
    }
    if ($defs.ContainsKey($campo)) { return (Resolve-ValorLista $ctx $defs[$campo] $campo $v $atual) }
    if ($campo -eq 'qualificacao' -and @(Get-ValoresLista $ctx 'Qualificacao').Count -gt 0) { return (Resolve-ValorLista $ctx 'Qualificacao' $campo $v $atual) }
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
        $repetida = Get-OperacaoAnterior $ctx $a 'clientes'
        if ($repetida) { return $repetida }
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
        if (Test-TemArg $a 'email') { [void](ConvertTo-Normalizado 'email' $a.email); $obs = Add-EmailsAdicionais $obs $script:EmailsAdicionais }
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
        Save-Operacao $ctx $a 'clientes' $id (Format-CodigoCliente $cod) 1
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
        switch ($acao) {
            'desativar' {
                if (-not [bool]$cli.ativo) { throw 'A empresa já está inativa.' }
                $nOp = [int]$ctx.Valor(('SELECT COUNT(*) FROM oportunidades WHERE id_cliente=? AND etapa NOT IN (' + $script:SqlEtapasEncerradas + ')'), (@((P 'inteiro' $id))))
                [void](Set-AtivoEm $ctx 'clientes' 'id=?' $id $false 'DESATIVACAO' ([string]$nOp + ' oportunidade(s) em aberto mantida(s)'))
                $nCt = Set-AtivoEm $ctx 'contatos' 'id_cliente=?' $id $false 'DESATIVACAO' 'cascata da desativação do cliente'
                $aviso = $null
                if ($nOp -gt 0) { $aviso = [string]$nOp + ' atendimento(s) em aberto continuam no funil; encerre um a um se não forem mais trabalhados.' }
                return [ordered]@{ gravado = $true; empresa = $cli.empresa; ativo = $false; contatos_desativados = $nCt; aviso = $aviso }
            }
            'reativar' {
                if ([bool]$cli.ativo) { throw 'A empresa já está ativa.' }
                # DEV-15 (seção 5.2): reativar a empresa reativa SÓ a empresa; contato volta um a um (atualizar_contato ativo=true)
                [void](Set-AtivoEm $ctx 'clientes' 'id=?' $id $true 'REATIVACAO' '')
                $inativos = [int]$ctx.Valor('SELECT COUNT(*) FROM contatos WHERE id_cliente=? AND ativo=False', @((P 'inteiro' $id)))
                return [ordered]@{ gravado = $true; empresa = $cli.empresa; ativo = $true; contatos_reativados = 0; contatos_inativos = $inativos
                                   aviso = $(if ($inativos -gt 0) { [string]$inativos + ' contato(s) continuam inativos: reative um a um (atualizar_contato ativo=true) só quem ainda trabalha na empresa.' } else { $null }) }
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
        $repetida = Get-OperacaoAnterior $ctx $a 'oportunidades'
        if ($repetida) { return $repetida }
        $ct = Get-ContatoNaTransacao $ctx $a
        if (-not [bool]$ct.ativo) { throw ('O contato ' + $ct.codigo + ' está inativo. Reative-o antes de abrir atendimento.') }
        $cli = $ctx.Consultar('SELECT * FROM clientes WHERE id=?', @((P 'inteiro' ([int]$ct.id_cliente))))[0]
        # DEV-11 (C29): conferido na transação; a empresa do atendimento é a do contato
        $mCt = Test-ContatoDoAtendimento $true ([bool]$ct.ativo) ([int]$ct.id_cliente) ([int]$cli.id)
        if ($mCt) { throw $mCt }
        if (-not [bool]$cli.ativo) { throw ('A empresa ' + $cli.empresa + ' está inativa.') }
        $novo = [ordered]@{}
        foreach ($c in $script:DefsAtendimento.Keys) { $novo[$c] = $null }
        $novo['dt_entrada'] = (Get-Date).Date
        foreach ($c in $script:DefsAtendimento.Keys) { if (Test-TemArg $a $c) { $novo[$c] = ConvertTo-CampoAtendimento $ctx $c $a.$c } }
        # DEV-10: retomada só em Perdido/Descartado; nascer encerrado em perda não tem próxima ação
        $papelNovo = Get-PapelEtapa ([string]$novo.etapa)
        $perdaNova = ($papelNovo.encerra -and $papelNovo.resultado -eq 'PERDA')
        $mRet = Test-RetomadaPermitida $papelNovo.encerra $perdaNova $novo.retomar_em $null
        if ($mRet) { throw $mRet }
        foreach ($c in (Get-TransicaoEtapa $false $papelNovo.encerra $perdaNova).limpar) { $novo[$c] = $null }
        # C34 (seção 5.2): a próxima ação do atendimento novo é hoje ou futura
        $mPx = Test-ProxAcaoPermitida $novo.dt_prox_acao $null
        if ($mPx) { throw $mPx }
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
        Save-Operacao $ctx $a 'oportunidades' $id $codigo 1
        return [ordered]@{ gravado = $true; codigo = $codigo; id = $id; versao = 1; empresa = $cli.empresa; contato = $ct.nome
                           empresa_promovida = $promovida; situacao = (Get-Situacao $true ([string]$novo.etapa) $novo.dt_prox_acao $novo.retomar_em)
                           avisos = @($avisos) }
    }
}

# ================================================================= DUPLICAR (DEV-28, ciclo 017)
# A regra da cópia num lugar só: o que passa do atendimento de origem para o novo. Igual a
# modCRM.CAMPOS_DUPLICAR no front (o verificar.py confere). O resto nasce vazio (valor, orçamento,
# proposta, agenda, desfecho, retomada, tentativas); dt_entrada = hoje; etapa = a inicial; o anterior
# é a origem. Duplicar encerrado é permitido.
$script:CamposDuplicar = 'responsavel,origem,prioridade,categoria,tipo_venda,maquina,familia,ultima_interacao'

# Etapa inicial: a primeira etapa ativa que não encerra, pela ordem da lista (nenhum nome no código).
function Get-EtapaInicial {
    $l = Invoke-Consulta "SELECT valor FROM listas WHERE tipo='Etapa' AND ativo=True AND encerra=False ORDER BY ordem, id" @()
    if ($l.Count -eq 0) { throw 'Não há etapa ativa que não encerre o atendimento: confira a lista Etapa. Nada foi gravado.' }
    return [string]$l[0].valor
}

# Argumentos do abrir_atendimento que duplicam o atendimento de origem (sem gravar nada).
function Get-DadosParaDuplicar([string]$codigo) {
    $l = Invoke-Consulta ('SELECT o.*, ct.codigo AS codigo_contato, ct.ativo AS contato_ativo FROM oportunidades o ' +
                          'LEFT JOIN contatos ct ON o.id_contato = ct.id WHERE o.codigo=?') @((P 'texto' $codigo))
    if ($l.Count -eq 0) { throw ('Atendimento ' + $codigo + ' não encontrado.') }
    $o = $l[0]
    $d = [ordered]@{ codigo_contato = [string]$o.codigo_contato }
    foreach ($c in $script:CamposDuplicar.Split(',')) { if (-not (Test-ValorVazio $o[$c])) { $d[$c] = $o[$c] } }
    $d['etapa'] = Get-EtapaInicial
    $d['codigo_atendimento_anterior'] = [string]$o.codigo
    return [pscustomobject]@{ dados = $d; contato_ativo = [bool]$o.contato_ativo }
}

function Invoke-DuplicarAtendimento($a) {
    $codigo = [string]$a.codigo
    $simular = $true; if (Test-TemArg $a 'simular') { $simular = [bool]$a.simular }
    $r = Get-DadosParaDuplicar $codigo
    $zerados = @('valor', 'orcamento', 'dt_proposta', 'prox_acao', 'dt_prox_acao', 'motivo_desfecho', 'dt_desfecho', 'retomar_em', 'tentativas')
    if ($simular) {
        $s = [ordered]@{ simulado = $true; duplicado_de = $codigo; dados = $r.dados; zerados = $zerados; dt_entrada = (Get-Date).ToString('yyyy-MM-dd') }
        if (-not $r.contato_ativo) { $s['aviso'] = 'O contato do atendimento de origem está inativo: reative-o ou abra o atendimento com outro contato da empresa.' }
        return $s
    }
    if (-not $r.contato_ativo) { throw ('O contato ' + $r.dados.codigo_contato + ' está inativo. Reative-o antes de duplicar (ou use abrir_atendimento com outro contato). Nada foi gravado.') }
    $args2 = [ordered]@{}
    foreach ($k in $r.dados.Keys) { $args2[$k] = $r.dados[$k] }
    foreach ($k in @('id_operacao', 'autor', 'promover_pre_cliente')) { if (Test-TemArg $a $k) { $args2[$k] = $a.$k } }
    $g = Invoke-AbrirAtendimento ([pscustomobject]$args2)
    $g['duplicado_de'] = $codigo
    return $g
}

function Invoke-GerenciarLista($a) {
    Assert-EsquemaGravavel
    $tipo = ([string]$a.tipo).Trim(); $valor = ([string]$a.valor).Trim(); $acao = [string]$a.acao
    if ($tipo -eq '' -or $valor -eq '') { throw 'Informe tipo e valor.' }
    if ($valor.Length -gt 120) { throw 'Valor aceita no máximo 120 caracteres.' }
    if ($acao -notin @('incluir', 'ativar', 'desativar', 'ordenar', 'definir_papel')) { throw 'acao deve ser incluir, ativar, desativar, ordenar ou definir_papel.' }
    return Invoke-Transacao {
        param($ctx)
        $tipos = @($ctx.Consultar('SELECT DISTINCT tipo FROM listas', @()) | ForEach-Object { [string]$_.tipo })
        $tipoOk = $tipos | Where-Object { $_ -ieq $tipo } | Select-Object -First 1
        if (-not $tipoOk) { throw ("Lista '" + $tipo + "' não existe. Listas: " + ($tipos -join ' | ')) }
        $ehEtapa = ($tipoOk -eq 'Etapa')
        # pela chave (C01/C02): "Pedído Fechado" é o mesmo valor que "Pedido Fechado"
        $chave = ConvertTo-ChaveLista $valor
        $ex = @($ctx.Consultar('SELECT * FROM listas WHERE tipo=?', @((P 'texto' $tipoOk))) | Where-Object { (ConvertTo-ChaveLista ([string]$_.valor)) -eq $chave })
        if ($ex.Count -gt 0 -and $acao -ne 'incluir') { $valor = [string]$ex[0].valor }
        $ordem = $null; if (Test-TemArg $a 'ordem') { $ordem = [int]$a.ordem }
        # C05: UF só entra se for uma das 27
        if ($tipoOk -eq 'UF' -and $acao -eq 'incluir' -and -not (Test-UFValida $valor)) { throw ("'" + $valor + "' não é uma UF do Brasil (C05).") }
        switch ($acao) {
            'incluir' {
                if ($ex.Count -gt 0) {
                    $sufixo = ''; if (-not [bool]$ex[0].ativo) { $sufixo = ' (inativo: use acao=ativar)' }
                    throw ("'" + $valor + "' já existe na lista " + $tipoOk + $sufixo + '.')
                }
                if ($null -eq $ordem) { $ordem = [int]$ctx.Valor('SELECT MAX(ordem) FROM listas WHERE tipo=?', @((P 'texto' $tipoOk))) + 1 }
                if ($ehEtapa) {
                    # etapa nova nasce aberta, no funil, sem família, a menos que o pedido diga o contrário (C04 vale)
                    $enc = (Test-TemArg $a 'encerra') -and [bool]$a.encerra
                    $res = ''; if (Test-TemArg $a 'resultado') { $res = ([string]$a.resultado).Trim().ToUpperInvariant() }
                    $erro = Test-PapelEtapaValido $enc $res; if ($erro) { throw $erro }
                    $quad = 'FUNIL'; if (Test-TemArg $a 'quadro') { $quad = ([string]$a.quadro).Trim().ToUpperInvariant() }
                    if ($quad -notin @('FUNIL', 'PROSPECCAO')) { throw 'quadro deve ser FUNIL ou PROSPECCAO.' }
                    $fam = (Test-TemArg $a 'exige_familia') -and [bool]$a.exige_familia
                    $resP = $null; if ($res -ne '') { $resP = $res }
                    [void]$ctx.Executar('INSERT INTO listas (tipo, valor, chave, ordem, ativo, encerra, resultado, quadro, exige_familia) VALUES (?,?,?,?,True,?,?,?,?)', @((P 'texto' $tipoOk), (P 'texto' $valor), (P 'texto' $chave), (P 'inteiro' $ordem), (P 'logico' $enc), (P 'texto' $resP), (P 'texto' $quad), (P 'logico' $fam)))
                } else {
                    [void]$ctx.Executar('INSERT INTO listas (tipo, valor, chave, ordem, ativo) VALUES (?,?,?,?,True)', @((P 'texto' $tipoOk), (P 'texto' $valor), (P 'texto' $chave), (P 'inteiro' $ordem)))
                }
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
                    # C03: não desativa a última etapa ativa de ganho ou de perda
                    if (-not $novo -and $ehEtapa -and -not (Test-DesativarEtapaPermitido (Get-PapeisEtapa $ctx) $valor)) {
                        throw ("'" + $valor + "' é a última etapa ativa de " + [string]$r.resultado + ': sem ela o funil não fecha. Ative outra antes (C03).')
                    }
                    [void]$ctx.Executar('UPDATE listas SET ativo=? WHERE id=?', @((P 'logico' $novo), (P 'inteiro' $id)))
                    $ac = 'DESATIVACAO'; if ($novo) { $ac = 'REATIVACAO' }
                    $ctx.Log('listas', $id, 'ativo', [string](-not $novo), [string]$novo, $ac)
                }
            }
            'definir_papel' {
                if (-not $ehEtapa) { throw 'definir_papel só existe para a lista Etapa: as outras listas não têm papel.' }
                if ($ex.Count -eq 0) { throw ("'" + $valor + "' não existe na lista Etapa.") }
                $r = $ex[0]; $id = [int]$r.id
                $novoEnc = [bool]$r.encerra; if (Test-TemArg $a 'encerra') { $novoEnc = [bool]$a.encerra }
                $novoRes = [string]$r.resultado; if (Test-TemArg $a 'resultado') { $novoRes = ([string]$a.resultado).Trim().ToUpperInvariant() }
                $novoQuad = [string]$r.quadro; if (Test-TemArg $a 'quadro') { $novoQuad = ([string]$a.quadro).Trim().ToUpperInvariant() }
                $novoFam = [bool]$r.exige_familia; if (Test-TemArg $a 'exige_familia') { $novoFam = [bool]$a.exige_familia }
                $erro = Test-PapelEtapaValido $novoEnc $novoRes; if ($erro) { throw $erro }
                if ($novoQuad -notin @('FUNIL', 'PROSPECCAO')) { throw 'quadro deve ser FUNIL ou PROSPECCAO.' }
                $antRes = [string]$r.resultado
                $mud = [ordered]@{}
                if ($novoEnc -ne [bool]$r.encerra) { $mud['encerra'] = @{ Tipo = 'logico'; Antigo = [bool]$r.encerra; Novo = $novoEnc } }
                if ($novoRes -ne $antRes) { $mud['resultado'] = @{ Tipo = 'texto'; Antigo = $antRes; Novo = $novoRes } }
                if ($novoQuad -ne [string]$r.quadro) { $mud['quadro'] = @{ Tipo = 'texto'; Antigo = [string]$r.quadro; Novo = $novoQuad } }
                if ($novoFam -ne [bool]$r.exige_familia) { $mud['exige_familia'] = @{ Tipo = 'logico'; Antigo = [bool]$r.exige_familia; Novo = $novoFam } }
                if ($mud.Count -eq 0) { return [ordered]@{ gravado = $false; mensagem = 'Nada a alterar.' } }
                # C03 também vale ao tirar o resultado de uma etapa ativa
                if ($antRes -ne '' -and $novoRes -ne $antRes -and [bool]$r.ativo -and -not (Test-DesativarEtapaPermitido (Get-PapeisEtapa $ctx) $valor)) {
                    throw ("'" + $valor + "' é a última etapa ativa de " + $antRes + ': mude a outra antes (C03).')
                }
                # mudar 'encerra' muda a situação de todo atendimento nessa etapa: pede confirmação com a contagem
                $afetados = 0
                if ($mud.Contains('encerra')) {
                    $afetados = [int]$ctx.Valor('SELECT COUNT(*) FROM oportunidades WHERE etapa=?', @((P 'texto' $valor)))
                    if ($afetados -gt 0 -and -not ((Test-TemArg $a 'confirmar') -and $a.confirmar -eq $true)) {
                        throw ("Mudar 'encerra' de '" + $valor + "' muda a situação de " + $afetados + ' atendimento(s). Repita com confirmar=true para gravar.')
                    }
                }
                # controle de versão: a linha tem de ter os valores antigos (a tabela não tem coluna versao)
                $filtro = 'id=? AND encerra=? AND quadro=? AND exige_familia=?'
                $pars = @((P 'inteiro' $id), (P 'logico' ([bool]$r.encerra)), (P 'texto' ([string]$r.quadro)), (P 'logico' ([bool]$r.exige_familia)))
                if ($antRes -eq '') { $filtro += ' AND (resultado IS NULL OR resultado=?)'; $pars += (P 'texto' '') } else { $filtro += ' AND resultado=?'; $pars += (P 'texto' $antRes) }
                $resGravar = $null; if ($novoRes -ne '') { $resGravar = $novoRes }
                $valoresNovos = @((P 'logico' $novoEnc), (P 'texto' $resGravar), (P 'texto' $novoQuad), (P 'logico' $novoFam))
                $n = $ctx.Executar(('UPDATE listas SET encerra=?, resultado=?, quadro=?, exige_familia=? WHERE ' + $filtro), ($valoresNovos + $pars))
                if ($n -ne 1) { throw "Conflito: a etapa '$valor' foi alterada por outra pessoa agora. Nada foi gravado; leia de novo com listar_opcoes." }
                foreach ($c in $mud.Keys) { $ctx.Log('listas', $id, $c, [string]$mud[$c].Antigo, [string]$mud[$c].Novo, 'ALTERAR') }
                return [ordered]@{ gravado = $true; tipo = $tipoOk; valor = $valor; acao = $acao; alterados = @($mud.Keys)
                    papel = [ordered]@{ encerra = $novoEnc; resultado = $novoRes; quadro = $novoQuad; exige_familia = $novoFam }
                    atendimentos_afetados = $afetados; valores_ativos = @(Get-ValoresLista $ctx $tipoOk) }
            }
        }
        return [ordered]@{ gravado = $true; tipo = $tipoOk; valor = $valor; acao = $acao; valores_ativos = @(Get-ValoresLista $ctx $tipoOk) }
    }
}

# Papel de todas as etapas (valor, resultado, ativo) para a regra C03.
function Get-PapeisEtapa($ctx) {
    $linhas = $ctx.Consultar("SELECT valor, ativo, resultado FROM listas WHERE tipo='Etapa'", @())
    return @($linhas | ForEach-Object { [pscustomobject]@{ valor = [string]$_.valor; ativo = [bool]$_.ativo; resultado = [string]$_.resultado } })
}

# ---------------------------------------------------------------- visitas (Fase 3, regras V1-V9 da spec 003)

# AAAA-MM-DD, ou AAAA-MM-DD HH:mm quando a hora importa (visita realizada, histórico C12)
function Get-DataIsoOuErro([string]$texto, [string]$campo) {
    $cultura = [Globalization.CultureInfo]::InvariantCulture
    foreach ($formato in @('yyyy-MM-dd', 'yyyy-MM-dd HH:mm')) {
        try { return [datetime]::ParseExact($texto.Trim(), $formato, $cultura) } catch { }
    }
    throw ($campo + " em AAAA-MM-DD (ou AAAA-MM-DD HH:mm): '" + $texto + "'.")
}

function Get-VisitaNaTransacao($ctx, [int]$id) {
    $l = $ctx.Consultar('SELECT * FROM visitas WHERE id=?', @((P 'inteiro' $id)))
    if ($l.Count -eq 0) { throw ('Visita não encontrada: ' + $id + '.') }
    return $l[0]
}

# C09: a data da visita cai no período do planejamento, quando ela tem um.
function Assert-DataDoPlanejamento($ctx, $idPlanejamento, $data) {
    if ($null -eq $idPlanejamento -or $null -eq $data) { return }
    $l = $ctx.Consultar('SELECT * FROM planejamentos WHERE id=?', @((P 'inteiro' ([int]$idPlanejamento))))
    if ($l.Count -eq 0) { throw ('Planejamento não encontrado: ' + $idPlanejamento + '.') }
    $ini = ([datetime]$l[0].data_inicio).ToString('yyyy-MM-dd')
    $fim = ([datetime]$l[0].data_fim).ToString('yyyy-MM-dd')
    $msg = Test-DataNoPeriodo ([datetime]$data).ToString('yyyy-MM-dd') $ini $fim
    if ($msg) { throw $msg }
}

function Invoke-CriarVisita($a) {
    Assert-EsquemaGravavel
    $codigo = [string]$a.codigo
    $dataAg = $null; if ((Test-TemArg $a 'data_agendada') -and [string]$a.data_agendada -ne '') { $dataAg = Get-DataIsoOuErro ([string]$a.data_agendada) 'data_agendada' }
    $idPl = $null; if (Test-TemArg $a 'id_planejamento') { $idPl = [int]$a.id_planejamento }
    return Invoke-Transacao {
        param($ctx)
        $op = Get-AtendimentoNaTransacao $ctx $codigo
        $papel = Get-PapelEtapa ([string]$op.etapa)
        $msg = Test-VisitaPodeCriar ([bool]$papel.encerra); if ($msg) { throw $msg }
        Assert-DataDoPlanejamento $ctx $idPl $dataAg
        $pend = [int]$ctx.Valor("SELECT COUNT(*) FROM visitas WHERE id_oportunidade=? AND situacao='PENDENTE'", @((P 'inteiro' ([int]$op.id))))
        $msg = Test-SegundaPendente $pend; if ($msg) { throw $msg }
        $tipo = $null; if (Test-TemArg $a 'tipo') { $tipo = Resolve-ValorLista $ctx 'TipoVisita' 'tipo' $a.tipo }
        $resp = $null; if (Test-TemArg $a 'responsavel') { $resp = Resolve-ValorLista $ctx 'Responsavel' 'responsavel' $a.responsavel }
        $hora = $null; if (Test-TemArg $a 'hora_agendada') { $hora = [string]$a.hora_agendada }
        $obs = $null; if (Test-TemArg $a 'observacoes') { $obs = [string]$a.observacoes }
        $agora = Get-Date
        [void]$ctx.Executar('INSERT INTO visitas (id_oportunidade, id_planejamento, situacao, data_agendada, hora_agendada, responsavel, tipo, observacoes, versao, criado_em, alterado_em, alterado_por) VALUES (?,?,?,?,?,?,?,?,1,?,?,?)', @(
            (P 'inteiro' ([int]$op.id)), (P 'inteiro' $idPl), (P 'texto' 'PENDENTE'), (P 'data' $dataAg), (P 'texto' $hora),
            (P 'texto' $resp), (P 'texto' $tipo), (P 'memo' $obs), (P 'data' $agora), (P 'data' $agora), (P 'texto' $script:UsuarioLog)))
        $id = [int]$ctx.Valor('SELECT @@IDENTITY', @())
        $ctx.Log('visitas', $id, 'situacao', '', 'PENDENTE', 'INSERIR')
        return [ordered]@{ gravado = $true; id = $id; codigo = $codigo; situacao = 'PENDENTE'; versao = 1 }
    }
}

$script:SituacoesVisita = @('PENDENTE', 'REALIZADA', 'CANCELADA')
$script:TiposCampoVisita = [ordered]@{
    situacao = 'texto'; data_agendada = 'data'; hora_agendada = 'texto'; responsavel = 'texto'; tipo = 'texto'
    observacoes = 'memo'; data_realizada = 'data'; resultado = 'texto'; relato = 'memo'; motivo_cancelamento = 'texto'
}

# Corpo de atualizar_visita dentro de uma transação já aberta (lotes de lançamento usam o mesmo caminho).
function Set-VisitaNoCtx($ctx, $a) {
    $id = [int]$a.id
    $v = Get-VisitaNaTransacao $ctx $id
    Assert-Versao $v $a
    $novo = [ordered]@{}; foreach ($k in $v.Keys) { $novo[$k] = $v[$k] }
    foreach ($campo in $script:TiposCampoVisita.Keys) {
        if (-not (Test-TemArg $a $campo)) { continue }
        $bruto = $a.$campo
        switch ($campo) {
            'situacao' {
                $s = ([string]$bruto).Trim().ToUpperInvariant()
                if ($s -notin $script:SituacoesVisita) { throw ("situacao deve ser PENDENTE, REALIZADA ou CANCELADA (recebi '" + $bruto + "').") }
                $novo[$campo] = $s
            }
            { $_ -in @('data_agendada', 'data_realizada') } {
                if ([string]$bruto -eq '') { $novo[$campo] = $null } else { $novo[$campo] = Get-DataIsoOuErro ([string]$bruto) $campo }
            }
            'tipo' { $novo[$campo] = Resolve-ValorLista $ctx 'TipoVisita' 'tipo' $bruto }
            'resultado' { $novo[$campo] = Resolve-ValorLista $ctx 'ResultadoVisita' 'resultado' $bruto }
            'responsavel' { $novo[$campo] = Resolve-ValorLista $ctx 'Responsavel' 'responsavel' $bruto }
            default { $novo[$campo] = [string]$bruto }
        }
    }
    $mud = [ordered]@{}
    foreach ($campo in $script:TiposCampoVisita.Keys) {
        if (-not (Test-Igual $v[$campo] $novo[$campo])) {
            $mud[$campo] = @{ Tipo = $script:TiposCampoVisita[$campo]; Novo = $novo[$campo]; Antigo = $v[$campo] }
        }
    }
    if ($mud.Count -eq 0) {
        return [ordered]@{ id = $id; gravado = $false; mensagem = 'Nada a alterar: os valores já estão iguais.'; versao = $v.versao }
    }
    $sit = [string]$novo.situacao
    if ($sit -eq 'REALIZADA') {
        $msg = Test-RealizadaValida ([string]$novo.relato) ([string]$novo.data_realizada); if ($msg) { throw $msg }
    }
    if ($sit -eq 'CANCELADA') {
        $msg = Test-CanceladaValida ([string]$novo.motivo_cancelamento); if ($msg) { throw $msg }
    }
    if ($sit -eq 'PENDENTE' -and [string]$v.situacao -ne 'PENDENTE') {
        $pend = [int]$ctx.Valor("SELECT COUNT(*) FROM visitas WHERE id_oportunidade=? AND situacao='PENDENTE' AND id<>?", @((P 'inteiro' ([int]$v.id_oportunidade)), (P 'inteiro' $id)))
        $msg = Test-SegundaPendente $pend; if ($msg) { throw $msg }
    }
    if ($mud.Contains('data_agendada')) { Assert-DataDoPlanejamento $ctx $v.id_planejamento $novo.data_agendada }
    # V6 e V7: ao virar REALIZADA, o histórico e a última interação do atendimento mudam na mesma transação
    if ($sit -eq 'REALIZADA' -and [string]$v.situacao -ne 'REALIZADA') {
        $opL = $ctx.Consultar('SELECT * FROM oportunidades WHERE id=?', @((P 'inteiro' ([int]$v.id_oportunidade))))
        $op = $opL[0]
        $texto = Get-TextoHistoricoVisita ([datetime]$novo.data_realizada) ([string]$novo.tipo) ([string]$novo.resultado) ([string]$novo.relato)
        $mo = [ordered]@{
            observacoes = @{ Tipo = 'memo'; Novo = (Add-Observacao $op.observacoes $texto); Antigo = $op.observacoes }
            ultima_interacao = @{ Tipo = 'data'; Novo = [datetime]$novo.data_realizada; Antigo = $op.ultima_interacao }
        }
        Update-ComVersao $ctx 'oportunidades' ([int]$op.id) ([int]$op.versao) $mo 'ALTERAR'
    }
    Update-ComVersao $ctx 'visitas' $id ([int]$v.versao) $mud 'ALTERAR'
    return [ordered]@{ id = $id; gravado = $true; situacao = $sit; versao = ([int]$v.versao + 1); alterados = @($mud.Keys) }
}

function Invoke-AtualizarVisita($a) {
    Assert-EsquemaGravavel
    return Invoke-Transacao {
        param($ctx)
        return Set-VisitaNoCtx $ctx $a
    }
}

function Invoke-ListarVisitas($a) {
    $w = New-Object System.Collections.ArrayList; $p = New-Object System.Collections.ArrayList
    if (Test-TemArg $a 'codigo') {
        $codigo = [string]$a.codigo
        if ($codigo -notmatch '^AT-\d{4}-\d{4}$') { throw 'Código inválido. Formato: AT-0000-0000.' }
        [void]$w.Add('o.codigo=?'); [void]$p.Add((P 'texto' $codigo))
    }
    if (Test-TemArg $a 'situacao') {
        $s = ([string]$a.situacao).Trim().ToUpperInvariant()
        if ($s -notin $script:SituacoesVisita) { throw 'situacao deve ser PENDENTE, REALIZADA ou CANCELADA.' }
        [void]$w.Add('v.situacao=?'); [void]$p.Add((P 'texto' $s))
    }
    if (Test-TemArg $a 'id_planejamento') { [void]$w.Add('v.id_planejamento=?'); [void]$p.Add((P 'inteiro' ([int]$a.id_planejamento))) }
    $where = ''; if ($w.Count -gt 0) { $where = ' WHERE ' + ($w -join ' AND ') }
    $n = Get-Limite $a
    # pendentes primeiro, depois a data agendada (como na tela de visitas)
    $sql = 'SELECT TOP ' + $n + ' v.id, v.situacao, v.data_agendada, v.hora_agendada, v.tipo, v.responsavel, v.resultado,' +
           ' v.data_realizada, v.relato, v.motivo_cancelamento, v.id_planejamento, v.versao, o.codigo AS codigo_atendimento,' +
           ' c.empresa, ct.nome AS contato, ct.telefone AS contato_telefone' +
           ' FROM ((visitas v INNER JOIN oportunidades o ON v.id_oportunidade = o.id)' +
           ' INNER JOIN clientes c ON o.id_cliente = c.id) INNER JOIN contatos ct ON o.id_contato = ct.id' +
           $where + ' ORDER BY IIf(v.situacao=''PENDENTE'', 0, 1), v.data_agendada, v.id'
    $l = Invoke-Consulta $sql $p.ToArray()
    return [ordered]@{ quantidade = $l.Count; visitas = $l }
}

function Invoke-AlterarSituacaoVisita($a) {
    return Invoke-AtualizarVisita $a
}

# ---------------------------------------------------------------- planejamento (Fase 4, regras P1-P4 da spec 003)

$script:SituacoesPlanejamento = @('ABERTO', 'ENCERRADO')

function Get-PlanejamentoNaTransacao($ctx, [int]$id) {
    $l = $ctx.Consultar('SELECT * FROM planejamentos WHERE id=?', @((P 'inteiro' $id)))
    if ($l.Count -eq 0) { throw ('Planejamento não encontrado: ' + $id + '.') }
    return $l[0]
}

# Cidade-base: tem de existir em municipios (P1).
function Assert-CidadeBase($ctx, [int]$codIbge) {
    $n = [int]$ctx.Valor('SELECT COUNT(*) FROM municipios WHERE cod_ibge=?', @((P 'inteiro' $codIbge)))
    if ($n -eq 0) { throw ('Cidade-base não encontrada (código IBGE ' + $codIbge + '). Use buscar_municipios ou resolver_cidade.') }
}

function Invoke-CriarPlanejamento($a) {
    Assert-EsquemaGravavel
    $descricao = ([string]$a.descricao).Trim()
    if ($descricao -eq '') { throw 'Informe a descrição do planejamento.' }
    if ($descricao.Length -gt 120) { throw 'Descrição aceita no máximo 120 caracteres.' }
    $cod = [int]$a.cod_municipio
    $raio = [int]$a.raio_km
    $msg = Test-RaioValido $raio; if ($msg) { throw $msg }
    $ini = [string]$a.data_inicio; $fim = [string]$a.data_fim
    $dIni = Get-DataIsoOuErro $ini 'data_inicio'; $dFim = Get-DataIsoOuErro $fim 'data_fim'
    $msg = Test-PeriodoValido $ini $fim; if ($msg) { throw $msg }
    $obs = $null; if (Test-TemArg $a 'observacoes') { $obs = [string]$a.observacoes }
    return Invoke-Transacao {
        param($ctx)
        Assert-CidadeBase $ctx $cod
        $resp = $null; if (Test-TemArg $a 'responsavel') { $resp = Resolve-ValorLista $ctx 'Responsavel' 'responsavel' $a.responsavel }
        $controle = [int]$ctx.Valor('SELECT MAX(controle) FROM planejamentos', @()) + 1
        $codigo = Get-CodigoPlanejamento (Get-Date).Year $controle
        $agora = Get-Date
        [void]$ctx.Executar('INSERT INTO planejamentos (codigo, controle, descricao, cod_municipio_base, raio_km, data_inicio, data_fim, responsavel, situacao, observacoes, versao, criado_em, alterado_em, alterado_por) VALUES (?,?,?,?,?,?,?,?,?,?,1,?,?,?)', @(
            (P 'texto' $codigo), (P 'inteiro' $controle), (P 'texto' $descricao), (P 'inteiro' $cod), (P 'inteiro' $raio),
            (P 'data' $dIni), (P 'data' $dFim), (P 'texto' $resp), (P 'texto' 'ABERTO'), (P 'memo' $obs),
            (P 'data' $agora), (P 'data' $agora), (P 'texto' $script:UsuarioLog)))
        $id = [int]$ctx.Valor('SELECT @@IDENTITY', @())
        $ctx.Log('planejamentos', $id, 'codigo', '', $codigo, 'INSERIR')
        return [ordered]@{ gravado = $true; id = $id; codigo = $codigo; situacao = 'ABERTO'; versao = 1 }
    }
}

function Invoke-AtualizarPlanejamento($a) {
    Assert-EsquemaGravavel
    $id = [int]$a.id
    return Invoke-Transacao {
        param($ctx)
        $p = Get-PlanejamentoNaTransacao $ctx $id
        Assert-Versao $p $a
        if ([string]$p.situacao -ne 'ABERTO') { throw 'Planejamento encerrado: reabra antes de alterar (alterar_situacao_planejamento).' }
        $novo = [ordered]@{}; foreach ($k in $p.Keys) { $novo[$k] = $p[$k] }
        if (Test-TemArg $a 'descricao') { $novo['descricao'] = ([string]$a.descricao).Trim() }
        if (Test-TemArg $a 'cod_municipio') { $novo['cod_municipio_base'] = [int]$a.cod_municipio }
        if (Test-TemArg $a 'raio_km') { $novo['raio_km'] = [int]$a.raio_km }
        if (Test-TemArg $a 'data_inicio') { $novo['data_inicio'] = Get-DataIsoOuErro ([string]$a.data_inicio) 'data_inicio' }
        if (Test-TemArg $a 'data_fim') { $novo['data_fim'] = Get-DataIsoOuErro ([string]$a.data_fim) 'data_fim' }
        if (Test-TemArg $a 'responsavel') { $novo['responsavel'] = Resolve-ValorLista $ctx 'Responsavel' 'responsavel' $a.responsavel }
        if (Test-TemArg $a 'observacoes') { $novo['observacoes'] = [string]$a.observacoes }
        if ([string]$novo['descricao'] -eq '') { throw 'Informe a descrição do planejamento.' }
        $msg = Test-RaioValido ([int]$novo['raio_km']); if ($msg) { throw $msg }
        $msg = Test-PeriodoValido (([datetime]$novo['data_inicio']).ToString('yyyy-MM-dd')) (([datetime]$novo['data_fim']).ToString('yyyy-MM-dd')); if ($msg) { throw $msg }
        Assert-CidadeBase $ctx ([int]$novo['cod_municipio_base'])
        $tipos = [ordered]@{ descricao = 'texto'; cod_municipio_base = 'inteiro'; raio_km = 'inteiro'; data_inicio = 'data'; data_fim = 'data'; responsavel = 'texto'; observacoes = 'memo' }
        $mud = [ordered]@{}
        foreach ($c in $tipos.Keys) {
            if (-not (Test-Igual $p[$c] $novo[$c])) { $mud[$c] = @{ Tipo = $tipos[$c]; Novo = $novo[$c]; Antigo = $p[$c] } }
        }
        if ($mud.Count -eq 0) { return [ordered]@{ id = $id; gravado = $false; mensagem = 'Nada a alterar.'; versao = $p.versao } }
        Update-ComVersao $ctx 'planejamentos' $id ([int]$p.versao) $mud 'ALTERAR'
        return [ordered]@{ id = $id; gravado = $true; versao = ([int]$p.versao + 1); alterados = @($mud.Keys) }
    }
}

# P3 (C16): encerrar recusa com visitas pendentes, a menos que cancelar_pendentes; reabrir volta para ABERTO.
function Invoke-AlterarSituacaoPlanejamento($a) {
    Assert-EsquemaGravavel
    $id = [int]$a.id
    $sit = ([string]$a.situacao).Trim().ToUpperInvariant()
    if ($sit -notin $script:SituacoesPlanejamento) { throw 'situacao deve ser ABERTO ou ENCERRADO.' }
    $cancelar = (Test-TemArg $a 'cancelar_pendentes') -and [bool]$a.cancelar_pendentes
    return Invoke-Transacao {
        param($ctx)
        $p = Get-PlanejamentoNaTransacao $ctx $id
        Assert-Versao $p $a
        if ([string]$p.situacao -eq $sit) { return [ordered]@{ id = $id; gravado = $false; mensagem = 'Já está nessa situação.'; versao = $p.versao } }
        $canceladas = 0
        if ($sit -eq 'ENCERRADO') {
            $pend = @($ctx.Consultar("SELECT id, versao FROM visitas WHERE id_planejamento=? AND situacao='PENDENTE'", @((P 'inteiro' $id))))
            $msg = Test-EncerramentoPermitido $pend.Count $cancelar; if ($msg) { throw $msg }
            foreach ($v in $pend) {
                Update-ComVersao $ctx 'visitas' ([int]$v.id) ([int]$v.versao) @{
                    situacao = @{ Tipo = 'texto'; Novo = 'CANCELADA'; Antigo = 'PENDENTE' }
                    motivo_cancelamento = @{ Tipo = 'texto'; Novo = 'planejamento encerrado'; Antigo = $null }
                } 'ALTERAR'
                $canceladas++
            }
        }
        Update-ComVersao $ctx 'planejamentos' $id ([int]$p.versao) ([ordered]@{
            situacao = @{ Tipo = 'texto'; Novo = $sit; Antigo = [string]$p.situacao } }) 'ALTERAR'
        return [ordered]@{ id = $id; gravado = $true; situacao = $sit; versao = ([int]$p.versao + 1); visitas_canceladas = $canceladas }
    }
}

# ---------------------------------------------------------------- candidatos e geração de visitas (Fase 4.4 e 4.9b)

# Municípios por UF|nome_busca, para achar a cidade de cada atendimento.
function Get-MapaMunicipios($ctx) {
    $mapa = @{}
    foreach ($m in $ctx.Consultar('SELECT cod_ibge, nome_busca, uf, lat, lon FROM municipios', @())) {
        $mapa[([string]$m.uf + '|' + [string]$m.nome_busca)] = $m
    }
    return $mapa
}

# Candidatos do planejamento: atendimentos em aberto a até raio_km da cidade-base.
# Ordem: qualificação (ordem da lista), dias parado (mais tempo primeiro), km (menor primeiro).
function Get-CandidatosPlanejamento($ctx, $pl) {
    $base = $ctx.Consultar('SELECT cod_ibge, lat, lon FROM municipios WHERE cod_ibge=?', @((P 'inteiro' ([int]$pl.cod_municipio_base))))
    if ($base.Count -eq 0) { throw 'Cidade-base do planejamento não está em municipios.' }
    $mapa = Get-MapaMunicipios $ctx
    $ordens = @{}
    foreach ($q in $ctx.Consultar("SELECT valor, ordem FROM listas WHERE tipo='Qualificacao'", @())) { $ordens[[string]$q.valor] = [int]$q.ordem }
    # pendentes de todos os atendimentos numa leitura só (nada de consulta por candidato)
    $pendPorAt = @{}
    foreach ($pv in $ctx.Consultar("SELECT id_oportunidade, COUNT(*) AS n FROM visitas WHERE situacao='PENDENTE' GROUP BY id_oportunidade", @())) { $pendPorAt[[int]$pv.id_oportunidade] = [int]$pv.n }
    $candidatos = New-Object System.Collections.ArrayList
    $naoReconhecidas = New-Object System.Collections.ArrayList
    $abertos = $ctx.Consultar('SELECT o.id, o.id_cliente, o.codigo, o.etapa, o.ultima_interacao, c.empresa, c.cidade, c.uf, c.qualificacao, ct.nome AS contato, ct.telefone AS contato_telefone FROM (oportunidades o INNER JOIN clientes c ON o.id_cliente = c.id) INNER JOIN contatos ct ON o.id_contato = ct.id', @())
    foreach ($o in $abertos) {
        if ((Get-PapelEtapa ([string]$o.etapa)).encerra) { continue }
        $chave = ([string]$o.uf + '|' + (ConvertTo-NomeBusca ([string]$o.cidade)))
        if (-not $mapa.ContainsKey($chave)) {
            [void]$naoReconhecidas.Add([ordered]@{ codigo = $o.codigo; empresa = $o.empresa; cidade = $o.cidade; uf = $o.uf })
            continue
        }
        $m = $mapa[$chave]
        $km = Get-DistanciaKm ([double]$base[0].lat) ([double]$base[0].lon) ([double]$m.lat) ([double]$m.lon)
        if ($km -gt [double]$pl.raio_km) { continue }
        $qual = [string]$o.qualificacao
        $ordemQ = 999; if ($ordens.ContainsKey($qual)) { $ordemQ = $ordens[$qual] }
        $pendentes = 0; if ($pendPorAt.ContainsKey([int]$o.id)) { $pendentes = $pendPorAt[[int]$o.id] }
        [void]$candidatos.Add([pscustomobject]@{
            codigo = [string]$o.codigo; empresa = [string]$o.empresa; cidade = [string]$o.cidade; uf = [string]$o.uf
            qualificacao = $qual; km = [math]::Round($km, 1); dias_parado = (Get-DiasParado $o.ultima_interacao)
            contato = [string]$o.contato; contato_telefone = [string]$o.contato_telefone
            visita_pendente = ($pendentes -gt 0); ordem_qualificacao = $ordemQ })
    }
    $ordenados = @($candidatos | Sort-Object -Property @{ e = 'ordem_qualificacao'; Ascending = $true }, @{ e = { if ($null -eq $_.dias_parado) { 0 } else { $_.dias_parado } }; Descending = $true }, @{ e = 'km'; Ascending = $true })

    # DEV-27 (ciclo 016): empresas ATIVAS do raio sem atendimento em aberto - pré-clientes inclusive (decisão 11).
    # Vêm depois dos candidatos, na ordem qualificação e km; a lista dos candidatos não muda.
    $comAberto = @{}
    foreach ($o in $abertos) { if (-not (Get-PapelEtapa ([string]$o.etapa)).encerra) { $comAberto[[int]$o.id_cliente] = $true } }
    $sem = New-Object System.Collections.ArrayList
    foreach ($c in $ctx.Consultar('SELECT id, codigo_cliente, estagio, empresa, cidade, uf, qualificacao, ativo FROM clientes', @())) {
        if ($null -ne $c.ativo -and -not [bool]$c.ativo) { continue }
        if ($comAberto.ContainsKey([int]$c.id)) { continue }
        $chave = ([string]$c.uf + '|' + (ConvertTo-NomeBusca ([string]$c.cidade)))
        if (-not $mapa.ContainsKey($chave)) { continue }
        $m = $mapa[$chave]
        $km = Get-DistanciaKm ([double]$base[0].lat) ([double]$base[0].lon) ([double]$m.lat) ([double]$m.lon)
        if ($km -gt [double]$pl.raio_km) { continue }
        $qual = [string]$c.qualificacao
        $ordemQ = 999; if ($ordens.ContainsKey($qual)) { $ordemQ = $ordens[$qual] }
        [void]$sem.Add([pscustomobject]@{
            id_cliente = [int]$c.id; codigo_cliente = $c.codigo_cliente; estagio = [string]$c.estagio; empresa = [string]$c.empresa
            cidade = [string]$c.cidade; uf = [string]$c.uf; qualificacao = $qual; km = [math]::Round($km, 1); ordem_qualificacao = $ordemQ })
    }
    $semOrd = @($sem | Sort-Object -Property ordem_qualificacao, km, empresa)
    return [pscustomobject]@{ candidatos = $ordenados; nao_reconhecidas = @($naoReconhecidas); sem_atendimento = $semOrd }
}

$script:LimiteSemAtendimento = 200
function Invoke-ListarCandidatos($a) {
    $id = [int]$a.id_planejamento
    return Invoke-Transacao {
        param($ctx)
        $pl = Get-PlanejamentoNaTransacao $ctx $id
        $r = Get-CandidatosPlanejamento $ctx $pl
        $sem = @($r.sem_atendimento)
        return [ordered]@{ id = $id; codigo = $pl.codigo; raio_km = $pl.raio_km; quantidade = $r.candidatos.Count
            candidatos = @($r.candidatos); nao_reconhecidas = $r.nao_reconhecidas
            quantidade_sem_atendimento = $sem.Count; sem_atendimento_truncado = ($sem.Count -gt $script:LimiteSemAtendimento)
            sem_atendimento = @($sem | Select-Object -First $script:LimiteSemAtendimento) }
    }
}

# Gera uma visita PENDENTE por atendimento no planejamento. simular (padrão true) só confere.
# Gravação tudo ou nada: se um atendimento é recusado (C08, C13), nada é gravado.
function Invoke-GerarVisitas($a) {
    Assert-EsquemaGravavel
    $id = [int]$a.id_planejamento
    $codigos = @($a.codigos)
    if ($codigos.Count -lt 1 -or $codigos.Count -gt 100) { throw 'Informe de 1 a 100 atendimentos (codigos).' }
    $simular = $true; if (Test-TemArg $a 'simular') { $simular = [bool]$a.simular }
    return Invoke-Transacao {
        param($ctx)
        $pl = Get-PlanejamentoNaTransacao $ctx $id
        if ([string]$pl.situacao -ne 'ABERTO') { throw 'Planejamento encerrado: reabra antes de gerar visitas.' }
        $linhas = @()
        $gravadas = @()
        foreach ($codigo in $codigos) {
            $msg = $null; $op = $null
            try { $op = Get-AtendimentoNaTransacao $ctx ([string]$codigo) } catch { $msg = $_.Exception.Message }
            if ($null -eq $msg) {
                $msg = Test-VisitaPodeCriar ([bool](Get-PapelEtapa ([string]$op.etapa)).encerra)
            }
            if ($null -eq $msg) {
                $pend = [int]$ctx.Valor("SELECT COUNT(*) FROM visitas WHERE id_oportunidade=? AND situacao='PENDENTE'", @((P 'inteiro' ([int]$op.id))))
                $msg = Test-SegundaPendente $pend
            }
            $linhas += [pscustomobject]@{ codigo = [string]$codigo; ok = ($null -eq $msg); mensagem = [string]$msg }
            if ($null -eq $msg -and -not $simular) {
                $agora = Get-Date
                [void]$ctx.Executar('INSERT INTO visitas (id_oportunidade, id_planejamento, situacao, versao, criado_em, alterado_em, alterado_por) VALUES (?,?,?,1,?,?,?)', @(
                    (P 'inteiro' ([int]$op.id)), (P 'inteiro' $id), (P 'texto' 'PENDENTE'), (P 'data' $agora), (P 'data' $agora), (P 'texto' $script:UsuarioLog)))
                $nova = [int]$ctx.Valor('SELECT @@IDENTITY', @())
                $ctx.Log('visitas', $nova, 'situacao', '', 'PENDENTE', 'INSERIR')
                $gravadas += $nova
            }
        }
        $recusadas = @($linhas | Where-Object { -not $_.ok })
        if (-not $simular -and $recusadas.Count -gt 0) {
            throw ('Nada gravado: ' + $recusadas.Count + ' atendimento(s) recusado(s) (ex.: ' + $recusadas[0].codigo + ': ' + $recusadas[0].mensagem + ').')
        }
        return [ordered]@{ simulado = $simular; planejamento = $pl.codigo; total = $linhas.Count
            ok = @($linhas | Where-Object { $_.ok }).Count; recusadas = $recusadas.Count; linhas = $linhas; visitas_criadas = $gravadas }
    }
}

# ---------------------------------------------------------------- lançamento em lote (Fase 4.5)

# Roda os itens na mesma transação, com as regras de atualizar_visita. simular (padrão true) desfaz no fim.
# Gravação tudo ou nada: um item recusado desfaz todos.
function Invoke-LoteVisitas($a) {
    $itens = @($a.itens)
    if ($itens.Count -lt 1 -or $itens.Count -gt 100) { throw 'Informe de 1 a 100 itens.' }
    $simular = $true; if (Test-TemArg $a 'simular') { $simular = [bool]$a.simular }
    if (-not $simular) { Assert-EsquemaGravavel }
    $resultados = New-Object System.Collections.ArrayList
    $marca = 'SIMULACAO-CONCLUIDA'
    try {
        [void](Invoke-Transacao {
            param($ctx)
            foreach ($it in $itens) {
                $linha = [ordered]@{ id = $it.id; ok = $true; mensagem = ''; versao = $null }
                try {
                    $r = Set-VisitaNoCtx $ctx $it
                    $linha['versao'] = $r.versao
                    if ($r.gravado -eq $false) { $linha['mensagem'] = [string]$r.mensagem }
                } catch {
                    $linha['ok'] = $false
                    $linha['mensagem'] = $_.Exception.Message
                }
                [void]$resultados.Add($linha)
            }
            $recusadas = @($resultados | Where-Object { -not $_.ok }).Count
            if ($recusadas -gt 0 -and -not $simular) {
                throw ('Nada gravado: ' + $recusadas + ' item(ns) recusado(s). Ex.: visita ' + $resultados[0].id + ': ' + (@($resultados | Where-Object { -not $_.ok })[0].mensagem))
            }
            if ($simular) { throw $marca }
            return $null
        })
    } catch {
        if ($_.Exception.Message -ne $marca) { throw }
    }
    $recusadas = @($resultados | Where-Object { -not $_.ok }).Count
    return [ordered]@{ simulado = $simular; total = $resultados.Count; ok = ($resultados.Count - $recusadas); recusadas = $recusadas; linhas = @($resultados) }
}

# lancar_datas: data e hora agendadas das visitas pendentes (V3, C09 no planejamento).
function Invoke-LancarDatas($a) { return Invoke-LoteVisitas $a }

# lancar_resultados: realizada (V4, V6, V7) ou cancelada (V5) em lote.
function Invoke-LancarResultados($a) { return Invoke-LoteVisitas $a }

# ---------------------------------------------------------------- leitura de planejamentos e municípios (Fase 4.9a e 4.9b)

function Invoke-ListarPlanejamentos($a) {
    $w = New-Object System.Collections.ArrayList; $p = New-Object System.Collections.ArrayList
    if (Test-TemArg $a 'situacao') {
        $s = ([string]$a.situacao).Trim().ToUpperInvariant()
        if ($s -notin $script:SituacoesPlanejamento) { throw 'situacao deve ser ABERTO ou ENCERRADO.' }
        [void]$w.Add('p.situacao=?'); [void]$p.Add((P 'texto' $s))
    }
    $where = ''; if ($w.Count -gt 0) { $where = ' WHERE ' + ($w -join ' AND ') }
    $n = Get-Limite $a
    $l = Invoke-Consulta ('SELECT TOP ' + $n + ' p.id, p.codigo, p.descricao, p.raio_km, p.data_inicio, p.data_fim, p.situacao, p.responsavel, p.versao, m.nome AS cidade_base, m.uf AS uf_base' +
                          ' FROM planejamentos p LEFT JOIN municipios m ON p.cod_municipio_base = m.cod_ibge' + $where + ' ORDER BY p.codigo DESC') $p.ToArray()
    return [ordered]@{ quantidade = $l.Count; planejamentos = $l }
}

function Invoke-ObterPlanejamento($a) {
    $sql = 'SELECT p.*, m.nome AS cidade_base, m.uf AS uf_base FROM planejamentos p LEFT JOIN municipios m ON p.cod_municipio_base = m.cod_ibge WHERE '
    if (Test-TemArg $a 'id') { $l = Invoke-Consulta ($sql + 'p.id=?') @((P 'inteiro' ([int]$a.id))) }
    elseif (Test-TemArg $a 'codigo') { $l = Invoke-Consulta ($sql + 'p.codigo=?') @((P 'texto' ([string]$a.codigo))) }
    else { throw 'Informe id ou codigo do planejamento.' }
    if ($l.Count -eq 0) { throw 'Planejamento não encontrado.' }
    $pl = $l[0]
    $cont = Invoke-Consulta "SELECT situacao, COUNT(*) AS n FROM visitas WHERE id_planejamento=? GROUP BY situacao" @((P 'inteiro' ([int]$pl.id)))
    $resumo = [ordered]@{ PENDENTE = 0; REALIZADA = 0; CANCELADA = 0 }
    foreach ($c in $cont) { $resumo[[string]$c.situacao] = [int]$c.n }
    $pl['visitas'] = $resumo
    return $pl
}

# Municípios por nome (sem acento, sem caixa), com UF opcional. Até 50 linhas.
function Invoke-BuscarMunicipios($a) {
    $texto = ConvertTo-NomeBusca ([string]$a.texto)
    if ($texto.Length -lt 2) { throw 'Informe pelo menos 2 letras do nome.' }
    $p = New-Object System.Collections.ArrayList
    $w = @('nome_busca LIKE ?'); [void]$p.Add((P 'texto' ('%' + $texto + '%')))
    if (Test-TemArg $a 'uf') { $w += 'uf=?'; [void]$p.Add((P 'texto' ([string]$a.uf).ToUpperInvariant())) }
    $l = Invoke-Consulta ('SELECT TOP 50 cod_ibge, nome, uf FROM municipios WHERE ' + ($w -join ' AND ') + ' ORDER BY nome') $p.ToArray()
    return [ordered]@{ quantidade = $l.Count; municipios = $l }
}

# Cidade informada pelo nome e UF: resolvida se houver exatamente uma; senão, a lista para escolher.
function Invoke-ResolverCidade($a) {
    $nome = ConvertTo-NomeBusca ([string]$a.cidade)
    $uf = ([string]$a.uf).Trim().ToUpperInvariant()
    if ($nome -eq '') { throw 'Informe a cidade.' }
    if ($uf -eq '') { throw 'Informe a UF.' }
    $l = Invoke-Consulta 'SELECT cod_ibge, nome, uf FROM municipios WHERE uf=? AND nome_busca=?' @((P 'texto' $uf), (P 'texto' $nome))
    if ($l.Count -eq 1) { return [ordered]@{ resolvida = $true; municipio = $l[0] } }
    return [ordered]@{ resolvida = $false; motivo = $(if ($l.Count -eq 0) { 'Cidade não encontrada na UF (use buscar_municipios).' } else { 'Mais de uma cidade com esse nome.' }); opcoes = $l }
}

# ---------------------------------------------------------------- exportação do roteiro (Fase 5)

$script:ColunasRoteiro = @('tipo', 'visita', 'atendimento', 'empresa', 'contato', 'telefone', 'cidade', 'uf', 'latitude', 'longitude', 'km_base', 'data_agendada', 'hora_agendada', 'tipo_visita', 'responsavel', 'situacao')

# CSV do roteiro: UTF-8 com BOM, ';', texto entre aspas, número com ponto decimal, data AAAA-MM-DD.
function Write-RoteiroCsv([string]$arquivo, $linhas) {
    $tmp = $arquivo + '.tmp'
    $cult = [System.Globalization.CultureInfo]::InvariantCulture
    $w = New-Object System.IO.StreamWriter($tmp, $false, (New-Object System.Text.UTF8Encoding($true)))
    try {
        $w.NewLine = "`r`n"
        $w.WriteLine(($script:ColunasRoteiro -join ';'))
        foreach ($l in $linhas) {
            $vals = New-Object System.Collections.ArrayList
            foreach ($c in $script:ColunasRoteiro) {
                $v = $l[$c]
                if ($null -eq $v -or $v -is [System.DBNull]) { [void]$vals.Add('""') }
                elseif ($v -is [double] -or $v -is [single] -or $v -is [decimal]) { [void]$vals.Add(([double]$v).ToString('0.######', $cult)) }
                elseif ($v -is [int] -or $v -is [long]) { [void]$vals.Add([string]$v) }
                else { [void]$vals.Add('"' + ([string]$v).Replace('"', '""') + '"') }
            }
            $w.WriteLine(($vals -join ';'))
        }
    } finally { $w.Close() }
    if (Test-Path -LiteralPath $arquivo) { Remove-Item -LiteralPath $arquivo -Force }
    Move-Item -LiteralPath $tmp -Destination $arquivo
}

function Assert-Esquema11 {
    $v = Get-VersaoEsquema
    if ($script:EsquemasSuportados -notcontains $v) { throw ("Banco no esquema '" + $v + "': o roteiro exige o esquema " + ($script:EsquemasSuportados -join ', ') + '. Rode a montagem da base (MONTAR-FRONTEND.bat).') }
}

# Roteiro do planejamento: a linha BASE primeiro, depois uma linha por visita PENDENTE, em ordem de data.
# Devolve só o caminho e o número de visitas (a ordem do roteiro é feita fora do sistema).
function Invoke-ExportarRoteiro($a) {
    Assert-Esquema11
    $id = 0
    if (Test-TemArg $a 'id_planejamento') { $id = [int]$a.id_planejamento }
    elseif (Test-TemArg $a 'codigo') { $id = [int](Invoke-Valor 'SELECT id FROM planejamentos WHERE codigo=?' @((P 'texto' ([string]$a.codigo)))) }
    if ($id -eq 0) { throw 'Informe id_planejamento ou codigo do planejamento.' }
    return Invoke-Transacao {
        param($ctx)
        $pl = Get-PlanejamentoNaTransacao $ctx $id
        $base = $ctx.Consultar('SELECT nome, uf, lat, lon FROM municipios WHERE cod_ibge=?', @((P 'inteiro' ([int]$pl.cod_municipio_base))))
        if ($base.Count -eq 0) { throw 'Cidade-base não está em municipios.' }
        $mapa = Get-MapaMunicipios $ctx
        $visitas = @($ctx.Consultar("SELECT v.id, v.data_agendada, v.hora_agendada, v.tipo, v.responsavel, o.codigo, c.empresa, c.cidade, c.uf, ct.nome AS contato, ct.telefone AS contato_telefone FROM ((visitas v INNER JOIN oportunidades o ON v.id_oportunidade = o.id) INNER JOIN clientes c ON o.id_cliente = c.id) INNER JOIN contatos ct ON o.id_contato = ct.id WHERE v.id_planejamento=? AND v.situacao='PENDENTE' ORDER BY v.data_agendada, v.hora_agendada, v.id", @((P 'inteiro' $id))))
        $linhas = New-Object System.Collections.ArrayList
        [void]$linhas.Add([ordered]@{ tipo = 'BASE'; visita = $null; atendimento = $null; empresa = $null; contato = $null; telefone = $null
            cidade = [string]$base[0].nome; uf = [string]$base[0].uf; latitude = [double]$base[0].lat; longitude = [double]$base[0].lon
            km_base = 0; data_agendada = $null; hora_agendada = $null; tipo_visita = $null; responsavel = $null; situacao = $null })
        foreach ($v in $visitas) {
            $m = $null
            $chave = ([string]$v.uf + '|' + (ConvertTo-NomeBusca ([string]$v.cidade)))
            if ($mapa.ContainsKey($chave)) { $m = $mapa[$chave] }
            $km = $null; $lat = $null; $lon = $null
            if ($null -ne $m) {
                $lat = [double]$m.lat; $lon = [double]$m.lon
                $km = [math]::Round((Get-DistanciaKm ([double]$base[0].lat) ([double]$base[0].lon) $lat $lon), 1)
            }
            $dt = $null; if ($null -ne $v.data_agendada) { $dt = ([datetime]$v.data_agendada).ToString('yyyy-MM-dd') }
            [void]$linhas.Add([ordered]@{ tipo = 'VISITA'; visita = [int]$v.id; atendimento = [string]$v.codigo; empresa = [string]$v.empresa
                contato = [string]$v.contato; telefone = [string]$v.contato_telefone; cidade = [string]$v.cidade; uf = [string]$v.uf
                latitude = $lat; longitude = $lon; km_base = $km; data_agendada = $dt; hora_agendada = [string]$v.hora_agendada
                tipo_visita = [string]$v.tipo; responsavel = [string]$v.responsavel; situacao = 'PENDENTE' })
        }
        $inicio = ([datetime]$pl.data_inicio).ToString('yyyy-MM-dd')
        $fim = ([datetime]$pl.data_fim).ToString('yyyy-MM-dd')
        # DEV-23: na raiz de extratos (execucao\extratos), uma subpasta com manifesto - nunca em Documentos
        $pasta = New-PastaExtrato ('roteiro-' + $pl.codigo)
        $man = New-ManifestoExtrato 'roteiro' $a 'roteiro de visitas'
        $man['planejamento'] = [string]$pl.codigo
        Write-ManifestoExtrato $pasta $man
        $nomeArq = ($pl.codigo + '_' + $inicio + '_a_' + $fim + '.csv')
        $arquivo = Join-Path $pasta $nomeArq
        Write-RoteiroCsv $arquivo @($linhas)
        $man['arquivos'] = [ordered]@{ roteiro = [ordered]@{ arquivo = $nomeArq; linhas = $visitas.Count; sha256 = (Get-HashArquivo $arquivo) } }
        $man['em_processamento'] = $false
        Write-ManifestoExtrato $pasta $man
        return [ordered]@{ caminho = $arquivo; pasta = $pasta; planejamento = [string]$pl.codigo; linhas = $visitas.Count }
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
                if ($defs[$c] -eq 'moeda') { $v = ConvertTo-Moeda $c $v } else { $v = ConvertTo-Inteiro $c $v }
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
$script:PropIdOperacao = [ordered]@{ type = 'string'; pattern = '^[A-Za-z0-9:_.|-]{8,120}$'; description = 'Identificador único desta operação (ex.: um GUID gerado por você). Repetir a chamada com o mesmo id e o mesmo conteúdo devolve o resultado da primeira, sem criar outro registro (reenvio seguro depois de resultado indeterminado).' }
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
            id_operacao = $script:PropIdOperacao
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
        description = ('desativar: a empresa e seus contatos ficam inativos (atendimentos em aberto continuam no funil). reativar: reativa SÓ a empresa; os contatos continuam inativos e voltam um a um por atualizar_contato (ativo=true). ' +
                       'excluir_pre_cliente: apaga de vez um pré-cliente sem contato nem atendimento (cliente com código nunca é excluído). Leia antes com obter_cliente e passe a "versao"; confirme com o usuário.')
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; required = @('acao', 'versao'); properties = [ordered]@{
            codigo_cliente = [ordered]@{ type = 'integer' }; id_cliente = [ordered]@{ type = 'integer' }
            versao = [ordered]@{ type = 'integer' }
            acao = [ordered]@{ type = 'string'; enum = @('desativar', 'reativar', 'excluir_pre_cliente') }
            autor = $script:PropAutor } }
    },
    [ordered]@{
        name = 'duplicar_atendimento'; title = 'Duplicar atendimento'; handler = 'Invoke-DuplicarAtendimento'; annotations = $script:Gravacao
        description = ('Abre um atendimento NOVO copiando do atendimento de origem: contato, responsável, origem, prioridade, categoria, tipo de venda, máquina, família e última interação. ' +
                       'Nascem vazios valor, orçamento, data da proposta, próxima ação, desfecho, motivo, retomada e tentativas; dt_entrada = hoje; etapa = a inicial; o anterior é a origem (encerrado também pode ser duplicado). ' +
                       'simular = true (padrão) só devolve os dados da cópia; false grava pelo mesmo caminho do abrir_atendimento (aceita id_operacao).')
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; required = @('codigo'); properties = [ordered]@{
            codigo = [ordered]@{ type = 'string'; pattern = '^AT-\d{4}-\d{4}$'; description = 'Atendimento de origem.' }
            simular = [ordered]@{ type = 'boolean'; default = $true; description = 'true: só mostra a cópia. false: grava.' }
            id_operacao = [ordered]@{ type = 'string'; maxLength = 120; description = 'Identificador da criação: repetir com o mesmo id não duplica.' }
            promover_pre_cliente = [ordered]@{ type = 'boolean'; description = 'Necessário se a empresa ainda for pré-cliente.' }
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
            id_operacao = $script:PropIdOperacao
            autor = $script:PropAutor } }
    },
    [ordered]@{
        name = 'gerenciar_lista'; title = 'Manter listas suspensas'; handler = 'Invoke-GerenciarLista'; annotations = $script:Gravacao
        description = ('Inclui, ativa, desativa ou reordena um valor de lista suspensa (Etapa, Responsavel, Familia...). Valor nunca é apagado nem renomeado: registros antigos continuam com ele. ' +
                       'Desativar Etapa ou Responsavel afeta o funil e a fila de todos: confirme com o usuário.')
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; required = @('tipo', 'valor', 'acao'); properties = [ordered]@{
            tipo = [ordered]@{ type = 'string' }; valor = [ordered]@{ type = 'string'; maxLength = 120 }
            acao = [ordered]@{ type = 'string'; enum = @('incluir', 'ativar', 'desativar', 'ordenar', 'definir_papel') }
            encerra = [ordered]@{ type = 'boolean' }; resultado = [ordered]@{ type = 'string'; enum = @('', 'GANHO', 'PERDA') }
            quadro = [ordered]@{ type = 'string'; enum = @('FUNIL', 'PROSPECCAO') }; exige_familia = [ordered]@{ type = 'boolean' }
            confirmar = [ordered]@{ type = 'boolean'; description = 'Obrigatório (true) ao mudar encerra de etapa com atendimentos.' }
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
            id_operacao = $script:PropIdOperacao
            autor = $script:PropAutor } }
    },
    [ordered]@{
        name = 'atualizar_contato'; title = 'Alterar contato'; handler = 'Invoke-AtualizarContato'; annotations = $script:Gravacao
        description = ('Altera nome, cargo, telefone, e-mail, observações ou situação (ativo) de um contato. Leia antes com obter_contato e passe a "versao". ' +
                       'Contato não se exclui: ativo=false desativa. O contato pertence a UMA empresa e não é transferido: se a pessoa mudou de empresa, desative-o e cadastre um novo (criar_contato) na empresa nova. Envie só os campos que mudam.')
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; required = @('versao'); properties = [ordered]@{
            codigo_contato = [ordered]@{ type = 'string'; pattern = '^[Cc][Tt]-\d{4}-\d{4}$' }; id_contato = [ordered]@{ type = 'integer' }
            versao = [ordered]@{ type = 'integer'; description = 'Versão lida em obter_contato' }
            nome = [ordered]@{ type = 'string'; maxLength = 120 }; cargo = [ordered]@{ type = @('string', 'null'); maxLength = 60 }
            telefone = [ordered]@{ type = @('string', 'null') }; email = [ordered]@{ type = @('string', 'null'); maxLength = 120 }
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
        name = 'criar_visita'; title = 'Criar visita'; handler = 'Invoke-CriarVisita'; annotations = $script:Gravacao
        description = 'Cria uma visita PENDENTE num atendimento aberto. Um atendimento tem no máximo uma visita pendente por vez. Com id_planejamento, a data tem de cair no período do planejamento.'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; required = @('codigo'); properties = [ordered]@{
            codigo = [ordered]@{ type = 'string'; pattern = '^AT-\d{4}-\d{4}$' }
            data_agendada = [ordered]@{ type = 'string'; description = 'AAAA-MM-DD' }
            hora_agendada = [ordered]@{ type = 'string'; maxLength = 20 }
            tipo = [ordered]@{ type = 'string'; description = 'Valor da lista TipoVisita (listar_opcoes)' }
            responsavel = [ordered]@{ type = 'string' }
            observacoes = [ordered]@{ type = 'string' }
            id_planejamento = [ordered]@{ type = 'integer' }
            autor = $script:PropAutor } }
    },
    [ordered]@{
        name = 'atualizar_visita'; title = 'Atualizar visita'; handler = 'Invoke-AtualizarVisita'; annotations = $script:Gravacao
        description = 'Muda a situação e os dados de uma visita (PENDENTE, REALIZADA ou CANCELADA). REALIZADA exige data e relato e grava o histórico do atendimento; CANCELADA exige motivo. Envie a versão lida (conflito se mudou).'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; required = @('id', 'versao'); properties = [ordered]@{
            id = [ordered]@{ type = 'integer' }
            versao = [ordered]@{ type = 'integer' }
            situacao = [ordered]@{ type = 'string'; enum = @('PENDENTE', 'REALIZADA', 'CANCELADA') }
            data_agendada = [ordered]@{ type = 'string'; description = 'AAAA-MM-DD' }
            hora_agendada = [ordered]@{ type = 'string'; maxLength = 20 }
            tipo = [ordered]@{ type = 'string' }
            responsavel = [ordered]@{ type = 'string' }
            observacoes = [ordered]@{ type = 'string' }
            data_realizada = [ordered]@{ type = 'string'; description = 'AAAA-MM-DD' }
            resultado = [ordered]@{ type = 'string' }
            relato = [ordered]@{ type = 'string' }
            motivo_cancelamento = [ordered]@{ type = 'string' }
            autor = $script:PropAutor } }
    },
    [ordered]@{
        name = 'listar_visitas'; title = 'Listar visitas'; handler = 'Invoke-ListarVisitas'; annotations = $script:SoLeitura
        description = 'Visitas, com o atendimento, a empresa e o contato (pendentes primeiro). Filtre por codigo do atendimento, situacao ou planejamento.'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; properties = [ordered]@{
            codigo = [ordered]@{ type = 'string'; pattern = '^AT-\d{4}-\d{4}$' }
            situacao = [ordered]@{ type = 'string'; enum = @('PENDENTE', 'REALIZADA', 'CANCELADA') }
            id_planejamento = [ordered]@{ type = 'integer' }
            limite = [ordered]@{ type = 'integer'; maximum = 200 } } }
    },
    [ordered]@{
        name = 'alterar_situacao_visita'; title = 'Alterar situação da visita'; handler = 'Invoke-AlterarSituacaoVisita'; annotations = $script:Gravacao
        description = 'Realiza ou cancela uma visita: REALIZADA exige data_realizada, resultado e relato (grava o histórico do atendimento); CANCELADA exige motivo_cancelamento. Envie a versão lida.'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; required = @('id', 'versao', 'situacao'); properties = [ordered]@{
            id = [ordered]@{ type = 'integer' }
            versao = [ordered]@{ type = 'integer' }
            situacao = [ordered]@{ type = 'string'; enum = @('PENDENTE', 'REALIZADA', 'CANCELADA') }
            data_realizada = [ordered]@{ type = 'string'; description = 'AAAA-MM-DD ou AAAA-MM-DD HH:mm' }
            resultado = [ordered]@{ type = 'string' }
            relato = [ordered]@{ type = 'string' }
            motivo_cancelamento = [ordered]@{ type = 'string' }
            autor = $script:PropAutor } }
    },
    [ordered]@{
        name = 'criar_planejamento'; title = 'Criar planejamento'; handler = 'Invoke-CriarPlanejamento'; annotations = $script:Gravacao
        description = 'Cria um planejamento por região: cidade-base (código IBGE), raio de 1 a 500 km e período (início não depois do fim). Código PL-AAAA-NNN gerado.'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; required = @('descricao', 'cod_municipio', 'raio_km', 'data_inicio', 'data_fim'); properties = [ordered]@{
            descricao = [ordered]@{ type = 'string'; maxLength = 120 }
            cod_municipio = [ordered]@{ type = 'integer'; description = 'Código IBGE da cidade-base' }
            raio_km = [ordered]@{ type = 'integer'; minimum = 1; maximum = 500 }
            data_inicio = [ordered]@{ type = 'string'; description = 'AAAA-MM-DD' }
            data_fim = [ordered]@{ type = 'string'; description = 'AAAA-MM-DD' }
            responsavel = [ordered]@{ type = 'string' }
            observacoes = [ordered]@{ type = 'string' }
            autor = $script:PropAutor } }
    },
    [ordered]@{
        name = 'atualizar_planejamento'; title = 'Atualizar planejamento'; handler = 'Invoke-AtualizarPlanejamento'; annotations = $script:Gravacao
        description = 'Altera um planejamento ABERTO (descrição, cidade-base, raio, período, responsável, observações). Envie a versão lida.'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; required = @('id', 'versao'); properties = [ordered]@{
            id = [ordered]@{ type = 'integer' }
            versao = [ordered]@{ type = 'integer' }
            descricao = [ordered]@{ type = 'string'; maxLength = 120 }
            cod_municipio = [ordered]@{ type = 'integer' }
            raio_km = [ordered]@{ type = 'integer'; minimum = 1; maximum = 500 }
            data_inicio = [ordered]@{ type = 'string' }
            data_fim = [ordered]@{ type = 'string' }
            responsavel = [ordered]@{ type = 'string' }
            observacoes = [ordered]@{ type = 'string' }
            autor = $script:PropAutor } }
    },
    [ordered]@{
        name = 'alterar_situacao_planejamento'; title = 'Encerrar ou reabrir planejamento'; handler = 'Invoke-AlterarSituacaoPlanejamento'; annotations = $script:Gravacao
        description = 'Encerra ou reabre um planejamento. Encerrar com visitas pendentes é recusado, salvo cancelar_pendentes=true (cancela as pendentes com o motivo "planejamento encerrado"). Envie a versão lida.'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; required = @('id', 'versao', 'situacao'); properties = [ordered]@{
            id = [ordered]@{ type = 'integer' }
            versao = [ordered]@{ type = 'integer' }
            situacao = [ordered]@{ type = 'string'; enum = @('ABERTO', 'ENCERRADO') }
            cancelar_pendentes = [ordered]@{ type = 'boolean' }
            autor = $script:PropAutor } }
    },
    [ordered]@{
        name = 'listar_candidatos'; title = 'Candidatos do planejamento'; handler = 'Invoke-ListarCandidatos'; annotations = $script:SoLeitura
        description = 'Atendimentos em aberto a até raio_km da cidade-base do planejamento, na ordem de visita (qualificação, dias parado, km). Cidades não reconhecidas vêm à parte. sem_atendimento: empresas ativas do raio sem atendimento em aberto (pré-clientes inclusive), até 200, com quantidade_sem_atendimento; para visitar, abra o atendimento antes (gerar_visitas só aceita atendimento).'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; required = @('id_planejamento'); properties = [ordered]@{
            id_planejamento = [ordered]@{ type = 'integer' } } }
    },
    [ordered]@{
        name = 'gerar_visitas'; title = 'Gerar visitas do planejamento'; handler = 'Invoke-GerarVisitas'; annotations = $script:Gravacao
        description = 'Gera uma visita PENDENTE por atendimento escolhido, no planejamento. simular=true (padrão) só confere e não grava. Com simular=false, se um atendimento for recusado (encerrado ou já com visita pendente), nada é gravado. Até 100 atendimentos.'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; required = @('id_planejamento', 'codigos'); properties = [ordered]@{
            id_planejamento = [ordered]@{ type = 'integer' }
            codigos = [ordered]@{ type = 'array'; minItems = 1; maxItems = 100; items = [ordered]@{ type = 'string'; pattern = '^AT-\d{4}-\d{4}$' } }
            simular = [ordered]@{ type = 'boolean' }
            autor = $script:PropAutor } }
    },
    [ordered]@{
        name = 'lancar_datas'; title = 'Lançar datas das visitas'; handler = 'Invoke-LancarDatas'; annotations = $script:Gravacao
        description = 'Lança data e hora agendadas em lote (até 100 visitas). simular=true (padrão) só confere. Com simular=false, se um item for recusado, nada é gravado. Cada item traz id e versao lidas.'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; required = @('itens'); properties = [ordered]@{
            simular = [ordered]@{ type = 'boolean' }
            itens = [ordered]@{ type = 'array'; minItems = 1; maxItems = 100; items = [ordered]@{ type = 'object'; additionalProperties = $false; required = @('id', 'versao', 'data_agendada'); properties = [ordered]@{
                id = [ordered]@{ type = 'integer' }
                versao = [ordered]@{ type = 'integer' }
                data_agendada = [ordered]@{ type = 'string'; description = 'AAAA-MM-DD' }
                hora_agendada = [ordered]@{ type = 'string'; maxLength = 20 } } } }
            autor = $script:PropAutor } }
    },
    [ordered]@{
        name = 'lancar_resultados'; title = 'Lançar resultados das visitas'; handler = 'Invoke-LancarResultados'; annotations = $script:Gravacao
        description = 'Lança o resultado em lote (até 100 visitas): REALIZADA exige data_realizada e relato (grava o histórico do atendimento); CANCELADA exige motivo_cancelamento. simular=true (padrão) só confere.'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; required = @('itens'); properties = [ordered]@{
            simular = [ordered]@{ type = 'boolean' }
            itens = [ordered]@{ type = 'array'; minItems = 1; maxItems = 100; items = [ordered]@{ type = 'object'; additionalProperties = $false; required = @('id', 'versao', 'situacao'); properties = [ordered]@{
                id = [ordered]@{ type = 'integer' }
                versao = [ordered]@{ type = 'integer' }
                situacao = [ordered]@{ type = 'string'; enum = @('REALIZADA', 'CANCELADA') }
                data_realizada = [ordered]@{ type = 'string'; description = 'AAAA-MM-DD ou AAAA-MM-DD HH:mm' }
                resultado = [ordered]@{ type = 'string' }
                relato = [ordered]@{ type = 'string' }
                motivo_cancelamento = [ordered]@{ type = 'string' } } } }
            autor = $script:PropAutor } }
    },
    [ordered]@{
        name = 'listar_planejamentos'; title = 'Listar planejamentos'; handler = 'Invoke-ListarPlanejamentos'; annotations = $script:SoLeitura
        description = 'Planejamentos por região, mais recentes primeiro, com a cidade-base e o período. Filtre por situação.'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; properties = [ordered]@{
            situacao = [ordered]@{ type = 'string'; enum = @('ABERTO', 'ENCERRADO') }
            limite = [ordered]@{ type = 'integer'; maximum = 200 } } }
    },
    [ordered]@{
        name = 'obter_planejamento'; title = 'Ficha do planejamento'; handler = 'Invoke-ObterPlanejamento'; annotations = $script:SoLeitura
        description = 'Um planejamento pelo id ou pelo código PL-AAAA-NNN, com o resumo das visitas (pendentes, realizadas, canceladas). Devolve versao.'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; properties = [ordered]@{
            id = [ordered]@{ type = 'integer' }
            codigo = [ordered]@{ type = 'string'; pattern = '^PL-\d{4}-\d{3}$' } } }
    },
    [ordered]@{
        name = 'buscar_municipios'; title = 'Buscar municípios'; handler = 'Invoke-BuscarMunicipios'; annotations = $script:SoLeitura
        description = 'Municípios pelo nome (sem acento e sem diferença de maiúsculas), com UF opcional. Devolve o código IBGE para criar um planejamento.'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; required = @('texto'); properties = [ordered]@{
            texto = [ordered]@{ type = 'string'; minLength = 2; maxLength = 80 }
            uf = [ordered]@{ type = 'string'; pattern = '^[A-Za-z]{2}$' } } }
    },
    [ordered]@{
        name = 'resolver_cidade'; title = 'Resolver cidade'; handler = 'Invoke-ResolverCidade'; annotations = $script:SoLeitura
        description = 'Acha o município pelo nome e pela UF. Se achar um só, devolve o código IBGE; se não, devolve as opções.'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; required = @('cidade', 'uf'); properties = [ordered]@{
            cidade = [ordered]@{ type = 'string'; maxLength = 80 }
            uf = [ordered]@{ type = 'string'; pattern = '^[A-Za-z]{2}$' } } }
    },
    [ordered]@{
        name = 'exportar_roteiro'; title = 'Exportar roteiro do planejamento'; handler = 'Invoke-ExportarRoteiro'; annotations = $script:Gravacao
        description = 'Grava o CSV do roteiro do planejamento numa subpasta nova de SISTEMA\DADOS\execucao\extratos, com manifesto (UTF-8 com BOM, ponto e vírgula): a linha BASE e uma linha por visita pendente. Devolve só o caminho e o número de visitas. Exige o esquema que o conector suporta (crm_status).'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; properties = [ordered]@{
            id_planejamento = [ordered]@{ type = 'integer' }
            codigo = [ordered]@{ type = 'string'; pattern = '^PL-\d{4}-\d{3}$' }
            autor = $script:PropAutor } }
    },
    [ordered]@{
        name = 'listar_opcoes'; title = 'Opções das listas suspensas'; handler = 'Invoke-ListarOpcoes'; annotations = $script:SoLeitura
        description = 'Valores aceitos em campos de lista (Etapa, Responsavel, Familia, Origem, Prioridade, Motivo, Segmento, UF, Categoria). Sem "tipo", devolve os tipos existentes. Consulte antes de gravar um campo de lista.'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; properties = [ordered]@{
            tipo = [ordered]@{ type = 'string'; description = 'Nome da lista, ex.: Etapa' }
            detalhar = [ordered]@{ type = 'boolean'; description = 'Traz também os inativos, o uso e, na Etapa, o papel.' } } }
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
            situacao = [ordered]@{ type = 'string'; enum = @(('A' + [char]0x00E7 + [char]0x00E3 + 'o atrasada'), 'Retomada a avaliar',
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
                       'Também traz retomadas_a_avaliar (Perdido/Descartado com retomar_em até o fim do período; só na primeira página). ' +
                       'Se passar do limite de tamanho, devolve truncado=true e proximo_cursor: repita com cursor=proximo_cursor (leitura dinâmica, ver limitacao). a_partir_de continua aceito.')
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; properties = [ordered]@{
            data_inicio = [ordered]@{ type = 'string'; pattern = '^\d{4}-\d{2}-\d{2}$'; description = 'AAAA-MM-DD, inclusive. Exige data_fim. Anterior a hoje vira hoje (ajustado).' }
            data_fim = [ordered]@{ type = 'string'; pattern = '^\d{4}-\d{2}-\d{2}$'; description = 'AAAA-MM-DD, inclusive. Exige data_inicio. Anterior a hoje é recusada. Período de até 92 dias.' }
            semana = [ordered]@{ type = 'integer'; minimum = 1; maximum = 53; description = 'Semana ISO 8601 (1 a 53). Exige ano. Semana encerrada é recusada; a corrente começa hoje.' }
            ano = [ordered]@{ type = 'integer'; minimum = 2000; maximum = 2100; description = 'Ano da semana ISO. Exige semana.' }
            responsavel = [ordered]@{ type = 'string'; maxLength = 60; description = 'Filtra um responsável (ex.: "Paulo").' }
            incluir_vencidas = [ordered]@{ type = 'boolean'; default = $true; description = 'Inclui os abertos com próxima ação anterior a hoje, com vencida=true.' }
            completo = [ordered]@{ type = 'boolean'; default = $false; description = 'true devolve todos os campos do atendimento, inclusive os técnicos.' }
            a_partir_de = [ordered]@{ type = 'integer'; minimum = 0; default = 0; description = 'Deslocamento (compatibilidade): prefira cursor.' }
            cursor = [ordered]@{ type = 'string'; maxLength = 2000; description = 'proximo_cursor da chamada anterior, com os mesmos filtros.' } } }
    },
    [ordered]@{
        name = 'exportar_base'; title = 'Extrato da base em CSV'; handler = 'Invoke-ExportarBase'
        annotations = [ordered]@{ readOnlyHint = $false; destructiveHint = $false; idempotentHint = $false; openWorldHint = $false }
        description = ('Grava a foto das tabelas do CRM em CSV (crm-<tabela>.csv + manifesto crm-extrato.json) numa subpasta nova de SISTEMA\DADOS\execucao\extratos, para análise por script. ' +
                       'Não altera o banco. Devolve o caminho da pasta e o envelope (arquivos, linhas, colunas, SHA-256) - nunca os dados. ' +
                       'Manifesto com solicitante (autor), estação, finalidade, instante e corte; consistencia = nao_comprovada (cada tabela é lida numa consulta). ' +
                       'Atendimentos saem com codigo_cliente, empresa, codigo_contato e os calculados situacao, quadro, dias_parado e ciclo. ' +
                       'CSV UTF-8, separador vírgula, campos entre aspas. Retenção: mês corrente e anterior (descartar_extratos).')
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; properties = [ordered]@{
            tabelas = [ordered]@{ type = 'array'; items = [ordered]@{ type = 'string' }; maxItems = 9; description = 'clientes, contatos, oportunidades, log_alteracoes, listas, metas, planejamentos, visitas, municipios. Padrão: clientes, contatos, oportunidades.' }
            finalidade = [ordered]@{ type = 'string'; maxLength = 200; description = 'Para que o extrato será usado (vai para o manifesto). Padrão: análise.' }
            autor = $script:PropAutor } }
    },
    [ordered]@{
        name = 'descartar_extratos'; title = 'Descartar extratos vencidos'; handler = 'Invoke-DescartarExtratos'
        annotations = [ordered]@{ readOnlyHint = $false; destructiveHint = $true; idempotentHint = $true; openWorldHint = $false }
        description = ('Apaga as subpastas de SISTEMA\DADOS\execucao\extratos criadas antes do mês anterior (seção 5.3), só as que têm manifesto válido e não estão em processamento. ' +
                       'simular = true (padrão) só lista o que seria apagado. Nada fora da raiz de extratos é tocado; o descarte fica registrado em descartes.log.')
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; properties = [ordered]@{
            simular = [ordered]@{ type = 'boolean'; default = $true; description = 'true: só lista. false: apaga.' }
            autor = $script:PropAutor } }
    },
    [ordered]@{
        name = 'atualizar_atendimento'; title = 'Atualizar atendimento'; handler = 'Invoke-AtualizarAtendimento'; annotations = $script:Gravacao
        description = ('Altera campos de um atendimento. Leia antes com obter_atendimento e passe a "versao" lida: se alguém alterou depois, a gravação é recusada em vez de sobrescrever. ' +
                       'Envie só os campos que mudam. Campos de lista precisam de valor existente (listar_opcoes). Datas em AAAA-MM-DD. ' +
                       'Etapa que encerra o atendimento (aba Listas, Encerra = Sim) exige motivo_desfecho e dt_desfecho.')
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
            codigo_atendimento_anterior = [ordered]@{ type = @('string', 'null'); pattern = '^AT-\d{4}-\d{4}$'; description = 'Atendimento de origem (retomada ou duplicação), da mesma empresa; null remove. Não pode fechar ciclo.' }
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
        description = 'Grava o resumo de contexto da empresa (ctx_*): síntese de uma linha, família aderente e caminho do CONTEXTO-GERAL.md. Campos exclusivos da IA; não conflitam com o vendedor. Exige a ctx_versao lida em obter_cliente: se outra sessão gravou antes, volta conflito com o contexto atual.'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; required = @('ctx_versao'); properties = [ordered]@{
            codigo_cliente = [ordered]@{ type = 'integer' }; id_cliente = [ordered]@{ type = 'integer' }
            ctx_versao = [ordered]@{ type = 'integer'; minimum = 0; description = 'ctx_versao lida em obter_cliente (DEV-17)' }
            ctx_resumo = [ordered]@{ type = @('string', 'null'); maxLength = 255 }
            ctx_familia = [ordered]@{ type = @('string', 'null'); maxLength = 255 }
            ctx_arquivo = [ordered]@{ type = @('string', 'null'); maxLength = 255; description = 'Caminho relativo do CONTEXTO-GERAL.md' }
            autor = $script:PropAutor } }
    },
    [ordered]@{
        name = 'atualizar_contexto_atendimento'; title = 'Atualizar contexto do atendimento'; handler = 'Invoke-AtualizarContextoAtendimento'; annotations = $script:Gravacao
        description = 'Grava o resumo de contexto de um atendimento (ctx_resumo e caminho do CONTEXTO-ATENDIMENTO.md). Campos exclusivos da IA. Exige a ctx_versao lida em obter_atendimento: se outra sessão gravou antes, volta conflito com o contexto atual.'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; required = @('codigo', 'ctx_versao'); properties = [ordered]@{
            ctx_versao = [ordered]@{ type = 'integer'; minimum = 0; description = 'ctx_versao lida em obter_atendimento (DEV-17)' }
            codigo = [ordered]@{ type = 'string'; pattern = '^AT-\d{4}-\d{4}$' }
            ctx_resumo = [ordered]@{ type = @('string', 'null'); maxLength = 255 }
            ctx_arquivo = [ordered]@{ type = @('string', 'null'); maxLength = 255 }
            autor = $script:PropAutor } }
    }
)
