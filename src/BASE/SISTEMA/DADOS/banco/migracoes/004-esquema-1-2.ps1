<#
  004 - esquema 1.1 -> 1.2 (ciclo 010): controle único, contexto versionado, autor técnico no log, operações.

  Migração ESTRUTURAL (EsquemaDe 1.1, EsquemaPara 1.2). Só a montagem aplica, numa CÓPIA do banco
  (execucao\publicacao ou execucao\ensaio), com backup validado antes; o banco de produção só é trocado no fim.

  1. INVENTÁRIO, antes de qualquer estrutura (DEV-16): controle repetido em contatos/oportunidades; atendimento
     anterior que não existe, de outra empresa, que aponta para si ou que forma ciclo. Qualquer achado PARA a
     migração com a lista (só códigos): corrigir dado real é operação, não desenvolvimento. Na montagem isso é o
     passo 7 com o banco de produção intacto; no ensaio, vai para o relatório.
  2. Estrutura (cada item só se faltar: idempotente):
     - índice único de controle: ix_cto_controle, ix_opo_controle (DEV-16);
     - ctx_versao LONG em clientes e oportunidades (DEV-17), começando em 0;
     - clientes.ctx_familia passa a TEXT(255) (DEV-10: nenhum campo trunca em silêncio);
     - log_alteracoes: usuario_tecnico, estacao, versao_cliente, autor_declarado (DEV-06 parte B);
     - tabela operacoes (id_operacao único) para o identificador de operação (DEV-13).
#>
@{
    Id = '004'
    Descricao = 'esquema 1.2: controle único, contexto versionado, autor técnico no log e operações'
    EsquemaDe = '1.1'
    EsquemaPara = '1.2'
    Aplicar = {
        param($cn, $tx, $log, $pasta)

        function Tem-Tabela([string]$nome) {
            try { [void](Invoke-Escalar $cn ('SELECT COUNT(*) FROM ' + $nome) @() $null); return $true } catch { return $false }
        }
        function Tem-Coluna([string]$tabela, [string]$coluna) {
            try { [void](Invoke-Escalar $cn ('SELECT TOP 1 ' + $coluna + ' FROM ' + $tabela) @() $null); return $true } catch { return $false }
        }
        function Tem-Indice([string]$nome) {
            foreach ($r in $cn.GetSchema('Indexes').Rows) { if ([string]$r['INDEX_NAME'] -eq $nome) { return $true } }
            return $false
        }
        function Exec([string]$sql) { [void](Invoke-Comando $cn $sql @() $null) }
        function Lista([string]$sql) {
            $c = New-ComandoBanco $cn $sql @() $null
            $r = $c.ExecuteReader(); $s = New-Object System.Collections.ArrayList
            while ($r.Read()) { $l = @(); for ($i = 0; $i -lt $r.FieldCount; $i++) { $v = $r.GetValue($i); $l += $(if ($v -is [DBNull]) { $null } else { $v }) }; [void]$s.Add($l) }
            $r.Close(); $c.Dispose()
            return , $s
        }

        # ---------------------------------------------------------------- 1. inventário
        $achados = New-Object System.Collections.ArrayList
        foreach ($t in @('contatos', 'oportunidades')) {
            foreach ($l in (Lista ('SELECT controle, COUNT(*) FROM ' + $t + ' GROUP BY controle HAVING COUNT(*) > 1'))) {
                [void]$achados.Add($t + ': controle ' + $l[0] + ' repetido em ' + $l[1] + ' registros')
            }
        }
        $ops = Lista 'SELECT id, codigo, id_cliente, id_atendimento_anterior FROM oportunidades'
        $porId = @{}; foreach ($o in $ops) { $porId[[int]$o[0]] = $o }
        foreach ($o in $ops) {
            if ($null -eq $o[3]) { continue }
            $ant = [int]$o[3]
            if ($ant -eq [int]$o[0]) { [void]$achados.Add($o[1] + ': atendimento anterior aponta para ele mesmo'); continue }
            if (-not $porId.ContainsKey($ant)) { [void]$achados.Add($o[1] + ': atendimento anterior (id ' + $ant + ') não existe'); continue }
            if ([int]$porId[$ant][2] -ne [int]$o[2]) { [void]$achados.Add($o[1] + ': atendimento anterior ' + $porId[$ant][1] + ' é de outra empresa') }
            # ciclo: seguir a corrente de anteriores
            $vistos = @{ ([int]$o[0]) = $true }; $p = $ant; $passos = 0
            while ($null -ne $p -and $porId.ContainsKey([int]$p) -and $passos -lt 10000) {
                if ($vistos.ContainsKey([int]$p)) { [void]$achados.Add($o[1] + ': atendimentos anteriores formam um ciclo'); break }
                $vistos[[int]$p] = $true; $p = $porId[[int]$p][3]; $passos++
            }
        }
        if ($achados.Count -gt 0) {
            $log.L('  INVENTÁRIO DA 004: ' + $achados.Count + ' achado(s) - a migração NÃO foi aplicada:')
            foreach ($a in $achados) { $log.L('    - ' + $a) }
            throw ('inventário da 004 com ' + $achados.Count + ' achado(s) (lista acima). Corrija os cadastros e rode de novo.')
        }
        $log.L('  inventário: nenhum controle repetido nem atendimento anterior inconsistente')

        # ---------------------------------------------------------------- 2. estrutura
        if (-not (Tem-Indice 'ix_cto_controle')) { Exec 'CREATE UNIQUE INDEX ix_cto_controle ON contatos (controle)' }
        if (-not (Tem-Indice 'ix_opo_controle')) { Exec 'CREATE UNIQUE INDEX ix_opo_controle ON oportunidades (controle)' }
        foreach ($t in @('clientes', 'oportunidades')) {
            if (-not (Tem-Coluna $t 'ctx_versao')) { Exec ('ALTER TABLE ' + $t + ' ADD COLUMN ctx_versao LONG') }
            Exec ('UPDATE ' + $t + ' SET ctx_versao = 0 WHERE ctx_versao IS NULL')
        }
        Exec 'ALTER TABLE clientes ALTER COLUMN ctx_familia TEXT(255)'
        foreach ($c in @('usuario_tecnico TEXT(50)', 'estacao TEXT(50)', 'versao_cliente TEXT(20)', 'autor_declarado TEXT(60)')) {
            if (-not (Tem-Coluna 'log_alteracoes' ($c.Split(' ')[0]))) { Exec ('ALTER TABLE log_alteracoes ADD COLUMN ' + $c) }
        }
        if (-not (Tem-Tabela 'operacoes')) {
            Exec ('CREATE TABLE operacoes (id AUTOINCREMENT PRIMARY KEY, id_operacao TEXT(120) NOT NULL, tabela TEXT(30), id_registro LONG,' +
                  ' codigo TEXT(20), versao LONG, hash_conteudo TEXT(64), criado_em DATETIME, origem TEXT(20), estacao TEXT(50))')
        }
        if (-not (Tem-Indice 'ix_ope_id')) { Exec 'CREATE UNIQUE INDEX ix_ope_id ON operacoes (id_operacao)' }
        if (-not (Tem-Indice 'ix_ope_criado')) { Exec 'CREATE INDEX ix_ope_criado ON operacoes (criado_em)' }
        $log.L('  estrutura 1.2: controle único, ctx_versao, ctx_familia 255, autor técnico no log, tabela operacoes')
    }
}
