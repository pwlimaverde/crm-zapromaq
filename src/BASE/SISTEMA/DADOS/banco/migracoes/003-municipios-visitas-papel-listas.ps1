<#
  003 - esquema 1.0 -> 1.1: municípios, planejamentos, visitas, papel das listas.

  Migração ESTRUTURAL (EsquemaDe 1.0, EsquemaPara 1.1). Quem a executa é a montagem, sobre uma
  CÓPIA do banco (SISTEMA\DADOS\execucao\publicacao); o banco de produção só é trocado no fim.

  Estrutura (spec 003, Contratos > Tabelas): cada tabela e coluna só é criada se faltar, então a
  migração pode rodar de novo sem erro e sem mudar nada (idempotente).

  Dados, numa transação só (tudo ou nada dentro da cópia):
    1. municípios: carga do municipios.csv (SHA-256 conferido contra a constante abaixo).
    2. chave de cada linha de listas (Get-ChaveLista).
    3. valor usado sem linha na lista (pela chave) entra INATIVO, com o texto mais usado (C18).
    4. variantes da mesma chave: fica a mais usada; empate -> ativa; empate -> menor id (C17).
       Os registros das outras variantes recebem o texto da que ficou (versao +1, log UNIFICACAO);
       a linha duplicada sai da lista (log UNIFICACAO, com o valor antigo).
    5. etapa especial ausente entra inativa; papel inicial de toda etapa sem papel.
  Índice único (tipo, chave) é criado no fim, depois da transação (DDL não é transacional).

  Contagem de uso é feita aqui, no PowerShell, com comparação EXATA de texto: o GROUP BY e o WHERE
  do Access não diferenciam maiúsculas (PEDIDO FECHADO e Pedido Fechado seriam o mesmo valor).
#>
@{
    Id = '003'
    Descricao = 'municípios, planejamentos, visitas e papel das listas (esquema 1.1)'
    EsquemaDe = '1.0'
    EsquemaPara = '1.1'
    Aplicar = {
        param($cn, $tx, $log, $pasta)

        # ---------------------------------------------------------------- helpers
        function Tem-Tabela([string]$nome) {
            try { [void](Invoke-Escalar $cn ('SELECT COUNT(*) FROM ' + $nome) @() $null); return $true } catch { return $false }
        }
        function Tem-Coluna([string]$tabela, [string]$coluna) {
            try { [void](Invoke-Escalar $cn ('SELECT TOP 1 ' + $coluna + ' FROM ' + $tabela) @() $null); return $true } catch { return $false }
        }
        function Tem-Indice([string]$nome) {
            $esquema = $cn.GetSchema('Indexes')
            foreach ($r in $esquema.Rows) { if ([string]$r['INDEX_NAME'] -eq $nome) { return $true } }
            return $false
        }
        function Exec([string]$sql) { [void](Invoke-Comando $cn $sql @() $null) }
        # leitura dentro da transação; devolve objetos com as colunas pelo nome (nulo = $null)
        function Ler([string]$sql, [array]$params = @()) {
            $c = New-ComandoBanco $cn $sql $params $t
            $r = $c.ExecuteReader()
            $saida = New-Object System.Collections.ArrayList
            while ($r.Read()) {
                $linha = [ordered]@{}
                for ($i = 0; $i -lt $r.FieldCount; $i++) {
                    $v = $r.GetValue($i); if ($v -is [DBNull]) { $v = $null }
                    $linha[$r.GetName($i).ToLowerInvariant()] = $v
                }
                [void]$saida.Add([pscustomobject]$linha)
            }
            $r.Close(); $c.Dispose()
            return , $saida
        }
        function Log-Alteracao([string]$tabela, [int]$id, [string]$campo, [string]$antigo, [string]$novo, [string]$acao) {
            [void](Invoke-Comando $cn ('INSERT INTO log_alteracoes (tabela, id_registro, campo, valor_antigo, valor_novo, acao, usuario, quando, origem)' +
                ' VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)') @(
                (New-Parametro 'texto' $tabela), (New-Parametro 'inteiro' $id), (New-Parametro 'texto' $campo),
                (New-Parametro 'memo' $antigo), (New-Parametro 'memo' $novo), (New-Parametro 'texto' $acao),
                (New-Parametro 'texto' 'MIGRACAO 003'), (New-Parametro 'data' (Get-Date)), (New-Parametro 'texto' 'MIGRACAO')) $t)
        }

        # ---------------------------------------------------------------- estrutura (cada parte só se faltar)
        if (-not (Tem-Tabela 'municipios')) {
            Exec ('CREATE TABLE municipios (cod_ibge LONG PRIMARY KEY, nome TEXT(80) NOT NULL, nome_busca TEXT(80) NOT NULL,' +
                  ' uf TEXT(2) NOT NULL, lat DOUBLE NOT NULL, lon DOUBLE NOT NULL, fonte_carga TEXT(60))')
            Exec 'CREATE INDEX ix_mun_uf_busca ON municipios (uf, nome_busca)'
        }
        if (-not (Tem-Tabela 'planejamentos')) {
            Exec ('CREATE TABLE planejamentos (id AUTOINCREMENT PRIMARY KEY, codigo TEXT(12) NOT NULL, controle LONG NOT NULL,' +
                  ' descricao TEXT(120) NOT NULL, cod_municipio_base LONG NOT NULL, raio_km INTEGER NOT NULL, data_inicio DATETIME,' +
                  ' data_fim DATETIME, responsavel TEXT(60), situacao TEXT(20) NOT NULL, observacoes MEMO, versao LONG,' +
                  ' criado_em DATETIME, alterado_em DATETIME, alterado_por TEXT(50))')
            Exec 'CREATE UNIQUE INDEX ix_pln_codigo ON planejamentos (codigo)'
            Exec 'CREATE UNIQUE INDEX ix_pln_controle ON planejamentos (controle)'
            Exec 'ALTER TABLE planejamentos ADD CONSTRAINT fk_pln_mun FOREIGN KEY (cod_municipio_base) REFERENCES municipios (cod_ibge)'
        }
        if (-not (Tem-Tabela 'visitas')) {
            Exec ('CREATE TABLE visitas (id AUTOINCREMENT PRIMARY KEY, id_oportunidade LONG NOT NULL, id_planejamento LONG,' +
                  ' situacao TEXT(20) NOT NULL, data_agendada DATETIME, hora_agendada TEXT(20), responsavel TEXT(60), tipo TEXT(40),' +
                  ' observacoes MEMO, data_realizada DATETIME, resultado TEXT(40), relato MEMO, motivo_cancelamento TEXT(120),' +
                  ' versao LONG, criado_em DATETIME, alterado_em DATETIME, alterado_por TEXT(50))')
            Exec 'CREATE INDEX ix_vis_opo_sit ON visitas (id_oportunidade, situacao)'
            Exec 'CREATE INDEX ix_vis_pln ON visitas (id_planejamento)'
            Exec 'CREATE INDEX ix_vis_data ON visitas (data_agendada)'
            Exec 'ALTER TABLE visitas ADD CONSTRAINT fk_vis_opo FOREIGN KEY (id_oportunidade) REFERENCES oportunidades (id)'
            Exec 'ALTER TABLE visitas ADD CONSTRAINT fk_vis_pln FOREIGN KEY (id_planejamento) REFERENCES planejamentos (id)'
        }
        $colunasListas = @(
            'ALTER TABLE listas ADD COLUMN chave TEXT(120)',
            'ALTER TABLE listas ADD COLUMN encerra YESNO',
            'ALTER TABLE listas ADD COLUMN resultado TEXT(10)',
            'ALTER TABLE listas ADD COLUMN quadro TEXT(12)',
            'ALTER TABLE listas ADD COLUMN exige_familia YESNO')
        foreach ($sql in $colunasListas) {
            if (-not (Tem-Coluna 'listas' (($sql -split ' ')[5]))) { Exec $sql }
        }
        $log.L('  estrutura 1.1 conferida (tabelas e colunas criadas só se faltavam)')

        # ---------------------------------------------------------------- dados (uma transação)
        $t = $cn.BeginTransaction()
        try {
            # 1. municípios -------------------------------------------------
            $shaCsv = 'bcf43574ebdd7c240ca21c02701101fff8599f60bbd4a187c443ff2b3a288515'
            $csv = Join-Path $pasta '..\referencia\municipios.csv'
            $bytes = [IO.File]::ReadAllBytes($csv)
            $sha = [BitConverter]::ToString([Security.Cryptography.SHA256]::Create().ComputeHash($bytes)).Replace('-', '').ToLowerInvariant()
            if ($sha -ne $shaCsv) { throw 'municipios.csv diferente do esperado pela migração 003 (SHA-256)' }
            $existentes = @{}
            foreach ($r in (Ler 'SELECT cod_ibge FROM municipios')) { $existentes[[int]$r.cod_ibge] = $true }
            $novos = 0
            foreach ($l in [IO.File]::ReadAllLines($csv, [Text.Encoding]::UTF8) | Select-Object -Skip 1) {
                $c = $l.Split(';')
                $cod = [int]$c[0]
                if ($existentes.ContainsKey($cod)) { continue }
                [void](Invoke-Comando $cn 'INSERT INTO municipios (cod_ibge, nome, nome_busca, uf, lat, lon, fonte_carga) VALUES (?, ?, ?, ?, ?, ?, ?)' @(
                    (New-Parametro 'inteiro' $cod), (New-Parametro 'texto' $c[1]), (New-Parametro 'texto' $c[2]), (New-Parametro 'texto' $c[3]),
                    (New-Parametro 'numero' ([double]::Parse($c[4], [Globalization.CultureInfo]::InvariantCulture))),
                    (New-Parametro 'numero' ([double]::Parse($c[5], [Globalization.CultureInfo]::InvariantCulture))),
                    (New-Parametro 'texto' ('IBGE/kelvins ' + $sha.Substring(0, 12)))) $t)
                $novos++
            }
            $log.L('  municípios: ' + $novos + ' carregado(s), ' + $existentes.Count + ' já existiam')

            # 2. chave de cada linha --------------------------------------
            foreach ($r in (Ler 'SELECT id, valor FROM listas WHERE chave IS NULL')) {
                [void](Invoke-Comando $cn 'UPDATE listas SET chave=? WHERE id=?' @((New-Parametro 'texto' (Get-ChaveLista $r.valor)), (New-Parametro 'inteiro' $r.id)) $t)
            }

            # 3. onde cada valor é usado (comparação exata) ---------------
            $colunasDaLista = Get-ColunasDaLista
            $usos = @{}
            foreach ($tipo in $colunasDaLista.Keys) {
                $usos[$tipo] = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.ArrayList]'
                foreach ($tc in $colunasDaLista[$tipo]) {
                    $partes = $tc.Split(':')
                    foreach ($r in (Ler ('SELECT id, ' + $partes[1] + ' AS v FROM ' + $partes[0] + ' WHERE ' + $partes[1] + ' IS NOT NULL'))) {
                        $v = [string]$r.v
                        if ($v -eq '') { continue }
                        if (-not $usos[$tipo].ContainsKey($v)) { $usos[$tipo][$v] = New-Object System.Collections.ArrayList }
                        [void]$usos[$tipo][$v].Add([pscustomobject]@{ tabela = $partes[0]; id = [int]$r.id; coluna = $partes[1] })
                    }
                }
            }
            function UsoDe([string]$tipo, [string]$valor) {
                if ($usos[$tipo].ContainsKey($valor)) { return $usos[$tipo][$valor].Count }
                return 0
            }

            # 4. valor usado sem linha na lista (pela chave): entra inativo (C18) -----
            $linhas = Ler 'SELECT id, tipo, valor, chave, ativo, ordem FROM listas'
            $existeNaLista = @{}
            $ordemMax = @{}
            foreach ($r in $linhas) {
                $existeNaLista[$r.tipo + '|' + $r.chave] = $true
                if (-not $ordemMax.ContainsKey($r.tipo) -or [int]$r.ordem -gt $ordemMax[$r.tipo]) { $ordemMax[$r.tipo] = [int]$r.ordem }
            }
            foreach ($tipo in @($usos.Keys)) {
                $porChave = @{}
                foreach ($valor in $usos[$tipo].Keys) {
                    $ch = Get-ChaveLista $valor
                    if ($existeNaLista.ContainsKey($tipo + '|' + $ch)) { continue }
                    if (-not $porChave.ContainsKey($ch)) { $porChave[$ch] = $valor; continue }
                    $atual = $porChave[$ch]
                    if ((UsoDe $tipo $valor) -gt (UsoDe $tipo $atual) -or ((UsoDe $tipo $valor) -eq (UsoDe $tipo $atual) -and [string]::CompareOrdinal($valor, $atual) -lt 0)) { $porChave[$ch] = $valor }
                }
                foreach ($ch in $porChave.Keys) {
                    $valor = $porChave[$ch]
                    $ordemMax[$tipo] = [int]$ordemMax[$tipo] + 1
                    [void](Invoke-Comando $cn 'INSERT INTO listas (tipo, valor, chave, ordem, ativo) VALUES (?, ?, ?, ?, ?)' @(
                        (New-Parametro 'texto' $tipo), (New-Parametro 'texto' $valor), (New-Parametro 'texto' $ch),
                        (New-Parametro 'inteiro' $ordemMax[$tipo]), (New-Parametro 'logico' $false)) $t)
                    $existeNaLista[$tipo + '|' + $ch] = $true
                    $log.L('  + ' + $tipo + ' "' + $valor + '" usado em ' + (UsoDe $tipo $valor) + ' registro(s) sem linha na lista: entrou INATIVO (C18)')
                }
            }

            # 5. variantes da mesma chave: fica a mais usada (C17) ------------
            $linhas = Ler 'SELECT id, tipo, valor, chave, ativo, ordem FROM listas'
            $grupos = @{}
            foreach ($r in $linhas) {
                $k = $r.tipo + '|' + $r.chave
                if (-not $grupos.ContainsKey($k)) { $grupos[$k] = New-Object System.Collections.ArrayList }
                [void]$grupos[$k].Add($r)
            }
            $unificadas = 0; $registrosAjustados = 0
            foreach ($k in $grupos.Keys) {
                $g = $grupos[$k]
                if ($g.Count -lt 2) { continue }
                $tipo = $g[0].tipo
                $vencedor = @($g | Sort-Object -Property @{ e = { UsoDe $tipo $_.valor }; Descending = $true },
                                                       @{ e = { if ([bool]$_.ativo) { 0 } else { 1 } } },
                                                       @{ e = { [int]$_.id } })[0]
                foreach ($outra in $g) {
                    if ([int]$outra.id -eq [int]$vencedor.id) { continue }
                    if ($outra.valor -cne $vencedor.valor -and $usos[$tipo].ContainsKey($outra.valor)) {
                        foreach ($reg in @($usos[$tipo][$outra.valor])) {
                            [void](Invoke-Comando $cn ("UPDATE " + $reg.tabela + " SET " + $reg.coluna + "=?, versao=versao+1, alterado_em=?, alterado_por=? WHERE id=?") @(
                                (New-Parametro 'texto' $vencedor.valor), (New-Parametro 'data' (Get-Date)), (New-Parametro 'texto' 'MIGRACAO 003'),
                                (New-Parametro 'inteiro' $reg.id)) $t)
                            Log-Alteracao $reg.tabela $reg.id $reg.coluna $outra.valor $vencedor.valor 'UNIFICACAO'
                            $registrosAjustados++
                        }
                    }
                    [void](Invoke-Comando $cn 'DELETE FROM listas WHERE id=?' @((New-Parametro 'inteiro' $outra.id)) $t)
                    Log-Alteracao 'listas' ([int]$outra.id) 'valor' $outra.valor $vencedor.valor 'UNIFICACAO'
                    $log.L('  unificado ' + $tipo + ': "' + $outra.valor + '" -> "' + $vencedor.valor + '"')
                    $unificadas++
                }
            }
            $log.L('  unificações: ' + $unificadas + ' linha(s) de lista; ' + $registrosAjustados + ' registro(s) ajustado(s)')

            # 6. etapas especiais ausentes entram inativas; papel inicial -----
            $especiais = @('Pedido Fechado', 'Perdido', 'Descartado', 'Contato Inicial', 'Sem Retorno', 'Retorno Agendado',
                           'Elaboração da Proposta', 'Negociação', 'Proposta em Análise', 'Proposta em Stand By')
            $linhas = Ler 'SELECT id, tipo, valor, chave, ativo, ordem FROM listas'
            $temEtapa = @{}
            foreach ($r in $linhas) { if ($r.tipo -eq 'Etapa') { $temEtapa[$r.chave] = $true } }
            foreach ($nome in $especiais) {
                $ch = Get-ChaveLista $nome
                if ($temEtapa.ContainsKey($ch)) { continue }
                $ordemEtapa = 0
                foreach ($r in $linhas) { if ($r.tipo -eq 'Etapa' -and [int]$r.ordem -gt $ordemEtapa) { $ordemEtapa = [int]$r.ordem } }
                [void](Invoke-Comando $cn 'INSERT INTO listas (tipo, valor, chave, ordem, ativo) VALUES (?, ?, ?, ?, ?)' @(
                    (New-Parametro 'texto' 'Etapa'), (New-Parametro 'texto' $nome), (New-Parametro 'texto' $ch),
                    (New-Parametro 'inteiro' ($ordemEtapa + 1)), (New-Parametro 'logico' $false)) $t)
                $temEtapa[$ch] = $true
                $log.L('  + Etapa especial ausente "' + $nome + '": entrou INATIVA com o papel dela')
            }
            # listas da visita (Fase 3): valores iniciais, só se a lista ainda estiver vazia (idempotente)
            $iniciais = [ordered]@{
                'TipoVisita'      = @('Visita', 'Retorno', 'Demonstração', 'Instalação')
                'ResultadoVisita' = @('Positivo', 'Neutro', 'Negativo', 'Sem contato')
            }
            foreach ($tipoLista in $iniciais.Keys) {
                $existentes = [int](Invoke-Escalar $cn 'SELECT COUNT(*) FROM listas WHERE tipo=?' @((New-Parametro 'texto' $tipoLista)) $t)
                if ($existentes -gt 0) { continue }
                $ordemVis = 0
                foreach ($v in $iniciais[$tipoLista]) {
                    $ordemVis++
                    [void](Invoke-Comando $cn 'INSERT INTO listas (tipo, valor, chave, ordem, ativo) VALUES (?, ?, ?, ?, ?)' @(
                        (New-Parametro 'texto' $tipoLista), (New-Parametro 'texto' $v), (New-Parametro 'texto' (Get-ChaveLista $v)),
                        (New-Parametro 'inteiro' $ordemVis), (New-Parametro 'logico' $true)) $t)
                }
                $log.L('  + lista ' + $tipoLista + ': ' + $ordemVis + ' valor(es) inicial(is)')
            }
            # 'quadro' é o marcador de "sem papel": a coluna YESNO criada por ALTER nasce Falso, não nula
            $papeis = 0
            foreach ($r in (Ler 'SELECT id, valor, chave FROM listas WHERE tipo=? AND quadro IS NULL' @((New-Parametro 'texto' 'Etapa')))) {
                $ch = [string]$r.chave
                $encerra = $false; $resultado = $null; $quadro = 'FUNIL'; $exige = $false
                if ($ch -eq 'PEDIDO FECHADO')                                                  { $encerra = $true; $resultado = 'GANHO'; $quadro = 'FUNIL' }
                elseif ($ch -eq 'PERDIDO')                                                     { $encerra = $true; $resultado = 'PERDA'; $quadro = 'FUNIL' }
                elseif ($ch -eq 'DESCARTADO')                                                  { $encerra = $true; $resultado = 'PERDA'; $quadro = 'PROSPECCAO' }
                elseif (@('CONTATO INICIAL', 'SEM RETORNO', 'RETORNO AGENDADO') -contains $ch) { $quadro = 'PROSPECCAO' }
                elseif (@('ELABORACAO DA PROPOSTA', 'NEGOCIACAO', 'PROPOSTA EM ANALISE', 'PROPOSTA EM STAND BY') -contains $ch) { $exige = $true }
                [void](Invoke-Comando $cn 'UPDATE listas SET encerra=?, resultado=?, quadro=?, exige_familia=? WHERE id=?' @(
                    (New-Parametro 'logico' $encerra), (New-Parametro 'texto' $resultado), (New-Parametro 'texto' $quadro),
                    (New-Parametro 'logico' $exige), (New-Parametro 'inteiro' $r.id)) $t)
                $papeis++
            }
            $log.L('  papel inicial gravado em ' + $papeis + ' etapa(s)')

            $t.Commit()
        } catch {
            $t.Rollback()
            throw
        }

        # ---------------------------------------------------------------- índice único (depois da transação)
        if (-not (Tem-Indice 'ix_lis_tipo_chave')) { Exec 'CREATE UNIQUE INDEX ix_lis_tipo_chave ON listas (tipo, chave)' }
        $log.L('  índice único (tipo, chave) conferido')
    }
}
