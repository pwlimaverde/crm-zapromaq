<#
  005 - normaliza os textos já gravados pela seção 5.1 do plano de 19/10/2026 (ciclo 009, DEV-10).

  Só DADOS (não muda versao_esquema): numa transação, cada texto comercial existente passa pela MESMA regra
  da gravação nova (ConvertTo-Normalizado, IA\mcp\lib\Regras.ps1 - espelho de modValidacao.Normalizar):
  maiúsculas, Trim e espaço repetido colapsado, quebras de linha preservadas nas observações, próxima ação e
  contexto, e URL/caminho/e-mail dentro do texto como estavam. E-mail com mais de um endereço fica com o
  primeiro; os demais vão para as observações (E-MAIL ADICIONAL). Listas, CNPJ, telefone e ctx_arquivo não mudam.

  Cada campo alterado sobe a versão do registro e entra no log com origem MIGRACAO (valor antigo e novo):
  nada se perde. A saída traz a PRÉVIA: quantos campos mudam por tabela e campo e até 3 exemplos (no ensaio,
  o relatório de execucao\ensaio mostra isso antes da publicação; tem dado real, não vai para o repositório).
  Idempotente: rodar de novo não muda nada (a regra aplicada duas vezes dá o mesmo texto).
#>
@{
    Id = '005'
    Descricao = 'normalização dos textos já gravados (seção 5.1: maiúsculas, um e-mail por campo)'
    EsquemaDe = $null
    EsquemaPara = $null
    Aplicar = {
        param($cn, $tx, $log, $pasta)
        $raizBase = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $pasta)))
        . (Join-Path $raizBase 'IA\mcp\lib\Regras.ps1')

        $alvos = [ordered]@{
            clientes      = @('empresa', 'cidade', 'uf', 'email', 'observacoes', 'ctx_resumo', 'ctx_familia')
            contatos      = @('nome', 'cargo', 'email', 'observacoes')
            oportunidades = @('orcamento', 'maquina', 'tipo_venda', 'prox_acao', 'observacoes', 'ctx_resumo')
        }
        $agora = Get-Date
        $total = 0
        foreach ($tabela in $alvos.Keys) {
            $campos = $alvos[$tabela]
            $c = New-ComandoBanco $cn ('SELECT id, ' + ($campos -join ', ') + ' FROM ' + $tabela + ' ORDER BY id') @() $tx
            $rd = $c.ExecuteReader()
            $linhas = New-Object System.Collections.ArrayList
            while ($rd.Read()) {
                $l = [ordered]@{ id = [int]$rd.GetValue(0) }
                for ($i = 0; $i -lt $campos.Count; $i++) { $v = $rd.GetValue($i + 1); $l[$campos[$i]] = $(if ($v -is [DBNull]) { $null } else { [string]$v }) }
                [void]$linhas.Add($l)
            }
            $rd.Close(); $c.Dispose()
            $porCampo = [ordered]@{}; $exemplos = [ordered]@{}
            foreach ($l in $linhas) {
                $mud = [ordered]@{}
                foreach ($campo in $campos) {
                    $antigo = $l[$campo]
                    if ($null -eq $antigo) { continue }
                    $script:EmailsAdicionais = ''
                    $novo = ConvertTo-Normalizado $campo $antigo
                    if ($campo -eq 'email' -and $script:EmailsAdicionais) {
                        $baseObs = if ($mud.Contains('observacoes')) { $mud['observacoes'].novo } else { $l['observacoes'] }
                        $obs = Add-EmailsAdicionais $baseObs $script:EmailsAdicionais
                        if ($obs -cne $l['observacoes']) { $mud['observacoes'] = @{ antigo = $l['observacoes']; novo = $obs } }
                    }
                    if ($campo -eq 'observacoes' -and $mud.Contains('observacoes')) { continue }
                    if ([string]$novo -cne [string]$antigo) { $mud[$campo] = @{ antigo = $antigo; novo = $novo } }
                }
                # observação ganhou e-mail e também precisa da normalização: aplica a regra sobre o resultado
                if ($mud.Contains('observacoes')) { $mud['observacoes'].novo = ConvertTo-Normalizado 'observacoes' $mud['observacoes'].novo }
                if ($mud.Count -eq 0) { continue }
                $sets = @(); $par = @()
                foreach ($k in $mud.Keys) { $sets += ($k + '=?'); $par += (New-Parametro $(if ($k -in @('observacoes', 'prox_acao')) { 'memo' } else { 'texto' }) $mud[$k].novo) }
                $par += (New-Parametro 'data' $agora); $par += (New-Parametro 'texto' 'MIGRACAO 005'); $par += (New-Parametro 'inteiro' $l.id)
                [void](Invoke-Comando $cn ('UPDATE ' + $tabela + ' SET ' + ($sets -join ', ') + ', versao=versao+1, alterado_em=?, alterado_por=? WHERE id=?') $par $tx)
                foreach ($k in $mud.Keys) {
                    # DEV-06 parte B (esquema 1.2, que a 004 já criou): estação, usuário técnico e versão em toda linha nova
                    [void](Invoke-Comando $cn ('INSERT INTO log_alteracoes (tabela, id_registro, campo, valor_antigo, valor_novo, acao, usuario, quando, origem,' +
                                               ' usuario_tecnico, estacao, versao_cliente) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)') @(
                        (New-Parametro 'texto' $tabela), (New-Parametro 'inteiro' $l.id), (New-Parametro 'texto' $k),
                        (New-Parametro 'memo' $mud[$k].antigo), (New-Parametro 'memo' $mud[$k].novo), (New-Parametro 'texto' 'NORMALIZACAO'),
                        (New-Parametro 'texto' 'MIGRACAO 005'), (New-Parametro 'data' $agora), (New-Parametro 'texto' 'MIGRACAO'),
                        (New-Parametro 'texto' ([Environment]::UserName)), (New-Parametro 'texto' ([Environment]::MachineName)),
                        (New-Parametro 'texto' 'MIGRACAO 005')) $tx)
                    if (-not $porCampo.Contains($k)) { $porCampo[$k] = 0; $exemplos[$k] = New-Object System.Collections.ArrayList }
                    $porCampo[$k]++
                    if ($exemplos[$k].Count -lt 3) { [void]$exemplos[$k].Add(([string]$mud[$k].antigo).Split("`n")[0].Trim() + ' -> ' + ([string]$mud[$k].novo).Split("`n")[0].Trim()) }
                    $total++
                }
            }
            foreach ($k in $porCampo.Keys) {
                $log.L('  ' + $tabela + '.' + $k + ': ' + $porCampo[$k] + ' valor(es) normalizado(s)')
                foreach ($e in $exemplos[$k]) { $log.L('      ex.: ' + $e) }
            }
        }
        $log.L('  normalização: ' + $total + ' campo(s) alterado(s), cada um no log com origem MIGRACAO')
    }
}
