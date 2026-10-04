<#
  001 - cria a lista "Qualificacao" a partir dos valores já usados nos cadastros.

  Pendência registrada em 17/09/2026 (docs\legado\estado-migracao.md): a ficha de
  clientes e o filtro "Qualif." usam a lista Qualificacao, que nunca foi criada
  na tabela listas - o combo aparecia vazio. Os valores vêm dos PRÓPRIOS
  cadastros (DISTINCT de clientes.qualificacao), e não digitados aqui: assim a
  lista casa letra por letra, acento por acento, com o que está gravado, e o
  filtro encontra os registros. Idempotente: não repete valor que já exista.
  Só dados: não muda versao_esquema.
#>
@{
    Id = '001'
    Descricao = 'lista Qualificacao a partir dos valores em uso'
    EsquemaDe = $null
    EsquemaPara = $null
    Aplicar = {
        param($cn, $tx, $log)
        $ordem = [int](Invoke-Escalar $cn "SELECT COUNT(*) FROM listas WHERE tipo='Qualificacao'" @() $tx)
        $leitor = (New-ComandoBanco $cn ("SELECT DISTINCT qualificacao FROM clientes WHERE qualificacao IS NOT NULL" +
                   " AND qualificacao NOT IN (SELECT valor FROM listas WHERE tipo='Qualificacao') ORDER BY qualificacao") @() $tx).ExecuteReader()
        $novos = New-Object System.Collections.ArrayList
        while ($leitor.Read()) { $v = ([string]$leitor.GetValue(0)).Trim(); if ($v) { [void]$novos.Add($v) } }
        $leitor.Close()
        foreach ($v in $novos) {
            $ordem++
            [void](Invoke-Comando $cn 'INSERT INTO listas (tipo, valor, ordem, ativo) VALUES (?, ?, ?, True)' `
                   @((New-Parametro 'texto' 'Qualificacao'), (New-Parametro 'texto' $v), (New-Parametro 'inteiro' $ordem)) $tx)
            $log.L('  + Qualificacao: ' + $v)
        }
        if ($novos.Count -eq 0) { $log.L('  nenhum valor novo (lista já completa ou base sem qualificação)') }
    }
}
