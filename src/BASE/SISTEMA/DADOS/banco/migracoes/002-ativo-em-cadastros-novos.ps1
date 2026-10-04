<#
  002 - reativa cadastros que nasceram INATIVOS por defeito do front legado.

  O defeito: a coluna ativo é Sim/Não; no Access esse tipo não aceita nulo e o
  padrão é FALSO. O Inserir do front 1.4 não preenchia ativo, então todo cliente
  e contato cadastrado pela ficha (e pelo cadastro em lote) depois da migração de
  16/09/2026 nasceu com ativo = Falso: a ficha mostrava [INATIVO] e oferecia
  "Reativar". O front novo grava ativo = Verdadeiro (modCRM.InserirEm).

  Critério conservador - só reativa quem NUNCA foi desativado de propósito:
    clientes: ativo = Falso, criado a partir de 16/09/2026, sem nenhum log de
              DESATIVACAO para o registro;
    contatos: idem, E a empresa do contato está ativa (contato desativado em
              cascata junto com a empresa fica como está).
  Cada correção entra no log_alteracoes (acao CORRECAO, origem MIGRACAO).
  Só dados: não muda versao_esquema.
#>
@{
    Id = '002'
    Descricao = 'reativa cadastros novos gravados com ativo = Falso (defeito do front 1.4)'
    EsquemaDe = $null
    EsquemaPara = $null
    Aplicar = {
        param($cn, $tx, $log)
        $desde = [datetime]::new(2026, 9, 16)
        $consultas = [ordered]@{
            clientes = ("SELECT id FROM clientes WHERE ativo=False AND criado_em>=?" +
                        " AND id NOT IN (SELECT id_registro FROM log_alteracoes WHERE tabela='clientes' AND acao='DESATIVACAO')")
            contatos = ("SELECT ct.id FROM contatos ct INNER JOIN clientes cl ON ct.id_cliente=cl.id" +
                        " WHERE ct.ativo=False AND cl.ativo=True AND ct.criado_em>=?" +
                        " AND ct.id NOT IN (SELECT id_registro FROM log_alteracoes WHERE tabela='contatos' AND acao='DESATIVACAO')")
        }
        foreach ($tabela in $consultas.Keys) {
            $ids = New-Object System.Collections.ArrayList
            $r = (New-ComandoBanco $cn $consultas[$tabela] @((New-Parametro 'data' $desde)) $tx).ExecuteReader()
            while ($r.Read()) { [void]$ids.Add([int]$r.GetValue(0)) }
            $r.Close()
            foreach ($id in $ids) {
                [void](Invoke-Comando $cn ('UPDATE ' + $tabela + ' SET ativo=True, versao=versao+1 WHERE id=?') @((New-Parametro 'inteiro' $id)) $tx)
                [void](Invoke-Comando $cn ('INSERT INTO log_alteracoes (tabela, id_registro, campo, valor_antigo, valor_novo, acao, usuario, quando, origem)' +
                                           ' VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)') @(
                    (New-Parametro 'texto' $tabela), (New-Parametro 'inteiro' $id), (New-Parametro 'texto' 'ativo'),
                    (New-Parametro 'memo' 'False'), (New-Parametro 'memo' 'True'), (New-Parametro 'texto' 'CORRECAO'),
                    (New-Parametro 'texto' 'migracao-002'), (New-Parametro 'data' (Get-Date)), (New-Parametro 'texto' 'MIGRACAO')) $tx)
            }
            $log.L('  ' + $tabela + ': ' + $ids.Count + ' registro(s) reativado(s)' + $(if ($ids.Count) { ' (ids ' + ($ids -join ', ') + ')' } else { '' }))
        }
    }
}
