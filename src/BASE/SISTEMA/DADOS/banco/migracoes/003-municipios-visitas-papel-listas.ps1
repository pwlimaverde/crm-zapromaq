<#
  003 - esquema 1.0 -> 1.1: municípios, planejamentos, visitas, colunas de papel em listas.

  Migração ESTRUTURAL (EsquemaDe 1.0, EsquemaPara 1.1): roda com o banco exclusivo, sem transação
  (DDL do ACE não é transacional). Quem executa é a montagem, sobre uma CÓPIA do banco
  (SISTEMA\DADOS\execucao\publicacao), e só troca o banco de produção no fim.

  Estrutura conforme spec 003 (Contratos > Tabelas). A carga dos municípios, a chave das listas,
  a unificação de variantes e o papel das etapas entram nas fatias 1.5 (dados).
  O índice único de listas (tipo, chave) é criado por último, depois que a chave estiver preenchida.
#>
@{
    Id = '003'
    Descricao = 'municípios, planejamentos, visitas e papel das listas (esquema 1.1)'
    EsquemaDe = '1.0'
    EsquemaPara = '1.1'
    Aplicar = {
        param($cn, $tx, $log)
        function Exec([string]$sql) { $cmd = $cn.CreateCommand(); $cmd.CommandText = $sql; [void]$cmd.ExecuteNonQuery(); $cmd.Dispose() }

        Exec ('CREATE TABLE municipios (cod_ibge LONG PRIMARY KEY, nome TEXT(80) NOT NULL, nome_busca TEXT(80) NOT NULL,' +
              ' uf TEXT(2) NOT NULL, lat DOUBLE NOT NULL, lon DOUBLE NOT NULL, fonte_carga TEXT(60))')
        Exec 'CREATE INDEX ix_mun_uf_busca ON municipios (uf, nome_busca)'

        Exec ('CREATE TABLE planejamentos (id AUTOINCREMENT PRIMARY KEY, codigo TEXT(12) NOT NULL, controle LONG NOT NULL,' +
              ' descricao TEXT(120) NOT NULL, cod_municipio_base LONG NOT NULL, raio_km INTEGER NOT NULL, data_inicio DATETIME,' +
              ' data_fim DATETIME, responsavel TEXT(60), situacao TEXT(20) NOT NULL, observacoes MEMO, versao LONG,' +
              ' criado_em DATETIME, alterado_em DATETIME, alterado_por TEXT(50))')
        Exec 'CREATE UNIQUE INDEX ix_pln_codigo ON planejamentos (codigo)'
        Exec 'CREATE UNIQUE INDEX ix_pln_controle ON planejamentos (controle)'
        Exec 'ALTER TABLE planejamentos ADD CONSTRAINT fk_pln_mun FOREIGN KEY (cod_municipio_base) REFERENCES municipios (cod_ibge)'

        Exec ('CREATE TABLE visitas (id AUTOINCREMENT PRIMARY KEY, id_oportunidade LONG NOT NULL, id_planejamento LONG,' +
              ' situacao TEXT(20) NOT NULL, data_agendada DATETIME, hora_agendada TEXT(20), responsavel TEXT(60), tipo TEXT(40),' +
              ' observacoes MEMO, data_realizada DATETIME, resultado TEXT(40), relato MEMO, motivo_cancelamento TEXT(120),' +
              ' versao LONG, criado_em DATETIME, alterado_em DATETIME, alterado_por TEXT(50))')
        Exec 'CREATE INDEX ix_vis_opo_sit ON visitas (id_oportunidade, situacao)'
        Exec 'CREATE INDEX ix_vis_pln ON visitas (id_planejamento)'
        Exec 'CREATE INDEX ix_vis_data ON visitas (data_agendada)'
        Exec 'ALTER TABLE visitas ADD CONSTRAINT fk_vis_opo FOREIGN KEY (id_oportunidade) REFERENCES oportunidades (id)'
        Exec 'ALTER TABLE visitas ADD CONSTRAINT fk_vis_pln FOREIGN KEY (id_planejamento) REFERENCES planejamentos (id)'

        Exec 'ALTER TABLE listas ADD COLUMN chave TEXT(120)'
        Exec 'ALTER TABLE listas ADD COLUMN encerra YESNO'
        Exec 'ALTER TABLE listas ADD COLUMN resultado TEXT(10)'
        Exec 'ALTER TABLE listas ADD COLUMN quadro TEXT(12)'
        Exec 'ALTER TABLE listas ADD COLUMN exige_familia YESNO'

        $log.L('  estrutura 1.1: municipios, planejamentos, visitas e colunas novas de listas criadas')
        # ix_lis_tipo_chave (UNIQUE): criado na fatia 1.5, depois da carga e da unificação das chaves
    }
}
