Attribute VB_Name = "modPlanejamento"
Option Explicit
'==========================================================
' modPlanejamento - regras do planejamento por regiao (Fase 4).
'
' Espelho de IA\mcp\lib\Regras.ps1 (ConvertTo-NomeBusca,
' Get-DistanciaKm, Test-RaioValido, Test-PeriodoValido,
' Test-EncerramentoPermitido, Get-CodigoPlanejamento): a mesma regra
' nos dois lados (C14, C15, C16 e P4 da spec 003).
'
' Funcoes puras: devolvem a mensagem da recusa, ou vazio.
'==========================================================

Private Const RAIO_TERRA_KM As Double = 6371#

' Nome de municipio para comparar: a mesma chave das listas.
Public Function NomeBusca(ByVal nome As String) As String
    NomeBusca = modValidacao.ChaveLista(nome)
End Function

' Distancia em km entre dois pontos (Haversine). Mesma cidade = 0.
Public Function DistanciaKm(ByVal lat1 As Double, ByVal lon1 As Double, _
                            ByVal lat2 As Double, ByVal lon2 As Double) As Double
    Dim graus As Double, dLat As Double, dLon As Double, a As Double
    graus = Atn(1#) * 4# / 180#
    dLat = (lat2 - lat1) * graus
    dLon = (lon2 - lon1) * graus
    a = Sin(dLat / 2#) * Sin(dLat / 2#) + Cos(lat1 * graus) * Cos(lat2 * graus) * Sin(dLon / 2#) * Sin(dLon / 2#)
    If a <= 0# Then
        DistanciaKm = 0#
    ElseIf a >= 1# Then
        DistanciaKm = RAIO_TERRA_KM * 3.14159265358979
    Else
        ' asin(x) = atn(x / sqr(1 - x^2))
        DistanciaKm = 2# * RAIO_TERRA_KM * Atn(Sqr(a) / Sqr(1# - a))
    End If
End Function

' P1 (C14): raio de 1 a 500 km.
Public Function RaioValido(ByVal raio As Long) As String
    If raio < 1 Or raio > 500 Then RaioValido = "Raio deve ser de 1 a 500 km (C14)."
End Function

' P2 (C15): periodo preenchido e inicio nao depois do fim (datas AAAA-MM-DD).
Public Function PeriodoValido(ByVal inicio As String, ByVal fim As String) As String
    If Trim$(inicio) = "" Or Trim$(fim) = "" Then
        PeriodoValido = "Informe inicio e fim do periodo (C15)."
    ElseIf inicio > fim Then
        PeriodoValido = "O inicio e depois do fim do periodo (C15)."
    End If
End Function

' P3 (C16): nao encerra com visita pendente sem cancelar_pendentes.
Public Function EncerramentoPermitido(ByVal pendentes As Long, ByVal cancelarPendentes As Boolean) As String
    If pendentes > 0 And Not cancelarPendentes Then
        EncerramentoPermitido = "Ha " & pendentes & " visita(s) pendente(s): cancele-as ou confirme o cancelamento (C16)."
    End If
End Function

' P4: codigo PL-AAAA-NNN.
Public Function CodigoPlanejamento(ByVal ano As Long, ByVal numero As Long) As String
    CodigoPlanejamento = "PL-" & ano & "-" & Format$(numero, "000")
End Function

'----------------------------------------------------------
' Gravacoes (fatia 4.3a). Mesmas regras do conector (C14, C15, C16, P4),
' numa transacao cada. Devolvem "" quando gravam, ou a mensagem da recusa.
'----------------------------------------------------------

' Cria um planejamento ABERTO. codMunicipio tem de existir em municipios (P1).
Public Function CriarPlanejamento(ByVal descricao As String, ByVal codMunicipio As Long, ByVal raio As Long, _
                                  ByVal inicio As Date, ByVal fim As Date, ByVal observacoes As String) As String
    Dim cn As Object, msg As String, controle As Long, codigo As String, id As Long, agora As Date
    descricao = Trim$(descricao)
    If descricao = "" Then CriarPlanejamento = "Informe a descricao do planejamento.": Exit Function
    If Len(descricao) > 120 Then CriarPlanejamento = "Descricao aceita no maximo 120 caracteres.": Exit Function
    msg = RaioValido(raio)
    If msg = "" Then msg = PeriodoValido(Format$(inicio, "yyyy-mm-dd"), Format$(fim, "yyyy-mm-dd"))
    If msg <> "" Then CriarPlanejamento = msg: Exit Function
    Set cn = modDB.AbrirTransacao()
    On Error GoTo Falha
    If CLng(modDB.ConsultarValorEm(cn, "SELECT COUNT(*) FROM municipios WHERE cod_ibge=?", 0, _
                                   Array(modDB.P(modDB.adInteger, codMunicipio)))) = 0 Then
        modDB.Desfazer cn
        CriarPlanejamento = "Cidade-base nao encontrada (codigo IBGE " & codMunicipio & ")."
        Exit Function
    End If
    controle = CLng(modDB.ConsultarValorEm(cn, "SELECT MAX(controle) FROM planejamentos", 0)) + 1
    codigo = CodigoPlanejamento(Year(Now), controle)
    agora = Now
    id = modDB.InserirEm(cn, "INSERT INTO planejamentos (codigo, controle, descricao, cod_municipio_base, raio_km, " & _
                             "data_inicio, data_fim, situacao, observacoes, versao, criado_em, alterado_em, alterado_por) " & _
                             "VALUES (?,?,?,?,?,?,?,'ABERTO',?,1,?,?,?)", _
                         Array(modDB.P(modDB.adVarWChar, codigo), modDB.P(modDB.adInteger, controle), _
                               modDB.P(modDB.adVarWChar, descricao), modDB.P(modDB.adInteger, codMunicipio), _
                               modDB.P(modDB.adInteger, raio), modDB.P(modDB.adDate, inicio), modDB.P(modDB.adDate, fim), _
                               modDB.P(modDB.adLongVarWChar, observacoes), modDB.P(modDB.adDate, agora), _
                               modDB.P(modDB.adDate, agora), modDB.P(modDB.adVarWChar, modConfig.UsuarioAtual())))
    modDB.LogEm cn, "planejamentos", id, "codigo", "", codigo, "INSERIR"
    modDB.Confirmar cn
    CriarPlanejamento = ""
    Exit Function
Falha:
    modDB.Desfazer cn
    CriarPlanejamento = "Nao gravou: " & Err.Description
End Function

' Encerra (ENCERRADO) ou reabre (ABERTO). C16: encerrar com visitas pendentes so com
' cancelarPendentes, que cancela cada uma com o motivo "planejamento encerrado".
Public Function AlterarSituacaoPlanejamento(ByVal id As Long, ByVal versaoLida As Long, ByVal situacao As String, _
                                            ByVal cancelarPendentes As Boolean) As String
    Dim cn As Object, d As Variant, sitAnt As String, versao As Long, pend As Variant, i As Long, n As Long
    Dim msg As String, agora As Date
    situacao = UCase$(Trim$(situacao))
    If situacao <> "ABERTO" And situacao <> "ENCERRADO" Then AlterarSituacaoPlanejamento = "situacao deve ser ABERTO ou ENCERRADO.": Exit Function
    Set cn = modDB.AbrirTransacao()
    On Error GoTo Falha
    d = modDB.ConsultarEm(cn, "SELECT situacao, versao FROM planejamentos WHERE id=?", Array(modDB.P(modDB.adInteger, id)))
    If IsEmpty(d) Then
        modDB.Desfazer cn
        AlterarSituacaoPlanejamento = "Planejamento nao encontrado."
        Exit Function
    End If
    sitAnt = CStr(d(1, 1)): versao = CLng(d(1, 2))
    If versao <> versaoLida Then
        modDB.Desfazer cn
        AlterarSituacaoPlanejamento = "Conflito de versao: o planejamento foi alterado depois da leitura. Nada foi gravado; leia de novo."
        Exit Function
    End If
    If sitAnt = situacao Then
        modDB.Desfazer cn
        AlterarSituacaoPlanejamento = ""
        Exit Function
    End If
    agora = Now
    If situacao = "ENCERRADO" Then
        pend = modDB.ConsultarEm(cn, "SELECT id, versao FROM visitas WHERE id_planejamento=? AND situacao='PENDENTE'", _
                                 Array(modDB.P(modDB.adInteger, id)))
        msg = EncerramentoPermitido(IIf(IsEmpty(pend), 0, UBound(pend, 1)), cancelarPendentes)
        If msg <> "" Then
            modDB.Desfazer cn
            AlterarSituacaoPlanejamento = msg
            Exit Function
        End If
        If Not IsEmpty(pend) Then
            For i = 1 To UBound(pend, 1)
                n = modDB.ExecutarEm(cn, "UPDATE visitas SET situacao='CANCELADA', motivo_cancelamento=?, versao=versao+1, " & _
                                         "alterado_em=?, alterado_por=? WHERE id=? AND versao=?", _
                                     Array(modDB.P(modDB.adVarWChar, "planejamento encerrado"), modDB.P(modDB.adDate, agora), _
                                           modDB.P(modDB.adVarWChar, modConfig.UsuarioAtual()), _
                                           modDB.P(modDB.adInteger, CLng(pend(i, 1))), modDB.P(modDB.adInteger, CLng(pend(i, 2)))))
                If n <> 1 Then
                    modDB.Desfazer cn
                    AlterarSituacaoPlanejamento = "Conflito numa visita pendente: nada foi gravado; leia de novo."
                    Exit Function
                End If
                modDB.LogEm cn, "visitas", CLng(pend(i, 1)), "situacao", "PENDENTE", "CANCELADA", "ALTERAR"
            Next i
        End If
    End If
    n = modDB.ExecutarEm(cn, "UPDATE planejamentos SET situacao=?, versao=versao+1, alterado_em=?, alterado_por=? " & _
                             "WHERE id=? AND versao=?", _
                         Array(modDB.P(modDB.adVarWChar, situacao), modDB.P(modDB.adDate, agora), _
                               modDB.P(modDB.adVarWChar, modConfig.UsuarioAtual()), _
                               modDB.P(modDB.adInteger, id), modDB.P(modDB.adInteger, versaoLida)))
    If n <> 1 Then
        modDB.Desfazer cn
        AlterarSituacaoPlanejamento = "Conflito de versao: nada foi gravado; leia de novo."
        Exit Function
    End If
    modDB.LogEm cn, "planejamentos", id, "situacao", sitAnt, situacao, "ALTERAR"
    modDB.Confirmar cn
    AlterarSituacaoPlanejamento = ""
    Exit Function
Falha:
    modDB.Desfazer cn
    AlterarSituacaoPlanejamento = "Nao gravou: " & Err.Description
End Function

' Altera um planejamento ABERTO (descricao, cidade-base, raio, periodo, responsavel, observacoes).
' Mesmas regras de criar: C14, C15 e cidade existente. Conflito pela versao lida.
Public Function AtualizarPlanejamento(ByVal id As Long, ByVal versaoLida As Long, ByVal descricao As String, _
                                      ByVal codMunicipio As Long, ByVal raio As Long, ByVal inicio As Date, _
                                      ByVal fim As Date, ByVal responsavel As String, ByVal observacoes As String) As String
    Dim cn As Object, d As Variant, sitAnt As String, msg As String, n As Long, agora As Date
    descricao = Trim$(descricao)
    If descricao = "" Then AtualizarPlanejamento = "Informe a descricao do planejamento.": Exit Function
    msg = RaioValido(raio)
    If msg = "" Then msg = PeriodoValido(Format$(inicio, "yyyy-mm-dd"), Format$(fim, "yyyy-mm-dd"))
    If msg <> "" Then AtualizarPlanejamento = msg: Exit Function
    Set cn = modDB.AbrirTransacao()
    On Error GoTo Falha
    d = modDB.ConsultarEm(cn, "SELECT situacao, versao FROM planejamentos WHERE id=?", Array(modDB.P(modDB.adInteger, id)))
    If IsEmpty(d) Then
        modDB.Desfazer cn
        AtualizarPlanejamento = "Planejamento nao encontrado."
        Exit Function
    End If
    sitAnt = CStr(d(1, 1))
    If CLng(d(1, 2)) <> versaoLida Then
        modDB.Desfazer cn
        AtualizarPlanejamento = "Conflito de versao: o planejamento foi alterado depois da leitura. Nada foi gravado; leia de novo."
        Exit Function
    End If
    If sitAnt <> "ABERTO" Then
        modDB.Desfazer cn
        AtualizarPlanejamento = "Planejamento encerrado: reabra antes de alterar."
        Exit Function
    End If
    If CLng(modDB.ConsultarValorEm(cn, "SELECT COUNT(*) FROM municipios WHERE cod_ibge=?", 0, _
                                   Array(modDB.P(modDB.adInteger, codMunicipio)))) = 0 Then
        modDB.Desfazer cn
        AtualizarPlanejamento = "Cidade-base nao encontrada (codigo IBGE " & codMunicipio & ")."
        Exit Function
    End If
    agora = Now
    n = modDB.ExecutarEm(cn, "UPDATE planejamentos SET descricao=?, cod_municipio_base=?, raio_km=?, data_inicio=?, " & _
                             "data_fim=?, responsavel=?, observacoes=?, versao=versao+1, alterado_em=?, alterado_por=? " & _
                             "WHERE id=? AND versao=?", _
                         Array(modDB.P(modDB.adVarWChar, descricao), modDB.P(modDB.adInteger, codMunicipio), _
                               modDB.P(modDB.adInteger, raio), modDB.P(modDB.adDate, inicio), modDB.P(modDB.adDate, fim), _
                               modDB.P(modDB.adVarWChar, responsavel), modDB.P(modDB.adLongVarWChar, observacoes), _
                               modDB.P(modDB.adDate, agora), modDB.P(modDB.adVarWChar, modConfig.UsuarioAtual()), _
                               modDB.P(modDB.adInteger, id), modDB.P(modDB.adInteger, versaoLida)))
    If n <> 1 Then
        modDB.Desfazer cn
        AtualizarPlanejamento = "Conflito de versao: nada foi gravado; leia de novo."
        Exit Function
    End If
    modDB.LogEm cn, "planejamentos", id, "descricao", "", descricao, "ALTERAR"
    modDB.Confirmar cn
    AtualizarPlanejamento = ""
    Exit Function
Falha:
    modDB.Desfazer cn
    AtualizarPlanejamento = "Nao gravou: " & Err.Description
End Function

' Candidatos do planejamento (espelho de listar_candidatos): atendimentos em aberto a ate o raio
' da cidade-base, ordem qualificacao, dias parado (mais tempo primeiro), km (menor primeiro).
' Devolve uma matriz base 1: id, codigo, empresa, cidade, uf, qualificacao, km, dias, pendente.
' Vazia quando nao ha candidatos.
Public Function Candidatos(ByVal idPlanejamento As Long) As Variant
    Dim d As Variant, base As Variant, mapa As Object, ordens As Object, abertos As Variant
    Dim i As Long, j As Long, k As Long, n As Long, chave As String, km As Double, m As Variant
    Dim linhas() As Variant, ordQ As Long, pend As Long, dias As Variant, t As Variant
    d = modDB.Consultar("SELECT cod_municipio_base, raio_km FROM planejamentos WHERE id=?", Array(modDB.P(modDB.adInteger, idPlanejamento)))
    If IsEmpty(d) Then Exit Function
    base = modDB.Consultar("SELECT lat, lon FROM municipios WHERE cod_ibge=?", Array(modDB.P(modDB.adInteger, CLng(d(1, 1)))))
    If IsEmpty(base) Then Exit Function
    Set mapa = CreateObject("Scripting.Dictionary")
    mapa.CompareMode = 1
    m = modDB.Consultar("SELECT nome_busca, uf, lat, lon FROM municipios")
    If Not IsEmpty(m) Then
        For i = 1 To UBound(m, 1)
            mapa(CStr(m(i, 2)) & "|" & CStr(m(i, 1))) = Array(m(i, 3), m(i, 4))
        Next i
    End If
    Set ordens = CreateObject("Scripting.Dictionary")
    ordens.CompareMode = 1
    m = modDB.Consultar("SELECT valor, ordem FROM listas WHERE tipo='Qualificacao'")
    If Not IsEmpty(m) Then
        For i = 1 To UBound(m, 1)
            ordens(CStr(m(i, 1))) = CLng(m(i, 2))
        Next i
    End If
    abertos = modDB.Consultar("SELECT o.id, o.codigo, o.etapa, o.ultima_interacao, c.empresa, c.cidade, c.uf, c.qualificacao " & _
                              "FROM oportunidades o INNER JOIN clientes c ON o.id_cliente = c.id")
    If IsEmpty(abertos) Then Exit Function
    ReDim linhas(1 To UBound(abertos, 1))
    n = 0
    For i = 1 To UBound(abertos, 1)
        If Not EtapaEncerra(CStr(abertos(i, 3))) Then
            chave = CStr(abertos(i, 7)) & "|" & NomeBusca(CStr(abertos(i, 6)))
            If mapa.Exists(chave) Then
                km = DistanciaKm(CDbl(base(1, 1)), CDbl(base(1, 2)), CDbl(mapa(chave)(0)), CDbl(mapa(chave)(1)))
                If km <= CDbl(d(1, 2)) Then
                    n = n + 1
                    ordQ = 999
                    If ordens.Exists(CStr(abertos(i, 8))) Then ordQ = ordens(CStr(abertos(i, 8)))
                    pend = modDB.ConsultarValor("SELECT COUNT(*) FROM visitas WHERE id_oportunidade=? AND situacao='PENDENTE'", 0, _
                                                Array(modDB.P(modDB.adInteger, CLng(abertos(i, 1)))))
                    dias = Null
                    If Not IsNull(abertos(i, 4)) Then dias = DateDiff("d", CDate(abertos(i, 4)), Date)
                    linhas(n) = Array(CLng(abertos(i, 1)), CStr(abertos(i, 2)), CStr(abertos(i, 5)), CStr(abertos(i, 6)), _
                                      CStr(abertos(i, 7)), CStr(abertos(i, 8)), Round(km, 1), dias, pend, ordQ)
                End If
            End If
        End If
    Next i
    If n = 0 Then Exit Function
    ' ordenacao simples (poucos candidatos): qualificacao, dias desc, km
    For i = 1 To n - 1
        For j = i + 1 To n
            If CandidatoAntesQue(linhas(j), linhas(i)) Then
                t = linhas(i): linhas(i) = linhas(j): linhas(j) = t
            End If
        Next j
    Next i
    Dim saida() As Variant
    ReDim saida(1 To n, 1 To 10)
    For i = 1 To n
        For k = 0 To 9
            If k = 0 Then saida(i, 1) = linhas(i)(0) Else saida(i, k + 1) = linhas(i)(k)
        Next k
    Next i
    Candidatos = saida
End Function

' Ordem de visita: qualificacao (ordem da lista), dias parado (mais tempo primeiro), km (menor primeiro).
Private Function CandidatoAntesQue(ByVal a As Variant, ByVal b As Variant) As Boolean
    If a(9) <> b(9) Then
        CandidatoAntesQue = (a(9) < b(9))
    ElseIf IsNull(a(7)) <> IsNull(b(7)) Or (Not IsNull(a(7)) And a(7) <> b(7)) Then
        Dim da As Long, db As Long
        da = 0: db = 0
        If Not IsNull(a(7)) Then da = a(7)
        If Not IsNull(b(7)) Then db = b(7)
        CandidatoAntesQue = (da > db)
    Else
        CandidatoAntesQue = (a(6) < b(6))
    End If
End Function

' Etapa encerra o atendimento? (papel da lista; etapa sem linha nao encerra)
Private Function EtapaEncerra(ByVal etapa As String) As Boolean
    Dim v As Variant
    v = modDB.ConsultarValor("SELECT encerra FROM listas WHERE tipo='Etapa' AND valor=?", False, _
                             Array(modDB.P(modDB.adVarWChar, etapa)))
    EtapaEncerra = CBool(v)
End Function

' Gera uma visita PENDENTE por atendimento escolhido, no planejamento, numa transacao.
' Qualquer recusa (encerrado, ja com pendente) desfaz tudo. Devolve "" ou a mensagem.
Public Function GerarVisitas(ByVal idPlanejamento As Long, ByVal idsOportunidade As Variant) As String
    Dim cn As Object, i As Long, pend As Long, id As Long, agora As Date, msg As String, etapa As String
    If IsEmpty(idsOportunidade) Then GerarVisitas = "Marque ao menos um atendimento.": Exit Function
    Set cn = modDB.AbrirTransacao()
    On Error GoTo Falha
    For i = LBound(idsOportunidade) To UBound(idsOportunidade)
        etapa = modDB.ConsultarValorEm(cn, "SELECT etapa FROM oportunidades WHERE id=?", "", _
                                       Array(modDB.P(modDB.adInteger, CLng(idsOportunidade(i)))))
        If etapa = "" Then
            msg = "Atendimento nao encontrado."
        ElseIf CBool(modDB.ConsultarValorEm(cn, "SELECT encerra FROM listas WHERE tipo='Etapa' AND valor=?", False, _
                                            Array(modDB.P(modDB.adVarWChar, etapa)))) Then
            msg = VisitaPodeCriar(True)
        Else
            pend = CLng(modDB.ConsultarValorEm(cn, "SELECT COUNT(*) FROM visitas WHERE id_oportunidade=? AND situacao='PENDENTE'", 0, _
                                               Array(modDB.P(modDB.adInteger, CLng(idsOportunidade(i))))))
            msg = SegundaPendente(pend)
        End If
        If msg <> "" Then
            modDB.Desfazer cn
            GerarVisitas = "Nada gravado (atendimento " & CLng(idsOportunidade(i)) & "): " & msg
            Exit Function
        End If
        agora = Now
        id = modDB.InserirEm(cn, "INSERT INTO visitas (id_oportunidade, id_planejamento, situacao, versao, criado_em, " & _
                                 "alterado_em, alterado_por) VALUES (?,?,'PENDENTE',1,?,?,?)", _
                             Array(modDB.P(modDB.adInteger, CLng(idsOportunidade(i))), modDB.P(modDB.adInteger, idPlanejamento), _
                                   modDB.P(modDB.adDate, agora), modDB.P(modDB.adDate, agora), _
                                   modDB.P(modDB.adVarWChar, modConfig.UsuarioAtual())))
        modDB.LogEm cn, "visitas", id, "situacao", "", "PENDENTE", "INSERIR"
    Next i
    modDB.Confirmar cn
    GerarVisitas = ""
    Exit Function
Falha:
    modDB.Desfazer cn
    GerarVisitas = "Nao gravou: " & Err.Description
End Function

' Grava o roteiro do planejamento (mesmo formato do conector): UTF-8 com BOM, ';', texto
' entre aspas, numero com ponto decimal, data AAAA-MM-DD. Linha BASE primeiro, depois uma
' linha por visita PENDENTE. Devolve "" e o caminho, ou a mensagem.
Public Function ExportarRoteiro(ByVal idPlanejamento As Long, ByVal pasta As String, ByRef caminho As String) As String
    Dim pl As Variant, base As Variant, vis As Variant, m As Variant, linhas As String, i As Long
    Dim mapa As Object, chave As String, lat As Variant, lon As Variant, km As Variant, dt As String
    Dim stream As Object, arq As String, hr As String
    pl = modDB.Consultar("SELECT codigo, data_inicio, data_fim, cod_municipio_base FROM planejamentos WHERE id=?", _
                         Array(modDB.P(modDB.adInteger, idPlanejamento)))
    If IsEmpty(pl) Then ExportarRoteiro = "Planejamento nao encontrado.": Exit Function
    base = modDB.Consultar("SELECT nome, uf, lat, lon FROM municipios WHERE cod_ibge=?", _
                           Array(modDB.P(modDB.adInteger, CLng(pl(1, 4)))))
    If IsEmpty(base) Then ExportarRoteiro = "Cidade-base nao esta em municipios.": Exit Function
    Set mapa = CreateObject("Scripting.Dictionary")
    mapa.CompareMode = 1
    m = modDB.Consultar("SELECT nome_busca, uf, lat, lon FROM municipios")
    If Not IsEmpty(m) Then
        For i = 1 To UBound(m, 1)
            mapa(CStr(m(i, 2)) & "|" & CStr(m(i, 1))) = Array(m(i, 3), m(i, 4))
        Next i
    End If
    vis = modDB.Consultar("SELECT v.id, v.data_agendada, v.hora_agendada, v.tipo, v.responsavel, o.codigo, c.empresa, " & _
                          "c.cidade, c.uf, ct.nome, ct.telefone FROM ((visitas v INNER JOIN oportunidades o ON v.id_oportunidade = o.id) " & _
                          "INNER JOIN clientes c ON o.id_cliente = c.id) INNER JOIN contatos ct ON o.id_contato = ct.id " & _
                          "WHERE v.id_planejamento=? AND v.situacao='PENDENTE' ORDER BY v.data_agendada, v.hora_agendada, v.id", _
                          Array(modDB.P(modDB.adInteger, idPlanejamento)))
    linhas = "tipo;visita;atendimento;empresa;contato;telefone;cidade;uf;latitude;longitude;km_base;data_agendada;hora_agendada;tipo_visita;responsavel;situacao" & vbCrLf
    linhas = linhas & """BASE"";;;;;;""" & CStr(base(1, 1)) & """;""" & CStr(base(1, 2)) & """;" & _
             Trim$(Str$(CDbl(base(1, 3)))) & ";" & Trim$(Str$(CDbl(base(1, 4)))) & ";0;;;;;" & vbCrLf
    If Not IsEmpty(vis) Then
        For i = 1 To UBound(vis, 1)
            chave = CStr(vis(i, 9)) & "|" & NomeBusca(CStr(vis(i, 8)))
            lat = Empty: lon = Empty: km = Empty
            If mapa.Exists(chave) Then
                lat = mapa(chave)(0): lon = mapa(chave)(1)
                km = Round(DistanciaKm(CDbl(base(1, 3)), CDbl(base(1, 4)), CDbl(lat), CDbl(lon)), 1)
            End If
            dt = "": hr = ""
            If Not IsNull(vis(i, 2)) Then dt = Format$(vis(i, 2), "yyyy-mm-dd")
            If Not IsNull(vis(i, 3)) Then hr = Format$(vis(i, 3), "hh:nn")
            linhas = linhas & """VISITA"";" & CStr(vis(i, 1)) & ";""" & CStr(vis(i, 6)) & """;""" & CStr(vis(i, 7)) & """;""" & _
                     Nz(vis(i, 10)) & """;""" & Nz(vis(i, 11)) & """;""" & CStr(vis(i, 8)) & """;""" & CStr(vis(i, 9)) & """;" & _
                     IIf(IsEmpty(lat), "", Trim$(Str$(CDbl(lat)))) & ";" & IIf(IsEmpty(lon), "", Trim$(Str$(CDbl(lon)))) & ";" & _
                     IIf(IsEmpty(km), "", Trim$(Str$(CDbl(km)))) & ";""" & dt & """;""" & hr & """;""" & _
                     Nz(vis(i, 4)) & """;""" & Nz(vis(i, 5)) & """;""PENDENTE""" & vbCrLf
        Next i
    End If
    On Error Resume Next
    If Dir(pasta, vbDirectory) = "" Then MkDir pasta
    On Error GoTo 0
    arq = pasta & "\" & CStr(pl(1, 1)) & "_" & Format$(pl(1, 2), "yyyy-mm-dd") & "_a_" & Format$(pl(1, 3), "yyyy-mm-dd") & ".csv"
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText linhas
    stream.SaveToFile arq, 2
    stream.Close
    caminho = arq
    ExportarRoteiro = ""
End Function

Private Function Nz(ByVal v As Variant) As String
    If IsNull(v) Or IsEmpty(v) Then Nz = "" Else Nz = CStr(v)
End Function
