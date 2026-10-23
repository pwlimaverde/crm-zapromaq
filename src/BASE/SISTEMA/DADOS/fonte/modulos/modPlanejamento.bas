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

' Municipios para a busca da cidade-base (frmVinculo modo municipio, fatia 4.7d).
' busca = "Nome" ou "Nome/UF"; compara pela chave (nome_busca), em qualquer parte
' do nome. Devolve ate 300 linhas base 1: cod_ibge, nome, uf, nome_busca - ou Empty.
Public Function ListarMunicipios(ByVal busca As String) As Variant
    Dim partes As Variant, nome As String, uf As String, sql As String, params As Variant
    busca = Trim$(busca)
    If busca <> "" Then
        partes = Split(busca, "/")
        nome = Trim$(partes(0))
        If UBound(partes) >= 1 Then uf = UCase$(Trim$(partes(1)))
    End If
    sql = "SELECT TOP 300 cod_ibge, nome, uf, nome_busca FROM municipios WHERE nome_busca LIKE ?"
    If uf <> "" Then
        sql = sql & " AND uf=?"
        params = Array(modDB.P(modDB.adVarWChar, modCRM.PadraoLike(NomeBusca(nome))), modDB.P(modDB.adVarWChar, uf))
    Else
        params = Array(modDB.P(modDB.adVarWChar, modCRM.PadraoLike(NomeBusca(nome))))
    End If
    ListarMunicipios = modDB.Consultar(sql & " ORDER BY nome, uf", params)
End Function

' Linha (base 1) do municipio de nome exatamente igual a busca, ou 0. Com UF na
' busca e uma linha exata so, a escolha e direta, sem abrir a lista.
Public Function MunicipioExato(ByVal lista As Variant, ByVal busca As String) As Long
    Dim i As Long, chave As String, achados As Long
    If IsEmpty(lista) Then Exit Function
    chave = NomeBusca(Trim$(Split(Trim$(busca) & "/", "/")(0)))
    For i = 1 To UBound(lista, 1)
        If CStr(lista(i, 4)) = chave Then
            achados = achados + 1
            MunicipioExato = i
        End If
    Next i
    If achados <> 1 Then MunicipioExato = 0
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
                                  ByVal inicio As Date, ByVal fim As Date, ByVal observacoes As String, _
                                  Optional ByVal responsavel As String = "") As String
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
                             "data_inicio, data_fim, situacao, responsavel, observacoes, versao, criado_em, alterado_em, alterado_por) " & _
                             "VALUES (?,?,?,?,?,?,?,'ABERTO',?,?,1,?,?,?)", _
                         Array(modDB.P(modDB.adVarWChar, codigo), modDB.P(modDB.adInteger, controle), _
                               modDB.P(modDB.adVarWChar, descricao), modDB.P(modDB.adInteger, codMunicipio), _
                               modDB.P(modDB.adInteger, raio), modDB.P(modDB.adDate, inicio), modDB.P(modDB.adDate, fim), _
                               modDB.P(modDB.adVarWChar, Trim$(responsavel)), modDB.P(modDB.adLongVarWChar, observacoes), _
                               modDB.P(modDB.adDate, agora), modDB.P(modDB.adDate, agora), _
                               modDB.P(modDB.adVarWChar, modConfig.UsuarioAtual())))
    modDB.LogEm cn, "planejamentos", id, "codigo", "", codigo, "INSERIR"
    modDB.Confirmar cn
    CriarPlanejamento = ""
    Exit Function
Falha:
    ' a descricao antes do Desfazer: o On Error de dentro dele limpa o Err
    Dim descErro As String
    descErro = Err.Description
    modDB.Desfazer cn
    CriarPlanejamento = "Nao gravou: " & descErro
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
    ' a descricao antes do Desfazer: o On Error de dentro dele limpa o Err
    Dim descErro As String
    descErro = Err.Description
    modDB.Desfazer cn
    AlterarSituacaoPlanejamento = "Nao gravou: " & descErro
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
    ' a descricao antes do Desfazer: o On Error de dentro dele limpa o Err
    Dim descErro As String
    descErro = Err.Description
    modDB.Desfazer cn
    AtualizarPlanejamento = "Nao gravou: " & descErro
End Function

' Candidatos do planejamento (espelho de listar_candidatos): atendimentos em aberto a ate o raio
' da cidade-base, ordem qualificacao, dias parado (mais tempo primeiro), km (menor primeiro).
' Devolve uma matriz base 1: id, codigo, empresa, cidade, uf, qualificacao, km, dias, pendente.
' Vazia quando nao ha candidatos. Uma conexao e leituras fixas (nenhuma consulta por atendimento):
' o papel da etapa vem do cache de modCRM e as pendentes de um GROUP BY. Cidade, UF e
' qualificacao sao opcionais no cliente: nulo nao quebra (sem cidade = fora; sem qualificacao = 999).
Public Function Candidatos(ByVal idPlanejamento As Long) As Variant
    Dim d As Variant, base As Variant, mapa As Object, ordens As Object, pendentes As Object, abertos As Variant
    Dim i As Long, j As Long, k As Long, n As Long, chave As String, km As Double, m As Variant
    Dim linhas() As Variant, ordQ As Long, pend As Long, dias As Variant, t As Variant, qual As String
    Dim cn As Object
    On Error GoTo Falha
    Set cn = modDB.AbrirLeitura()
    d = modDB.ConsultarEm(cn, "SELECT cod_municipio_base, raio_km FROM planejamentos WHERE id=?", Array(modDB.P(modDB.adInteger, idPlanejamento)))
    If Not IsEmpty(d) Then
        base = modDB.ConsultarEm(cn, "SELECT lat, lon FROM municipios WHERE cod_ibge=?", Array(modDB.P(modDB.adInteger, CLng(d(1, 1)))))
    End If
    If IsEmpty(d) Or IsEmpty(base) Then
        modDB.FecharLeitura cn
        Exit Function
    End If
    Set mapa = CreateObject("Scripting.Dictionary")
    mapa.CompareMode = 1
    m = modDB.ConsultarEm(cn, "SELECT nome_busca, uf, lat, lon FROM municipios")
    If Not IsEmpty(m) Then
        For i = 1 To UBound(m, 1)
            mapa(CStr(m(i, 2)) & "|" & CStr(m(i, 1))) = Array(m(i, 3), m(i, 4))
        Next i
    End If
    Set ordens = CreateObject("Scripting.Dictionary")
    ordens.CompareMode = 1
    m = modDB.ConsultarEm(cn, "SELECT valor, ordem FROM listas WHERE tipo='Qualificacao'")
    If Not IsEmpty(m) Then
        For i = 1 To UBound(m, 1)
            If Not IsNull(m(i, 2)) Then ordens(CStr(m(i, 1))) = CLng(m(i, 2))
        Next i
    End If
    Set pendentes = CreateObject("Scripting.Dictionary")
    m = modDB.ConsultarEm(cn, "SELECT id_oportunidade, COUNT(*) FROM visitas WHERE situacao='PENDENTE' GROUP BY id_oportunidade")
    If Not IsEmpty(m) Then
        For i = 1 To UBound(m, 1)
            pendentes(CLng(m(i, 1))) = CLng(m(i, 2))
        Next i
    End If
    abertos = modDB.ConsultarEm(cn, "SELECT o.id, o.codigo, o.etapa, o.ultima_interacao, c.empresa, c.cidade, c.uf, c.qualificacao " & _
                                    "FROM oportunidades o INNER JOIN clientes c ON o.id_cliente = c.id")
    modDB.FecharLeitura cn
    Set cn = Nothing
    If IsEmpty(abertos) Then Exit Function
    ReDim linhas(1 To UBound(abertos, 1))
    n = 0
    For i = 1 To UBound(abertos, 1)
        If Not modCRM.EtapaEncerra(modDB.Nz(abertos(i, 3))) Then
            chave = modDB.Nz(abertos(i, 7)) & "|" & NomeBusca(modDB.Nz(abertos(i, 6)))
            If mapa.Exists(chave) Then
                km = DistanciaKm(CDbl(base(1, 1)), CDbl(base(1, 2)), CDbl(mapa(chave)(0)), CDbl(mapa(chave)(1)))
                If km <= CDbl(d(1, 2)) Then
                    n = n + 1
                    qual = modDB.Nz(abertos(i, 8))
                    ordQ = 999
                    If ordens.Exists(qual) Then ordQ = ordens(qual)
                    pend = 0
                    If pendentes.Exists(CLng(abertos(i, 1))) Then pend = pendentes(CLng(abertos(i, 1)))
                    dias = Null
                    If Not IsNull(abertos(i, 4)) Then dias = DateDiff("d", CDate(abertos(i, 4)), Date)
                    linhas(n) = Array(CLng(abertos(i, 1)), CStr(abertos(i, 2)), modDB.Nz(abertos(i, 5)), modDB.Nz(abertos(i, 6)), _
                                      modDB.Nz(abertos(i, 7)), qual, Round(km, 1), dias, pend, ordQ)
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
    Exit Function
Falha:
    Dim numErro As Long, descErro As String
    numErro = Err.Number: descErro = Err.Description
    If Not cn Is Nothing Then modDB.FecharLeitura cn
    Err.Raise numErro, "modPlanejamento.Candidatos", descErro
End Function

' Ordem de visita: qualificacao (ordem da lista), dias parado (mais tempo primeiro; sem data = 0,
' como no conector), km (menor primeiro).
Private Function CandidatoAntesQue(ByVal a As Variant, ByVal b As Variant) As Boolean
    Dim da As Long, db As Long
    If Not IsNull(a(7)) Then da = a(7)
    If Not IsNull(b(7)) Then db = b(7)
    If a(9) <> b(9) Then
        CandidatoAntesQue = (a(9) < b(9))
    ElseIf da <> db Then
        CandidatoAntesQue = (da > db)
    Else
        CandidatoAntesQue = (a(6) < b(6))
    End If
End Function

' Gera uma visita PENDENTE por atendimento escolhido, no planejamento, numa transacao.
' Qualquer recusa (encerrado, ja com pendente) desfaz tudo. Devolve "" ou a mensagem.
Public Function GerarVisitas(ByVal idPlanejamento As Long, ByVal idsOportunidade As Variant) As String
    Dim cn As Object, i As Long, pend As Long, id As Long, agora As Date, msg As String, etapa As String
    If IsEmpty(idsOportunidade) Then GerarVisitas = "Marque ao menos um atendimento.": Exit Function
    Set cn = modDB.AbrirTransacao()
    On Error GoTo Falha
    ' como no conector: planejamento encerrado nao recebe visitas novas (e tem de existir)
    msg = CStr(modDB.ConsultarValorEm(cn, "SELECT situacao FROM planejamentos WHERE id=?", "", _
                                      Array(modDB.P(modDB.adInteger, idPlanejamento))))
    If msg = "" Then
        modDB.Desfazer cn
        GerarVisitas = "Planejamento nao encontrado."
        Exit Function
    ElseIf msg <> "ABERTO" Then
        modDB.Desfazer cn
        GerarVisitas = "Planejamento encerrado: reabra antes de gerar visitas."
        Exit Function
    End If
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
    ' a descricao antes do Desfazer: o On Error de dentro dele limpa o Err
    Dim descErro As String
    descErro = Err.Description
    modDB.Desfazer cn
    GerarVisitas = "Nao gravou: " & descErro
End Function

' Grava o roteiro do planejamento (mesmo formato do conector): UTF-8 com BOM, ';', texto
' entre aspas, numero com ponto decimal, data AAAA-MM-DD. Linha BASE primeiro, depois uma
' linha por visita PENDENTE. Devolve "" e o caminho, ou a mensagem.
Public Function ExportarRoteiro(ByVal idPlanejamento As Long, ByVal pasta As String, ByRef caminho As String) As String
    Dim pl As Variant, base As Variant, vis As Variant, m As Variant, linhas As String, i As Long
    Dim mapa As Object, chave As String, lat As Variant, lon As Variant, km As Variant, dt As String
    Dim stream As Object, arq As String, hr As String, msg As String
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
            linhas = linhas & """VISITA"";" & CStr(vis(i, 1)) & ";""" & Aspas(CStr(vis(i, 6))) & """;""" & Aspas(CStr(vis(i, 7))) & """;""" & _
                     Aspas(Nz(vis(i, 10))) & """;""" & Aspas(Nz(vis(i, 11))) & """;""" & Aspas(CStr(vis(i, 8))) & """;""" & _
                     Aspas(CStr(vis(i, 9))) & """;" & _
                     IIf(IsEmpty(lat), "", Trim$(Str$(CDbl(lat)))) & ";" & IIf(IsEmpty(lon), "", Trim$(Str$(CDbl(lon)))) & ";" & _
                     IIf(IsEmpty(km), "", Trim$(Str$(CDbl(km)))) & ";""" & dt & """;""" & hr & """;""" & _
                     Aspas(Nz(vis(i, 4))) & """;""" & Aspas(Nz(vis(i, 5))) & """;""PENDENTE""" & vbCrLf
        Next i
    End If
    msg = GarantirPasta(pasta)
    If msg <> "" Then ExportarRoteiro = msg: Exit Function
    arq = pasta & "\" & CStr(pl(1, 1)) & "_" & Format$(pl(1, 2), "yyyy-mm-dd") & "_a_" & Format$(pl(1, 3), "yyyy-mm-dd") & ".csv"
    On Error GoTo FalhaGravar
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText linhas
    stream.SaveToFile arq, 2
    stream.Close
    caminho = arq
    ' DEV-23: manifesto ao lado do CSV (o descarte do conector so apaga pasta com manifesto valido)
    Dim nVis As Long
    If Not IsEmpty(vis) Then nVis = UBound(vis, 1)
    GravarManifestoRoteiro pasta, CStr(pl(1, 1)), Mid$(arq, InStrRev(arq, "\") + 1), nVis
    ExportarRoteiro = ""
    Exit Function
FalhaGravar:
    ExportarRoteiro = "Nao foi possivel gravar o roteiro (" & arq & "): " & Err.Description
End Function

'----------------------------------------------------------
' DEV-23 (ciclo 012, secao 5.3): extratos e roteiros ficam na raiz
' de extratos, ao lado do banco (<BASE>\SISTEMA\DADOS\execucao\extratos),
' uma subpasta por extrato - nunca em Documentos (sincronizado).
'----------------------------------------------------------
Public Function PastaRaizExtratos() As String
    Dim c As String, k As Long
    c = modConfig.CaminhoBanco()
    k = InStrRev(c, "\")
    If k = 0 Then Exit Function
    PastaRaizExtratos = Left$(c, k) & "SISTEMA\DADOS\execucao\extratos"
End Function

' Subpasta nova (ainda nao criada) para o roteiro: AAAAMMDD-HHMMSS-<6 hex>-roteiro-<id>.
Public Function NovaPastaRoteiro(ByVal idPlanejamento As Long) As String
    Dim raiz As String
    raiz = PastaRaizExtratos()
    If raiz = "" Then Exit Function
    Randomize
    NovaPastaRoteiro = raiz & "\" & Format$(Now, "yyyymmdd-hhnnss") & "-" & _
                       LCase$(Right$("000000" & Hex$(CLng(Rnd * 16777215)), 6)) & "-roteiro-" & idPlanejamento
End Function

' Manifesto JSON (UTF-8) no formato do conector: tipo, instante, solicitante, estacao e finalidade.
' O hash do arquivo nao e calculado aqui (o VBA nao tem SHA-256 nativo): o manifesto diz isso.
Private Sub GravarManifestoRoteiro(ByVal pasta As String, ByVal codigo As String, ByVal arquivo As String, ByVal visitas As Long)
    Dim j As String, stream As Object
    On Error Resume Next
    j = "{" & vbCrLf & _
        "  ""tipo"": ""extrato-crm""," & vbCrLf & _
        "  ""versao_manifesto"": 1," & vbCrLf & _
        "  ""conteudo"": ""roteiro""," & vbCrLf & _
        "  ""origem"": ""FRONT""," & vbCrLf & _
        "  ""criado_em"": """ & Format$(Now, "yyyy-mm-dd") & "T" & Format$(Now, "hh:nn:ss") & """," & vbCrLf & _
        "  ""solicitante"": """ & Json(modConfig.UsuarioAtual()) & """," & vbCrLf & _
        "  ""estacao"": """ & Json(Environ$("COMPUTERNAME")) & """," & vbCrLf & _
        "  ""usuario_tecnico"": """ & Json(Environ$("USERNAME")) & """," & vbCrLf & _
        "  ""versao_cliente"": ""FRONT " & Json(modConfig.VERSAO_FRONT) & """," & vbCrLf & _
        "  ""finalidade"": ""roteiro de visitas""," & vbCrLf & _
        "  ""planejamento"": """ & Json(codigo) & """," & vbCrLf & _
        "  ""arquivos"": { ""roteiro"": { ""arquivo"": """ & Json(arquivo) & """, ""visitas"": " & visitas & ", ""sha256"": null } }," & vbCrLf & _
        "  ""hash"": ""nao calculado pelo front""," & vbCrLf & _
        "  ""retencao"": ""mes corrente e anterior (secao 5.3)""," & vbCrLf & _
        "  ""em_processamento"": false" & vbCrLf & "}"
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText j
    stream.SaveToFile pasta & "\crm-extrato.json", 2
    stream.Close
End Sub

' Texto dentro de uma cadeia JSON: barra invertida e aspas escapadas.
Private Function Json(ByVal s As String) As String
    Json = Replace(Replace(s, "\", "\\"), """", "\""")
End Function

' Aspas dobradas dentro de um campo entre aspas do CSV (mesmo formato do conector).
Private Function Aspas(ByVal texto As String) As String
    Aspas = Replace(texto, """", """""")
End Function

' Cria a pasta e as que faltam acima dela. "" = existe; senao, a mensagem.
Private Function GarantirPasta(ByVal pasta As String) As String
    Dim fso As Object, pai As String
    On Error GoTo Falha
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(pasta) Then
        pai = fso.GetParentFolderName(pasta)
        If pai <> "" Then
            GarantirPasta = GarantirPasta(pai)
            If GarantirPasta <> "" Then Exit Function
        End If
        fso.CreateFolder pasta
    End If
    Exit Function
Falha:
    GarantirPasta = "Nao foi possivel criar a pasta do roteiro (" & pasta & "): " & Err.Description
End Function

Private Function Nz(ByVal v As Variant) As String
    If IsNull(v) Or IsEmpty(v) Then Nz = "" Else Nz = CStr(v)
End Function
