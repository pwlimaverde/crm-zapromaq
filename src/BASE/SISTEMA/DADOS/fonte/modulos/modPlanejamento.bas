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
