Attribute VB_Name = "modVisita"
Option Explicit
'==========================================================
' modVisita - regras das visitas ao atendimento (Fase 3).
'
' Espelho de IA\mcp\lib\Regras.ps1 (Test-VisitaPodeCriar,
' Test-DataNoPeriodo, Test-RealizadaValida, Test-CanceladaValida,
' Get-TextoHistoricoVisita, Test-SegundaPendente): a mesma regra
' nos dois lados, casos C08 a C13 da spec 003 (secao Regras de visita).
'
' Funcoes puras: devolvem a mensagem da recusa, ou vazio quando a
' regra passa. As gravacoes ficam em modVisita (fatia 3.2b).
'==========================================================

' C08: atendimento encerrado nao recebe visita.
Public Function VisitaPodeCriar(ByVal atendimentoEncerrado As Boolean) As String
    If atendimentoEncerrado Then VisitaPodeCriar = "Atendimento encerrado nao recebe visita (C08)."
End Function

' C09: data dentro do periodo do planejamento (inclusive as bordas).
Public Function DataNoPeriodo(ByVal data As Date, ByVal inicio As Date, ByVal fim As Date) As String
    If data < inicio Or data > fim Then
        DataNoPeriodo = "Data " & Format$(data, "dd/mm/yyyy") & " fora do periodo do planejamento (C09)."
    End If
End Function

' C10: realizada exige data e relato.
Public Function RealizadaValida(ByVal relato As String, ByVal dataRealizada As Variant) As String
    If Trim$(relato) = "" Or IsNull(dataRealizada) Then RealizadaValida = "Visita realizada exige data e relato (C10)."
End Function

' C11: cancelada exige motivo.
Public Function CanceladaValida(ByVal motivo As String) As String
    If Trim$(motivo) = "" Then CanceladaValida = "Visita cancelada exige motivo (C11)."
End Function

' C12: linha do historico do atendimento quando a visita e realizada.
Public Function TextoHistorico(ByVal quando As Date, ByVal tipo As String, ByVal resultado As String, _
                               ByVal relato As String) As String
    Dim resumo As String
    resumo = relato
    If Len(resumo) > 200 Then resumo = Left$(resumo, 200)
    TextoHistorico = "VISITA " & Format$(quando, "dd/mm/yyyy hh:nn") & " - " & tipo & " - " & resultado & ": " & resumo
End Function

' C13: so uma visita PENDENTE por atendimento por vez.
Public Function SegundaPendente(ByVal pendentesExistentes As Long) As String
    If pendentesExistentes > 0 Then SegundaPendente = "Ja ha visita pendente neste atendimento: realize ou cancele antes (C13)."
End Function

' Lancamento (frmLancamento): para onde vai a data/hora digitada. Visita PENDENTE e
' agendamento (data e hora agendadas, sem data realizada - como lancar_datas no conector);
' REALIZADA/CANCELADA guardam a data em data_realizada e mantem o agendamento.
' Devolve Array(data_agendada, hora_agendada, data_realizada).
Public Function CamposLancamento(ByVal situacao As String, ByVal quando As Variant, ByVal hora As String, _
                                 ByVal dataAgAtual As Variant, ByVal horaAtual As String) As Variant
    If UCase$(Trim$(situacao)) = "PENDENTE" Then
        If IsNull(quando) Or IsEmpty(quando) Then
            CamposLancamento = Array(Null, Trim$(hora), Null)
        Else
            CamposLancamento = Array(DateValue(quando), Trim$(hora), Null)
        End If
    Else
        CamposLancamento = Array(dataAgAtual, horaAtual, quando)
    End If
End Function

'----------------------------------------------------------
' Gravacoes (fatia 3.2b). Mesmas regras do conector (C08 a C13), numa
' transacao cada. Devolvem "" quando gravam, ou a mensagem da recusa.
'----------------------------------------------------------

' Cria visita PENDENTE no atendimento (C08 e C13). Sem planejamento aqui.
Public Function CriarVisita(ByVal idOportunidade As Long, ByVal dataAgendada As Variant, _
                            ByVal hora As String, ByVal tipo As String, ByVal responsavel As String, _
                            ByVal observacoes As String) As String
    Dim cn As Object, etapa As String, encerra As Variant, pend As Variant, msg As String, id As Long, agora As Date
    Set cn = modDB.AbrirTransacao()
    On Error GoTo Falha
    etapa = modDB.ConsultarValorEm(cn, "SELECT etapa FROM oportunidades WHERE id=?", "", _
                                   Array(modDB.P(modDB.adInteger, idOportunidade)))
    If etapa = "" Then
        modDB.Desfazer cn
        CriarVisita = "Atendimento nao encontrado."
        Exit Function
    End If
    encerra = modDB.ConsultarValorEm(cn, "SELECT encerra FROM listas WHERE tipo='Etapa' AND valor=?", False, _
                                     Array(modDB.P(modDB.adVarWChar, etapa)))
    msg = VisitaPodeCriar(CBool(encerra))
    If msg = "" Then
        pend = modDB.ConsultarValorEm(cn, "SELECT COUNT(*) FROM visitas WHERE id_oportunidade=? AND situacao='PENDENTE'", 0, _
                                      Array(modDB.P(modDB.adInteger, idOportunidade)))
        msg = SegundaPendente(CLng(pend))
    End If
    If msg <> "" Then
        modDB.Desfazer cn
        CriarVisita = msg
        Exit Function
    End If
    agora = Now
    id = modDB.InserirEm(cn, "INSERT INTO visitas (id_oportunidade, situacao, data_agendada, hora_agendada, responsavel, " & _
                             "tipo, observacoes, versao, criado_em, alterado_em, alterado_por) VALUES (?,?,?,?,?,?,?,1,?,?,?)", _
                         Array(modDB.P(modDB.adInteger, idOportunidade), modDB.P(modDB.adVarWChar, "PENDENTE"), _
                               modDB.P(modDB.adDate, dataAgendada), modDB.P(modDB.adVarWChar, hora), _
                               modDB.P(modDB.adVarWChar, responsavel), modDB.P(modDB.adVarWChar, tipo), _
                               modDB.P(modDB.adLongVarWChar, observacoes), modDB.P(modDB.adDate, agora), _
                               modDB.P(modDB.adDate, agora), modDB.P(modDB.adVarWChar, modConfig.UsuarioAtual())))
    modDB.LogEm cn, "visitas", id, "situacao", "", "PENDENTE", "INSERIR"
    modDB.Confirmar cn
    CriarVisita = ""
    Exit Function
Falha:
    ' a descricao antes do Desfazer: o On Error de dentro dele limpa o Err
    Dim descErro As String
    descErro = Err.Description
    modDB.Desfazer cn
    CriarVisita = "Nao gravou: " & descErro
End Function

' Muda a situacao e os dados da visita. V4 e V5 (C10, C11); C13 ao voltar a PENDENTE.
' Ao virar REALIZADA: historico do atendimento (V6) e ultima interacao (V7), na mesma transacao.
' versaoLida: a versao que o usuario leu; se mudou, recusa (conflito).
Public Function AtualizarVisita(ByVal idVisita As Long, ByVal versaoLida As Long, ByVal situacao As String, _
                                ByVal dataAgendada As Variant, ByVal hora As String, ByVal tipo As String, _
                                ByVal responsavel As String, ByVal observacoes As String, _
                                ByVal dataRealizada As Variant, ByVal resultado As String, ByVal relato As String, _
                                ByVal motivo As String) As String
    Dim cn As Object, d As Variant, idOpo As Long, sitAnt As String, msg As String, n As Long, agora As Date
    Dim opObs As String, opVersao As Long, texto As String, pend As Variant
    Dim idPl As Variant, dataAnt As Variant, per As Variant, mudou As Boolean
    situacao = UCase$(Trim$(situacao))
    If InStr(1, "|PENDENTE|REALIZADA|CANCELADA|", "|" & situacao & "|", vbBinaryCompare) = 0 Then
        AtualizarVisita = "situacao deve ser PENDENTE, REALIZADA ou CANCELADA."
        Exit Function
    End If
    Set cn = modDB.AbrirTransacao()
    On Error GoTo Falha
    d = modDB.ConsultarEm(cn, "SELECT id_oportunidade, situacao, versao, id_planejamento, data_agendada FROM visitas WHERE id=?", _
                          Array(modDB.P(modDB.adInteger, idVisita)))
    If IsEmpty(d) Then
        modDB.Desfazer cn
        AtualizarVisita = "Visita nao encontrada."
        Exit Function
    End If
    idOpo = CLng(d(1, 1)): sitAnt = CStr(d(1, 2))
    idPl = d(1, 4): dataAnt = d(1, 5)
    If CLng(d(1, 3)) <> versaoLida Then
        modDB.Desfazer cn
        AtualizarVisita = "Conflito de versao: a visita foi alterada depois da leitura. Nada foi gravado; leia de novo."
        Exit Function
    End If
    If situacao = "REALIZADA" Then msg = RealizadaValida(relato, dataRealizada)
    If msg = "" And situacao = "CANCELADA" Then msg = CanceladaValida(motivo)
    ' C09 (como no conector): data agendada nova de visita de planejamento cai no periodo dele
    If msg = "" And Not IsNull(idPl) And Not IsNull(dataAgendada) And Not IsEmpty(dataAgendada) Then
        mudou = True
        If Not IsNull(dataAnt) Then mudou = (CDate(dataAnt) <> CDate(dataAgendada))
        If mudou Then
            per = modDB.ConsultarEm(cn, "SELECT data_inicio, data_fim FROM planejamentos WHERE id=?", _
                                    Array(modDB.P(modDB.adInteger, CLng(idPl))))
            If Not IsEmpty(per) Then msg = DataNoPeriodo(CDate(dataAgendada), CDate(per(1, 1)), CDate(per(1, 2)))
        End If
    End If
    If msg = "" And situacao = "PENDENTE" And sitAnt <> "PENDENTE" Then
        pend = modDB.ConsultarValorEm(cn, "SELECT COUNT(*) FROM visitas WHERE id_oportunidade=? AND situacao='PENDENTE' AND id<>?", 0, _
                                      Array(modDB.P(modDB.adInteger, idOpo), modDB.P(modDB.adInteger, idVisita)))
        msg = SegundaPendente(CLng(pend))
    End If
    If msg <> "" Then
        modDB.Desfazer cn
        AtualizarVisita = msg
        Exit Function
    End If
    agora = Now
    n = modDB.ExecutarEm(cn, "UPDATE visitas SET situacao=?, data_agendada=?, hora_agendada=?, responsavel=?, tipo=?, " & _
                             "observacoes=?, data_realizada=?, resultado=?, relato=?, motivo_cancelamento=?, " & _
                             "versao=versao+1, alterado_em=?, alterado_por=? WHERE id=? AND versao=?", _
                         Array(modDB.P(modDB.adVarWChar, situacao), modDB.P(modDB.adDate, dataAgendada), _
                               modDB.P(modDB.adVarWChar, hora), modDB.P(modDB.adVarWChar, responsavel), _
                               modDB.P(modDB.adVarWChar, tipo), modDB.P(modDB.adLongVarWChar, observacoes), _
                               modDB.P(modDB.adDate, dataRealizada), modDB.P(modDB.adVarWChar, resultado), _
                               modDB.P(modDB.adLongVarWChar, relato), modDB.P(modDB.adVarWChar, motivo), _
                               modDB.P(modDB.adDate, agora), modDB.P(modDB.adVarWChar, modConfig.UsuarioAtual()), _
                               modDB.P(modDB.adInteger, idVisita), modDB.P(modDB.adInteger, versaoLida)))
    If n <> 1 Then
        modDB.Desfazer cn
        AtualizarVisita = "Conflito de versao: nada foi gravado; leia de novo."
        Exit Function
    End If
    modDB.LogEm cn, "visitas", idVisita, "situacao", sitAnt, situacao, "ALTERAR"
    If situacao = "REALIZADA" And sitAnt <> "REALIZADA" Then
        d = modDB.ConsultarEm(cn, "SELECT observacoes, versao FROM oportunidades WHERE id=?", _
                              Array(modDB.P(modDB.adInteger, idOpo)))
        opObs = modDB.Nz(d(1, 1)): opVersao = CLng(d(1, 2))
        texto = TextoHistorico(CDate(dataRealizada), tipo, resultado, relato)
        If opObs <> "" Then opObs = opObs & vbCrLf
        opObs = opObs & "[" & Format$(Date, "dd/mm/yyyy") & " " & modConfig.UsuarioAtual() & "] " & texto
        n = modDB.ExecutarEm(cn, "UPDATE oportunidades SET observacoes=?, ultima_interacao=?, versao=versao+1, " & _
                                 "alterado_em=?, alterado_por=? WHERE id=? AND versao=?", _
                             Array(modDB.P(modDB.adLongVarWChar, opObs), modDB.P(modDB.adDate, CDate(dataRealizada)), _
                                   modDB.P(modDB.adDate, agora), modDB.P(modDB.adVarWChar, modConfig.UsuarioAtual()), _
                                   modDB.P(modDB.adInteger, idOpo), modDB.P(modDB.adInteger, opVersao)))
        If n <> 1 Then
            modDB.Desfazer cn
            AtualizarVisita = "Conflito de versao no atendimento: nada foi gravado; leia de novo."
            Exit Function
        End If
        modDB.LogEm cn, "oportunidades", idOpo, "ultima_interacao", "", CStr(dataRealizada), "ALTERAR"
    End If
    modDB.Confirmar cn
    AtualizarVisita = ""
    Exit Function
Falha:
    ' a descricao antes do Desfazer: o On Error de dentro dele limpa o Err
    Dim descErro As String
    descErro = Err.Description
    modDB.Desfazer cn
    AtualizarVisita = "Nao gravou: " & descErro
End Function

' Quantas visitas o atendimento tem, em qualquer situacao (V9): o rotulo Visitas (n).
Public Function ContarVisitas(ByVal idOportunidade As Long) As Long
    ContarVisitas = CLng(modDB.ConsultarValor("SELECT COUNT(*) FROM visitas WHERE id_oportunidade=?", 0, _
                                              Array(modDB.P(modDB.adInteger, idOportunidade))))
End Function
