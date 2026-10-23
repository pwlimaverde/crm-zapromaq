Attribute VB_Name = "modCalc"
'==========================================================
' modCalc - campos calculados
'
' Traducao literal das formulas da planilha 3.6. Conferido
' registro a registro contra os 487 atendimentos reais em
' 16/09/2026: 487 de 487 iguais nos sete campos.
' Alterar aqui muda indicador do Painel - reconfira.
'==========================================================
Option Explicit

'----------------------------------------------------------
' Original: coluna C da aba Oportunidades
'----------------------------------------------------------
Public Function Situacao(ByVal temContato As Boolean, ByVal etapa As String, _
                         ByVal dtProxAcao As Variant, ByVal retomarEm As Variant) As String
    If Not temContato Then Exit Function
    If etapa = "" Then Situacao = "Sem etapa definida": Exit Function
    Situacao = SituacaoDaEtapa(modCRM.EtapaEncerra(etapa), dtProxAcao, retomarEm)
End Function

'----------------------------------------------------------
' Regra de prazo sem consultar a lista: quem chama diz se a etapa encerra.
' O autoteste usa esta funcao (ele nao acessa o banco).
'----------------------------------------------------------
Public Function SituacaoDaEtapa(ByVal encerrada As Boolean, ByVal dtProxAcao As Variant, _
                               ByVal retomarEm As Variant) As String
    Dim fimSemana As Date

    ' DEV-21 (ciclo 012): a agenda do aberto e so a proxima acao (retomar_em nao conta). Perdido ou
    ' Descartado cuja retomada chegou vai para avaliacao, sem reabrir o atendimento.
    If encerrada Then
        SituacaoDaEtapa = "Encerrado"
        If Not IsNull(retomarEm) Then
            If IsDate(retomarEm) Then
                If CDate(retomarEm) <= Date Then SituacaoDaEtapa = "Retomada a avaliar"
            End If
        End If
        Exit Function
    End If

    If IsNull(dtProxAcao) Then SituacaoDaEtapa = "Sem próxima ação": Exit Function
    If Not IsDate(dtProxAcao) Then SituacaoDaEtapa = "Sem próxima ação": Exit Function

    If CDate(dtProxAcao) < Date Then SituacaoDaEtapa = "Ação atrasada": Exit Function

    ' fim da semana corrente: domingo
    fimSemana = Date - Weekday(Date, vbMonday) + 7
    If CDate(dtProxAcao) <= fimSemana Then
        SituacaoDaEtapa = "Ação nesta semana"
    Else
        SituacaoDaEtapa = "Em dia"
    End If
End Function

'----------------------------------------------------------
Public Function Quadro(ByVal etapa As String) As String
    If etapa = "" Then Exit Function
    Quadro = QuadroDoCodigo(modCRM.EtapaQuadro(etapa))
End Function

' Texto do quadro a partir do codigo gravado na lista (PROSPECCAO ou FUNIL).
Public Function QuadroDoCodigo(ByVal codigo As String) As String
    If codigo = "PROSPECCAO" Then
        QuadroDoCodigo = "1. Prospecção"
    Else
        QuadroDoCodigo = "2. Funil comercial"
    End If
End Function

Public Function DiasParado(ByVal ultimaInteracao As Variant) As Variant
    If IsNull(ultimaInteracao) Then DiasParado = Null: Exit Function
    If Not IsDate(ultimaInteracao) Then DiasParado = Null: Exit Function
    DiasParado = CLng(Date - CDate(ultimaInteracao))
End Function

Public Function Ciclo(ByVal dtProposta As Variant, ByVal dtDesfecho As Variant) As Variant
    If IsNull(dtProposta) Or IsNull(dtDesfecho) Then Ciclo = Null: Exit Function
    If Not IsDate(dtProposta) Or Not IsDate(dtDesfecho) Then Ciclo = Null: Exit Function
    Ciclo = CLng(CDate(dtDesfecho) - CDate(dtProposta))
End Function

Public Function MesReferencia(ByVal dt As Variant) As Variant
    If IsNull(dt) Then MesReferencia = Null: Exit Function
    If Not IsDate(dt) Then MesReferencia = Null: Exit Function
    MesReferencia = Year(CDate(dt)) * 100 + Month(CDate(dt))
End Function

'----------------------------------------------------------
' Texto junto com a cor: vermelho nunca e o unico aviso.
'----------------------------------------------------------
Public Function CorDaSituacao(ByVal situacaoTexto As String) As Long
    ' cor de texto da situacao: a paleta unica fica no tema (modTema)
    CorDaSituacao = modTema.CorSituacao(situacaoTexto, "texto")
End Function
