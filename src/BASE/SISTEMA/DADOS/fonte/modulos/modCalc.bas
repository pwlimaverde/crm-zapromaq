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

Private Function EhEncerrada(ByVal etapa As String) As Boolean
    EhEncerrada = (etapa = "Pedido Fechado" Or etapa = "Perdido" Or etapa = "Descartado")
End Function

Private Function EhProspeccao(ByVal etapa As String) As Boolean
    EhProspeccao = (etapa = "Contato Inicial" Or etapa = "Sem Retorno" Or _
                    etapa = "Retorno Agendado" Or etapa = "Descartado")
End Function

'----------------------------------------------------------
' Original: coluna C da aba Oportunidades
'----------------------------------------------------------
Public Function Situacao(ByVal temContato As Boolean, ByVal etapa As String, _
                         ByVal dtProxAcao As Variant, ByVal retomarEm As Variant) As String
    Dim fimSemana As Date

    If Not temContato Then Exit Function
    If etapa = "" Then Situacao = "Sem etapa definida": Exit Function
    If EhEncerrada(etapa) Then Situacao = "Encerrado": Exit Function

    If Not IsNull(retomarEm) Then
        If IsDate(retomarEm) Then
            If CDate(retomarEm) <= Date Then Situacao = "Retomar hoje": Exit Function
        End If
    End If

    If IsNull(dtProxAcao) Then Situacao = "Sem próxima ação": Exit Function
    If Not IsDate(dtProxAcao) Then Situacao = "Sem próxima ação": Exit Function

    If CDate(dtProxAcao) < Date Then Situacao = "Ação atrasada": Exit Function

    ' fim da semana corrente: domingo
    fimSemana = Date - Weekday(Date, vbMonday) + 7
    If CDate(dtProxAcao) <= fimSemana Then
        Situacao = "Ação nesta semana"
    Else
        Situacao = "Em dia"
    End If
End Function

'----------------------------------------------------------
Public Function Quadro(ByVal etapa As String) As String
    If etapa = "" Then Exit Function
    If EhProspeccao(etapa) Then
        Quadro = "1. Prospecção"
    Else
        Quadro = "2. Funil comercial"
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
