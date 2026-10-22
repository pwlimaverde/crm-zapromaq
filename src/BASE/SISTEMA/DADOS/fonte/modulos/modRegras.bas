Attribute VB_Name = "modRegras"
Option Explicit
'==========================================================
' modRegras - regras de etapa da secao 5.2 do plano de
' 19/10/2026 (ciclo 009, DEV-10), iguais as do conector
' (Get-TransicaoEtapa e Test-RetomadaPermitida, Regras.ps1).
' Casos C25 a C27 nos dois testes. Sem banco: quem chama
' passa o papel das etapas (encerra, resultado PERDA).
'
'   encerrar em Perdido ou Descartado (encerra + PERDA) limpa a
'   proxima acao e a data dela;
'   reabrir (de etapa que encerra para etapa aberta) limpa
'   motivo, data do desfecho e retomada, e devolve a orientacao;
'   retomar_em so existe em Perdido e Descartado: gravar a data
'   (nova ou alterada) em outra etapa e recusado.
' Os valores anteriores vao para o log pela gravacao normal
' (o campo que muda tem log, antigo e novo).
'==========================================================

Public Function OrientacaoReabertura() As String
    OrientacaoReabertura = "Reabrir serve para corrigir um encerramento indevido: motivo, data do desfecho e " & _
        "retomada ser" & ChrW$(227) & "o limpos (os valores anteriores ficam no hist" & ChrW$(243) & "rico)." & vbCrLf & vbCrLf & _
        "Interesse novo " & ChrW$(233) & " atendimento NOVO: use Duplicar no atendimento encerrado."
End Function

' Aplica a transicao em d (dicionario do que vai ser gravado). Devolve a
' orientacao quando e reabertura ("" nos outros casos).
Public Function TransicaoEtapa(ByVal encerraAntiga As Boolean, ByVal encerraNova As Boolean, _
                               ByVal perdaNova As Boolean, ByVal d As Object) As String
    If encerraNova And perdaNova Then
        If d.Exists("prox_acao") Then d("prox_acao") = Null
        If d.Exists("dt_prox_acao") Then d("dt_prox_acao") = Null
    End If
    If encerraAntiga And Not encerraNova Then
        If d.Exists("motivo_desfecho") Then d("motivo_desfecho") = Null
        If d.Exists("dt_desfecho") Then d("dt_desfecho") = Null
        If d.Exists("retomar_em") Then d("retomar_em") = Null
        TransicaoEtapa = OrientacaoReabertura()
    End If
End Function

' C26: retomar_em nova ou alterada fora de Perdido/Descartado e recusada.
Public Function CriticarRetomada(ByVal encerraNova As Boolean, ByVal perdaNova As Boolean, _
                                 ByVal retomarNova As Variant, ByVal retomarLida As Variant) As String
    If IsNull(retomarNova) Or IsEmpty(retomarNova) Then Exit Function
    If encerraNova And perdaNova Then Exit Function
    If Not (IsNull(retomarLida) Or IsEmpty(retomarLida)) Then
        If IsDate(retomarLida) Then If CDate(retomarLida) = CDate(retomarNova) Then Exit Function
    End If
    CriticarRetomada = "A data de retomada s" & ChrW$(243) & " existe em atendimento Perdido ou Descartado. " & _
        "Para um atendimento aberto, use a pr" & ChrW$(243) & "xima a" & ChrW$(231) & ChrW$(227) & "o (data da pr" & ChrW$(243) & "xima a" & ChrW$(231) & ChrW$(227) & "o)."
End Function


' C29 (DEV-11): o contato do atendimento existe, esta ativo e e da empresa
' do atendimento. Mesma regra de Test-ContatoDoAtendimento (conector).
Public Function CriticarContatoDoAtendimento(ByVal existe As Boolean, ByVal ativo As Boolean, _
                                             ByVal idClienteContato As Long, ByVal idClienteAtendimento As Long) As String
    If Not existe Then CriticarContatoDoAtendimento = "Contato n" & ChrW$(227) & "o encontrado: nada foi gravado.": Exit Function
    If idClienteContato <> idClienteAtendimento Then
        CriticarContatoDoAtendimento = "O contato " & ChrW$(233) & " de outra empresa: o atendimento s" & ChrW$(243) & " usa contato da empresa dele. Nada foi gravado."
        Exit Function
    End If
    If Not ativo Then CriticarContatoDoAtendimento = "O contato est" & ChrW$(225) & " inativo: escolha um contato ativo da empresa. Nada foi gravado."
End Function


' Mesmo valor, para conflito por campo (DEV-15): nulo = vazio; numero e data
' comparados como tal; texto exato.
Public Function MesmoValor(ByVal a As Variant, ByVal b As Variant) As Boolean
    Dim va As Boolean, vb As Boolean
    va = IsNull(a) Or IsEmpty(a): If Not va Then va = (Len(CStr(a)) = 0)
    vb = IsNull(b) Or IsEmpty(b): If Not vb Then vb = (Len(CStr(b)) = 0)
    If va Or vb Then MesmoValor = (va And vb): Exit Function
    If VarType(a) = vbDate Or VarType(b) = vbDate Then
        If IsDate(a) And IsDate(b) Then MesmoValor = (CDate(a) = CDate(b)): Exit Function
    End If
    If IsNumeric(a) And IsNumeric(b) And VarType(a) <> vbString And VarType(b) <> vbString Then MesmoValor = (CDbl(a) = CDbl(b)): Exit Function
    MesmoValor = (StrComp(CStr(a), CStr(b), vbBinaryCompare) = 0)
End Function


' C33 (DEV-16): o atendimento anterior forma ciclo? Segue a corrente de
' anteriores a partir do novo anterior (ids() e anteriores() em paralelo,
' os atendimentos da empresa); chegar no proprio atendimento = ciclo.
Public Function FormaCiclo(ByVal idAtual As Long, ByVal idAnterior As Long, ByVal ids As Variant, ByVal anteriores As Variant) As Boolean
    Dim p As Long, passos As Long, i As Long, achou As Boolean
    If idAtual = 0 Or idAnterior = 0 Then Exit Function
    p = idAnterior
    Do While p <> 0 And passos < 10000
        If p = idAtual Then FormaCiclo = True: Exit Function
        achou = False
        If IsArray(ids) Then
            For i = LBound(ids) To UBound(ids)
                If CLng(ids(i)) = p Then p = CLng(anteriores(i)): achou = True: Exit For
            Next i
        End If
        If Not achou Then Exit Function
        passos = passos + 1
    Loop
End Function
