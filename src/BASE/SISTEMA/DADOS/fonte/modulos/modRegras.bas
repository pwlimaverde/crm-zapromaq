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
