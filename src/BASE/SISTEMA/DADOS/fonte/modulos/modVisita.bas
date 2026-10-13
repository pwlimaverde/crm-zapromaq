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
