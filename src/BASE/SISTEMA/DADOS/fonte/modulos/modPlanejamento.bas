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
