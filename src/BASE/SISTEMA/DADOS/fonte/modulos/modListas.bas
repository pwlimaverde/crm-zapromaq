Attribute VB_Name = "modListas"
Option Explicit
'==========================================================
' modListas - regras do gerenciamento de listas (Fase 2).
'
' Espelho de IA\mcp\lib\Regras.ps1 (Test-UFValida,
' Test-PapelEtapaValido, Test-DesativarEtapaPermitido): a mesma
' regra nos dois lados. Casos C03, C04 e C05 de spec 003.
' Este modulo so tem funcoes puras; quem grava e modListas
' (fatia 2.2).
'==========================================================

Private Const UFS_BRASIL As String = "AC AL AP AM BA CE DF ES GO MA MT MS MG PA PB PR PE PI RJ RN RS RO RR SC SP SE TO"

' C05: UF so entra se for uma das 27 (maiuscula, sem espaco sobrando).
Public Function UFValida(ByVal uf As String) As Boolean
    If Len(uf) <> 2 Then Exit Function
    UFValida = (InStr(1, " " & UFS_BRASIL & " ", " " & UCase$(uf) & " ", vbBinaryCompare) > 0)
End Function

' C04: resultado (GANHO/PERDA) so existe em etapa que encerra.
' Devolve a mensagem de recusa, ou vazio quando o papel e valido.
Public Function PapelEtapaValido(ByVal encerra As Boolean, ByVal resultado As String) As String
    If resultado <> "" And resultado <> "GANHO" And resultado <> "PERDA" Then
        PapelEtapaValido = "Resultado deve ser GANHO, PERDA ou vazio."
    ElseIf resultado <> "" And Not encerra Then
        PapelEtapaValido = "Resultado so pode ser GANHO ou PERDA em etapa que encerra o atendimento (C04)."
    End If
End Function

' C03: desativar uma etapa com resultado so e permitido se sobrar outra
' ATIVA com o mesmo resultado. chaves = ChaveLista de cada etapa;
' resultados e ativos alinhados a ela.
Public Function DesativarEtapaPermitido(ByVal valor As String, ByVal chaves As Variant, _
                                        ByVal resultados As Variant, ByVal ativos As Variant) As Boolean
    Dim chave As String, resAlvo As String, achou As Boolean, i As Long
    chave = modValidacao.ChaveLista(valor)
    DesativarEtapaPermitido = True
    For i = LBound(chaves) To UBound(chaves)
        If chaves(i) = chave Then
            resAlvo = resultados(i)
            achou = True
            Exit For
        End If
    Next i
    If Not achou Then Exit Function
    If resAlvo = "" Then Exit Function
    For i = LBound(chaves) To UBound(chaves)
        If chaves(i) <> chave And ativos(i) And resultados(i) = resAlvo Then Exit Function
    Next i
    DesativarEtapaPermitido = False
End Function
