Attribute VB_Name = "modListas"
Option Explicit
'==========================================================
' modListas - gerenciamento das listas suspensas (Fase 2).
'
' Regras puras (UFValida, PapelEtapaValido, DesativarEtapaPermitido)
' espelham IA\mcp\lib\Regras.ps1: a mesma regra nos dois lados
' (casos C02 a C05 da spec 003).
'
' Gravacoes (IncluirValor, DefinirAtivo, DefinirOrdem): uma transacao
' por acao, com LogEm na mesma transacao. Devolvem "" quando gravam e
' a mensagem da recusa quando nao gravam (nada fica pela metade).
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

' Incluir um valor novo no fim da lista. C02: chave ja existente recusa;
' C05: UF so das 27. Valor nasce ativo.
Public Function IncluirValor(ByVal tipo As String, ByVal valor As String) As String
    Dim cn As Object, d As Variant, i As Long, ordem As Long, id As Long
    Dim chave As String, ativo As Boolean
    valor = Trim$(valor)
    If valor = "" Then IncluirValor = "Informe o valor.": Exit Function
    If Len(valor) > 120 Then IncluirValor = "Valor aceita no maximo 120 caracteres.": Exit Function
    If tipo = "UF" And Not UFValida(valor) Then IncluirValor = "'" & valor & "' nao e uma UF do Brasil (C05).": Exit Function
    chave = modValidacao.ChaveLista(valor)
    Set cn = modDB.AbrirTransacao()
    On Error GoTo Falha
    d = modDB.ConsultarEm(cn, "SELECT valor, ativo, ordem FROM listas WHERE tipo=? ORDER BY ordem", _
                          Array(modDB.P(modDB.adVarWChar, tipo)))
    If Not IsEmpty(d) Then
        For i = 1 To UBound(d, 1)
            If modValidacao.ChaveLista(CStr(d(i, 1))) = chave Then
                ativo = CBool(d(i, 2))
                modDB.Desfazer cn
                IncluirValor = "'" & valor & "' ja existe na lista " & tipo & IIf(ativo, "", " (inativo: use Reativar)") & "."
                Exit Function
            End If
            If CLng(d(i, 3)) >= ordem Then ordem = CLng(d(i, 3))
        Next i
    End If
    id = modDB.InserirEm(cn, "INSERT INTO listas (tipo, valor, ordem, ativo) VALUES (?,?,?,True)", _
                         Array(modDB.P(modDB.adVarWChar, tipo), modDB.P(modDB.adVarWChar, valor), _
                               modDB.P(modDB.adInteger, ordem + 1)))
    modDB.LogEm cn, "listas", id, "valor", "", tipo & ": " & valor, "INSERIR"
    modDB.Confirmar cn
    modCRM.LimparCacheListas
    Exit Function
Falha:
    modDB.Desfazer cn
    IncluirValor = "Nao gravou: " & Err.Description
End Function

' Desativar ou reativar um valor. C03: nao desativa a ultima etapa ativa
' de ganho ou de perda. Conflito: o estado antigo tem de ainda ser o lido.
Public Function DefinirAtivo(ByVal tipo As String, ByVal valor As String, ByVal ativar As Boolean) As String
    Dim cn As Object, d As Variant, i As Long, n As Long, id As Long, achou As Boolean
    Dim chave As String, ativoAntigo As Boolean, chaves() As String, resultados() As String, ativos() As Boolean
    Dim resDoAlvo As String
    chave = modValidacao.ChaveLista(valor)
    Set cn = modDB.AbrirTransacao()
    On Error GoTo Falha
    d = modDB.ConsultarEm(cn, "SELECT id, valor, ativo, resultado FROM listas WHERE tipo=?", _
                          Array(modDB.P(modDB.adVarWChar, tipo)))
    If IsEmpty(d) Then modDB.Desfazer cn: DefinirAtivo = "Lista " & tipo & " nao existe.": Exit Function
    ReDim chaves(1 To UBound(d, 1)): ReDim resultados(1 To UBound(d, 1)): ReDim ativos(1 To UBound(d, 1))
    For i = 1 To UBound(d, 1)
        chaves(i) = modValidacao.ChaveLista(CStr(d(i, 2)))
        ativos(i) = CBool(d(i, 3))
        If IsNull(d(i, 4)) Then resultados(i) = "" Else resultados(i) = CStr(d(i, 4))
        If chaves(i) = chave Then
            achou = True: id = CLng(d(i, 1)): ativoAntigo = ativos(i): resDoAlvo = resultados(i)
            valor = CStr(d(i, 2))
        End If
    Next i
    If Not achou Then modDB.Desfazer cn: DefinirAtivo = "'" & valor & "' nao existe na lista " & tipo & ".": Exit Function
    If ativoAntigo = ativar Then
        modDB.Desfazer cn
        DefinirAtivo = "'" & valor & "' ja esta no estado pedido."
        Exit Function
    End If
    If Not ativar And tipo = "Etapa" Then
        If Not DesativarEtapaPermitido(valor, chaves, resultados, ativos) Then
            modDB.Desfazer cn
            DefinirAtivo = "'" & valor & "' e a ultima etapa ativa de " & _
                           resDoAlvo & ": sem ela o funil nao fecha. Ative outra antes (C03)."
            Exit Function
        End If
    End If
    n = modDB.ExecutarEm(cn, "UPDATE listas SET ativo=? WHERE id=? AND ativo=?", _
                         Array(modDB.P(modDB.adBoolean, ativar), modDB.P(modDB.adInteger, id), _
                               modDB.P(modDB.adBoolean, ativoAntigo)))
    If n <> 1 Then
        modDB.Desfazer cn
        DefinirAtivo = "Conflito: '" & valor & "' foi alterado por outra pessoa agora. Nada foi gravado; leia de novo."
        Exit Function
    End If
    If ativar Then
        modDB.LogEm cn, "listas", id, "ativo", "False", "True", "REATIVACAO"
    Else
        modDB.LogEm cn, "listas", id, "ativo", "True", "False", "DESATIVACAO"
    End If
    modDB.Confirmar cn
    modCRM.LimparCacheListas
    Exit Function
Falha:
    modDB.Desfazer cn
    DefinirAtivo = "Nao gravou: " & Err.Description
End Function
