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
    ' a chave vai junto: o papel da etapa e lido pela chave e a montagem recusa lista sem chave.
    ' Etapa nova nasce aberta, no funil, sem familia (como no conector), com YESNO explicito.
    If tipo = "Etapa" Then
        id = modDB.InserirEm(cn, "INSERT INTO listas (tipo, valor, chave, ordem, ativo, encerra, quadro, exige_familia) " & _
                                 "VALUES (?,?,?,?,True,False,'FUNIL',False)", _
                             Array(modDB.P(modDB.adVarWChar, tipo), modDB.P(modDB.adVarWChar, valor), _
                                   modDB.P(modDB.adVarWChar, chave), modDB.P(modDB.adInteger, ordem + 1)))
    Else
        id = modDB.InserirEm(cn, "INSERT INTO listas (tipo, valor, chave, ordem, ativo) VALUES (?,?,?,?,True)", _
                             Array(modDB.P(modDB.adVarWChar, tipo), modDB.P(modDB.adVarWChar, valor), _
                                   modDB.P(modDB.adVarWChar, chave), modDB.P(modDB.adInteger, ordem + 1)))
    End If
    modDB.LogEm cn, "listas", id, "valor", "", tipo & ": " & valor, "INSERIR"
    modDB.Confirmar cn
    modCRM.LimparCacheListas
    Exit Function
Falha:
    ' a descricao antes do Desfazer: o On Error de dentro dele limpa o Err
    Dim descErro As String
    descErro = Err.Description
    modDB.Desfazer cn
    IncluirValor = "Nao gravou: " & descErro
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
    ' a descricao antes do Desfazer: o On Error de dentro dele limpa o Err
    Dim descErro As String
    descErro = Err.Description
    modDB.Desfazer cn
    DefinirAtivo = "Nao gravou: " & descErro
End Function

' Quantos atendimentos estao nesta etapa (o front mostra antes de pedir confirmacao).
Public Function AtendimentosDaEtapa(ByVal valor As String) As Long
    AtendimentosDaEtapa = CLng(modDB.ConsultarValor("SELECT COUNT(*) FROM oportunidades WHERE etapa=?", 0, _
                                                    Array(modDB.P(modDB.adVarWChar, valor))))
End Function

' Muda o papel de uma etapa (encerra, resultado, quadro, familia). C04: resultado
' so em etapa que encerra. C03: tirar o resultado de uma etapa ativa exige outra
' ativa com o mesmo resultado. Mudar 'encerra' muda a situacao dos atendimentos
' dessa etapa: sem confirmar, recusa com a contagem. Conflito: os valores antigos
' tem de ainda ser os lidos. Devolve "" quando grava (ou nada muda).
Public Function DefinirPapel(ByVal valor As String, ByVal encerra As Boolean, ByVal resultado As String, _
                             ByVal quadro As String, ByVal familia As Boolean, ByVal confirmar As Boolean) As String
    Dim cn As Object, d As Variant, i As Long, n As Long, id As Long, achou As Boolean
    Dim chave As String, valorCanon As String, chaves() As String, resultados() As String, ativos() As Boolean
    Dim encAnt As Boolean, resAnt As String, quadAnt As String, famAnt As Boolean, ativoAnt As Boolean
    Dim afetados As Long, filtro As String, sql As String

    resultado = UCase$(Trim$(resultado))
    quadro = UCase$(Trim$(quadro))
    DefinirPapel = PapelEtapaValido(encerra, resultado)
    If DefinirPapel <> "" Then Exit Function
    If quadro <> "FUNIL" And quadro <> "PROSPECCAO" Then DefinirPapel = "quadro deve ser FUNIL ou PROSPECCAO.": Exit Function
    chave = modValidacao.ChaveLista(valor)
    Set cn = modDB.AbrirTransacao()
    On Error GoTo Falha
    d = modDB.ConsultarEm(cn, "SELECT id, valor, encerra, resultado, quadro, exige_familia, ativo FROM listas WHERE tipo='Etapa'")
    If IsEmpty(d) Then modDB.Desfazer cn: DefinirPapel = "Lista Etapa vazia.": Exit Function
    ReDim chaves(1 To UBound(d, 1)): ReDim resultados(1 To UBound(d, 1)): ReDim ativos(1 To UBound(d, 1))
    For i = 1 To UBound(d, 1)
        chaves(i) = modValidacao.ChaveLista(CStr(d(i, 2)))
        If IsNull(d(i, 4)) Then resultados(i) = "" Else resultados(i) = CStr(d(i, 4))
        ativos(i) = CBool(d(i, 7))
        If chaves(i) = chave Then
            achou = True: id = CLng(d(i, 1)): valorCanon = CStr(d(i, 2))
            encAnt = CBool(d(i, 3)): resAnt = resultados(i): quadAnt = CStr(d(i, 5))
            famAnt = CBool(d(i, 6)): ativoAnt = ativos(i)
        End If
    Next i
    If Not achou Then modDB.Desfazer cn: DefinirPapel = "'" & valor & "' nao existe na lista Etapa.": Exit Function
    If encAnt = encerra And resAnt = resultado And quadAnt = quadro And famAnt = familia Then
        modDB.Desfazer cn
        DefinirPapel = ""
        Exit Function
    End If
    If resAnt <> "" And resultado <> resAnt And ativoAnt Then
        If Not DesativarEtapaPermitido(valorCanon, chaves, resultados, ativos) Then
            modDB.Desfazer cn
            DefinirPapel = "'" & valorCanon & "' e a ultima etapa ativa de " & resAnt & ": mude a outra antes (C03)."
            Exit Function
        End If
    End If
    If encerra <> encAnt Then
        afetados = AtendimentosDaEtapa(valorCanon)
        If afetados > 0 And Not confirmar Then
            modDB.Desfazer cn
            DefinirPapel = "Mudar 'encerra' de '" & valorCanon & "' muda a situacao de " & afetados & _
                           " atendimento(s). Confirme para gravar."
            Exit Function
        End If
    End If
    ' conflito: a linha tem de ter os valores antigos (a tabela nao tem coluna versao)
    filtro = " WHERE id=? AND encerra=? AND quadro=? AND exige_familia=?"
    If resAnt = "" Then
        filtro = filtro & " AND (resultado IS NULL OR resultado='')"
    Else
        filtro = filtro & " AND resultado=?"
    End If
    sql = "UPDATE listas SET encerra=?, resultado=?, quadro=?, exige_familia=?" & filtro
    Dim ps As Variant
    If resAnt = "" Then
        ps = Array(modDB.P(modDB.adBoolean, encerra), modDB.P(modDB.adVarWChar, IIf(resultado = "", Null, resultado)), _
                   modDB.P(modDB.adVarWChar, quadro), modDB.P(modDB.adBoolean, familia), _
                   modDB.P(modDB.adInteger, id), modDB.P(modDB.adBoolean, encAnt), _
                   modDB.P(modDB.adVarWChar, quadAnt), modDB.P(modDB.adBoolean, famAnt))
    Else
        ps = Array(modDB.P(modDB.adBoolean, encerra), modDB.P(modDB.adVarWChar, IIf(resultado = "", Null, resultado)), _
                   modDB.P(modDB.adVarWChar, quadro), modDB.P(modDB.adBoolean, familia), _
                   modDB.P(modDB.adInteger, id), modDB.P(modDB.adBoolean, encAnt), _
                   modDB.P(modDB.adVarWChar, quadAnt), modDB.P(modDB.adBoolean, famAnt), _
                   modDB.P(modDB.adVarWChar, resAnt))
    End If
    n = modDB.ExecutarEm(cn, sql, ps)
    If n <> 1 Then
        modDB.Desfazer cn
        DefinirPapel = "Conflito: '" & valorCanon & "' foi alterado por outra pessoa agora. Nada foi gravado; leia de novo."
        Exit Function
    End If
    If encerra <> encAnt Then modDB.LogEm cn, "listas", id, "encerra", CStr(encAnt), CStr(encerra), "ALTERAR"
    If resultado <> resAnt Then modDB.LogEm cn, "listas", id, "resultado", resAnt, resultado, "ALTERAR"
    If quadro <> quadAnt Then modDB.LogEm cn, "listas", id, "quadro", quadAnt, quadro, "ALTERAR"
    If familia <> famAnt Then modDB.LogEm cn, "listas", id, "exige_familia", CStr(famAnt), CStr(familia), "ALTERAR"
    modDB.Confirmar cn
    modCRM.LimparCacheListas
    Exit Function
Falha:
    ' a descricao antes do Desfazer: o On Error de dentro dele limpa o Err
    Dim descErro As String
    descErro = Err.Description
    modDB.Desfazer cn
    DefinirPapel = "Nao gravou: " & descErro
End Function

' Lista de tipos para o combo do gerenciamento.
Public Function TiposDeLista() As Variant
    TiposDeLista = modDB.Consultar("SELECT DISTINCT tipo FROM listas ORDER BY tipo")
End Function

' Valores de uma lista na ordem de exibicao (inativos tambem), com o papel.
' Colunas: id, valor, ordem, ativo, encerra, resultado, quadro, exige_familia.
Public Function LinhasDaLista(ByVal tipo As String) As Variant
    LinhasDaLista = modDB.Consultar("SELECT id, valor, ordem, ativo, encerra, resultado, quadro, exige_familia " & _
                                    "FROM listas WHERE tipo=? ORDER BY ordem, valor", _
                                    Array(modDB.P(modDB.adVarWChar, tipo)))
End Function

' Quantos registros usam o valor (mesmo mapa de Get-ColunasDaLista do build).
Public Function UsoDoValor(ByVal tipo As String, ByVal valor As String) As Long
    Dim alvos As String, a As Variant, partes As Variant
    Select Case tipo
        Case "UF": alvos = "clientes.uf"
        Case "Segmento": alvos = "clientes.segmento"
        Case "Qualificacao": alvos = "clientes.qualificacao"
        Case "Responsavel": alvos = "clientes.responsavel|oportunidades.responsavel"
        Case "Etapa": alvos = "oportunidades.etapa"
        Case "Familia": alvos = "oportunidades.familia"
        Case "Categoria": alvos = "oportunidades.categoria"
        Case "Origem": alvos = "oportunidades.origem"
        Case "Prioridade": alvos = "oportunidades.prioridade"
        Case "Motivo": alvos = "oportunidades.motivo_desfecho"
    End Select
    If alvos = "" Then Exit Function
    For Each a In Split(alvos, "|")
        partes = Split(a, ".")
        UsoDoValor = UsoDoValor + modDB.ConsultarValor("SELECT COUNT(*) FROM " & partes(0) & " WHERE " & partes(1) & "=?", 0, _
                                                       Array(modDB.P(modDB.adVarWChar, valor)))
    Next a
End Function

' Sobe ou desce um valor na ordem de exibicao. Caso comum (ordens diferentes): troca
' a ordem so entre as duas linhas. Empate: renumera a lista 1..n. Uma transacao; cada
' ordem alterada vai ao log.
Public Function TrocarOrdem(ByVal tipo As String, ByVal valor As String, ByVal subir As Boolean) As String
    Dim cn As Object, d As Variant, i As Long, k As Long, n As Long, pos As Long, viz As Long, n2 As Long
    Dim ids() As Long, antig() As Long, nomes() As String, troca As Long, trocaS As String
    Dim oPos As Long, oViz As Long
    Set cn = modDB.AbrirTransacao()
    On Error GoTo Falha
    d = modDB.ConsultarEm(cn, "SELECT id, valor, ordem FROM listas WHERE tipo=? ORDER BY ordem, valor", _
                          Array(modDB.P(modDB.adVarWChar, tipo)))
    If IsEmpty(d) Then modDB.Desfazer cn: TrocarOrdem = "Lista " & tipo & " vazia.": Exit Function
    n = UBound(d, 1)
    ReDim ids(1 To n): ReDim antig(1 To n): ReDim nomes(1 To n)
    For i = 1 To n
        ids(i) = CLng(d(i, 1)): nomes(i) = CStr(d(i, 2)): antig(i) = CLng(d(i, 3))
        If modValidacao.ChaveLista(nomes(i)) = modValidacao.ChaveLista(valor) Then pos = i
    Next i
    If pos = 0 Then modDB.Desfazer cn: TrocarOrdem = "'" & valor & "' nao existe na lista " & tipo & ".": Exit Function
    If subir Then viz = pos - 1 Else viz = pos + 1
    If viz < 1 Or viz > n Then
        modDB.Desfazer cn
        If subir Then TrocarOrdem = "Ja esta no topo da lista." Else TrocarOrdem = "Ja esta no fim da lista."
        Exit Function
    End If
    oPos = antig(pos): oViz = antig(viz)
    If oPos <> oViz Then
        n2 = modDB.ExecutarEm(cn, "UPDATE listas SET ordem=? WHERE id=? AND ordem=?", _
                              Array(modDB.P(modDB.adInteger, oViz), modDB.P(modDB.adInteger, ids(pos)), modDB.P(modDB.adInteger, oPos)))
        If n2 = 1 Then n2 = modDB.ExecutarEm(cn, "UPDATE listas SET ordem=? WHERE id=? AND ordem=?", _
                              Array(modDB.P(modDB.adInteger, oPos), modDB.P(modDB.adInteger, ids(viz)), modDB.P(modDB.adInteger, oViz)))
        If n2 <> 1 Then
            modDB.Desfazer cn
            TrocarOrdem = "Conflito: a lista mudou agora. Nada foi gravado; leia de novo."
            Exit Function
        End If
        modDB.LogEm cn, "listas", ids(pos), "ordem", CStr(oPos), CStr(oViz), "ALTERAR"
        modDB.LogEm cn, "listas", ids(viz), "ordem", CStr(oViz), CStr(oPos), "ALTERAR"
    Else
        troca = ids(pos): ids(pos) = ids(viz): ids(viz) = troca
        troca = antig(pos): antig(pos) = antig(viz): antig(viz) = troca
        trocaS = nomes(pos): nomes(pos) = nomes(viz): nomes(viz) = trocaS
        For k = 1 To n
            If antig(k) <> k Then
                n2 = modDB.ExecutarEm(cn, "UPDATE listas SET ordem=? WHERE id=? AND ordem=?", _
                                      Array(modDB.P(modDB.adInteger, k), modDB.P(modDB.adInteger, ids(k)), _
                                            modDB.P(modDB.adInteger, antig(k))))
                If n2 <> 1 Then
                    modDB.Desfazer cn
                    TrocarOrdem = "Conflito: a lista mudou agora. Nada foi gravado; leia de novo."
                    Exit Function
                End If
                modDB.LogEm cn, "listas", ids(k), "ordem", CStr(antig(k)), CStr(k), "ALTERAR"
            End If
        Next k
    End If
    modDB.Confirmar cn
    modCRM.LimparCacheListas
    Exit Function
Falha:
    ' a descricao antes do Desfazer: o On Error de dentro dele limpa o Err
    Dim descErro As String
    descErro = Err.Description
    modDB.Desfazer cn
    TrocarOrdem = "Nao gravou: " & descErro
End Function
