Attribute VB_Name = "modAcoes"
'==========================================================
' modAcoes - ativar, desativar e excluir registro.
'
' Chamado pela FICHA, nao pela grade: e na ficha que se ve o
' registro inteiro antes de decidir.
'
' Regras que vieram do comercial, nao do codigo:
'
'  1. PRE-CLIENTE nao tem contato nem oportunidade. Exclui de
'     vez. Mesmo assim a rotina CONFERE dependente antes, na
'     mesma transacao do DELETE.
'
'  2. CLIENTE nunca e excluido: desativa, e os contatos
'     vinculados desativam em cascata na MESMA transacao.
'
'  3. As OPORTUNIDADES EM ABERTO de um cliente desativado
'     FICAM COMO ESTAO - decisao do comercial em 17/09/2026.
'     O aviso diz quantas sao, para a decisao ser consciente.
'
'  4. Contato, na grade dele, so inativa. Oportunidade nao se
'     desativa: encerra-se pela etapa.
'
' Toda mudanca de ativo sobe a versao do registro (quem estiver
' com a ficha aberta para editar recebe o aviso de conflito) e
' grava log de CADA registro afetado, na mesma transacao - o log
' e tambem o que faz as outras estacoes atualizarem a grade.
'==========================================================
Option Explicit

'----------------------------------------------------------
' O QUE A FICHA PERGUNTA ANTES DE MOSTRAR O BOTAO
'   pre-cliente ....... Excluir   (nao tem dependente)
'   registro ativo .... Desativar
'   registro inativo .. Reativar
'----------------------------------------------------------
Public Function RotuloStatus(ByVal tabela As String, ByVal idRegistro As Long, _
                             ByVal ativo As Boolean) As String
    If idRegistro = 0 Then RotuloStatus = "Desativar": Exit Function
    If tabela = "clientes" Then
        If EhPreCliente(idRegistro) Then RotuloStatus = "Excluir": Exit Function
    End If
    If ativo Then RotuloStatus = "Desativar" Else RotuloStatus = "Reativar"
End Function

'----------------------------------------------------------
' Executa o que o rotulo promete. Devolve True se mexeu no
' banco - a ficha recarrega a lista so nesse caso.
'----------------------------------------------------------
Public Function AcaoStatus(ByVal tabela As String, ByVal idRegistro As Long, _
                           ByVal ativo As Boolean) As Boolean
    If idRegistro = 0 Then Exit Function
    If tabela = "oportunidades" Then Exit Function
    If Not modMenu.Pronto() Then Exit Function

    On Error GoTo erro
    If tabela = "clientes" Then
        If EhPreCliente(idRegistro) Then
            AcaoStatus = ExcluirPreCliente(idRegistro)
        ElseIf ativo Then
            AcaoStatus = DesativarCliente(idRegistro)
        Else
            AcaoStatus = ReativarCliente(idRegistro)
        End If
    Else
        AcaoStatus = AlternarSimples(tabela, idRegistro, ativo)
    End If
    Exit Function
erro:
    MsgBox "Não foi possível concluir a operação. Nada foi alterado." & vbCrLf & vbCrLf & _
           Err.Description, vbExclamation, "CRM Zapromaq"
End Function

'----------------------------------------------------------
' Contato: sem cascata.
'----------------------------------------------------------
Private Function AlternarSimples(ByVal tabela As String, ByVal idRegistro As Long, _
                                 ByVal ativo As Boolean) As Boolean
    Dim msg As String, cn As Object
    If ativo Then
        msg = "Desativar este registro? Ele sai das listas de lançamento," & vbCrLf & _
              "mas continua no histórico e nas consultas."
    Else
        msg = "Reativar este registro? Ele volta a aparecer nas listas de lançamento."
    End If
    If MsgBox(msg, vbYesNo + vbQuestion, "CRM Zapromaq") <> vbYes Then Exit Function

    On Error GoTo falha
    Set cn = modDB.AbrirTransacao()
    If MudarAtivo(cn, tabela, "id=?", idRegistro, Not ativo, IIf(ativo, "DESATIVACAO", "REATIVACAO"), "") = 0 Then
        modDB.Desfazer cn
        MsgBox "O registro já estava " & IIf(ativo, "inativo", "ativo") & ". Atualize a lista.", vbInformation, "CRM Zapromaq"
        Exit Function
    End If
    modDB.Confirmar cn
    AlternarSimples = True
    Exit Function
falha:
    Dim nErr As Long, sErr As String
    nErr = Err.Number: sErr = Err.Description
    modDB.Desfazer cn
    Err.Raise nErr, "modAcoes.AlternarSimples", sErr
End Function

'----------------------------------------------------------
' Reativar cliente traz junto os contatos que estao inativos.
' Sem isso o cliente volta sem ninguem para ligar, e a falta
' so aparece na hora de lancar atendimento.
'----------------------------------------------------------
Public Function ReativarCliente(ByVal idCliente As Long) As Boolean
    Dim nCt As Long, cn As Object

    nCt = Contar("SELECT COUNT(*) FROM contatos WHERE id_cliente=? AND ativo=False", idCliente)
    If MsgBox("Reativar este cliente?" & vbCrLf & vbCrLf & TextoDoCliente(idCliente) & vbCrLf & vbCrLf & _
              IIf(nCt > 0, "Os " & nCt & " contato(s) inativos dele voltam junto.", _
                           "Ele não tem contato inativo para voltar junto."), _
              vbYesNo + vbQuestion, "CRM Zapromaq") <> vbYes Then Exit Function

    On Error GoTo falha
    Set cn = modDB.AbrirTransacao()
    MudarAtivo cn, "clientes", "id=?", idCliente, True, "REATIVACAO", ""
    nCt = MudarAtivo(cn, "contatos", "id_cliente=?", idCliente, True, "REATIVACAO", "cascata da reativação do cliente")
    modDB.Confirmar cn
    ReativarCliente = True
    Exit Function
falha:
    Dim nErr As Long, sErr As String
    nErr = Err.Number: sErr = Err.Description
    modDB.Desfazer cn
    Err.Raise nErr, "modAcoes.ReativarCliente", sErr
End Function

'==========================================================
' DESATIVACAO DE CLIENTE EM CASCATA
' Contatos vao junto, na mesma transacao. Oportunidades ficam.
'==========================================================
Public Function DesativarCliente(ByVal idCliente As Long) As Boolean
    Dim nCt As Long, nOp As Long, aviso As String, cn As Object

    nCt = Contar("SELECT COUNT(*) FROM contatos WHERE id_cliente=? AND ativo=True", idCliente)
    nOp = Contar("SELECT COUNT(*) FROM oportunidades WHERE id_cliente=?" & _
                 " AND etapa NOT IN ('Pedido Fechado','Perdido','Descartado')", idCliente)

    aviso = "Desativar este cliente?" & vbCrLf & vbCrLf & TextoDoCliente(idCliente) & vbCrLf & vbCrLf & _
            "Junto com ele são desativados " & nCt & " contato(s)."
    If nOp > 0 Then
        aviso = aviso & vbCrLf & vbCrLf & _
                "ATENÇÃO: " & nOp & " oportunidade(s) em aberto CONTINUAM abertas e seguem" & _
                " somando no funil. Encerre uma por uma se não for mais trabalhar nelas."
    End If
    If MsgBox(aviso, vbYesNo + vbQuestion, "CRM Zapromaq") <> vbYes Then Exit Function

    On Error GoTo falha
    Set cn = modDB.AbrirTransacao()
    MudarAtivo cn, "clientes", "id=?", idCliente, False, "DESATIVACAO", _
               nOp & " oportunidade(s) em aberto mantida(s)"
    nCt = MudarAtivo(cn, "contatos", "id_cliente=?", idCliente, False, "DESATIVACAO", "cascata da desativação do cliente")
    modDB.Confirmar cn

    MsgBox "Cliente desativado." & vbCrLf & nCt & " contato(s) desativado(s) em cascata." & _
           IIf(nOp > 0, vbCrLf & nOp & " oportunidade(s) em aberto permanecem no funil.", ""), _
           vbInformation, "CRM Zapromaq"
    DesativarCliente = True
    Exit Function
falha:
    Dim nErr As Long, sErr As String
    nErr = Err.Number: sErr = Err.Description
    modDB.Desfazer cn
    Err.Raise nErr, "modAcoes.DesativarCliente", sErr
End Function

'==========================================================
' EXCLUSAO DEFINITIVA DE PRE-CLIENTE
' Conferencia de dependentes, DELETE e log na mesma transacao:
' o log so e gravado se a exclusao aconteceu de fato.
'==========================================================
Public Function ExcluirPreCliente(ByVal idCliente As Long) As Boolean
    Dim nCt As Long, nOp As Long, texto As String, cn As Object, n As Long

    texto = TextoDoCliente(idCliente)
    If MsgBox("Excluir DE VEZ este pré-cliente?" & vbCrLf & vbCrLf & texto & vbCrLf & vbCrLf & _
              "Não há como desfazer pelo sistema - só pelo backup.", _
              vbYesNo + vbExclamation + vbDefaultButton2, "CRM Zapromaq") <> vbYes Then Exit Function

    On Error GoTo falha
    Set cn = modDB.AbrirTransacao()
    nCt = CLng(modDB.ConsultarValorEm(cn, "SELECT COUNT(*) FROM contatos WHERE id_cliente=?", 0, _
                                      Array(modDB.P(modDB.adInteger, idCliente))))
    nOp = CLng(modDB.ConsultarValorEm(cn, "SELECT COUNT(*) FROM oportunidades WHERE id_cliente=?", 0, _
                                      Array(modDB.P(modDB.adInteger, idCliente))))
    If nCt > 0 Or nOp > 0 Then
        modDB.Desfazer cn
        MsgBox "Este pré-cliente tem registro vinculado e por isso não será excluído:" & vbCrLf & _
               "  contatos: " & nCt & vbCrLf & "  oportunidades: " & nOp & vbCrLf & vbCrLf & _
               "Pré-cliente não deveria ter vínculo. Confira o cadastro antes - " & _
               "provavelmente ele deveria estar como Cliente.", vbExclamation, "CRM Zapromaq"
        Exit Function
    End If
    n = modDB.ExecutarEm(cn, "DELETE FROM clientes WHERE id=? AND codigo_cliente IS NULL", _
                         Array(modDB.P(modDB.adInteger, idCliente)))
    If n <> 1 Then
        modDB.Desfazer cn
        MsgBox "O registro não foi excluído: ele não existe mais ou já virou cliente. Atualize a lista.", _
               vbInformation, "CRM Zapromaq"
        Exit Function
    End If
    modDB.LogEm cn, "clientes", idCliente, "(registro)", texto, "", "EXCLUSAO"
    modDB.Confirmar cn
    ExcluirPreCliente = True
    Exit Function
falha:
    Dim nErr As Long, sErr As String
    nErr = Err.Number: sErr = Err.Description
    modDB.Desfazer cn
    Err.Raise nErr, "modAcoes.ExcluirPreCliente", sErr
End Function

'----------------------------------------------------------
' Muda ativo dos registros que casam com o filtro (e ainda nao
' estao no estado novo), sobe a versao e grava UM log por
' registro. Devolve quantos mudaram.
'----------------------------------------------------------
Private Function MudarAtivo(ByVal cn As Object, ByVal tabela As String, ByVal filtro As String, _
                            ByVal valorFiltro As Long, ByVal novo As Boolean, ByVal acao As String, _
                            ByVal detalhe As String) As Long
    Dim ids As Variant, i As Long, n As Long
    ids = modDB.ConsultarEm(cn, "SELECT id FROM " & tabela & " WHERE " & filtro & " AND ativo=?", _
                            Array(modDB.P(modDB.adInteger, valorFiltro), modDB.P(modDB.adBoolean, Not novo)))
    If IsEmpty(ids) Then Exit Function
    For i = 1 To UBound(ids, 1)
        n = n + modDB.ExecutarEm(cn, "UPDATE " & tabela & " SET ativo=?, versao=versao+1, alterado_em=?," & _
                                     " alterado_por=? WHERE id=? AND ativo=?", _
            Array(modDB.P(modDB.adBoolean, novo), modDB.P(modDB.adDate, Now), _
                  modDB.P(modDB.adVarWChar, modConfig.UsuarioAtual()), _
                  modDB.P(modDB.adInteger, CLng(ids(i, 1))), modDB.P(modDB.adBoolean, Not novo)))
        modDB.LogEm cn, tabela, CLng(ids(i, 1)), "ativo", IIf(novo, "False", "True"), _
                    IIf(novo, "True", "False") & IIf(detalhe = "", "", " (" & detalhe & ")"), acao
    Next i
    MudarAtivo = n
End Function

'----------------------------------------------------------
Private Function EhPreCliente(ByVal idCliente As Long) As Boolean
    Dim v As Variant
    v = modDB.ConsultarValor("SELECT codigo_cliente FROM clientes WHERE id=?", Null, _
                             Array(modDB.P(modDB.adInteger, idCliente)))
    ' pre-cliente e quem nao tem codigo (o estagio deriva do codigo)
    EhPreCliente = IsNull(v)
End Function

Private Function TextoDoCliente(ByVal idCliente As Long) As String
    TextoDoCliente = modCRM.ResumoCliente(idCliente)
End Function

Private Function Contar(ByVal sql As String, ByVal idCliente As Long) As Long
    Dim v As Variant
    v = modDB.ConsultarValor(sql, 0, Array(modDB.P(modDB.adInteger, idCliente)))
    If IsNumeric(v) Then Contar = CLng(v)
End Function
