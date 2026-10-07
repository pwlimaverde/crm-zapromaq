Attribute VB_Name = "modCRM"
'==========================================================
' modCRM - regras de negocio
'
' Codigos, promocao de pre-cliente, gravacao com controle de
' versao e as consultas das grades. Regras em
' docs\contrato-banco.md.
'
' O codigo visual (CT-CCCC-NNNN / AT-CCCC-NNNN) e CONGELADO
' no ato da criacao e gravado como campo. Nunca se recalcula:
' ele nomeia pasta em 02 - CLIENTES, e recalcular renomearia
' o atendimento quando o contato troca de empresa.
'
' Toda gravacao acontece numa TRANSACAO com o log junto
' (modDB.AbrirTransacao / LogEm). Numero sequencial (controle,
' codigo_cliente) e calculado DENTRO da transacao que insere;
' se dois usuarios pegarem o mesmo MAX+1, o indice unico
' recusa o segundo e a rotina repete com o numero seguinte,
' em vez de mostrar erro.
'==========================================================
Option Explicit

Private Const TENTATIVAS As Long = 5

' Listas suspensas (etapa, responsavel, origem...) mudam raramente e
' eram lidas UMA CONSULTA POR COMBO a cada ficha aberta. Agora: uma
' consulta traz todas, guardadas por 10 minutos.
Private mListas As Object
Private mListasLidasEm As Date
Private Const VALIDADE_LISTAS_MIN As Long = 10

'==========================================================
' CODIGOS
'==========================================================
Public Function CodigoVisualCliente(ByVal codigo As Variant) As String
    If IsNull(codigo) Or IsEmpty(codigo) Then CodigoVisualCliente = "": Exit Function
    CodigoVisualCliente = Format$(CLng(codigo), "0000")
End Function

Public Function MontarCodigo(ByVal prefixo As String, ByVal codCliente As Variant, _
                             ByVal controle As Variant) As String
    If IsNull(codCliente) Or IsNull(controle) Then Exit Function
    MontarCodigo = prefixo & "-" & Format$(CLng(codCliente), "0000") & _
                   "-" & Format$(CLng(controle), "0000")
End Function

' MAX+1 lido NA transacao que vai inserir (NZ nao existe via OLE DB).
Private Function ProximoNumeroEm(ByVal cn As Object, ByVal tabela As String, ByVal coluna As String) As Long
    Dim v As Variant
    v = modDB.ConsultarValorEm(cn, "SELECT MAX(" & coluna & ") FROM " & tabela, Null)
    If IsNull(v) Then ProximoNumeroEm = 1 Else ProximoNumeroEm = CLng(v) + 1
End Function

'==========================================================
' LISTAS SUSPENSAS
'==========================================================
Public Function ObterLista(ByVal tipo As String) As Variant
    Dim d As Variant, i As Long, t As String, c As Collection, k As Variant, m() As Variant
    If mListas Is Nothing Or DateDiff("n", mListasLidasEm, Now) >= VALIDADE_LISTAS_MIN Then
        d = modDB.Consultar("SELECT tipo, valor FROM listas WHERE ativo=True ORDER BY tipo, ordem, valor, id")
        Set mListas = CreateObject("Scripting.Dictionary")
        mListas.CompareMode = 1
        If Not IsEmpty(d) Then
            For i = 1 To UBound(d, 1)
                t = modDB.Nz(d(i, 1))
                If Not mListas.Exists(t) Then mListas.Add t, New Collection
                mListas(t).Add d(i, 2)
            Next i
        End If
        ' cada tipo vira matriz base 1 (linha, 1): o formato de sempre
        For Each k In mListas.Keys
            Set c = mListas(k)
            ReDim m(1 To c.Count, 1 To 1)
            For i = 1 To c.Count
                m(i, 1) = c(i)
            Next i
            mListas(k) = m
        Next k
        mListasLidasEm = Now
    End If
    If mListas.Exists(tipo) Then ObterLista = mListas(tipo)
End Function

' Depois de alterar a tabela listas, releia na proxima ficha.
Public Sub EsquecerListas()
    Set mListas = Nothing
End Sub

'----------------------------------------------------------
' Texto digitado na busca como padrao LIKE literal: _ % e [
' sao curingas do ACE e mudavam o resultado ("A_B" achava "AXB").
'----------------------------------------------------------
Public Function PadraoLike(ByVal texto As String) As String
    Dim s As String
    s = Replace(texto, "[", "[[]")
    s = Replace(s, "_", "[_]")
    s = Replace(s, "%", "[%]")
    PadraoLike = "%" & s & "%"
End Function

'==========================================================
' MONTAGEM DE SQL A PARTIR DO SCHEMA
'==========================================================
Private Function TipoAdo(ByVal tipoSchema As String) As Long
    Select Case tipoSchema
        Case "M": TipoAdo = modDB.adLongVarWChar
        Case "N": TipoAdo = modDB.adInteger
        Case "C": TipoAdo = modDB.adCurrency
        Case "D": TipoAdo = modDB.adDate
        Case "K": TipoAdo = modDB.adInteger
        Case Else: TipoAdo = modDB.adVarWChar
    End Select
End Function

' Campos que o formulario grava (editavel = 1): "campo|tipo"
Private Function CamposGravaveis(ByVal tabela As String) As Variant
    Dim campos As Variant, i As Long, spec As Variant
    Dim saida() As String, n As Long
    campos = modSchema.CamposDe(tabela)
    ReDim saida(0 To UBound(campos) - LBound(campos))
    For i = LBound(campos) To UBound(campos)
        spec = Split(campos(i), "|")
        If spec(6) = "1" Then
            saida(n) = spec(0) & "|" & spec(2)
            n = n + 1
        End If
    Next i
    ReDim Preserve saida(0 To n - 1)
    CamposGravaveis = saida
End Function

Private Function ValorDe(ByVal d As Object, ByVal campo As String) As Variant
    If d.Exists(campo) Then ValorDe = d(campo) Else ValorDe = Null
End Function

Private Function TipoDoExtra(ByVal campo As String) As Long
    Select Case campo
        Case "codigo", "estagio": TipoDoExtra = modDB.adVarWChar
        Case Else: TipoDoExtra = modDB.adInteger
    End Select
End Function

Private Function ParaMatriz(ByVal c As Collection) As Variant
    Dim r() As Variant, i As Long
    If c.Count = 0 Then ParaMatriz = Empty: Exit Function
    ReDim r(0 To c.Count - 1)
    For i = 1 To c.Count
        r(i - 1) = c(i)
    Next i
    ParaMatriz = r
End Function

'----------------------------------------------------------
' INSERT a partir do dicionario de dados, NA TRANSACAO, com o
' log de insercao. Devolve o id novo.
'----------------------------------------------------------
Private Function InserirEm(ByVal cn As Object, ByVal tabela As String, ByVal d As Object, _
                           ByVal extras As Object, ByVal detalhe As String) As Long
    Dim lista As Variant, i As Long, spec As Variant, chave As Variant
    Dim colunas As String, marcas As String, ps As New Collection, campo As String

    lista = CamposGravaveis(tabela)
    For i = LBound(lista) To UBound(lista)
        spec = Split(lista(i), "|")
        campo = spec(0)
        ' Campo que ja vem em extras NAO se repete: o vinculo (id_cliente,
        ' id_contato) e gravavel na ficha e tambem e definido pela rotina
        ' que cria o registro. Repetido, o INSERT sai com a mesma coluna
        ' duas vezes e o ACE recusa.
        If Not extras.Exists(campo) Then
            colunas = colunas & IIf(colunas = "", "", ", ") & campo
            marcas = marcas & IIf(marcas = "", "", ",") & "?"
            ps.Add modDB.P(TipoAdo(CStr(spec(1))), ValorDe(d, campo))
        End If
    Next i
    For Each chave In extras.Keys
        colunas = colunas & ", " & chave
        marcas = marcas & ",?"
        ps.Add modDB.P(TipoDoExtra(CStr(chave)), extras(chave))
    Next chave
    ' ativo e Sim/Nao: no Access nao aceita nulo e o padrao e FALSO. Sem
    ' esta linha, todo cadastro novo nascia INATIVO (defeito do legado,
    ' corrigido na migracao 002 para os registros ja gravados).
    If tabela = "clientes" Or tabela = "contatos" Then
        colunas = colunas & ", ativo"
        marcas = marcas & ",True"
    End If
    colunas = colunas & ", versao, criado_em, alterado_em, alterado_por"
    marcas = marcas & ",1,?,?,?"
    ps.Add modDB.P(modDB.adDate, Now)
    ps.Add modDB.P(modDB.adDate, Now)
    ps.Add modDB.P(modDB.adVarWChar, modConfig.UsuarioAtual())

    InserirEm = modDB.InserirEm(cn, "INSERT INTO " & tabela & " (" & colunas & ") VALUES (" & marcas & ")", _
                                ParaMatriz(ps))
    modDB.LogEm cn, tabela, InserirEm, "", "", detalhe, "INSERIR"
End Function

'----------------------------------------------------------
' UPDATE com versao, gravando SO o que mudou e registrando no
' log cada campo alterado (valor antigo e novo).
'   1 = gravado
'   0 = alguem alterou desde a leitura. NAO sobrescreva.
'   2 = nada mudou: nao grava, nao sobe versao, nao gera log
'----------------------------------------------------------
Public Function Atualizar(ByVal tabela As String, ByVal d As Object, _
                          ByVal idRegistro As Long, ByVal versaoLida As Long) As Long
    Dim lista As Variant, i As Long, spec As Variant, campo As String, tipo As String
    Dim cn As Object, atual As Variant, sel As String, sets As String
    Dim ps As New Collection, mudou As New Collection, novo As Variant, antigo As Variant, m As Variant
    Dim n As Long

    lista = CamposGravaveis(tabela)
    For i = LBound(lista) To UBound(lista)
        sel = sel & IIf(sel = "", "", ", ") & Split(lista(i), "|")(0)
    Next i

    On Error GoTo falha
    Set cn = modDB.AbrirTransacao()
    atual = modDB.ConsultarEm(cn, "SELECT " & sel & " FROM " & tabela & " WHERE id=? AND versao=?", _
                              Array(modDB.P(modDB.adInteger, idRegistro), modDB.P(modDB.adInteger, versaoLida)))
    If IsEmpty(atual) Then
        modDB.Desfazer cn
        Atualizar = 0
        Exit Function
    End If

    For i = LBound(lista) To UBound(lista)
        spec = Split(lista(i), "|")
        campo = spec(0): tipo = spec(1)
        novo = ValorDe(d, campo)
        antigo = atual(1, i - LBound(lista) + 1)
        If Not Iguais(antigo, novo, tipo) Then
            sets = sets & IIf(sets = "", "", ", ") & campo & "=?"
            ps.Add modDB.P(TipoAdo(tipo), novo)
            mudou.Add Array(campo, modDB.TextoLog(antigo), modDB.TextoLog(novo))
        End If
    Next i

    If mudou.Count = 0 Then
        modDB.Desfazer cn
        Atualizar = 2
        Exit Function
    End If

    ps.Add modDB.P(modDB.adDate, Now)
    ps.Add modDB.P(modDB.adVarWChar, modConfig.UsuarioAtual())
    ps.Add modDB.P(modDB.adInteger, idRegistro)
    ps.Add modDB.P(modDB.adInteger, versaoLida)
    n = modDB.ExecutarEm(cn, "UPDATE " & tabela & " SET " & sets & _
                         ", versao=versao+1, alterado_em=?, alterado_por=? WHERE id=? AND versao=?", ParaMatriz(ps))
    If n <> 1 Then
        modDB.Desfazer cn
        Atualizar = 0
        Exit Function
    End If
    For Each m In mudou
        modDB.LogEm cn, tabela, idRegistro, CStr(m(0)), CStr(m(1)), CStr(m(2)), "ALTERAR"
    Next m
    modDB.Confirmar cn
    Atualizar = 1
    Exit Function
falha:
    Dim nErr As Long, sErr As String
    nErr = Err.Number: sErr = Err.Description
    modDB.Desfazer cn
    Err.Raise nErr, "modCRM.Atualizar", sErr
End Function

' Mesmo valor, considerando o tipo: nulo = vazio; numero e data
' comparados como numero e data; texto comparado exatamente.
Private Function Iguais(ByVal a As Variant, ByVal b As Variant, ByVal tipo As String) As Boolean
    Dim va As Boolean, vb As Boolean
    va = IsNull(a) Or IsEmpty(a): If Not va Then va = (Len(CStr(a)) = 0)
    vb = IsNull(b) Or IsEmpty(b): If Not vb Then vb = (Len(CStr(b)) = 0)
    If va Or vb Then Iguais = (va And vb): Exit Function
    Select Case tipo
        Case "N", "C", "K"
            If IsNumeric(a) And IsNumeric(b) Then Iguais = (CDbl(a) = CDbl(b)): Exit Function
        Case "D"
            If IsDate(a) And IsDate(b) Then Iguais = (CDate(a) = CDate(b)): Exit Function
    End Select
    Iguais = (StrComp(CStr(a), CStr(b), vbBinaryCompare) = 0)
End Function

'==========================================================
' PROMOCAO DE PRE-CLIENTE
' Atribui o proximo codigo e muda o estagio na MESMA
' transacao. O UPDATE exige codigo_cliente ainda nulo: quem
' promover primeiro leva; o indice unico recusa numero repetido.
'
' PromoverEm devolve o codigo atribuido, ou 0 se a empresa ja
' tinha codigo (alguem promoveu antes).
'==========================================================
Private Function PromoverEm(ByVal cn As Object, ByVal idCliente As Long) As Long
    Dim codigo As Long, n As Long
    codigo = ProximoNumeroEm(cn, "clientes", "codigo_cliente")
    n = modDB.ExecutarEm(cn, "UPDATE clientes SET codigo_cliente=?, estagio=?, versao=versao+1," & _
                             " alterado_em=?, alterado_por=? WHERE id=? AND codigo_cliente IS NULL", _
        Array(modDB.P(modDB.adInteger, codigo), modDB.P(modDB.adVarWChar, "Cliente"), _
              modDB.P(modDB.adDate, Now), modDB.P(modDB.adVarWChar, modConfig.UsuarioAtual()), _
              modDB.P(modDB.adInteger, idCliente)))
    If n <> 1 Then Exit Function
    modDB.LogEm cn, "clientes", idCliente, "codigo_cliente", "", CStr(codigo), "PROMOVER"
    modDB.LogEm cn, "clientes", idCliente, "estagio", "Pré-cliente", "Cliente", "PROMOVER"
    PromoverEm = codigo
End Function

Public Function PromoverCliente(ByVal idCliente As Long, ByRef codigoAtribuido As Long) As Boolean
    Dim tentativa As Long, duplicou As Boolean
    On Error GoTo erro
    For tentativa = 1 To TENTATIVAS
        duplicou = False
        codigoAtribuido = TentarPromover(idCliente, duplicou)
        If codigoAtribuido > 0 Then PromoverCliente = True: Exit Function
        If Not duplicou Then
            MsgBox "Esta empresa já é cliente e já tem código.", vbInformation, "CRM Zapromaq"
            Exit Function
        End If
    Next tentativa
    MsgBox "Não foi possível atribuir o código depois de " & TENTATIVAS & " tentativas. Tente de novo em instantes.", _
           vbExclamation, "CRM Zapromaq"
    Exit Function
erro:
    MsgBox "Erro ao promover a empresa a cliente:" & vbCrLf & vbCrLf & Err.Description, vbCritical, "CRM Zapromaq"
End Function

Private Function TentarPromover(ByVal idCliente As Long, ByRef duplicou As Boolean) As Long
    Dim cn As Object
    On Error GoTo falha
    Set cn = modDB.AbrirTransacao()
    TentarPromover = PromoverEm(cn, idCliente)
    modDB.Confirmar cn
    Exit Function
falha:
    Dim nErr As Long, sErr As String
    nErr = Err.Number: sErr = Err.Description
    modDB.Desfazer cn
    If modDB.EhDuplicidade(sErr) Then duplicou = True: Exit Function
    Err.Raise nErr, "modCRM.PromoverCliente", sErr
End Function

'==========================================================
' CRIACAO DE REGISTROS
' Um caminho so por tipo de registro: o vendedor nunca digita
' codigo. Promocao (se preciso), numero sequencial, INSERT e
' log vao na MESMA transacao - ou tudo, ou nada.
'==========================================================
Public Function AbrirAtendimento(ByVal idContato As Long, ByVal d As Object, _
                                 ByRef codigoGerado As String) As Long
    Dim idCliente As Long, tentativa As Long, duplicou As Boolean, novoId As Long

    idCliente = CLng(modDB.ConsultarValor("SELECT id_cliente FROM contatos WHERE id=?", 0, _
                                         Array(modDB.P(modDB.adInteger, idContato))))
    If idCliente = 0 Then
        MsgBox "Contato não encontrado.", vbExclamation, "CRM Zapromaq"
        Exit Function
    End If
    If Not ConfirmarPromocao(idCliente, "Abrir o atendimento") Then Exit Function

    For tentativa = 1 To TENTATIVAS
        duplicou = False
        novoId = TentarCriar("oportunidades", idCliente, idContato, d, codigoGerado, duplicou)
        If novoId > 0 Then AbrirAtendimento = novoId: Exit Function
        If Not duplicou Then Exit Function
    Next tentativa
    Err.Raise vbObjectError + 513, "modCRM.AbrirAtendimento", _
              "Não foi possível gerar um código único depois de " & TENTATIVAS & " tentativas. Tente de novo."
End Function

Public Function CriarContato(ByVal idCliente As Long, ByVal d As Object, _
                             ByRef codigoGerado As String) As Long
    Dim tentativa As Long, duplicou As Boolean, novoId As Long
    ' pre-cliente nao pode ser beco sem saida: cadastrar o primeiro
    ' contato e prospectar, e prospectar promove a empresa.
    If Not ConfirmarPromocao(idCliente, "Cadastrar um contato") Then Exit Function

    For tentativa = 1 To TENTATIVAS
        duplicou = False
        novoId = TentarCriar("contatos", idCliente, 0, d, codigoGerado, duplicou)
        If novoId > 0 Then CriarContato = novoId: Exit Function
        If Not duplicou Then Exit Function
    Next tentativa
    Err.Raise vbObjectError + 513, "modCRM.CriarContato", _
              "Não foi possível gerar um código único depois de " & TENTATIVAS & " tentativas. Tente de novo."
End Function

Public Function CriarCliente(ByVal d As Object, ByVal comoCliente As Boolean, _
                             ByRef codigoGerado As Long) As Long
    Dim tentativa As Long, duplicou As Boolean, novoId As Long
    ' Pre-cliente nao recebe numero: nao ha o que repetir. Cliente
    ' recebe MAX+1; so a colisao DESSE numero justifica nova tentativa.
    For tentativa = 1 To TENTATIVAS
        duplicou = False
        novoId = TentarCriarCliente(d, comoCliente, codigoGerado, duplicou)
        If novoId > 0 Then CriarCliente = novoId: Exit Function
        If Not duplicou Then Exit Function
    Next tentativa
    Err.Raise vbObjectError + 513, "modCRM.CriarCliente", _
              "Não foi possível gerar um código único depois de " & TENTATIVAS & " tentativas. Tente de novo."
End Function

Private Function TentarCriarCliente(ByVal d As Object, ByVal comoCliente As Boolean, _
                                    ByRef codigoGerado As Long, ByRef duplicou As Boolean) As Long
    Dim cn As Object, extras As Object
    On Error GoTo falha
    Set cn = modDB.AbrirTransacao()
    Set extras = CreateObject("Scripting.Dictionary")
    If comoCliente Then
        codigoGerado = ProximoNumeroEm(cn, "clientes", "codigo_cliente")
        extras("codigo_cliente") = codigoGerado
        extras("estagio") = "Cliente"
    Else
        codigoGerado = 0
        extras("estagio") = "Pré-cliente"
    End If
    TentarCriarCliente = InserirEm(cn, "clientes", d, extras, "cadastro novo pelo formulário")
    modDB.Confirmar cn
    Exit Function
falha:
    Dim nErr As Long, sErr As String
    nErr = Err.Number: sErr = Err.Description
    modDB.Desfazer cn
    ' CNPJ repetido tambem e "duplicidade", mas nao se resolve repetindo:
    ' so o numero de cliente (codigo_cliente) justifica nova tentativa.
    If comoCliente And modDB.EhDuplicidade(sErr) And Not CnpjJaExiste(ValorDe(d, "cnpj")) Then
        duplicou = True
        Exit Function
    End If
    Err.Raise nErr, "modCRM.CriarCliente", sErr
End Function

Private Function CnpjJaExiste(ByVal cnpj As Variant) As Boolean
    If IsNull(cnpj) Or IsEmpty(cnpj) Then Exit Function
    If Len(CStr(cnpj)) = 0 Then Exit Function
    CnpjJaExiste = Not IsNull(modDB.ConsultarValor("SELECT id FROM clientes WHERE cnpj=?", Null, _
                                                   Array(modDB.P(modDB.adVarWChar, CStr(cnpj)))))
End Function

' Pre-cliente: pergunta antes de promover. Devolve False se o
' usuario desistiu. Cliente: segue sem perguntar.
Private Function ConfirmarPromocao(ByVal idCliente As Long, ByVal acao As String) As Boolean
    Dim codCliente As Variant
    codCliente = modDB.ConsultarValor("SELECT codigo_cliente FROM clientes WHERE id=?", Null, _
                                      Array(modDB.P(modDB.adInteger, idCliente)))
    If Not IsNull(codCliente) Then ConfirmarPromocao = True: Exit Function
    ConfirmarPromocao = (MsgBox("Esta empresa ainda é pré-cliente." & vbCrLf & vbCrLf & _
                         acao & " promove a empresa a cliente e atribui o próximo código. Continuar?", _
                         vbYesNo + vbQuestion, "CRM Zapromaq") = vbYes)
End Function

'----------------------------------------------------------
' Uma tentativa de criar contato ou atendimento. Tudo numa
' transacao: promove se ainda for pre-cliente, pega o proximo
' controle, monta o codigo congelado, insere e registra.
' Colisao no indice unico: devolve 0 com duplicou = True.
'----------------------------------------------------------
Private Function TentarCriar(ByVal tabela As String, ByVal idCliente As Long, ByVal idContato As Long, _
                             ByVal d As Object, ByRef codigoGerado As String, ByRef duplicou As Boolean) As Long
    Dim cn As Object, codCliente As Variant, controle As Long, extras As Object

    On Error GoTo falha
    Set cn = modDB.AbrirTransacao()
    codCliente = modDB.ConsultarValorEm(cn, "SELECT codigo_cliente FROM clientes WHERE id=?", Null, _
                                        Array(modDB.P(modDB.adInteger, idCliente)))
    If IsNull(codCliente) Then
        codCliente = PromoverEm(cn, idCliente)
        If codCliente = 0 Then
            ' alguem promoveu no meio do caminho: le o codigo que ficou
            codCliente = modDB.ConsultarValorEm(cn, "SELECT codigo_cliente FROM clientes WHERE id=?", Null, _
                                                Array(modDB.P(modDB.adInteger, idCliente)))
        End If
    End If
    If IsNull(codCliente) Then Err.Raise vbObjectError + 514, , "Empresa sem código de cliente."

    Set extras = CreateObject("Scripting.Dictionary")
    controle = ProximoNumeroEm(cn, tabela, "controle")
    extras("controle") = controle
    extras("id_cliente") = idCliente
    If tabela = "oportunidades" Then
        codigoGerado = MontarCodigo("AT", codCliente, controle)
        extras("id_contato") = idContato
    Else
        codigoGerado = MontarCodigo("CT", codCliente, controle)
    End If
    extras("codigo") = codigoGerado

    TentarCriar = InserirEm(cn, tabela, d, extras, "cadastro novo pelo formulário")
    modDB.Confirmar cn
    Exit Function
falha:
    Dim nErr As Long, sErr As String
    nErr = Err.Number: sErr = Err.Description
    modDB.Desfazer cn
    codigoGerado = ""
    If modDB.EhDuplicidade(sErr) Then duplicou = True: Exit Function
    Err.Raise nErr, "modCRM.TentarCriar", sErr
End Function

'==========================================================
' CONSULTAS DAS GRADES
' A grade ja chega ORDENADA POR CODIGO - e assim que se procura
' um registro no papel e no telefone. Pre-cliente nao tem codigo:
' vai depois dos clientes, por nome. Toda ordenacao termina no
' id: TOP do Access nao desempata e a navegacao Anterior/Proximo
' precisa de uma ordem estavel.
' Uma consulta devolve o conjunto. Nunca N consultas dentro
' de um laco: a 26 ms por ciclo, 500 linhas sao 13 segundos.
'
' ids (opcional): so os registros dessa lista - usado para
' reescrever na grade so as linhas que mudaram.
'==========================================================
Public Function SQLClientes(ByVal busca As String, ByVal estagio As String, _
                            ByVal responsavel As String, ByRef params As Variant, _
                            Optional ByVal qualificacao As String = "", _
                            Optional ByVal ordem As String = "", _
                            Optional ByVal ids As Variant, _
                            Optional ByVal situacao As String = "") As String
    Dim w As String, ps As New Collection

    w = FiltroSituacao("cl.ativo", situacao)
    If Trim$(busca) <> "" Then
        ' So procura no CNPJ se o texto tiver digito. Sem isso, "%%"
        ' casa com QUALQUER cnpj preenchido e a busca devolve a base toda.
        If modValidacao.SoDigitos(busca) <> "" Then
            w = w & " AND (cl.empresa LIKE ? OR cl.cidade LIKE ? OR cl.cnpj LIKE ?)"
            ps.Add modDB.P(modDB.adVarWChar, PadraoLike(busca))
            ps.Add modDB.P(modDB.adVarWChar, PadraoLike(busca))
            ps.Add modDB.P(modDB.adVarWChar, PadraoLike(modValidacao.SoDigitos(busca)))
        Else
            w = w & " AND (cl.empresa LIKE ? OR cl.cidade LIKE ?)"
            ps.Add modDB.P(modDB.adVarWChar, PadraoLike(busca))
            ps.Add modDB.P(modDB.adVarWChar, PadraoLike(busca))
        End If
    End If
    If Trim$(estagio) <> "" Then w = w & " AND cl.estagio=?": ps.Add modDB.P(modDB.adVarWChar, estagio)
    If Trim$(responsavel) <> "" Then w = w & " AND cl.responsavel=?": ps.Add modDB.P(modDB.adVarWChar, responsavel)
    If Trim$(qualificacao) <> "" Then w = w & " AND cl.qualificacao=?": ps.Add modDB.P(modDB.adVarWChar, qualificacao)
    w = w & FiltroIds("cl.id", ids, ps)

    params = ParaMatriz(ps)
    SQLClientes = "SELECT cl.id, cl.codigo_cliente, cl.estagio, cl.empresa, cl.cnpj," & _
                  " cl.cidade, cl.uf, cl.segmento, cl.responsavel, cl.aderencia, cl.porte," & _
                  " cl.qualificacao, cl.observacoes, cl.ctx_resumo, cl.ctx_familia," & _
                  " cl.ctx_atualizado_em, cl.telefone, cl.email, cl.ativo, cl.versao" & _
                  " FROM clientes cl WHERE 1=1" & w & _
                  OrdemClientes(ordem)
End Function

Public Function SQLContatos(ByVal busca As String, ByRef params As Variant, _
                            Optional ByVal ordem As String = "", _
                            Optional ByVal ids As Variant, _
                            Optional ByVal situacao As String = "") As String
    Dim w As String, ps As New Collection, k As Long
    w = FiltroSituacao("ct.ativo", situacao)
    If Trim$(busca) <> "" Then
        w = w & " AND (ct.nome LIKE ? OR cl.empresa LIKE ? OR ct.email LIKE ? OR ct.codigo LIKE ?)"
        For k = 1 To 4
            ps.Add modDB.P(modDB.adVarWChar, PadraoLike(busca))
        Next k
    End If
    w = w & FiltroIds("ct.id", ids, ps)
    params = ParaMatriz(ps)
    SQLContatos = "SELECT ct.id, ct.codigo, ct.id_cliente, cl.empresa, ct.nome, ct.cargo," & _
                  " ct.telefone, ct.email, cl.cidade, cl.uf, ct.observacoes, ct.ativo, ct.versao," & _
                  " cl.estagio, cl.codigo_cliente" & _
                  " FROM contatos ct LEFT JOIN clientes cl ON ct.id_cliente=cl.id" & _
                  " WHERE 1=1" & w & OrdemContatos(ordem)
End Function

Public Function SQLOportunidades(ByVal busca As String, ByVal etapa As String, _
                                 ByVal responsavel As String, ByVal somenteAbertas As Boolean, _
                                 ByRef params As Variant, _
                                 Optional ByVal origem As String = "", _
                                 Optional ByVal ordem As String = "", _
                                 Optional ByVal ids As Variant) As String
    Dim w As String, ps As New Collection, k As Long

    If Trim$(busca) <> "" Then
        w = w & " AND (op.codigo LIKE ? OR op.orcamento LIKE ? OR cl.empresa LIKE ?" & _
                " OR ct.nome LIKE ? OR op.maquina LIKE ?)"
        For k = 1 To 5
            ps.Add modDB.P(modDB.adVarWChar, PadraoLike(busca))
        Next k
    End If
    If Trim$(etapa) <> "" Then w = w & " AND op.etapa=?": ps.Add modDB.P(modDB.adVarWChar, etapa)
    If Trim$(responsavel) <> "" Then w = w & " AND op.responsavel=?": ps.Add modDB.P(modDB.adVarWChar, responsavel)
    If Trim$(origem) <> "" Then w = w & " AND op.origem=?": ps.Add modDB.P(modDB.adVarWChar, origem)
    If somenteAbertas Then w = w & " AND op.etapa NOT IN ('Pedido Fechado','Perdido','Descartado')"
    w = w & FiltroIds("op.id", ids, ps)

    params = ParaMatriz(ps)
    SQLOportunidades = "SELECT op.id, op.codigo, op.id_contato, op.id_cliente, cl.empresa," & _
        " cl.cidade, cl.uf, cl.segmento, ct.nome AS contato, ct.cargo, ct.telefone, ct.email," & _
        " op.orcamento, op.etapa, op.responsavel, op.maquina, op.familia, op.categoria," & _
        " op.tipo_venda, op.valor, op.origem, op.prioridade, op.dt_entrada, op.dt_proposta," & _
        " op.ultima_interacao, op.tentativas, op.prox_acao, op.dt_prox_acao, op.retomar_em," & _
        " op.motivo_desfecho, op.dt_desfecho, op.observacoes, op.ctx_resumo, op.versao, cl.estagio" & _
        " FROM (oportunidades op LEFT JOIN contatos ct ON op.id_contato=ct.id)" & _
        " LEFT JOIN clientes cl ON op.id_cliente=cl.id" & _
        " WHERE 1=1" & w & OrdemOportunidades(ordem)
End Function

' Consulta completa de uma tabela de grade (sem filtro de tela),
' opcionalmente restrita a uma lista de ids.
Public Function SQLGrade(ByVal tabela As String, ByRef params As Variant, Optional ByVal ids As Variant) As String
    Select Case tabela
        Case "clientes":      SQLGrade = SQLClientes("", "", "", params, "", "", ids)
        Case "contatos":      SQLGrade = SQLContatos("", params, "", ids)
        Case "oportunidades": SQLGrade = SQLOportunidades("", "", "", False, params, "", "", ids)
    End Select
End Function

'----------------------------------------------------------
' Situacao do cadastro (item 3): "ativos", "inativos" ou
' "todos"/vazio (sem filtro). Constante booleana do proprio
' SQL, nao parametro: nao vem de digitacao.
'----------------------------------------------------------
Private Function FiltroSituacao(ByVal coluna As String, ByVal situacao As String) As String
    Select Case LCase$(situacao)
        Case "ativos":   FiltroSituacao = " AND " & coluna & "=True"
        Case "inativos": FiltroSituacao = " AND " & coluna & "=False"
    End Select
End Function

Private Function FiltroIds(ByVal coluna As String, ByVal ids As Variant, ByVal ps As Collection) As String
    Dim marcas As String, pIds As Variant, i As Long
    If IsMissing(ids) Then Exit Function
    If Not IsArray(ids) Then Exit Function
    marcas = modDB.MarcadoresIds(ids, pIds)
    If marcas = "" Then Exit Function
    For i = LBound(pIds) To UBound(pIds)
        ps.Add pIds(i)
    Next i
    FiltroIds = " AND " & coluna & " IN (" & marcas & ")"
End Function

'==========================================================
' NAVEGACAO ENTRE FICHAS
'
' Devolve os ids de um vinculo, na ordem em que a grade os
' mostraria. E o mesmo formato que frmCRM.Abrir recebe da
' grade, entao a ficha aberta por aqui ja vem com Anterior
' e Proximo funcionando dentro do conjunto.
'==========================================================
Public Function IDsDeContatosDoCliente(ByVal idCliente As Long) As Variant
    IDsDeContatosDoCliente = ColunaDeIDs( _
        "SELECT id FROM contatos WHERE id_cliente=? ORDER BY codigo, id", idCliente)
End Function

Public Function IDsDeAtendimentosDoCliente(ByVal idCliente As Long) As Variant
    IDsDeAtendimentosDoCliente = ColunaDeIDs( _
        "SELECT id FROM oportunidades WHERE id_cliente=? ORDER BY codigo DESC, id DESC", idCliente)
End Function

Public Function IDsDeAtendimentosDoContato(ByVal idContato As Long) As Variant
    IDsDeAtendimentosDoContato = ColunaDeIDs( _
        "SELECT id FROM oportunidades WHERE id_contato=? ORDER BY codigo DESC, id DESC", idContato)
End Function

Private Function ColunaDeIDs(ByVal sql As String, ByVal idVinculo As Long) As Variant
    Dim d As Variant, i As Long, ids() As Long
    If idVinculo = 0 Then Exit Function
    d = modDB.Consultar(sql, Array(modDB.P(modDB.adInteger, idVinculo)))
    If IsEmpty(d) Then Exit Function
    ReDim ids(1 To UBound(d, 1))
    For i = 1 To UBound(d, 1)
        ids(i) = CLng(modDB.Nz(d(i, 1), "0"))
    Next i
    ColunaDeIDs = ids
End Function

'==========================================================
' ORDENACAO DA CONSULTA
'
' Ordenar no BANCO, e nao na lista da tela, e o que faz a
' navegacao Anterior/Proximo seguir a mesma ordem que se ve.
' Data nula por ultimo: em Access, ORDER BY joga Nulo na
' frente, e "sem proxima acao" no topo da fila de trabalho
' empurra para baixo justamente o que tem prazo.
'==========================================================
Private Function OrdemClientes(ByVal ordem As String) As String
    Select Case LCase$(ordem)
        Case "empresa":  OrdemClientes = " ORDER BY cl.empresa, cl.id"
        Case "cidade":   OrdemClientes = " ORDER BY cl.uf, cl.cidade, cl.empresa, cl.id"
        Case "aderencia": OrdemClientes = " ORDER BY IIf(cl.aderencia Is Null,1,0), cl.aderencia DESC, cl.empresa, cl.id"
        Case Else:       OrdemClientes = " ORDER BY IIf(cl.codigo_cliente Is Null,1,0), cl.codigo_cliente, cl.empresa, cl.id"
    End Select
End Function

Private Function OrdemContatos(ByVal ordem As String) As String
    Select Case LCase$(ordem)
        Case "empresa": OrdemContatos = " ORDER BY cl.empresa, ct.nome, ct.id"
        Case "nome":    OrdemContatos = " ORDER BY ct.nome, ct.id"
        Case Else:      OrdemContatos = " ORDER BY ct.codigo, ct.id"
    End Select
End Function

Private Function OrdemOportunidades(ByVal ordem As String) As String
    Select Case LCase$(ordem)
        Case "prox":    OrdemOportunidades = " ORDER BY IIf(op.dt_prox_acao Is Null,1,0), op.dt_prox_acao, op.id"
        Case "valor":   OrdemOportunidades = " ORDER BY IIf(op.valor Is Null,1,0), op.valor DESC, op.id"
        Case "entrada": OrdemOportunidades = " ORDER BY IIf(op.dt_entrada Is Null,1,0), op.dt_entrada DESC, op.id DESC"
        Case "empresa": OrdemOportunidades = " ORDER BY cl.empresa, op.codigo DESC, op.id DESC"
        Case "etapa":   OrdemOportunidades = " ORDER BY op.etapa, IIf(op.dt_prox_acao Is Null,1,0), op.dt_prox_acao, op.id"
        Case Else:      OrdemOportunidades = " ORDER BY op.codigo DESC, op.id DESC"
    End Select
End Function

'==========================================================
' RESUMO DE VINCULO - uma consulta, ao trocar o vinculo
'==========================================================
Public Function ResumoContato(ByVal idContato As Long) As String
    Dim d As Object
    Set d = DadosDoContato(idContato)
    If d Is Nothing Then Exit Function
    ResumoContato = TextoResumoContato(d("nome"), d("cargo"), d("empresa"), d("cidade"), d("uf"), d("estagio"))
End Function

'----------------------------------------------------------
' Tudo o que a ficha do atendimento mostra do contato, numa
' consulta: trocar o contato atualiza cargo, telefone e e-mail
' e da a empresa dele (a troca so vale dentro da mesma empresa).
' Devolve um dicionario campo -> valor, ou Nothing.
'----------------------------------------------------------
Public Function DadosDoContato(ByVal idContato As Long) As Object
    Dim d As Variant, r As Object, nomes As Variant, i As Long
    If idContato = 0 Then Exit Function
    d = modDB.Consultar( _
        "SELECT ct.id_cliente, ct.nome, ct.cargo, ct.telefone, ct.email, cl.empresa, cl.cidade, cl.uf, cl.estagio" & _
        " FROM contatos ct LEFT JOIN clientes cl ON ct.id_cliente=cl.id WHERE ct.id=?", _
        Array(modDB.P(modDB.adInteger, idContato)))
    If IsEmpty(d) Then Exit Function
    ' gCampos tem o indice da coluna na matriz (como em IndiceDeCampos)
    nomes = modDB.gCampos
    Set r = CreateObject("Scripting.Dictionary")
    For i = LBound(nomes) To UBound(nomes)
        r(LCase$(nomes(i))) = d(1, i)
    Next i
    Set DadosDoContato = r
End Function

'----------------------------------------------------------
' Troca de contato de um atendimento ja gravado: so dentro da
' mesma empresa. O codigo AT- guarda o codigo do cliente e e
' congelado (nomeia a pasta em 02 - CLIENTES), e o id_cliente do
' atendimento nao e gravavel pela ficha - trocar de empresa
' deixaria o atendimento apontando para duas empresas.
' idClienteAtual = 0: atendimento novo, ainda sem empresa.
'----------------------------------------------------------
'----------------------------------------------------------
' Quando o botao de um campo de vinculo da ficha fica ativo.
' Sempre so na edicao. O do contato, so em atendimento JA
' GRAVADO: o novo escolhe empresa e contato pelo Procurar, em
' dois passos.
'----------------------------------------------------------
Public Function BotaoVinculoAtivo(ByVal nomeCampo As String, ByVal editando As Boolean, _
                                  ByVal novo As Boolean, ByVal idContato As Long) As Boolean
    If Not editando Then Exit Function
    Select Case nomeCampo
        Case "id_contato": BotaoVinculoAtivo = Not novo
        Case Else:         BotaoVinculoAtivo = True
    End Select
End Function

Public Function CriticarTrocaContato(ByVal idClienteAtual As Long, ByVal idClienteNovo As Long) As String
    If idClienteAtual = 0 Then Exit Function
    If idClienteNovo = idClienteAtual Then Exit Function
    CriticarTrocaContato = "O contato escolhido é de outra empresa." & vbCrLf & vbCrLf & _
        "O atendimento só troca de contato dentro da mesma empresa: o código AT- guarda " & _
        "o código do cliente e não muda. Se o atendimento foi aberto na empresa errada, " & _
        "encerre-o como Descartado e abra um novo apontando para ele como anterior."
End Function

Public Function ResumoCliente(ByVal idCliente As Long) As String
    Dim d As Variant
    If idCliente = 0 Then Exit Function
    d = modDB.Consultar("SELECT empresa, cidade, uf, estagio, codigo_cliente FROM clientes WHERE id=?", _
                        Array(modDB.P(modDB.adInteger, idCliente)))
    If IsEmpty(d) Then Exit Function
    ResumoCliente = TextoResumoCliente(d(1, 1), d(1, 2), d(1, 3), d(1, 4), d(1, 5))
End Function

' Mesmo texto, montado de valores que a ficha ja tem na memoria
' (linha da grade): navegar entre registros nao consulta o banco.
Public Function TextoResumoContato(ByVal nome As Variant, ByVal cargo As Variant, ByVal empresa As Variant, _
                                   ByVal cidade As Variant, ByVal uf As Variant, ByVal estagio As Variant) As String
    TextoResumoContato = modDB.Nz(nome) & IIf(modDB.Nz(cargo) = "", "", " (" & modDB.Nz(cargo) & ")") & _
                         "  -  " & modDB.Nz(empresa) & "  " & modDB.Nz(cidade) & "/" & modDB.Nz(uf) & _
                         "  [" & modDB.Nz(estagio) & "]"
End Function

Public Function TextoResumoCliente(ByVal empresa As Variant, ByVal cidade As Variant, ByVal uf As Variant, _
                                   ByVal estagio As Variant, ByVal codigo As Variant) As String
    TextoResumoCliente = modDB.Nz(empresa) & "  " & modDB.Nz(cidade) & "/" & modDB.Nz(uf) & _
                         "  [" & modDB.Nz(estagio) & " " & CodigoVisualCliente(codigo) & "]"
End Function

'----------------------------------------------------------
' Busca de vinculo: o vendedor procura por nome, nao decora id.
' TOP 50 com desempate por id (TOP do Access nao desempata).
'----------------------------------------------------------
Public Function BuscarClientes(ByVal texto As String) As Variant
    ' mesmo cuidado do SQLClientes: CNPJ so entra se houver digito
    If modValidacao.SoDigitos(texto) <> "" Then
        BuscarClientes = modDB.Consultar( _
            "SELECT TOP 50 id, codigo_cliente, empresa, cidade, uf, estagio FROM clientes" & _
            " WHERE ativo=True AND (empresa LIKE ? OR cnpj LIKE ?) ORDER BY empresa, id", _
            Array(modDB.P(modDB.adVarWChar, PadraoLike(texto)), _
                  modDB.P(modDB.adVarWChar, PadraoLike(modValidacao.SoDigitos(texto)))))
    Else
        BuscarClientes = modDB.Consultar( _
            "SELECT TOP 50 id, codigo_cliente, empresa, cidade, uf, estagio FROM clientes" & _
            " WHERE ativo=True AND empresa LIKE ? ORDER BY empresa, id", _
            Array(modDB.P(modDB.adVarWChar, PadraoLike(texto))))
    End If
End Function

Public Function BuscarContatos(ByVal texto As String) As Variant
    BuscarContatos = modDB.Consultar( _
        "SELECT TOP 50 ct.id, ct.codigo, ct.nome, cl.empresa, ct.cargo FROM contatos ct" & _
        " LEFT JOIN clientes cl ON ct.id_cliente=cl.id" & _
        " WHERE ct.ativo=True AND (ct.nome LIKE ? OR cl.empresa LIKE ? OR ct.codigo LIKE ?)" & _
        " ORDER BY cl.empresa, ct.nome, ct.id", _
        Array(modDB.P(modDB.adVarWChar, PadraoLike(texto)), _
              modDB.P(modDB.adVarWChar, PadraoLike(texto)), _
              modDB.P(modDB.adVarWChar, PadraoLike(texto))))
End Function

'==========================================================
' ESCOLHA DE VINCULO (frmVinculo)
' Lista para escolher contato (atendimento novo) ou empresa
' (contato novo). So ativos; TOP com desempate por id. Texto
' vazio traz os primeiros da lista, por empresa.
'==========================================================
' idCliente > 0: so os contatos dessa empresa (2o passo do atendimento novo)
Public Function ListarContatosParaVinculo(ByVal texto As String, Optional ByVal idCliente As Long = 0) As Variant
    Dim w As String, ps As New Collection, k As Long, m() As Variant
    If idCliente > 0 Then
        w = " AND ct.id_cliente=?"
        ps.Add modDB.P(modDB.adInteger, idCliente)
    End If
    If Trim$(texto) <> "" Then
        w = w & " AND (ct.nome LIKE ? OR cl.empresa LIKE ? OR ct.codigo LIKE ? OR cl.cidade LIKE ? OR ct.cargo LIKE ?)"
        For k = 1 To 5
            ps.Add modDB.P(modDB.adVarWChar, PadraoLike(texto))
        Next k
    End If
    ListarContatosParaVinculo = modDB.Consultar( _
        "SELECT TOP 300 ct.id, ct.codigo, ct.nome, ct.cargo, cl.empresa, cl.cidade, cl.uf, ct.telefone, ct.id_cliente" & _
        " FROM contatos ct LEFT JOIN clientes cl ON ct.id_cliente=cl.id" & _
        " WHERE ct.ativo=True" & w & " ORDER BY cl.empresa, ct.nome, ct.id", ParaMatriz(ps))
End Function

Public Function ListarClientesParaVinculo(ByVal texto As String) As Variant
    Dim w As String, ps As Variant, dig As String
    If Trim$(texto) <> "" Then
        dig = modValidacao.SoDigitos(texto)
        If dig <> "" Then
            w = " AND (empresa LIKE ? OR cidade LIKE ? OR cnpj LIKE ? OR Format(codigo_cliente,'0000') LIKE ?)"
            ps = Array(modDB.P(modDB.adVarWChar, PadraoLike(texto)), modDB.P(modDB.adVarWChar, PadraoLike(texto)), _
                       modDB.P(modDB.adVarWChar, PadraoLike(dig)), modDB.P(modDB.adVarWChar, PadraoLike(dig)))
        Else
            w = " AND (empresa LIKE ? OR cidade LIKE ?)"
            ps = Array(modDB.P(modDB.adVarWChar, PadraoLike(texto)), modDB.P(modDB.adVarWChar, PadraoLike(texto)))
        End If
    End If
    ListarClientesParaVinculo = modDB.Consultar( _
        "SELECT TOP 300 id, codigo_cliente, empresa, cidade, uf, estagio, cnpj, responsavel" & _
        " FROM clientes WHERE ativo=True" & w & " ORDER BY empresa, id", ps)
End Function
