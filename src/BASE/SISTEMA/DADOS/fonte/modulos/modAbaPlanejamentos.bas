Attribute VB_Name = "modAbaPlanejamentos"
'==========================================================
' modAbaPlanejamentos - aba Planejamentos (fatia 4.8).
'
' Seletor do planejamento em C4 (lista de validacao que aponta
' para uma faixa oculta da propria aba, nao para um texto
' "a,b,c": nao depende do separador de lista do Windows) e a
' grade das visitas dele, SOMENTE LEITURA, com a coluna oculta
' _id, como as grades de modGrade (aba protegida, so o seletor
' aceita digitacao, nada de mesclagem na area de dados).
'
' Duplo clique numa visita abre o lancamento dela. No front o
' Lancamento e janela propria (frmLancamento), a mesma que a
' pagina Visitas do frmPlanejamento abre; nao ha pagina
' Lancamento dentro do frmPlanejamento.
'
' O feed (modFeed) remonta a aba quando visitas ou planejamentos
' mudam - e uma consulta so, e a aba mostra um planejamento.
' Remontar desfaz filtro feito pelo usuario nesta aba: aceito,
' porque a aba e pequena e mostrar o que mudou vale mais.
'==========================================================
Option Explicit

Public Const NOME_ABA As String = "Planejamentos"
Public Const LIN_SELETOR As Long = 4
Public Const COL_SELETOR As Long = 3          ' C4 (mesclada ate F4)
Private Const LIN_TITULO As Long = 2
Private Const LIN_STATUS As Long = 5
Private Const LIN_CAB As Long = 7
Private Const COL_INI As Long = 2
Private Const COL_OPCOES As Long = 60         ' texto do seletor | id do planejamento (ocultas)
Private Const NCOLS As Long = 10              ' 9 visiveis + _id
Private Const CAB_ID As String = "_id"

'==========================================================
' BOTOES (aba Inicio e a propria aba)
'==========================================================
Public Sub GradePlanejamentos()
    Dim ws As Worksheet
    If Not modMenu.Pronto() Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(NOME_ABA)
    ws.Activate
    Montar ws, False
End Sub

Public Sub PlanejamentosAtualizar()
    If Not modMenu.Pronto() Then Exit Sub
    Montar ThisWorkbook.Worksheets(NOME_ABA), False
End Sub

' Abre o planejamento (cabecalho, Visitas, Candidatos); na volta a aba
' mostra o que foi gravado la.
Public Sub PlanejamentosAbrirForm()
    If Not modMenu.Pronto() Then Exit Sub
    Dim f As New frmPlanejamento
    f.Show
    Montar ThisWorkbook.Worksheets(NOME_ABA), True
End Sub

Public Function EhAba(ByVal ws As Object) As Boolean
    EhAba = (ws.Name = NOME_ABA)
End Function

'==========================================================
' EVENTOS (ThisWorkbook)
'==========================================================
' Mudou o seletor: mostra as visitas do planejamento escolhido.
Public Sub Mudou(ByVal ws As Object, ByVal alvo As Object)
    If Intersect(alvo, ws.Cells(LIN_SELETOR, COL_SELETOR)) Is Nothing Then Exit Sub
    If Not modMenu.Pronto() Then Exit Sub
    Montar ws, False
End Sub

' Duplo clique: True quando caiu numa visita (o Excel nao entra em
' edicao); fora da grade, o duplo clique normal continua valendo.
Public Function DuploClique(ByVal ws As Object, ByVal alvo As Object) As Boolean
    Dim id As Long
    id = IdDaVisitaNaLinha(ws, alvo)
    If id = 0 Then Exit Function
    DuploClique = True
    If Not modMenu.Pronto() Then Exit Function
    Dim f As New frmLancamento
    f.Abrir id
    ' a janela e modal: na volta, a aba mostra o que foi lancado
    Montar ws, False
End Function

' Feed: visitas ou planejamentos mudaram. So remonta a aba que ja foi
' carregada; nunca mostra MsgBox (roda pelo OnTime).
Public Sub AtualizarPeloFeed()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(NOME_ABA)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    If ws.ListObjects.Count = 0 And CStr(modDB.Nz(ws.Cells(LIN_SELETOR, COL_SELETOR).Value2)) = "" Then Exit Sub
    Montar ws, True
End Sub

'==========================================================
' LINHA -> VISITA (coluna _id, nunca a posicao da linha)
'==========================================================
Public Function IdDaVisitaNaLinha(ByVal ws As Object, ByVal alvo As Object) As Long
    Dim lo As Object, j As Long, v As Variant
    If ws.ListObjects.Count = 0 Then Exit Function
    Set lo = ws.ListObjects(1)
    If lo.DataBodyRange Is Nothing Then Exit Function
    If Intersect(alvo, lo.DataBodyRange) Is Nothing Then Exit Function
    For j = 1 To lo.ListColumns.Count
        If CStr(lo.ListColumns(j).Name) = CAB_ID Then
            v = ws.Cells(alvo.Row, lo.Range.Column + j - 1).Value2
            If IsNumeric(v) Then IdDaVisitaNaLinha = CLng(v)
            Exit Function
        End If
    Next j
End Function

' Texto do seletor para o planejamento (o mesmo da lista de C4).
Public Function TextoDoSeletor(ByVal codigo As String, ByVal descricao As String, ByVal situacao As String) As String
    TextoDoSeletor = codigo & "  -  " & descricao & "  (" & situacao & ")"
End Function

' Situacao exibida da visita (spec 003): PENDENTE sem data = A AGENDAR,
' com data = AGENDADA; REALIZADA e CANCELADA como gravadas.
Public Function SituacaoExibida(ByVal situacao As String, ByVal dataAgendada As Variant) As String
    If situacao = "PENDENTE" Then
        If IsNull(dataAgendada) Or IsEmpty(dataAgendada) Then SituacaoExibida = "A AGENDAR" Else SituacaoExibida = "AGENDADA"
    Else
        SituacaoExibida = situacao
    End If
End Function

'==========================================================
' MONTAGEM
'==========================================================
Public Sub Montar(ByVal ws As Worksheet, ByVal silencioso As Boolean)
    Dim cn As Object, pl As Variant, d As Variant, escolhido As String
    Dim i As Long, np As Long, nl As Long, pos As Long, idPl As Long, pend As Long
    Dim arr() As Variant, ops() As Variant, rng As Range, lo As ListObject
    Dim eventos As Boolean, idAnterior As Long

    eventos = Application.EnableEvents
    On Error GoTo trata
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    escolhido = CStr(modDB.Nz(ws.Cells(LIN_SELETOR, COL_SELETOR).Value2))
    If IsNumeric(ws.Cells(1, COL_OPCOES + 2).Value2) Then idAnterior = CLng(ws.Cells(1, COL_OPCOES + 2).Value2)

    ' uma conexao para as duas leituras
    Set cn = modDB.AbrirLeitura()
    pl = modDB.ConsultarEm(cn, "SELECT id, codigo, descricao, situacao FROM planejamentos ORDER BY codigo DESC")
    If IsEmpty(pl) Then np = 0 Else np = UBound(pl, 1)
    ' escolhido: o da celula, se ainda existe; senao o primeiro ABERTO; senao o primeiro
    pos = 0
    For i = 1 To np
        If TextoDoSeletor(CStr(pl(i, 2)), modDB.Nz(pl(i, 3)), CStr(pl(i, 4))) = escolhido Then pos = i
    Next i
    ' o texto do seletor leva a situacao e a descricao: se mudaram (encerrou, renomeou), o texto da celula
    ' nao casa mais; o id guardado na faixa oculta mantem o mesmo planejamento escolhido
    If pos = 0 And idAnterior > 0 Then
        For i = 1 To np
            If CLng(pl(i, 1)) = idAnterior Then pos = i
        Next i
    End If
    If pos = 0 Then
        For i = np To 1 Step -1
            If CStr(pl(i, 4)) = "ABERTO" Then pos = i
        Next i
    End If
    If pos = 0 And np > 0 Then pos = 1
    If pos > 0 Then
        idPl = CLng(pl(pos, 1))
        d = modDB.ConsultarEm(cn, "SELECT v.id, v.situacao, v.data_agendada, v.hora_agendada, o.codigo, c.empresa, " & _
                                  "c.cidade, c.uf, ct.nome, v.tipo, v.resultado " & _
                                  "FROM ((visitas v INNER JOIN oportunidades o ON v.id_oportunidade = o.id) " & _
                                  "INNER JOIN clientes c ON o.id_cliente = c.id) LEFT JOIN contatos ct ON o.id_contato = ct.id " & _
                                  "WHERE v.id_planejamento=? ORDER BY IIf(v.situacao='PENDENTE', 0, 1), v.data_agendada, v.id", _
                              Array(modDB.P(modDB.adInteger, idPl)))
    End If
    modDB.FecharLeitura cn
    Set cn = Nothing
    If IsEmpty(d) Then nl = 0 Else nl = UBound(d, 1)

    modGrade.Desproteger ws
    Limpar ws
    Cabecalho ws

    ' opcoes do seletor (faixa oculta) e a celula C4
    If np > 0 Then
        ReDim ops(1 To np, 1 To 2)
        For i = 1 To np
            ops(i, 1) = TextoDoSeletor(CStr(pl(i, 2)), modDB.Nz(pl(i, 3)), CStr(pl(i, 4)))
            ops(i, 2) = CLng(pl(i, 1))
        Next i
        ws.Range(ws.Cells(1, COL_OPCOES), ws.Cells(np, COL_OPCOES + 1)).Value2 = ops
        ws.Cells(LIN_SELETOR, COL_SELETOR).Value2 = ops(pos, 1)
        ws.Cells(1, COL_OPCOES + 2).Value2 = idPl
        With ws.Cells(LIN_SELETOR, COL_SELETOR).Validation
            .Delete
            ' xlValidateList = 3, xlValidAlertStop = 1, xlBetween = 1
            .Add Type:=3, AlertStyle:=1, Operator:=1, _
                 Formula1:="=" & ws.Range(ws.Cells(1, COL_OPCOES), ws.Cells(np, COL_OPCOES)).Address
            .InCellDropdown = True
            .InputTitle = "Planejamento"
            .InputMessage = "Escolha o planejamento para ver as visitas."
        End With
    End If
    ws.Columns(COL_OPCOES).Hidden = True
    ws.Columns(COL_OPCOES + 1).Hidden = True
    ws.Columns(COL_OPCOES + 2).Hidden = True

    ' grade das visitas
    ReDim arr(1 To nl + 1, 1 To NCOLS)
    arr(1, 1) = "Situa" & ChrW$(&HE7) & ChrW$(&HE3) & "o"
    arr(1, 2) = "Data"
    arr(1, 3) = "Hora"
    arr(1, 4) = "Atendimento"
    arr(1, 5) = "Empresa"
    arr(1, 6) = "Cidade/UF"
    arr(1, 7) = "Contato"
    arr(1, 8) = "Tipo"
    arr(1, 9) = "Resultado"
    arr(1, 10) = CAB_ID
    For i = 1 To nl
        arr(i + 1, 1) = SituacaoExibida(CStr(d(i, 2)), d(i, 3))
        If IsNull(d(i, 3)) Then arr(i + 1, 2) = "" Else arr(i + 1, 2) = CDate(d(i, 3))
        arr(i + 1, 3) = modDB.Nz(d(i, 4))
        arr(i + 1, 4) = modDB.Nz(d(i, 5))
        arr(i + 1, 5) = modDB.Nz(d(i, 6))
        arr(i + 1, 6) = modDB.Nz(d(i, 7)) & IIf(modDB.Nz(d(i, 8)) = "", "", "/" & modDB.Nz(d(i, 8)))
        arr(i + 1, 7) = modDB.Nz(d(i, 9))
        arr(i + 1, 8) = modDB.Nz(d(i, 10))
        arr(i + 1, 9) = modDB.Nz(d(i, 11))
        arr(i + 1, 10) = CLng(d(i, 1))
        If CStr(d(i, 2)) = "PENDENTE" Then pend = pend + 1
    Next i
    Set rng = ws.Range(ws.Cells(LIN_CAB, COL_INI), ws.Cells(LIN_CAB + nl, COL_INI + NCOLS - 1))
    ' hora e atendimento sao texto: "09:00" nao vira numero
    ws.Range(ws.Cells(LIN_CAB + 1, COL_INI + 2), ws.Cells(LIN_CAB + nl + 1, COL_INI + 3)).NumberFormat = "@"
    rng.Value2 = arr
    ' xlSrcRange = 1, xlYes = 1
    Set lo = ws.ListObjects.Add(1, rng, , 1)
    lo.Name = "tblPlanejamentos"
    lo.TableStyle = "TableStyleLight1"
    lo.ShowTableStyleRowStripes = True
    Formatar ws, lo, nl

    If pos = 0 Then
        ws.Cells(LIN_STATUS, COL_INI).Value2 = "Nenhum planejamento gravado. Use Abrir planejamento para criar o primeiro."
    Else
        ws.Cells(LIN_STATUS, COL_INI).Value2 = nl & " visita(s), " & pend & " pendente(s)  " & ChrW$(&HB7) & _
            "  consulta de " & Format$(Now, "dd/mm/yyyy hh:nn") & "  " & ChrW$(&HB7) & _
            "  duplo clique numa visita abre o lan" & ChrW$(&HE7) & "amento"
    End If

    Congelar ws
    modGrade.Proteger ws
    Application.ScreenUpdating = True
    Application.EnableEvents = eventos
    Exit Sub

trata:
    Dim sErro As String
    sErro = Err.Description
    On Error Resume Next
    If Not cn Is Nothing Then modDB.FecharLeitura cn
    ws.Cells(LIN_STATUS, COL_INI).Value2 = "Erro ao montar a aba: " & sErro
    ' a aba NUNCA fica desprotegida (mesma regra das grades)
    modGrade.Proteger ws
    Application.ScreenUpdating = True
    Application.EnableEvents = eventos
    On Error GoTo 0
    If Not silencioso Then
        MsgBox "Nao foi possivel montar a aba Planejamentos." & vbCrLf & vbCrLf & sErro, vbExclamation, "CRM Zapromaq"
    End If
End Sub

' Mesma ordem do modGrade.Limpar: tabela (Delete), filtro, reexibir, limpar.
Private Sub Limpar(ByVal ws As Worksheet)
    Dim i As Long
    On Error Resume Next
    For i = ws.ListObjects.Count To 1 Step -1
        ws.ListObjects(i).Delete
    Next i
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Cells.Rows.Hidden = False
    ws.Cells.Columns.Hidden = False
    ws.Range(ws.Cells(LIN_SELETOR, COL_SELETOR), ws.Cells(LIN_SELETOR, COL_SELETOR + 3)).UnMerge
    On Error GoTo 0
    ws.Cells.Clear
    ws.Cells.Locked = True
End Sub

' Faixa de cima: titulo (forma do build, ao lado do logo), rotulo e
' seletor (C4:F4 mesclada - mesclagem no cabecalho e permitida).
Private Sub Cabecalho(ByVal ws As Worksheet)
    ws.Columns(1).ColumnWidth = 2.25
    ws.Rows(1).RowHeight = 4
    ws.Rows(LIN_TITULO).RowHeight = 28
    ws.Rows(3).RowHeight = 34
    ws.Rows(LIN_SELETOR).RowHeight = 22
    ws.Rows(LIN_STATUS).RowHeight = 16
    ws.Rows(6).RowHeight = 10
    ws.Range(ws.Cells(1, 1), ws.Cells(6, 40)).Interior.Color = modTema.COR_FUNDO

    With ws.Cells(LIN_SELETOR, COL_INI)
        .Value2 = "Planejamento"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Bold = True
        .Font.Color = modTema.COR_AZUL
        .HorizontalAlignment = -4152             ' xlRight
    End With
    With ws.Range(ws.Cells(LIN_SELETOR, COL_SELETOR), ws.Cells(LIN_SELETOR, COL_SELETOR + 3))
        .Merge
        .Interior.Color = modTema.COR_BRANCO
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .HorizontalAlignment = -4131             ' xlLeft
        .VerticalAlignment = -4108               ' xlCenter
        .IndentLevel = 1
        .ShrinkToFit = True
        .Locked = False                          ' unica celula digitavel da aba
        .Borders.Color = modTema.COR_VERDE
        .Borders.Weight = 2
    End With
    With ws.Cells(LIN_STATUS, COL_INI)
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Color = modTema.COR_TEXTO2
    End With
End Sub

Private Sub Formatar(ByVal ws As Worksheet, ByVal lo As ListObject, ByVal nl As Long)
    Dim larguras As Variant, j As Long
    With lo.HeaderRowRange
        .Interior.Color = modTema.COR_AZUL
        .Font.Name = "Segoe UI"
        .Font.Color = modTema.COR_BRANCO
        .Font.Bold = True
        .HorizontalAlignment = -4131             ' xlLeft
        .RowHeight = 22
        .VerticalAlignment = -4108               ' xlCenter
    End With
    larguras = Array(14, 12, 8, 16, 40, 24, 22, 16, 18)
    For j = 0 To 8
        ws.Columns(COL_INI + j).ColumnWidth = larguras(j)
    Next j
    ws.Columns(COL_INI + 1).NumberFormat = "dd/mm/yyyy"
    ws.Columns(COL_INI + NCOLS - 1).Hidden = True      ' _id: oculta, nunca apagada
    If nl > 0 Then
        With lo.DataBodyRange
            .Font.Name = "Segoe UI"
            .Font.Size = 10
            .VerticalAlignment = -4108           ' xlCenter
            .WrapText = False
            .ShrinkToFit = True
            .RowHeight = 15
            .Locked = True
        End With
    End If
End Sub

Private Sub Congelar(ByVal ws As Worksheet)
    On Error Resume Next
    If Not ActiveSheet Is ws Then Exit Sub
    ActiveWindow.FreezePanes = False
    ws.Cells(LIN_CAB + 1, COL_INI).Select
    ActiveWindow.FreezePanes = True
    ws.Cells(LIN_SELETOR, COL_SELETOR).Select
    On Error GoTo 0
End Sub
