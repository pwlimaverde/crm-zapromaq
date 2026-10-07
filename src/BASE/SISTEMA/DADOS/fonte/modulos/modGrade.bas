Attribute VB_Name = "modGrade"
'==========================================================
' modGrade - grade de consulta nas abas Clientes, Contatos
'            e Oportunidades.
'
' Decisoes que nao se alteram sem perder a garantia:
'
'  1. A grade e SOMENTE LEITURA. Digitar na celula nao passa
'     por validacao nem pelo controle de versao - foi isso
'     que corrompeu a planilha compartilhada. A aba fica
'     PROTEGIDA e so a celula de pesquisa aceita digitacao.
'
'  2. Nao existe celula mesclada na area de dados. AutoFiltro
'     e ordenacao nao funcionam sobre mesclagem.
'
'  3. UMA consulta por clique em Atualizar. Quem filtra dali
'     em diante e o Excel, em memoria.
'
'  4. O vinculo linha->registro e uma COLUNA id OCULTA dentro
'     da tabela, nunca a posicao da linha na planilha. O id
'     viaja com a linha em qualquer ordenacao ou filtro. A
'     versao anterior mapeava posicao de linha para posicao
'     na matriz da consulta, e era isso que perdia a
'     referencia ao ordenar.
'
'  5. Ordenar NAO grava nada no banco. Tabela relacional nao
'     tem ordem de linha; ordem e visual de cada usuario. Com
'     a aba protegida o menu do filtro nao ordena, entao o
'     clique no cabecalho desprotege, ordena e reprotege.
'
'  6. NAO existe coluna de icone. Duplo clique na linha abre
'     a ficha, e e na ficha que se ativa, desativa ou exclui
'     o registro. Coluna de icone custava largura em toda
'     linha para repetir o que o duplo clique ja fazia.
'
'  7. SITUACAO (Clientes e Contatos): a aba abre em Ativos. A
'     celula com lista ao lado da pesquisa (Ativos, Inativos,
'     Todos) filtra a coluna oculta _a, como a pesquisa filtra
'     _b; inativo aparece em vermelho. A consulta continua
'     trazendo tudo: quem filtra e o Excel, em memoria.
'==========================================================
Option Explicit

Public Const LIN_MODO As Long = 1       ' MINIMO / COMPLETO + estado da ordenacao
Public Const LIN_TITULO As Long = 2
Public Const LIN_BOTOES As Long = 3
Public Const LIN_BUSCA As Long = 4
Public Const LIN_STATUS As Long = 5
Public Const LIN_CAB As Long = 7
Public Const COL_INI As Long = 2        ' coluna A e so respiro
Public Const COL_SITUACAO As Long = 6   ' celula de situacao: F e G mescladas, na linha da busca
Private Const COL_LISTA As Long = 70    ' itens da lista de situacao, na linha LIN_MODO (4 pt)

Private Const SENHA As String = "zpm"   ' travamento contra digitacao, nao e seguranca

' Paleta da marca, extraida do logo institucional
' Cores: modTema (gerado de layout\tema.json), a mesma paleta das telas.

Private Const CAB_ID As String = "_id"
Private Const CAB_BUSCA As String = "_b"
Private Const CAB_ATIVO As String = "_a"         ' ATIVO / INATIVO (so clientes e contatos)
Private Const SIT_PADRAO As String = "Ativos"

Private mDados As Variant
Private mCampos As Variant
Private mIndice As Object              ' nome da coluna -> posicao em mDados

'==========================================================
' PONTOS DE ENTRADA DOS BOTOES
'==========================================================
Public Sub GradeClientes()
    AbrirGrade "Clientes"
End Sub

Public Sub GradeContatos()
    AbrirGrade "Contatos"
End Sub

Public Sub GradeOportunidades()
    AbrirGrade "Oportunidades"
End Sub

Public Sub GradeAtualizar()
    If Not modMenu.Pronto() Then Exit Sub
    Montar ActiveSheet
End Sub

Public Sub GradeVerTudo()
    Dim ws As Worksheet
    If Not modMenu.Pronto() Then Exit Sub
    Set ws = ActiveSheet
    If TabelaDaAba(ws) = "" Then Exit Sub
    Desproteger ws
    If ModoDe(ws) = "COMPLETO" Then
        ws.Cells(LIN_MODO, COL_INI).Value2 = "MINIMO"
    Else
        ws.Cells(LIN_MODO, COL_INI).Value2 = "COMPLETO"
    End If
    Montar ws
End Sub

Public Sub GradeLimparFiltros()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    If TabelaDaAba(ws) = "" Then Exit Sub
    Desproteger ws
    On Error Resume Next
    Application.EnableEvents = False
    ws.Cells(LIN_BUSCA, COL_INI + 1).Value2 = ""
    ws.Cells(LIN_MODO, COL_INI + 1).Value2 = ""
    If ws.ListObjects.Count > 0 Then ws.ListObjects(1).AutoFilter.ShowAllData
    ' "limpo" e o estado de abertura: so os ativos
    If TemSituacao(TabelaDaAba(ws)) Then
        ws.Cells(LIN_BUSCA, COL_SITUACAO).Value2 = SIT_PADRAO
        If ws.ListObjects.Count > 0 Then AplicarSituacao ws, ws.ListObjects(1)
    End If
    Application.EnableEvents = True
    On Error GoTo 0
    Proteger ws
    Status ws
End Sub

Public Sub GradeAbrirFicha()
    AbrirFichaDaLinha ActiveSheet, Selection
End Sub

Public Sub GradeVoltar()
    ThisWorkbook.Worksheets("Inicio").Activate
End Sub

'==========================================================
' ABRIR / MONTAR
'==========================================================
Private Sub AbrirGrade(ByVal nomeAba As String)
    Dim ws As Worksheet
    If Not modMenu.Pronto() Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(nomeAba)
    ws.Activate
    Montar ws
End Sub

Public Function TabelaDaAba(ByVal ws As Object) As String
    Select Case ws.Name
        Case "Clientes":      TabelaDaAba = "clientes"
        Case "Contatos":      TabelaDaAba = "contatos"
        Case "Oportunidades": TabelaDaAba = "oportunidades"
    End Select
End Function

Private Function ModoDe(ByVal ws As Worksheet) As String
    ModoDe = UCase$(Trim$(CStr(modDB.Nz(ws.Cells(LIN_MODO, COL_INI).Value2, "MINIMO"))))
    If ModoDe <> "COMPLETO" Then ModoDe = "MINIMO"
End Function

'----------------------------------------------------------
' Refaz a aba inteira: cabecalho da marca, consulta, tabela
' do Excel, colunas ocultas de controle, formatos, protecao.
'----------------------------------------------------------
Private Sub Montar(ByVal ws As Worksheet)
    Dim tabela As String, sql As String, params As Variant
    Dim cols As Variant, arr As Variant, p As Variant
    Dim i As Long, j As Long, nc As Long, nl As Long, nt As Long
    Dim rng As Range, lo As ListObject
    Dim completo As Boolean, txt As String
    Dim criterios As Variant, buscaAnterior As String, ordemAnterior As String
    Dim situacaoAnterior As String

    tabela = TabelaDaAba(ws)
    If tabela = "" Then Exit Sub

    On Error GoTo trata
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Cursor = 2                       ' xlWait

    Desproteger ws
    completo = (ModoDe(ws) = "COMPLETO")
    cols = modSchema.ColunasPlanilha(tabela, completo)
    nc = UBound(cols) - LBound(cols) + 1
    nt = nc + 3                                  ' + _b + _id + _a

    sql = modCRM.SQLGrade(tabela, params)
    mDados = modDB.Consultar(sql, params)
    mCampos = modDB.gCampos
    Set mIndice = modDB.IndiceDeCampos()
    If IsEmpty(mDados) Then nl = 0 Else nl = UBound(mDados, 1)

    criterios = GuardarFiltros(ws)
    buscaAnterior = CStr(modDB.Nz(ws.Cells(LIN_BUSCA, COL_INI + 1).Value2))
    ordemAnterior = CStr(modDB.Nz(ws.Cells(LIN_MODO, COL_INI + 1).Value2))
    situacaoAnterior = SituacaoValida(ws.Cells(LIN_BUSCA, COL_SITUACAO).Value2)

    Limpar ws
    Cabecalho ws, tabela, completo

    ReDim arr(1 To nl + 1, 1 To nt)
    For j = 1 To nc
        arr(1, j) = modSchema.Parte(CStr(cols(LBound(cols) + j - 1)), 1)
    Next j
    arr(1, nc + 1) = CAB_BUSCA
    arr(1, nc + 2) = CAB_ID
    arr(1, nc + 3) = CAB_ATIVO

    For i = 1 To nl
        txt = ""
        For j = 1 To nc
            p = Split(CStr(cols(LBound(cols) + j - 1)), "|")
            arr(i + 1, j) = CelulaDaColuna(i, CStr(p(0)), CStr(p(3)))
            txt = txt & " " & CStr(arr(i + 1, j))
        Next j
        ' A pesquisa continua encontrando o que NAO esta na tela:
        ' observacoes, proxima acao, maquina, CNPJ e e-mail entram na
        ' coluna oculta de busca mesmo sem coluna visivel.
        txt = txt & " " & modDB.Nz(CampoDaLinha(i, "observacoes")) & _
                    " " & modDB.Nz(CampoDaLinha(i, "prox_acao")) & _
                    " " & modDB.Nz(CampoDaLinha(i, "maquina")) & _
                    " " & modDB.Nz(CampoDaLinha(i, "cnpj")) & _
                    " " & modDB.Nz(CampoDaLinha(i, "email")) & _
                    " " & modDB.Nz(CampoDaLinha(i, "categoria"))
        arr(i + 1, nc + 1) = UCase$(txt)
        arr(i + 1, nc + 2) = CLng(modDB.Nz(CampoDaLinha(i, "id"), "0"))
        arr(i + 1, nc + 3) = TextoAtivo(i, tabela)
    Next i

    Set rng = ws.Range(ws.Cells(LIN_CAB, COL_INI), ws.Cells(LIN_CAB + nl, COL_INI + nt - 1))
    rng.Value2 = arr

    ' xlSrcRange = 1, xlYes = 1
    Set lo = ws.ListObjects.Add(1, rng, , 1)
    lo.Name = "tbl" & ws.Name
    lo.TableStyle = "TableStyleLight1"
    lo.ShowTableStyleRowStripes = True

    Formatar ws, lo, cols, nc, nt, nl, tabela
    CaixaPesquisa ws
    CaixaSituacao ws, tabela, situacaoAnterior

    ' Filtro, pesquisa e ordenacao sobrevivem ao Atualizar: refazer
    ' tres filtros a cada clique era o que tornava a edicao em
    ' sequencia impraticavel.
    ws.Cells(LIN_MODO, COL_INI + 1).Value2 = ordemAnterior
    ws.Cells(LIN_BUSCA, COL_INI + 1).Value2 = buscaAnterior
    ReaplicarOrdem ws, lo, ordemAnterior
    ReaplicarFiltros ws, lo, criterios
    ReaplicarPesquisa ws, lo, buscaAnterior
    If TemSituacao(tabela) Then AplicarSituacao ws, lo
    Status ws
    Congelar ws
    Proteger ws

    Application.Cursor = -4143                   ' xlDefault
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

trata:
    Dim sErro As String
    sErro = Err.Description
    Application.Cursor = -4143
    ws.Cells(LIN_STATUS, COL_INI).Value2 = "Erro ao montar a grade: " & sErro
    ' a aba NUNCA fica desprotegida: grade editavel e o que corrompia a 3.6
    Proteger ws
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Nao foi possivel montar a grade." & vbCrLf & vbCrLf & sErro, _
           vbExclamation, "CRM Zapromaq"
End Sub

'----------------------------------------------------------
' Zera a aba antes de montar.
'
' Ordem que NAO se altera: apagar a TABELA (Delete, nao Unlist),
' desligar o filtro e REEXIBIR linhas e colunas antes de limpar.
' Com filtro ou pesquisa ativos as linhas ficam ocultas, e
' escrever a matriz numa faixa com linhas ocultas espalhava os
' valores: a grade saia com o cabecalho da consulta anterior e
' colunas #N/D (era o defeito do "Ver tudo").
'----------------------------------------------------------
Private Sub Limpar(ByVal ws As Worksheet)
    Dim i As Long
    On Error Resume Next
    For i = ws.ListObjects.Count To 1 Step -1
        ws.ListObjects(i).Delete
    Next i
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Cells.Rows.Hidden = False
    ws.Cells.Columns.Hidden = False
    On Error GoTo 0
    ws.Cells.Clear
    ws.Cells.FormatConditions.Delete
    ws.Cells.Locked = True
End Sub

'----------------------------------------------------------
' Faixa de identidade visual - paleta do logo institucional:
' azul #172A67 no titulo e no cabecalho, verde #58B030 no
' realce, cinza claro de fundo.
'----------------------------------------------------------
Private Sub Cabecalho(ByVal ws As Worksheet, ByVal tabela As String, ByVal completo As Boolean)
    ws.Columns(1).ColumnWidth = 2.25
    ws.Rows(LIN_MODO).RowHeight = 4
    ws.Rows(LIN_TITULO).RowHeight = 28
    ws.Rows(LIN_BOTOES).RowHeight = 34
    ws.Rows(LIN_BUSCA).RowHeight = 22
    ws.Rows(LIN_STATUS).RowHeight = 16
    ws.Rows(6).RowHeight = 10

    ws.Cells(LIN_MODO, COL_INI).Value2 = IIf(completo, "COMPLETO", "MINIMO")
    ws.Range(ws.Cells(LIN_MODO, COL_INI), ws.Cells(LIN_MODO, COL_INI + 1)).Font.Color = modTema.COR_BRANCO

    TituloAoLadoDoLogo ws, modSchema.TituloDe(tabela)

    With ws.Cells(LIN_BUSCA, COL_INI)
        .Value2 = "Pesquisar"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Bold = True
        .Font.Color = modTema.COR_AZUL
        .HorizontalAlignment = -4152             ' xlRight
    End With
    With ws.Cells(LIN_BUSCA, COL_INI + 1)
        .Value2 = ""
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Interior.Color = modTema.COR_BRANCO
        .Locked = False                          ' unica celula digitavel da aba
        .Borders.Color = modTema.COR_VERDE
        .Borders.Weight = 2                      ' xlThin
    End With

    With ws.Cells(LIN_STATUS, COL_INI)
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Color = modTema.COR_TEXTO2
    End With

    ws.Range(ws.Cells(1, 1), ws.Cells(6, 80)).Interior.Color = modTema.COR_FUNDO
    ws.Cells(LIN_BUSCA, COL_INI + 1).Interior.Color = modTema.COR_BRANCO
End Sub

'----------------------------------------------------------
' CAIXA DE PESQUISA
'
' A caixa ocupava UMA celula, e a largura dessa celula muda
' de aba para aba - em Clientes a coluna Empresa tem 50 e em
' Oportunidades a Situacao tem 18. Resultado: a mesma caixa
' aparecia de tres tamanhos.
'
' Agora o span e fixo - colunas C a E nas tres abas.
' Mesclagem na FAIXA DE CABECALHO e permitida;
' o que nao pode e mesclar na area de dados, que quebra
' AutoFiltro e ordenacao.
'----------------------------------------------------------
'----------------------------------------------------------
' TITULO DA ABA
'
' E uma FORMA, nao uma celula: posicionada em pontos logo ao
' lado do logo. Em celula, a posicao dependia da largura das
' colunas - que muda de aba para aba e ao montar a grade -, e o
' titulo ora ficava embaixo do logo, ora longe demais dele.
'----------------------------------------------------------
Private Sub TituloAoLadoDoLogo(ByVal ws As Worksheet, ByVal texto As String)
    Dim sh As Shape, esq As Double, topo As Double, cx As Shape
    On Error Resume Next
    ws.Shapes("tituloAba").Delete
    esq = 92: topo = ws.Rows(LIN_TITULO).Top
    For Each sh In ws.Shapes
        If sh.Name = "logoAba" Then esq = sh.Left + sh.Width + 12
    Next sh
    ' msoTextOrientationHorizontal = 1
    Set cx = ws.Shapes.AddTextbox(1, esq, topo, 460, ws.Rows(LIN_TITULO).Height)
    cx.Name = "tituloAba"
    cx.Line.Visible = False
    cx.Fill.Visible = False
    cx.Placement = 3                                  ' xlFreeFloating
    With cx.TextFrame2
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
        .VerticalAnchor = 3                           ' msoAnchorMiddle
        .WordWrap = False
        With .TextRange
            .Text = texto
            .Font.Name = modTema.FONTE_SEMIBOLD
            .Font.Size = 15
            .Font.Bold = True
            .Font.Fill.ForeColor.RGB = modTema.COR_AZUL
        End With
    End With
    On Error GoTo 0
End Sub

Private Sub CaixaPesquisa(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Range(ws.Cells(LIN_BUSCA, COL_INI), ws.Cells(LIN_BUSCA, COL_INI + 40)).UnMerge

    ' Span FIXO: rotulo em B, caixa de C ate E, igual nas tres abas.
    ' A primeira versao mesclava ate somar uma largura em pontos, e como
    ' a largura das colunas muda por aba, a caixa mudava de tamanho.
    With ws.Range(ws.Cells(LIN_BUSCA, COL_INI + 1), ws.Cells(LIN_BUSCA, COL_INI + 3))
        .Merge
        .Interior.Color = modTema.COR_BRANCO
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .HorizontalAlignment = -4131             ' xlLeft
        .VerticalAlignment = -4108               ' xlCenter
        .IndentLevel = 1
        .Locked = False
        .Borders.Color = modTema.COR_VERDE
        .Borders.Weight = 2
    End With
    On Error GoTo 0
End Sub

'----------------------------------------------------------
' SITUACAO DO CADASTRO (Clientes e Contatos)
'
' Celula com lista na linha da pesquisa, a direita da caixa
' (F e G mescladas; mesclagem no cabecalho e permitida). O
' valor fica na propria celula e sobrevive ao Atualizar, como
' a pesquisa. A lista aponta para 3 celulas da linha de
' controle, nao para um texto "a,b,c": referencia de faixa nao
' depende do separador de lista do Windows de cada estacao.
'----------------------------------------------------------
Private Function TemSituacao(ByVal tabela As String) As Boolean
    TemSituacao = (tabela = "clientes" Or tabela = "contatos")
End Function

' Valor da celula, se for uma das opcoes; senao o padrao (Ativos).
Private Function SituacaoValida(ByVal v As Variant) As String
    Dim op As Variant
    SituacaoValida = SIT_PADRAO
    For Each op In modCRM.OpcoesSituacao()
        If CStr(modDB.Nz(v)) = op Then SituacaoValida = op
    Next op
End Function

Private Function TextoAtivo(ByVal linha As Long, ByVal tabela As String) As String
    Dim v As Variant
    If Not TemSituacao(tabela) Then Exit Function
    v = CampoDaLinha(linha, "ativo")
    ' CBool e nao CStr: o texto de um Boolean depende do idioma do Office
    If IsNull(v) Then TextoAtivo = "ATIVO" Else TextoAtivo = IIf(CBool(v), "ATIVO", "INATIVO")
End Function

Private Sub CaixaSituacao(ByVal ws As Worksheet, ByVal tabela As String, ByVal valor As String)
    Dim faixa As Range, lista As Range
    If Not TemSituacao(tabela) Then Exit Sub
    On Error Resume Next
    Set lista = ws.Range(ws.Cells(LIN_MODO, COL_LISTA), ws.Cells(LIN_MODO, COL_LISTA + 2))
    lista.Value2 = modCRM.OpcoesSituacao()
    lista.Font.Color = modTema.COR_FUNDO
    Set faixa = ws.Range(ws.Cells(LIN_BUSCA, COL_SITUACAO), ws.Cells(LIN_BUSCA, COL_SITUACAO + 1))
    faixa.Merge
    With faixa
        .Interior.Color = modTema.COR_BRANCO
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .HorizontalAlignment = -4131             ' xlLeft
        .VerticalAlignment = -4108               ' xlCenter
        .IndentLevel = 1
        .ShrinkToFit = True
        .Locked = False
        .Borders.Color = modTema.COR_VERDE
        .Borders.Weight = 2
        ' mostra "Situacao: Ativos"; o valor da celula continua "Ativos"
        .NumberFormat = """Situa" & ChrW$(&HE7) & ChrW$(&HE3) & "o: ""@"
    End With
    ws.Cells(LIN_BUSCA, COL_SITUACAO).Value2 = valor
    With ws.Cells(LIN_BUSCA, COL_SITUACAO).Validation
        .Delete
        ' xlValidateList = 3, xlValidAlertStop = 1, xlBetween = 1
        .Add Type:=3, AlertStyle:=1, Operator:=1, Formula1:="=" & lista.Address
        .InCellDropdown = True
        .InputTitle = "Situa" & ChrW$(&HE7) & ChrW$(&HE3) & "o"
        .InputMessage = "Ativos, Inativos ou Todos."
    End With
    On Error GoTo 0
End Sub

' Filtra a coluna _a pela celula de situacao (aba desprotegida por quem chama).
Private Sub AplicarSituacao(ByVal ws As Worksheet, ByVal lo As ListObject)
    Dim idx As Long
    idx = IndiceColuna(lo, CAB_ATIVO)
    If idx = 0 Then Exit Sub
    On Error Resume Next
    Select Case SituacaoValida(ws.Cells(LIN_BUSCA, COL_SITUACAO).Value2)
        Case "Ativos":   lo.Range.AutoFilter Field:=idx, Criteria1:="ATIVO"
        Case "Inativos": lo.Range.AutoFilter Field:=idx, Criteria1:="INATIVO"
        Case Else:       lo.Range.AutoFilter Field:=idx
    End Select
    On Error GoTo 0
End Sub

' Mudou a celula de situacao (Workbook_SheetChange).
Public Sub FiltrarSituacao(ByVal ws As Object)
    If Not TemSituacao(TabelaDaAba(ws)) Then Exit Sub
    If ws.ListObjects.Count = 0 Then Exit Sub
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Desproteger ws
    AplicarSituacao ws, ws.ListObjects(1)
    Proteger ws
    Status ws
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Private Sub Status(ByVal ws As Worksheet)
    Dim lo As ListObject, total As Long, visiveis As Long, t As String
    If ws.ListObjects.Count = 0 Then Exit Sub
    Set lo = ws.ListObjects(1)
    If Not lo.DataBodyRange Is Nothing Then
        total = lo.DataBodyRange.Rows.Count
        ' SUBTOTAL(103) conta so as linhas visiveis, em TODOS os blocos.
        ' SpecialCells(visivel).Rows.Count contava so o primeiro bloco:
        ' com filtro, o "X de Y" saia errado.
        On Error Resume Next
        visiveis = total
        visiveis = CLng(Application.WorksheetFunction.Subtotal(103, _
                        lo.ListColumns(IndiceColuna(lo, CAB_ID)).DataBodyRange))
        On Error GoTo 0
    End If
    t = CStr(modDB.Nz(ws.Cells(LIN_BUSCA, COL_INI + 1).Value2))
    ws.Cells(LIN_STATUS, COL_INI).Font.Color = modTema.COR_TEXTO2
    ws.Cells(LIN_STATUS, COL_INI).Value2 = _
        Format$(visiveis, "#,##0") & " de " & Format$(total, "#,##0") & " registro(s)  " & _
        ChrW$(&HB7) & "  consulta de " & Format$(Now, "dd/mm/yyyy hh:nn") & "  " & _
        ChrW$(&HB7) & "  " & IIf(ModoDe(ws) = "COMPLETO", "todas as colunas", "colunas de trabalho") & _
        IIf(t = "", "", "  " & ChrW$(&HB7) & "  pesquisa: " & t) & _
        IIf(TemSituacao(TabelaDaAba(ws)), "  " & ChrW$(&HB7) & "  " & _
            LCase$(SituacaoValida(ws.Cells(LIN_BUSCA, COL_SITUACAO).Value2)), "") & "  " & ChrW$(&HB7) & _
        "  clique no cabecalho ordena  " & ChrW$(&HB7) & "  duplo clique abre a ficha"
End Sub

Private Sub Formatar(ByVal ws As Worksheet, ByVal lo As ListObject, ByVal cols As Variant, _
                     ByVal nc As Long, ByVal nt As Long, ByVal nl As Long, ByVal tabela As String)
    Dim j As Long, p As Variant, col As Range, colSit As Long

    With lo.HeaderRowRange
        .Interior.Color = modTema.COR_AZUL
        .Font.Name = "Segoe UI"
        .Font.Color = modTema.COR_BRANCO
        .Font.Bold = True
        .HorizontalAlignment = -4131             ' xlLeft
        .RowHeight = 22
        .VerticalAlignment = -4108               ' xlCenter
    End With

    For j = 1 To nc
        p = Split(CStr(cols(LBound(cols) + j - 1)), "|")
        Set col = ws.Columns(COL_INI + j - 1)
        col.ColumnWidth = CDbl(p(2))
        Select Case CStr(p(3))
            Case "C": col.NumberFormat = "#,##0.00"
            Case "D": col.NumberFormat = "dd/mm/yyyy"
            Case "N": col.NumberFormat = "0"
        End Select
        If CStr(p(0)) = "situacao" Then colSit = j
    Next j

    ' colunas de controle: ocultas, nunca apagadas
    ws.Columns(COL_INI + nc).Hidden = True       ' _b
    ws.Columns(COL_INI + nc + 1).Hidden = True    ' _id
    ws.Columns(COL_INI + nc + 2).Hidden = True    ' _a

    If nl > 0 Then
        With lo.DataBodyRange
            .Font.Name = "Segoe UI"
            .Font.Size = 10
            .VerticalAlignment = -4108           ' xlCenter
            .WrapText = False
            ' reduzir para caber: nome comprido encolhe na celula em vez
            ' de sumir atras da coluna seguinte ou virar #####
            .ShrinkToFit = True
            .RowHeight = 15
            .Locked = True
        End With
    End If

    Condicionais ws, lo, cols, nc, nl, tabela
End Sub

'----------------------------------------------------------
' FORMATACAO CONDICIONAL
'
' Reproduz o que a 3.6 comunicava, sem o excesso dela: la
' eram 7 cores de situacao MAIS uma cor de fundo por etapa
' pintando a linha inteira, em 11 tons. Tabela toda colorida
' para de comunicar - o olho nao acha mais o que e urgente.
'
' Aqui o destaque e da ACAO, que e o que se trabalha:
' atrasada, retomar hoje e nesta semana tem fundo; o resto
' vive de cor de texto. Perdido e Descartado ESMAECEM a linha
' inteira em cinza, para sair da frente sem sair da tabela.
'
' Sobre a implementacao: a versao anterior usava
' FormatConditions.Add(xlTextString, , texto) e nao pintava
' nada - o tipo texto exige o operador, e o erro ficava
' engolido pelo On Error Resume Next. Agora e xlExpression
' com formula, que nao depende de parametro opcional em
' ordem certa.
'----------------------------------------------------------
Private Sub Condicionais(ByVal ws As Worksheet, ByVal lo As ListObject, _
                         ByVal cols As Variant, ByVal nc As Long, ByVal nl As Long, _
                         ByVal tabela As String)
    Dim jSit As Long, jEta As Long, j As Long, p As Variant
    Dim colSit As Range, lin1 As Long, refSit As String, refEta As String

    If nl = 0 Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    ' Cadastro inativo: a linha inteira em vermelho (cor perigo do tema).
    ' Formula simples, sem funcao: nao depende do idioma do Excel.
    If TemSituacao(tabela) Then
        RegraLinha lo.DataBodyRange, "=$" & LetraDaColuna(ws, COL_INI + nc + 2) & lo.DataBodyRange.Row & _
                   "=""INATIVO""", -1, modTema.COR_PERIGO, False
        Exit Sub
    End If
    If tabela <> "oportunidades" Then Exit Sub

    For j = 1 To nc
        p = Split(CStr(cols(LBound(cols) + j - 1)), "|")
        If CStr(p(0)) = "situacao" Then jSit = j
        If CStr(p(0)) = "etapa" Then jEta = j
    Next j

    lin1 = lo.DataBodyRange.Row
    If jEta > 0 Then refEta = "$" & LetraDaColuna(ws, COL_INI + jEta - 1) & lin1
    If jSit > 0 Then refSit = "$" & LetraDaColuna(ws, COL_INI + jSit - 1) & lin1

    ' 1. linha esmaecida: negocio que saiu. Entra primeiro para
    '    ficar EMBAIXO das regras de acao na ordem de prioridade.
    ' Duas regras simples, sem OR(): a formula de FormatConditions e
    ' lida no idioma e no separador do Excel instalado (OU e ; no
    ' Excel em portugues) - com OR a regra falhava calada e Perdido
    ' e Descartado nunca esmaeciam.
    If refEta <> "" Then
        RegraLinha lo.DataBodyRange, "=" & refEta & "=""Perdido""", -1, modTema.COR_DESABILITADO_TEXTO, False
        RegraLinha lo.DataBodyRange, "=" & refEta & "=""Descartado""", -1, modTema.COR_DESABILITADO_TEXTO, False
    End If

    If refSit = "" Then Exit Sub
    Set colSit = lo.ListColumns(jSit).DataBodyRange

    ' 2. situacao - do menos para o mais urgente
    RegraLinha colSit, "=" & refSit & "=""Encerrado""", _
               -1, modTema.CorSituacao("Encerrado", "texto"), False
    RegraLinha colSit, "=" & refSit & "=""Em dia""", _
               -1, modTema.CorSituacao("Em dia", "texto"), False
    RegraLinha colSit, "=" & refSit & "=""Sem etapa definida""", _
               -1, modTema.CorSituacao("Sem etapa definida", "texto"), True
    RegraLinha colSit, "=" & refSit & "=""" & TXT_SEM_ACAO() & """", _
               modTema.CorSituacao(TXT_SEM_ACAO(), "fundo"), modTema.CorSituacao(TXT_SEM_ACAO(), "texto"), True
    RegraLinha colSit, "=" & refSit & "=""" & TXT_SEMANA() & """", _
               modTema.CorSituacao(TXT_SEMANA(), "fundo"), modTema.CorSituacao(TXT_SEMANA(), "texto"), False
    RegraLinha colSit, "=" & refSit & "=""Retomar hoje""", _
               modTema.CorSituacao("Retomar hoje", "fundo"), modTema.CorSituacao("Retomar hoje", "texto"), True
    RegraLinha colSit, "=" & refSit & "=""" & TXT_ATRASADA() & """", _
               modTema.CorSituacao(TXT_ATRASADA(), "fundo"), modTema.CorSituacao(TXT_ATRASADA(), "texto"), True
End Sub

' Os textos vem de modCalc.Situacao. Montados por ChrW para a
' fonte seguir em ASCII puro: acento gravado direto no .bas
' depende da pagina de codigo na importacao.
Private Function TXT_ATRASADA() As String
    TXT_ATRASADA = "A" & ChrW$(&HE7) & ChrW$(&HE3) & "o atrasada"
End Function

Private Function TXT_SEMANA() As String
    TXT_SEMANA = "A" & ChrW$(&HE7) & ChrW$(&HE3) & "o nesta semana"
End Function

Private Function TXT_SEM_ACAO() As String
    TXT_SEM_ACAO = "Sem pr" & ChrW$(&HF3) & "xima a" & ChrW$(&HE7) & ChrW$(&HE3) & "o"
End Function

Private Function LetraDaColuna(ByVal ws As Worksheet, ByVal indice As Long) As String
    Dim s As String
    s = ws.Cells(1, indice).Address(True, False)
    LetraDaColuna = Left$(s, InStr(s, "$") - 1)
End Function

'----------------------------------------------------------
' fundo = -1 significa nao mexer no preenchimento: cor de
' texto sozinha basta e deixa a grade sobria.
'----------------------------------------------------------
Private Sub RegraLinha(ByVal alvo As Range, ByVal formula As String, _
                       ByVal fundo As Long, ByVal cor As Long, ByVal negrito As Boolean)
    Dim fc As Object
    On Error Resume Next
    Set fc = alvo.FormatConditions.Add(2, , formula)     ' xlExpression = 2
    If fc Is Nothing Then Exit Sub
    If fundo <> -1 Then fc.Interior.Color = fundo
    fc.Font.Color = cor
    fc.Font.Bold = negrito
    fc.StopIfTrue = False
    On Error GoTo 0
End Sub

Private Sub Congelar(ByVal ws As Worksheet)
    On Error Resume Next
    If Not ActiveSheet Is ws Then Exit Sub
    ActiveWindow.FreezePanes = False
    ws.Cells(LIN_CAB + 1, COL_INI).Select
    ActiveWindow.FreezePanes = True
    On Error GoTo 0
End Sub

'==========================================================
' PROTECAO
' UserInterfaceOnly nao sobrevive a fechar e reabrir o
' arquivo, entao nao se confia nele: toda rotina que mexe na
' aba desprotege e reprotege explicitamente.
'==========================================================
Public Sub Desproteger(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Unprotect SENHA
    On Error GoTo 0
End Sub

Public Sub Proteger(ByVal ws As Worksheet)
    On Error Resume Next
    ws.EnableSelection = 0                       ' xlNoRestrictions
    ws.Protect Password:=SENHA, DrawingObjects:=False, Contents:=True, Scenarios:=False, _
               AllowFiltering:=True, AllowSorting:=False, UserInterfaceOnly:=True
    On Error GoTo 0
End Sub

'==========================================================
' ORDENACAO POR SCRIPT
' Planilha protegida nao aceita ordenar pelo menu do filtro
' nem com AllowSorting. Clique no cabecalho: desprotege,
' ordena, reprotege. Nada disso toca no banco.
'==========================================================
Public Sub OrdenarPorCabecalho(ByVal ws As Object, ByVal alvo As Object)
    Dim lo As ListObject, j As Long, nc As Long
    Dim estado As String, dir As Long, cab As String

    If TabelaDaAba(ws) = "" Then Exit Sub
    If ws.ListObjects.Count = 0 Then Exit Sub
    Set lo = ws.ListObjects(1)
    If alvo.Cells.Count > 1 Then Exit Sub
    If Intersect(alvo, lo.HeaderRowRange) Is Nothing Then Exit Sub

    j = alvo.Column - lo.Range.Column + 1
    nc = lo.ListColumns.Count
    cab = CStr(lo.ListColumns(j).Name)
    ' coluna de controle nao ordena
    If cab = CAB_BUSCA Or cab = CAB_ID Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    estado = CStr(modDB.Nz(ws.Cells(LIN_MODO, COL_INI + 1).Value2))
    If estado = j & "|1" Then dir = 2 Else dir = 1      ' xlDescending / xlAscending

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Desproteger ws
    On Error Resume Next
    With lo.Sort
        .SortFields.Clear
        .SortFields.Add Key:=lo.ListColumns(j).DataBodyRange, Order:=dir
        .Header = 1                              ' xlYes
        .MatchCase = False
        .Apply
    End With
    On Error GoTo 0
    ws.Cells(LIN_MODO, COL_INI + 1).Value2 = j & "|" & dir
    Proteger ws
    ws.Cells(LIN_CAB + 1, COL_INI).Select
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

'==========================================================
' PESQUISA
' Filtra a coluna oculta _b, que concatena o conteudo
' visivel da linha. AutoFiltro nativo nao sabe procurar em
' qualquer coluna; essa coluna resolve com um filtro so.
'==========================================================
Public Sub Pesquisar(ByVal ws As Object, ByVal alvo As Object)
    Dim lo As ListObject, t As String, idx As Long

    If TabelaDaAba(ws) = "" Then Exit Sub
    If Intersect(alvo, ws.Cells(LIN_BUSCA, COL_INI + 1)) Is Nothing Then Exit Sub
    If ws.ListObjects.Count = 0 Then Exit Sub
    Set lo = ws.ListObjects(1)
    idx = IndiceColuna(lo, CAB_BUSCA)
    If idx = 0 Then Exit Sub

    t = TextoFiltro(CStr(modDB.Nz(ws.Cells(LIN_BUSCA, COL_INI + 1).Value2)))

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Desproteger ws
    On Error Resume Next
    If t = "" Then
        lo.Range.AutoFilter Field:=idx
    Else
        lo.Range.AutoFilter Field:=idx, Criteria1:="*" & t & "*"
    End If
    On Error GoTo 0
    Proteger ws
    Status ws
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Private Function IndiceColuna(ByVal lo As ListObject, ByVal nome As String) As Long
    Dim j As Long
    For j = 1 To lo.ListColumns.Count
        If CStr(lo.ListColumns(j).Name) = nome Then
            IndiceColuna = j
            Exit Function
        End If
    Next j
End Function

'==========================================================
' CLIQUE NOS ICONES
'==========================================================
Public Sub CliqueNaAba(ByVal ws As Object, ByVal alvo As Object)
    Dim lo As ListObject

    If Not Application.EnableEvents Then Exit Sub
    If TabelaDaAba(ws) = "" Then Exit Sub
    If ws.ListObjects.Count = 0 Then Exit Sub
    If alvo.Cells.Count > 1 Then Exit Sub
    Set lo = ws.ListObjects(1)

    If Not Intersect(alvo, lo.HeaderRowRange) Is Nothing Then OrdenarPorCabecalho ws, alvo
End Sub

'==========================================================
' MENU DO BOTAO DIREITO
' Ativar, desativar e excluir NAO entram aqui: sao acoes que
' gravam, e gravacao acontece na ficha, onde se ve o registro
' inteiro antes de decidir.
'==========================================================
Public Sub CopiarSelecao()
    On Error Resume Next
    Selection.Copy
    On Error GoTo 0
End Sub

' A celula esta na area de dados da grade? (o menu proprio so
' substitui o do Excel ali; no resto da aba o menu padrao vale)
Public Function NaAreaDeDados(ByVal ws As Object, ByVal alvo As Object) As Boolean
    Dim lo As ListObject
    If TabelaDaAba(ws) = "" Then Exit Function
    If ws.ListObjects.Count = 0 Then Exit Function
    Set lo = ws.ListObjects(1)
    If lo.DataBodyRange Is Nothing Then Exit Function
    NaAreaDeDados = Not (Intersect(alvo, lo.DataBodyRange) Is Nothing)
End Function

Public Sub MenuDaLinha(ByVal ws As Object, ByVal alvo As Object)
    Dim lo As ListObject, bar As Object, bt As Object

    If ws.ListObjects.Count = 0 Then Exit Sub
    Set lo = ws.ListObjects(1)
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If Intersect(alvo, lo.DataBodyRange) Is Nothing Then Exit Sub

    On Error Resume Next
    Application.CommandBars("mnuCRMLinha").Delete
    Set bar = Application.CommandBars.Add(Name:="mnuCRMLinha", Position:=5, Temporary:=True)
    If bar Is Nothing Then Exit Sub

    Set bt = bar.Controls.Add(Type:=1)
    bt.Caption = "Abrir ficha"
    bt.OnAction = "modGrade.GradeAbrirFicha"

    ' copiar telefone, e-mail ou linhas inteiras: a aba protegida
    ' deixa copiar, mas o menu padrao estava bloqueado
    Set bt = bar.Controls.Add(Type:=1)
    bt.Caption = "Copiar"
    bt.FaceId = 19
    bt.OnAction = "modGrade.CopiarSelecao"

    Set bt = bar.Controls.Add(Type:=1)
    bt.Caption = "Atualizar a grade"
    bt.BeginGroup = True
    bt.OnAction = "modGrade.GradeAtualizar"

    Set bt = bar.Controls.Add(Type:=1)
    bt.Caption = "Limpar filtros e pesquisa"
    bt.OnAction = "modGrade.GradeLimparFiltros"

    bar.ShowPopup
    On Error GoTo 0
End Sub

'==========================================================
' LEITURA DA MATRIZ DA CONSULTA
'==========================================================
' Busca da coluna pelo dicionario: montar 500 linhas x 20 colunas
' varrendo a lista de nomes a cada celula custava centenas de
' milhares de comparacoes de texto.
Private Function CampoDaLinha(ByVal linha As Long, ByVal nome As String) As Variant
    CampoDaLinha = Null
    If mIndice Is Nothing Then Exit Function
    If mIndice.Exists(nome) Then CampoDaLinha = mDados(linha, mIndice(nome))
End Function

'----------------------------------------------------------
' Numero e data saem como numero e data, nunca como texto:
' e o que faz o filtro "maior que", a ordenacao e a soma da
' barra de status funcionarem de verdade.
'----------------------------------------------------------
Private Function CelulaDaColuna(ByVal linha As Long, ByVal nome As String, _
                               ByVal tipo As String) As Variant
    Dim v As Variant

    Select Case nome
        Case "situacao"
            CelulaDaColuna = modCalc.Situacao(True, CStr(modDB.Nz(CampoDaLinha(linha, "etapa"))), _
                                             CampoDaLinha(linha, "dt_prox_acao"), _
                                             CampoDaLinha(linha, "retomar_em"))
            Exit Function
        Case "codigo_visual"
            CelulaDaColuna = modCRM.CodigoVisualCliente(CampoDaLinha(linha, "codigo_cliente"))
            Exit Function
    End Select

    v = CampoDaLinha(linha, nome)
    If IsNull(v) Then CelulaDaColuna = "": Exit Function

    Select Case tipo
        Case "C", "N"
            If IsNumeric(v) Then CelulaDaColuna = CDbl(v) Else CelulaDaColuna = ""
        Case "D"
            If IsDate(v) Then CelulaDaColuna = CDate(v) Else CelulaDaColuna = ""
        Case "F"
            CelulaDaColuna = modValidacao.FormatarTelefone(v)
        Case "J"
            ' o banco guarda digitos; a mascara e so exibicao
            CelulaDaColuna = modValidacao.FormatarCNPJ(v)
        Case Else
            CelulaDaColuna = Limitar(CStr(v))
    End Select
End Function

Private Function Limitar(ByVal s As String) As String
    s = Replace(Replace(s, vbCrLf, " "), vbLf, " ")
    If Len(s) > 240 Then s = Left$(s, 237) & "..."
    Limitar = s
End Function

'==========================================================
' LINHA -> REGISTRO
' Le o id da coluna oculta da propria linha. Nao usa, em
' nenhuma hipotese, a posicao da linha na planilha.
'==========================================================
Public Function IDdaLinha(ByVal ws As Object, ByVal alvo As Object) As Long
    Dim lo As ListObject, idx As Long, v As Variant
    If ws.ListObjects.Count = 0 Then Exit Function
    Set lo = ws.ListObjects(1)
    If lo.DataBodyRange Is Nothing Then Exit Function
    If Intersect(alvo, lo.DataBodyRange) Is Nothing Then Exit Function
    idx = IndiceColuna(lo, CAB_ID)
    If idx = 0 Then Exit Function
    ' Endereco FISICO da celula, nao posicao relativa dentro da
    ' tabela: qualquer deslocamento da tabela na aba deixaria a
    ' conta relativa lendo o id da linha vizinha - e abrir a ficha
    ' do registro errado e o pior defeito que esta grade pode ter.
    v = ws.Cells(alvo.Row, lo.Range.Column + idx - 1).Value2
    If IsNumeric(v) Then IDdaLinha = CLng(v)
End Function

Public Function ValorDaLinha(ByVal ws As Object, ByVal alvo As Object, _
                             ByVal cabecalho As String) As String
    Dim lo As ListObject, idx As Long
    If ws.ListObjects.Count = 0 Then Exit Function
    Set lo = ws.ListObjects(1)
    If lo.DataBodyRange Is Nothing Then Exit Function
    idx = IndiceColuna(lo, cabecalho)
    If idx = 0 Then Exit Function
    ValorDaLinha = CStr(modDB.Nz(ws.Cells(alvo.Row, lo.Range.Column + idx - 1).Value2))
End Function

Public Sub AbrirFichaDaLinha(ByVal ws As Object, ByVal alvo As Object)
    Dim tabela As String, lo As ListObject, idReg As Long

    tabela = TabelaDaAba(ws)
    If tabela = "" Then Exit Sub
    If Not modMenu.Pronto() Then Exit Sub
    If ws.ListObjects.Count = 0 Then
        MsgBox "A grade ainda nao foi carregada. Clique em Atualizar.", _
               vbInformation, "CRM Zapromaq"
        Exit Sub
    End If

    Set lo = ws.ListObjects(1)
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If Intersect(alvo, lo.DataBodyRange) Is Nothing Then
        MsgBox "Clique numa linha da grade primeiro.", vbInformation, "CRM Zapromaq"
        Exit Sub
    End If

    idReg = IDdaLinha(ws, alvo)
    If idReg = 0 Then
        MsgBox "Nao consegui identificar o registro desta linha." & vbCrLf & _
               "Clique em Atualizar e tente de novo.", vbExclamation, "CRM Zapromaq"
        Exit Sub
    End If

    ' A ficha e modal: a execucao volta para ca quando ela fecha.
    ' Se algo foi gravado la, a grade e remontada - senao a linha
    ' continuaria mostrando o valor antigo ate alguem clicar em
    ' Atualizar, e a tela estaria mentindo.
    modConfig.gGravados = ""
    Dim f As New frmCRM
    f.Abrir tabela, 0, idReg, IDsVisiveis(ws)
    AtualizarGravados ws
End Sub

'==========================================================
' O FILTRO SOBREVIVE AO ATUALIZAR
'
' Guarda os criterios de cada coluna antes de refazer a
' tabela e devolve depois. So criterio simples e de lista
' voltam - filtro de cor e "os 10 primeiros" nao, e por isso
' a reaplicacao e silenciosa: melhor voltar sem um filtro
' raro do que estourar erro no meio de uma consulta.
'==========================================================
Private Function GuardarFiltros(ByVal ws As Worksheet) As Variant
    Dim lo As ListObject, f As Object, i As Long, guardados() As String

    If ws.ListObjects.Count = 0 Then Exit Function
    Set lo = ws.ListObjects(1)
    On Error Resume Next
    If lo.AutoFilter Is Nothing Then Exit Function
    ReDim guardados(1 To lo.ListColumns.Count, 1 To 3)
    For i = 1 To lo.ListColumns.Count
        Set f = lo.AutoFilter.Filters(i)
        ' coluna de controle (_b, _id, _a) volta pela propria celula, e o
        ' indice dela muda no "Ver tudo": reaplicada por posicao, cairia
        ' numa coluna visivel e esconderia tudo
        If Left$(CStr(lo.ListColumns(i).Name), 1) = "_" Then GoTo proxima
        If f.On Then
            guardados(i, 1) = TextoCriterio(f.Criteria1)
            guardados(i, 2) = CStr(f.Operator)
            If f.Operator = 1 Or f.Operator = 2 Then   ' xlAnd / xlOr
                guardados(i, 3) = TextoCriterio(f.Criteria2)
            End If
        End If
proxima:
    Next i
    On Error GoTo 0
    GuardarFiltros = guardados
End Function

' Criteria1 pode ser array (filtro de lista com varios itens
' marcados). Vira texto separado por barra vertical e volta
' como array na reaplicacao.
Private Function TextoCriterio(ByVal c As Variant) As String
    Dim i As Long, s As String
    On Error Resume Next
    If IsArray(c) Then
        For i = LBound(c) To UBound(c)
            s = s & IIf(s = "", "", "|") & CStr(c(i))
        Next i
        TextoCriterio = "[]" & s
    Else
        TextoCriterio = CStr(c)
    End If
    On Error GoTo 0
End Function

Private Sub ReaplicarFiltros(ByVal ws As Worksheet, ByVal lo As ListObject, ByVal criterios As Variant)
    Dim i As Long, c1 As String, op As Long

    If IsEmpty(criterios) Then Exit Sub
    If Not IsArray(criterios) Then Exit Sub

    On Error Resume Next
    For i = 1 To UBound(criterios, 1)
        If i <= lo.ListColumns.Count Then
            c1 = criterios(i, 1)
            If c1 <> "" Then
                op = 0
                If criterios(i, 2) <> "" Then op = CLng(criterios(i, 2))
                If Left$(c1, 2) = "[]" Then
                    ' xlFilterValues = 7
                    lo.Range.AutoFilter Field:=i, Criteria1:=Split(Mid$(c1, 3), "|"), Operator:=7
                ElseIf op = 1 Or op = 2 Then
                    lo.Range.AutoFilter Field:=i, Criteria1:=c1, Operator:=op, _
                                        Criteria2:=criterios(i, 3)
                Else
                    lo.Range.AutoFilter Field:=i, Criteria1:=c1
                End If
            End If
        End If
    Next i
    On Error GoTo 0
End Sub

Private Sub ReaplicarPesquisa(ByVal ws As Worksheet, ByVal lo As ListObject, ByVal texto As String)
    Dim idx As Long
    If Trim$(texto) = "" Then Exit Sub
    idx = IndiceColuna(lo, CAB_BUSCA)
    If idx = 0 Then Exit Sub
    On Error Resume Next
    lo.Range.AutoFilter Field:=idx, Criteria1:="*" & TextoFiltro(texto) & "*"
    On Error GoTo 0
End Sub

' * ? e ~ sao curingas do AutoFiltro: digitados, valem como letra.
Private Function TextoFiltro(ByVal s As String) As String
    s = UCase$(Trim$(s))
    s = Replace(s, "~", "~~")
    s = Replace(s, "*", "~*")
    s = Replace(s, "?", "~?")
    TextoFiltro = s
End Function

Private Sub ReaplicarOrdem(ByVal ws As Worksheet, ByVal lo As ListObject, ByVal estado As String)
    Dim p As Variant, j As Long, dir As Long
    If InStr(estado, "|") = 0 Then Exit Sub
    p = Split(estado, "|")
    j = CLng(Val(p(0)))
    dir = CLng(Val(p(1)))
    If j < 1 Or j > lo.ListColumns.Count Then Exit Sub
    If dir <> 1 And dir <> 2 Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    On Error Resume Next
    With lo.Sort
        .SortFields.Clear
        .SortFields.Add Key:=lo.ListColumns(j).DataBodyRange, Order:=dir
        .Header = 1
        .MatchCase = False
        .Apply
    End With
    On Error GoTo 0
End Sub

'----------------------------------------------------------
' VOLTA DA FICHA
'
' UMA consulta para todos os registros gravados, e cada linha
' reescrita de uma vez, como matriz.
'
' A primeira versao consultava a tabela inteira UMA VEZ POR
' REGISTRO: dez atendimentos editados eram dez consultas de
' 490 linhas, e a planilha ficava segundos travada ao fechar
' a ficha. Disparar isso a cada Salvar nao resolveria nada -
' seriam as mesmas dez consultas, so distribuidas, travando
' a ficha no meio da edicao. O que resolve e consultar uma
' vez e escrever em bloco.
'
' Remonta a aba inteira so quando alguma linha nao pode ser
' resolvida: registro excluido tem de sair da tabela, e isso
' a reescrita de celula nao faz. Remontar apaga filtro e
' ordenacao, entao e ultimo recurso.
'----------------------------------------------------------
Private Sub AtualizarGravados(ByVal ws As Worksheet)
    Dim ids As Variant, i As Long, pendentes As Object

    If modConfig.gGravados = "" Then Exit Sub
    ids = Split(modConfig.gGravados, ",")
    modConfig.gGravados = ""

    ' ids distintos - com "Salvar e prox." o mesmo registro pode
    ' ter sido salvo duas vezes
    Set pendentes = CreateObject("Scripting.Dictionary")
    For i = LBound(ids) To UBound(ids)
        If Trim$(CStr(ids(i))) <> "" Then
            If Not pendentes.Exists(CLng(ids(i))) Then pendentes.Add CLng(ids(i)), 0
        End If
    Next i
    AtualizarRegistros ws, pendentes, True
End Sub

'----------------------------------------------------------
' Reescreve na grade SO as linhas destes ids (dicionario com
' os ids como chave), com UMA consulta restrita a eles. Usado
' na volta da ficha e pelo feed de alteracoes (modFeed).
'
' Id que nao esta na grade (registro novo) ou nao volta da
' consulta (registro excluido):
'   remontarSeFaltar = True  -> remonta a aba (volta da ficha)
'   remontarSeFaltar = False -> devolve quantos faltaram, para
'     o feed avisar na linha de status sem tirar o usuario do lugar
'----------------------------------------------------------
Public Function AtualizarRegistros(ByVal ws As Worksheet, ByVal pendentes As Object, _
                                   ByVal remontarSeFaltar As Boolean) As Long
    Dim i As Long, j As Long, k As Long, faltaram As Long
    Dim lo As ListObject, tabela As String, sql As String, params As Variant
    Dim cols As Variant, completo As Boolean, nc As Long
    Dim idxID As Long, idsPlan As Variant, linhaPlan As Long
    Dim achou As Long, arr As Variant, txt As String, remontar As Boolean
    Dim idGravado As Long, chave As Variant, esp As Variant, lista As Variant

    If pendentes Is Nothing Then Exit Function
    If pendentes.Count = 0 Then Exit Function
    tabela = TabelaDaAba(ws)
    If tabela = "" Then Exit Function
    If ws.ListObjects.Count = 0 Then GoTo semGrade
    Set lo = ws.ListObjects(1)
    If lo.DataBodyRange Is Nothing Then GoTo semGrade
    idxID = IndiceColuna(lo, CAB_ID)
    If idxID = 0 Then GoTo semGrade

    ' UMA consulta, so dos ids pedidos (acima de 200, recarrega tudo)
    If pendentes.Count > 200 Then
        If remontarSeFaltar Then Montar ws
        AtualizarRegistros = pendentes.Count
        Exit Function
    End If
    lista = pendentes.Keys
    sql = modCRM.SQLGrade(tabela, params, lista)
    mDados = modDB.Consultar(sql, params)
    mCampos = modDB.gCampos
    Set mIndice = modDB.IndiceDeCampos()

    ' a coluna de id lida em bloco, nao celula por celula. Com UMA
    ' linha so, Value2 devolve um valor e nao uma matriz: embrulha.
    If lo.DataBodyRange.Rows.Count = 1 Then
        ReDim idsPlan(1 To 1, 1 To 1)
        idsPlan(1, 1) = lo.ListColumns(idxID).DataBodyRange.Value2
    Else
        idsPlan = lo.ListColumns(idxID).DataBodyRange.Value2
    End If

    completo = (ModoDe(ws) = "COMPLETO")
    cols = modSchema.ColunasPlanilha(tabela, completo)
    nc = UBound(cols) - LBound(cols) + 1

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Desproteger ws
    ' erro no meio da escrita: a aba NAO pode ficar desprotegida
    ' nem o Excel com eventos desligados (foi o A04)
    On Error GoTo falhaEscrita

    For Each chave In pendentes.Keys
        idGravado = CLng(chave)

        achou = 0
        If Not IsEmpty(mDados) Then
            For i = 1 To UBound(mDados, 1)
                If CLng(modDB.Nz(CampoDaLinha(i, "id"), "0")) = idGravado Then achou = i: Exit For
            Next i
        End If

        linhaPlan = 0
        For k = 1 To UBound(idsPlan, 1)
            If IsNumeric(idsPlan(k, 1)) Then
                If CLng(idsPlan(k, 1)) = idGravado Then
                    linhaPlan = lo.DataBodyRange.Row + k - 1
                    Exit For
                End If
            End If
        Next k

        If achou = 0 Or linhaPlan = 0 Then
            remontar = True
            faltaram = faltaram + 1
        Else
            ReDim arr(1 To 1, 1 To nc + 1)
            txt = ""
            For j = 1 To nc
                esp = Split(CStr(cols(LBound(cols) + j - 1)), "|")
                arr(1, j) = CelulaDaColuna(achou, CStr(esp(0)), CStr(esp(3)))
                txt = txt & " " & CStr(arr(1, j))
            Next j
            txt = txt & " " & modDB.Nz(CampoDaLinha(achou, "observacoes")) & _
                        " " & modDB.Nz(CampoDaLinha(achou, "prox_acao")) & _
                        " " & modDB.Nz(CampoDaLinha(achou, "maquina")) & _
                        " " & modDB.Nz(CampoDaLinha(achou, "cnpj")) & _
                        " " & modDB.Nz(CampoDaLinha(achou, "email")) & _
                        " " & modDB.Nz(CampoDaLinha(achou, "categoria"))
            arr(1, nc + 1) = UCase$(txt)
            ' uma escrita por linha: nc celulas de uma vez
            ws.Range(ws.Cells(linhaPlan, COL_INI), _
                     ws.Cells(linhaPlan, COL_INI + nc)).Value2 = arr
            ' situacao (_a, depois de _id): desativado em outra estacao fica
            ' vermelho na hora; sai da lista no proximo filtro ou Atualizar
            ws.Cells(linhaPlan, COL_INI + nc + 2).Value2 = TextoAtivo(achou, tabela)
        End If
    Next chave
    On Error GoTo 0

    Proteger ws
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    If remontar And remontarSeFaltar Then
        Montar ws
    Else
        Status ws
    End If
    AtualizarRegistros = faltaram
    Exit Function
semGrade:
    If remontarSeFaltar Then Montar ws
    AtualizarRegistros = pendentes.Count
    Exit Function
falhaEscrita:
    Dim nErr As Long, sErr As String
    nErr = Err.Number: sErr = Err.Description
    Proteger ws
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Err.Raise nErr, "modGrade.AtualizarRegistros", sErr
End Function

'----------------------------------------------------------
' Aviso do feed (modFeed): registro novo ou excluido por outra
' estacao ou pela IA. Nao remonta a aba sozinho - remontar
' apaga filtro e ordenacao no meio do trabalho do usuario.
'----------------------------------------------------------
Public Sub AvisarPendentes(ByVal ws As Worksheet, ByVal quantos As Long)
    Dim c As Range
    Set c = ws.Cells(LIN_STATUS, COL_INI)
    Desproteger ws
    Application.EnableEvents = False
    c.Value2 = ChrW$(&H25B2) & " " & quantos & " registro(s) novo(s) ou exclu" & ChrW$(&HED) & "do(s) " & _
               "em outra esta" & ChrW$(&HE7) & ChrW$(&HE3) & "o ou pela IA " & ChrW$(&HB7) & _
               " clique em Atualizar (" & Format$(Now, "hh:nn") & ")   " & ChrW$(&HB7) & "   " & _
               CStr(modDB.Nz(c.Value2))
    c.Font.Color = modTema.COR_ALERTA
    Application.EnableEvents = True
    Proteger ws
End Sub

'----------------------------------------------------------
' CONJUNTO VISIVEL DA GRADE
'
' Devolve os ids das linhas que estao NA TELA, na ordem em
' que estao - o que sobrou do filtro, como foi ordenado.
' E esse conjunto que a ficha percorre no Anterior/Proximo:
' o filtro combinado da aba vale mais que qualquer combo que
' eu pudesse colocar na ficha, porque ele combina QUALQUER
' coluna sem precisar de codigo novo.
'----------------------------------------------------------
Private Function IDsVisiveis(ByVal ws As Worksheet) As Variant
    Dim lo As ListObject, idx As Long, area As Range, lin As Range
    Dim ids() As Long, n As Long, v As Variant, k As Long

    If ws.ListObjects.Count = 0 Then Exit Function
    Set lo = ws.ListObjects(1)
    If lo.DataBodyRange Is Nothing Then Exit Function
    idx = IndiceColuna(lo, CAB_ID)
    If idx = 0 Then Exit Function

    ReDim ids(1 To lo.DataBodyRange.Rows.Count)

    ' Cada bloco de linhas visiveis lido de uma vez (uma chamada ao
    ' Excel por bloco, nao uma por linha). A coluna _id e OCULTA:
    ' SpecialCells(visivel) nela nao devolveria nada - por isso as
    ' linhas visiveis vem da 1a coluna (codigo, sempre visivel; com a
    ' tabela inteira, uma coluna do meio oculta partiria os blocos e
    ' repetiria ids) e cruzam com a coluna de id.
    On Error Resume Next
    For Each area In lo.ListColumns(1).DataBodyRange.SpecialCells(12).Areas   ' xlCellTypeVisible
        v = Intersect(area.EntireRow, lo.ListColumns(idx).DataBodyRange).Value2
        If IsArray(v) Then
            For k = 1 To UBound(v, 1)
                If IsNumeric(v(k, 1)) Then
                    If CLng(v(k, 1)) > 0 Then n = n + 1: ids(n) = CLng(v(k, 1))
                End If
            Next k
        ElseIf IsNumeric(v) Then
            If CLng(v) > 0 Then n = n + 1: ids(n) = CLng(v)
        End If
    Next area
    On Error GoTo 0

    If n = 0 Then Exit Function
    ReDim Preserve ids(1 To n)
    IDsVisiveis = ids
End Function

'----------------------------------------------------------
' Digitou na grade: a aba esta protegida, entao isso so
' acontece se a protecao tiver sido removida a mao. Refaz a
' grade e avisa - celula editada sem validacao e sem versao
' foi a origem das perdas na planilha compartilhada.
'----------------------------------------------------------
Public Sub AvisarSomenteLeitura(ByVal ws As Object, ByVal alvo As Object)
    Dim lo As ListObject
    If Not Application.EnableEvents Then Exit Sub
    If TabelaDaAba(ws) = "" Then Exit Sub
    If ws.ListObjects.Count = 0 Then Exit Sub

    Set lo = ws.ListObjects(1)
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If Intersect(alvo, lo.DataBodyRange) Is Nothing Then Exit Sub

    MsgBox "Esta grade e somente leitura - nada foi gravado no banco." & vbCrLf & vbCrLf & _
           "Para alterar, de duplo clique na linha e edite pela ficha, " & _
           "que confere a versao do registro antes de gravar.", _
           vbInformation, "CRM Zapromaq"
    Montar ws
End Sub
