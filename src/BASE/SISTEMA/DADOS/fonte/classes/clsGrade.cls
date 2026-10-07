VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==========================================================
' clsGrade - grade de leitura desenhada com Labels num Frame.
'
' Substitui a ListBox, que nao colore linha nem celula: aqui ha
' cabecalho na cor da marca, linhas alternadas, linha
' selecionada destacada e a coluna de situacao com ponto e
' texto coloridos (a cor nunca e o unico sinal: o texto vai
' junto). So as linhas VISIVEIS existem como Label; rolar so
' troca o texto delas - 500 registros nao viram 3.000 Labels.
'
' Interface parecida com a da ListBox (Clear, AddItem, List,
' ListCount, ListIndex, ColumnCount, ColumnWidths), para o
' codigo da ficha mudar pouco. Mudar ListIndex avisa o
' formulario por UI_GradeClique(nome), como o Click da ListBox;
' duplo clique chama UI_GradeDuploClique(nome).
'
' Teclado (como a ListBox): setas, PgUp/PgDn, Home/End e Enter
' (= duplo clique). Quem recebe as teclas e uma caixa de texto
' invisivel dentro do Frame, que ganha o foco no clique da linha.
' Limite do MSForms: sem roda do mouse (so com subclassing,
' que derruba o Excel). Rola pela barra, pelos cliques e teclas.
'==========================================================
Option Explicit

Private Const ALT_CAB As Single = 18
Private Const ALT_LIN As Single = 16
Private Const LARG_BARRA As Single = 12
Private Const MARGEM As Single = 3

Private mForm As Object
Private mFrame As MSForms.Frame
Private WithEvents mBarra As MSForms.ScrollBar
Attribute mBarra.VB_VarHelpID = -1
Private WithEvents mTecla As MSForms.TextBox
Attribute mTecla.VB_VarHelpID = -1
Private mCab() As MSForms.Label
Private mCel() As MSForms.Label            ' (linha visivel, coluna): texto, transparente
Private mFundoLinha() As MSForms.Label     ' fundo de cada linha visivel (zebra/selecao)
Private mCapas As Collection
Private mDados As Collection               ' cada item: Array de textos
Private mLarguras() As Single
Private mDireita() As Boolean
Private mColunas As Long
Private mVisiveis As Long
Private mTopo As Long                      ' indice (0) do primeiro registro na tela
Private mSel As Long                       ' indice (0) selecionado; -1 = nenhum
Private mColSelo As Long                   ' coluna de situacao (0 = nenhuma; 1..n)
Private mMontada As Boolean

'==========================================================
Public Sub Ligar(ByVal frm As Object, ByVal quadro As MSForms.Frame)
    Set mForm = frm
    Set mFrame = quadro
    Set mDados = New Collection
    mSel = -1
    mFrame.Caption = ""
    mFrame.BorderStyle = 1                     ' fmBorderStyleSingle
    mFrame.SpecialEffect = 0                   ' flat
    mFrame.BorderColor = modTema.COR_BORDA_SUAVE
    mFrame.BackColor = modTema.COR_BRANCO
    mFrame.ScrollBars = 0
    Set mBarra = mFrame.Controls.Add("Forms.ScrollBar.1", "sb_" & mFrame.Name, True)
    mBarra.Orientation = 0                     ' vertical
    mBarra.Min = 0: mBarra.Max = 0: mBarra.SmallChange = 1
    ' receptor de teclado: 1x1 pt, sem borda nem fundo, travado
    ' (KeyDown funciona com Locked; nada digitado aparece)
    Set mTecla = mFrame.Controls.Add("Forms.TextBox.1", "tk_" & mFrame.Name, True)
    mTecla.Left = 0: mTecla.Top = 0: mTecla.Width = 1: mTecla.Height = 1
    mTecla.BorderStyle = 0: mTecla.SpecialEffect = 0: mTecla.BackStyle = 0
    mTecla.Locked = True
End Sub

' Leva o foco para a grade (setas passam a mover a selecao).
Public Sub Focar()
    On Error Resume Next
    mTecla.SetFocus
    On Error GoTo 0
End Sub

Public Property Get Nome() As String
    Nome = mFrame.Name
End Property

'----------------------------------------------------------
' Colunas: larguras em pt separadas por ";" (como ColumnWidths
' da ListBox). Recria os Labels - chamar antes de carregar.
'----------------------------------------------------------
Public Property Let ColumnCount(ByVal n As Long)
    mColunas = n
End Property

Public Property Get ColumnCount() As Long
    ColumnCount = mColunas
End Property

Public Property Let ColumnWidths(ByVal larguras As String)
    Dim p As Variant, i As Long
    p = Split(Replace(larguras, "pt", ""), ";")
    mColunas = UBound(p) - LBound(p) + 1
    ReDim mLarguras(1 To mColunas)
    ReDim mDireita(1 To mColunas)
    For i = 1 To mColunas
        mLarguras(i) = CSng(Val(Trim$(p(LBound(p) + i - 1))))
    Next i
    Montar
End Property

Public Sub AlinharDireita(ByVal coluna As Long)
    If coluna >= 1 And coluna <= mColunas Then mDireita(coluna) = True
    If mMontada Then Redesenhar
End Sub

' A cor do selo vem do estado ANTES de " - ": "SUSPEITO - FORA" e
' "SUSPEITO - ENTRA" (lote) pintam como SUSPEITO, sem uma cor por rotulo.
Public Function ChaveSelo(ByVal texto As String) As String
    Dim p As Long
    p = InStr(texto, " - ")
    If p > 0 Then ChaveSelo = Left$(texto, p - 1) Else ChaveSelo = texto
End Function

Public Property Let ColunaSelo(ByVal coluna As Long)
    mColSelo = coluna
End Property

Public Sub Cabecalho(ByVal titulos As Variant)
    Dim i As Long
    If Not mMontada Then Exit Sub
    For i = 1 To mColunas
        mCab(i).Caption = CStr(titulos(LBound(titulos) + i - 1))
    Next i
End Sub

'----------------------------------------------------------
Private Sub Montar()
    Dim i As Long, j As Long, x As Single, lb As MSForms.Label, capa As clsLinhaGrade
    Dim larguraUtil As Single

    ' desmonta o que houver (troca de colunas)
    For i = mFrame.Controls.Count - 1 To 0 Step -1
        If mFrame.Controls(i).Name <> mBarra.Name And mFrame.Controls(i).Name <> mTecla.Name Then mFrame.Controls.Remove i
    Next i
    larguraUtil = mFrame.InsideWidth - LARG_BARRA
    mVisiveis = Int((mFrame.InsideHeight - ALT_CAB) / ALT_LIN)
    If mVisiveis < 1 Then mVisiveis = 1
    mBarra.Left = larguraUtil: mBarra.Top = 0
    mBarra.Width = LARG_BARRA: mBarra.Height = mFrame.InsideHeight

    ReDim mCab(1 To mColunas)
    ReDim mCel(1 To mVisiveis, 1 To mColunas)
    ReDim mFundoLinha(1 To mVisiveis)
    Set mCapas = New Collection
    ' Label do MSForms nao centraliza na vertical: fundo em um Label,
    ' texto em outro, transparente e deslocado alguns pontos para baixo.
    Set lb = mFrame.Controls.Add("Forms.Label.1", "gcf", True)
    Posicionar lb, 0, 0, larguraUtil, ALT_CAB
    lb.BackColor = modTema.COR_AZUL
    x = 0
    For j = 1 To mColunas
        Set lb = mFrame.Controls.Add("Forms.Label.1", "gc_" & j, True)
        Posicionar lb, x + MARGEM, 3, IIf(j = mColunas, larguraUtil - x, mLarguras(j)) - MARGEM, ALT_CAB - 3
        lb.BackStyle = 0
        lb.ForeColor = modTema.COR_BRANCO
        lb.Font.Name = modTema.FONTE_PADRAO: lb.Font.Size = 8.5: lb.Font.Bold = True
        Set mCab(j) = lb
        x = x + mLarguras(j)
    Next j
    For i = 1 To mVisiveis
        Set lb = mFrame.Controls.Add("Forms.Label.1", "gf_" & i, True)
        Posicionar lb, 0, ALT_CAB + (i - 1) * ALT_LIN, larguraUtil, ALT_LIN
        Set mFundoLinha(i) = lb
        x = 0
        For j = 1 To mColunas
            Set lb = mFrame.Controls.Add("Forms.Label.1", "g" & i & "_" & j, True)
            Posicionar lb, x + MARGEM, ALT_CAB + (i - 1) * ALT_LIN + 2, _
                       IIf(j = mColunas, larguraUtil - x, mLarguras(j)) - MARGEM * 2, ALT_LIN - 2
            lb.BackStyle = 0
            lb.Font.Name = modTema.FONTE_PADRAO: lb.Font.Size = 8.5
            lb.WordWrap = False
            Set mCel(i, j) = lb
            x = x + mLarguras(j)
        Next j
        ' capa transparente por cima da linha: um clique vale a linha inteira
        Set lb = mFrame.Controls.Add("Forms.Label.1", "gk_" & i, True)
        Posicionar lb, 0, ALT_CAB + (i - 1) * ALT_LIN, larguraUtil, ALT_LIN
        lb.BackStyle = 0
        Set capa = New clsLinhaGrade
        capa.Ligar Me, lb, i
        mCapas.Add capa
    Next i
    mMontada = True
    Redesenhar
End Sub

Private Sub Posicionar(ByVal lb As MSForms.Label, ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single)
    lb.Left = x: lb.Top = y
    If w < 1 Then w = 1
    lb.Width = w: lb.Height = h
    lb.BackStyle = 1
    lb.BackColor = modTema.COR_BRANCO
    lb.BorderStyle = 0
    lb.SpecialEffect = 0
    lb.TextAlign = 1
End Sub

'==========================================================
' DADOS
'==========================================================
Public Sub Clear()
    Set mDados = New Collection
    mTopo = 0: mSel = -1
    AjustarBarra
    If mMontada Then Redesenhar
End Sub

Public Sub AddItem(Optional ByVal primeiro As String = "")
    Dim linha() As String
    ReDim linha(0 To IIf(mColunas > 0, mColunas - 1, 0))
    linha(0) = primeiro
    mDados.Add linha
    AjustarBarra
    If mDados.Count - 1 - mTopo < mVisiveis Then DesenharLinha mDados.Count - 1 - mTopo + 1
End Sub

' Carga de uma vez (matriz base 1: linha, coluna) - uma redesenhada so.
Public Sub Carregar(ByVal matriz As Variant)
    Dim i As Long, j As Long, linha() As String
    Set mDados = New Collection
    mTopo = 0: mSel = -1
    If IsArray(matriz) Then
        For i = LBound(matriz, 1) To UBound(matriz, 1)
            ReDim linha(0 To mColunas - 1)
            For j = 1 To mColunas
                If j - 1 + LBound(matriz, 2) <= UBound(matriz, 2) Then linha(j - 1) = CStr(matriz(i, j - 1 + LBound(matriz, 2)))
            Next j
            mDados.Add linha
        Next i
    End If
    AjustarBarra
    If mMontada Then Redesenhar
End Sub

Public Property Get List(ByVal linha As Long, ByVal coluna As Long) As String
    List = mDados(linha + 1)(coluna)
End Property

Public Property Let List(ByVal linha As Long, ByVal coluna As Long, ByVal valor As String)
    Dim a() As String
    a = mDados(linha + 1)
    a(coluna) = valor
    mDados.Remove linha + 1
    If linha + 1 > mDados.Count Then mDados.Add a Else mDados.Add a, , linha + 1
    If linha >= mTopo And linha < mTopo + mVisiveis Then DesenharLinha linha - mTopo + 1
End Property

Public Property Get ListCount() As Long
    ListCount = mDados.Count
End Property

Public Property Get ListIndex() As Long
    ListIndex = mSel
End Property

' Como na ListBox: mudar a selecao avisa o formulario (UI_GradeClique).
Public Property Let ListIndex(ByVal indice As Long)
    Dim anterior As Long
    If indice < -1 Or indice >= mDados.Count Then Exit Property
    anterior = mSel
    mSel = indice
    If mSel >= 0 Then
        If mSel < mTopo Then mTopo = mSel
        If mSel >= mTopo + mVisiveis Then mTopo = mSel - mVisiveis + 1
        If mBarra.Value <> mTopo Then mBarra.Value = mTopo Else Redesenhar
    Else
        Redesenhar
    End If
    ' como a ListBox: so avisa quando a selecao MUDA
    If anterior <> mSel Then CallByName mForm, "UI_GradeClique", VbMethod, mFrame.Name
End Property

'==========================================================
' DESENHO
'==========================================================
Private Sub AjustarBarra()
    Dim m As Long
    m = mDados.Count - mVisiveis
    If m < 0 Then m = 0
    mBarra.Max = m
    mBarra.LargeChange = IIf(mVisiveis > 1, mVisiveis - 1, 1)
    mBarra.Enabled = (m > 0)
    If mTopo > m Then mTopo = m
End Sub

Private Sub Redesenhar()
    Dim i As Long
    If Not mMontada Then Exit Sub
    For i = 1 To mVisiveis
        DesenharLinha i
    Next i
End Sub

Private Sub DesenharLinha(ByVal pos As Long)
    Dim idx As Long, j As Long, fundo As Long, frente As Long, t As String, lb As MSForms.Label
    If Not mMontada Then Exit Sub
    If pos < 1 Or pos > mVisiveis Then Exit Sub
    idx = mTopo + pos - 1
    If idx = mSel Then
        fundo = modTema.COR_SELECAO: frente = modTema.COR_SELECAO_TEXTO
    ElseIf idx Mod 2 = 1 Then
        fundo = modTema.COR_ZEBRA: frente = modTema.COR_TEXTO
    Else
        fundo = modTema.COR_BRANCO: frente = modTema.COR_TEXTO
    End If
    mFundoLinha(pos).BackColor = fundo
    For j = 1 To mColunas
        Set lb = mCel(pos, j)
        If idx < mDados.Count Then t = mDados(idx + 1)(j - 1) Else t = ""
        lb.ForeColor = frente
        lb.Font.Bold = False
        lb.TextAlign = IIf(mDireita(j), 3, 1)
        If j = mColSelo And t <> "" Then
            ' ponto + texto na cor da situacao (Segoe UI tem o U+25CF)
            lb.Caption = ChrW$(&H25CF) & " " & t
            lb.ForeColor = modTema.CorSituacao(ChaveSelo(t), "texto")
            lb.Font.Bold = True
        Else
            lb.Caption = t
        End If
    Next j
End Sub

'==========================================================
' EVENTOS
'==========================================================
Private Sub mBarra_Change()
    mTopo = mBarra.Value
    Redesenhar
End Sub

Private Sub mBarra_Scroll()
    mTopo = mBarra.Value
    Redesenhar
End Sub

Public Sub CliqueNaLinha(ByVal pos As Long)
    Dim idx As Long
    idx = mTopo + pos - 1
    If idx < mDados.Count Then ListIndex = idx
    Focar
End Sub

Private Sub mTecla_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim alvo As Long
    If mDados.Count = 0 Then Exit Sub
    alvo = mSel
    Select Case KeyCode
        Case 38: alvo = mSel - 1                          ' seta acima
        Case 40: alvo = mSel + 1                          ' seta abaixo
        Case 33: alvo = mSel - (mVisiveis - 1)            ' PgUp
        Case 34: alvo = mSel + (mVisiveis - 1)            ' PgDn
        Case 36: alvo = 0                                 ' Home
        Case 35: alvo = mDados.Count - 1                  ' End
        Case 13                                           ' Enter = duplo clique
            KeyCode = 0
            If mSel >= 0 Then CallByName mForm, "UI_GradeDuploClique", VbMethod, mFrame.Name
            Exit Sub
        Case Else: Exit Sub
    End Select
    KeyCode = 0
    If alvo < 0 Then alvo = 0
    If alvo > mDados.Count - 1 Then alvo = mDados.Count - 1
    If alvo <> mSel Then ListIndex = alvo
End Sub

Public Sub DuploCliqueNaLinha(ByVal pos As Long)
    Dim idx As Long
    idx = mTopo + pos - 1
    If idx >= mDados.Count Then Exit Sub
    If idx <> mSel Then ListIndex = idx
    CallByName mForm, "UI_GradeDuploClique", VbMethod, mFrame.Name
End Sub
