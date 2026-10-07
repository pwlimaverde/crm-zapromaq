VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==========================================================
' clsUI - liga os componentes visuais de um formulario.
'
' O build cria os controles e marca cada um no Tag:
'   "ui:botao:<variante>"        Label de TEXTO do botao (tem o nome
'                                do botao); acompanham bg_<nome> e ico_<nome>
'   "ui:grade"                   Frame que vira clsGrade
'   "ui:glifo:<hex>"             Label de icone: Caption regravado ao abrir
' No UserForm_Initialize:   mUI.Preparar Me
' No UserForm_MouseMove:    mUI.Soltar
' Depois de habilitar/desabilitar botoes: mUI.Pintar
'==========================================================
Option Explicit

Private mForm As Object
Private mBotoes As Collection
Private mGrades As Collection
Private mRealcado As clsBotao

Public Sub Preparar(ByVal frm As Object)
    Dim c As Object, p As Variant, b As clsBotao, icone As Object, fundo As Object, g As clsGrade
    Set mForm = frm
    Set mBotoes = New Collection
    Set mGrades = New Collection
    For Each c In frm.Controls
        If Left$(CStr(c.Tag), 3) = "ui:" Then
            p = Split(CStr(c.Tag), ":")
            Select Case p(1)
                Case "botao"
                    Set icone = Nothing: Set fundo = Nothing
                    On Error Resume Next
                    Set icone = frm.Controls("ico_" & c.Name)
                    Set fundo = frm.Controls("bg_" & c.Name)
                    On Error GoTo 0
                    Set b = New clsBotao
                    b.Ligar frm, Me, fundo, c, icone, CStr(p(2))
                    mBotoes.Add b, c.Name
                Case "grade"
                    Set g = New clsGrade
                    g.Ligar frm, c
                    mGrades.Add g, c.Name
                Case "glifo"
                    ' ICONE (Segoe MDL2). Duas coisas, nesta ordem:
                    '  1. Charset = 2 (simbolo). Com o charset padrao o texto e
                    '     tratado como ANSI/DBCS e o codigo cujo segundo byte
                    '     forma par valido sai como OUTRO caractere: E76C
                    '     aparecia como cifrao de euro; E821 e E77B, como
                    '     ideograma. Conferido no Excel 2019.
                    '  2. Caption regravado do codigo guardado na Tag.
                    c.Font.Charset = 2
                    c.Caption = ChrW$(CLng("&H" & p(2)))
            End Select
        End If
    Next c
End Sub

'----------------------------------------------------------
' Botao criado em TEMPO DE EXECUCAO (ex.: o botao de cada campo de
' vinculo da ficha, criado com os campos). Controle criado por
' Controls.Add nao dispara o <nome>_Click do formulario, entao o
' clique no TEXTO tambem e repassado por UI_Acao("<nome>").
' Religar o mesmo nome substitui o botao anterior.
'----------------------------------------------------------
Public Sub LigarBotao(ByVal fundo As Object, ByVal texto As Object, ByVal icone As Object, _
                      ByVal variante As String)
    Dim b As clsBotao
    On Error Resume Next
    mBotoes.Remove texto.Name
    On Error GoTo 0
    Set b = New clsBotao
    b.Ligar mForm, Me, fundo, texto, icone, variante, True
    mBotoes.Add b, texto.Name
End Sub

Public Function Grade(ByVal nome As String) As clsGrade
    Set Grade = mGrades(nome)
End Function

Public Sub Pintar()
    Dim b As clsBotao
    If mBotoes Is Nothing Then Exit Sub
    For Each b In mBotoes
        b.Pintar
    Next b
End Sub

Public Sub Realcar(ByVal b As clsBotao)
    If Not mRealcado Is Nothing Then
        If Not mRealcado Is b Then mRealcado.Realcar False
    End If
    b.Realcar True
    Set mRealcado = b
End Sub

Public Sub Soltar()
    If mRealcado Is Nothing Then Exit Sub
    mRealcado.Realcar False
    Set mRealcado = Nothing
End Sub
