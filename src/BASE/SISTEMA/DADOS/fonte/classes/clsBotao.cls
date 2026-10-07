VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBotao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==========================================================
' clsBotao - botao moderno feito de tres Labels do MSForms.
'
'   bg_<nome>   fundo do botao (cor da variante, realce do mouse)
'   <nome>      o TEXTO, transparente e centralizado na vertical.
'               Tem o nome do botao (ex.: cmdSalvar): o clique nele
'               dispara o evento do proprio formulario
'               (cmdSalvar_Click), e o codigo continua usando
'               cmdSalvar.Caption e cmdSalvar.Enabled como antes.
'   ico_<nome>  icone (fonte Segoe MDL2 Assets), a esquerda
' Cliques no fundo e no icone sao repassados ao formulario por
' UI_Acao("<nome>"), que chama o mesmo evento.
'
' MSForms nao centraliza texto na vertical nem tem evento de
' "mouse saiu": por isso as tres camadas, e por isso o
' formulario chama clsUI.Soltar no proprio MouseMove.
' Cores: modTema.CorBotao(variante, parte), geradas do tema.
'==========================================================
Option Explicit

Private WithEvents mFundo As MSForms.Label
Attribute mFundo.VB_VarHelpID = -1
Private WithEvents mTexto As MSForms.Label
Attribute mTexto.VB_VarHelpID = -1
Private WithEvents mIcone As MSForms.Label
Attribute mIcone.VB_VarHelpID = -1
Private mForm As Object
Private mUI As clsUI
Private mVariante As String
Private mRealcado As Boolean
Private mEstado As Long                    ' -1 nunca pintado; 0 desabilitado; 1 habilitado
Private mRepassarTexto As Boolean          ' botao criado em execucao (clsUI.LigarBotao)

Public Sub Ligar(ByVal frm As Object, ByVal ui As clsUI, ByVal fundo As MSForms.Label, _
                 ByVal texto As MSForms.Label, ByVal icone As MSForms.Label, ByVal variante As String, _
                 Optional ByVal repassarTexto As Boolean = False)
    mRepassarTexto = repassarTexto
    Set mForm = frm
    Set mUI = ui
    Set mFundo = fundo
    Set mTexto = texto
    Set mIcone = icone
    mVariante = variante
    mEstado = -1
    Pintar
End Sub

Public Property Get Nome() As String
    Nome = mTexto.Name
End Property

'----------------------------------------------------------
' Repinta conforme o Enabled do Label de TEXTO (o que o codigo
' do formulario liga e desliga). clsUI.Pintar chama isto em
' todos os botoes depois de cada troca de modo.
'----------------------------------------------------------
Public Sub Pintar()
    Dim hab As Long, fundo As Long, frente As Long, borda As Long
    hab = IIf(mTexto.Enabled, 1, 0)
    If hab = mEstado And Not mRealcado Then Exit Sub
    mEstado = hab
    mRealcado = False
    If hab = 1 Then
        fundo = modTema.CorBotao(mVariante, "fundo")
        frente = modTema.CorBotao(mVariante, "texto")
    Else
        fundo = modTema.COR_DESABILITADO
        frente = modTema.COR_DESABILITADO_TEXTO
    End If
    If Not mFundo Is Nothing Then
        mFundo.BackColor = fundo
        borda = modTema.CorBotao(mVariante, "borda")
        If borda >= 0 And hab = 1 Then
            mFundo.BorderStyle = 1: mFundo.BorderColor = borda
        Else
            mFundo.BorderStyle = 0
        End If
    End If
    mTexto.ForeColor = frente
    If Not mIcone Is Nothing Then
        mIcone.ForeColor = frente
        mIcone.Enabled = (hab = 1)
    End If
End Sub

Public Sub Realcar(ByVal sim As Boolean)
    If sim = mRealcado Then Exit Sub
    If Not mTexto.Enabled Then Exit Sub
    If mFundo Is Nothing Then Exit Sub
    mRealcado = sim
    If sim Then
        mFundo.BackColor = modTema.CorBotao(mVariante, "hover")
    Else
        mFundo.BackColor = modTema.CorBotao(mVariante, "fundo")
    End If
End Sub

Private Sub Acionar()
    If Not mTexto.Enabled Then Exit Sub
    CallByName mForm, "UI_Acao", VbMethod, mTexto.Name
End Sub

Private Sub mFundo_Click()
    Acionar
End Sub

Private Sub mIcone_Click()
    Acionar
End Sub

' So no botao criado em execucao: no do designer o clique no texto
' ja dispara o <nome>_Click do formulario, e repassar chamaria duas vezes.
Private Sub mTexto_Click()
    If mRepassarTexto Then Acionar
End Sub

Private Sub mFundo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mUI.Realcar Me
End Sub

Private Sub mTexto_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mUI.Realcar Me
End Sub

Private Sub mIcone_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mUI.Realcar Me
End Sub
