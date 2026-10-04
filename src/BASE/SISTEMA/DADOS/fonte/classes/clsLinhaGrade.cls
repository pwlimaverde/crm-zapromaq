VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLinhaGrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==========================================================
' clsLinhaGrade - Label transparente por cima de uma linha
' visivel da grade: recebe clique e duplo clique da linha
' inteira e repassa a clsGrade com o numero da linha na tela.
'==========================================================
Option Explicit

Private WithEvents mCapa As MSForms.Label
Attribute mCapa.VB_VarHelpID = -1
Private mGrade As clsGrade
Private mPosicao As Long

Public Sub Ligar(ByVal grade As clsGrade, ByVal capa As MSForms.Label, ByVal posicao As Long)
    Set mGrade = grade
    Set mCapa = capa
    mPosicao = posicao
End Sub

Private Sub mCapa_Click()
    mGrade.CliqueNaLinha mPosicao
End Sub

Private Sub mCapa_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    mGrade.DuploCliqueNaLinha mPosicao
End Sub
