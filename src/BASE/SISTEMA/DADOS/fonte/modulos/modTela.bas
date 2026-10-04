Attribute VB_Name = "modTela"
'==========================================================
' modTela - encaixe do formulario na tela da estacao.
'
' As telas sao desenhadas para 100% (referencia 1366x768). Em
' escala 125%/150% do Windows, ou tela menor, o formulario nao
' caberia. Zoom do UserForm reduz controles e fontes juntos
' (referencias/01-vba-msforms/zoom-property.md); o tamanho da
' janela e ajustado a mao, porque Zoom nao muda Width/Height.
'
' A conta usa a AREA UTIL DA TELA EM PONTOS (pixels x 72 / dpi).
' Ela da o mesmo resultado com o Excel ciente ou nao de DPI:
' sem suporte, o Windows entrega pixels e dpi "logicos"; com
' suporte, fisicos - a razao e a mesma. So a DPI enganaria
' (processo sem suporte sempre ve 96). Defeito A15 do plano.
'==========================================================
Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare PtrSafe Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" _
    (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As RECT, ByVal fuWinIni As Long) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long

Private Const SPI_GETWORKAREA As Long = 48
Private Const LOGPIXELSX As Long = 88
Private Const FOLGA As Single = 12          ' pt de respiro em volta da janela

Public Function DPI() As Long
    Dim dc As LongPtr
    dc = GetDC(0)
    DPI = GetDeviceCaps(dc, LOGPIXELSX)
    ReleaseDC 0, dc
    If DPI <= 0 Then DPI = 96
End Function

' Area de trabalho (sem a barra de tarefas) em pontos.
Public Sub AreaUtil(ByRef largura As Single, ByRef altura As Single)
    Dim r As RECT, d As Long
    d = DPI()
    If SystemParametersInfo(SPI_GETWORKAREA, 0, r, 0) = 0 Then
        largura = 1366 * 72 / 96: altura = 728 * 72 / 96
        Exit Sub
    End If
    largura = (r.Right - r.Left) * 72 / d
    altura = (r.Bottom - r.Top) * 72 / d
End Sub

' Zoom (10 a 100) para uma janela de largura x altura pt caber na tela.
Public Function ZoomIdeal(ByVal largura As Single, ByVal altura As Single) As Long
    Dim w As Single, h As Single, z As Double
    AreaUtil w, h
    z = 100
    If largura + FOLGA > w Then z = (w - FOLGA) / largura * 100
    If altura + FOLGA > h Then If (h - FOLGA) / altura * 100 < z Then z = (h - FOLGA) / altura * 100
    If z < 10 Then z = 10
    ZoomIdeal = Int(z)
End Function

'----------------------------------------------------------
' Aplica o zoom e redimensiona a janela na mesma proporcao.
' Chamar no UserForm_Initialize, depois de montar os controles.
'----------------------------------------------------------
Public Sub Encaixar(ByVal frm As Object)
    Dim larg As Single, alt As Single, z As Long
    larg = frm.Width: alt = frm.Height
    z = ZoomIdeal(larg, alt)
    If z >= 100 Then Exit Sub
    frm.Zoom = z
    frm.Width = larg * z / 100
    frm.Height = alt * z / 100
End Sub

' Texto para diagnostico (tela de calibracao e Sobre).
Public Function Diagnostico(ByVal frm As Object) As String
    Dim w As Single, h As Single
    AreaUtil w, h
    Diagnostico = "Area interna " & Format$(frm.InsideWidth, "0.0") & " x " & Format$(frm.InsideHeight, "0.0") & _
                  " pt | janela " & Format$(frm.Width, "0") & " x " & Format$(frm.Height, "0") & _
                  " pt | " & DPI() & " dpi | area util " & Format$(w, "0") & " x " & Format$(h, "0") & _
                  " pt | zoom " & frm.Zoom
End Function
