Attribute VB_Name = "modDiag"
Option Explicit
'==========================================================
' modDiag - diagnostico tecnico do front (DEV-06 parte A, ciclo 007)
'
' Antes o erro so aparecia na tela e se perdia. Agora cada falha de
' gravacao (e cada recusa por manutencao, esquema ou backup) vira uma
' linha em <BASE>\SISTEMA\DADOS\execucao\logs\front-<ESTACAO>-AAAAMM.log:
'   instante fuso  [ESTACAO\usuario vVERSAO]  tela|operacao|registro|resultado  erro N: texto
' (ciclo 018, DEV-06A: o instante sai com o fuso da estacao, ex. -03:00, como no conector)
' Sem conteudo comercial: so nomes de tela, tabela, operacao, id/codigo e o
' erro do motor. Um arquivo por estacao e por mes; a propria estacao apaga
' os seus anteriores ao mes anterior (retencao da secao 5.3).
' Falha ao gravar o diagnostico nunca bloqueia a operacao do usuario.
'==========================================================
Private mRetencaoFeita As Boolean
Public gDiagFalhas As Long

Private Type SYSTEMTIME_DIAG
    Partes(0 To 7) As Integer
End Type
Private Type TIME_ZONE_INFORMATION
    Bias As Long
    StandardName(0 To 63) As Byte
    StandardDate As SYSTEMTIME_DIAG
    StandardBias As Long
    DaylightName(0 To 63) As Byte
    DaylightDate As SYSTEMTIME_DIAG
    DaylightBias As Long
End Type
Private Declare PtrSafe Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

' Diferenca do relogio da estacao para o UTC, em minutos (Brasilia = -180).
Public Function FusoMinutos() As Long
    Dim tz As TIME_ZONE_INFORMATION, r As Long, bias As Long
    On Error GoTo falhou
    r = GetTimeZoneInformation(tz)
    bias = tz.Bias
    If r = 2 Then bias = bias + tz.DaylightBias Else If r = 1 Then bias = bias + tz.StandardBias
    FusoMinutos = -bias
    Exit Function
falhou:
    FusoMinutos = 0
End Function

' Fuso no formato ISO 8601: -180 -> "-03:00"; 330 -> "+05:30". Pura.
Public Function TextoFuso(ByVal minutos As Long) As String
    Dim s As String, m As Long
    If minutos < 0 Then s = "-" Else s = "+"
    m = Abs(minutos)
    TextoFuso = s & Format$(m \ 60, "00") & ":" & Format$(m Mod 60, "00")
End Function

Public Function Estacao() As String
    Dim e As String
    On Error Resume Next
    e = Environ$("COMPUTERNAME")
    On Error GoTo 0
    If e = "" Then e = "ESTACAO"
    Estacao = SoNome(e)
End Function

' troca o que nao e letra, digito, _ ou - por _ (mesmo criterio do conector)
Private Function SoNome(ByVal s As String) As String
    Dim i As Long, c As String, r As String
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If c Like "[A-Za-z0-9_-]" Then r = r & c Else r = r & "_"
    Next i
    SoNome = r
End Function

Public Function NomeArquivoDiag(ByVal est As String, ByVal quando As Date) As String
    NomeArquivoDiag = "front-" & est & "-" & Format$(quando, "yyyymm") & ".log"
End Function

' O arquivo e desta estacao e de antes do mes anterior? (retencao: mes corrente e o anterior)
Public Function DeveApagarDiag(ByVal nome As String, ByVal est As String, ByVal agora As Date) As Boolean
    Dim prefixo As String, ym As String, corte As Long
    prefixo = "front-" & est & "-"
    If LCase$(Left$(nome, Len(prefixo))) <> LCase$(prefixo) Then Exit Function
    If LCase$(Right$(nome, 4)) <> ".log" Then Exit Function
    ym = Mid$(nome, Len(prefixo) + 1, Len(nome) - Len(prefixo) - 4)
    If Len(ym) <> 6 Or Not IsNumeric(ym) Then Exit Function
    corte = CLng(Format$(DateAdd("m", -1, DateSerial(Year(agora), Month(agora), 1)), "yyyymm"))
    DeveApagarDiag = (CLng(ym) < corte)
End Function

' Uma linha sem quebra e sem excesso (o texto do erro do ADO pode ter varias linhas)
Public Function LinhaDiag(ByVal quando As Date, ByVal fusoMin As Long, ByVal est As String, ByVal usuario As String, _
                          ByVal versao As String, ByVal tela As String, ByVal operacao As String, ByVal registro As String, _
                          ByVal numero As Long, ByVal erro As String, ByVal resultado As String) As String
    Dim t As String
    t = Replace(Replace(erro, vbCr, " "), vbLf, " ")
    If Len(t) > 400 Then t = Left$(t, 400) & "..."
    LinhaDiag = Format$(quando, "yyyy-mm-dd hh:nn:ss") & " " & TextoFuso(fusoMin) & "  [" & est & "\" & usuario & " v" & versao & "]  " & _
                tela & "|" & operacao & "|" & registro & "|" & resultado & "  erro " & numero & ": " & t
End Function

Public Function PastaLogs() As String
    Dim c As String, k As Long
    c = modConfig.CaminhoBanco()
    k = InStrRev(c, "\")
    If k = 0 Then Exit Function
    PastaLogs = Left$(c, k) & "SISTEMA\DADOS\execucao\logs\"
End Function

Public Sub Registrar(ByVal tela As String, ByVal operacao As String, ByVal registro As String, _
                     ByVal numero As Long, ByVal erro As String, Optional ByVal resultado As String = "erro")
    Dim pasta As String, f As Integer, est As String
    On Error GoTo falhou
    pasta = PastaLogs()
    If pasta = "" Then Exit Sub
    If Dir(pasta, vbDirectory) = "" Then MkDir Left$(pasta, Len(pasta) - 1)
    est = Estacao()
    If Not mRetencaoFeita Then ApagarAntigos pasta, est
    f = FreeFile
    Open pasta & NomeArquivoDiag(est, Now) For Append Access Write Lock Write As #f
    Print #f, LinhaDiag(Now, FusoMinutos(), est, modConfig.UsuarioAtual(), modConfig.VERSAO_FRONT, tela, operacao, registro, numero, erro, resultado)
    Close #f
    Exit Sub
falhou:
    gDiagFalhas = gDiagFalhas + 1
    On Error Resume Next
    If f <> 0 Then Close #f
End Sub

Private Sub ApagarAntigos(ByVal pasta As String, ByVal est As String)
    Dim nome As String, nomes As String, lista As Variant, i As Long
    mRetencaoFeita = True
    On Error Resume Next
    nome = Dir(pasta & "front-" & est & "-*.log")
    Do While nome <> ""
        If DeveApagarDiag(nome, est, Now) Then nomes = nomes & nome & "|"
        nome = Dir()
    Loop
    If nomes = "" Then Exit Sub
    lista = Split(Left$(nomes, Len(nomes) - 1), "|")
    For i = LBound(lista) To UBound(lista)
        Kill pasta & lista(i)
    Next i
End Sub
