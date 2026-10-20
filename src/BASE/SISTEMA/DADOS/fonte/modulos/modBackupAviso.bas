Attribute VB_Name = "modBackupAviso"
Option Explicit
'==========================================================
' modBackupAviso - aviso de backup (spec 006; regra por janela no ciclo 007)
'
' O backup (backup-banco.ps1) grava em <BASE>\SISTEMA\DADOS\execucao\backup
'   crm_<diario|periodico>_AAAAMMDD-HHMMSS_<ESTACAO>.accdb   (+ .validado)
' diario = copia completa (com exclusividade); periodico = simples.
'
' Regra (secao 5.3 do plano de 19/10/2026), a mesma de lib\Backup.ps1 e do
' crm_status (Get-AvisoBackupJanela) - os tres testes usam os casos J1 a J8:
'   janelas 9h, 12h30 e 16h, de segunda a sexta; a janela encerra 30 min
'   depois do inicio; uma copia conta para a janela se o instante estiver em
'   [inicio, inicio + 60 min);
'   (a) a ultima janela encerrada nao tem copia nenhuma;
'   (b) o ultimo dia util cujas 16h30 ja passaram nao tem copia COMPLETA
'       VALIDADA (a simples nao apaga este aviso).
' Se a pasta nao puder ser lida (rede fora do ar, pasta inexistente) nao diz
' nada: o aviso de banco inacessivel ja cobre esse caso.
'==========================================================
Private Const ENCERRA_MIN As Long = 30
Private Const CONTA_MIN As Long = 60

' "crm_diario_20261018-081530_EST1.accdb" -> 18/10/2026 08:15:30; aceita o
' formato antigo HHMM. 0 = nome fora do padrao.
Public Function InstanteDoNome(ByVal nome As String) As Date
    Dim p1 As Long, p2 As Long, s As String, hora As String
    On Error GoTo fora
    p1 = InStr(nome, "_")
    If p1 = 0 Then Exit Function
    p2 = InStr(p1 + 1, nome, "_")
    If p2 = 0 Then Exit Function
    s = Mid$(nome, p2 + 1)
    If Mid$(s, 9, 1) <> "-" Then Exit Function
    If Not IsNumeric(Left$(s, 8)) Then Exit Function
    If IsNumeric(Mid$(s, 10, 6)) And InStr(Mid$(s, 10, 6), "_") = 0 Then
        hora = Mid$(s, 10, 6)
    ElseIf IsNumeric(Mid$(s, 10, 4)) Then
        hora = Mid$(s, 10, 4) & "00"
    Else
        Exit Function
    End If
    InstanteDoNome = DateSerial(CInt(Left$(s, 4)), CInt(Mid$(s, 5, 2)), CInt(Mid$(s, 7, 2))) + _
                     TimeSerial(CInt(Left$(hora, 2)), CInt(Mid$(hora, 3, 2)), CInt(Mid$(hora, 5, 2)))
    Exit Function
fora:
    InstanteDoNome = 0
End Function

' "crm_diario_..." -> "diario"
Public Function TipoDoNome(ByVal nome As String) As String
    Dim p1 As Long, p2 As Long
    p1 = InStr(nome, "_")
    If p1 = 0 Then Exit Function
    p2 = InStr(p1 + 1, nome, "_")
    If p2 = 0 Then Exit Function
    TipoDoNome = Mid$(nome, p1 + 1, p2 - p1 - 1)
End Function

Private Function DiaUtil(ByVal d As Date) As Boolean
    DiaUtil = (Weekday(d, vbMonday) <= 5)
End Function

Private Function Janelas() As Variant
    Janelas = Array(TimeSerial(16, 0, 0), TimeSerial(12, 30, 0), TimeSerial(9, 0, 0))   ' da mais tarde para a mais cedo
End Function

'----------------------------------------------------------
' Motivos do aviso, separados por "; " ("" = backups em dia).
' instantes/tipos/validados: listas paralelas (Empty = nenhuma copia).
'----------------------------------------------------------
Public Function MotivosAviso(ByVal agora As Date, ByVal instantes As Variant, _
                             ByVal tipos As Variant, ByVal validados As Variant) As String
    Dim k As Long, j As Variant, dia As Date, ini As Date, ultima As Date, diaB As Date
    Dim i As Long, tem As Boolean, achou As Boolean, m As String
    ' (a) a ultima janela encerrada
    For k = 0 To 10
        dia = Int(agora) - k
        If DiaUtil(dia) Then
            For Each j In Janelas()
                ini = dia + j
                If agora >= ini + TimeSerial(0, ENCERRA_MIN, 0) - TimeSerial(0, 0, 1) / 2 Then ultima = ini: achou = True: Exit For
            Next j
        End If
        If achou Then Exit For
    Next k
    If achou Then
        tem = False
        If IsArray(instantes) Then
            For i = LBound(instantes) To UBound(instantes)
                If instantes(i) >= ultima And instantes(i) < ultima + TimeSerial(0, CONTA_MIN, 0) Then tem = True
            Next i
        End If
        If Not tem Then m = "a janela das " & Format$(ultima, "hh:nn") & " de " & Format$(ultima, "dd/mm/yyyy") & " terminou sem c" & ChrW$(243) & "pia"
    End If
    ' (b) o ultimo dia util cujas 16h30 ja passaram
    achou = False
    For k = 0 To 10
        dia = Int(agora) - k
        If DiaUtil(dia) And agora >= dia + TimeSerial(16, 30, 0) - TimeSerial(0, 0, 1) / 2 Then diaB = dia: achou = True: Exit For
    Next k
    If achou Then
        tem = False
        If IsArray(instantes) Then
            For i = LBound(instantes) To UBound(instantes)
                If tipos(i) = "diario" And validados(i) And Int(instantes(i)) = diaB Then tem = True
            Next i
        End If
        If Not tem Then
            If m <> "" Then m = m & "; "
            m = m & "o dia " & Format$(diaB, "dd/mm/yyyy") & " terminou sem c" & ChrW$(243) & "pia completa validada"
        End If
    End If
    MotivosAviso = m
End Function

' Pasta de backup, a partir da pasta do banco (Config!B2): <BASE>\SISTEMA\DADOS\execucao\backup\
Public Function PastaBackup() As String
    Dim c As String, k As Long
    c = modConfig.CaminhoBanco()
    k = InStrRev(c, "\")
    If k = 0 Then Exit Function
    PastaBackup = Left$(c, k) & "SISTEMA\DADOS\execucao\backup\"
End Function

Private Function Existe(ByVal caminho As String) As Boolean
    On Error Resume Next
    Existe = ((GetAttr(caminho) And vbDirectory) = 0)
    If Err.Number <> 0 Then Existe = False
    On Error GoTo 0
End Function

' DEV-03: a copia exclusiva do backup esta segurando o arquivo do banco agora
Public Function BackupEmCurso() As Boolean
    Dim pasta As String
    pasta = PastaBackup()
    If pasta = "" Then Exit Function
    BackupEmCurso = Existe(pasta & "copiando.txt")
End Function

' Motivos do aviso para a pasta real ("" = em dia). Erro de leitura devolve "-" (sem opiniao).
Public Function MotivosDaPasta() As String
    Dim pasta As String, nome As String, nomes() As String, n As Long, i As Long
    Dim instantes() As Date, tipos() As String, validados() As Boolean, d As Date, q As Long
    On Error GoTo semLeitura
    pasta = PastaBackup()
    If pasta = "" Then MotivosDaPasta = "-": Exit Function
    If Dir(pasta, vbDirectory) = "" Then MotivosDaPasta = "-": Exit Function
    ' primeiro so os nomes (Dir nao pode ser chamado de novo no meio da enumeracao)
    nome = Dir(pasta & "crm_*.accdb")
    Do While nome <> ""
        ReDim Preserve nomes(0 To n)
        nomes(n) = nome
        n = n + 1
        nome = Dir()
    Loop
    If n = 0 Then
        MotivosDaPasta = MotivosAviso(Now, Empty, Empty, Empty)
        Exit Function
    End If
    For i = 0 To n - 1
        d = InstanteDoNome(nomes(i))
        If d > 0 And LCase$(Right$(nomes(i), 6)) = ".accdb" Then
            ReDim Preserve instantes(0 To q): ReDim Preserve tipos(0 To q): ReDim Preserve validados(0 To q)
            instantes(q) = d
            tipos(q) = TipoDoNome(nomes(i))
            validados(q) = Existe(pasta & nomes(i) & ".validado")
            q = q + 1
        End If
    Next i
    If q = 0 Then
        MotivosDaPasta = MotivosAviso(Now, Empty, Empty, Empty)
    Else
        MotivosDaPasta = MotivosAviso(Now, instantes, tipos, validados)
    End If
    Exit Function
semLeitura:
    MotivosDaPasta = "-"
End Function

' Chamado na abertura (ThisWorkbook): avisa uma vez, so quando ha motivo.
Public Sub AvisarSeAtrasado()
    Dim m As String
    On Error Resume Next
    m = MotivosDaPasta()
    If m = "" Or m = "-" Then Exit Sub
    MsgBox "Backup do banco atrasado: " & m & "." & vbCrLf & vbCrLf & _
           "Avise quem cuida dos backups (FAZER-BACKUP.bat e TESTAR-RESTAURACAO.bat, " & _
           "na pasta SISTEMA\DADOS\operacao).", vbExclamation, "CRM Zapromaq - backup"
End Sub
