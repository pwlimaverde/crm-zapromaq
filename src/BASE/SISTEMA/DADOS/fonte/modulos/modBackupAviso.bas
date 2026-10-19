Attribute VB_Name = "modBackupAviso"
Option Explicit
'==========================================================
' modBackupAviso - alerta de backup (spec 006)
'
' O backup (backup-banco.ps1) grava ao lado de cada copia validada uma marca
'   crm_<tipo>_AAAAMMDD-HHMMSS_<ESTACAO>.accdb.validado
' em <BASE>\SISTEMA\DADOS\execucao\backup. Ao abrir o front, se nao ha copia
' validada ha mais de 36 h (RPO 24 h + folga), avisa uma vez. Se a pasta nao
' puder ser lida (rede fora do ar, pasta inexistente) nao diz nada: o aviso
' de banco inacessivel ja cobre esse caso.
'==========================================================
Public Const HORAS_ALERTA_BACKUP As Long = 36

' "crm_diario_20261018-081530_EST1.accdb.validado" -> 18/10/2026 08:15:30 (0 = nome fora do padrao)
Public Function InstanteDoNome(ByVal nome As String) As Date
    Dim p1 As Long, p2 As Long, s As String
    On Error GoTo fora
    p1 = InStr(nome, "_")
    If p1 = 0 Then Exit Function
    p2 = InStr(p1 + 1, nome, "_")
    If p2 = 0 Then Exit Function
    s = Mid$(nome, p2 + 1)
    If Mid$(s, 9, 1) <> "-" Then Exit Function
    If Not IsNumeric(Left$(s, 8)) Or Not IsNumeric(Mid$(s, 10, 6)) Then Exit Function
    InstanteDoNome = DateSerial(CInt(Left$(s, 4)), CInt(Mid$(s, 5, 2)), CInt(Mid$(s, 7, 2))) + _
                     TimeSerial(CInt(Mid$(s, 10, 2)), CInt(Mid$(s, 12, 2)), CInt(Mid$(s, 14, 2)))
    Exit Function
fora:
    InstanteDoNome = 0
End Function

Public Function Atrasado(ByVal ultimo As Date, ByVal agora As Date) As Boolean
    If ultimo = 0 Then
        Atrasado = True
    Else
        Atrasado = ((agora - ultimo) * 24 > HORAS_ALERTA_BACKUP)
    End If
End Function

' Pasta de backup, a partir da pasta do banco (Config!B2): <BASE>\SISTEMA\DADOS\execucao\backup\
Public Function PastaBackup() As String
    Dim c As String, k As Long
    c = modConfig.CaminhoBanco()
    k = InStrRev(c, "\")
    If k = 0 Then Exit Function
    PastaBackup = Left$(c, k) & "SISTEMA\DADOS\execucao\backup\"
End Function

' Instante da copia validada mais recente (0 = nenhuma). Erro de leitura devolve -1 (sem opiniao).
Public Function UltimoValidado() As Date
    Dim pasta As String, nome As String, d As Date, melhor As Date
    On Error GoTo semLeitura
    pasta = PastaBackup()
    If pasta = "" Then UltimoValidado = -1: Exit Function
    If Dir(pasta, vbDirectory) = "" Then UltimoValidado = -1: Exit Function
    nome = Dir(pasta & "crm_*.validado")
    Do While nome <> ""
        d = InstanteDoNome(nome)
        If d > melhor Then melhor = d
        nome = Dir()
    Loop
    UltimoValidado = melhor
    Exit Function
semLeitura:
    UltimoValidado = -1
End Function

Public Sub AvisarSeAtrasado()
    Dim u As Date
    On Error Resume Next
    u = UltimoValidado()
    If u = -1 Then Exit Sub
    If Atrasado(u, Now) Then
        MsgBox IIf(u = 0, "Nao ha nenhum backup validado do banco.", _
               "O ultimo backup validado do banco e de " & Format$(u, "dd/mm/yyyy hh:nn") & ".") & vbCrLf & vbCrLf & _
               "O limite e de " & HORAS_ALERTA_BACKUP & " horas. Avise quem cuida dos backups " & _
               "(FAZER-BACKUP.bat e TESTAR-RESTAURACAO.bat, na pasta SISTEMA\DADOS\operacao).", _
               vbExclamation, "CRM Zapromaq - backup"
    End If
End Sub
