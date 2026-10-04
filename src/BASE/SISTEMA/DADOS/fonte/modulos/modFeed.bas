Attribute VB_Name = "modFeed"
'==========================================================
' modFeed - atualizacao automatica das estacoes (D07).
'
' A cada INTERVALO_S segundos consulta MAX(id) de
' log_alteracoes. Se subiu, busca numa consulta so QUAIS
' registros mudaram desde a ultima vez e reescreve na grade
' so essas linhas (modGrade.AtualizarRegistros). Toda
' gravacao - desta estacao, de outra ou da IA - passa pelo
' log dentro da transacao, entao o log e o feed.
'
' Custo: uma consulta de MAX por intervalo (~26 ms na rede).
' Nada mudou = nada mais acontece.
'
' Regras do Application.OnTime (referencias/04-excel, C05):
'  - a rotina fica em modulo padrao, publica;
'  - cancelar exige o MESMO Procedure e o MESMO EarliestTime:
'    a hora agendada fica guardada em mProxima;
'  - sem cancelar no fechamento, o Excel REABRE o arquivo para
'    rodar a rotina agendada (ThisWorkbook.BeforeClose chama
'    Parar);
'  - so roda com o Excel em modo Pronto: digitando numa
'    celula, espera. Sem LatestTime de proposito: se a janela
'    passasse, a rotina nao rodaria e a corrente de
'    agendamentos morreria. Nao empilha porque cada rodada
'    agenda so a seguinte, no fim.
'
' Registro novo ou excluido nao entra sozinho na grade (isso
' exigiria remontar a aba e perder filtro e ordenacao): a
' linha de status avisa e o usuario clica em Atualizar.
' Ficha aberta num registro alterado por outro recebe aviso
' (frmCRM.AvisoExterno) - o controle de versao continua sendo
' quem impede a sobrescrita.
'==========================================================
Option Explicit

Private Const INTERVALO_S As Long = 25
Private Const ROTINA As String = "modFeed.Verificar"

Private mProxima As Date           ' hora agendada: cancelar exige a mesma
Private mAgendado As Boolean
Private mAtivo As Boolean
Private mUltimoLog As Long         ' maior id de log_alteracoes ja processado
Private mFalhasSeguidas As Long
Private mDesejado As Boolean       ' Iniciar ja rodou nesta sessao

'----------------------------------------------------------
' Liga o feed (Workbook_Open, depois de conferir o banco).
' A grade aberta agora e a referencia: o que veio antes dela
' ja esta na tela ou aparece no proximo Atualizar.
'----------------------------------------------------------
Public Sub Iniciar()
    If modConfig.gModoModelo Then Exit Sub
    If mAtivo Then Exit Sub
    mDesejado = True
    On Error GoTo semBanco
    mUltimoLog = MaiorId()
    mAtivo = True
    mFalhasSeguidas = 0
    Agendar
    Exit Sub
semBanco:
    ' sem banco agora: tenta de novo no proximo intervalo
    mUltimoLog = -1
    mAtivo = True
    Agendar
End Sub

'----------------------------------------------------------
' Desliga e cancela o agendamento pendente (BeforeClose).
'----------------------------------------------------------
Public Sub Parar()
    mAtivo = False
    If Not mAgendado Then Exit Sub
    On Error Resume Next
    Application.OnTime EarliestTime:=mProxima, Procedure:=NomeRotina(), Schedule:=False
    On Error GoTo 0
    mAgendado = False
End Sub

'----------------------------------------------------------
' O BeforeClose para o feed, mas o fechamento ainda pode ser
' cancelado (Cancelar na pergunta de salvar). A proxima troca
' de aba religa, sem perder a referencia do que ja foi visto.
'----------------------------------------------------------
Public Sub Retomar()
    If Not mDesejado Or mAtivo Then Exit Sub
    If modConfig.gModoModelo Then Exit Sub
    mAtivo = True
    On Error Resume Next
    Agendar
    On Error GoTo 0
End Sub

Public Function Ligado() As Boolean
    Ligado = mAtivo
End Function

Private Function NomeRotina() As String
    ' qualificado pelo arquivo: com outra pasta aberta que tenha um
    ' modFeed, o Excel poderia chamar a rotina errada
    NomeRotina = "'" & ThisWorkbook.Name & "'!" & ROTINA
End Function

Private Sub Agendar()
    Dim espera As Long
    ' banco fora do ar: espaca as tentativas (25 s, 50 s ... ate 5 min)
    espera = INTERVALO_S * (1 + mFalhasSeguidas)
    If espera > 300 Then espera = 300
    mProxima = Now + TimeSerial(0, 0, espera)
    Application.OnTime EarliestTime:=mProxima, Procedure:=NomeRotina()
    mAgendado = True
End Sub

'----------------------------------------------------------
' Chamada pelo OnTime. Nunca mostra MsgBox: rodando sozinha a
' cada 25 s, uma caixa de erro travaria o usuario em loop.
'----------------------------------------------------------
Public Sub Verificar()
    mAgendado = False
    If Not mAtivo Then Exit Sub
    On Error GoTo falha
    Processar
    mFalhasSeguidas = 0
    Agendar
    Exit Sub
falha:
    mFalhasSeguidas = mFalhasSeguidas + 1
    On Error Resume Next
    Application.StatusBar = "CRM Zapromaq: atualiza" & ChrW$(&HE7) & ChrW$(&HE3) & "o autom" & ChrW$(&HE1) & _
                            "tica sem banco agora (" & Format$(Now, "hh:nn") & "); tentando de novo."
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    If mAtivo Then Agendar
End Sub

Private Function MaiorId() As Long
    MaiorId = CLng(modDB.ConsultarValor("SELECT MAX(id) FROM log_alteracoes", 0))
End Function

'----------------------------------------------------------
Private Sub Processar()
    Dim ate As Long, d As Variant, i As Long, tabela As String, idReg As Long
    Dim porTabela As Object, deOutros As Object, chave As Variant

    ate = MaiorId()
    If mUltimoLog < 0 Then mUltimoLog = ate: Exit Sub   ' voltou o banco: nova referencia
    If ate <= mUltimoLog Then Exit Sub

    ' UMA consulta: o que mudou, de quem. Faixa fechada (ate) para
    ' nao perder o que entrar entre esta consulta e a proxima.
    d = modDB.Consultar("SELECT tabela, id_registro, usuario, origem FROM log_alteracoes " & _
                        "WHERE id > ? AND id <= ?", _
                        Array(modDB.P(modDB.adInteger, mUltimoLog), modDB.P(modDB.adInteger, ate)))
    mUltimoLog = ate
    If IsEmpty(d) Then Exit Sub

    Set porTabela = CreateObject("Scripting.Dictionary")
    Set deOutros = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(d, 1)
        tabela = LCase$(modDB.Nz(d(i, 1)))
        idReg = CLng(modDB.Nz(d(i, 2), "0"))
        If tabela <> "" And idReg > 0 Then
            If Not porTabela.Exists(tabela) Then porTabela.Add tabela, CreateObject("Scripting.Dictionary")
            If Not porTabela(tabela).Exists(idReg) Then porTabela(tabela).Add idReg, 0
            ' gravacao desta propria estacao nao e aviso para a ficha
            If Not (modDB.Nz(d(i, 3)) = modConfig.UsuarioAtual() And modDB.Nz(d(i, 4)) = "FRONT") Then
                If Not deOutros.Exists(tabela) Then deOutros.Add tabela, CreateObject("Scripting.Dictionary")
                If Not deOutros(tabela).Exists(idReg) Then deOutros(tabela).Add idReg, 0
            End If
        End If
    Next i

    For Each chave In porTabela.Keys
        AtualizarAba CStr(chave), porTabela(chave)
    Next chave
    For Each chave In deOutros.Keys
        AvisarFichas CStr(chave), deOutros(chave)
    Next chave
End Sub

'----------------------------------------------------------
' Reescreve as linhas na aba da tabela; o que faltou (novo ou
' excluido) vira aviso na linha de status.
'----------------------------------------------------------
Private Sub AtualizarAba(ByVal tabela As String, ByVal ids As Object)
    Dim ws As Worksheet, faltaram As Long
    For Each ws In ThisWorkbook.Worksheets
        If modGrade.TabelaDaAba(ws) = tabela Then
            If ws.ListObjects.Count > 0 Then
                faltaram = modGrade.AtualizarRegistros(ws, ids, False)
                If faltaram > 0 Then modGrade.AvisarPendentes ws, faltaram
            End If
        End If
    Next ws
End Sub

Private Sub AvisarFichas(ByVal tabela As String, ByVal ids As Object)
    Dim f As Object
    For Each f In VBA.UserForms
        If TypeName(f) = "frmCRM" Then
            On Error Resume Next
            f.AvisoExterno tabela, ids
            On Error GoTo 0
        End If
    Next f
End Sub
