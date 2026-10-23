Attribute VB_Name = "modFeed"
'==========================================================
' modFeed - atualizacao automatica das estacoes (D07).
'
' A cada INTERVALO_S segundos consulta MAX(id) de
' log_alteracoes. Se subiu, busca numa consulta so QUAIS
' registros mudaram desde a ultima vez e reescreve na grade
' so essas linhas (modGrade.AtualizarPeloFeed). Toda
' gravacao - desta estacao, de outra ou da IA - passa pelo
' log dentro da transacao, entao o log e o feed.
'
' Custo: duas leituras curtas por intervalo, numa conexao
' (MAX do log e a geracao da base). Nada mudou = nada mais.
'
' Feed confiavel (ciclo 011, DEV-18):
'  - o cursor (mUltimoLog) so avanca DEPOIS de aplicar todas
'    as abas e avisos: se algo falha no meio, a proxima rodada
'    le o mesmo lote de novo (reescrever linha e idempotente);
'  - commit fora de ordem: id do log que faltou na faixa lida
'    (transacao ainda aberta) vira LACUNA e e procurado de novo
'    por LACUNA_S segundos; depois disso era transacao desfeita;
'  - geracao da base (config.geracao_base) trocada ou MAX(id)
'    menor que o cursor (banco restaurado): referencia zerada e
'    as abas carregadas sao remontadas.
' Dependencias e dia (DEV-19): Dependentes() diz o que cada
' tabela alterada invalida; a virada do dia remonta as abas
' (situacao e dias parado sao calculados aqui no front).
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
' Ficha aberta num registro alterado por outro recebe aviso
' (frmCRM.AvisoExterno) - os dados da ficha em edicao nao sao
' trocados; o controle de versao impede a sobrescrita.
'==========================================================
Option Explicit

Private Const INTERVALO_S As Long = 25
Private Const ROTINA As String = "modFeed.Verificar"
Private Const LACUNA_S As Long = 300       ' quanto tempo um id que faltou e procurado
Private Const MAX_LACUNAS As Long = 200
Private Const MAX_IDS As Long = 200        ' acima disso os dependentes nao sao buscados um a um

Private mProxima As Date           ' hora agendada: cancelar exige a mesma
Private mAgendado As Boolean
Private mAtivo As Boolean
Private mReferencia As Boolean     ' cursor e geracao ja lidos do banco
Private mUltimoLog As Long         ' maior id de log_alteracoes ja aplicado
Private mLacunas As Object         ' id do log que faltou -> instante em que faltou
Private mGeracao As String         ' config.geracao_base da referencia
Private mUltimaCarga As Date       ' ultima vez que a tela foi atualizada COM sucesso
Private mDiaCalculo As Date        ' dia em que situacao e dias parado foram calculados
Private mFalhaSimulada As Boolean  ' teste de uso: a proxima aplicacao falha
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
    mReferencia = False
    On Error Resume Next
    LerReferencia
    On Error GoTo 0
    mAtivo = True
    mFalhasSeguidas = 0
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

' Ultima atualizacao EFETIVA da tela (so depois de aplicar com sucesso).
' A ultima consulta ao banco fica em modConfig.UltimaConsulta.
Public Function UltimaCarga() As Date
    UltimaCarga = mUltimaCarga
End Function

' Dia em que situacao e dias parado foram calculados nas abas abertas.
Public Function DiaDoCalculo() As Date
    DiaDoCalculo = mDiaCalculo
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

' Processa o que chegou no log agora, sem agendar nada. O Verificar usa isto;
' o teste de uso chama direto, sem ligar o relogio do feed.
Public Sub ProcessarAgora()
    Processar
End Sub

' Teste de uso (DEV-18/19): a proxima aplicacao falha no meio; e o dia do calculo
' volta um dia, como se a meia-noite tivesse passado com o CRM aberto.
Public Sub SimularFalhaNaProximaAplicacao()
    mFalhaSimulada = True
End Sub

Public Sub SimularDiaAnterior()
    mDiaCalculo = Date - 1
End Sub

Private Sub LerReferencia()
    Dim cn As Object
    mReferencia = False
    Set cn = modDB.AbrirLeitura()
    On Error GoTo falha
    mUltimoLog = MaiorIdEm(cn)
    mGeracao = GeracaoEm(cn)
    Set mLacunas = CreateObject("Scripting.Dictionary")
    If mDiaCalculo = 0 Then mDiaCalculo = Date
    modDB.FecharLeitura cn
    mReferencia = True
    Exit Sub
falha:
    Dim n As Long, s As String
    n = Err.Number: s = Err.Description
    modDB.FecharLeitura cn
    Err.Raise n, "modFeed.LerReferencia", s
End Sub

Private Function MaiorIdEm(ByVal cn As Object) As Long
    MaiorIdEm = CLng(modDB.ConsultarValorEm(cn, "SELECT MAX(id) FROM log_alteracoes", 0))
End Function

Private Function GeracaoEm(ByVal cn As Object) As String
    GeracaoEm = CStr(modDB.ConsultarValorEm(cn, "SELECT valor FROM config WHERE chave='geracao_base'", ""))
End Function

'----------------------------------------------------------
' LOGICA PURA (testada no modAutoteste)
'----------------------------------------------------------

' Referencia do cursor invalida: a base trocou de geracao (restaurada ou
' substituida) ou o log tem menos do que ja foi visto (cursor a frente).
Public Function ReferenciaInvalida(ByVal geracaoGuardada As String, ByVal geracaoAtual As String, _
                                   ByVal cursor As Long, ByVal maiorId As Long) As Boolean
    If geracaoGuardada <> "" And geracaoAtual <> "" And geracaoGuardada <> geracaoAtual Then
        ReferenciaInvalida = True
    ElseIf maiorId < cursor Then
        ReferenciaInvalida = True
    End If
End Function

' Lacunas depois de aplicar a faixa (de, ate]: as antigas que nao apareceram
' e ainda estao no prazo continuam; todo id da faixa que nao foi visto entra
' (transacao ainda aberta na hora da leitura, ou desfeita).
Public Function NovasLacunas(ByVal de As Long, ByVal ate As Long, ByVal vistos As Object, _
                             ByVal antigas As Object, ByVal agora As Date) As Object
    Dim r As Object, k As Variant, i As Long
    Set r = CreateObject("Scripting.Dictionary")
    If Not antigas Is Nothing Then
        For Each k In antigas.Keys
            If Not vistos.Exists(CLng(k)) Then
                If DateDiff("s", antigas(k), agora) < LACUNA_S Then r.Add CLng(k), antigas(k)
            End If
        Next k
    End If
    For i = de + 1 To ate
        If r.Count >= MAX_LACUNAS Then Exit For
        If Not vistos.Exists(i) Then If Not r.Exists(i) Then r.Add i, agora
    Next i
    Set NovasLacunas = r
End Function

' Mapa de dependencias (DEV-19): o que a alteracao de uma tabela invalida
' alem das proprias linhas. "painel" = numeros do Painel; "cache" = listas
' das combos. Separado por virgula.
Public Function Dependentes(ByVal tabela As String) As String
    Select Case LCase$(tabela)
        Case "clientes":      Dependentes = "contatos,oportunidades,painel"
        Case "contatos":      Dependentes = "oportunidades,painel"
        Case "oportunidades": Dependentes = "painel"
        Case "listas":        Dependentes = "cache"
    End Select
End Function

'----------------------------------------------------------
Private Sub Processar()
    Dim cn As Object, ate As Long, geracao As String, inicio As Long, k As Variant
    Dim d As Variant, i As Long, idLog As Long, tabela As String, idReg As Long
    Dim vistos As Object, porTabela As Object, deOutros As Object, extras As Object
    Dim virouDia As Boolean, v As Variant

    ' primeira rodada (ou o banco voltou): o que esta na tela e a referencia
    If Not mReferencia Then LerReferencia: Exit Sub

    Set cn = modDB.AbrirLeitura()
    On Error GoTo falhaLeitura
    ate = MaiorIdEm(cn)
    geracao = GeracaoEm(cn)

    If ReferenciaInvalida(mGeracao, geracao, mUltimoLog, ate) Then
        modDB.FecharLeitura cn
        RecarregarTudo
        mUltimoLog = ate: mGeracao = geracao
        Set mLacunas = CreateObject("Scripting.Dictionary")
        mDiaCalculo = Date: mUltimaCarga = Now
        Exit Sub
    End If
    If mGeracao = "" Then mGeracao = geracao
    virouDia = (mDiaCalculo <> Date)

    If ate <= mUltimoLog And mLacunas.Count = 0 Then
        modDB.FecharLeitura cn
        If virouDia Then
            RecarregarTudo
            mDiaCalculo = Date: mUltimaCarga = Now
        End If
        Exit Sub
    End If

    ' faixa lida: do menor id ainda em lacuna (commit atrasado) ate o MAX
    inicio = mUltimoLog
    For Each k In mLacunas.Keys
        If CLng(k) - 1 < inicio Then inicio = CLng(k) - 1
    Next k
    d = modDB.ConsultarEm(cn, "SELECT id, tabela, id_registro, usuario, origem FROM log_alteracoes " & _
                              "WHERE id > ? AND id <= ?", _
                              Array(modDB.P(modDB.adInteger, inicio), modDB.P(modDB.adInteger, ate)))

    Set vistos = CreateObject("Scripting.Dictionary")
    Set porTabela = CreateObject("Scripting.Dictionary")
    Set deOutros = CreateObject("Scripting.Dictionary")
    If Not IsEmpty(d) Then
        For i = 1 To UBound(d, 1)
            idLog = CLng(d(i, 1))
            ' abaixo do cursor, so o que era lacuna: o resto ja foi aplicado
            If idLog > mUltimoLog Or mLacunas.Exists(idLog) Then
                If Not vistos.Exists(idLog) Then vistos.Add idLog, 0
                tabela = LCase$(modDB.Nz(d(i, 2)))
                idReg = CLng(modDB.Nz(d(i, 3), "0"))
                If tabela <> "" And idReg > 0 Then
                    Juntar porTabela, tabela, idReg
                    ' gravacao desta propria estacao nao e aviso para a ficha
                    If Not (modDB.Nz(d(i, 4)) = modConfig.UsuarioAtual() And modDB.Nz(d(i, 5)) = "FRONT") Then
                        Juntar deOutros, tabela, idReg
                    End If
                End If
            End If
        Next i
    End If

    ' dependentes que estao na tela: uma consulta por tabela dependente
    Set extras = IdsDependentes(cn, porTabela)
    modDB.FecharLeitura cn
    For Each k In extras.Keys
        For Each v In extras(k).Keys
            Juntar porTabela, CStr(k), CLng(v)
        Next v
    Next k

    ' ---- aplicar: qualquer erro daqui para baixo deixa o cursor onde estava
    Aplicar porTabela, deOutros, virouDia

    Set mLacunas = NovasLacunas(mUltimoLog, ate, vistos, mLacunas, Now)
    If ate > mUltimoLog Then mUltimoLog = ate
    If virouDia Then mDiaCalculo = Date
    mUltimaCarga = Now
    Exit Sub
falhaLeitura:
    Dim n As Long, s As String
    n = Err.Number: s = Err.Description
    modDB.FecharLeitura cn
    Err.Raise n, "modFeed.Processar", s
End Sub

Private Sub Juntar(ByVal grupos As Object, ByVal tabela As String, ByVal id As Long)
    If Not grupos.Exists(tabela) Then grupos.Add tabela, CreateObject("Scripting.Dictionary")
    If Not grupos(tabela).Exists(id) Then grupos(tabela).Add id, 0
End Sub

' Ids de linhas que dependem das alteradas (telefone do contato e empresa
' aparecem no atendimento...), so para as abas carregadas.
Private Function IdsDependentes(ByVal cn As Object, ByVal porTabela As Object) As Object
    Dim r As Object
    Set r = CreateObject("Scripting.Dictionary")
    If porTabela.Exists("clientes") Then
        Buscar cn, r, "contatos", "id_cliente", porTabela("clientes")
        Buscar cn, r, "oportunidades", "id_cliente", porTabela("clientes")
    End If
    If porTabela.Exists("contatos") Then
        Buscar cn, r, "oportunidades", "id_contato", porTabela("contatos")
    End If
    Set IdsDependentes = r
End Function

Private Sub Buscar(ByVal cn As Object, ByVal r As Object, ByVal tabela As String, _
                   ByVal coluna As String, ByVal ids As Object)
    Dim marcas As String, params As Variant, d As Variant, i As Long
    If Not AbaCarregada(tabela) Then Exit Sub
    If ids.Count = 0 Or ids.Count > MAX_IDS Then Exit Sub
    marcas = modDB.MarcadoresIds(ids.Keys, params)
    ' tabela e coluna vem do proprio codigo (nunca de digitacao)
    d = modDB.ConsultarEm(cn, "SELECT id FROM " & tabela & " WHERE " & coluna & " IN (" & marcas & ")", params)
    If IsEmpty(d) Then Exit Sub
    If Not r.Exists(tabela) Then r.Add tabela, CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(d, 1)
        If Not r(tabela).Exists(CLng(d(i, 1))) Then r(tabela).Add CLng(d(i, 1)), 0
    Next i
End Sub

Private Function AbaCarregada(ByVal tabela As String) As Boolean
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If modGrade.TabelaDaAba(ws) = tabela Then
            If ws.ListObjects.Count > 0 Then AbaCarregada = True: Exit Function
        End If
    Next ws
End Function

Private Sub Aplicar(ByVal porTabela As Object, ByVal deOutros As Object, ByVal virouDia As Boolean)
    Dim chave As Variant, deps As String

    If mFalhaSimulada Then
        mFalhaSimulada = False
        Err.Raise vbObjectError + 611, "modFeed.Aplicar", "falha simulada na primeira aba (teste de uso)"
    End If

    For Each chave In porTabela.Keys
        deps = "," & Dependentes(CStr(chave)) & ","
        If InStr(deps, ",cache,") > 0 Then modCRM.LimparCacheListas
        If InStr(deps, ",painel,") > 0 Then modPainel.MarcarVelho
    Next chave

    If virouDia Then
        RecarregarAbas
    Else
        For Each chave In porTabela.Keys
            AtualizarAba CStr(chave), porTabela(chave)
        Next chave
    End If
    For Each chave In deOutros.Keys
        AvisarFichas CStr(chave), deOutros(chave)
    Next chave
    ' aba Planejamentos (fatia 4.8): uma consulta, so se ja foi carregada
    If porTabela.Exists("visitas") Or porTabela.Exists("planejamentos") Then modAbaPlanejamentos.AtualizarPeloFeed
End Sub

' Base restaurada ou dia novo: tudo o que esta na tela e refeito, com filtro,
' pesquisa, ordem e selecao preservados.
Private Sub RecarregarTudo()
    modCRM.LimparCacheListas
    modPainel.MarcarVelho
    RecarregarAbas
    modAbaPlanejamentos.AtualizarPeloFeed
End Sub

Private Sub RecarregarAbas()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If modGrade.TabelaDaAba(ws) <> "" Then
            If ws.ListObjects.Count > 0 Then modGrade.Remontar ws
        End If
    Next ws
End Sub

'----------------------------------------------------------
' Reescreve as linhas na aba da tabela, reaplica filtro e
' ordem; linha nova ou excluida remonta a aba (DEV-20).
'----------------------------------------------------------
Private Sub AtualizarAba(ByVal tabela As String, ByVal ids As Object)
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If modGrade.TabelaDaAba(ws) = tabela Then
            If ws.ListObjects.Count > 0 Then modGrade.AtualizarLinhasDoFeed ws, ids
        End If
    Next ws
End Sub

' Telas abertas que sabem receber o aviso: a ficha (frmCRM), as visitas
' do atendimento (frmVisitas) e o planejamento (frmPlanejamento). Cada
' uma decide o que fazer com a tabela e os ids (fatia 3.3b).
Private Sub AvisarFichas(ByVal tabela As String, ByVal ids As Object)
    Dim f As Object
    For Each f In VBA.UserForms
        Select Case TypeName(f)
            Case "frmCRM", "frmVisitas", "frmPlanejamento"
                On Error Resume Next
                f.AvisoExterno tabela, ids
                On Error GoTo 0
        End Select
    Next f
End Sub
