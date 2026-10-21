Attribute VB_Name = "modDB"
'==========================================================
' modDB - acesso ao banco Access (ACE OLEDB) por ADO
'
' Tres regras que nao se alteram (docs\contrato-banco.md):
'   1. Conexao CURTA, por OPERACAO DO USUARIO - nunca por
'      linha, campo ou combo. Medido: 26 ms por ciclo na
'      rede; dentro de um laco de 500 linhas sao 13 segundos.
'   2. PARAMETROS, sempre. Concatenar valor quebra com
'      apostrofo (ha "INDUSTRIA D'ANGELO LTDA" na base) e
'      abre porta para erro de tipo.
'   3. Data NUNCA como literal. O literal do ACE e mm/dd/yyyy
'      e a estacao escreve dd/mm/yyyy: #03/01/2026# vira
'      1o de marco sem erro e sem aviso.
'
' Duas formas de uso:
'   - operacao simples: Consultar, ConsultarValor, Executar
'     (cada chamada abre e fecha a propria conexao);
'   - operacao que grava mais de uma coisa: AbrirTransacao,
'     depois ...Em(cn, ...) e LogEm, e por fim Confirmar ou
'     Desfazer. Tudo ou nada, e o log vai JUNTO com o dado.
'==========================================================
Option Explicit

' Constantes ADO (late binding - sem referencia a biblioteca)
Public Const adVarWChar As Long = 202
Public Const adLongVarWChar As Long = 203
Public Const adDate As Long = 7
Public Const adInteger As Long = 3
Public Const adCurrency As Long = 6
Public Const adBoolean As Long = 11
' Gravacao recusada por manutencao, esquema ou backup em curso (DEV-03, DEV-08)
Public Const ERRO_ESCRITA_RECUSADA As Long = vbObjectError + 7001
Private Const adParamInput As Long = 1
Private Const adCmdText As Long = 1

' Nomes das colunas da ultima consulta (base 1)
Public gCampos As Variant

' Provedor que abriu da ultima vez (16.0 ou 12.0). Sem isto, com a
' rede fora do ar, TODA conexao esperava o tempo limite duas vezes
' (16.0 e depois 12.0) - ate 30 s de Excel travado por clique.
Private mProvedor As String

'----------------------------------------------------------
' Monta um parametro: use sempre P() ao montar a chamada.
'   modDB.Executar sql, Array(P(adVarWChar, nome), P(adDate, dt))
'----------------------------------------------------------
Public Function P(ByVal tipo As Long, ByVal valor As Variant) As Variant
    P = Array(tipo, valor)
End Function

Public Function StringConexao(Optional ByVal provedor As String = "Microsoft.ACE.OLEDB.16.0") As String
    StringConexao = "Provider=" & provedor & ";" & _
                    "Data Source=" & modConfig.CaminhoBanco() & ";" & _
                    "Persist Security Info=False;"
End Function

Public Function NovaConexao() As Object
    Dim cn As Object
    Set cn = CreateObject("ADODB.Connection")
    cn.ConnectionTimeout = 15
    If mProvedor <> "" Then
        ' provedor ja conhecido: uma tentativa so
        cn.Open StringConexao(mProvedor)
        Set NovaConexao = cn
        Exit Function
    End If
    On Error GoTo tentaAntigo
    cn.Open StringConexao("Microsoft.ACE.OLEDB.16.0")
    mProvedor = "Microsoft.ACE.OLEDB.16.0"
    Set NovaConexao = cn
    Exit Function
tentaAntigo:
    ' estacao com ACE 12 em vez de 16 (so na primeira conexao)
    Err.Clear
    On Error GoTo 0
    cn.Open StringConexao("Microsoft.ACE.OLEDB.12.0")
    mProvedor = "Microsoft.ACE.OLEDB.12.0"
    Set NovaConexao = cn
End Function

'----------------------------------------------------------
' LEITURA EM LOTE - uma conexao para varias consultas.
' Abrir conexao e o que custa (arquivo na rede, .laccdb): o
' Painel fazia ~20 consultas com 20 conexoes. Agora:
'   Set cn = modDB.AbrirLeitura()
'   d = modDB.ConsultarEm(cn, ...) ... (quantas precisar)
'   modDB.FecharLeitura cn
'----------------------------------------------------------
Public Function AbrirLeitura() As Object
    Set AbrirLeitura = NovaConexao()
End Function

Public Sub FecharLeitura(ByVal cn As Object)
    On Error Resume Next
    If Not cn Is Nothing Then If cn.State = 1 Then cn.Close
    On Error GoTo 0
    modConfig.RegistrarUltimaConsulta
End Sub

'----------------------------------------------------------
' Indice nome da coluna -> posicao (base 1) da ultima consulta.
' Procurar a coluna pelo nome, celula por celula, custava
' centenas de milhares de comparacoes numa grade de 500 linhas.
'----------------------------------------------------------
Public Function IndiceDeCampos() As Object
    Dim d As Object, i As Long
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1                            ' vbTextCompare
    If Not IsEmpty(gCampos) Then
        For i = LBound(gCampos) To UBound(gCampos)
            If Not d.Exists(gCampos(i)) Then d.Add gCampos(i), i
        Next i
    End If
    Set IndiceDeCampos = d
End Function

'----------------------------------------------------------
Private Function NovoComando(ByVal cn As Object, ByVal sql As String, _
                             ByVal params As Variant) As Object
    Dim cmd As Object, i As Long, par As Object, spec As Variant
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = cn
    cmd.CommandType = adCmdText
    cmd.CommandText = sql
    If Not IsEmpty(params) Then
        If IsArray(params) Then
            For i = LBound(params) To UBound(params)
                spec = params(i)
                Set par = cmd.CreateParameter("p" & i, CLng(spec(0)), adParamInput, _
                                              TamanhoDe(CLng(spec(0))))
                If IsNull(spec(1)) Or IsEmpty(spec(1)) Then
                    par.Value = Null
                ElseIf VarType(spec(1)) = vbString Then
                    ' texto vazio grava NULO, nunca cadeia vazia: o indice
                    ' unico de CNPJ aceita varios nulos, nao varias vazias
                    If Len(Trim$(CStr(spec(1)))) = 0 Then par.Value = Null Else par.Value = spec(1)
                Else
                    par.Value = spec(1)
                End If
                cmd.Parameters.Append par
            Next i
        End If
    End If
    Set NovoComando = cmd
End Function

Private Function TamanhoDe(ByVal tipo As Long) As Long
    Select Case tipo
        Case adVarWChar:     TamanhoDe = 255
        Case adLongVarWChar: TamanhoDe = 1073741823
        Case Else:           TamanhoDe = 0
    End Select
End Function

'----------------------------------------------------------
' Recordset -> matriz base 1 (linha, campo); Empty se vazio.
' Guarda os nomes das colunas em gCampos.
'----------------------------------------------------------
Private Function LerMatriz(ByVal rs As Object) As Variant
    Dim i As Long, nomes() As String
    ReDim nomes(1 To rs.Fields.Count)
    For i = 0 To rs.Fields.Count - 1
        nomes(i + 1) = rs.Fields(i).Name
    Next i
    gCampos = nomes
    If rs.EOF Then
        LerMatriz = Empty
    Else
        LerMatriz = Transpor(rs.GetRows)
    End If
End Function

'==========================================================
' OPERACAO SIMPLES - cada chamada abre e fecha a conexao
'==========================================================
Public Function Consultar(ByVal sql As String, Optional ByVal params As Variant) As Variant
    Dim cn As Object, cmd As Object, rs As Object
    On Error GoTo trata
    Set cn = NovaConexao()
    Set cmd = NovoComando(cn, sql, params)
    Set rs = cmd.Execute
    Consultar = LerMatriz(rs)
    rs.Close: Set rs = Nothing
    Set cmd = Nothing
    cn.Close: Set cn = Nothing
    modConfig.RegistrarUltimaConsulta
    Exit Function
trata:
    Dim nErr As Long, sErr As String
    nErr = Err.Number: sErr = Err.Description
    LimparRecursos rs, cn
    Set cmd = Nothing
    Err.Raise nErr, "modDB.Consultar", sErr & vbCrLf & vbCrLf & "SQL: " & sql
End Function

Public Function ConsultarValor(ByVal sql As String, Optional ByVal padrao As Variant = Null, _
                               Optional ByVal params As Variant) As Variant
    Dim d As Variant
    d = Consultar(sql, params)
    If IsEmpty(d) Then
        ConsultarValor = padrao
    ElseIf IsNull(d(1, 1)) Then
        ConsultarValor = padrao
    Else
        ConsultarValor = d(1, 1)
    End If
End Function

'----------------------------------------------------------
' INSERT / UPDATE / DELETE avulso - devolve linhas afetadas.
' Para gravar dado de negocio use a transacao com LogEm:
' gravacao sem log nao aparece no feed das outras estacoes.
'----------------------------------------------------------
Public Function Executar(ByVal sql As String, Optional ByVal params As Variant) As Long
    Dim cn As Object, cmd As Object, n As Long
    On Error GoTo trata
    Set cn = NovaConexao()
    Set cmd = NovoComando(cn, sql, params)
    cmd.Execute n
    Set cmd = Nothing
    cn.Close: Set cn = Nothing
    Executar = n
    Exit Function
trata:
    Dim nErr As Long, sErr As String
    nErr = Err.Number: sErr = Err.Description
    LimparRecursos Nothing, cn
    Set cmd = Nothing
    Err.Raise nErr, "modDB.Executar", sErr & vbCrLf & vbCrLf & "SQL: " & sql
End Function

'----------------------------------------------------------
' INSERT que devolve o id novo, com o log de insercao na
' mesma transacao. @@IDENTITY e POR CONEXAO: tem de ser lido
' na MESMA conexao. MAX(id) numa conexao nova devolveria o id
' de outro usuario.
'----------------------------------------------------------
Public Function InserirComId(ByVal sql As String, Optional ByVal params As Variant, _
                             Optional ByVal logTabela As String = "", _
                             Optional ByVal logDetalhe As String = "") As Long
    Dim cn As Object, novoId As Long
    On Error GoTo trata
    Set cn = AbrirTransacao()
    novoId = InserirEm(cn, sql, params)
    If logTabela <> "" Then LogEm cn, logTabela, novoId, "", "", logDetalhe, "INSERIR"
    Confirmar cn
    InserirComId = novoId
    Exit Function
trata:
    Dim nErr As Long, sErr As String
    nErr = Err.Number: sErr = Err.Description
    Desfazer cn
    Err.Raise nErr, "modDB.InserirComId", sErr & vbCrLf & vbCrLf & "SQL: " & sql
End Function

'==========================================================
' TRANSACAO - uma conexao, varias instrucoes, tudo ou nada
'
' Padrao de uso (capture o erro ANTES de Desfazer: o On Error
' Resume Next de dentro dele limpa o objeto Err):
'
'   On Error GoTo falha
'   Set cn = modDB.AbrirTransacao()
'   n = modDB.ExecutarEm(cn, sql, params)
'   modDB.LogEm cn, "clientes", id, "campo", antigo, novo, "ALTERAR"
'   modDB.Confirmar cn
'   Exit Function
' falha:
'   nErr = Err.Number: sErr = Err.Description
'   modDB.Desfazer cn
'   Err.Raise nErr, "...", sErr
'==========================================================
Public Function AbrirTransacao() As Object
    Dim cn As Object, d As Variant, i As Long, esq As String, man As String
    Dim motivo As String, camposAntes As Variant
    Set cn = NovaConexao()
    cn.BeginTrans
    ' DEV-08: toda gravacao do front passa por aqui. Esquema e manutencao sao
    ' conferidos DENTRO da transacao, no ponto de escrita: a ficha aberta antes
    ' de uma migracao (ou da montagem) nao grava depois dela. gCampos e
    ' preservado: quem chamou pode estar no meio da leitura de outra consulta.
    camposAntes = gCampos
    On Error GoTo falhou
    d = ConsultarEm(cn, "SELECT chave, valor FROM config WHERE chave IN ('versao_esquema','manutencao')")
    On Error GoTo 0
    gCampos = camposAntes
    If Not IsEmpty(d) Then
        For i = 1 To UBound(d, 1)
            If Nz(d(i, 1)) = "versao_esquema" Then esq = Nz(d(i, 2)) Else man = Nz(d(i, 2))
        Next i
    End If
    motivo = MotivoRecusaEscrita(esq, man)
    If motivo <> "" Then
        Desfazer cn
        modDiag.Registrar "gravacao", "inicio", "", ERRO_ESCRITA_RECUSADA, motivo, "recusada"
        Err.Raise ERRO_ESCRITA_RECUSADA, "modDB.AbrirTransacao", motivo
    End If
    Set AbrirTransacao = cn
    Exit Function
falhou:
    Dim nErr As Long, sErr As String
    nErr = Err.Number: sErr = Err.Description
    gCampos = camposAntes
    Desfazer cn
    Err.Raise nErr, "modDB.AbrirTransacao", sErr
End Function

'----------------------------------------------------------
' DEV-08: motivo para recusar a escrita (vazio = pode gravar).
' Mesma regra de Get-MotivoRecusaEscrita (conector, Banco.ps1).
' manutencao = "estacao|usuario|PID|inicio|motivo" (config).
'----------------------------------------------------------
Public Function MotivoRecusaEscrita(ByVal esquema As String, ByVal manutencao As String) As String
    Dim partes As Variant, quem As String
    If Len(Trim$(manutencao)) > 0 Then
        partes = Split(manutencao, "|")
        If UBound(partes) >= 4 Then quem = " (" & partes(4) & ", por " & partes(0) & "\" & partes(1) & " desde " & partes(3) & ")"
        MotivoRecusaEscrita = "Sistema em manuten" & ChrW$(231) & ChrW$(227) & "o" & quem & _
            ": grava" & ChrW$(231) & ChrW$(227) & "o recusada, nada foi gravado. Tente de novo quando a manuten" & _
            ChrW$(231) & ChrW$(227) & "o terminar."
        Exit Function
    End If
    If esquema <> modConfig.VERSAO_ESQUEMA Then
        MotivoRecusaEscrita = "O banco est" & ChrW$(225) & " no esquema " & esquema & " e esta c" & ChrW$(243) & _
            "pia do CRM espera o " & modConfig.VERSAO_ESQUEMA & ": nada foi gravado. Feche e abra pelo INICIAR-CRM."
    End If
End Function

Public Sub Confirmar(ByVal cn As Object)
    On Error GoTo falhou
    cn.CommitTrans
    cn.Close
    Exit Sub
falhou:
    Dim nErr As Long, sErr As String, sOrig As String
    nErr = Err.Number: sErr = Err.Description: sOrig = Err.Source
    modDiag.Registrar "gravacao", "commit", "", nErr, sErr
    Err.Raise nErr, sOrig, sErr
End Sub

' "UPDATE clientes SET ..." -> "UPDATE clientes" (diagnostico sem dado)
Public Function OperacaoDoSql(ByVal sql As String) As String
    Dim palavras As Variant, s As String
    s = Trim$(Replace(Replace(sql, vbCr, " "), vbLf, " "))
    palavras = Split(s, " ")
    If UBound(palavras) < 0 Then Exit Function
    Select Case UCase$(palavras(0))
        Case "UPDATE": If UBound(palavras) >= 1 Then OperacaoDoSql = "UPDATE " & palavras(1)
        Case "INSERT", "DELETE": If UBound(palavras) >= 2 Then OperacaoDoSql = UCase$(palavras(0)) & " " & palavras(2)
        Case Else: OperacaoDoSql = UCase$(palavras(0))
    End Select
End Function

Public Sub Desfazer(ByVal cn As Object)
    On Error Resume Next
    If cn Is Nothing Then Exit Sub
    If cn.State = 1 Then
        cn.RollbackTrans
        cn.Close
    End If
    On Error GoTo 0
End Sub

Public Function ExecutarEm(ByVal cn As Object, ByVal sql As String, _
                           Optional ByVal params As Variant) As Long
    Dim cmd As Object, n As Long, msg As String
    On Error GoTo falhou
    Set cmd = NovoComando(cn, sql, params)
    cmd.Execute n
    ExecutarEm = n
    Exit Function
falhou:
    Dim nErr As Long, sErr As String, sOrig As String
    nErr = Err.Number: sErr = Err.Description: sOrig = Err.Source
    ' DEV-03: com o backup segurando o arquivo, o ACE abriu so leitura e a
    ' gravacao falhou antes de gravar: quem pediu recebe a mensagem controlada
    msg = MensagemEscritaRecusada(sErr, modBackupAviso.BackupEmCurso())
    ' DEV-06A: so o verbo e a tabela do SQL (os valores sao parametros, nunca entram no texto)
    modDiag.Registrar "gravacao", OperacaoDoSql(sql), "", nErr, sErr, IIf(msg <> "", "recusada (backup)", "erro")
    If msg <> "" Then Err.Raise ERRO_ESCRITA_RECUSADA, "modDB.ExecutarEm", msg
    Err.Raise nErr, sOrig, sErr
End Function

'----------------------------------------------------------
' DEV-03: erro de escrita durante a copia exclusiva do backup
' vira a mensagem controlada ("" = nao e o caso). Mesma regra de
' Get-MensagemEscritaRecusada (conector, Banco.ps1).
'----------------------------------------------------------
Public Function MensagemEscritaRecusada(ByVal erro As String, ByVal backupEmCurso As Boolean) As String
    Dim s As String
    If Not backupEmCurso Then Exit Function
    s = LCase$(erro)
    If InStr(s, "atualiz") = 0 And InStr(s, "updat") = 0 And InStr(s, "bloquead") = 0 And InStr(s, "locked") = 0 _
       And InStr(s, "em uso") = 0 And InStr(s, "in use") = 0 And InStr(s, "somente leitura") = 0 And InStr(s, "read-only") = 0 Then Exit Function
    MensagemEscritaRecusada = "C" & ChrW$(243) & "pia de seguran" & ChrW$(231) & "a (backup) do banco em andamento: grava" & _
        ChrW$(231) & ChrW$(227) & "o recusada, nada foi gravado. A leitura continua normal; tente gravar de novo em um ou dois minutos."
End Function

Public Function ConsultarEm(ByVal cn As Object, ByVal sql As String, _
                            Optional ByVal params As Variant) As Variant
    Dim cmd As Object, rs As Object
    Set cmd = NovoComando(cn, sql, params)
    Set rs = cmd.Execute
    ConsultarEm = LerMatriz(rs)
    rs.Close
End Function

Public Function ConsultarValorEm(ByVal cn As Object, ByVal sql As String, _
                                 Optional ByVal padrao As Variant = Null, _
                                 Optional ByVal params As Variant) As Variant
    Dim d As Variant
    d = ConsultarEm(cn, sql, params)
    If IsEmpty(d) Then
        ConsultarValorEm = padrao
    ElseIf IsNull(d(1, 1)) Then
        ConsultarValorEm = padrao
    Else
        ConsultarValorEm = d(1, 1)
    End If
End Function

' INSERT na transacao; devolve o id novo (@@IDENTITY da MESMA conexao).
Public Function InserirEm(ByVal cn As Object, ByVal sql As String, _
                          Optional ByVal params As Variant) As Long
    Dim rs As Object
    ExecutarEm cn, sql, params
    Set rs = cn.Execute("SELECT @@IDENTITY")
    InserirEm = CLng(rs.Fields(0).Value)
    rs.Close
End Function

'----------------------------------------------------------
' Uma linha de auditoria, na transacao do dado. E tambem o
' feed que faz as outras estacoes atualizarem a grade.
'----------------------------------------------------------
Public Sub LogEm(ByVal cn As Object, ByVal tabela As String, ByVal idRegistro As Long, _
                 ByVal campo As String, ByVal valorAntigo As String, ByVal valorNovo As String, _
                 ByVal acao As String)
    ExecutarEm cn, "INSERT INTO log_alteracoes (tabela, id_registro, campo, valor_antigo," & _
                   " valor_novo, acao, usuario, quando, origem) VALUES (?,?,?,?,?,?,?,?,?)", _
        Array(P(adVarWChar, Left$(tabela, 30)), P(adInteger, idRegistro), P(adVarWChar, Left$(campo, 40)), _
              P(adLongVarWChar, valorAntigo), P(adLongVarWChar, valorNovo), P(adVarWChar, Left$(acao, 20)), _
              P(adVarWChar, modConfig.UsuarioAtual()), P(adDate, Now), P(adVarWChar, "FRONT"))
End Sub

'----------------------------------------------------------
' Colisao em indice unico (dois usuarios pegaram o mesmo
' MAX+1). Quem chama repete com o proximo numero. O ACE
' responde "valores duplicados" / "duplicate values".
'----------------------------------------------------------
Public Function EhDuplicidade(ByVal descricao As String) As Boolean
    Dim s As String
    s = LCase$(descricao)
    EhDuplicidade = (InStr(s, "duplica") > 0 Or InStr(s, "duplicate") > 0)
End Function

'----------------------------------------------------------
' Texto de valor para o log: data dd/mm/aaaa, nulo vazio.
'----------------------------------------------------------
Public Function TextoLog(ByVal v As Variant) As String
    If IsNull(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbDate Then
        If v = Int(v) Then TextoLog = Format$(v, "dd/mm/yyyy") Else TextoLog = Format$(v, "dd/mm/yyyy hh:nn")
    Else
        TextoLog = CStr(v)
    End If
End Function

'----------------------------------------------------------
Private Sub LimparRecursos(ByVal rs As Object, ByVal cn As Object)
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State = 1 Then rs.Close
    If Not cn Is Nothing Then If cn.State = 1 Then cn.Close
    On Error GoTo 0
End Sub

Private Function Transpor(ByVal m As Variant) As Variant
    Dim i As Long, j As Long, r As Variant
    ReDim r(1 To UBound(m, 2) + 1, 1 To UBound(m, 1) + 1)
    For i = 0 To UBound(m, 2)
        For j = 0 To UBound(m, 1)
            r(i + 1, j + 1) = m(j, i)
        Next j
    Next i
    Transpor = r
End Function

Public Function IndiceCampo(ByVal nome As String) As Long
    Dim i As Long
    If IsEmpty(gCampos) Then Exit Function
    For i = LBound(gCampos) To UBound(gCampos)
        If LCase$(gCampos(i)) = LCase$(nome) Then IndiceCampo = i: Exit Function
    Next i
End Function

Public Function Nz(ByVal v As Variant, Optional ByVal padrao As String = "") As String
    ' NZ() do Access NAO existe via OLE DB - tratamento e sempre aqui.
    If IsNull(v) Or IsEmpty(v) Then Nz = padrao Else Nz = CStr(v)
End Function

'----------------------------------------------------------
' Marcadores "?,?,?" e parametros para um filtro IN de ids.
' Ids vem do proprio sistema, mas seguem parametrizados como
' todo o resto. Mais de 200 ids: quem chama recarrega tudo.
'----------------------------------------------------------
Public Function MarcadoresIds(ByVal ids As Variant, ByRef params As Variant) As String
    Dim i As Long, n As Long, lista() As Variant, m As String
    If Not IsArray(ids) Then Exit Function
    ReDim lista(0 To UBound(ids) - LBound(ids))
    For i = LBound(ids) To UBound(ids)
        m = m & IIf(m = "", "", ",") & "?"
        lista(n) = P(adInteger, CLng(ids(i)))
        n = n + 1
    Next i
    params = lista
    MarcadoresIds = m
End Function

'----------------------------------------------------------
Public Function BancoOK(Optional ByVal silencioso As Boolean = False) As Boolean
    Dim cn As Object
    On Error GoTo falhou
    If Not modConfig.CaminhoConfigurado() Then GoTo falhou
    If Dir(modConfig.CaminhoBanco()) = "" Then GoTo falhou
    Set cn = NovaConexao()
    cn.Close
    Set cn = Nothing
    BancoOK = True
    Exit Function
falhou:
    BancoOK = False
    If Not silencioso Then
        MsgBox "Nao foi possivel acessar o banco de dados." & vbCrLf & vbCrLf & _
               "Caminho: " & IIf(modConfig.CaminhoConfigurado(), modConfig.CaminhoBanco(), "(nao configurado: Config!B2 esta vazio; use a copia publicada pelo MONTAR-FRONTEND)") & vbCrLf & vbCrLf & _
               "Ultima consulta bem-sucedida: " & modConfig.UltimaConsulta() & vbCrLf & vbCrLf & _
               "Verifique a rede e o caminho na aba Config.", vbExclamation, "CRM Zapromaq"
    End If
End Function
