Attribute VB_Name = "modRelatorio"
'==========================================================
' modRelatorio - retrato do momento para a direcao.
'
' Um clique (botao "Exportar relatorio" na aba Painel) gera
' uma pasta com data e hora em Documentos\CRM Zapromaq\Relatorios:
'
'   indicadores.csv    os numeros do Painel (bloco;chave;valores)
'   movimento_mensal.csv  entradas, fechados e perdidos, 12 meses
'   atendimentos.csv   todos os atendimentos, com a situacao
'                      calculada e os dias parado
'   clientes.csv       empresas (clientes e pre-clientes)
'   painel.pdf         o Painel como esta na tela, para imprimir
'   LEIA-ME.txt        o que e cada arquivo e um pedido pronto
'                      para montar o relatorio com o Claude
'
' Formato pensado para ser lido por maquina (Claude) sem
' ambiguidade: UTF-8, separador ";", data AAAA-MM-DD, ponto
' decimal, sem separador de milhar. Nada e gravado no banco.
'==========================================================
Option Explicit

Private Const SEP As String = ";"

' Botao da aba Painel
Public Sub ExportarRelatorio()
    GerarRelatorio False
End Sub

'----------------------------------------------------------
' Gera o relatorio e devolve a pasta ("" se nao gerou).
' silencioso = True: sem caixas de mensagem (teste automatico).
'----------------------------------------------------------
Public Function GerarRelatorio(Optional ByVal silencioso As Boolean = False) As String
    Dim pasta As String, cn As Object, d As Variant, sql As String, params As Variant
    Dim nAt As Long, nCl As Long, inicio As Single

    If Not modMenu.Pronto() Then Exit Function
    On Error GoTo falha
    inicio = Timer
    Application.Cursor = 2                        ' xlWait
    Application.StatusBar = "CRM Zapromaq: gerando o relat" & ChrW$(&HF3) & "rio..."

    ' 1. Painel atualizado agora: o PDF e o indicadores.csv sao do mesmo instante
    modPainel.AtualizarBasePainel

    pasta = PastaNova()

    ' 2. indicadores = a aba BasePainel (mesmos numeros do Painel)
    GravarTexto pasta & "indicadores.csv", CsvDaBasePainel()

    ' 3. movimento mensal (mesma regra do Painel)
    GravarTexto pasta & "movimento_mensal.csv", CsvMensal()

    ' 4. tabelas completas, numa conexao
    Set cn = modDB.AbrirLeitura()
    sql = modCRM.SQLGrade("oportunidades", params)
    d = modDB.ConsultarEm(cn, sql, params)
    nAt = Linhas(d)
    GravarTexto pasta & "atendimentos.csv", CsvAtendimentos(d, modDB.gCampos)
    sql = modCRM.SQLGrade("clientes", params)
    d = modDB.ConsultarEm(cn, sql, params)
    nCl = Linhas(d)
    GravarTexto pasta & "clientes.csv", CsvTabela(d, modDB.gCampos)
    modDB.FecharLeitura cn: Set cn = Nothing

    ' 5. Painel em PDF (paisagem, uma pagina de largura)
    PainelEmPdf pasta & "painel.pdf"

    ' 6. instrucoes
    GravarTexto pasta & "LEIA-ME.txt", TextoLeiaMe(nAt, nCl)

    Application.Cursor = -4143
    Application.StatusBar = False
    GerarRelatorio = pasta
    If silencioso Then Exit Function
    If MsgBox("Relat" & ChrW$(&HF3) & "rio gerado em " & Format$(Timer - inicio, "0.0") & " s:" & vbCrLf & vbCrLf & _
              pasta & vbCrLf & vbCrLf & _
              "indicadores, movimento mensal, " & nAt & " atendimento(s), " & nCl & " empresa(s)," & vbCrLf & _
              "o Painel em PDF e o LEIA-ME com o pedido pronto para o Claude." & vbCrLf & vbCrLf & _
              "Abrir a pasta?", vbYesNo + vbInformation, "CRM Zapromaq") = vbYes Then
        Shell "explorer.exe """ & pasta & """", vbNormalFocus
    End If
    Exit Function
falha:
    Dim sErr As String
    sErr = Err.Description
    modDB.FecharLeitura cn
    Application.Cursor = -4143
    Application.StatusBar = False
    If silencioso Then Err.Raise vbObjectError + 600, "modRelatorio", sErr
    MsgBox "N" & ChrW$(&HE3) & "o foi poss" & ChrW$(&HED) & "vel gerar o relat" & ChrW$(&HF3) & "rio." & vbCrLf & vbCrLf & sErr, _
           vbExclamation, "CRM Zapromaq"
End Function

'----------------------------------------------------------
' Documentos\CRM Zapromaq\Relatorios\AAAA-MM-DD_HHMM\
' (Documentos pelo registro: com OneDrive, e redirecionado)
'----------------------------------------------------------
Private Function PastaNova() As String
    Dim sh As Object, docs As String, p As String
    On Error Resume Next
    Set sh = CreateObject("WScript.Shell")
    docs = sh.SpecialFolders("MyDocuments")
    On Error GoTo 0
    If docs = "" Then docs = Environ$("USERPROFILE") & "\Documents"
    p = docs & "\CRM Zapromaq"
    If Dir(p, vbDirectory) = "" Then MkDir p
    p = p & "\Relatorios"
    If Dir(p, vbDirectory) = "" Then MkDir p
    p = p & "\" & Format$(Now, "yyyy-mm-dd_hhnn")
    If Dir(p, vbDirectory) = "" Then MkDir p
    PastaNova = p & "\"
End Function

' UTF-8 com BOM (acentos certos no Excel e no Claude)
Private Sub GravarTexto(ByVal arquivo As String, ByVal texto As String)
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2                                   ' adTypeText
    st.Charset = "utf-8"
    st.Open
    st.WriteText texto
    st.SaveToFile arquivo, 2                      ' adSaveCreateOverWrite
    st.Close
End Sub

'----------------------------------------------------------
' Valor -> campo CSV: data AAAA-MM-DD, numero com ponto,
' texto entre aspas quando precisa.
'----------------------------------------------------------
Public Function CampoCsv(ByVal v As Variant) As String
    Dim s As String
    If IsNull(v) Or IsEmpty(v) Then Exit Function
    Select Case VarType(v)
        Case vbDate
            If v = Int(v) Then s = Format$(v, "yyyy-mm-dd") Else s = Format$(v, "yyyy-mm-dd hh:nn")
        Case vbBoolean
            s = IIf(v, "sim", "nao")
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte
            s = Trim$(Str$(v))                      ' Str usa sempre ponto decimal
            If Left$(s, 1) = "." Then s = "0" & s
            If Left$(s, 2) = "-." Then s = "-0" & Mid$(s, 2)
        Case Else
            s = Replace(Replace(CStr(v), vbCrLf, " "), vbLf, " ")
            s = Replace(s, vbCr, " ")
            If InStr(s, SEP) > 0 Or InStr(s, """") > 0 Then s = """" & Replace(s, """", """""") & """"
    End Select
    CampoCsv = s
End Function

Private Function Linhas(ByVal d As Variant) As Long
    If Not IsEmpty(d) Then Linhas = UBound(d, 1)
End Function

' Matriz de consulta -> CSV (sem colunas internas de controle)
Private Function CsvTabela(ByVal d As Variant, ByVal nomes As Variant) As String
    Dim i As Long, j As Long, linha As String, r As String, usar() As Boolean
    ReDim usar(LBound(nomes) To UBound(nomes))
    For j = LBound(nomes) To UBound(nomes)
        usar(j) = (LCase$(nomes(j)) <> "versao")
        If usar(j) Then linha = linha & IIf(linha = "", "", SEP) & nomes(j)
    Next j
    r = linha & vbCrLf
    If Not IsEmpty(d) Then
        For i = 1 To UBound(d, 1)
            linha = ""
            For j = LBound(nomes) To UBound(nomes)
                If usar(j) Then linha = linha & IIf(j = LBound(nomes), "", SEP) & CampoCsv(d(i, j))
            Next j
            r = r & linha & vbCrLf
        Next i
    End If
    CsvTabela = r
End Function

' Atendimentos + colunas calculadas (as mesmas da grade e do Painel)
Private Function CsvAtendimentos(ByVal d As Variant, ByVal nomes As Variant) As String
    Dim base As String, linhas As Variant, i As Long, idx As Object, extra As String, r As String
    Set idx = CreateObject("Scripting.Dictionary")
    idx.CompareMode = 1
    For i = LBound(nomes) To UBound(nomes)
        If Not idx.Exists(nomes(i)) Then idx.Add nomes(i), i
    Next i
    base = CsvTabela(d, nomes)
    linhas = Split(base, vbCrLf)
    r = linhas(0) & SEP & "situacao" & SEP & "quadro" & SEP & "dias_parado" & SEP & "em_aberto" & vbCrLf
    If Not IsEmpty(d) Then
        For i = 1 To UBound(d, 1)
            extra = SEP & CampoCsv(modCalc.Situacao(True, modDB.Nz(d(i, idx("etapa"))), d(i, idx("dt_prox_acao")), d(i, idx("retomar_em")))) & _
                    SEP & CampoCsv(modCalc.Quadro(modDB.Nz(d(i, idx("etapa"))))) & _
                    SEP & CampoCsv(modCalc.DiasParado(d(i, idx("ultima_interacao")))) & _
                    SEP & IIf(EmAberto(modDB.Nz(d(i, idx("etapa")))), "sim", "nao")
            r = r & linhas(i) & extra & vbCrLf
        Next i
    End If
    CsvAtendimentos = r
End Function

Private Function EmAberto(ByVal etapa As String) As Boolean
    EmAberto = Not (etapa = "Pedido Fechado" Or etapa = "Perdido" Or etapa = "Descartado")
End Function

' A aba BasePainel ja e bloco;chave;valor1..3 - com cabecalho legivel
Private Function CsvDaBasePainel() As String
    Dim ws As Worksheet, v As Variant, i As Long, j As Long, r As String, linha As String
    Set ws = ThisWorkbook.Worksheets("BasePainel")
    v = ws.UsedRange.Value2
    If Not IsArray(v) Then Exit Function
    r = "bloco;chave;valor1;valor2;valor3" & vbCrLf
    For i = 2 To UBound(v, 1)
        If CStr(v(i, 1)) <> "" And CStr(v(i, 1)) <> "ATUALIZADO_EM" Then
            linha = ""
            For j = 1 To 5
                If j <= UBound(v, 2) Then linha = linha & IIf(j = 1, "", SEP) & CampoCsv(v(i, j)) Else linha = linha & SEP
            Next j
            r = r & linha & vbCrLf
        End If
    Next i
    CsvDaBasePainel = r
End Function

Private Function CsvMensal() As String
    Dim ent As Variant, des As Variant, m As Variant, i As Long, r As String, rotulo As String
    ent = modDB.Consultar("SELECT YEAR(op.dt_entrada)*100+MONTH(op.dt_entrada) AS mes, COUNT(*) AS qt" & _
                          " FROM oportunidades op WHERE op.dt_entrada IS NOT NULL" & _
                          " GROUP BY YEAR(op.dt_entrada)*100+MONTH(op.dt_entrada)")
    des = modDB.Consultar("SELECT YEAR(op.dt_desfecho)*100+MONTH(op.dt_desfecho) AS mes, op.etapa," & _
                          " COUNT(*) AS qt, SUM(op.valor) AS total FROM oportunidades op" & _
                          " WHERE op.dt_desfecho IS NOT NULL" & _
                          " GROUP BY YEAR(op.dt_desfecho)*100+MONTH(op.dt_desfecho), op.etapa")
    m = modPainel.Mensal(ent, des)
    r = "mes;entradas;fechados;valor_fechado;perdidos_ou_descartados" & vbCrLf
    For i = 1 To 12
        ' "09/2026" -> "2026-09"
        rotulo = Right$(CStr(m(i, 1)), 4) & "-" & Left$(CStr(m(i, 1)), 2)
        r = r & rotulo & SEP & CampoCsv(m(i, 2)) & SEP & CampoCsv(m(i, 3)) & SEP & CampoCsv(m(i, 4)) & SEP & CampoCsv(m(i, 5)) & vbCrLf
    Next i
    CsvMensal = r
End Function

Private Sub PainelEmPdf(ByVal arquivo As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Painel")
    On Error Resume Next
    With ws.PageSetup
        .Orientation = 2                          ' xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .CenterFooter = "CRM Zapromaq - Painel comercial - " & Format$(Now, "dd/mm/yyyy hh:nn")
    End With
    On Error GoTo 0
    ws.ExportAsFixedFormat 0, arquivo             ' xlTypePDF
End Sub

Private Function TextoLeiaMe(ByVal nAt As Long, ByVal nCl As Long) As String
    Dim t As String
    t = "RELATORIO DO CRM ZAPROMAQ - retrato de " & Format$(Now, "dd/mm/yyyy hh:nn") & vbCrLf & _
        "Gerado por " & modConfig.UsuarioAtual() & " (versao do CRM " & modConfig.VERSAO_FRONT & ")" & vbCrLf & vbCrLf & _
        "ARQUIVOS (CSV: UTF-8, separador ponto e virgula, data AAAA-MM-DD, ponto decimal)" & vbCrLf & _
        "  indicadores.csv        numeros do Painel: bloco;chave;valor1;valor2;valor3" & vbCrLf & _
        "                         blocos: total (totais gerais), situacao (acompanhamento), etapa (qtd;valor)," & vbCrLf & _
        "                         responsavel/familia/categoria/segmento (so em aberto), origem, mes_entrada," & vbCrLf & _
        "                         mes_desfecho (mes;etapa;qtd;valor), fila (estagio;aderencia;qtd), metas" & vbCrLf & _
        "  movimento_mensal.csv   ultimos 12 meses: entradas, fechados, valor fechado, perdidos/descartados" & vbCrLf & _
        "  atendimentos.csv       " & nAt & " atendimento(s), todas as colunas + situacao, quadro, dias_parado, em_aberto" & vbCrLf & _
        "  clientes.csv           " & nCl & " empresa(s): clientes e pre-clientes" & vbCrLf & _
        "  painel.pdf             o Painel como estava na tela" & vbCrLf & vbCrLf & _
        "DEFINICOES" & vbCrLf & _
        "  Em aberto = etapa diferente de Pedido Fechado, Perdido e Descartado." & vbCrLf & _
        "  Situacao: Acao atrasada (proxima acao vencida), Retomar hoje, Acao nesta semana, Em dia," & vbCrLf & _
        "  Sem proxima acao, Encerrado. Valores em reais (R$)." & vbCrLf & vbCrLf
    t = t & "PEDIDO SUGERIDO PARA O CLAUDE (anexe os CSVs e, se quiser, o painel.pdf)" & vbCrLf & _
        "  Monte um relatorio executivo de 1 a 2 paginas para a direcao da Zapromaq com base nos" & vbCrLf & _
        "  arquivos anexos do CRM (retrato de " & Format$(Now, "dd/mm/yyyy") & "). Inclua: resumo do funil em aberto" & vbCrLf & _
        "  (quantidade e valor por etapa), pedidos fechados e perdidos no ultimo mes e a tendencia dos 12 meses," & vbCrLf & _
        "  desempenho por responsavel, atendimentos com acao atrasada ou sem proxima acao (os de maior valor)," & vbCrLf & _
        "  principais familias de maquina e origens, e 3 a 5 pontos de atencao com recomendacao." & vbCrLf & _
        "  Use tabelas curtas, numeros em reais no formato brasileiro e nao invente dados que nao estejam nos arquivos." & vbCrLf
    TextoLeiaMe = t
End Function
