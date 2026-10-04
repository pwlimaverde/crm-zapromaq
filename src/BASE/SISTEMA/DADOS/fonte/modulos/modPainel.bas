Attribute VB_Name = "modPainel"
'==========================================================
' modPainel - indicadores por agregacao no banco
'
' A rotina publica chama-se AtualizarBasePainel, e nao
' AtualizarIndicadores: esse nome pertence ao modMenu, que e o
' ponto de entrada do botao. Dois Public com o mesmo nome fazem
' o VBA recusar a chamada com "Nome repetido encontrado".
'
' Nao existe mais aba espelho. As consultas GROUP BY (algumas
' dezenas de linhas, nao 5.000) vao para a aba oculta
' BasePainel (conferencia: testes\conferir-painel.ps1) e o
' modulo DESENHA a aba Painel com elas (DesenharPainel, abaixo).
' Medido em 16/09/2026: GROUP BY por etapa sobre 500 registros
' em 46 ms; por responsavel com filtro, 4 ms.
'
' Se a atualizacao falhar, o conjunto anterior PERMANECE e e
' marcado como desatualizado. Numero velho nunca e mostrado
' como atual.
'==========================================================
Option Explicit

Private Const ABA_BASE As String = "BasePainel"
Private Const ABA_PAINEL As String = "Painel"
Private Const MINUTOS_VALIDADE As Long = 15

Private mAtualizadoEm As Date            ' ultima atualizacao nesta sessao
Private mTempos As String                ' "etapa=ms; ..." da ultima atualizacao
Private mMarca As Single

' desenho da aba Painel: grade de 3 paineis de 4 colunas (B:E, G:J, L:O)
Private Const LIN_CARTOES As Long = 5
Private Const P1 As Long = 2
Private Const P2 As Long = 7
Private Const P3 As Long = 12
Private Const LINHAS_GRAFICO_MIN As Long = 13
Private Const ABERTAS As String = " AND op.etapa NOT IN ('Pedido Fechado','Perdido','Descartado')"

Public Sub AtualizarBasePainel()
    Dim ws As Worksheet, inicio As Single, linha As Long, cn As Object
    Dim calcAnterior As XlCalculation

    If Not modDB.BancoOK() Then Exit Sub
    inicio = Timer
    mTempos = "": mMarca = Timer

    On Error GoTo trata
    Set ws = ThisWorkbook.Worksheets(ABA_BASE)

    calcAnterior = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' monta tudo em memoria antes de tocar na aba:
    ' se a consulta falhar, o painel anterior fica intacto.
    ' UMA conexao para as ~20 consultas (antes, uma por consulta).
    Set cn = modDB.AbrirLeitura()
    Dim blocos As Object
    Set blocos = CreateObject("Scripting.Dictionary")
    blocos.Add "etapa", modDB.ConsultarEm(cn,  _
        "SELECT op.etapa, COUNT(*) AS qt, SUM(op.valor) AS total" & _
        " FROM oportunidades op GROUP BY op.etapa ORDER BY op.etapa")
    blocos.Add "responsavel", modDB.ConsultarEm(cn,  _
        "SELECT op.responsavel, COUNT(*) AS qt, SUM(op.valor) AS total" & _
        " FROM oportunidades op WHERE 1=1" & ABERTAS & " GROUP BY op.responsavel")
    blocos.Add "familia", modDB.ConsultarEm(cn,  _
        "SELECT op.familia, COUNT(*) AS qt, SUM(op.valor) AS total" & _
        " FROM oportunidades op WHERE 1=1" & ABERTAS & " GROUP BY op.familia")
    blocos.Add "categoria", modDB.ConsultarEm(cn,  _
        "SELECT op.categoria, COUNT(*) AS qt, SUM(op.valor) AS total" & _
        " FROM oportunidades op WHERE 1=1" & ABERTAS & " GROUP BY op.categoria")
    blocos.Add "origem", modDB.ConsultarEm(cn,  _
        "SELECT op.origem, COUNT(*) AS qt FROM oportunidades op GROUP BY op.origem")
    blocos.Add "segmento", modDB.ConsultarEm(cn,  _
        "SELECT cl.segmento, COUNT(*) AS qt FROM oportunidades op" & _
        " LEFT JOIN clientes cl ON op.id_cliente=cl.id WHERE 1=1" & ABERTAS & _
        " GROUP BY cl.segmento")
    blocos.Add "mes_entrada", modDB.ConsultarEm(cn,  _
        "SELECT YEAR(op.dt_entrada)*100+MONTH(op.dt_entrada) AS mes, COUNT(*) AS qt" & _
        " FROM oportunidades op WHERE op.dt_entrada IS NOT NULL" & _
        " GROUP BY YEAR(op.dt_entrada)*100+MONTH(op.dt_entrada)")
    blocos.Add "mes_desfecho", modDB.ConsultarEm(cn,  _
        "SELECT YEAR(op.dt_desfecho)*100+MONTH(op.dt_desfecho) AS mes, op.etapa," & _
        " COUNT(*) AS qt, SUM(op.valor) AS total FROM oportunidades op" & _
        " WHERE op.dt_desfecho IS NOT NULL" & _
        " GROUP BY YEAR(op.dt_desfecho)*100+MONTH(op.dt_desfecho), op.etapa")
    blocos.Add "fila", modDB.ConsultarEm(cn,  _
        "SELECT cl.estagio, cl.aderencia, COUNT(*) AS qt FROM clientes cl" & _
        " GROUP BY cl.estagio, cl.aderencia")
    blocos.Add "ordem_etapa", modDB.ConsultarEm(cn,  _
        "SELECT valor, ordem FROM listas WHERE tipo='Etapa' ORDER BY ordem, valor")
    blocos.Add "metas", modDB.ConsultarEm(cn,  _
        "SELECT ano, mes, meta_faturamento, meta_pedidos, meta_propostas, meta_contatos" & _
        " FROM metas ORDER BY ano, mes")

    ' Totais: uma consulta escalar por vez. Juntar tudo num SELECT de
    ' subconsultas sem FROM e pedir para o ACE recusar, e cada escalar
    ' custa poucos milissegundos - nao ha o que economizar aqui.
    Dim totais As Object
    Set totais = CreateObject("Scripting.Dictionary")
    totais("empresas") = modDB.ConsultarValorEm(cn, "SELECT COUNT(*) FROM clientes", 0)
    totais("clientes") = modDB.ConsultarValorEm(cn, "SELECT COUNT(*) FROM clientes WHERE estagio='Cliente'", 0)
    totais("preclientes") = modDB.ConsultarValorEm(cn, "SELECT COUNT(*) FROM clientes WHERE estagio='Pré-cliente'", 0)
    totais("contatos") = modDB.ConsultarValorEm(cn, "SELECT COUNT(*) FROM contatos", 0)
    totais("atendimentos") = modDB.ConsultarValorEm(cn, "SELECT COUNT(*) FROM oportunidades", 0)
    totais("abertos") = modDB.ConsultarValorEm(cn, "SELECT COUNT(*) FROM oportunidades op WHERE 1=1" & ABERTAS, 0)
    totais("valor_aberto") = modDB.ConsultarValorEm(cn, "SELECT SUM(op.valor) FROM oportunidades op WHERE 1=1" & ABERTAS, 0)
    totais("valor_ganho") = modDB.ConsultarValorEm(cn, "SELECT SUM(valor) FROM oportunidades WHERE etapa='Pedido Fechado'", 0)

    ' Situacao de acompanhamento (atrasada, nesta semana, em dia...):
    ' a regra depende do dia de hoje e esta em modCalc, nao em SQL.
    ' Uma consulta so, contada em memoria.
    Dim situacoes As Object
    Set situacoes = ContarPorSituacao(cn)
    modDB.FecharLeitura cn: Set cn = Nothing
    Marcar "consultas"


    ' consulta deu certo: agora sim reescreve a aba
    ws.Cells.ClearContents
    linha = 1
    ws.Cells(linha, 1).Value = "BLOCO": ws.Cells(linha, 2).Value = "CHAVE"
    ws.Cells(linha, 3).Value = "VALOR1": ws.Cells(linha, 4).Value = "VALOR2"
    ws.Cells(linha, 5).Value = "VALOR3"
    linha = 2

    linha = EscreverTotais(ws, linha, totais)
    linha = EscreverContagem(ws, linha, "situacao", situacoes)
    Dim nome As Variant
    For Each nome In blocos.Keys
        linha = EscreverBloco(ws, linha, CStr(nome), blocos(nome))
    Next nome

    ws.Cells(linha + 1, 1).Value = "ATUALIZADO_EM"
    ws.Cells(linha + 1, 2).Value = Now
    ws.Cells(linha + 1, 3).Value = "OK"

    Marcar "BasePainel"
    DesenharPainel totais, situacoes, blocos
    mAtualizadoEm = Now

    Application.Calculation = calcAnterior
    Application.Calculate
    Application.ScreenUpdating = True
    Application.StatusBar = "Indicadores atualizados em " & Format$(Timer - inicio, "0.0") & _
                            " s  -  " & Format$(Now, "dd/mm/yyyy hh:nn")
    Exit Sub

trata:
    ' captura ANTES de qualquer On Error: MarcarDesatualizado usa
    ' On Error GoTo 0, e isso limpa o objeto Err
    Dim nErro As Long, sErro As String, sOrigem As String
    nErro = Err.Number
    sErro = Err.Description
    sOrigem = Err.Source
    modDB.FecharLeitura cn

    Application.Calculation = calcAnterior
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MarcarDesatualizado

    MsgBox "Não foi possível atualizar os indicadores." & vbCrLf & vbCrLf & _
           "Erro " & nErro & IIf(sOrigem = "", "", " em " & sOrigem) & vbCrLf & _
           sErro & vbCrLf & vbCrLf & _
           "Os números na tela são os da última atualização e estão marcados como desatualizados.", _
           vbExclamation, "CRM Zapromaq"
End Sub

'----------------------------------------------------------
' Ao abrir a aba Painel (ThisWorkbook.Workbook_SheetActivate):
' atualiza sozinho se os numeros tem mais de 15 minutos nesta
' sessao - o Painel nunca abre vazio nem velho sem aviso.
'----------------------------------------------------------
Public Sub AtualizarSeVelho()
    If modConfig.gModoModelo Then Exit Sub
    If mAtualizadoEm <> 0 Then
        If DateDiff("n", mAtualizadoEm, Now) < MINUTOS_VALIDADE Then Exit Sub
    End If
    AtualizarBasePainel
End Sub

'----------------------------------------------------------
Private Function EscreverTotais(ByVal ws As Worksheet, ByVal linha As Long, _
                                ByVal d As Object) As Long
    Dim chave As Variant
    If d Is Nothing Then EscreverTotais = linha: Exit Function
    For Each chave In d.Keys
        ws.Cells(linha, 1).Value = "total"
        ws.Cells(linha, 2).Value = chave
        ws.Cells(linha, 3).Value = TrocaNulo(d(chave))
        linha = linha + 1
    Next chave
    EscreverTotais = linha
End Function

Private Function EscreverContagem(ByVal ws As Worksheet, ByVal linha As Long, _
                                  ByVal nome As String, ByVal d As Object) As Long
    Dim chave As Variant
    If d Is Nothing Then EscreverContagem = linha: Exit Function
    For Each chave In d.Keys
        ws.Cells(linha, 1).Value = nome
        ws.Cells(linha, 2).Value = chave
        ws.Cells(linha, 3).Value = d(chave)
        linha = linha + 1
    Next chave
    EscreverContagem = linha
End Function

Private Function EscreverBloco(ByVal ws As Worksheet, ByVal linha As Long, _
                               ByVal nome As String, ByVal d As Variant) As Long
    Dim i As Long, j As Long, colunas As Long
    If IsEmpty(d) Then EscreverBloco = linha: Exit Function
    colunas = UBound(d, 2)
    If colunas > 4 Then colunas = 4
    For i = 1 To UBound(d, 1)
        ws.Cells(linha, 1).Value = nome
        For j = 1 To colunas
            ws.Cells(linha, 1 + j).Value = TrocaNulo(d(i, j))
        Next j
        linha = linha + 1
    Next i
    EscreverBloco = linha
End Function

Private Function TrocaNulo(ByVal v As Variant) As Variant
    If IsNull(v) Then TrocaNulo = "" Else TrocaNulo = v
End Function

Private Sub MarcarDesatualizado()
    Dim ws As Worksheet, ultima As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(ABA_BASE)
    If ws Is Nothing Then Exit Sub
    ultima = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ws.Cells(ultima, 3).Value = "DESATUALIZADO desde " & Format$(Now, "dd/mm/yyyy hh:nn")
    On Error GoTo 0
End Sub

'----------------------------------------------------------
' Situacao de acompanhamento: calculada aqui, porque a regra
' de "esta semana" depende do dia e nao cabe bem em SQL.
'----------------------------------------------------------
Public Function ContarPorSituacao(ByVal cn As Object) As Object
    Dim d As Variant, i As Long, s As String, r As Object
    Set r = CreateObject("Scripting.Dictionary")
    d = modDB.ConsultarEm(cn, "SELECT etapa, dt_prox_acao, retomar_em FROM oportunidades")
    If IsEmpty(d) Then Set ContarPorSituacao = r: Exit Function
    For i = 1 To UBound(d, 1)
        s = modCalc.Situacao(True, modDB.Nz(d(i, 1)), d(i, 2), d(i, 3))
        If s <> "" Then
            If r.Exists(s) Then r(s) = r(s) + 1 Else r.Add s, 1
        End If
    Next i
    Set ContarPorSituacao = r
End Function

'==========================================================
' DESENHO DA ABA PAINEL
'
' Grade fixa de 14 colunas iguais (B:O), em tres paineis de 4
' colunas com uma de respiro entre eles:
'     B:E  |  G:J  |  L:O
' Cartoes: 6 de 2 colunas. Tabelas: rotulo ocupa 2 colunas,
' numeros 1. Graficos ocupam o espaco de paineis inteiros, ao
' lado da tabela que representam. Valores (nao formulas) e
' graficos nativos do Excel 2019.
'
'   cartoes de totais
'   funil por etapa ............................ + grafico
'   movimento dos ultimos 12 meses ............. + grafico
'   acompanhamento | por responsavel | por familia
'   origem         | segmento        | estagio x aderencia
'
' Os botoes da aba (build) ficam nas linhas 1-3; aqui so se
' limpa da linha 4 para baixo.
'==========================================================

Private Sub DesenharPainel(ByVal totais As Object, ByVal situacoes As Object, ByVal blocos As Object)
    Dim ws As Worksheet, ch As ChartObject, lin As Long, fim As Long, f2 As Long, f3 As Long
    Dim ini As Long, d As Variant

    Set ws = ThisWorkbook.Worksheets(ABA_PAINEL)
    ws.Unprotect
    For Each ch In ws.ChartObjects
        ch.Delete
    Next ch
    With ws.Range(ws.Rows(LIN_CARTOES - 1), ws.Rows(500))
        .UnMerge
        .Clear
        .RowHeight = ws.StandardHeight
    End With
    ws.Cells.Font.Name = modTema.FONTE_PADRAO
    ws.Cells.Interior.Color = modTema.COR_FUNDO
    ActiveWindowSemGrade ws
    Colunas ws
    Marcar "limpeza"

    ws.Range("B2").Value = "Painel comercial"
    ws.Range("B2").Font.Size = 16: ws.Range("B2").Font.Bold = True
    ws.Range("B2").Font.Color = modTema.COR_AZUL
    ws.Range("B3").Value = "Atualizado em " & Format$(Now, "dd/mm/yyyy hh:nn") & _
                           "   " & ChrW$(&HB7) & "   em aberto = etapa diferente de Pedido Fechado, Perdido e Descartado"
    ws.Range("B3").Font.Size = 9: ws.Range("B3").Font.Color = modTema.COR_TEXTO2

    ' ---------------- cartoes (6 x 2 colunas)
    Cartao ws, 2, "Empresas", totais("empresas"), "#,##0"
    Cartao ws, 4, "Clientes", totais("clientes"), "#,##0"
    Cartao ws, 7, "Pr" & ChrW$(&HE9) & "-clientes", totais("preclientes"), "#,##0"
    Cartao ws, 9, "Atendimentos em aberto", totais("abertos"), "#,##0"
    Cartao ws, 12, "Valor em aberto (R$)", totais("valor_aberto"), "#,##0"
    Cartao ws, 14, "Pedidos fechados (R$)", totais("valor_ganho"), "#,##0"

    ' ---------------- funil + grafico
    lin = LIN_CARTOES + 4
    d = Funil(blocos("etapa"), blocos("ordem_etapa"))
    fim = Tabela(ws, lin, P1, "Funil por etapa (todos os atendimentos)", _
                 Array("Etapa", "Qtd", "Valor (R$)"), Array(2, 1, 1), d)
    If fim > lin + 1 Then
        Grafico ws, lin, P2, 15, fim, 57, "Atendimentos por etapa", _
                ws.Range(ws.Cells(lin + 2, P1), ws.Cells(fim, P1)), Array("Atendimentos"), _
                Array(ws.Range(ws.Cells(lin + 2, P1 + 2), ws.Cells(fim, P1 + 2)))
    End If
    lin = FimComGrafico(fim, lin) + 2

    ' ---------------- movimento mensal + grafico
    ' tabela: Mes | Entradas | Fechados | Valor fechado (4 colunas do painel B:E);
    ' os perdidos entram no grafico, ao lado, como terceira serie
    d = Mensal(blocos("mes_entrada"), blocos("mes_desfecho"))
    fim = Tabela(ws, lin, P1, "Movimento dos " & ChrW$(&HFA) & "ltimos 12 meses", _
                 Array("M" & ChrW$(&HEA) & "s", "Entradas", "Fechados", "Valor (R$)"), _
                 Array(1, 1, 1, 1), d)
    Grafico ws, lin, P2, 15, fim, 51, "Entradas, pedidos fechados e perdidos por m" & ChrW$(&HEA) & "s", _
            ws.Range(ws.Cells(lin + 2, P1), ws.Cells(fim, P1)), Array("Entradas", "Fechados", "Perdidos"), _
            Array(ws.Range(ws.Cells(lin + 2, P1 + 1), ws.Cells(fim, P1 + 1)), _
                  ws.Range(ws.Cells(lin + 2, P1 + 2), ws.Cells(fim, P1 + 2)), _
                  ColunaMensal(d, 5))
    lin = FimComGrafico(fim, lin) + 2
    Marcar "cartoes, funil e meses"

    ' ---------------- acompanhamento | responsavel | familia
    fim = Tabela(ws, lin, P1, "Acompanhamento", Array("Situa" & ChrW$(&HE7) & ChrW$(&HE3) & "o", "Qtd", "%"), _
                 Array(2, 1, 1), ComPercentual(DeDicionario(situacoes)))
    PintarSituacoes ws, lin + 2, fim, P1
    f2 = Tabela(ws, lin, P2, "Em aberto por respons" & ChrW$(&HE1) & "vel", _
                Array("Respons" & ChrW$(&HE1) & "vel", "Qtd", "Valor (R$)"), Array(2, 1, 1), Ordenado(blocos("responsavel"), 3))
    f3 = Tabela(ws, lin, P3, "Em aberto por fam" & ChrW$(&HED) & "lia", _
                Array("Fam" & ChrW$(&HED) & "lia", "Qtd", "Valor (R$)"), Array(2, 1, 1), Ordenado(blocos("familia"), 3))
    lin = Maior(Maior(fim, f2), f3) + 2

    ' ---------------- origem | segmento | estagio x aderencia
    fim = Tabela(ws, lin, P1, "Origem (todos)", Array("Origem", "Qtd", "%"), Array(2, 1, 1), _
                 ComPercentual(Ordenado(blocos("origem"), 2)))
    f2 = Tabela(ws, lin, P2, "Em aberto por segmento", Array("Segmento", "Qtd", "%"), Array(2, 1, 1), _
                ComPercentual(Ordenado(blocos("segmento"), 2)))
    f3 = Tabela(ws, lin, P3, "Empresas por est" & ChrW$(&HE1) & "gio e ader" & ChrW$(&HEA) & "ncia", _
                Array("Est" & ChrW$(&HE1) & "gio", "Ader" & ChrW$(&HEA) & "ncia", "Qtd"), Array(2, 1, 1), _
                SemNotaComoTraco(blocos("fila")))

    If ActiveSheet Is ws Then ws.Range("A1").Select
    ws.Protect DrawingObjects:=False, Contents:=True, UserInterfaceOnly:=True
    Marcar "paineis de baixo"
End Sub

' Tempo por etapa da ultima atualizacao (diagnostico: testes\testar-uso.ps1).
Public Function TemposUltimaAtualizacao() As String
    TemposUltimaAtualizacao = mTempos
End Function

Private Sub Marcar(ByVal etapa As String)
    mTempos = mTempos & IIf(mTempos = "", "", "; ") & etapa & "=" & Format$((Timer - mMarca) * 1000, "0") & " ms"
    mMarca = Timer
End Sub

Private Sub ActiveWindowSemGrade(ByVal ws As Worksheet)
    On Error Resume Next
    If ActiveSheet Is ws Then ActiveWindow.DisplayGridlines = False
    On Error GoTo 0
End Sub

' 14 colunas iguais (B:O); F e K sao o respiro entre os paineis.
Private Sub Colunas(ByVal ws As Worksheet)
    Dim c As Long
    ws.Columns(1).ColumnWidth = 2
    For c = 2 To 15
        ws.Columns(c).ColumnWidth = 11
    Next c
    ws.Columns(6).ColumnWidth = 2
    ws.Columns(11).ColumnWidth = 2
    ws.Columns(16).ColumnWidth = 2
End Sub

Private Function Maior(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then Maior = a Else Maior = b
End Function

' Linha onde a secao termina: a tabela ou o grafico, o que for mais alto.
Private Function FimComGrafico(ByVal fimTabela As Long, ByVal linTopo As Long) As Long
    FimComGrafico = Maior(fimTabela, linTopo + LINHAS_GRAFICO_MIN)
End Function

'----------------------------------------------------------
' Cartao de 2 colunas: rotulo pequeno em cima, numero grande
' embaixo, fundo branco com borda suave.
'----------------------------------------------------------
Private Sub Cartao(ByVal ws As Worksheet, ByVal col As Long, ByVal rotulo As String, _
                   ByVal valor As Variant, ByVal formato As String)
    Dim r As Range
    Set r = ws.Range(ws.Cells(LIN_CARTOES, col), ws.Cells(LIN_CARTOES + 1, col + 1))
    r.Interior.Color = modTema.COR_BRANCO
    ws.Range(ws.Cells(LIN_CARTOES, col), ws.Cells(LIN_CARTOES, col + 1)).Merge
    ws.Range(ws.Cells(LIN_CARTOES + 1, col), ws.Cells(LIN_CARTOES + 1, col + 1)).Merge
    r.BorderAround LineStyle:=1, Weight:=2, Color:=modTema.COR_BORDA_SUAVE
    With ws.Cells(LIN_CARTOES, col)
        .Value = rotulo
        .Font.Size = 8: .Font.Color = modTema.COR_TEXTO2
        .IndentLevel = 1
    End With
    With ws.Cells(LIN_CARTOES + 1, col)
        If IsNull(valor) Or IsEmpty(valor) Then valor = 0
        .Value = valor
        .NumberFormat = formato
        .HorizontalAlignment = xlLeft
        .IndentLevel = 1
        .Font.Size = 16: .Font.Bold = True: .Font.Color = modTema.COR_AZUL
    End With
    ws.Rows(LIN_CARTOES).RowHeight = 16
    ws.Rows(LIN_CARTOES + 1).RowHeight = 28
End Sub

'----------------------------------------------------------
' Tabela com titulo, cabecalho azul e linhas zebradas.
' spans: quantas colunas da grade cada coluna ocupa (mescla).
' dados: matriz base 1 (linha, coluna) ou Empty. Devolve a
' ultima linha usada.
'----------------------------------------------------------
Private Function Tabela(ByVal ws As Worksheet, ByVal lin As Long, ByVal col As Long, ByVal titulo As String, _
                        ByVal cab As Variant, ByVal spans As Variant, ByVal dados As Variant) As Long
    Dim nc As Long, i As Long, j As Long, n As Long, c As Long, largura As Long, cab1 As String
    Dim inicio() As Long, fimCol As Long

    nc = UBound(cab) - LBound(cab) + 1
    ReDim inicio(1 To nc)
    c = col
    For j = 1 To nc
        inicio(j) = c
        c = c + spans(LBound(spans) + j - 1)
    Next j
    fimCol = c - 1

    ws.Cells(lin, col).Value = titulo
    ws.Cells(lin, col).Font.Bold = True: ws.Cells(lin, col).Font.Size = 10
    ws.Cells(lin, col).Font.Color = modTema.COR_AZUL

    If IsEmpty(dados) Then n = 0 Else n = UBound(dados, 1)
    ' mescla de cada celula que ocupa mais de uma coluna (cabecalho e dados)
    For j = 1 To nc
        largura = spans(LBound(spans) + j - 1)
        If largura > 1 Then
            For i = 0 To IIf(n = 0, 1, n)
                ws.Range(ws.Cells(lin + 1 + i, inicio(j)), ws.Cells(lin + 1 + i, inicio(j) + largura - 1)).Merge
            Next i
        End If
        ws.Cells(lin + 1, inicio(j)).Value = cab(LBound(cab) + j - 1)
    Next j
    With ws.Range(ws.Cells(lin + 1, col), ws.Cells(lin + 1, fimCol))
        .Interior.Color = modTema.COR_AZUL: .Font.Color = modTema.COR_BRANCO: .Font.Bold = True: .Font.Size = 9
    End With
    For j = 2 To nc
        ws.Cells(lin + 1, inicio(j)).HorizontalAlignment = xlRight
    Next j

    If n = 0 Then
        ws.Cells(lin + 2, col).Value = "(sem dados)"
        ws.Cells(lin + 2, col).Font.Color = modTema.COR_TEXTO2
        ws.Cells(lin + 2, col).Font.Italic = True
        Tabela = lin + 2
        Exit Function
    End If

    ' 1a coluna e rotulo: texto sempre ("10/2025" viraria data)
    ws.Range(ws.Cells(lin + 2, col), ws.Cells(lin + 1 + n, col)).NumberFormat = "@"
    For i = 1 To n
        For j = 1 To nc
            If j <= UBound(dados, 2) Then
                If j = 1 Then
                    ws.Cells(lin + 1 + i, inicio(j)).Value = Limpo(dados(i, j))
                Else
                    ws.Cells(lin + 1 + i, inicio(j)).Value = Numero(dados(i, j))
                End If
            End If
        Next j
        ws.Range(ws.Cells(lin + 1 + i, col), ws.Cells(lin + 1 + i, fimCol)).Interior.Color = _
            IIf(i Mod 2 = 0, modTema.COR_ZEBRA, modTema.COR_BRANCO)
    Next i
    With ws.Range(ws.Cells(lin + 2, col), ws.Cells(lin + 1 + n, fimCol))
        .Font.Size = 9: .Font.Color = modTema.COR_TEXTO
    End With
    For j = 2 To nc
        cab1 = CStr(cab(LBound(cab) + j - 1))
        With ws.Range(ws.Cells(lin + 2, inicio(j)), ws.Cells(lin + 1 + n, inicio(j)))
            If cab1 = "%" Then .NumberFormat = "0%" Else .NumberFormat = "#,##0"
            .HorizontalAlignment = xlRight
        End With
    Next j
    ws.Range(ws.Cells(lin + 1, col), ws.Cells(lin + 1 + n, fimCol)).BorderAround LineStyle:=1, Weight:=2, Color:=modTema.COR_BORDA_SUAVE
    Tabela = lin + 1 + n
End Function

' Coluna numerica: nulo e 0; texto (ex.: "-") passa como veio.
Private Function Numero(ByVal v As Variant) As Variant
    If IsNull(v) Or IsEmpty(v) Then Numero = 0 Else Numero = v
End Function

Private Function Limpo(ByVal v As Variant) As Variant
    If IsNull(v) Or IsEmpty(v) Then
        Limpo = "(n" & ChrW$(&HE3) & "o informado)"
    ElseIf VarType(v) = vbString And Trim$(v) = "" Then
        Limpo = "(n" & ChrW$(&HE3) & "o informado)"
    Else
        Limpo = v
    End If
End Function

Private Sub PintarSituacoes(ByVal ws As Worksheet, ByVal ini As Long, ByVal fim As Long, ByVal col As Long)
    Dim i As Long, s As String
    For i = ini To fim
        s = CStr(ws.Cells(i, col).Value)
        If s <> "" And s <> "(sem dados)" Then
            ws.Cells(i, col).Font.Color = modTema.CorSituacao(s, "texto")
            ws.Cells(i, col).Font.Bold = True
        End If
    Next i
End Sub

'----------------------------------------------------------
' Grafico ocupando as colunas colIni..colFim, da linha do
' titulo da tabela ate o fim dela (minimo LINHAS_GRAFICO_MIN).
' Series montadas uma a uma (nome, valores, categorias).
'----------------------------------------------------------
Private Sub Grafico(ByVal ws As Worksheet, ByVal linTopo As Long, ByVal colIni As Long, ByVal colFim As Long, _
                    ByVal linFim As Long, ByVal tipo As Long, ByVal titulo As String, _
                    ByVal categorias As Range, ByVal nomes As Variant, ByVal valores As Variant)
    Dim co As ChartObject, s As Object, k As Long, esq As Double, topo As Double, cores As Variant, maximo As Double
    If linFim < linTopo + LINHAS_GRAFICO_MIN Then linFim = linTopo + LINHAS_GRAFICO_MIN
    esq = ws.Cells(linTopo, colIni).Left
    topo = ws.Cells(linTopo, colIni).Top
    Set co = ws.ChartObjects.Add(esq, topo, ws.Cells(linTopo, colFim + 1).Left - esq, _
                                 ws.Cells(linFim + 1, colIni).Top - topo)
    cores = Array(modTema.COR_AZUL_MEDIO, modTema.COR_VERDE_ESCURO, modTema.COR_PERIGO)
    With co.Chart
        .ChartType = tipo
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
        For k = LBound(valores) To UBound(valores)
            Set s = .SeriesCollection.NewSeries
            s.Name = nomes(k)
            s.Values = valores(k)
            s.XValues = categorias
            s.Format.Fill.ForeColor.RGB = cores((k - LBound(valores)) Mod 3)
            If Application.WorksheetFunction.Max(valores(k)) > maximo Then maximo = Application.WorksheetFunction.Max(valores(k))
        Next k
        .HasTitle = True
        .ChartTitle.Text = titulo
        .ChartTitle.Font.Size = 10
        .ChartTitle.Font.Color = modTema.COR_AZUL
        .HasLegend = (UBound(valores) > LBound(valores))
        If .HasLegend Then .Legend.Position = xlLegendPositionBottom
        If tipo = 57 Then .Axes(xlCategory).ReversePlotOrder = True     ' funil: 1a etapa em cima
        ' quantidade e inteira: sem passo fracionario (0 0 0 1 1 1)
        .Axes(xlValue).MinimumScale = 0
        .Axes(xlValue).MajorUnit = PassoEixo(maximo)
        .Axes(xlValue).TickLabels.NumberFormat = "0"
        .ChartArea.Format.Line.ForeColor.RGB = modTema.COR_BORDA_SUAVE
    End With
    co.Placement = xlFreeFloating
End Sub

Private Function PassoEixo(ByVal m As Double) As Double
    If m <= 10 Then
        PassoEixo = 1
    ElseIf m <= 50 Then
        PassoEixo = 5
    Else
        PassoEixo = Application.WorksheetFunction.Ceiling(m / 10, 10)
    End If
End Function

'----------------------------------------------------------
' Dados de cada bloco
'----------------------------------------------------------
Private Function DeDicionario(ByVal d As Object) As Variant
    Dim r() As Variant, k As Variant, i As Long
    If d Is Nothing Then Exit Function
    If d.Count = 0 Then Exit Function
    ReDim r(1 To d.Count, 1 To 2)
    For Each k In d.Keys
        i = i + 1: r(i, 1) = k: r(i, 2) = d(k)
    Next k
    DeDicionario = Ordenado(r, 2)
End Function

' (rotulo, qtd) -> (rotulo, qtd, fracao do total)
Private Function ComPercentual(ByVal d As Variant) As Variant
    Dim r() As Variant, i As Long, total As Double
    If IsEmpty(d) Then Exit Function
    For i = 1 To UBound(d, 1)
        total = total + Num(d(i, 2))
    Next i
    ReDim r(1 To UBound(d, 1), 1 To 3)
    For i = 1 To UBound(d, 1)
        r(i, 1) = d(i, 1): r(i, 2) = Num(d(i, 2))
        If total > 0 Then r(i, 3) = Num(d(i, 2)) / total Else r(i, 3) = 0
    Next i
    ComPercentual = r
End Function

' Aderencia em branco aparece como "-" (e nao como 0, que seria nota)
Private Function SemNotaComoTraco(ByVal d As Variant) As Variant
    Dim i As Long
    If IsEmpty(d) Then Exit Function
    For i = 1 To UBound(d, 1)
        If IsNull(d(i, 2)) Then d(i, 2) = "-"
    Next i
    SemNotaComoTraco = d
End Function

' Ordena por uma coluna numerica, maior primeiro (poucas linhas: insercao).
Private Function Ordenado(ByVal d As Variant, ByVal col As Long) As Variant
    Dim i As Long, j As Long, k As Long, t As Variant
    If IsEmpty(d) Then Exit Function
    For i = LBound(d, 1) + 1 To UBound(d, 1)
        j = i
        Do While j > LBound(d, 1)
            If Num(d(j, col)) <= Num(d(j - 1, col)) Then Exit Do
            For k = LBound(d, 2) To UBound(d, 2)
                t = d(j, k): d(j, k) = d(j - 1, k): d(j - 1, k) = t
            Next k
            j = j - 1
        Loop
    Next i
    Ordenado = d
End Function

Private Function Num(ByVal v As Variant) As Double
    If IsNull(v) Or IsEmpty(v) Then Exit Function
    If IsNumeric(v) Then Num = CDbl(v)
End Function

' Etapas na ordem da lista (a ordem do funil), depois as que nao estao nela.
Private Function Funil(ByVal d As Variant, ByVal ordem As Variant) As Variant
    Dim r() As Variant, n As Long, i As Long, j As Long, usado() As Boolean
    If IsEmpty(d) Then Exit Function
    ReDim r(1 To UBound(d, 1), 1 To 3)
    ReDim usado(1 To UBound(d, 1))
    If Not IsEmpty(ordem) Then
        For i = 1 To UBound(ordem, 1)
            For j = 1 To UBound(d, 1)
                If Not usado(j) And modDB.Nz(d(j, 1)) = modDB.Nz(ordem(i, 1)) Then
                    n = n + 1: r(n, 1) = d(j, 1): r(n, 2) = d(j, 2): r(n, 3) = d(j, 3): usado(j) = True
                End If
            Next j
        Next i
    End If
    For j = 1 To UBound(d, 1)
        If Not usado(j) Then n = n + 1: r(n, 1) = d(j, 1): r(n, 2) = d(j, 2): r(n, 3) = d(j, 3)
    Next j
    Funil = r
End Function

Private Function ChaveMes(ByVal i As Long) As Long
    ' i = 0 e o mes atual; 11 e onze meses atras
    Dim dt As Date
    dt = DateAdd("m", -i, DateSerial(Year(Date), Month(Date), 1))
    ChaveMes = Year(dt) * 100 + Month(dt)
End Function

Private Function RotuloMes(ByVal chave As Long) As String
    RotuloMes = Format$(chave Mod 100, "00") & "/" & (chave \ 100)
End Function

' Uma coluna da matriz mensal como vetor (serie de grafico sem celula)
Private Function ColunaMensal(ByVal m As Variant, ByVal col As Long) As Variant
    Dim r(1 To 12) As Double, i As Long
    For i = 1 To 12
        r(i) = Num(m(i, col))
    Next i
    ColunaMensal = r
End Function

' Mes | Entradas | Fechados | Valor fechado | Perdidos (12 meses)
Public Function Mensal(ByVal entradas As Variant, ByVal desfechos As Variant) As Variant
    Dim r(1 To 12, 1 To 5) As Variant, i As Long, j As Long, c As Long, etapa As String
    For i = 1 To 12
        c = ChaveMes(12 - i)
        r(i, 1) = RotuloMes(c): r(i, 2) = 0: r(i, 3) = 0: r(i, 4) = 0: r(i, 5) = 0
        If Not IsEmpty(entradas) Then
            For j = 1 To UBound(entradas, 1)
                If CLng(Num(entradas(j, 1))) = c Then r(i, 2) = Num(entradas(j, 2))
            Next j
        End If
        If Not IsEmpty(desfechos) Then
            For j = 1 To UBound(desfechos, 1)
                If CLng(Num(desfechos(j, 1))) = c Then
                    etapa = modDB.Nz(desfechos(j, 2))
                    If etapa = "Pedido Fechado" Then
                        r(i, 3) = r(i, 3) + Num(desfechos(j, 3)): r(i, 4) = r(i, 4) + Num(desfechos(j, 4))
                    ElseIf etapa = "Perdido" Or etapa = "Descartado" Then
                        r(i, 5) = r(i, 5) + Num(desfechos(j, 3))
                    End If
                End If
            Next j
        End If
    Next i
    Mensal = r
End Function
