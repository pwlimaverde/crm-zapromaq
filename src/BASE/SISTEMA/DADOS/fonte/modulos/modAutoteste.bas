Attribute VB_Name = "modAutoteste"
'==========================================================
' modAutoteste - conferencia automatica do front, sem banco.
'
' O build (montar-frontend.ps1) chama Autoteste depois de
' compilar e so publica se voltar "OK". Chamar uma rotina de
' cada modulo tambem obriga o VBA a compilar o modulo inteiro:
' erro de sintaxe que a verificacao estatica nao pegou aparece
' aqui, na maquina do build, e nao na mao do vendedor.
'
' Pode ser rodado na estacao a qualquer momento:
'   Alt+F11 > Ctrl+G > ?Autoteste("")
'
' Regras: nada aqui acessa o banco nem grava na planilha.
' Acentos por ChrW para a fonte seguir em ASCII.
'==========================================================
Option Explicit

Private mFalhas As String
Private mTestes As Long

Public Function Autoteste(Optional ByVal versaoEsperada As String = "") As String
    mFalhas = ""
    mTestes = 0
    On Error GoTo quebrou

    ' ---- modConfig
    If versaoEsperada <> "" Then
        Confere "VERSAO_FRONT", modConfig.VERSAO_FRONT, versaoEsperada
    End If
    Confere "VERSAO_ESQUEMA", modConfig.VERSAO_ESQUEMA, "1.0"

    ' ---- modDB (so utilitarios)
    Confere "Nz nulo", modDB.Nz(Null, "x"), "x"
    Confere "Nz valor", modDB.Nz(12), "12"

    ' ---- modValidacao
    Confere "Normalizar empresa", modValidacao.Normalizar("empresa", "  ind  d'angelo  "), "IND D'ANGELO"
    Confere "Normalizar email", modValidacao.Normalizar("email", " A@B.COM "), "a@b.com"
    Confere "Normalizar cnpj", modValidacao.Normalizar("cnpj", "11.222.333/0001-81"), "11222333000181"
    Confere "Normalizar vazio vira nulo", IsNull(modValidacao.Normalizar("cidade", "   ")), True
    Confere "CNPJ valido", modValidacao.CNPJValido("11222333000181"), True
    Confere "CNPJ invalido", modValidacao.CNPJValido("11222333000182"), False
    Confere "CNPJ repetido", modValidacao.CNPJValido("11111111111111"), False
    Confere "FormatarCNPJ", modValidacao.FormatarCNPJ("11222333000181"), "11.222.333/0001-81"
    Confere "FormatarTelefone 11", modValidacao.FormatarTelefone("41999990000"), "(41) 99999-0000"
    Confere "FormatarTelefone 10", modValidacao.FormatarTelefone("4133330000"), "(41) 3333-0000"
    Confere "Telefone curto recusado", (modValidacao.CriticarTelefone("123") <> ""), True
    Confere "Telefone certo aceito", modValidacao.CriticarTelefone("41999990000"), ""

    ' ---- modCalc (regras da planilha 3.6)
    Confere "Situacao encerrado", modCalc.Situacao(True, "Perdido", Null, Null), "Encerrado"
    Confere "Situacao sem etapa", modCalc.Situacao(True, "", Null, Null), "Sem etapa definida"
    Confere "Situacao atrasada", modCalc.Situacao(True, "Contato Inicial", Date - 1, Null), _
            "A" & ChrW$(&HE7) & ChrW$(&HE3) & "o atrasada"
    Confere "Situacao retomar", modCalc.Situacao(True, "Contato Inicial", Null, Date), "Retomar hoje"
    Confere "Situacao sem acao", modCalc.Situacao(True, "Contato Inicial", Null, Null), _
            "Sem pr" & ChrW$(&HF3) & "xima a" & ChrW$(&HE7) & ChrW$(&HE3) & "o"
    Confere "Situacao em dia", modCalc.Situacao(True, "Contato Inicial", Date + 30, Null), "Em dia"
    Confere "Quadro prospeccao", modCalc.Quadro("Contato Inicial"), "1. Prospec" & ChrW$(&HE7) & ChrW$(&HE3) & "o"
    Confere "Quadro funil", modCalc.Quadro("Pedido Fechado"), "2. Funil comercial"
    Confere "Dias parado", modCalc.DiasParado(Date - 10), 10
    Confere "Ciclo", modCalc.Ciclo(DateSerial(2026, 1, 1), DateSerial(2026, 1, 31)), 30
    Confere "Mes referencia", modCalc.MesReferencia(DateSerial(2026, 9, 19)), 202609

    ' ---- modCRM (sem banco)
    Confere "MontarCodigo", modCRM.MontarCodigo("AT", 11, 148), "AT-0011-0148"
    Confere "CodigoVisualCliente", modCRM.CodigoVisualCliente(7), "0007"
    Confere "CodigoVisualCliente nulo", modCRM.CodigoVisualCliente(Null), ""
    ' item 6: o atendimento so troca de contato dentro da mesma empresa
    ' (o codigo AT- e congelado e nomeia a pasta do cliente)
    Confere "troca de contato na mesma empresa", modCRM.CriticarTrocaContato(3, 3), ""
    Confere "troca de contato para outra empresa", (modCRM.CriticarTrocaContato(3, 4) <> ""), True
    Confere "atendimento novo nao confere empresa", modCRM.CriticarTrocaContato(0, 4), ""

    ' ---- modSchema
    Confere "CamposDe clientes", (UBound(modSchema.CamposDe("clientes")) >= 10), True
    Confere "CamposDe oportunidades", (UBound(modSchema.CamposDe("oportunidades")) >= 20), True
    Confere "ColunasPlanilha", (UBound(modSchema.ColunasPlanilha("oportunidades", False)) >= 5), True
    Confere "Parte", modSchema.Parte("a|b|c", 1), "b"
    ' a ficha guarda um vinculo POR CAMPO K: com dois campos K na mesma
    ' tabela, um id unico gravaria o contato no lugar do outro vinculo
    Confere "CamposK oportunidades", modSchema.CamposK("oportunidades"), "id_contato"
    Confere "CamposK contatos", modSchema.CamposK("contatos"), "id_cliente"
    Confere "CamposK clientes", modSchema.CamposK("clientes"), ""

    ' ---- modGrade / modAcoes / modLote
    Confere "TabelaDaAba", modGrade.TabelaDaAba(ThisWorkbook.Worksheets("Clientes")), "clientes"
    Confere "TabelaDaAba fora", modGrade.TabelaDaAba(ThisWorkbook.Worksheets("Inicio")), ""
    Confere "RotuloStatus", modAcoes.RotuloStatus("contatos", 0, True), "Desativar"
    Confere "NomeNormalizado", modLote.NomeNormalizado("Ind" & ChrW$(&HFA) & "stria D'" & ChrW$(&HC2) & "ngelo LTDA"), "D ANGELO"

    ' ---- tema e tela (visual)
    Confere "COR_AZUL do tema", modTema.COR_AZUL, 6760983
    Confere "CorBotao primario", modTema.CorBotao("primario", "fundo"), modTema.COR_AZUL
    Confere "CorBotao sem borda", modTema.CorBotao("primario", "borda"), -1
    ' AscW devolve Integer COM sinal: acima de &H7FFF vem negativo
    Confere "Glifo", AscW(modTema.Glifo(modTema.ICO_NOVO)) And &HFFFF&, &HE710&
    Confere "COR_BRANCO", modTema.COR_BRANCO, 16777215
    Confere "Zoom cabe", modTela.ZoomIdeal(100, 100), 100
    Confere "DPI positiva", (modTela.DPI() > 0), True
    Confere "feed desligado no build", modFeed.Ligado(), False
    Confere "versao 1.10 > 1.9", (modConfig.NumeroVersao("1.10") > modConfig.NumeroVersao("1.9")), True
    Confere "LIKE literal", modCRM.PadraoLike("A_B[1]%"), "%A[_]B[[]1][%]%"
    Confere "resumo de contato sem banco", modCRM.TextoResumoContato("ANA", "COMPRAS", "EMPRESA X", "JAU", "SP", "Cliente"), _
            "ANA (COMPRAS)  -  EMPRESA X  JAU/SP  [Cliente]"
    modFeed.Verificar                          ' desligado: sai sem consultar nem agendar
    modFeed.Parar

    ' ---- formularios: instanciar compila o modulo do formulario
    Dim f As Object
    Set f = New frmCRM: Unload f: Set f = Nothing
    Set f = New frmLote: Unload f: Set f = Nothing
    Set f = New frmCalibracao: Unload f: Set f = Nothing
    Set f = New frmVinculo: Unload f: Set f = Nothing
    ContratoFormularios

    If mFalhas = "" Then Autoteste = "OK " & mTestes & " testes" Else Autoteste = "FALHA" & mFalhas
    Exit Function
quebrou:
    Autoteste = "FALHA" & mFalhas & vbLf & "erro " & Err.Number & " durante o autoteste: " & Err.Description
End Function

Private Sub Confere(ByVal nome As String, ByVal obtido As Variant, ByVal esperado As Variant)
    mTestes = mTestes + 1
    If IsNull(obtido) Or IsNull(esperado) Then
        If IsNull(obtido) And IsNull(esperado) Then Exit Sub
    ElseIf VarType(obtido) = VarType(esperado) Or (IsNumeric(obtido) And IsNumeric(esperado)) Then
        If obtido = esperado Then Exit Sub
    ElseIf Texto(obtido) = Texto(esperado) Then
        Exit Sub
    End If
    mFalhas = mFalhas & vbLf & nome & ": esperado [" & Texto(esperado) & "], obtido [" & Texto(obtido) & "]"
End Sub

'----------------------------------------------------------
' CONTRATO DOS FORMULARIOS
' Os formularios sao modais: chamar o metodo mostraria a tela e
' travaria o build. Com a variavel TIPADA, um metodo que nao
' existe e erro de compilacao - o build para antes de publicar.
' O If False garante que nada roda; o comportamento se confere
' na estacao (itens marcados na spec).
'----------------------------------------------------------
Private Sub ContratoFormularios()
    Dim fv As frmVinculo
    If False Then
        Set fv = New frmVinculo
        Call fv.EscolherContatoDe(0)
    End If
End Sub

' CStr(Null) e erro: o relato da falha nao pode quebrar o autoteste.
Private Function Texto(ByVal v As Variant) As String
    If IsNull(v) Then Texto = "(nulo)" Else Texto = CStr(v)
End Function
