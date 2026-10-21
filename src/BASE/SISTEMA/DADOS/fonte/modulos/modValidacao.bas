Attribute VB_Name = "modValidacao"
'==========================================================
' modValidacao - normalizacao e regras de preenchimento
'
' Ponto UNICO de normalizacao: o mesmo valor entra igual,
' venha do formulario ou do executor da fila.
' Padrao definido na secao 5.1 do estudo de migracao.
'==========================================================
Option Explicit
'----------------------------------------------------------
' Normaliza pelo NOME do campo (seção 5.1 do plano de
' 19/10/2026, ciclo 009 - a mesma regra de ConvertTo-Normalizado
' no conector, casos C22 a C24):
'   - todo texto comercial em MAIÚSCULAS, com Trim e espaço
'     horizontal repetido colapsado; acento preservado;
'   - observações, próxima ação e contexto: as quebras de linha
'     ficam; URL, caminho e e-mail dentro do texto ficam como
'     digitados;
'   - e-mail minúsculo, um endereço por campo (os demais saem
'     por EmailsAdicionais e vão para as observações);
'   - CNPJ e telefone só dígitos;
'   - campo de lista: o valor exato da lista (só apara);
'   - ctx_arquivo: caminho, como gravado.
'----------------------------------------------------------
Public Function Normalizar(ByVal campo As String, ByVal valor As Variant) As Variant
    Dim s As String
    If IsNull(valor) Then Normalizar = Null: Exit Function
    s = Replace(CStr(valor), Chr$(160), " ")

    Select Case LCase$(campo)
        Case "email"
            s = LCase$(PrimeiroEmail(s))
        Case "cnpj", "telefone"
            s = SoDigitos(s)
        Case "etapa", "responsavel", "familia", "origem", "categoria", "prioridade", _
             "motivo_desfecho", "qualificacao", "segmento"
            s = ColapsarEspacos(s)
        Case "ctx_arquivo"
            s = Trim$(s)
        Case "observacoes", "prox_acao", "ctx_resumo"
            s = TextoComercial(s, True)
        Case Else
            s = TextoComercial(s, False)
    End Select

    If Len(s) = 0 Then Normalizar = Null Else Normalizar = s
End Function

' Texto comercial: maiúsculas por palavra, menos URL, caminho e e-mail.
Public Function TextoComercial(ByVal s As String, ByVal multilinha As Boolean) As String
    Dim linhas As Variant, i As Long, r As String, ln As String
    s = Replace(Replace(s, vbCrLf, vbLf), vbCr, vbLf)
    linhas = Split(s, vbLf)
    For i = LBound(linhas) To UBound(linhas)
        ln = MaiusculasProtegidas(ColapsarEspacos(CStr(linhas(i))))
        If multilinha Then
            r = r & IIf(i > LBound(linhas), vbCrLf, "") & ln
        ElseIf ln <> "" Then
            r = r & IIf(r = "", "", " ") & ln
        End If
    Next i
    If multilinha Then
        ' linhas vazias nas pontas saem; as do meio ficam
        Do While Left$(r, 2) = vbCrLf
            r = Mid$(r, 3)
        Loop
        Do While Right$(r, 2) = vbCrLf
            r = Left$(r, Len(r) - 2)
        Loop
    End If
    TextoComercial = r
End Function

' Espaço e tabulação repetidos viram um espaço; aparado nas pontas.
Public Function ColapsarEspacos(ByVal s As String) As String
    s = Trim$(Replace(s, vbTab, " "))
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    ColapsarEspacos = s
End Function

Private Function MaiusculasProtegidas(ByVal s As String) As String
    Dim p As Variant, i As Long, t As String
    If s = "" Then Exit Function
    p = Split(s, " ")
    For i = LBound(p) To UBound(p)
        t = CStr(p(i))
        If Not (InStr(t, "@") > 0 Or InStr(t, "://") > 0 Or InStr(t, "\") > 0 Or LCase$(Left$(t, 4)) = "www.") Then t = UCase$(t)
        p(i) = t
    Next i
    MaiusculasProtegidas = Join(p, " ")
End Function

' "a@x.com; b@y.com" -> "a@x.com" (o primeiro endereço)
Public Function PrimeiroEmail(ByVal s As String) As String
    Dim p As Variant
    p = ListaEmails(s)
    If UBound(p) >= 0 Then PrimeiroEmail = CStr(p(0))
End Function

' Os demais endereços, separados por "; " ("" = só havia um)
Public Function EmailsAdicionais(ByVal s As String) As String
    Dim p As Variant, i As Long, r As String
    p = ListaEmails(s)
    For i = 1 To UBound(p)
        r = r & IIf(r = "", "", "; ") & LCase$(CStr(p(i)))
    Next i
    EmailsAdicionais = r
End Function

Private Function ListaEmails(ByVal s As String) As Variant
    Dim p As Variant, i As Long, r As String
    s = Replace(Replace(Replace(s, ";", " "), ",", " "), "/", " ")
    p = Split(ColapsarEspacos(s), " ")
    For i = LBound(p) To UBound(p)
        If Trim$(CStr(p(i))) <> "" Then r = r & IIf(r = "", "", "|") & Trim$(CStr(p(i)))
    Next i
    If r = "" Then ListaEmails = Split("", "|") Else ListaEmails = Split(r, "|")
End Function

' Endereços além do primeiro vão para as observações, numa linha própria.
Public Function ObservacoesComEmails(ByVal observacoes As Variant, ByVal adicionais As String) As Variant
    If adicionais = "" Then ObservacoesComEmails = observacoes: Exit Function
    If IsNull(observacoes) Or IsEmpty(observacoes) Then
        ObservacoesComEmails = "E-MAIL ADICIONAL: " & adicionais
    ElseIf InStr(CStr(observacoes), adicionais) > 0 Then
        ObservacoesComEmails = observacoes
    Else
        ObservacoesComEmails = CStr(observacoes) & vbCrLf & "E-MAIL ADICIONAL: " & adicionais
    End If
End Function

' Chave de comparacao de lista (C01): sem caixa, sem acento, espacos aparados e colapsados.
' Espelho de Regras.ps1 (ConvertTo-ChaveLista) - a mesma regra nos dois lados.
Public Function ChaveLista(ByVal texto As String) As String
    Dim s As String, origem As String, destino As String, i As Long
    s = UCase$(Trim$(Replace(CStr(texto), Chr$(160), " ")))
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    ' letras acentuadas maiusculas (Latin-1) e a letra sem acento de cada uma
    origem = ChrW(193) & ChrW(192) & ChrW(194) & ChrW(195) & ChrW(196) & ChrW(201) & ChrW(200) & ChrW(202) & ChrW(203) & ChrW(205) & ChrW(204) & ChrW(206) & ChrW(207)
    origem = origem & ChrW(211) & ChrW(210) & ChrW(212) & ChrW(213) & ChrW(214) & ChrW(218) & ChrW(217) & ChrW(219) & ChrW(220) & ChrW(199) & ChrW(209)
    destino = "AAAAAEEEEIIIIOOOOOUUUUCN"
    For i = 1 To Len(origem)
        s = Replace(s, Mid$(origem, i, 1), Mid$(destino, i, 1))
    Next i
    ChaveLista = s
End Function

Public Function SoDigitos(ByVal s As String) As String
    Dim i As Long, c As String, r As String
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If c >= "0" And c <= "9" Then r = r & c
    Next i
    SoDigitos = r
End Function

'----------------------------------------------------------
' Mascara so na exibicao - o banco guarda digitos.
'----------------------------------------------------------
Public Function FormatarCNPJ(ByVal v As Variant) As String
    Dim s As String
    s = SoDigitos(modDB.Nz(v))
    If Len(s) <> 14 Then FormatarCNPJ = s: Exit Function
    FormatarCNPJ = Mid$(s, 1, 2) & "." & Mid$(s, 3, 3) & "." & Mid$(s, 6, 3) & _
                   "/" & Mid$(s, 9, 4) & "-" & Mid$(s, 13, 2)
End Function

Public Function FormatarTelefone(ByVal v As Variant) As String
    Dim s As String
    s = SoDigitos(modDB.Nz(v))
    Select Case Len(s)
        Case 11: FormatarTelefone = "(" & Mid$(s, 1, 2) & ") " & Mid$(s, 3, 5) & "-" & Mid$(s, 8, 4)
        Case 10: FormatarTelefone = "(" & Mid$(s, 1, 2) & ") " & Mid$(s, 3, 4) & "-" & Mid$(s, 7, 4)
        Case Else: FormatarTelefone = s
    End Select
End Function

'----------------------------------------------------------
Public Function CNPJValido(ByVal v As Variant) As Boolean
    Dim s As String, i As Long, soma As Long, d1 As Long, d2 As Long, peso As Long
    s = SoDigitos(modDB.Nz(v))
    If Len(s) <> 14 Then Exit Function
    If s = String$(14, Mid$(s, 1, 1)) Then Exit Function

    peso = 5: soma = 0
    For i = 1 To 12
        soma = soma + CLng(Mid$(s, i, 1)) * peso
        peso = peso - 1
        If peso < 2 Then peso = 9
    Next i
    d1 = 11 - (soma Mod 11)
    If d1 >= 10 Then d1 = 0

    peso = 6: soma = 0
    For i = 1 To 13
        soma = soma + CLng(Mid$(s, i, 1)) * peso
        peso = peso - 1
        If peso < 2 Then peso = 9
    Next i
    d2 = 11 - (soma Mod 11)
    If d2 >= 10 Then d2 = 0

    CNPJValido = (CLng(Mid$(s, 13, 1)) = d1 And CLng(Mid$(s, 14, 1)) = d2)
End Function

'----------------------------------------------------------
' Converte texto digitado em valor tipado. Devolve Null
' quando vazio - NUNCA zero: campo em branco nao e zero.
'----------------------------------------------------------
Public Function ComoNumero(ByVal texto As String) As Variant
    Dim s As String
    s = Trim$(texto)
    If s = "" Then ComoNumero = Null: Exit Function
    s = Replace(s, "R$", "")
    s = Trim$(Replace(s, ".", ""))
    s = Replace(s, ",", ".")
    If Not IsNumeric(Replace(s, ".", Application.International(xlDecimalSeparator))) Then
        ComoNumero = Null
    Else
        ComoNumero = CDbl(Val(s))
    End If
End Function

Public Function ComoData(ByVal texto As String) As Variant
    Dim s As String
    s = Trim$(texto)
    If s = "" Then ComoData = Null: Exit Function
    If Not IsDate(s) Then ComoData = Null: Exit Function
    ComoData = CDate(s)
End Function

' Data digitada como AAAA-MM-DD existe no calendario? DateSerial "corrige" 2026-02-30 para
' 02/03/2026 sem avisar; o conector (ParseExact) recusa. Mesma regra nas telas.
Public Function DataExiste(ByVal ano As Long, ByVal mes As Long, ByVal dia As Long) As Boolean
    Dim d As Date
    If ano < 1900 Or ano > 9999 Or mes < 1 Or mes > 12 Or dia < 1 Or dia > 31 Then Exit Function
    d = DateSerial(ano, mes, dia)
    DataExiste = (Year(d) = ano And Month(d) = mes And Day(d) = dia)
End Function

Public Function ComoInteiro(ByVal texto As String) As Variant
    Dim s As String
    s = Trim$(texto)
    If s = "" Then ComoInteiro = Null: Exit Function
    If Not IsNumeric(s) Then ComoInteiro = Null: Exit Function
    ComoInteiro = CLng(Val(s))
End Function

'----------------------------------------------------------
' DEV-09 (ciclo 008): texto digitado -> valor tipado, SEM perda
' silenciosa. Devolve Array(estado, valor, mensagem):
'   "vazio"    valor Null: limpar o campo de propósito
'   "valido"   valor convertido
'   "invalido" valor Null e a mensagem: quem chama NÃO grava
'              (e não apaga o valor anterior)
' tipo: "D" data, "C" moeda, "N" inteiro. Mesma matriz (C17 a
' C21) do conector, em Regras.ps1.
'----------------------------------------------------------
Public Function Converter(ByVal tipo As String, ByVal texto As String) As Variant
    Dim s As String
    s = Trim$(Replace(texto, Chr$(160), " "))
    If s = "" Then Converter = Array("vazio", Null, ""): Exit Function
    Select Case tipo
        Case "D": Converter = ConverterData(s)
        Case "C": Converter = ConverterMoeda(s)
        Case "N": Converter = ConverterInteiro(s)
        Case Else: Converter = Array("valido", s, "")
    End Select
End Function

' dd/mm/aaaa (dia e mês com 1 ou 2 dígitos) ou aaaa-mm-dd; o dia tem de existir.
Private Function ConverterData(ByVal s As String) As Variant
    Dim partes As Variant, d As Long, m As Long, a As Long
    If s Like "####-##-##" Then
        a = CLng(Left$(s, 4)): m = CLng(Mid$(s, 6, 2)): d = CLng(Mid$(s, 9, 2))
    ElseIf s Like "#/#/####" Or s Like "##/#/####" Or s Like "#/##/####" Or s Like "##/##/####" Then
        partes = Split(s, "/"): d = CLng(partes(0)): m = CLng(partes(1)): a = CLng(partes(2))
    Else
        ConverterData = Array("invalido", Null, "data fora do formato dd/mm/aaaa"): Exit Function
    End If
    If Not DataExiste(a, m, d) Then ConverterData = Array("invalido", Null, "data que não existe no calendário"): Exit Function
    ConverterData = Array("valido", DateSerial(a, m, d), "")
End Function

' Vírgula é o separador decimal (até 2 casas); ponto só como milhar, em grupos de 3.
' "12.50" é recusado: era o erro que virava 1250 sem aviso.
Private Function ConverterMoeda(ByVal s As String) As Variant
    Dim neg As Boolean, i As Long, c As String, pv As Long, inteira As String, dec As String
    Dim grupos As Variant, v As Currency
    s = Replace(Replace(s, "R$", ""), " ", "")
    If Left$(s, 1) = "-" Then neg = True: s = Mid$(s, 2)
    If s = "" Then ConverterMoeda = Array("invalido", Null, "valor sem número"): Exit Function
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If Not (c Like "#" Or c = "." Or c = ",") Then ConverterMoeda = Array("invalido", Null, "valor com caractere que não é número"): Exit Function
    Next i
    pv = InStr(s, ",")
    If pv > 0 Then
        If InStr(pv + 1, s, ",") > 0 Then ConverterMoeda = Array("invalido", Null, "valor com mais de uma vírgula"): Exit Function
        inteira = Left$(s, pv - 1): dec = Mid$(s, pv + 1)
        If InStr(dec, ".") > 0 Then ConverterMoeda = Array("invalido", Null, "separadores misturados (use 1.250,50)"): Exit Function
    Else
        inteira = s
    End If
    If Len(dec) > 2 Then ConverterMoeda = Array("invalido", Null, "mais de 2 casas decimais"): Exit Function
    If InStr(inteira, ".") > 0 Then
        grupos = Split(inteira, ".")
        If Len(grupos(0)) < 1 Or Len(grupos(0)) > 3 Then ConverterMoeda = Array("invalido", Null, "ponto só separa milhar (1.250); centavos vão depois da vírgula"): Exit Function
        For i = 1 To UBound(grupos)
            If Len(grupos(i)) <> 3 Then ConverterMoeda = Array("invalido", Null, "ponto só separa milhar (1.250); centavos vão depois da vírgula"): Exit Function
        Next i
        inteira = Replace(inteira, ".", "")
    End If
    If inteira = "" Then inteira = "0"
    If Len(inteira) > 12 Then ConverterMoeda = Array("invalido", Null, "valor acima do limite"): Exit Function
    v = CCur(inteira)
    If dec <> "" Then v = v + CCur(dec) / (10 ^ Len(dec))
    If v > 999999999999.99@ Then ConverterMoeda = Array("invalido", Null, "valor acima do limite"): Exit Function
    If neg Then v = -v
    ConverterMoeda = Array("valido", v, "")
End Function

' Só dígitos (sinal opcional), dentro do Long do banco.
Private Function ConverterInteiro(ByVal s As String) As Variant
    Dim neg As Boolean, i As Long
    If Left$(s, 1) = "-" Then neg = True: s = Mid$(s, 2)
    If s = "" Then ConverterInteiro = Array("invalido", Null, "número sem dígitos"): Exit Function
    For i = 1 To Len(s)
        If Not (Mid$(s, i, 1) Like "#") Then ConverterInteiro = Array("invalido", Null, "número inteiro, sem vírgula nem ponto"): Exit Function
    Next i
    If Len(s) > 10 Then ConverterInteiro = Array("invalido", Null, "número fora do limite"): Exit Function
    If CDbl(s) > 2147483647# Then ConverterInteiro = Array("invalido", Null, "número fora do limite"): Exit Function
    If neg Then ConverterInteiro = Array("valido", -CLng(s), "") Else ConverterInteiro = Array("valido", CLng(s), "")
End Function

'==========================================================
' REGRAS DE PREENCHIMENTO
' Devolvem "" quando esta tudo certo, ou a mensagem do que
' falta. Mensagem sempre diz o campo, nunca "dados invalidos".
'==========================================================
Public Function CriticarCliente(ByVal d As Object) As String
    If modDB.Nz(d("empresa")) = "" Then CriticarCliente = "Informe a empresa.": Exit Function
    CriticarCliente = CriticarTelefone(d("telefone"))
    If CriticarCliente <> "" Then Exit Function
    If Not NotaDe1a3(d("aderencia")) Then CriticarCliente = "Aderência vai de 1 a 3 (ou em branco).": Exit Function
    If Not NotaDe1a3(d("porte")) Then CriticarCliente = "Porte vai de 1 a 3 (ou em branco).": Exit Function
    If modDB.Nz(d("cnpj")) <> "" Then
        If Not CNPJValido(d("cnpj")) Then
            CriticarCliente = "CNPJ inválido. Deixe em branco se ainda não tiver o número correto."
            Exit Function
        End If
    End If
End Function

Private Function NotaDe1a3(ByVal v As Variant) As Boolean
    If IsNull(v) Or IsEmpty(v) Then NotaDe1a3 = True: Exit Function
    If Not IsNumeric(v) Then Exit Function
    NotaDe1a3 = (CDbl(v) >= 1 And CDbl(v) <= 3 And CDbl(v) = Int(CDbl(v)))
End Function

Public Function CriticarTelefone(ByVal v As Variant) As String
    Dim s As String
    s = SoDigitos(modDB.Nz(v))
    If s = "" Then Exit Function
    If Len(s) < 10 Or Len(s) > 11 Then
        CriticarTelefone = "Telefone deve ter DDD mais o número: 10 ou 11 dígitos." & vbCrLf & vbCrLf & _
                           "Nome de quem atende vai no campo Contato ou em Observações, " & _
                           "e um segundo número também vai em Observações."
    End If
End Function

Public Function CriticarContato(ByVal d As Object) As String
    If modDB.Nz(d("nome")) = "" Then CriticarContato = "Informe o nome do contato (use GERAL se for o contato geral da empresa).": Exit Function
    If modDB.Nz(d("id_cliente")) = "" Then CriticarContato = "Informe a empresa do contato.": Exit Function
    CriticarContato = CriticarTelefone(d("telefone"))
End Function

Public Function CriticarOportunidade(ByVal d As Object) As String
    Dim etapa As String
    If modDB.Nz(d("id_contato")) = "" Then CriticarOportunidade = "Informe o contato.": Exit Function
    etapa = modDB.Nz(d("etapa"))
    If etapa = "" Then CriticarOportunidade = "Informe a etapa.": Exit Function
    If modDB.Nz(d("responsavel")) = "" Then CriticarOportunidade = "Informe o responsável.": Exit Function

    If Not IsNull(d("valor")) Then
        If CDbl(d("valor")) < 0 Then CriticarOportunidade = "O valor não pode ser negativo.": Exit Function
    End If

    ' datas coerentes entre si
    If Not IsNull(d("dt_proposta")) And Not IsNull(d("dt_entrada")) Then
        If CDate(d("dt_proposta")) < CDate(d("dt_entrada")) Then
            CriticarOportunidade = "A data da proposta é anterior à data de entrada."
            Exit Function
        End If
    End If
    If Not IsNull(d("dt_desfecho")) And Not IsNull(d("dt_entrada")) Then
        If CDate(d("dt_desfecho")) < CDate(d("dt_entrada")) Then
            CriticarOportunidade = "A data do desfecho é anterior à data de entrada."
            Exit Function
        End If
    End If

    ' o que a etapa exige
    ' o que a etapa exige: o papel vem da lista (Encerra e Exige familia)
    If modCRM.EtapaEncerra(etapa) Then
        If modDB.Nz(d("motivo_desfecho")) = "" Then
            CriticarOportunidade = "Etapa """ & etapa & """ exige o motivo de desfecho."
            Exit Function
        End If
        If IsNull(d("dt_desfecho")) Then
            CriticarOportunidade = "Etapa """ & etapa & """ exige a data do desfecho."
            Exit Function
        End If
    End If
    If modCRM.EtapaExigeFamilia(etapa) Then
        If modDB.Nz(d("familia")) = "" Then
            CriticarOportunidade = "Etapa """ & etapa & """ exige a família de máquina."
            Exit Function
        End If
    End If
End Function

'----------------------------------------------------------
' Aviso, nao impedimento: o vendedor decide seguir.
'----------------------------------------------------------
Public Function AvisarOportunidade(ByVal d As Object) As String
    Dim etapa As String
    etapa = modDB.Nz(d("etapa"))
    If modCRM.EtapaResultado(etapa) = "GANHO" And IsNull(d("valor")) Then
        AvisarOportunidade = "Pedido fechado sem valor informado. Deseja salvar assim mesmo?"
        Exit Function
    End If
    If Not modCRM.EtapaEncerra(etapa) And IsNull(d("dt_prox_acao")) And IsNull(d("retomar_em")) Then
        AvisarOportunidade = "Sem próxima ação e sem data de retomada: este atendimento vai sumir da fila da manhã. Salvar assim mesmo?"
    End If
End Function

