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
' Normaliza pelo NOME do campo. Campo desconhecido so leva
' Trim - nunca se inventa regra por semelhanca de nome.
'----------------------------------------------------------
Public Function Normalizar(ByVal campo As String, ByVal valor As Variant) As Variant
    Dim s As String
    If IsNull(valor) Then Normalizar = Null: Exit Function
    s = Trim$(Replace(CStr(valor), Chr$(160), " "))
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    Select Case LCase$(campo)
        ' MAIUSCULAS: identificacao e cadastro curto
        Case "empresa", "cidade", "uf", "maquina", "orcamento"
            s = UCase$(s)
        ' minusculas
        Case "email"
            s = LCase$(s)
        ' so digitos
        Case "cnpj"
            s = SoDigitos(s)
        ' telefone: SO DIGITOS, um numero por campo.
        ' Nome de pessoa vai no campo Contato ou em Observacoes -
        ' na migracao, 344 cadastros traziam nome junto do numero.
        Case "telefone"
            s = SoDigitos(s)
        ' como digitado: nome, cargo, responsavel, observacoes,
        ' prox_acao, ctx_resumo e as listas suspensas
    End Select

    If Len(s) = 0 Then Normalizar = Null Else Normalizar = s
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

Public Function ComoInteiro(ByVal texto As String) As Variant
    Dim s As String
    s = Trim$(texto)
    If s = "" Then ComoInteiro = Null: Exit Function
    If Not IsNumeric(s) Then ComoInteiro = Null: Exit Function
    ComoInteiro = CLng(Val(s))
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
    Select Case etapa
        Case "Pedido Fechado", "Perdido", "Descartado"
            If modDB.Nz(d("motivo_desfecho")) = "" Then
                CriticarOportunidade = "Etapa """ & etapa & """ exige o motivo de desfecho."
                Exit Function
            End If
            If IsNull(d("dt_desfecho")) Then
                CriticarOportunidade = "Etapa """ & etapa & """ exige a data do desfecho."
                Exit Function
            End If
        Case "Elaboração da Proposta", "Negociação", "Proposta em Análise", "Proposta em Stand By"
            If modDB.Nz(d("familia")) = "" Then
                CriticarOportunidade = "Etapa """ & etapa & """ exige a família de máquina."
                Exit Function
            End If
    End Select
End Function

'----------------------------------------------------------
' Aviso, nao impedimento: o vendedor decide seguir.
'----------------------------------------------------------
Public Function AvisarOportunidade(ByVal d As Object) As String
    Dim etapa As String
    etapa = modDB.Nz(d("etapa"))
    If etapa = "Pedido Fechado" And IsNull(d("valor")) Then
        AvisarOportunidade = "Pedido fechado sem valor informado. Deseja salvar assim mesmo?"
        Exit Function
    End If
    If Not modCalcEncerrada(etapa) And IsNull(d("dt_prox_acao")) And IsNull(d("retomar_em")) Then
        AvisarOportunidade = "Sem próxima ação e sem data de retomada: este atendimento vai sumir da fila da manhã. Salvar assim mesmo?"
    End If
End Function

Private Function modCalcEncerrada(ByVal etapa As String) As Boolean
    modCalcEncerrada = (etapa = "Pedido Fechado" Or etapa = "Perdido" Or etapa = "Descartado")
End Function
