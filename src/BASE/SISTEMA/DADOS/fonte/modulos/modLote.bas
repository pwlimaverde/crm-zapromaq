Attribute VB_Name = "modLote"
'==========================================================
' modLote - cadastro em lote de PRE-CLIENTES, a partir da
'           saida do projeto de qualificacao.
'
' Entrada: o bloco copiado da planilha de qualificacao,
' colado inteiro. Uma linha por empresa, campos separados
' por TABULACAO, nesta ordem confirmada com o comercial:
'
'   1 aderencia   2 porte        3 qualificacao  4 empresa
'   5 cidade      6 estado       7 segmento      8 telefone
'   9 e-mail     10 CNPJ        11 observacoes
'
' REGRA DE NEGOCIO: pre-cliente nao tem contato nem
' oportunidade. Contato e atendimento so existem depois que
' o pre-cliente vira cliente, pela ficha.
'
' QUATRO ESTADOS POR LINHA, e nenhum deles e opiniao:
'   OK .......... cadastra. Duplo clique alterna para IGNORAR.
'   SUSPEITO .... comeca FORA e avisa por que desconfiou.
'                 Duplo clique alterna para ENTRAR.
'   DUPLICADO ... NAO cadastra. CNPJ ja esta na base ou
'                 repetido dentro do proprio lote.
'   ERRO ........ NAO cadastra. Falta dado obrigatorio.
'
' Por que SUSPEITO comeca fora: decisao de 06/10/2026, que
' REVERTE a do comercial de 17/09/2026 (SUSPEITO cadastrava).
' Relatos do comercial de 18 a 29/09/2026: o suspeito que
' entrava sozinho virava cadastro repetido que ninguem revia.
' Passa a comecar fora e so entra por escolha consciente
' (duplo clique). OK tambem alterna: tirar uma linha boa do
' lote nao pode exigir apagar a linha da planilha e colar de novo.
'
' O que o indice unico de CNPJ cobre e o que NAO cobre:
' ele recusa CNPJ repetido no banco - mas 54% da base nao
' tem CNPJ, e ali a unica barreira e a conferencia por nome
' feita aqui.
'==========================================================
Option Explicit

Private Const COL_ADERENCIA As Long = 0
Private Const COL_PORTE As Long = 1
Private Const COL_QUALIF As Long = 2
Private Const COL_EMPRESA As Long = 3
Private Const COL_CIDADE As Long = 4
Private Const COL_UF As Long = 5
Private Const COL_SEGMENTO As Long = 6
Private Const COL_TELEFONE As Long = 7
Private Const COL_EMAIL As Long = 8
Private Const COL_CNPJ As Long = 9
Private Const COL_OBS As Long = 10
Private Const COLUNAS As Long = 11

Private mLinhas As Collection
Private mSegmentos As Object      ' listas Segmento e Qualificacao, lidas UMA vez por conferencia
Private mLote As String           ' DEV-12: LOTE-AAAAMMDD-HHMMSS-<ESTACAO>, gravado no log de cada linha

'==========================================================
' LEITURA E CONFERENCIA
'==========================================================
Public Function Analisar(ByVal texto As String) As Long
    Dim linhas As Variant, i As Long, bruta As String
    Dim cnpjBase As Object, nomeBase As Object
    Dim cnpjLote As Object, nomeLote As Object
    Dim d As Object

    Set mLinhas = New Collection
    mLote = "LOTE-" & Format$(Now, "yyyymmdd-hhnnss") & "-" & modDiag.Estacao()
    If Trim$(texto) = "" Then Exit Function

    Set cnpjBase = CreateObject("Scripting.Dictionary")
    Set nomeBase = CreateObject("Scripting.Dictionary")
    Set cnpjLote = CreateObject("Scripting.Dictionary")
    Set nomeLote = CreateObject("Scripting.Dictionary")
    CarregarBase cnpjBase, nomeBase
    CarregarSegmentos

    texto = Replace(texto, vbCrLf, vbLf)
    texto = Replace(texto, vbCr, vbLf)
    linhas = Split(texto, vbLf)

    For i = LBound(linhas) To UBound(linhas)
        bruta = CStr(linhas(i))
        If Trim$(bruta) <> "" Then
            Set d = Interpretar(bruta, i + 1)
            Conferir d, cnpjBase, nomeBase, cnpjLote, nomeLote
            mLinhas.Add d
        End If
    Next i

    Analisar = mLinhas.Count
End Function

'----------------------------------------------------------
' UMA consulta para a base inteira. Conferir CNPJ por linha
' seria uma ida a rede por empresa - a 26 ms, um lote de 200
' levaria 5 segundos so nisso.
'----------------------------------------------------------
Private Sub CarregarBase(ByRef cnpjBase As Object, ByRef nomeBase As Object)
    Dim dados As Variant, i As Long, c As String, n As String

    dados = modDB.Consultar("SELECT empresa, cnpj FROM clientes")
    If IsEmpty(dados) Then Exit Sub

    For i = 1 To UBound(dados, 1)
        c = modValidacao.SoDigitos(modDB.Nz(dados(i, 2)))
        n = NomeNormalizado(modDB.Nz(dados(i, 1)))
        If c <> "" Then
            If Not cnpjBase.Exists(c) Then cnpjBase.Add c, modDB.Nz(dados(i, 1))
        End If
        If n <> "" Then
            If Not nomeBase.Exists(n) Then nomeBase.Add n, modDB.Nz(dados(i, 1))
        End If
    Next i
End Sub

Private Function Interpretar(ByVal bruta As String, ByVal numero As Long) As Object
    Dim p As Variant, d As Object

    Set d = CreateObject("Scripting.Dictionary")
    d("linha") = numero
    d("status") = "OK"
    d("motivo") = ""
    d("ignorar") = False

    p = Split(bruta, vbTab)

    ' mesma normalizacao do formulario (modValidacao.Normalizar):
    ' o mesmo valor entra igual, venha da ficha ou do lote
    d("aderencia") = Pedaco(p, COL_ADERENCIA)
    d("porte") = Pedaco(p, COL_PORTE)
    d("qualificacao") = Texto(modValidacao.Normalizar("qualificacao", Pedaco(p, COL_QUALIF)))
    d("empresa") = Texto(modValidacao.Normalizar("empresa", Pedaco(p, COL_EMPRESA)))
    d("cidade") = Texto(modValidacao.Normalizar("cidade", Pedaco(p, COL_CIDADE)))
    d("uf") = Texto(modValidacao.Normalizar("uf", Pedaco(p, COL_UF)))
    d("segmento") = Texto(modValidacao.Normalizar("segmento", Pedaco(p, COL_SEGMENTO)))
    d("telefone") = Texto(modValidacao.Normalizar("telefone", Pedaco(p, COL_TELEFONE)))
    d("email") = Texto(modValidacao.Normalizar("email", Pedaco(p, COL_EMAIL)))
    d("cnpj") = Texto(modValidacao.Normalizar("cnpj", Pedaco(p, COL_CNPJ)))
    d("observacoes") = Texto(modValidacao.Normalizar("observacoes", Pedaco(p, COL_OBS)))
    ' seção 5.1: um e-mail por campo; os demais vão para as observações
    d("observacoes") = Texto(modValidacao.ObservacoesComEmails(IIf(d("observacoes") = "", Null, d("observacoes")), _
                                                                 modValidacao.EmailsAdicionais(Pedaco(p, COL_EMAIL))))

    ' DEV-12: exatamente 11 colunas (coluna a mais é dado deslocado, não sobra)
    If (UBound(p) - LBound(p) + 1) <> COLUNAS Then
        Marcar d, "ERRO", "esperadas " & COLUNAS & " colunas, vieram " & _
                          (UBound(p) - LBound(p) + 1)
    End If

    Set Interpretar = d
End Function

Private Function Texto(ByVal v As Variant) As String
    If Not IsNull(v) Then Texto = CStr(v)
End Function

Private Function Pedaco(ByVal p As Variant, ByVal indice As Long) As String
    If indice > UBound(p) Then Exit Function
    Pedaco = Trim$(Replace(CStr(p(indice)), ChrW$(&HA0), " "))
End Function

'----------------------------------------------------------
Private Sub Conferir(ByRef d As Object, ByRef cnpjBase As Object, ByRef nomeBase As Object, _
                     ByRef cnpjLote As Object, ByRef nomeLote As Object)
    Dim c As String, n As String

    Dim erroCampos As String
    ' DEV-12: as mesmas regras do cadastro pela ficha (C30, C31)
    erroCampos = CriticaCampos(d)
    If erroCampos <> "" Then Marcar d, "ERRO", erroCampos

    c = d("cnpj")
    n = NomeNormalizado(d("empresa"))

    If c <> "" Then
        If Len(c) <> 14 Then
            Marcar d, "ERRO", "CNPJ com " & Len(c) & " digitos"
        ElseIf cnpjBase.Exists(c) Then
            Marcar d, "DUPLICADO", "CNPJ ja cadastrado: " & cnpjBase(c)
        ElseIf cnpjLote.Exists(c) Then
            ' a primeira linha ja foi conferida: o estado dela e o final
            Marcar d, "DUPLICADO", MotivoRepetidoNoLote(CStr(cnpjLote(c)("linha")), CStr(cnpjLote(c)("status")))
        End If
        ' guarda a LINHA (o dicionario), nao so o numero: o motivo de quem
        ' repetir depois precisa saber se ela entra ou comeca fora
        If Not cnpjLote.Exists(c) Then cnpjLote.Add c, d
    Else
        ' Sem CNPJ o indice unico do banco nao protege nada:
        ' o Access aceita varios nulos. Sobra o nome.
        If n <> "" Then
            If nomeBase.Exists(n) Then
                Marcar d, "SUSPEITO", "nome parecido com o ja cadastrado: " & nomeBase(n)
            ElseIf nomeLote.Exists(n) Then
                Marcar d, "SUSPEITO", "nome repetido na linha " & nomeLote(n) & " deste lote"
            End If
        End If
    End If

    If n <> "" Then
        If Not nomeLote.Exists(n) Then nomeLote.Add n, d("linha")
    End If

    ' Telefone segue a regra da ficha: DDD + numero, 10 ou 11 digitos,
    ' um por campo. Fora disso o numero NAO e descartado: vai para as
    ' observacoes (como na migracao de 16/09) e a linha fica SUSPEITA.
    If d("telefone") <> "" Then
        If Len(d("telefone")) < 10 Or Len(d("telefone")) > 11 Then
            d("observacoes") = Trim$(d("observacoes") & IIf(d("observacoes") = "", "", " ") & _
                               "[telefone do lote: " & d("telefone") & "]")
            Marcar d, "SUSPEITO", "telefone fora de 10-11 digitos (movido para observacoes)"
            d("telefone") = ""
        End If
    End If

    ' DEV-12: valor fora da lista não entra (a ficha também só aceita valor da lista)
    If d("segmento") <> "" Then
        If Not NaLista("Segmento", d("segmento")) Then Marcar d, "ERRO", "segmento fora da lista: " & d("segmento")
    End If
    If d("qualificacao") <> "" Then
        If Not NaLista("Qualificacao", d("qualificacao")) Then Marcar d, "ERRO", "qualifica" & ChrW$(231) & ChrW$(227) & "o fora da lista: " & d("qualificacao")
    End If

    ' depois de TODAS as conferencias: o estado final decide se comeca fora
    d("ignorar") = ComecaIgnorada(CStr(d("status")))
End Sub

'----------------------------------------------------------
' REGRAS DE ENTRADA DE CADA LINHA (sem banco; ver o cabecalho)
'----------------------------------------------------------
Public Function ComecaIgnorada(ByVal status As String) As Boolean
    ComecaIgnorada = (status = "SUSPEITO")
End Function

' So OK e SUSPEITO alternam: DUPLICADO e ERRO nao sao opiniao.
Public Function PodeAlternar(ByVal status As String) As Boolean
    PodeAlternar = (status = "OK" Or status = "SUSPEITO")
End Function

' Textos da tela (frmLote) que explicam a regra: ficam aqui, junto dela.
Public Function TextoAjuda() As String
    ' cabe nas 2 linhas de lblAjuda (736 x 26 pt)
    TextoAjuda = "Cole o bloco da planilha de qualificação, uma empresa por linha: aderência, porte, " & _
        "qualificação, empresa, cidade, estado, segmento, telefone, e-mail, CNPJ, observações. " & _
        "Duplo clique alterna OK e SUSPEITO entre entrar e ficar fora; Suspeito começa fora."
End Function

Public Function AvisoNaoAlterna() As String
    AvisoNaoAlterna = "Só OK e SUSPEITO alternam." & vbCrLf & vbCrLf & _
        "DUPLICADO e ERRO não entram: um já está na base, o outro está " & _
        "sem dado obrigatório. Corrija na planilha e confira de novo."
End Function

' Motivo do DUPLICADO por CNPJ repetido no proprio lote. A primeira
' linha e a que vale; se ela e SUSPEITO, comeca FORA e, sem aviso,
' nenhuma das duas entraria - entao o motivo pede para inclui-la.
Public Function MotivoRepetidoNoLote(ByVal linha As String, ByVal statusDela As String) As String
    MotivoRepetidoNoLote = "CNPJ repetido na linha " & linha & " deste lote"
    If statusDela = "SUSPEITO" Then
        MotivoRepetidoNoLote = MotivoRepetidoNoLote & " (a linha " & linha & " e SUSPEITO e comeca fora:" & _
                               " para cadastrar a empresa, inclua a linha " & linha & " com duplo clique)"
    End If
End Function

' O rotulo da lista mostra o estado EFETIVO da linha.
Public Function RotuloLinha(ByVal status As String, ByVal ignorar As Boolean) As String
    Select Case status
        Case "OK":       RotuloLinha = IIf(ignorar, "OK - IGNORAR", "OK")
        Case "SUSPEITO": RotuloLinha = IIf(ignorar, "SUSPEITO - FORA", "SUSPEITO - ENTRA")
        Case Else:       RotuloLinha = status
    End Select
End Function

'----------------------------------------------------------
' ERRO e DUPLICADO nao viram SUSPEITO depois: o estado pior
' manda, senao uma conferencia posterior apagaria a anterior.
'----------------------------------------------------------
Private Sub Marcar(ByRef d As Object, ByVal status As String, ByVal motivo As String)
    If Peso(status) > Peso(CStr(d("status"))) Then d("status") = status
    If d("motivo") = "" Then
        d("motivo") = motivo
    Else
        d("motivo") = d("motivo") & "; " & motivo
    End If
End Sub

Private Function Peso(ByVal status As String) As Long
    Select Case status
        Case "OK": Peso = 0
        Case "SUSPEITO": Peso = 1
        Case "DUPLICADO": Peso = 2
        Case "ERRO": Peso = 3
    End Select
End Function

'----------------------------------------------------------
' DEV-12: críticas de campo da linha, as mesmas da ficha, sem banco
' (C30 DV do CNPJ, C31 nota inteira e UF da lista; tamanhos do banco).
' Devolve "" ou os motivos separados por "; ".
'----------------------------------------------------------
Public Function CriticaCampos(ByVal d As Object) As String
    Dim m As String
    If CStr(d("empresa")) = "" Then m = m & "; sem nome de empresa"
    If CStr(d("uf")) <> "" Then
        If Not modListas.UFValida(CStr(d("uf"))) Then m = m & "; UF fora da lista: " & d("uf")
    End If
    If Not NotaValida(CStr(d("aderencia"))) Then m = m & "; aderência deve ser inteira de 1 a 3"
    If Not NotaValida(CStr(d("porte"))) Then m = m & "; porte deve ser inteiro de 1 a 3"
    If CStr(d("cnpj")) <> "" Then
        If Len(CStr(d("cnpj"))) <> 14 Then
            m = m & "; CNPJ com " & Len(CStr(d("cnpj"))) & " dígitos"
        ElseIf Not modValidacao.CNPJValido(CStr(d("cnpj"))) Then
            m = m & "; dígito verificador do CNPJ não fecha"
        End If
    End If
    If Len(CStr(d("empresa"))) > 120 Then m = m & "; empresa passa de 120 caracteres"
    If Len(CStr(d("cidade"))) > 60 Then m = m & "; cidade passa de 60 caracteres"
    If Len(CStr(d("email"))) > 120 Then m = m & "; e-mail passa de 120 caracteres"
    If Len(CStr(d("segmento"))) > 60 Then m = m & "; segmento passa de 60 caracteres"
    If Len(CStr(d("qualificacao"))) > 40 Then m = m & "; qualificação passa de 40 caracteres"
    If m <> "" Then CriticaCampos = Mid$(m, 3)
End Function

' Nota 1 a 3, INTEIRA (DEV-09): "1,5" era arredondada por CLng e entrava como 2.
Public Function NotaValida(ByVal v As String) As Boolean
    Dim conv As Variant
    conv = modValidacao.Converter("N", v)
    If conv(0) = "vazio" Then NotaValida = True: Exit Function
    If conv(0) = "invalido" Then Exit Function
    NotaValida = (conv(1) >= 1 And conv(1) <= 3)
End Function

' A lista e lida UMA vez por conferencia (CarregarSegmentos), nao
' uma vez por linha: a 26 ms por ida a rede, 200 linhas seriam 5 s.
Private Sub CarregarSegmentos()
    Dim d As Variant, i As Long, tipo As Variant
    ' Segmento e Qualificação: chave "tipo|valor" (o tipo separa as duas listas)
    Set mSegmentos = CreateObject("Scripting.Dictionary")
    For Each tipo In Array("Segmento", "Qualificacao")
        d = modCRM.ObterLista(CStr(tipo))
        If Not IsEmpty(d) Then
            mSegmentos(CStr(tipo)) = True
            For i = 1 To UBound(d, 1)
                mSegmentos(CStr(tipo) & "|" & LCase$(modDB.Nz(d(i, 1)))) = True
            Next i
        End If
    Next tipo
End Sub

' Lista vazia (não cadastrada) não recusa nada: é o mesmo critério do conector.
Private Function NaLista(ByVal tipo As String, ByVal valor As String) As Boolean
    If mSegmentos Is Nothing Then NaLista = True: Exit Function
    If Not mSegmentos.Exists(tipo) Then NaLista = True: Exit Function
    NaLista = mSegmentos.Exists(tipo & "|" & LCase$(valor))
End Function

'----------------------------------------------------------
' Normalizacao do nome, so para desconfiar de duplicata:
' maiusculas, sem acento, sem pontuacao, sem forma juridica
' e sem espaco duplo. Nao e o nome gravado - o cadastro fica
' com o nome como veio.
'----------------------------------------------------------
Public Function NomeNormalizado(ByVal s As String) As String
    Dim i As Long, c As String, r As String, antes As String
    Dim formas As Variant, f As Variant

    s = UCase$(Trim$(s))
    antes = "AAAAAEEEEIIIIOOOOOUUUUCN"
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        Select Case c
            Case ChrW$(&HC1), ChrW$(&HC0), ChrW$(&HC3), ChrW$(&HC2), ChrW$(&HC4): c = "A"
            Case ChrW$(&HC9), ChrW$(&HC8), ChrW$(&HCA), ChrW$(&HCB): c = "E"
            Case ChrW$(&HCD), ChrW$(&HCC), ChrW$(&HCE), ChrW$(&HCF): c = "I"
            Case ChrW$(&HD3), ChrW$(&HD2), ChrW$(&HD5), ChrW$(&HD4), ChrW$(&HD6): c = "O"
            Case ChrW$(&HDA), ChrW$(&HD9), ChrW$(&HDB), ChrW$(&HDC): c = "U"
            Case ChrW$(&HC7): c = "C"
            Case ChrW$(&HD1): c = "N"
        End Select
        If (c >= "A" And c <= "Z") Or (c >= "0" And c <= "9") Then
            r = r & c
        Else
            r = r & " "
        End If
    Next i

    formas = Array(" LTDA ", " ME ", " EPP ", " EIRELI ", " MEI ", " SA ", " S A ", _
                   " CIA ", " INDUSTRIA ", " COMERCIO ", " E ", " DE ", " DA ", " DO ", _
                   " DAS ", " DOS ")
    r = " " & r & " "
    For Each f In formas
        Do While InStr(r, f) > 0
            r = Replace(r, f, " ")
        Loop
    Next f
    Do While InStr(r, "  ") > 0
        r = Replace(r, "  ", " ")
    Loop
    NomeNormalizado = Trim$(r)
End Function

'==========================================================
' LEITURA PELO FORMULARIO
'==========================================================
Public Function Quantidade() As Long
    If mLinhas Is Nothing Then Exit Function
    Quantidade = mLinhas.Count
End Function

Public Function Valor(ByVal i As Long, ByVal campo As String) As String
    If mLinhas Is Nothing Then Exit Function
    If i < 1 Or i > mLinhas.Count Then Exit Function
    Valor = CStr(mLinhas(i)(campo))
End Function

Public Function Ignorada(ByVal i As Long) As Boolean
    If mLinhas Is Nothing Then Exit Function
    If i < 1 Or i > mLinhas.Count Then Exit Function
    Ignorada = CBool(mLinhas(i)("ignorar"))
End Function

'----------------------------------------------------------
' OK e SUSPEITO se alternam (PodeAlternar); ERRO e DUPLICADO
' nao entram de jeito nenhum.
'----------------------------------------------------------
Public Function AlternarIgnorar(ByVal i As Long) As Boolean
    If mLinhas Is Nothing Then Exit Function
    If i < 1 Or i > mLinhas.Count Then Exit Function
    If Not PodeAlternar(CStr(mLinhas(i)("status"))) Then Exit Function
    mLinhas(i)("ignorar") = Not CBool(mLinhas(i)("ignorar"))
    AlternarIgnorar = True
End Function

Public Function Contagem(ByVal status As String) As Long
    Dim i As Long, n As Long
    If mLinhas Is Nothing Then Exit Function
    For i = 1 To mLinhas.Count
        If CStr(mLinhas(i)("status")) = status Then n = n + 1
    Next i
    Contagem = n
End Function

Public Function ACadastrar() As Long
    Dim i As Long, n As Long
    If mLinhas Is Nothing Then Exit Function
    For i = 1 To mLinhas.Count
        If Entra(i) Then n = n + 1
    Next i
    ACadastrar = n
End Function

Private Function Entra(ByVal i As Long) As Boolean
    Dim s As String
    s = CStr(mLinhas(i)("status"))
    If s <> "OK" And s <> "SUSPEITO" Then Exit Function
    Entra = Not CBool(mLinhas(i)("ignorar"))
End Function

'==========================================================
' GRAVACAO
'
' Uma insercao por linha, cada uma com seu id devolvido por
' @@IDENTITY na propria conexao. Nao existe INSERT em bloco
' que devolva id no ACE, e id e o que o cadastro do contato
' vai usar depois.
'==========================================================
Public Function Gravar(ByRef gravados As Long, ByRef falhas As String) As Boolean
    Dim i As Long, d As Object, novo As Long, cod As Long

    gravados = 0
    falhas = ""
    If mLinhas Is Nothing Then Exit Function
    If Not modMenu.Pronto() Then Exit Function

    Application.Cursor = 2                        ' xlWait
    For i = 1 To mLinhas.Count
        If Entra(i) Then
            Set d = CreateObject("Scripting.Dictionary")
            d("empresa") = Nulo(Valor(i, "empresa"))
            d("cnpj") = Nulo(Valor(i, "cnpj"))
            d("cidade") = Nulo(Valor(i, "cidade"))
            d("uf") = Nulo(Valor(i, "uf"))
            d("segmento") = Nulo(Valor(i, "segmento"))
            d("telefone") = Nulo(Valor(i, "telefone"))
            d("email") = Nulo(Valor(i, "email"))
            d("qualificacao") = Nulo(Valor(i, "qualificacao"))
            d("aderencia") = NuloNumero(Valor(i, "aderencia"))
            d("porte") = NuloNumero(Valor(i, "porte"))
            d("observacoes") = Nulo(Valor(i, "observacoes"))

            ' DEV-12: o MESMO validador da ficha, logo antes do INSERT
            Dim critica As String
            critica = modValidacao.CriticarCliente(d)
            If critica <> "" Then
                falhas = falhas & "linha " & Valor(i, "linha") & ": " & critica & vbCrLf
                mLinhas(i)("status") = "ERRO"
                mLinhas(i)("motivo") = critica
                GoTo proximaLinha
            End If
            ' DEV-13: identificador ESTÁVEL por linha (CNPJ, nome e cidade): colar e gravar de novo a
            ' mesma linha devolve a empresa já cadastrada em vez de duplicar
            modCRM.gIdOperacao = Left$("LOTE|" & Valor(i, "cnpj") & "|" & NomeNormalizado(Valor(i, "empresa")) & "|" & Valor(i, "cidade"), 120)
            modCRM.gHashOperacao = modDB.HashTexto(modDB.TextoDoDicionario(d))
            On Error Resume Next
            Err.Clear
            ' identificador do lote e da linha no log: o reprocessamento da mesma colagem é rastreável
            novo = modCRM.CriarCliente(d, False, cod, "lote " & mLote & " linha " & Valor(i, "linha"))
            If Err.Number <> 0 Then
                falhas = falhas & "linha " & Valor(i, "linha") & ": " & Err.Description & vbCrLf
                mLinhas(i)("status") = "ERRO"
                mLinhas(i)("motivo") = Err.Description
                Err.Clear
            ElseIf novo > 0 Then
                gravados = gravados + 1
                mLinhas(i)("status") = "GRAVADO"
                mLinhas(i)("motivo") = "id " & novo
            End If
            On Error GoTo 0
        End If
proximaLinha:
    Next i
    modCRM.gIdOperacao = "": modCRM.gHashOperacao = ""
    Application.Cursor = -4143                    ' xlDefault

    Gravar = (gravados > 0)
End Function

' Texto vazio vira NULO, nunca cadeia vazia: o indice unico
' de CNPJ aceita varios nulos, mas recusaria a segunda vazia.
Private Function Nulo(ByVal s As String) As Variant
    If Trim$(s) = "" Then Nulo = Null Else Nulo = s
End Function

Private Function NuloNumero(ByVal s As String) As Variant
    If Trim$(s) = "" Then NuloNumero = Null: Exit Function
    If Not IsNumeric(s) Then NuloNumero = Null: Exit Function
    NuloNumero = CLng(s)
End Function
