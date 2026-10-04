Attribute VB_Name = "modConfig"
'==========================================================
' modConfig - caminhos, versao e identidade do ambiente
'
' CRM Comercial Zapromaq
' Nenhuma letra de unidade aqui: o caminho e UNC e mora na
' aba Config do frontend.
'==========================================================
Option Explicit

' VERSAO_FRONT sobe a cada montagem publicada. O montar-frontend
' grava este mesmo numero na tabela config do banco, e o
' Workbook_Open compara: copia antiga na estacao avisa sozinha.
Public Const VERSAO_FRONT As String = "1.4"
Public Const VERSAO_ESQUEMA As String = "1.0"
Public Const ABA_CONFIG As String = "Config"

' Ligado pelo Workbook_Open quando o arquivo esta rodando da pasta
' da rede. Enquanto estiver ligado, nada grava - nem no banco nem
' no proprio arquivo.
Public gModoModelo As Boolean

' Ligado pela ficha sempre que algo foi gravado, e lido pela
' grade quando a ficha fecha. Sem isto, editar pela ficha e
' fechar deixava a LINHA DA GRADE com o valor antigo ate
' alguem clicar em Atualizar - e a tela mentia.
' Variavel de modulo, nao propriedade do formulario: Unload
' destroi o objeto do formulario e levaria a resposta junto.
' Lista de ids gravados pela ficha enquanto ela esteve aberta,
' separados por virgula. Nao e um Sim/Nao: com "Salvar e prox."
' o vendedor edita varios registros sem fechar a ficha, e a
' grade precisa saber TODOS para reescrever cada linha.
Public gGravados As String

' Hora da ultima consulta bem-sucedida, em memoria. A celula
' Config!B4 so e regravada no maximo uma vez por minuto: gravar
' a cada consulta marcava o arquivo como alterado e disparava
' os eventos de celula da pasta a toa.
Private mUltimaConsulta As Date

' Tabela config do banco (versao_esquema, versao_front...) lida numa
' consulta so e guardada por 5 minutos: conferir o esquema a CADA
' clique custava duas conexoes antes de qualquer acao.
Private mConfig As Object
Private mConfigLidaEm As Date
Private Const VALIDADE_CONFIG_MIN As Long = 5

' Usado apenas se a aba Config nao existir ou estiver vazia.
' O build grava o caminho UNC real em Config!B2.
Private Const CAMINHO_PADRAO As String = _
    "\\servidor\COMERCIAL\1 - COMERCIAL ZAPROMAQ\01 - CRM\01 - CONTROLE\BASE\crm_zapromaq.accdb"

'----------------------------------------------------------
Public Function CaminhoBanco() As String
    Dim v As String
    On Error Resume Next
    v = CStr(ThisWorkbook.Worksheets(ABA_CONFIG).Range("B2").Value)
    On Error GoTo 0
    If Trim$(v) = "" Then v = CAMINHO_PADRAO
    CaminhoBanco = v
End Function

Public Function Ambiente() As String
    Dim v As String
    On Error Resume Next
    v = CStr(ThisWorkbook.Worksheets(ABA_CONFIG).Range("B3").Value)
    On Error GoTo 0
    If Trim$(v) = "" Then v = "PRODUCAO"
    Ambiente = UCase$(Trim$(v))
End Function

Public Function UsuarioAtual() As String
    Dim u As String
    On Error Resume Next
    u = Environ$("USERNAME")
    On Error GoTo 0
    If u = "" Then u = Application.UserName
    UsuarioAtual = Left$(u, 50)
End Function

'----------------------------------------------------------
' Front so grava se o esquema do banco for o esperado.
'----------------------------------------------------------
'----------------------------------------------------------
' Valor da tabela config do banco, com cache de 5 minutos.
' forcar = True rele agora (abertura do arquivo).
'----------------------------------------------------------
Public Function ValorConfig(ByVal chave As String, Optional ByVal forcar As Boolean = False) As String
    Dim d As Variant, i As Long
    If forcar Or mConfig Is Nothing Or DateDiff("n", mConfigLidaEm, Now) >= VALIDADE_CONFIG_MIN Then
        Set mConfig = CreateObject("Scripting.Dictionary")
        mConfig.CompareMode = 1
        d = modDB.Consultar("SELECT chave, valor FROM config")
        If Not IsEmpty(d) Then
            For i = 1 To UBound(d, 1)
                mConfig(modDB.Nz(d(i, 1))) = modDB.Nz(d(i, 2))
            Next i
        End If
        mConfigLidaEm = Now
    End If
    If mConfig.Exists(chave) Then ValorConfig = mConfig(chave)
End Function

' Houve consulta bem-sucedida ha menos de N segundos?
Public Function ConsultouHaPouco(ByVal segundos As Long) As Boolean
    If mUltimaConsulta = 0 Then Exit Function
    ConsultouHaPouco = (DateDiff("s", mUltimaConsulta, Now) < segundos)
End Function

Public Function EsquemaCompativel(Optional ByVal silencioso As Boolean = False, _
                                  Optional ByVal forcar As Boolean = False) As Boolean
    Dim v As Variant
    On Error GoTo falhou
    v = ValorConfig("versao_esquema", forcar)
    If CStr(v) = VERSAO_ESQUEMA Then
        EsquemaCompativel = True
        Exit Function
    End If
falhou:
    EsquemaCompativel = False
    If Not silencioso Then
        MsgBox "Esta copia do CRM espera o esquema " & VERSAO_ESQUEMA & _
               " e o banco esta em " & CStr(v) & "." & vbCrLf & vbCrLf & _
               "Atualize o arquivo do CRM antes de continuar. Gravar assim pode corromper dados.", _
               vbCritical, "CRM Zapromaq"
    End If
End Function

'----------------------------------------------------------
' COPIA DESATUALIZADA
'
' O banco guarda a versao do front publicado. A copia local
' que ficou para tras avisa na abertura - antes dependia de
' alguem reparar na data do arquivo, e ninguem repara.
'
' Nao bloqueia: versao de front diferente nao corrompe dado,
' quem corrompe e esquema incompativel, e isso o
' EsquemaCompativel ja barra. Aqui e aviso, nao tranca.
'----------------------------------------------------------
Public Function VersaoPublicada() As String
    On Error Resume Next
    VersaoPublicada = ValorConfig("versao_front")
    On Error GoTo 0
End Function

Public Sub ConferirVersaoCopia()
    Dim publicada As String
    On Error Resume Next
    publicada = VersaoPublicada()
    On Error GoTo 0
    If publicada = "" Then Exit Sub
    ' so avisa se ESTA copia for mais antiga: uma copia mais nova (teste,
    ' ou publicada antes de o banco registrar) nao manda ninguem "atualizar"
    If NumeroVersao(VERSAO_FRONT) >= NumeroVersao(publicada) Then Exit Sub

    MsgBox "Esta copia do CRM e a versao " & VERSAO_FRONT & _
           " e a publicada na rede e a " & publicada & "." & vbCrLf & vbCrLf & _
           "Feche o CRM e abra pelo atalho INICIAR-CRM: ele copia a versao nova " & _
           "e abre sozinho, sem voce ir na pasta da rede." & vbCrLf & vbCrLf & _
           "Da para continuar trabalhando nesta - o banco e o mesmo - mas o que " & _
           "foi corrigido depois da sua copia nao esta aqui.", _
           vbExclamation, "CRM Zapromaq  -  copia desatualizada"
End Sub

' "1.10" > "1.9": compara maior e menor como numeros, nao como texto
Public Function NumeroVersao(ByVal v As String) As Double
    Dim p As Variant
    p = Split(Trim$(v) & ".0", ".")
    NumeroVersao = Val(p(0)) * 100000 + Val(p(1))
End Function

Public Sub RegistrarUltimaConsulta()
    Dim antes As Date
    antes = mUltimaConsulta
    mUltimaConsulta = Now
    If mUltimaConsulta - antes < TimeSerial(0, 1, 0) Then Exit Sub
    On Error Resume Next
    Dim ev As Boolean, salvo As Boolean
    ev = Application.EnableEvents
    salvo = ThisWorkbook.Saved
    Application.EnableEvents = False
    ThisWorkbook.Worksheets(ABA_CONFIG).Range("B4").Value = mUltimaConsulta
    ' anotacao interna: nao pode marcar o arquivo como alterado (com o
    ' feed consultando a cada 25 s, TODO fechamento perguntaria se salva)
    ThisWorkbook.Saved = salvo
    Application.EnableEvents = ev
    On Error GoTo 0
End Sub

Public Function UltimaConsulta() As String
    Dim v As Variant
    If mUltimaConsulta > 0 Then
        UltimaConsulta = Format$(mUltimaConsulta, "dd/mm/yyyy hh:nn")
        Exit Function
    End If
    On Error Resume Next
    v = ThisWorkbook.Worksheets(ABA_CONFIG).Range("B4").Value
    On Error GoTo 0
    If IsEmpty(v) Or IsNull(v) Or Not IsDate(v) Then
        UltimaConsulta = "nunca"
    Else
        UltimaConsulta = Format$(v, "dd/mm/yyyy hh:nn")
    End If
End Function
