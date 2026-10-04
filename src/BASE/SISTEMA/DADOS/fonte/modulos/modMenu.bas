Attribute VB_Name = "modMenu"
'==========================================================
' modMenu - pontos de entrada dos botoes
'==========================================================
Option Explicit

Public Sub AbrirClientes()
    If Not Pronto() Then Exit Sub
    Dim f As New frmCRM
    f.Abrir "clientes"
End Sub

Public Sub AbrirContatos()
    If Not Pronto() Then Exit Sub
    Dim f As New frmCRM
    f.Abrir "contatos"
End Sub

Public Sub AbrirOportunidades()
    If Not Pronto() Then Exit Sub
    Dim f As New frmCRM
    f.Abrir "oportunidades"
End Sub

'----------------------------------------------------------
' "Nova ficha" das abas: abre a ficha JA em modo Novo (antes
' abria em leitura e ainda era preciso clicar em Novo).
'----------------------------------------------------------
Public Sub NovoCliente()
    If Not Pronto() Then Exit Sub
    Dim f As New frmCRM
    f.Abrir "clientes", 0, 0, , True
End Sub

Public Sub NovoContato()
    If Not Pronto() Then Exit Sub
    Dim f As New frmCRM
    f.Abrir "contatos", 0, 0, , True
End Sub

Public Sub NovoAtendimento()
    If Not Pronto() Then Exit Sub
    Dim f As New frmCRM
    f.Abrir "oportunidades", 0, 0, , True
End Sub

Public Sub AtualizarIndicadores()
    If Not Pronto() Then Exit Sub
    modPainel.AtualizarBasePainel
    ' vindo da aba Inicio: leva ao Painel (recem-atualizado, nao redesenha)
    If Not ActiveSheet Is ThisWorkbook.Worksheets("Painel") Then ThisWorkbook.Worksheets("Painel").Activate
End Sub

'----------------------------------------------------------
' Banco acessivel E esquema compativel. Front de versao
' errada nao grava.
'----------------------------------------------------------
Public Function Pronto() As Boolean
    If modConfig.gModoModelo Then
        MsgBox "Este é o arquivo MODELO, aberto direto da pasta da rede." & vbCrLf & vbCrLf & _
               "Trabalhar por aqui é o que corrompia a planilha antiga: duas pessoas " & _
               "no mesmo arquivo." & vbCrLf & vbCrLf & _
               "Feche, abra o modelo de novo e responda SIM quando ele oferecer " & _
               "instalar sua cópia local.", vbExclamation, "CRM Zapromaq"
        Exit Function
    End If
    ' Consulta bem-sucedida ha menos de 1 minuto = banco no ar: nao
    ' abre conexao de teste. O esquema vem do cache de 5 minutos.
    ' Antes eram duas conexoes ANTES de cada acao do usuario.
    If Not modConfig.ConsultouHaPouco(60) Then
        If Not modDB.BancoOK() Then Exit Function
    End If
    If Not modConfig.EsquemaCompativel() Then Exit Function
    Pronto = True
End Function

'==========================================================
' FILA DA MANHA - acoes vencidas e de hoje, por valor
'==========================================================
Public Sub FilaDaManha()
    Dim d As Variant, i As Long, msg As String, total As Long, limite As Long
    If Not Pronto() Then Exit Sub

    d = modDB.Consultar( _
        "SELECT op.codigo, cl.empresa, ct.nome, op.etapa, op.prox_acao, op.dt_prox_acao," & _
        " op.retomar_em, op.valor, op.responsavel" & _
        " FROM (oportunidades op LEFT JOIN contatos ct ON op.id_contato=ct.id)" & _
        " LEFT JOIN clientes cl ON op.id_cliente=cl.id" & _
        " WHERE op.etapa NOT IN ('Pedido Fechado','Perdido','Descartado')" & _
        " AND ((op.dt_prox_acao IS NOT NULL AND op.dt_prox_acao <= ?)" & _
        "   OR (op.retomar_em IS NOT NULL AND op.retomar_em <= ?))" & _
        " ORDER BY IIf(op.dt_prox_acao Is Null,1,0), op.dt_prox_acao, op.valor DESC, op.id", _
        Array(modDB.P(modDB.adDate, Date), modDB.P(modDB.adDate, Date)))

    If IsEmpty(d) Then
        MsgBox "Nenhuma ação vencida ou para hoje. Fila limpa." & vbCrLf & vbCrLf & _
               "Consulta de " & Format$(Now, "dd/mm/yyyy hh:nn") & ".", _
               vbInformation, "Fila da manhã"
        Exit Sub
    End If

    total = UBound(d, 1)
    limite = total
    If limite > 15 Then limite = 15

    For i = 1 To limite
        msg = msg & modDB.Nz(d(i, 1)) & "  " & _
              Left$(modDB.Nz(d(i, 2)) & Space$(30), 30) & "  " & _
              Format$(modDB.Nz(d(i, 8), "0"), "#,##0") & vbCrLf & _
              "     " & modDB.Nz(d(i, 5), "(sem próxima ação descrita)") & _
              "   [" & modDB.Nz(d(i, 9)) & "]" & vbCrLf
    Next i
    If total > limite Then msg = msg & vbCrLf & "... e mais " & (total - limite) & "."

    MsgBox total & " ação(ões) vencida(s) ou para hoje:" & vbCrLf & vbCrLf & msg, _
           vbInformation, "Fila da manhã  -  " & Format$(Date, "dd/mm/yyyy")
End Sub

'==========================================================
' (A fila de arquivos da IA e o botao "Aplicar pendencias"
' foram descontinuados: a IA grava pelo servidor MCP em
' BASE\IA, pelo mesmo motor e com as mesmas regras do front.
' As alteracoes dela chegam as estacoes pelo log.)
'==========================================================
'==========================================================
' CADASTRO EM LOTE DE PRE-CLIENTES
' Entrada da saida do projeto de qualificacao. A conferencia
' e a gravacao estao em modLote.
'==========================================================
Public Sub CadastroEmLote()
    If Not Pronto() Then Exit Sub
    Dim f As New frmLote
    f.Show
End Sub

'----------------------------------------------------------
' Tela de calibracao visual (Alt+F8 > AbrirCalibracao): nao usa
' o banco. Serve para comparar a tela real da estacao com a
' previa gerada no desenvolvimento (execucao\previa).
'----------------------------------------------------------
Public Sub AbrirCalibracao()
    Dim f As New frmCalibracao
    f.Show
End Sub

Public Sub SobreOSistema()
    MsgBox "CRM Comercial Zapromaq" & vbCrLf & vbCrLf & _
           "Versão do front....: " & modConfig.VERSAO_FRONT & vbCrLf & _
           "Esquema esperado...: " & modConfig.VERSAO_ESQUEMA & vbCrLf & _
           "Ambiente...........: " & modConfig.Ambiente() & vbCrLf & _
           "Banco..............: " & modConfig.CaminhoBanco() & vbCrLf & _
           "Última consulta....: " & modConfig.UltimaConsulta() & vbCrLf & _
           "Usuário............: " & modConfig.UsuarioAtual(), _
           vbInformation, "CRM Zapromaq"
End Sub
