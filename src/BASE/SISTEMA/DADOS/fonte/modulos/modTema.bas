Attribute VB_Name = "modTema"
'==========================================================
' modTema - GERADO por build\gerar-visual.ps1 a partir de
'           fonte\layout\tema.json e icones.json.
' NAO EDITE A MAO: altere o JSON e rode o gerador.
'==========================================================
Option Explicit

Public Const FONTE_PADRAO As String = "Segoe UI"
Public Const FONTE_SEMIBOLD As String = "Segoe UI Semibold"
Public Const FONTE_FIXA As String = "Consolas"
Public Const FONTE_ICONES As String = "Segoe MDL2 Assets"

Public Const COR_AZUL As Long = 6760983   ' #172A67
Public Const COR_AZUL2 As Long = 7421717   ' #153F71
Public Const COR_AZUL_MEDIO As Long = 10243615   ' #1F4E9C
Public Const COR_AZUL_CLARO As Long = 16314088   ' #E8EEF8
Public Const COR_VERDE As Long = 3190872   ' #58B030
Public Const COR_VERDE_ESCURO As Long = 2264638   ' #3E8E22
Public Const COR_BRANCO As Long = 16777215   ' #FFFFFF
Public Const COR_FUNDO As Long = 16381941   ' #F5F7F9
Public Const COR_FUNDO_CARD As Long = 16777215   ' #FFFFFF
Public Const COR_TEXTO As Long = 5256995   ' #233750
Public Const COR_TEXTO2 As Long = 8746083   ' #637485
Public Const COR_TEXTO_CLARO As Long = 15259081   ' #C9D5E8
Public Const COR_BORDA As Long = 13946310   ' #C6CDD4
Public Const COR_BORDA_SUAVE As Long = 15393757   ' #DDE3EA
Public Const COR_ZEBRA As Long = 16579063   ' #F7F9FC
Public Const COR_SELECAO As Long = 16443612   ' #DCE8FA
Public Const COR_SELECAO_TEXTO As Long = 7024143   ' #0F2E6B
Public Const COR_DESABILITADO As Long = 16118254   ' #EEF1F5
Public Const COR_DESABILITADO_TEXTO As Long = 11839130   ' #9AA6B4
Public Const COR_PERIGO As Long = 2832832   ' #C0392B
Public Const COR_ALERTA As Long = 28344   ' #B86E00
Public Const COR_INFO As Long = 11361291   ' #0B5CAD
Public Const COR_TOM_BOTAO1 As Long = 8140831   ' #1F387C
Public Const COR_TOM_BOTAO2 As Long = 9194029   ' #2D4A8C
Public Const COR_TOM_BOTAO3 As Long = 10247742   ' #3E5E9C
Public Const COR_TOM_BOTAO4 As Long = 11301972   ' #5474AC
Public Const COR_TOM_BOTAO5 As Long = 12356718   ' #6E8CBC
Public Const COR_VERDE_BOTAO As Long = 4625996   ' #4C9646

Public Const ICO_NOVO As Long = &HE710&   ' Add
Public Const ICO_EDITAR As Long = &HE70F&   ' Edit
Public Const ICO_SALVAR As Long = &HE74E&   ' Save
Public Const ICO_SALVAR_PROXIMO As Long = &HE893&   ' Next
Public Const ICO_CANCELAR As Long = &HE711&   ' Cancel
Public Const ICO_FECHAR As Long = &HE8BB&   ' ChromeClose
Public Const ICO_BUSCAR As Long = &HE721&   ' Search
Public Const ICO_FILTRO As Long = &HE71C&   ' Filter
Public Const ICO_ORDENAR As Long = &HE8CB&   ' Sort
Public Const ICO_ATUALIZAR As Long = &HE72C&   ' Refresh
Public Const ICO_EXCLUIR As Long = &HE74D&   ' Delete
Public Const ICO_CLIENTE As Long = &HE821&   ' Work
Public Const ICO_CONTATO As Long = &HE77B&   ' Contact
Public Const ICO_CONTATO_CARTAO As Long = &HE779&   ' ContactInfo
Public Const ICO_PESSOAS As Long = &HE716&   ' People
Public Const ICO_ATENDIMENTO As Long = &HE8BD&   ' Message
Public Const ICO_CALENDARIO As Long = &HE787&   ' Calendar
Public Const ICO_HOJE As Long = &HE8D1&   ' GotoToday
Public Const ICO_TELEFONE As Long = &HE717&   ' Phone
Public Const ICO_EMAIL As Long = &HE715&   ' Mail
Public Const ICO_ANTERIOR As Long = &HE76B&   ' ChevronLeft
Public Const ICO_PROXIMO As Long = &HE76C&   ' ChevronRight
Public Const ICO_PRIMEIRO As Long = &HE892&   ' Previous
Public Const ICO_ULTIMO As Long = &HE893&   ' Next
Public Const ICO_ACIMA As Long = &HE70E&   ' ChevronUp
Public Const ICO_ABAIXO As Long = &HE70D&   ' ChevronDown
Public Const ICO_VINCULO As Long = &HE71B&   ' Link
Public Const ICO_PROMOVER As Long = &HE8FA&   ' AddFriend
Public Const ICO_DESATIVAR As Long = &HE72E&   ' Lock
Public Const ICO_REATIVAR As Long = &HE785&   ' Unlock
Public Const ICO_OK As Long = &HE73E&   ' CheckMark
Public Const ICO_INFO As Long = &HE946&   ' Info
Public Const ICO_ALERTA As Long = &HE7BA&   ' Warning
Public Const ICO_IMPORTANTE As Long = &HE8C9&   ' Important
Public Const ICO_DOCUMENTO As Long = &HE8A5&   ' Document
Public Const ICO_LISTA As Long = &HE8FD&   ' BulletedList
Public Const ICO_ETIQUETA As Long = &HE8EC&   ' Tag
Public Const ICO_BANDEIRA As Long = &HE7C1&   ' Flag
Public Const ICO_RELOGIO As Long = &HE916&   ' Stopwatch
Public Const ICO_INICIO As Long = &HE80F&   ' Home
Public Const ICO_GRUPO As Long = &HE902&   ' Group
Public Const ICO_IMPORTAR As Long = &HE8B5&   ' Import
Public Const ICO_PAINEL As Long = &HE9D9&   ' Diagnostic
Public Const ICO_PINO As Long = &HE718&   ' Pin
Public Const ICO_ABRIR As Long = &HE8A7&   ' OpenInNewWindow
Public Const ICO_PONTO As Long = &HEA3B&   ' CircleFill

' Cor de uma parte (fundo, texto, hover, borda) de uma variante de botao. -1 = nao se aplica.
Public Function CorBotao(ByVal variante As String, ByVal parte As String) As Long
    CorBotao = -1
    Select Case LCase$(variante) & "|" & LCase$(parte)
        Case "primario|fundo": CorBotao = 6760983
        Case "primario|texto": CorBotao = 16777215
        Case "primario|hover": CorBotao = 9388324
        Case "sucesso|fundo": CorBotao = 2264638
        Case "sucesso|texto": CorBotao = 16777215
        Case "sucesso|hover": CorBotao = 3123279
        Case "neutro|fundo": CorBotao = 16777215
        Case "neutro|texto": CorBotao = 6760983
        Case "neutro|hover": CorBotao = 16314088
        Case "neutro|borda": CorBotao = 13946310
        Case "perigo|fundo": CorBotao = 16777215
        Case "perigo|texto": CorBotao = 2832832
        Case "perigo|hover": CorBotao = 15395579
        Case "perigo|borda": CorBotao = 11581157
        Case "cabecalho|fundo": CorBotao = 8273955
        Case "cabecalho|texto": CorBotao = 16777215
        Case "cabecalho|hover": CorBotao = 10704175
        Case "rodape|fundo": CorBotao = 16777215
        Case "rodape|texto": CorBotao = 6760983
        Case "rodape|hover": CorBotao = 16314088
        Case "claro|fundo": CorBotao = 16777215
        Case "claro|texto": CorBotao = 6760983
        Case "claro|hover": CorBotao = 16314088
        Case "sucessoclaro|fundo": CorBotao = 3190872
        Case "sucessoclaro|texto": CorBotao = 16777215
        Case "sucessoclaro|hover": CorBotao = 4441195
        Case "perigoforte|fundo": CorBotao = 2832832
        Case "perigoforte|texto": CorBotao = 16777215
        Case "perigoforte|hover": CorBotao = 4150488
    End Select
End Function

' Cor do selo de uma situacao (parte: texto ou fundo). Situacao desconhecida: azul.
Public Function CorSituacao(ByVal situacao As String, ByVal parte As String) As Long
    Select Case situacao & "|" & LCase$(parte)
        Case "Ação atrasada|texto": CorSituacao = 2829212
        Case "Ação atrasada|fundo": CorSituacao = 14869243
        Case "Retomar hoje|texto": CorSituacao = 23178
        Case "Retomar hoje|fundo": CorSituacao = 13167101
        Case "Ação nesta semana|texto": CorSituacao = 1063531
        Case "Ação nesta semana|fundo": CorSituacao = 15529727
        Case "Em dia|texto": CorSituacao = 3828524
        Case "Em dia|fundo": CorSituacao = 15004902
        Case "Sem próxima ação|texto": CorSituacao = 7760735
        Case "Sem próxima ação|fundo": CorSituacao = 15921647
        Case "Sem etapa definida|texto": CorSituacao = 2832832
        Case "Sem etapa definida|fundo": CorSituacao = 15395579
        Case "Encerrado|texto": CorSituacao = 10392714
        Case "Encerrado|fundo": CorSituacao = 16118769
        Case "OK|texto": CorSituacao = 3828524
        Case "OK|fundo": CorSituacao = 15004902
        Case "SUSPEITO|texto": CorSituacao = 23178
        Case "SUSPEITO|fundo": CorSituacao = 13167101
        Case "DUPLICADO|texto": CorSituacao = 7760735
        Case "DUPLICADO|fundo": CorSituacao = 15921647
        Case "ERRO|texto": CorSituacao = 2829212
        Case "ERRO|fundo": CorSituacao = 14869243
        Case "IGNORAR|texto": CorSituacao = 10392714
        Case "IGNORAR|fundo": CorSituacao = 16118769
        Case "GRAVADO|texto": CorSituacao = 10243615
        Case "GRAVADO|fundo": CorSituacao = 16314088
        Case Else: If LCase$(parte) = "fundo" Then CorSituacao = COR_AZUL_CLARO Else CorSituacao = COR_AZUL
    End Select
End Function

Public Function Glifo(ByVal codigo As Long) As String
    Glifo = ChrW$(codigo)
End Function
