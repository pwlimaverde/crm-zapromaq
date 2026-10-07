Attribute VB_Name = "modSchema"
'==========================================================
' modSchema - metadados dos campos de cada ficha
'
' Formato: campo|Rotulo|tipo|lista|secao|obrigatorio|editavel
'   tipo:  T texto   M memo    N numero   C moeda
'          D data    L lista   K vinculo  R somente leitura
'   obrigatorio: 1 sim  0 nao
'   editavel:    1 sim  0 nao (vem por JOIN ou e calculado)
'
' Uma instrucao Add por campo, de proposito: o VBA aceita no
' maximo 25 continuacoes de linha numa mesma instrucao, e um
' Array( ) encadeado estoura esse limite quando a ficha cresce.
' O erro aparece so na importacao, como HRESULT 0x800A9D00.
'
' Acrescentar campo = uma linha Add aqui + uma coluna no banco.
' Campo que muda consulta, Painel ou regra exige mais do que
' isso: confira antes de prometer que e so uma linha.
'==========================================================
Option Explicit

Private Sub Add(ByRef v As Variant, ByVal spec As String)
    If IsEmpty(v) Then
        ReDim v(0 To 0)
    Else
        ReDim Preserve v(0 To UBound(v) + 1)
    End If
    v(UBound(v)) = spec
End Sub

Public Function TituloDe(ByVal tabela As String) As String
    Select Case tabela
        Case "clientes":      TituloDe = "Clientes  -  empresas e fila de prospecção"
        Case "contatos":      TituloDe = "Contatos  -  pessoas vinculadas à empresa"
        Case "oportunidades": TituloDe = "Oportunidades  -  atendimentos"
    End Select
End Function

Public Function ChaveDe(ByVal tabela As String) As String
    ChaveDe = "id"
End Function

'----------------------------------------------------------
Public Function CamposDe(ByVal tabela As String) As Variant
    Dim c As Variant

    Select Case tabela

    Case "clientes"
    Add c, "codigo_cliente|Código|R||1. Identificação|0|0"
    Add c, "estagio|Estágio|R||1. Identificação|0|0"
    Add c, "empresa|Empresa|T||1. Identificação|1|1"
    Add c, "cnpj|CNPJ|T||1. Identificação|0|1"
    Add c, "cidade|Cidade|T||2. Localização|0|1"
    Add c, "uf|UF|L|UF|2. Localização|0|1"
    Add c, "segmento|Segmento|L|Segmento|2. Localização|0|1"
    Add c, "telefone|Telefone geral|T||2. Localização|0|1"
    Add c, "email|E-mail geral|T||2. Localização|0|1"
    Add c, "responsavel|Responsável|L|Responsavel|3. Prospecção|0|1"
    Add c, "aderencia|Aderência (1-3)|N||3. Prospecção|0|1"
    Add c, "porte|Porte (1-3)|N||3. Prospecção|0|1"
    Add c, "qualificacao|Qualificação|L|Qualificacao|3. Prospecção|0|1"
    Add c, "ctx_resumo|Contexto|R||4. Contexto|0|0"
    Add c, "ctx_familia|Família aderente|R||4. Contexto|0|0"
    Add c, "ctx_atualizado_em|Atualizado em|R||4. Contexto|0|0"
    Add c, "observacoes|Observações|M||5. Observações|0|1"

    Case "contatos"
    Add c, "codigo|Código|R||1. Identificação|0|0"
    Add c, "id_cliente|Empresa|K||1. Identificação|1|1"
    Add c, "nome|Contato|T||1. Identificação|1|1"
    Add c, "cargo|Cargo|T||1. Identificação|0|1"
    Add c, "telefone|Telefone|T||2. Comunicação|0|1"
    Add c, "email|E-mail|T||2. Comunicação|0|1"
    Add c, "observacoes|Observações|M||3. Observações|0|1"

    Case "oportunidades"
    Add c, "codigo|Atendimento|R||1. Cliente e contato|0|0"
    Add c, "id_contato|Contato|K||1. Cliente e contato|1|1"
    Add c, "empresa|Empresa|R||1. Cliente e contato|0|0"
    Add c, "cidade|Cidade|R||1. Cliente e contato|0|0"
    Add c, "uf|UF|R||1. Cliente e contato|0|0"
    Add c, "cargo|Cargo|R||1. Cliente e contato|0|0"
    Add c, "telefone|Telefone|R||1. Cliente e contato|0|0"
    Add c, "email|E-mail|R||1. Cliente e contato|0|0"
    Add c, "orcamento|Orçamento|T||2. Negociação|0|1"
    Add c, "etapa|Etapa|L|Etapa|2. Negociação|1|1"
    Add c, "responsavel|Responsável|L|Responsavel|2. Negociação|1|1"
    Add c, "maquina|Máquina / descrição|T||2. Negociação|0|1"
    Add c, "familia|Família de máquina|L|Familia|2. Negociação|0|1"
    Add c, "categoria|Categoria|L|Categoria|2. Negociação|0|1"
    Add c, "tipo_venda|Tipo de venda|T||2. Negociação|0|1"
    Add c, "valor|Valor (R$)|C||2. Negociação|0|1"
    Add c, "origem|Origem|L|Origem|3. Acompanhamento|0|1"
    Add c, "prioridade|Prioridade|L|Prioridade|3. Acompanhamento|0|1"
    Add c, "dt_entrada|Data de entrada|D||3. Acompanhamento|0|1"
    Add c, "dt_proposta|Data da proposta|D||3. Acompanhamento|0|1"
    Add c, "ultima_interacao|Última interação|D||3. Acompanhamento|0|1"
    Add c, "tentativas|Tentativas|N||3. Acompanhamento|0|1"
    Add c, "prox_acao|Próxima ação|M||4. Próxima ação e desfecho|0|1"
    Add c, "dt_prox_acao|Data próx. ação|D||4. Próxima ação e desfecho|0|1"
    Add c, "retomar_em|Retomar em|D||4. Próxima ação e desfecho|0|1"
    Add c, "motivo_desfecho|Motivo de desfecho|L|Motivo|4. Próxima ação e desfecho|0|1"
    Add c, "dt_desfecho|Data do desfecho|D||4. Próxima ação e desfecho|0|1"
    Add c, "ctx_resumo|Contexto|R||5. Observações e contexto|0|0"
    Add c, "observacoes|Observações|M||5. Observações e contexto|0|1"

    End Select

    CamposDe = c
End Function

'----------------------------------------------------------
' Campos de VINCULO (tipo K) da tabela, na ordem da ficha,
' separados por virgula. O primeiro e o vinculo principal (o
' que o atendimento novo e o contato novo escolhem logo ao
' abrir). A ficha guarda um id POR CAMPO: um id unico para a
' tabela inteira gravaria o contato no lugar de outro vinculo.
'----------------------------------------------------------
Public Function CamposK(ByVal tabela As String) As String
    Dim campos As Variant, i As Long, spec As Variant
    campos = CamposDe(tabela)
    If IsEmpty(campos) Then Exit Function
    For i = LBound(campos) To UBound(campos)
        spec = Split(campos(i), "|")
        If spec(2) = "K" Then CamposK = CamposK & IIf(CamposK = "", "", ",") & spec(0)
    Next i
End Function

'----------------------------------------------------------
' Colunas da grade: expressao|Titulo|largura
'----------------------------------------------------------
Public Function ColunasDe(ByVal tabela As String) As Variant
    Dim c As Variant

    ' Lista do FORMULARIO. Duas regras aqui:
    '
    ' 1. A soma cabe em 478 pt - a ListBox tem 498 e a barra de
    '    rolagem come cerca de 16.
    ' 2. A largura segue o CONTEUDO, nao o titulo. "AT-0489-0491" sao
    '    12 caracteres e ocupam cerca de 62 pt em Segoe UI 9; dar 86
    '    deixava um vao entre Atendimento e Empresa maior que o vao
    '    entre as outras colunas, e era isso que parecia desalinhado.
    '
    ' Prox. acao saiu da lista de oportunidades: a coluna Situacao ja
    ' diz se esta atrasada, nesta semana ou em dia, que e a decisao
    ' que se toma olhando a lista. A data exata se le na ficha.

    Select Case tabela

    Case "clientes"
    Add c, "codigo_visual|C" & Chr$(243) & "digo|46"
    Add c, "estagio|Est" & Chr$(225) & "gio|62"
    Add c, "empresa|Empresa|196"
    Add c, "cidade|Cidade|94"
    Add c, "uf|UF|22"
    Add c, "responsavel|Resp.|56"

    Case "contatos"
    Add c, "codigo|C" & Chr$(243) & "digo|76"
    Add c, "empresa|Empresa|160"
    Add c, "nome|Contato|104"
    Add c, "cargo|Cargo|60"
    Add c, "telefone|Telefone|78"

    Case "oportunidades"
    Add c, "codigo|Atendimento|68"
    Add c, "empresa|Empresa|172"
    Add c, "etapa|Etapa|96"
    Add c, "valor|Valor|58"
    Add c, "situacao|Situa" & Chr$(231) & Chr$(227) & "o|84"

    End Select

    ColunasDe = c
End Function

'----------------------------------------------------------
Public Function SecoesDe(ByVal tabela As String) As Variant
    Dim campos As Variant, i As Long, spec As Variant
    Dim vistas As Object, saida() As String, n As Long
    Set vistas = CreateObject("Scripting.Dictionary")
    campos = CamposDe(tabela)
    If IsEmpty(campos) Then Exit Function
    ReDim saida(0 To UBound(campos) - LBound(campos))
    For i = LBound(campos) To UBound(campos)
        spec = Split(campos(i), "|")
        If Not vistas.Exists(spec(4)) Then
            vistas.Add spec(4), 1
            saida(n) = spec(4)
            n = n + 1
        End If
    Next i
    ReDim Preserve saida(0 To n - 1)
    SecoesDe = saida
End Function

'----------------------------------------------------------
' Colunas da GRADE DA PLANILHA (abas Clientes, Contatos e
' Oportunidades). Formato: campo|Titulo|largura|tipo
'   tipo: T texto  N numero  C moeda  D data  F telefone
'
' Dois conjuntos por tabela. O minimo e o que cabe na tela
' sem rolagem horizontal - e o que se usa o dia inteiro. O
' completo sai no botao "Ver tudo" e serve para conferencia,
' nao para trabalho continuo.
'
' Texto longo (observacoes, proxima acao, categoria) NAO entra
' na grade nem no Ver tudo: uma celula com 240 caracteres
' empurra todas as outras colunas e a tabela deixa de ser
' legivel. Esse conteudo se le na ficha, inteiro.
'
' Campo aqui tem de existir na consulta correspondente de
' modCRM (SQLClientes, SQLContatos, SQLOportunidades), fora
' os dois calculados: codigo_visual e situacao.
'----------------------------------------------------------
Public Function ColunasPlanilha(ByVal tabela As String, _
                                ByVal completo As Boolean) As Variant
    Dim c As Variant

    Select Case tabela

    ' Ordem das colunas: o que se FILTRA vem primeiro, como na 3.6.
    ' Codigo, estagio e responsavel a esquerda; nome e endereco depois.
    ' Mudar a ordem aqui muda a grade - e so mover a linha.
    Case "clientes"
    Add c, "codigo_visual|C" & Chr$(243) & "digo|10|T"
    Add c, "estagio|Est" & Chr$(225) & "gio|13|T"
    Add c, "responsavel|Resp.|13|T"
    Add c, "empresa|Empresa|50|T"
    Add c, "cidade|Cidade|22|T"
    Add c, "uf|UF|5|T"
    If completo Then
        Add c, "cnpj|CNPJ|20|J"
        Add c, "segmento|Segmento|22|T"
        Add c, "qualificacao|Qualifica" & Chr$(231) & Chr$(227) & "o|16|T"
        Add c, "aderencia|Ader.|7|N"
        Add c, "porte|Porte|7|N"
        Add c, "telefone|Telefone|16|F"
        Add c, "email|E-mail|28|T"
        Add c, "ctx_familia|Fam" & Chr$(237) & "lia aderente|24|T"
        Add c, "ctx_atualizado_em|Contexto em|13|D"
    End If

    Case "contatos"
    Add c, "codigo|C" & Chr$(243) & "digo|13|T"
    Add c, "empresa|Empresa|46|T"
    Add c, "nome|Contato|26|T"
    Add c, "cargo|Cargo|20|T"
    Add c, "telefone|Telefone|16|F"
    If completo Then
        Add c, "email|E-mail|30|T"
        Add c, "cidade|Cidade|22|T"
        Add c, "uf|UF|5|T"
    End If

    ' Mesma logica da 3.6: atendimento, situacao, etapa e responsavel
    ' a esquerda, porque e por eles que se filtra o dia inteiro.
    Case "oportunidades"
    Add c, "codigo|Atendimento|14|T"
    Add c, "situacao|Situa" & Chr$(231) & Chr$(227) & "o|18|T"
    Add c, "etapa|Etapa|20|T"
    Add c, "responsavel|Resp.|13|T"
    Add c, "empresa|Empresa|42|T"
    Add c, "contato|Contato|24|T"
    Add c, "valor|Valor (R$)|15|C"
    Add c, "dt_prox_acao|Pr" & Chr$(243) & "x. a" & Chr$(231) & Chr$(227) & "o|13|D"
    If completo Then
        Add c, "cidade|Cidade|22|T"
        Add c, "uf|UF|5|T"
        Add c, "orcamento|Or" & Chr$(231) & "amento|14|T"
        Add c, "maquina|M" & Chr$(225) & "quina|34|T"
        Add c, "familia|Fam" & Chr$(237) & "lia|20|T"
        Add c, "tipo_venda|Tipo de venda|16|T"
        Add c, "origem|Origem|16|T"
        Add c, "prioridade|Prioridade|13|T"
        Add c, "dt_entrada|Entrada|12|D"
        Add c, "dt_proposta|Proposta|12|D"
        Add c, "ultima_interacao|" & Chr$(218) & "lt. intera" & Chr$(231) & Chr$(227) & "o|13|D"
        Add c, "tentativas|Tent.|7|N"
        Add c, "retomar_em|Retomar em|13|D"
        Add c, "motivo_desfecho|Motivo|22|T"
        Add c, "dt_desfecho|Desfecho|12|D"
    End If

    End Select

    ColunasPlanilha = c
End Function

Public Function Parte(ByVal spec As String, ByVal indice As Long) As String
    Dim p As Variant
    p = Split(spec, "|")
    If indice <= UBound(p) Then Parte = p(indice)
End Function
