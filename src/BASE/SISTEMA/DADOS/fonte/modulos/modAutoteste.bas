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
    Confere "VERSAO_ESQUEMA", modConfig.VERSAO_ESQUEMA, "1.1"

    ' ---- modDB (so utilitarios)
    Confere "Nz nulo", modDB.Nz(Null, "x"), "x"
    Confere "Nz valor", modDB.Nz(12), "12"

    ' ---- modValidacao
    Confere "Normalizar empresa", modValidacao.Normalizar("empresa", "  ind  d'angelo  "), "IND D'ANGELO"
    ' ---- chave de lista (C01): mesma regra do conector (Regras.ps1)
    Confere "ChaveLista C01 caixa e espaco", modValidacao.ChaveLista("pedido  fechado "), "PEDIDO FECHADO"
    Confere "ChaveLista C01 acento", modValidacao.ChaveLista("Ped" & ChrW(237) & "do Fechado"), "PEDIDO FECHADO"
    Confere "ChaveLista C01 til e cedilha", modValidacao.ChaveLista("Negocia" & ChrW(231) & ChrW(227) & "o " & ChrW(199) & " " & ChrW(209)), "NEGOCIACAO C N"
    Confere "ChaveLista C01 vazio", modValidacao.ChaveLista("   "), ""
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
    Confere "Situacao encerrado", modCalc.SituacaoDaEtapa(True, Null, Null), "Encerrado"
    Confere "Situacao sem etapa", modCalc.Situacao(True, "", Null, Null), "Sem etapa definida"
    Confere "Situacao atrasada", modCalc.SituacaoDaEtapa(False, Date - 1, Null), _
            "A" & ChrW$(&HE7) & ChrW$(&HE3) & "o atrasada"
    Confere "Situacao retomar", modCalc.SituacaoDaEtapa(False, Null, Date), "Retomar hoje"
    Confere "Situacao sem acao", modCalc.SituacaoDaEtapa(False, Null, Null), _
            "Sem pr" & ChrW$(&HF3) & "xima a" & ChrW$(&HE7) & ChrW$(&HE3) & "o"
    Confere "Situacao em dia", modCalc.SituacaoDaEtapa(False, Date + 30, Null), "Em dia"
    Confere "Quadro prospeccao", modCalc.QuadroDoCodigo("PROSPECCAO"), "1. Prospec" & ChrW$(&HE7) & ChrW$(&HE3) & "o"
    Confere "Quadro funil", modCalc.QuadroDoCodigo("FUNIL"), "2. Funil comercial"
    Confere "C06 valor inativo gravado: sem a marca", modCRM.ValorDaLista("Perdido (inativo)"), "Perdido"
    Confere "C06 valor ativo: nada a tirar", modCRM.ValorDaLista("Perdido"), "Perdido"
    ' ---- regras de lista da Fase 2 (C03, C04, C05): mesma regra do conector (modListas)
    Confere "C05 UF valida", modListas.UFValida("sp"), True
    Confere "C05 UF fora das 27", modListas.UFValida("XX"), False
    Confere "C05 UF com espaco sobrando", modListas.UFValida(" SP "), False
    Confere "C04 resultado com encerra Nao e recusado", modListas.PapelEtapaValido(False, "GANHO") <> "", True
    Confere "C04 PERDA com encerra Sim passa", modListas.PapelEtapaValido(True, "PERDA"), ""
    Confere "C04 sem resultado e sem encerra passa", modListas.PapelEtapaValido(False, ""), ""
    Confere "C04 resultado fora de GANHO/PERDA e recusado", modListas.PapelEtapaValido(True, "TALVEZ") <> "", True
    Confere "C03 desativar a unica GANHO e recusado", modListas.DesativarEtapaPermitido("Pedido Fechado", Array("PEDIDO FECHADO", "PERDIDO"), Array("GANHO", "PERDA"), Array(True, True)), False
    Confere "C03 uma de duas PERDA pode sair", modListas.DesativarEtapaPermitido("Perdido", Array("PERDIDO", "DESCARTADO"), Array("PERDA", "PERDA"), Array(True, True)), True
    Confere "C03 etapa sem resultado pode sair", modListas.DesativarEtapaPermitido("Negociacao", Array("NEGOCIACAO"), Array(""), Array(True)), True
    Confere "C03 a outra PERDA ja inativa: a ultima e recusada", modListas.DesativarEtapaPermitido("Perdido", Array("PERDIDO", "DESCARTADO"), Array("PERDA", "PERDA"), Array(True, False)), False
    ' ---- regras de visita (C08 a C13): mesma regra do conector (modVisita)
    Confere "C08 visita em encerrado e recusada", modVisita.VisitaPodeCriar(True) <> "", True
    Confere "C08 visita em aberto passa", modVisita.VisitaPodeCriar(False), ""
    Confere "C09 data fora do periodo e recusada", modVisita.DataNoPeriodo(DateSerial(2026, 10, 20), DateSerial(2026, 10, 1), DateSerial(2026, 10, 15)) <> "", True
    Confere "C09 data na borda passa", modVisita.DataNoPeriodo(DateSerial(2026, 10, 15), DateSerial(2026, 10, 1), DateSerial(2026, 10, 15)), ""
    Confere "C10 realizada sem relato e recusada", modVisita.RealizadaValida("", DateSerial(2026, 10, 14)) <> "", True
    Confere "C10 realizada completa passa", modVisita.RealizadaValida("Visitou", DateSerial(2026, 10, 14)), ""
    Confere "C11 cancelada sem motivo e recusada", modVisita.CanceladaValida("") <> "", True
    Confere "C12 texto do historico", modVisita.TextoHistorico(DateSerial(2026, 10, 14) + TimeSerial(9, 30, 0), "Visita", "Positivo", "Visitou"), "VISITA 14/10/2026 09:30 - Visita - Positivo: Visitou"
    Confere "C13 segunda pendente e recusada", modVisita.SegundaPendente(1) <> "", True
    Confere "C13 sem pendente passa", modVisita.SegundaPendente(0), ""
    ' C02 compara pela chave: variante com acento e o mesmo valor; C07 etapa que encerra = Encerrado
    Confere "C02 variante com acento tem a mesma chave", modValidacao.ChaveLista("Ped" & ChrW(237) & "do Fechado") = modValidacao.ChaveLista("Pedido Fechado"), True
    Confere "C07 etapa que encerra fica Encerrado", modCalc.SituacaoDaEtapa(True, DateSerial(2099, 1, 1), Null), "Encerrado"
    ' data digitada nas telas: dia que nao existe e recusado (como o ParseExact do conector)
    Confere "data 2026-02-30 nao existe", modValidacao.DataExiste(2026, 2, 30), False
    Confere "data 2026-10-32 nao existe", modValidacao.DataExiste(2026, 10, 32), False
    Confere "data 2026-13-01 nao existe", modValidacao.DataExiste(2026, 13, 1), False
    Confere "data 2028-02-29 existe", modValidacao.DataExiste(2028, 2, 29), True
    ' lancamento: pendente com data = agendar (data e hora agendadas); realizada = data realizada
    Dim lc As Variant
    lc = modVisita.CamposLancamento("PENDENTE", DateSerial(2026, 10, 25) + TimeSerial(9, 0, 0), "09:00", Null, "")
    Confere "lancamento PENDENTE: a data vai para a agendada", lc(0), DateSerial(2026, 10, 25)
    Confere "lancamento PENDENTE: a hora vai para a agendada", lc(1), "09:00"
    Confere "lancamento PENDENTE: sem data realizada", IsNull(lc(2)), True
    lc = modVisita.CamposLancamento("REALIZADA", DateSerial(2026, 10, 25) + TimeSerial(9, 30, 0), "09:30", DateSerial(2026, 10, 24), "08:00")
    Confere "lancamento REALIZADA: a data vai para a realizada", lc(2), DateSerial(2026, 10, 25) + TimeSerial(9, 30, 0)
    Confere "lancamento REALIZADA: a agendada fica como estava", lc(0), DateSerial(2026, 10, 24)
    ' ---- planejamento (Fase 4): mesma regra do conector (modPlanejamento)
    Confere "nome_busca sem acento e maiusculas", modPlanejamento.NomeBusca("S" & ChrW(227) & "o Lu" & ChrW(237) & "s"), "SAO LUIS"
    Confere "distancia mesma cidade e zero", modPlanejamento.DistanciaKm(-23.5505, -46.6333, -23.5505, -46.6333), 0
    Confere "distancia SP a RJ entre 340 e 370 km", (modPlanejamento.DistanciaKm(-23.5505, -46.6333, -22.9068, -43.1729) > 340) And (modPlanejamento.DistanciaKm(-23.5505, -46.6333, -22.9068, -43.1729) < 370), True
    Confere "C14 raio zero e recusado", modPlanejamento.RaioValido(0) <> "", True
    Confere "C14 raio 501 e recusado", modPlanejamento.RaioValido(501) <> "", True
    Confere "C14 raio 500 passa", modPlanejamento.RaioValido(500), ""
    Confere "C15 inicio depois do fim e recusado", modPlanejamento.PeriodoValido("2026-10-20", "2026-10-10") <> "", True
    Confere "C15 periodo de um dia passa", modPlanejamento.PeriodoValido("2026-10-10", "2026-10-10"), ""
    Confere "C16 encerrar com pendentes sem cancelar e recusado", modPlanejamento.EncerramentoPermitido(2, False) <> "", True
    Confere "C16 encerrar cancelando pendentes passa", modPlanejamento.EncerramentoPermitido(2, True), ""
    Confere "P4 codigo PL-2026-007", modPlanejamento.CodigoPlanejamento(2026, 7), "PL-2026-007"
    Confere "Dias parado", modCalc.DiasParado(Date - 10), 10
    Confere "Ciclo", modCalc.Ciclo(DateSerial(2026, 1, 1), DateSerial(2026, 1, 31)), 30
    Confere "Mes referencia", modCalc.MesReferencia(DateSerial(2026, 9, 19)), 202609

    ' ---- modCRM (sem banco)
    Confere "MontarCodigo", modCRM.MontarCodigo("AT", 11, 148), "AT-0011-0148"
    Confere "CodigoVisualCliente", modCRM.CodigoVisualCliente(7), "0007"
    Confere "CodigoVisualCliente nulo", modCRM.CodigoVisualCliente(Null), ""
    ' item 3: filtro de situacao do cadastro (o SQL e so texto: sem banco)
    Dim ps As Variant
    Confere "SQL clientes ativos", (InStr(modCRM.SQLClientes("", "", "", ps, situacao:="ativos"), "cl.ativo=True") > 0), True
    Confere "SQL clientes inativos", (InStr(modCRM.SQLClientes("", "", "", ps, situacao:="inativos"), "cl.ativo=False") > 0), True
    Confere "SQL clientes todos", (InStr(modCRM.SQLClientes("", "", "", ps, situacao:="todos"), "cl.ativo=") = 0), True
    Confere "SQL contatos ativos", (InStr(modCRM.SQLContatos("", ps, situacao:="ativos"), "ct.ativo=True") > 0), True
    Confere "SQL contatos inativos", (InStr(modCRM.SQLContatos("", ps, situacao:="inativos"), "ct.ativo=False") > 0), True
    ' combo da ficha e lista da aba: as mesmas opcoes, Ativos primeiro (padrao)
    Confere "OpcoesSituacao", Join(modCRM.OpcoesSituacao(), "|"), "Ativos|Inativos|Todos"
    Confere "ChaveSituacao Inativos", modCRM.ChaveSituacao("Inativos"), "inativos"
    Confere "ChaveSituacao vazio", modCRM.ChaveSituacao(""), "todos"
    ' item 6: o atendimento so troca de contato dentro da mesma empresa
    ' (o codigo AT- e congelado e nomeia a pasta do cliente)
    Confere "troca de contato na mesma empresa", modCRM.CriticarTrocaContato(False, 3, 3), ""
    Confere "troca de contato para outra empresa", (modCRM.CriticarTrocaContato(False, 3, 4) <> ""), True
    Confere "atendimento novo nao confere empresa", modCRM.CriticarTrocaContato(True, 0, 4), ""
    ' gravado sem empresa conhecida (0): recusa, nunca "qualquer empresa" - era o
    ' furo do atendimento novo que nao entrava na lista e era editado em seguida
    Confere "atendimento gravado sem empresa recusa", (modCRM.CriticarTrocaContato(False, 0, 4) <> ""), True
    ' item 4: o anterior nao pode ser o proprio atendimento nem de outra empresa
    Confere "anterior vazio aceito", modCRM.CriticarAnterior(10, 0, 3, 0), ""
    Confere "anterior da mesma empresa aceito", modCRM.CriticarAnterior(10, 7, 3, 3), ""
    Confere "anterior no atendimento novo", modCRM.CriticarAnterior(0, 7, 3, 3), ""
    Confere "anterior = o proprio recusado", (modCRM.CriticarAnterior(10, 10, 3, 3) <> ""), True
    Confere "anterior de outra empresa recusado", (modCRM.CriticarAnterior(10, 7, 3, 4) <> ""), True
    ' anterior que nao existe mais (empresa 0): recusa, mas nao diz "outra empresa"
    Confere "anterior inexistente recusado", (modCRM.CriticarAnterior(10, 7, 3, 0) <> ""), True
    Confere "anterior inexistente sem 'empresa'", (InStr(modCRM.CriticarAnterior(10, 7, 3, 0), "empresa") = 0), True
    ' so se confere o anterior que MUDOU (ou no atendimento novo): o que a IA
    ' gravou antes da regra nao trava a edicao de outro campo
    Confere "anterior inalterado nao confere", modCRM.AnteriorPrecisaConferir(False, 7, 7), False
    Confere "anterior trocado confere", modCRM.AnteriorPrecisaConferir(False, 7, 8), True
    Confere "anterior no novo confere", modCRM.AnteriorPrecisaConferir(True, 0, 8), True
    Confere "anterior removido nao confere", modCRM.AnteriorPrecisaConferir(False, 7, 0), False
    ' botao de cada campo de vinculo: so na edicao; o do contato so em
    ' atendimento ja gravado (o novo escolhe empresa e contato em 2 passos)
    Confere "botao contato na edicao", modCRM.BotaoVinculoAtivo("id_contato", True, False, 5), True
    Confere "botao contato no atendimento novo", modCRM.BotaoVinculoAtivo("id_contato", True, True, 0), False
    Confere "botao contato em leitura", modCRM.BotaoVinculoAtivo("id_contato", False, False, 5), False
    Confere "botao empresa do contato novo", modCRM.BotaoVinculoAtivo("id_cliente", True, True, 0), True
    ' o anterior e da empresa do contato: sem contato nao ha o que listar
    Confere "botao anterior sem contato", modCRM.BotaoVinculoAtivo("id_atendimento_anterior", True, True, 0), False
    Confere "botao anterior com contato", modCRM.BotaoVinculoAtivo("id_atendimento_anterior", True, True, 5), True
    ' troca de contato numa empresa so com contatos INATIVOS: a lista (so
    ' ativos) vem vazia, mas a mensagem nao pode dizer "nao tem contato"
    Confere "sem contato nenhum", (InStr(modCRM.TextoSemContato("EMPRESA X", 0), "ainda n" & ChrW$(&HE3) & "o tem contato cadastrado") > 0), True
    Confere "so contatos inativos", (InStr(modCRM.TextoSemContato("EMPRESA X", 2), "2 contato(s) inativo(s)") > 0), True
    Confere "so inativos nao diz 'nao tem contato cadastrado'", (InStr(modCRM.TextoSemContato("EMPRESA X", 2), "tem contato cadastrado") = 0), True
    ' D20: telefone da grade de Oportunidades - o do contato; sem ele, o geral
    ' da empresa; sem os dois, vazio. Sempre formatado.
    Confere "telefone do contato", modCRM.TelefoneDoAtendimento("41999990000", "4133330000"), "(41) 99999-0000"
    Confere "contato sem telefone usa o da empresa", modCRM.TelefoneDoAtendimento(Null, "4133330000"), "(41) 3333-0000"
    Confere "contato com telefone em branco usa o da empresa", modCRM.TelefoneDoAtendimento("  ", "4133330000"), "(41) 3333-0000"
    Confere "sem telefone nenhum fica vazio", modCRM.TelefoneDoAtendimento(Null, Null), ""

    ' ---- modSchema
    Confere "CamposDe clientes", (UBound(modSchema.CamposDe("clientes")) >= 10), True
    Confere "CamposDe oportunidades", (UBound(modSchema.CamposDe("oportunidades")) >= 20), True
    Confere "ColunasPlanilha", (UBound(modSchema.ColunasPlanilha("oportunidades", False)) >= 5), True
    Confere "Parte", modSchema.Parte("a|b|c", 1), "b"
    ' D20: Telefone logo depois de Contato, na grade minima e no Ver tudo
    Confere "Telefone depois de Contato (minima)", (InStr(CamposDaGrade("oportunidades", False), ",contato:Contato,telefone_atendimento:Telefone,") > 0), True
    Confere "Telefone depois de Contato (Ver tudo)", (InStr(CamposDaGrade("oportunidades", True), ",contato:Contato,telefone_atendimento:Telefone,") > 0), True
    ' a ficha guarda um vinculo POR CAMPO K: com dois campos K na mesma
    ' tabela, um id unico gravaria o contato no lugar do outro vinculo
    Confere "CamposK oportunidades", modSchema.CamposK("oportunidades"), "id_contato,id_atendimento_anterior"
    Confere "CamposK contatos", modSchema.CamposK("contatos"), "id_cliente"
    Confere "CamposK clientes", modSchema.CamposK("clientes"), ""

    ' ---- modGrade / modAcoes / modLote
    Confere "TabelaDaAba", modGrade.TabelaDaAba(ThisWorkbook.Worksheets("Clientes")), "clientes"
    Confere "TabelaDaAba fora", modGrade.TabelaDaAba(ThisWorkbook.Worksheets("Inicio")), ""
    Confere "RotuloStatus", modAcoes.RotuloStatus("contatos", 0, True), "Desativar"
    ' item 5 (06/10/2026): SUSPEITO comeca fora; OK e SUSPEITO alternam
    Confere "lote: SUSPEITO comeca ignorado", modLote.ComecaIgnorada("SUSPEITO"), True
    Confere "lote: OK comeca entrando", modLote.ComecaIgnorada("OK"), False
    Confere "lote: OK alterna", modLote.PodeAlternar("OK"), True
    Confere "lote: SUSPEITO alterna", modLote.PodeAlternar("SUSPEITO"), True
    Confere "lote: DUPLICADO nao alterna", modLote.PodeAlternar("DUPLICADO"), False
    Confere "lote: ERRO nao alterna", modLote.PodeAlternar("ERRO"), False
    Confere "lote: rotulo SUSPEITO incluido", modLote.RotuloLinha("SUSPEITO", False), "SUSPEITO - ENTRA"
    Confere "lote: rotulo SUSPEITO fora", modLote.RotuloLinha("SUSPEITO", True), "SUSPEITO - FORA"
    Confere "lote: rotulo OK ignorado", modLote.RotuloLinha("OK", True), "OK - IGNORAR"
    Confere "lote: rotulo OK", modLote.RotuloLinha("OK", False), "OK"
    Confere "lote: rotulo DUPLICADO", modLote.RotuloLinha("DUPLICADO", False), "DUPLICADO"
    ' DUPLICADO no lote que aponta para uma SUSPEITO (comeca fora) diz
    ' que e preciso incluir a outra; apontando para OK, so o numero
    Confere "lote: repetido de OK", modLote.MotivoRepetidoNoLote("3", "OK"), "CNPJ repetido na linha 3 deste lote"
    Confere "lote: repetido de SUSPEITO pede incluir", (InStr(modLote.MotivoRepetidoNoLote("3", "SUSPEITO"), "inclua a linha 3") > 0), True
    Confere "lote: ajuda explica o duplo clique", (InStr(modLote.TextoAjuda(), "Suspeito come" & ChrW$(&HE7) & "a fora") > 0), True
    Confere "lote: aviso de quem nao alterna", (Left$(modLote.AvisoNaoAlterna(), 25) = "S" & ChrW$(&HF3) & " OK e SUSPEITO alternam"), True
    ' o selo da grade pinta pelo estado antes do " - " (SUSPEITO - FORA = SUSPEITO)
    Dim g As New clsGrade
    Confere "selo pelo prefixo", g.ChaveSelo("SUSPEITO - FORA"), "SUSPEITO"
    Confere "selo sem prefixo", g.ChaveSelo("Em dia"), "Em dia"
    ' linha fora do lote pinta como IGNORAR (cinza), nao como OK (verde)
    Confere "selo OK ignorado", g.ChaveSelo("OK - IGNORAR"), "IGNORAR"
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
    Dim fv As frmVinculo, fl As frmListas, ui As clsUI
    If False Then
        Set fv = New frmVinculo
        Set fl = New frmListas
        Call fv.EscolherContatoDe(0)
        Call fv.EscolherAtendimentoDe(0, 0)
        ' botao criado em tempo de execucao (campos de vinculo da ficha)
        Set ui = New clsUI
        ui.LigarBotao Nothing, Nothing, Nothing, "neutro"
    End If
End Sub

' ",campo:Titulo,campo:Titulo," das colunas da grade, na ordem.
Private Function CamposDaGrade(ByVal tabela As String, ByVal completo As Boolean) As String
    Dim c As Variant, s As String
    For Each c In modSchema.ColunasPlanilha(tabela, completo)
        s = s & "," & modSchema.Parte(CStr(c), 0) & ":" & modSchema.Parte(CStr(c), 1)
    Next c
    CamposDaGrade = s & ","
End Function

' CStr(Null) e erro: o relato da falha nao pode quebrar o autoteste.
Private Function Texto(ByVal v As Variant) As String
    If IsNull(v) Then Texto = "(nulo)" Else Texto = CStr(v)
End Function
