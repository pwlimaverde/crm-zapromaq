<#
  testar-uso.ps1 - teste de fumaça do front montado, sem interação.

  Abre uma CÓPIA do .xlsm de teste (eventos desligados: o Workbook_Open com
  MsgBox não roda), monta as três grades e o Painel pelas próprias macros e
  confere o resultado: linhas carregadas, status, Painel preenchido. Exporta
  as abas em PDF (execucao\teste\uso-*.pdf) para conferência visual.
  Um vigia encerra o Excel se alguma caixa de mensagem travar a execução.

  Uso: testar-uso.ps1 [-Arquivo <xlsm>]   (padrão: execucao\teste\CRM_Zapromaq.xlsm)
#>
param([string]$Arquivo = '', [switch]$SemRelatorio)
$ErrorActionPreference = 'Stop'
$dados = Split-Path -Parent $PSScriptRoot
. (Join-Path $dados 'lib\Comum.ps1')
. (Join-Path $dados 'build\lib\Excel.ps1')
$raiz = Get-RaizBase $PSScriptRoot
if (-not $Arquivo) { $Arquivo = Join-Path $dados 'execucao\teste\CRM_Zapromaq.xlsm' }
$pasta = Get-PastaExecucao $raiz 'teste'
$copia = Join-Path $env:TEMP ('crm-uso-' + $PID + '.xlsm')
Copy-Item -LiteralPath $Arquivo -Destination $copia -Force

$falhas = 0
function Conferir([string]$nome, [bool]$ok, [string]$det = '') {
    if ($ok) { Write-Host ('  OK    ' + $nome) } else { $script:falhas++; Write-Host ('  FALHA ' + $nome + '  ' + $det) }
}

$xl = $null; $idExcel = 0; $sinal = Join-Path $env:TEMP ('crm-uso-' + $PID + '.ok')
# o teste da troca de contato põe um formulário temporário na cópia (modelo de objeto do VBA)
$vbomOriginal = Get-AcessoVBOM
if ($vbomOriginal -ne 1) { Set-AcessoVBOM 1 }
try {
    $xl = New-Excel ([ref]$idExcel)
    $xl.DisplayAlerts = $false
    $xl.EnableEvents = $false
    $vigia = Start-Vigia $idExcel 240 $sinal
    $wb = $xl.Workbooks.Open($copia)
    $n = $wb.Name
    foreach ($m in 'GradeClientes', 'GradeContatos', 'GradeOportunidades') {
        $ini = Get-Date
        $xl.Run("'" + $n + "'!modGrade." + $m)
        $ms = [int]((Get-Date) - $ini).TotalMilliseconds
        $aba = $m.Substring(5)
        $ws = $wb.Worksheets.Item($aba)
        $linhas = if ($ws.ListObjects.Count -gt 0 -and $ws.ListObjects.Item(1).DataBodyRange) { $ws.ListObjects.Item(1).DataBodyRange.Rows.Count } else { 0 }
        Conferir ($aba + ': ' + $linhas + ' linha(s) em ' + $ms + ' ms') ($linhas -gt 0) ([string]$ws.Range('B5').Value2)
        Conferir ($aba + ': aba protegida') ([bool]$ws.ProtectContents)
        $ws.ExportAsFixedFormat(0, (Join-Path $pasta ('uso-' + $aba + '.pdf')))
    }
    # item 3: Clientes e Contatos abrem em Ativos; Todos mostra os inativos (em vermelho);
    # Limpar filtros volta para Ativos. Célula de situação: F4 (modGrade.COL_SITUACAO).
    foreach ($aba in 'Clientes', 'Contatos') {
        $ws = $wb.Worksheets.Item($aba); $lo = $ws.ListObjects.Item(1)
        $total = $lo.DataBodyRange.Rows.Count
        $vis = [int]$xl.WorksheetFunction.Subtotal(103, $lo.ListColumns.Item('_id').DataBodyRange)
        Conferir ($aba + ': abre em Ativos (' + $vis + ' de ' + $total + ')') ($vis -lt $total -and [string]$ws.Range('F4').Value2 -eq 'Ativos') ('situação=' + [string]$ws.Range('F4').Value2 + '; banco de demonstração sem inativos? rode testes\criar-banco-demo.ps1 -Recriar')
        $ws.Range('F4').Value2 = 'Todos'
        $xl.Run("'" + $n + "'!modGrade.FiltrarSituacao", $ws)
        $vis = [int]$xl.WorksheetFunction.Subtotal(103, $lo.ListColumns.Item('_id').DataBodyRange)
        Conferir ($aba + ': Todos mostra os ' + $total) ($vis -eq $total) ([string]$vis)
        $regras = 0
        foreach ($fc in $lo.DataBodyRange.FormatConditions) { if ([string]$fc.Formula1 -like '*INATIVO*') { $regras++ } }
        Conferir ($aba + ': regra vermelha dos inativos') ($regras -ge 1)
        $ws.Activate()
        $xl.Run("'" + $n + "'!modGrade.GradeLimparFiltros")
        $vis = [int]$xl.WorksheetFunction.Subtotal(103, $lo.ListColumns.Item('_id').DataBodyRange)
        Conferir ($aba + ': Limpar filtros volta para Ativos') ($vis -lt $total -and [string]$ws.Range('F4').Value2 -eq 'Ativos') ([string]$vis)
        # célula apagada como o Delete faz (ClearContents na área da célula, aba protegida):
        # em Contatos a mesclagem F:G caía sobre a _b oculta e o Excel recusava
        $xl.EnableEvents = $false
        $apagou = $true
        try { $ws.Range('F4').MergeArea.ClearContents() | Out-Null } catch { $apagou = $false }
        Conferir ($aba + ': Delete na célula de situação') $apagou ('área=' + $ws.Range('F4').MergeArea.Address())
        $xl.Run("'" + $n + "'!modGrade.FiltrarSituacao", $ws)
        $vis = [int]$xl.WorksheetFunction.Subtotal(103, $lo.ListColumns.Item('_id').DataBodyRange)
        Conferir ($aba + ': célula vazia volta para Ativos') ($vis -lt $total -and [string]$ws.Range('F4').Value2 -eq 'Ativos') ('situação=[' + [string]$ws.Range('F4').Value2 + ']')
    }
    # as macros acima religam os eventos; o resto do teste roda como começou (sem eventos)
    $xl.EnableEvents = $false

    # D20: coluna Telefone na grade de Oportunidades, logo depois de Contato - o telefone
    # do contato; sem ele, o geral da empresa; sem os dois, vazio. Cada linha contra o
    # banco; o banco de demonstração tem ao menos um atendimento de cada caso.
    $lo = $wb.Worksheets.Item('Oportunidades').ListObjects.Item(1)
    $cabs = @(foreach ($c in $lo.ListColumns) { [string]$c.Name })
    $iCto = [array]::IndexOf($cabs, 'Contato')
    Conferir 'Oportunidades: Telefone logo depois de Contato' ($iCto -ge 0 -and $cabs[$iCto + 1] -eq 'Telefone') ($cabs -join '|')
    if ($cabs -contains 'Telefone') {
        function Fone($v) {
            $s = if ($null -eq $v -or $v -is [DBNull]) { '' } else { ([string]$v) -replace '\D', '' }
            if ($s.Length -eq 11) { return '(' + $s.Substring(0, 2) + ') ' + $s.Substring(2, 5) + '-' + $s.Substring(7, 4) }
            if ($s.Length -eq 10) { return '(' + $s.Substring(0, 2) + ') ' + $s.Substring(2, 4) + '-' + $s.Substring(6, 4) }
            return $s
        }
        $tb = $xl.Run("'" + $n + "'!modDB.Consultar", 'SELECT op.id, ct.telefone, cl.telefone FROM (oportunidades op LEFT JOIN contatos ct ON op.id_contato=ct.id) LEFT JOIN clientes cl ON op.id_cliente=cl.id')
        $esp = @{}; $casos = @{ contato = 0; empresa = 0; vazio = 0 }
        for ($k = 1; $k -le $tb.GetLength(0); $k++) {
            $fc = Fone $tb.GetValue($k, 2); $fe = Fone $tb.GetValue($k, 3)
            if ($fc) { $esp[[int]$tb.GetValue($k, 1)] = $fc; $casos.contato++ }
            elseif ($fe) { $esp[[int]$tb.GetValue($k, 1)] = $fe; $casos.empresa++ }
            else { $esp[[int]$tb.GetValue($k, 1)] = ''; $casos.vazio++ }
        }
        $vIds = $lo.ListColumns.Item('_id').DataBodyRange.Value2
        $vTel = $lo.ListColumns.Item('Telefone').DataBodyRange.Value2
        $erradas = @()
        for ($k = 1; $k -le $vIds.GetLength(0); $k++) {
            $id = [int]$vIds.GetValue($k, 1); $obt = [string]$vTel.GetValue($k, 1)
            if ($obt -ne $esp[$id]) { $erradas += ('id ' + $id + ': [' + $obt + '] esperado [' + $esp[$id] + ']') }
        }
        Conferir ('Oportunidades: Telefone = contato (' + $casos.contato + '), empresa (' + $casos.empresa + '), vazio (' + $casos.vazio + ')') ($erradas.Count -eq 0 -and $casos.empresa -ge 1 -and $casos.vazio -ge 1) (($erradas -join '; ') + ' (sem caso de empresa/vazio? rode testes\criar-banco-demo.ps1 -Recriar)')
    }

    # item 6: os dados do contato (troca na ficha) vêm numa consulta só
    $lo = $wb.Worksheets.Item('Contatos').ListObjects.Item(1)
    $idCto = [int]$lo.ListColumns.Item('_id').DataBodyRange.Cells.Item(1, 1).Value2
    $nomeCto = [string]$lo.ListColumns.Item('Contato').DataBodyRange.Cells.Item(1, 1).Value2
    $dc = $xl.Run("'" + $n + "'!modCRM.DadosDoContato", $idCto)
    Conferir ('DadosDoContato(' + $idCto + ') traz nome e empresa') ($null -ne $dc -and [string]$dc.Item('nome') -eq $nomeCto -and [int]$dc.Item('id_cliente') -gt 0) $nomeCto

    # item 4: lista do atendimento anterior = os da mesma empresa, mais recentes primeiro,
    # sem o próprio. No banco de demonstração o cliente 1 tem 3; o id 1 é o mais recente.
    $la = $xl.Run("'" + $n + "'!modCRM.ListarAtendimentosDoCliente", 1, 0, '')
    $qt = if ($la -is [array]) { $la.GetLength(0) } else { 0 }
    Conferir ('anteriores do cliente 1: ' + $qt + ', o mais recente primeiro') ($qt -ge 2 -and [int]$la.GetValue(1, 1) -eq 1)
    $lb = $xl.Run("'" + $n + "'!modCRM.ListarAtendimentosDoCliente", 1, 1, '')
    $ids = if ($lb -is [array]) { @(for ($k = 1; $k -le $lb.GetLength(0); $k++) { [int]$lb.GetValue($k, 1) }) } else { @() }
    Conferir 'anteriores sem o próprio atendimento' ($ids.Count -eq $qt - 1 -and $ids -notcontains 1) ($ids -join ',')

    # Troca de contato pelo lápis (atendimento gravado) numa empresa só com
    # contatos INATIVOS - o caminho do relato de 10/10/2026: a lista (só ativos)
    # vem vazia; o aviso tem de dizer quantos inativos há, e não "não tem contato";
    # "Sim" cadastra o contato na hora, e a escolha acontece ANTES do Show do
    # frmVinculo, com a ficha modal por cima: Me.Hide nele era o erro 402.
    # Nesta cópia: um formulário modal faz o papel da ficha e um timer do Windows
    # responde Sim ao aviso, preenche o nome na ficha nova e salva. O contato de
    # teste é apagado no fim (banco de demonstração volta como estava).
    $sem = $xl.Run("'" + $n + "'!modDB.Consultar", 'SELECT TOP 1 ct.id_cliente FROM contatos ct WHERE ct.ativo=False AND ct.id_cliente NOT IN (SELECT id_cliente FROM contatos WHERE ativo=True) ORDER BY ct.id_cliente')
    if ($sem -is [array]) {
        $idSem = [int]$sem.GetValue(1, 1)
        $lv = $xl.Run("'" + $n + "'!modCRM.ListarContatosParaVinculo", '', $idSem)
        Conferir ('troca de contato: empresa ' + $idSem + ' só com inativos tem a lista vazia') (-not ($lv -is [array]))
        $vbp = $wb.VBProject
        $mt = $vbp.VBComponents.Add(1); $mt.Name = 'modTesteUso'
        $mt.CodeModule.AddFromString(@'
Option Explicit
Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
Private Declare PtrSafe Function KillTimer Lib "user32" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr) As Long
Private Declare PtrSafe Function FindWindowW Lib "user32" (ByVal cls As LongPtr, ByVal titulo As LongPtr) As LongPtr
Private Declare PtrSafe Function SendMessageW Lib "user32" (ByVal hWnd As LongPtr, ByVal msg As Long, ByVal wp As LongPtr, ByVal lp As LongPtr) As LongPtr
Private Declare PtrSafe Function GetDlgItemTextW Lib "user32" (ByVal hDlg As LongPtr, ByVal id As Long, ByVal buf As LongPtr, ByVal tam As Long) As Long
Public gCliente As Long, gAviso As String, gResultado As String
Private mTimer As LongPtr, mTiques As Long, mOcupado As Boolean, mSalvou As Boolean

Public Function Rodar(ByVal idCliente As Long) As String
    Dim f As New frmTesteModal
    gCliente = idCliente
    mTimer = SetTimer(0, 0, 300, AddressOf Tique)
    f.Show
    Unload f
    KillTimer 0, mTimer
    Rodar = gResultado & "|aviso=" & Replace(gAviso, vbCrLf, " ")
End Function

Public Sub Tique(ByVal h As LongPtr, ByVal m As Long, ByVal id As LongPtr, ByVal t As Long)
    Dim dlg As LongPtr, f As Object, buf As String, n As Long
    On Error Resume Next
    If mOcupado Then Exit Sub
    mOcupado = True
    mTiques = mTiques + 1
    If mTiques > 150 Then KillTimer 0, mTimer
    dlg = FindWindowW(StrPtr("#32770"), StrPtr("CRM Zapromaq"))
    If dlg <> 0 Then
        buf = String$(1024, vbNullChar)
        n = GetDlgItemTextW(dlg, &HFFFF&, StrPtr(buf), 1024)
        If gAviso = "" Then gAviso = Left$(buf, n) Else gResultado = gResultado & "|outra caixa: " & Left$(buf, n)
        SendMessageW dlg, &H111, 6, 0
        SendMessageW dlg, &H111, 1, 0
    ElseIf Not mSalvou Then
        For Each f In VBA.UserForms
            If TypeName(f) = "frmCRM" And f.Visible Then
                f.Controls("c_nome").Value = "CONTATO TESTE USO"
                mSalvou = True
                f.UI_Acao "cmdSalvar"
            End If
        Next f
    End If
    mOcupado = False
End Sub
'@)
        $ft = $vbp.VBComponents.Add(3); $ft.Name = 'frmTesteModal'
        $ft.CodeModule.AddFromString(@'
Private Sub UserForm_Activate()
    Dim v As frmVinculo, ok As Boolean
    Set v = New frmVinculo
    On Error Resume Next
    ok = v.EscolherContatoDe(modTesteUso.gCliente)
    modTesteUso.gResultado = "erro=" & Err.Number & " " & Err.Description & "|ok=" & ok & "|id=" & v.IdEscolhido & modTesteUso.gResultado
    On Error GoTo 0
    Unload v
    Me.Hide
End Sub
'@)
        $r = [string]$xl.Run("'" + $n + "'!modTesteUso.Rodar", $idSem)
        Write-Host ('  troca de contato: ' + $r)
        Conferir 'troca de contato: aviso diz que os contatos são inativos' ($r -match 'contato\(s\) inativo\(s\)') $r
        Conferir 'troca de contato: cadastro na hora volta escolhido, sem erro 402' ($r -like 'erro=0 *' -and $r -match '\|ok=(True|Verdadeiro)\|id=[1-9]') $r
        # banco de demonstração volta como estava
        $ex = "'" + $n + "'!modDB.Executar"
        [void]$xl.Run($ex, "DELETE FROM log_alteracoes WHERE tabela='contatos' AND id_registro IN (SELECT id FROM contatos WHERE nome='CONTATO TESTE USO')")
        [void]$xl.Run($ex, "DELETE FROM contatos WHERE nome='CONTATO TESTE USO'")
    } else { Conferir 'banco de demonstração com empresa só de contatos inativos' $false 'rode testes\criar-banco-demo.ps1 -Recriar' }

    # Fase 2.2a: gravacao de listas (C02, C03, C05) pelo front, com volta ao estado original
    $tl = "'" + $n + "'!modListas."
    $ri = [string]$xl.Run($tl + 'IncluirValor', 'Familia', 'TESTE FASE 2')
    Conferir 'listas: incluir valor novo grava (vazio = ok)' ($ri -eq '') $ri
    $rd = [string]$xl.Run($tl + 'IncluirValor', 'Familia', 'teste  fase 2')
    Conferir 'listas C02: mesma chave e recusada' ($rd -match 'ja existe') $rd
    $ruf = [string]$xl.Run($tl + 'IncluirValor', 'UF', 'XX')
    Conferir 'listas C05: UF fora das 27 e recusada' ($ruf -match 'C05') $ruf
    $rf = [string]$xl.Run($tl + 'DefinirAtivo', 'Familia', 'TESTE FASE 2', $false)
    Conferir 'listas: desativar valor de teste grava' ($rf -eq '') $rf
    $rg = [string]$xl.Run($tl + 'DefinirAtivo', 'Familia', 'TESTE FASE 2', $true)
    Conferir 'listas: reativar valor de teste grava' ($rg -eq '') $rg
    $rc = [string]$xl.Run($tl + 'DefinirAtivo', 'Etapa', 'Pedido Fechado', $false)
    Conferir 'listas C03: desativar a unica etapa GANHO e recusado' ($rc -match 'C03') $rc
    $nAtivo = [int]$xl.Run("'" + $n + "'!modDB.ConsultarValor", "SELECT COUNT(*) FROM listas WHERE tipo='Etapa' AND valor='Pedido Fechado' AND ativo=True")
    Conferir 'listas C03: a etapa recusada continua ativa' ($nAtivo -eq 1) ([string]$nAtivo)
    # volta: apaga o valor de teste e o log dele (o banco de demonstracao volta como estava)
    [void]$xl.Run("'" + $n + "'!modDB.Executar", "DELETE FROM log_alteracoes WHERE tabela='listas' AND id_registro IN (SELECT id FROM listas WHERE valor='TESTE FASE 2')")
    [void]$xl.Run("'" + $n + "'!modDB.Executar", "DELETE FROM listas WHERE valor='TESTE FASE 2'")

    # Fase 2.2b: papel da etapa pelo front (C03, C04 e confirmacao de 'encerra'), com volta ao estado original
    $tp = "'" + $n + "'!modListas."
    $maxLogAntes = [int]$xl.Run("'" + $n + "'!modDB.ConsultarValor", "SELECT MAX(id) FROM log_alteracoes", 0)
    $rp1 = [string]$xl.Run($tp + 'DefinirPapel', 'Negociação', $true, '', 'FUNIL', $true, $false)
    Conferir 'papel: mudar encerra com atendimentos sem confirmar e recusado' ($rp1 -match 'Confirme') $rp1
    $rp2 = [string]$xl.Run($tp + 'DefinirPapel', 'Negociação', $false, 'GANHO', 'FUNIL', $true, $false)
    Conferir 'papel C04: resultado com encerra Nao e recusado' ($rp2 -match 'C04') $rp2
    $rp3 = [string]$xl.Run($tp + 'DefinirPapel', 'Pedido Fechado', $true, '', 'FUNIL', $false, $false)
    Conferir 'papel C03: tirar o resultado da unica etapa GANHO e recusado' ($rp3 -match 'C03') $rp3
    $rp4 = [string]$xl.Run($tp + 'DefinirPapel', 'Negociação', $false, '', 'PROSPECCAO', $true, $false)
    Conferir 'papel: mudar quadro grava (vazio = ok)' ($rp4 -eq '') $rp4
    $rp5 = [string]$xl.Run($tp + 'DefinirPapel', 'Negociação', $false, '', 'FUNIL', $true, $false)
    Conferir 'papel: volta o quadro para FUNIL' ($rp5 -eq '') $rp5
    $nLog = [int]$xl.Run("'" + $n + "'!modDB.ConsultarValor", "SELECT COUNT(*) FROM log_alteracoes WHERE tabela='listas' AND id > " + $maxLogAntes, 0)
    Conferir 'papel: a mudanca de quadro foi registrada no log (ida e volta)' ($nLog -ge 2) ([string]$nLog)
    # volta: o log criado pelo teste sai (o banco de demonstracao volta como estava)
    [void]$xl.Run("'" + $n + "'!modDB.Executar", "DELETE FROM log_alteracoes WHERE tabela='listas' AND id > " + $maxLogAntes)

    # Fase 2.3: outra estacao grava em listas; o cache so muda quando o feed processa o log
    $tf = "'" + $n + "'!"
    $mfd = $wb.VBProject.VBComponents.Add(1); $mfd.Name = 'modTesteFeed'
    $mfd.CodeModule.AddFromString(@'
Option Explicit
' grava como outra estacao (origem IA): um valor novo e o log, na mesma transacao
Public Function OutraEstacao() As Long
    Dim cn As Object, id As Long
    Set cn = modDB.AbrirTransacao()
    On Error GoTo Falha
    id = modDB.InserirEm(cn, "INSERT INTO listas (tipo, valor, ordem, ativo) VALUES (?,?,?,True)", _
                         Array(modDB.P(modDB.adVarWChar, "Familia"), modDB.P(modDB.adVarWChar, "TESTE FEED"), _
                               modDB.P(modDB.adInteger, 999)))
    modDB.ExecutarEm cn, "INSERT INTO log_alteracoes (tabela, id_registro, campo, valor_antigo, valor_novo, acao, usuario, quando, origem) VALUES (?,?,?,?,?,?,?,?,?)", _
        Array(modDB.P(modDB.adVarWChar, "listas"), modDB.P(modDB.adInteger, id), modDB.P(modDB.adVarWChar, "valor"), _
              modDB.P(modDB.adLongVarWChar, ""), modDB.P(modDB.adLongVarWChar, "Familia: TESTE FEED"), _
              modDB.P(modDB.adVarWChar, "INSERIR"), modDB.P(modDB.adVarWChar, "OUTRA ESTACAO"), _
              modDB.P(modDB.adDate, Now), modDB.P(modDB.adVarWChar, "IA"))
    modDB.Confirmar cn
    OutraEstacao = id
    Exit Function
Falha:
    modDB.Desfazer cn
    OutraEstacao = -1
End Function
'@)
    $maxLogF = [int]$xl.Run($tf + 'modDB.ConsultarValor', "SELECT MAX(id) FROM log_alteracoes", 0)
    [void]$xl.Run($tf + 'modFeed.ProcessarAgora')
    [void]$xl.Run($tf + 'modCRM.ObterLista', 'Familia')
    $idOutra = [int]$xl.Run($tf + 'modTesteFeed.OutraEstacao')
    Conferir 'feed: a outra estacao gravou o valor' ($idOutra -gt 0) ([string]$idOutra)
    $lstAntes = $xl.Run($tf + 'modCRM.ObterLista', 'Familia')
    $vAntes = $false
    for ($i = 1; $i -le $lstAntes.GetLength(0); $i++) { if ([string]$lstAntes.GetValue($i, 1) -eq 'TESTE FEED') { $vAntes = $true } }
    Conferir 'feed: antes de processar o log, o cache ainda nao mostra o valor' (-not $vAntes)
    [void]$xl.Run($tf + 'modFeed.ProcessarAgora')
    $lstDepois = $xl.Run($tf + 'modCRM.ObterLista', 'Familia')
    $vDepois = $false
    for ($i = 1; $i -le $lstDepois.GetLength(0); $i++) { if ([string]$lstDepois.GetValue($i, 1) -eq 'TESTE FEED') { $vDepois = $true } }
    Conferir 'feed: depois de processar o log, o cache mostra o valor da outra estacao' $vDepois
    # volta: apaga o valor e o log criados pelo teste
    [void]$xl.Run($tf + 'modDB.Executar', "DELETE FROM listas WHERE valor='TESTE FEED'")
    [void]$xl.Run($tf + 'modDB.Executar', "DELETE FROM log_alteracoes WHERE id > " + $maxLogF)
    [void]$xl.Run($tf + 'modCRM.LimparCacheListas')

    # Fase 3.2b: gravacao de visitas pelo front (C08, C10, C11, C13, V6, V7, conflito), com volta ao estado original
    $tv = "'" + $n + "'!"
    $idOpV = [int]$xl.Run($tv + 'modDB.ConsultarValor', "SELECT TOP 1 id FROM oportunidades WHERE etapa='Negociação' ORDER BY id", 0)
    $opAntV = $xl.Run($tv + 'modDB.Consultar', "SELECT observacoes, ultima_interacao, versao FROM oportunidades WHERE id=" + $idOpV)
    $maxLogV2 = [int]$xl.Run($tv + 'modDB.ConsultarValor', "SELECT MAX(id) FROM log_alteracoes", 0)
    $rv1 = [string]$xl.Run($tv + 'modVisita.CriarVisita', $idOpV, [datetime]'2026-10-20', '09:00', 'Visita', 'Vendedor A', 'teste uso visita')
    Conferir 'visita: criar pendente grava (vazio = ok)' ($rv1 -eq '') $rv1
    $rv2 = [string]$xl.Run($tv + 'modVisita.CriarVisita', $idOpV, [datetime]'2026-10-21', '10:00', 'Visita', 'Vendedor A', 'teste uso visita')
    Conferir 'visita C13: segunda pendente no mesmo atendimento e recusada' ($rv2 -match 'C13') $rv2
    $idVis = [int]$xl.Run($tv + 'modDB.ConsultarValor', "SELECT MAX(id) FROM visitas WHERE id_oportunidade=" + $idOpV, 0)
    $ra = [string]$xl.Run($tv + 'modVisita.AtualizarVisita', $idVis, 9, 'CANCELADA', [datetime]'2026-10-20', '09:00', 'Visita', 'Vendedor A', 'teste uso visita', $null, '', '', 'x')
    Conferir 'visita: versao lida antiga e recusada (conflito)' ($ra -match 'Conflito') $ra
    $rc = [string]$xl.Run($tv + 'modVisita.AtualizarVisita', $idVis, 1, 'CANCELADA', [datetime]'2026-10-20', '09:00', 'Visita', 'Vendedor A', 'teste uso visita', $null, '', '', '')
    Conferir 'visita C11: cancelar sem motivo e recusado' ($rc -match 'C11') $rc
    $rr = [string]$xl.Run($tv + 'modVisita.AtualizarVisita', $idVis, 1, 'REALIZADA', [datetime]'2026-10-20', '09:00', 'Visita', 'Vendedor A', 'teste uso visita', [datetime]'2026-10-20 09:30', 'Positivo', 'Visitou no front', '')
    Conferir 'visita: realizar grava (vazio = ok)' ($rr -eq '') $rr
    $histV = [string]$xl.Run($tv + 'modDB.ConsultarValor', "SELECT observacoes FROM oportunidades WHERE id=" + $idOpV, '')
    Conferir 'visita V6: histórico do atendimento ganhou VISITA 20/10/2026 09:30' ($histV -match 'VISITA 20/10/2026 09:30') $histV
    $ultV = $xl.Run($tv + 'modDB.ConsultarValor', "SELECT ultima_interacao FROM oportunidades WHERE id=" + $idOpV, $null)
    Conferir 'visita V7: última interação é a data realizada (20/10/2026 09:30)' (([datetime]$ultV).ToString('yyyy-MM-dd HH:mm') -eq '2026-10-20 09:30') ([string]$ultV)
    # volta: visita do teste, histórico e última interação do atendimento, e o log criado pelo teste
    [void]$xl.Run($tv + 'modDB.Executar', "DELETE FROM visitas WHERE id_oportunidade=" + $idOpV)
    $obsAnt = $opAntV.GetValue(1, 1); $ultAnt = $opAntV.GetValue(1, 2); $verAnt = [int]$opAntV.GetValue(1, 3)
    # parametros no mesmo formato do modDB.P (tipo ADO, valor): nulo continua nulo
    $pRest = [object[]]@( [object[]]@(203, $obsAnt), [object[]]@(7, $ultAnt), [object[]]@(3, $verAnt), [object[]]@(3, $idOpV) )
    [void]$xl.Run($tv + 'modDB.Executar', "UPDATE oportunidades SET observacoes=?, ultima_interacao=?, versao=? WHERE id=?", $pRest)
    [void]$xl.Run($tv + 'modDB.Executar', "DELETE FROM log_alteracoes WHERE id > " + $maxLogV2)

    # item 5: lote fictício com uma linha de cada estado (nada é gravado: só Analisar)
    #   1 OK; 2 SUSPEITO (nome de empresa da base, sem CNPJ); 3 DUPLICADO (CNPJ da linha 1); 4 ERRO (3 colunas)
    $t = "`t"
    $lote = @(('2' + $t + '2' + $t + $t + 'LOTE TESTE UM LTDA' + $t + 'CURITIBA' + $t + 'PR' + $t + $t + $t + $t + '11222333000181' + $t),
              ('2' + $t + '2' + $t + $t + 'METALURGICA EXEMPLO' + $t + 'CURITIBA' + $t + 'PR' + $t + $t + $t + $t + $t),
              ('1' + $t + '1' + $t + $t + 'LOTE TESTE DOIS LTDA' + $t + 'CURITIBA' + $t + 'PR' + $t + $t + $t + $t + '11222333000181' + $t),
              ('1' + $t + '1' + $t + 'SO TRES COLUNAS')) -join "`r`n"
    $rl = "'" + $n + "'!modLote."
    $q = [int]$xl.Run($rl + 'Analisar', $lote)
    # foreach, não ForEach-Object: o $_ do pipeline chega ao COM embrulhado (PSObject) e o Run trava
    $est = @(foreach ($i in 1..4) { [string]$xl.Run($rl + 'Valor', $i, 'status') }) -join ','
    Conferir ('lote: 4 linhas nos 4 estados (' + $est + ')') ($q -eq 4 -and $est -eq 'OK,SUSPEITO,DUPLICADO,ERRO') $est
    Conferir 'lote: total inicial conta só a OK' ([int]$xl.Run($rl + 'ACadastrar') -eq 1) ([string]$xl.Run($rl + 'ACadastrar'))
    $ok1 = [bool]$xl.Run($rl + 'AlternarIgnorar', 1)
    Conferir 'lote: duplo clique na OK tira a linha' ($ok1 -and [int]$xl.Run($rl + 'ACadastrar') -eq 0)
    $ok2 = [bool]$xl.Run($rl + 'AlternarIgnorar', 2)
    Conferir 'lote: duplo clique na SUSPEITO inclui' ($ok2 -and [int]$xl.Run($rl + 'ACadastrar') -eq 1 -and -not [bool]$xl.Run($rl + 'Ignorada', 2))
    Conferir 'lote: DUPLICADO e ERRO não alternam' (-not [bool]$xl.Run($rl + 'AlternarIgnorar', 3) -and -not [bool]$xl.Run($rl + 'AlternarIgnorar', 4))
    # DUPLICADO cuja primeira linha é SUSPEITO (começa fora): o motivo pede incluir a primeira
    #   1 SUSPEITO (telefone com 9 dígitos); 2 DUPLICADO (CNPJ da linha 1). Fictício, nada é gravado.
    $lote = @(('2' + $t + '2' + $t + $t + 'LOTE TESTE TRES LTDA' + $t + 'CURITIBA' + $t + 'PR' + $t + $t + '419999999' + $t + $t + '11222333000181' + $t),
              ('2' + $t + '2' + $t + $t + 'LOTE TESTE QUATRO LTDA' + $t + 'CURITIBA' + $t + 'PR' + $t + $t + $t + $t + '11222333000181' + $t)) -join "`r`n"
    [void]$xl.Run($rl + 'Analisar', $lote)
    $mot = [string]$xl.Run($rl + 'Valor', 2, 'motivo')
    Conferir 'lote: DUPLICADO de SUSPEITO pede incluir a linha 1' ([string]$xl.Run($rl + 'Valor', 1, 'status') -eq 'SUSPEITO' -and $mot -match 'inclua a linha 1') $mot

    $ini = Get-Date
    $xl.Run("'" + $n + "'!modPainel.AtualizarBasePainel")
    $ms = [int]((Get-Date) - $ini).TotalMilliseconds
    $p = $wb.Worksheets.Item('Painel')
    $tempos = [string]$xl.Run("'" + $n + "'!modPainel.TemposUltimaAtualizacao")
    Write-Host ('  tempos do Painel: ' + $tempos)
    # sem a marca final, o desenho quebrou no meio (e a caixa de erro prendeu o Excel)
    Conferir 'Painel desenhado até o fim' ($tempos -like '*paineis de baixo=*') $tempos
    Conferir 'Painel em menos de 5 s' ($ms -lt 5000) ($ms.ToString() + ' ms')
    Conferir ('Painel desenhado em ' + $ms + ' ms') ([string]$p.Range('B3').Value2 -like 'Atualizado em*') ([string]$p.Range('B3').Value2)
    Conferir ('Painel: cartão Empresas = ' + $p.Range('B6').Value2) ($p.Range('B6').Value2 -gt 0)
    Conferir ('Painel: ' + $p.ChartObjects().Count + ' gráfico(s)') ($p.ChartObjects().Count -ge 2)
    # "Ver tudo" com filtro e pesquisa ativos: as linhas ficam ocultas e a
    # grade ja saiu com o cabecalho da consulta anterior e colunas #N/D
    $ws = $wb.Worksheets.Item('Oportunidades')
    $lo = $ws.ListObjects.Item(1)
    $ws.Unprotect('zpm')
    $lo.Range.AutoFilter(3, [string]$lo.ListColumns.Item(3).DataBodyRange.Cells.Item(1, 1).Value2)
    $ws.Cells.Item(4, 3).Value2 = 'A'
    $ws.Protect('zpm')
    foreach ($passo in 'COMPLETO', 'MINIMO') {
        $xl.Run("'" + $n + "'!modGrade.GradeVerTudo")
        $cabs = @($wb.Worksheets.Item('Oportunidades').ListObjects.Item(1).ListColumns | ForEach-Object { $_.Name })
        $ruim = @($cabs | Where-Object { $_ -like '#N/D*' -or $_ -like 'Coluna*' -or $_ -like '*2' })
        Conferir ('Ver tudo (' + $passo + ') com filtro e pesquisa') ($ruim.Count -eq 0) ($cabs -join '|')
    }

    # relatorio: pasta com CSVs, PDF e LEIA-ME (no build, -SemRelatorio pula:
    # publicar nao pode criar pasta de relatorio em Documentos sem o usuario pedir)
    if (-not $SemRelatorio) {
    $pastaRel = [string]$xl.Run("'" + $n + "'!modRelatorio.GerarRelatorio", $true)
    foreach ($a in 'indicadores.csv', 'movimento_mensal.csv', 'atendimentos.csv', 'clientes.csv', 'painel.pdf', 'LEIA-ME.txt') {
        $arq = Join-Path $pastaRel $a
        Conferir ('relatório: ' + $a) ((Test-Path -LiteralPath $arq) -and (Get-Item -LiteralPath $arq).Length -gt 0) $arq
    }
    Write-Host ('  relatório em ' + $pastaRel)
    }
    $p.PageSetup.Orientation = 2; $p.PageSetup.Zoom = $false; $p.PageSetup.FitToPagesWide = 1; $p.PageSetup.FitToPagesTall = 1
    $p.ExportAsFixedFormat(0, (Join-Path $pasta 'uso-Painel.pdf'))
    $wb.Close($false)
} catch {
    $falhas++
    Write-Host ('ERRO: ' + $_.Exception.Message)
    Write-Host ('  onde: ' + $_.InvocationInfo.ScriptLineNumber + ': ' + $_.InvocationInfo.Line.Trim())
    if (-not (Get-Process -Id $idExcel -ErrorAction SilentlyContinue)) { Write-Host '  (o Excel travou numa caixa de mensagem e foi encerrado pelo vigia)' }
} finally {
    if ($vigia) { Stop-Vigia $vigia $sinal }
    Close-Excel $xl $idExcel
    if ($vbomOriginal -ne 1) { Set-AcessoVBOM $vbomOriginal }
    Remove-Item -LiteralPath $copia -Force -ErrorAction SilentlyContinue
}
Write-Host ''
if ($falhas -eq 0) { Write-Host 'TUDO CERTO. PDFs em execucao\teste\uso-*.pdf'; exit 0 }
Write-Host ($falhas.ToString() + ' falha(s).'); exit 1
