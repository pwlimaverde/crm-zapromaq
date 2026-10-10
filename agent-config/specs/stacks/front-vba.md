# Stack: front — Excel 2019 / VBA 7 / MSForms

## Papel

A planilha `CRM_Zapromaq.xlsm` que cada estação usa (cópia local em `Documentos\CRM Zapromaq`):
abas de grade (Clientes, Contatos, Oportunidades), Painel, ficha (`frmCRM`), cadastro em lote,
vínculo. O `.xlsm` **não é fonte**: é montado do zero pelo build a partir de `fonte\`.

## Versões e ambiente

- Excel **2019** 64 bits nas estações (VBA 7, `LongPtr`/`PtrSafe` se houver `Declare`).
  Esta máquina de desenvolvimento tem Excel 365 — compila aqui não garante compila lá.
- Tela-alvo 1366×768 com escala variável; projeta-se a 100% e `modTela.Encaixar` ajusta o `Zoom`.
- Sem referências adicionais: ADO em *late binding* (`CreateObject("ADODB.Connection")`).

## Onde fica

```
src/BASE/SISTEMA/DADOS/fonte/
  modulos/*.bas        modConfig, modDB, modSchema, modCRM, modCalc, modValidacao, modGrade,
                       modPainel, modRelatorio, modAcoes, modLote, modMenu, modFeed, modTela,
                       modAutoteste, modTema (GERADO)
  classes/*.cls        clsUI, clsBotao, clsGrade, clsLinhaGrade
  formularios/*.txt    código de frmCRM, frmLote, frmVinculo, frmCalibracao
  pasta-de-trabalho/ThisWorkbook.txt
  layout/              (ver visual-layout.md)
```

Papéis: `modDB` é o **único** acesso a dados; `modSchema` gera a ficha e as colunas;
`modCRM` regras e SQL das listagens; `modCalc` campos calculados; `modValidacao` normalização;
`modGrade` abas de grade; `modPainel` agregações; `modFeed` atualização a cada 25 s.

## Comandos

```powershell
python src\BASE\SISTEMA\DADOS\ferramentas\verificacao\verificar.py        # sintaxe (ANTLR), encoding, chamadas
src\BASE\SISTEMA\MONTAR-FRONTEND.bat -Teste                                # compila + autoteste + teste de uso
powershell -File src\BASE\SISTEMA\DADOS\testes\diagnosticar-compilacao.ps1  # quando não compila
```

## Estilo

- **Windows-1252 + CRLF.** Acento fora da página → `ChrW$(&H....)`. Ler/gravar sempre em cp1252.
- `Option Explicit` em tudo; nomes em português; constantes no topo.
- Máximo de 25 continuações por instrução (por isso um `Add` por campo em `modSchema`).
- VBA não diferencia maiúsculas: parâmetro/local com nome de rotina do módulo quebra a chamada
  (o `verificar.py` acusa).
- Cor do Excel = `R + G*256 + B*65536`. Paleta: azul `#172A67`, azul 2 `#153F71`, verde `#58B030`
  (vem de `modTema`, gerado do `tema.json`).
- Componentes de tela: botão = três Labels (`bg_<nome>`, `<nome>`, `ico_<nome>`) ligados por
  `Tag` (`ui:botao:<variante>`) via `clsUI`; depois de mudar `Enabled`, `mUI.Pintar`.
  `clsGrade` substitui a ListBox (`ColumnWidths`, `Cabecalho`, `Carregar`, `ListIndex`).
- Abas criadas por automação não têm CodeName: eventos em `ThisWorkbook` (`Workbook_Sheet*`).
- **Vínculos da ficha por campo:** cada campo K (`modSchema.CamposK`) guarda o próprio id
  (`frmCRM.mVinculos`, sincronizado por `DefinirValor`); nunca um id único por tabela.
  Cada campo K ganha, em tempo de execução, os botões `kb_<campo>` (trocar) e `kv_<campo>`
  (abrir a ficha em leitura), ligados por `clsUI.LigarBotao` (o clique no texto chega por
  `UI_Acao`, porque controle de `Controls.Add` não dispara o evento do formulário). Regra de
  quando ficam ativos: `modCRM.BotaoVinculoAtivo`. Atendimento gravado só troca de contato
  dentro da mesma empresa (`frmVinculo.EscolherContatoDe`, checagem ao salvar). A empresa do
  atendimento vem do **banco** no Editar (`mClienteRegistro`), nunca só da lista; empresa 0 em
  registro gravado é recusada (`CriticarTrocaContato`). O anterior só se confere quando muda
  ou no atendimento novo (`modCRM.AnteriorPrecisaConferir`).
- **Situação do cadastro** (Clientes/Contatos): grade com coluna oculta `_a` (depois de
  `_id`) e célula com lista em F4 (`modGrade.COL_SITUACAO`); a aba abre em Ativos, inativo em
  vermelho; valor fora da lista volta para Ativos; na ficha, `cboSituacao` vai no SQL
  (`modCRM.OpcoesSituacao`/`ChaveSituacao`).
- **Edição de fonte VBA com acento:** a ferramenta Edit regrava o arquivo inteiro como UTF-8
  e troca os acentos cp1252 por U+FFFD. Em arquivo com byte > 127 (ex.: `modCRM`, `frmCRM`,
  `modLote`), troque o trecho por script que lê e grava em cp1252, e confira com `git diff`.
- Nome de variável do formulário encobre função do VBA: em `frmCRM`, `Mid$` é `mID` — use
  `VBA.Mid$`.

## Testes

- **Unidade:** `modAutoteste.Autoteste` — `Confere "nome", obtido, esperado`. Sem banco, sem
  gravar na planilha. Roda no build depois de compilar; o build só publica com "OK".
  Na estação: `Alt+F11 › Ctrl+G › ?Autoteste("")`. Formulários são modais: a API pública deles
  entra em `ContratoFormularios` (chamada tipada dentro de `If False`: método inexistente =
  erro de compilação no build); o comportamento é item 🖥.
- **Uso:** `testes/testar-uso.ps1` monta as grades e o Painel com dados reais do banco
  apontado (no desenvolvimento, o banco de demonstração) e confere situação Ativos/Inativos/
  Todos, lote, `DadosDoContato` e a lista do atendimento anterior.
- **Visual:** prévia + `frmCalibracao` (Alt+F8 › `AbrirCalibracao`) na estação 🖥.

## Limites

- Sempre: toda gravação via `modDB` com log na mesma transação (o `modFeed` depende disso);
  uma conexão por operação do usuário; normalização só em `modValidacao`.
- Perguntar antes: mudar nome de controle (contrato com o layout JSON), o texto/célula de
  `Início!B7`, a proteção da cópia modelo.
- Nunca: recurso do Microsoft 365; OCX/DLL externos; URL; subclassing (`SetWindowsHookEx`);
  `LoadPicture` de disco em tempo de execução; célula mesclada na área de dados da grade;
  gravar `ctx_*` pelo formulário; editar `modTema.bas` à mão.

## Revisão

- [ ] Arquivo em cp1252 + CRLF; acentos certos depois de importar.
- [ ] Nenhuma consulta dentro de laço; `ScreenUpdating` desligado em lote.
- [ ] Toda chamada `modX.Rotina` existe; nada encobre nome de rotina.
- [ ] Nada do 365; `Declare` com `PtrSafe`.
- [ ] Erro tratado sem engolir (`On Error Resume Next` só em volta da linha que precisa).
- [ ] Nome de controle citado no código existe no layout (o `verificar.py` confere).

## Referências offline

`agent-config/doc_dev/planejamento/referencias/` 01-vba-msforms, 02-vba-linguagem,
03-vbe-modelo-extensibilidade, 04-excel-modelo-objetos, 05-excel-ribbon-customui, 14-office-2019.
