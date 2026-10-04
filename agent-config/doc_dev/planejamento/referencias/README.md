# Referências técnicas do projeto

Documentação oficial baixada em **2026-09-19** para consulta offline. Cada arquivo começa com um comentário com a URL de origem e a data.

- **Markdown original**: copiado do repositório público da documentação (Microsoft Docs, Model Context Protocol, Anthropic).
- **`.html.md`**: página oficial convertida de HTML para Markdown (quando não há repositório público). Tabelas e links podem sair simplificados; na dúvida, abra a URL de origem.

Atualizar (máquina de desenvolvimento, com internet): `python atualizar-referencias.py` — ou só algumas pastas: `python atualizar-referencias.py mcp vba`.
O resultado da última execução fica em `fontes-baixadas.json` (arquivos e falhas).

## Onde consultar cada assunto

| Pasta | Para que serve |
|---|---|
| [`01-vba-msforms`](01-vba-msforms/) | UserForm e controles MSForms: Zoom, Picture (formatos bmp/gif/jpg), BackStyle, BorderColor, SpecialEffect, eventos de mouse. Base dos componentes visuais (Fase 4/5). |
| [`02-vba-linguagem`](02-vba-linguagem/) | VBA: 64 bits, Declare/PtrSafe (API de DPI), RGB, Like, IIf, DateAdd/DateDiff. |
| [`03-vbe-modelo-extensibilidade`](03-vbe-modelo-extensibilidade/) | Modelo de extensibilidade do VBE (VBIDE) usado pelo build: Import, AddFromString, InsertLines etc.; habilitar acesso ao modelo de objeto do VBA. |
| [`04-excel-modelo-objetos`](04-excel-modelo-objetos/) | Excel: OnTime (feed de alterações), eventos da pasta de trabalho, SaveAs, proteção, formas, formatação condicional. |
| [`05-excel-ribbon-customui`](05-excel-ribbon-customui/) | Faixa de opções própria (customUI) gravada dentro do .xlsm (opcional, Fase 5). |
| [`06-access-sql-ace`](06-access-sql-ace/) | SQL do Access/ACE: DDL, relacionamentos, TOP (não desempata), LIKE, datas em formato EUA; página do Access Database Engine 2016 Redistributable. |
| [`07-ado`](07-ado/) | ADO usado pelo front VBA: Connection, Command, Parameter, CreateParameter, DataTypeEnum, transações. |
| [`08-dotnet-oledb`](08-dotnet-oledb/) | System.Data.OleDb usado pelos scripts PowerShell e pelo MCP: pool (OLE DB Services=-4), parâmetros, tipos, transações, @@IDENTITY. |
| [`09-powershell-5.1`](09-powershell-5.1/) | Windows PowerShell 5.1: JSON, codificação (BOM), política de execução, powershell.exe, escopos e blocos de script. |
| [`10-edge-headless`](10-edge-headless/) | Chromium/Edge headless: gerar imagens de fundo e exportar DOM a partir das páginas de prévia (Fase 4). |
| [`11-windows-icones-dpi`](11-windows-icones-dpi/) | Ícones Segoe MDL2 Assets / Fluent Icons (tabela de glifos) e APIs de DPI (GetDC, GetDeviceCaps, GetDpiForSystem). |
| [`12-mcp-especificacao`](12-mcp-especificacao/) | Especificação MCP: revisão atual 2026-07-28 (sem estado, server/discover, _meta, resultType, cache) e legado 2025-11-25 (initialize). Base do servidor dual-era. |
| [`13-claude-desktop`](13-claude-desktop/) | Claude Desktop: servidores MCP locais (claude_desktop_config.json, logs), extensões MCPB, Agent Skills (SKILL.md) e como enviar skills. |
| [`14-office-2019`](14-office-2019/) | Office 2019: ciclo de vida (suporte encerrado em 14/10/2025) e novidades do Excel 2019 (o que existe e o que não existe). |

## Arquivos

### 01-vba-msforms

- [UserForm](01-vba-msforms/userform-object.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Reference/User-Interface-Help/userform-object.md`
- [Label](01-vba-msforms/label-control.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Reference/User-Interface-Help/label-control.md`
- [TextBox](01-vba-msforms/textbox-control.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Reference/User-Interface-Help/textbox-control.md`
- [ComboBox](01-vba-msforms/combobox-control.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Reference/User-Interface-Help/combobox-control.md`
- [ListBox](01-vba-msforms/listbox-control.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Reference/User-Interface-Help/listbox-control.md`
- [Image](01-vba-msforms/image-control.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Reference/User-Interface-Help/image-control.md`
- [Frame](01-vba-msforms/frame-control.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Reference/User-Interface-Help/frame-control.md`
- [ScrollBar](01-vba-msforms/scrollbar-control.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Reference/User-Interface-Help/scrollbar-control.md`
- [CommandButton](01-vba-msforms/commandbutton-control.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Reference/User-Interface-Help/commandbutton-control.md`
- [CheckBox](01-vba-msforms/checkbox-control.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Reference/User-Interface-Help/checkbox-control.md`
- [Zoom](01-vba-msforms/zoom-property.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Reference/User-Interface-Help/zoom-property.md`
- [Picture](01-vba-msforms/picture-property.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Reference/User-Interface-Help/picture-property.md`
- [PictureSizeMode](01-vba-msforms/picturesizemode-property.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Reference/User-Interface-Help/picturesizemode-property.md`
- [BackStyle](01-vba-msforms/backstyle-property-microsoft-forms.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Reference/User-Interface-Help/backstyle-property-microsoft-forms.md`
- [BorderColor](01-vba-msforms/bordercolor-property.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Reference/User-Interface-Help/bordercolor-property.md`
- [BorderStyle](01-vba-msforms/borderstyle-property.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Reference/User-Interface-Help/borderstyle-property.md`
- [SpecialEffect](01-vba-msforms/specialeffect-property.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Reference/User-Interface-Help/specialeffect-property.md`
- [MouseMove](01-vba-msforms/mousemove-event.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Reference/User-Interface-Help/mousemove-event.md`
- [MousePointer](01-vba-msforms/mousepointer-property.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Reference/User-Interface-Help/mousepointer-property.md`
- [Controls.Add](01-vba-msforms/add-method-microsoft-forms.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Reference/User-Interface-Help/add-method-microsoft-forms.md`
- [Controls](01-vba-msforms/controls-collection-microsoft-forms.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Reference/User-Interface-Help/controls-collection-microsoft-forms.md`
- [ColumnWidths](01-vba-msforms/columnwidths-property.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Reference/User-Interface-Help/columnwidths-property.md`
- [Show](01-vba-msforms/show-method.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Reference/User-Interface-Help/show-method.md`
- [Change](01-vba-msforms/change-event.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Reference/User-Interface-Help/change-event.md`
- [AutoSize](01-vba-msforms/autosize-property.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Reference/User-Interface-Help/autosize-property.md`
- [Visible](01-vba-msforms/visible-property-microsoft-forms.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Reference/User-Interface-Help/visible-property-microsoft-forms.md`
- [Value](01-vba-msforms/value-property-microsoft-forms.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Reference/User-Interface-Help/value-property-microsoft-forms.md`
- [Imagens no Image](01-vba-msforms/figuras-no-controle-image.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Concepts/Forms/things-you-can-do-with-a-picture-on-an-image-control.md`

### 02-vba-linguagem

- [VBA 64 bits](02-vba-linguagem/vba-64-bits.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Concepts/Getting-Started/64-bit-visual-basic-for-applications-overview.md`
- [PtrSafe](02-vba-linguagem/ptrsafe.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Reference/User-Interface-Help/ptrsafe-keyword.md`
- [Declare](02-vba-linguagem/declare.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Reference/User-Interface-Help/declare-statement.md`
- [RGB](02-vba-linguagem/rgb.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Reference/User-Interface-Help/rgb-function.md`
- [Like](02-vba-linguagem/like.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Reference/User-Interface-Help/like-operator.md`
- [IIf](02-vba-linguagem/iif.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Reference/User-Interface-Help/iif-function.md`
- [DateAdd](02-vba-linguagem/dateadd.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Reference/User-Interface-Help/dateadd-function.md`
- [DateDiff](02-vba-linguagem/datediff.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Reference/User-Interface-Help/datediff-function.md`

### 03-vbe-modelo-extensibilidade

- [Referência](03-vbe-modelo-extensibilidade/visual-basic-add-in-model-reference.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Reference/User-Interface-Help/visual-basic-add-in-model-reference.md`
- [Import](03-vbe-modelo-extensibilidade/import-method-vba-add-in-object-model.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Reference/User-Interface-Help/import-method-vba-add-in-object-model.md`
- [Export](03-vbe-modelo-extensibilidade/export-method-vba-add-in-object-model.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Reference/User-Interface-Help/export-method-vba-add-in-object-model.md`
- [Add](03-vbe-modelo-extensibilidade/add-method-vba-add-in-object-model.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Reference/User-Interface-Help/add-method-vba-add-in-object-model.md`
- [Remove](03-vbe-modelo-extensibilidade/remove-method-vba-add-in-object-model.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Reference/User-Interface-Help/remove-method-vba-add-in-object-model.md`
- [AddFromString](03-vbe-modelo-extensibilidade/addfromstring-method-vba-add-in-object-model.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Reference/User-Interface-Help/addfromstring-method-vba-add-in-object-model.md`
- [AddFromFile](03-vbe-modelo-extensibilidade/addfromfile-method-vba-add-in-object-model.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Reference/User-Interface-Help/addfromfile-method-vba-add-in-object-model.md`
- [InsertLines](03-vbe-modelo-extensibilidade/insertlines-method-vba-add-in-object-model.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Reference/User-Interface-Help/insertlines-method-vba-add-in-object-model.md`
- [DeleteLines](03-vbe-modelo-extensibilidade/deletelines-method-vba-add-in-object-model.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Reference/User-Interface-Help/deletelines-method-vba-add-in-object-model.md`
- [ReplaceLine](03-vbe-modelo-extensibilidade/replaceline-method-vba-add-in-object-model.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Reference/User-Interface-Help/replaceline-method-vba-add-in-object-model.md`
- [objects-visual-basic-add-in-model](03-vbe-modelo-extensibilidade/objects-visual-basic-add-in-model.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Reference/Visual-Basic-Add-in-Model/objects-visual-basic-add-in-model.md`
- [collections-visual-basic-add-in-model](03-vbe-modelo-extensibilidade/collections-visual-basic-add-in-model.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Reference/Visual-Basic-Add-in-Model/collections-visual-basic-add-in-model.md`
- [methods-visual-basic-add-in-model](03-vbe-modelo-extensibilidade/methods-visual-basic-add-in-model.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Reference/Visual-Basic-Add-in-Model/methods-visual-basic-add-in-model.md`
- [properties-visual-basic-add-in-model](03-vbe-modelo-extensibilidade/properties-visual-basic-add-in-model.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Reference/Visual-Basic-Add-in-Model/properties-visual-basic-add-in-model.md`
- [events-visual-basic-add-in-model](03-vbe-modelo-extensibilidade/events-visual-basic-add-in-model.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Reference/Visual-Basic-Add-in-Model/events-visual-basic-add-in-model.md`
- [Workbook.VBProject](03-vbe-modelo-extensibilidade/excel-workbook-vbproject.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/api/Excel.Workbook.VBProject.md`
- [Macros e acesso ao modelo VBA](03-vbe-modelo-extensibilidade/confiar-acesso-modelo-objeto-vba.html.md) — `https://support.microsoft.com/en-us/office/enable-or-disable-macros-in-microsoft-365-files-12b036fd-d140-4e74-b45e-16fed1a7e5c6`

### 04-excel-modelo-objetos

- [Excel.Application.OnTime](04-excel-modelo-objetos/application.ontime.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/api/Excel.Application.OnTime.md`
- [Excel.Application.EnableEvents](04-excel-modelo-objetos/application.enableevents.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/api/Excel.Application.EnableEvents.md`
- [Excel.Application.ScreenUpdating](04-excel-modelo-objetos/application.screenupdating.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/api/Excel.Application.ScreenUpdating.md`
- [Excel.Application.AutomationSecurity](04-excel-modelo-objetos/application.automationsecurity.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/api/Excel.Application.AutomationSecurity.md`
- [Excel.Workbook.Open](04-excel-modelo-objetos/workbook.open.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/api/Excel.Workbook.Open.md`
- [Excel.Workbook.BeforeClose](04-excel-modelo-objetos/workbook.beforeclose.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/api/Excel.Workbook.BeforeClose.md`
- [Excel.Workbook.BeforeSave](04-excel-modelo-objetos/workbook.beforesave.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/api/Excel.Workbook.BeforeSave.md`
- [Excel.Workbook.SaveAs](04-excel-modelo-objetos/workbook.saveas.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/api/Excel.Workbook.SaveAs.md`
- [Excel.Workbook.SheetBeforeDoubleClick](04-excel-modelo-objetos/workbook.sheetbeforedoubleclick.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/api/Excel.Workbook.SheetBeforeDoubleClick.md`
- [Excel.Workbook.SheetSelectionChange](04-excel-modelo-objetos/workbook.sheetselectionchange.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/api/Excel.Workbook.SheetSelectionChange.md`
- [Excel.Workbook.SheetChange](04-excel-modelo-objetos/workbook.sheetchange.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/api/Excel.Workbook.SheetChange.md`
- [Excel.Workbook.SheetBeforeRightClick](04-excel-modelo-objetos/workbook.sheetbeforerightclick.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/api/Excel.Workbook.SheetBeforeRightClick.md`
- [Excel.Worksheet.CodeName](04-excel-modelo-objetos/worksheet.codename.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/api/Excel.Worksheet.CodeName.md`
- [Excel.Worksheet.Protect](04-excel-modelo-objetos/worksheet.protect.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/api/Excel.Worksheet.Protect.md`
- [Excel.Shapes.AddShape](04-excel-modelo-objetos/shapes.addshape.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/api/Excel.Shapes.AddShape.md`
- [Excel.Shape.Placement](04-excel-modelo-objetos/shape.placement.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/api/Excel.Shape.Placement.md`
- [Excel.Shape.OnAction](04-excel-modelo-objetos/shape.onaction.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/api/Excel.Shape.OnAction.md`
- [Excel.FormatConditions.Add](04-excel-modelo-objetos/formatconditions.add.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/api/Excel.FormatConditions.Add.md`
- [Excel.Window.DisplayGridlines](04-excel-modelo-objetos/window.displaygridlines.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/api/Excel.Window.DisplayGridlines.md`

### 05-excel-ribbon-customui

- [Ribbon](05-excel-ribbon-customui/visao-geral-ribbon.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Library-Reference/Concepts/overview-of-the-office-fluent-ribbon.md`
- [customUI no pacote](05-excel-ribbon-customui/ribbon-por-arquivo-open-xml.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Library-Reference/Concepts/customize-the-office-fluent-ribbon-by-using-an-open-xml-formats-file.md`
- [IRibbonUI](05-excel-ribbon-customui/iribbonui.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/api/Office.IRibbonUI.md`
- [IRibbonControl](05-excel-ribbon-customui/iribboncontrol.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/api/Office.IRibbonControl.md`

### 06-access-sql-ace

- [create-and-delete-tables-and-indexes-using-access-sql](06-access-sql-ace/create-and-delete-tables-and-indexes-using-access-sql.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/access/Concepts/Structured-Query-Language/create-and-delete-tables-and-indexes-using-access-sql.md`
- [define-relationships-between-tables-using-access-sql](06-access-sql-ace/define-relationships-between-tables-using-access-sql.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/access/Concepts/Structured-Query-Language/define-relationships-between-tables-using-access-sql.md`
- [modify-a-table-s-design-using-access-sql](06-access-sql-ace/modify-a-table-s-design-using-access-sql.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/access/Concepts/Structured-Query-Language/modify-a-table-s-design-using-access-sql.md`
- [insert-update-and-delete-records-from-a-table-using-access-sql](06-access-sql-ace/insert-update-and-delete-records-from-a-table-using-access-sql.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/access/Concepts/Structured-Query-Language/insert-update-and-delete-records-from-a-table-using-access-sql.md`
- [retrieve-records-using-access-sql](06-access-sql-ace/retrieve-records-using-access-sql.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/access/Concepts/Structured-Query-Language/retrieve-records-using-access-sql.md`
- [all-distinct-distinctrow-top-predicates-microsoft-access-sql](06-access-sql-ace/all-distinct-distinctrow-top-predicates-microsoft-access-sql.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/access/Concepts/Structured-Query-Language/all-distinct-distinctrow-top-predicates-microsoft-access-sql.md`
- [where-clause-microsoft-access-sql](06-access-sql-ace/where-clause-microsoft-access-sql.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/access/Concepts/Structured-Query-Language/where-clause-microsoft-access-sql.md`
- [order-by-clause-microsoft-access-sql](06-access-sql-ace/order-by-clause-microsoft-access-sql.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/access/Concepts/Structured-Query-Language/order-by-clause-microsoft-access-sql.md`
- [group-by-clause-microsoft-access-sql](06-access-sql-ace/group-by-clause-microsoft-access-sql.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/access/Concepts/Structured-Query-Language/group-by-clause-microsoft-access-sql.md`
- [having-clause-microsoft-access-sql](06-access-sql-ace/having-clause-microsoft-access-sql.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/access/Concepts/Structured-Query-Language/having-clause-microsoft-access-sql.md`
- [from-clause-microsoft-access-sql](06-access-sql-ace/from-clause-microsoft-access-sql.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/access/Concepts/Structured-Query-Language/from-clause-microsoft-access-sql.md`
- [perform-joins-using-access-sql](06-access-sql-ace/perform-joins-using-access-sql.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/access/Concepts/Structured-Query-Language/perform-joins-using-access-sql.md`
- [like-operator-microsoft-access-sql](06-access-sql-ace/like-operator-microsoft-access-sql.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/access/Concepts/Structured-Query-Language/like-operator-microsoft-access-sql.md`
- [use-aggregate-functions-to-work-with-values-in-access-sql](06-access-sql-ace/use-aggregate-functions-to-work-with-values-in-access-sql.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/access/Concepts/Structured-Query-Language/use-aggregate-functions-to-work-with-values-in-access-sql.md`
- [use-international-date-formats-in-sql-statements](06-access-sql-ace/use-international-date-formats-in-sql-statements.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/access/Concepts/Structured-Query-Language/use-international-date-formats-in-sql-statements.md`
- [build-sql-statements-that-include-variables-and-controls](06-access-sql-ace/build-sql-statements-that-include-variables-and-controls.md) — `https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/access/Concepts/Structured-Query-Language/build-sql-statements-that-include-variables-and-controls.md`
- [ACE redistribuível](06-access-sql-ace/access-database-engine-2016-redistributable.html.md) — `https://www.microsoft.com/en-us/download/details.aspx?id=54920`

### 07-ado

- [connection-object-ado](07-ado/connection-object-ado.html.md) — `https://learn.microsoft.com/en-us/previous-versions/sql/ado/reference/ado-api/connection-object-ado?view=sql-server-ver15`
- [command-object-ado](07-ado/command-object-ado.html.md) — `https://learn.microsoft.com/en-us/previous-versions/sql/ado/reference/ado-api/command-object-ado?view=sql-server-ver15`
- [parameter-object](07-ado/parameter-object.html.md) — `https://learn.microsoft.com/en-us/previous-versions/sql/ado/reference/ado-api/parameter-object?view=sql-server-ver15`
- [createparameter-method-ado](07-ado/createparameter-method-ado.html.md) — `https://learn.microsoft.com/en-us/previous-versions/sql/ado/reference/ado-api/createparameter-method-ado?view=sql-server-ver15`
- [parameters-collection-ado](07-ado/parameters-collection-ado.html.md) — `https://learn.microsoft.com/en-us/previous-versions/sql/ado/reference/ado-api/parameters-collection-ado?view=sql-server-ver15`
- [append-method-ado](07-ado/append-method-ado.html.md) — `https://learn.microsoft.com/en-us/previous-versions/sql/ado/reference/ado-api/append-method-ado?view=sql-server-ver15`
- [datatypeenum](07-ado/datatypeenum.html.md) — `https://learn.microsoft.com/en-us/previous-versions/sql/ado/reference/ado-api/datatypeenum?view=sql-server-ver15`
- [execute-method-ado-command](07-ado/execute-method-ado-command.html.md) — `https://learn.microsoft.com/en-us/previous-versions/sql/ado/reference/ado-api/execute-method-ado-command?view=sql-server-ver15`
- [begintrans-committrans-and-rollbacktrans-methods-ado](07-ado/begintrans-committrans-and-rollbacktrans-methods-ado.html.md) — `https://learn.microsoft.com/en-us/previous-versions/sql/ado/reference/ado-api/begintrans-committrans-and-rollbacktrans-methods-ado?view=sql-server-ver15`
- [recordset-object-ado](07-ado/recordset-object-ado.html.md) — `https://learn.microsoft.com/en-us/previous-versions/sql/ado/reference/ado-api/recordset-object-ado?view=sql-server-ver15`
- [getrows-method-ado](07-ado/getrows-method-ado.html.md) — `https://learn.microsoft.com/en-us/previous-versions/sql/ado/reference/ado-api/getrows-method-ado?view=sql-server-ver15`
- [commandtypeenum](07-ado/commandtypeenum.html.md) — `https://learn.microsoft.com/en-us/previous-versions/sql/ado/reference/ado-api/commandtypeenum?view=sql-server-ver15`
- [cursorlocationenum](07-ado/cursorlocationenum.html.md) — `https://learn.microsoft.com/en-us/previous-versions/sql/ado/reference/ado-api/cursorlocationenum?view=sql-server-ver15`

### 08-dotnet-oledb

- [ole-db-odbc-and-oracle-connection-pooling](08-dotnet-oledb/ole-db-odbc-and-oracle-connection-pooling.md) — `https://raw.githubusercontent.com/dotnet/docs/main/docs/framework/data/adonet/ole-db-odbc-and-oracle-connection-pooling.md`
- [configuring-parameters-and-parameter-data-types](08-dotnet-oledb/configuring-parameters-and-parameter-data-types.md) — `https://raw.githubusercontent.com/dotnet/docs/main/docs/framework/data/adonet/configuring-parameters-and-parameter-data-types.md`
- [ole-db-data-type-mappings](08-dotnet-oledb/ole-db-data-type-mappings.md) — `https://raw.githubusercontent.com/dotnet/docs/main/docs/framework/data/adonet/ole-db-data-type-mappings.md`
- [local-transactions](08-dotnet-oledb/local-transactions.md) — `https://raw.githubusercontent.com/dotnet/docs/main/docs/framework/data/adonet/local-transactions.md`
- [retrieving-identity-or-autonumber-values](08-dotnet-oledb/retrieving-identity-or-autonumber-values.md) — `https://raw.githubusercontent.com/dotnet/docs/main/docs/framework/data/adonet/retrieving-identity-or-autonumber-values.md`
- [system.data.oledb.oledbtype](08-dotnet-oledb/system.data.oledb.oledbtype.html.md) — `https://learn.microsoft.com/en-us/dotnet/api/system.data.oledb.oledbtype?view=netframework-4.8.1`
- [system.data.oledb.oledbparameter](08-dotnet-oledb/system.data.oledb.oledbparameter.html.md) — `https://learn.microsoft.com/en-us/dotnet/api/system.data.oledb.oledbparameter?view=netframework-4.8.1`
- [system.data.oledb.oledbconnection.connectionstring](08-dotnet-oledb/system.data.oledb.oledbconnection.connectionstring.html.md) — `https://learn.microsoft.com/en-us/dotnet/api/system.data.oledb.oledbconnection.connectionstring?view=netframework-4.8.1`
- [system.data.oledb.oledbtransaction](08-dotnet-oledb/system.data.oledb.oledbtransaction.html.md) — `https://learn.microsoft.com/en-us/dotnet/api/system.data.oledb.oledbtransaction?view=netframework-4.8.1`
- [system.data.oledb.oledbenumerator](08-dotnet-oledb/system.data.oledb.oledbenumerator.html.md) — `https://learn.microsoft.com/en-us/dotnet/api/system.data.oledb.oledbenumerator?view=netframework-4.8.1`

### 09-powershell-5.1

- [ConvertTo-Json](09-powershell-5.1/ConvertTo-Json.md) — `https://raw.githubusercontent.com/MicrosoftDocs/PowerShell-Docs/main/reference/5.1/Microsoft.PowerShell.Utility/ConvertTo-Json.md`
- [ConvertFrom-Json](09-powershell-5.1/ConvertFrom-Json.md) — `https://raw.githubusercontent.com/MicrosoftDocs/PowerShell-Docs/main/reference/5.1/Microsoft.PowerShell.Utility/ConvertFrom-Json.md`
- [about_Character_Encoding](09-powershell-5.1/about_Character_Encoding.md) — `https://raw.githubusercontent.com/MicrosoftDocs/PowerShell-Docs/main/reference/5.1/Microsoft.PowerShell.Core/About/about_Character_Encoding.md`
- [about_Execution_Policies](09-powershell-5.1/about_Execution_Policies.md) — `https://raw.githubusercontent.com/MicrosoftDocs/PowerShell-Docs/main/reference/5.1/Microsoft.PowerShell.Core/About/about_Execution_Policies.md`
- [about_PowerShell_exe](09-powershell-5.1/about_PowerShell_exe.md) — `https://raw.githubusercontent.com/MicrosoftDocs/PowerShell-Docs/main/reference/5.1/Microsoft.PowerShell.Core/About/about_PowerShell_exe.md`
- [about_Scripts](09-powershell-5.1/about_Scripts.md) — `https://raw.githubusercontent.com/MicrosoftDocs/PowerShell-Docs/main/reference/5.1/Microsoft.PowerShell.Core/About/about_Scripts.md`
- [about_Scopes](09-powershell-5.1/about_Scopes.md) — `https://raw.githubusercontent.com/MicrosoftDocs/PowerShell-Docs/main/reference/5.1/Microsoft.PowerShell.Core/About/about_Scopes.md`
- [about_Script_Blocks](09-powershell-5.1/about_Script_Blocks.md) — `https://raw.githubusercontent.com/MicrosoftDocs/PowerShell-Docs/main/reference/5.1/Microsoft.PowerShell.Core/About/about_Script_Blocks.md`
- [about_Try_Catch_Finally](09-powershell-5.1/about_Try_Catch_Finally.md) — `https://raw.githubusercontent.com/MicrosoftDocs/PowerShell-Docs/main/reference/5.1/Microsoft.PowerShell.Core/About/about_Try_Catch_Finally.md`
- [about_Automatic_Variables](09-powershell-5.1/about_Automatic_Variables.md) — `https://raw.githubusercontent.com/MicrosoftDocs/PowerShell-Docs/main/reference/5.1/Microsoft.PowerShell.Core/About/about_Automatic_Variables.md`
- [about_Preference_Variables](09-powershell-5.1/about_Preference_Variables.md) — `https://raw.githubusercontent.com/MicrosoftDocs/PowerShell-Docs/main/reference/5.1/Microsoft.PowerShell.Core/About/about_Preference_Variables.md`

### 10-edge-headless

- [Headless](10-edge-headless/chrome-headless.html.md) — `https://developer.chrome.com/docs/chromium/headless`

### 11-windows-icones-dpi

- [Segoe MDL2 Assets](11-windows-icones-dpi/segoe-mdl2-assets.md) — `https://raw.githubusercontent.com/MicrosoftDocs/windows-dev-docs/docs/hub/apps/design/iconography/segoe-ui-symbol-font.md`
- [Segoe Fluent Icons](11-windows-icones-dpi/segoe-fluent-icons.md) — `https://raw.githubusercontent.com/MicrosoftDocs/windows-dev-docs/docs/hub/apps/design/iconography/segoe-fluent-icons-font.md`
- [GetDeviceCaps](11-windows-icones-dpi/getdevicecaps.md) — `https://raw.githubusercontent.com/MicrosoftDocs/sdk-api/docs/sdk-api-src/content/wingdi/nf-wingdi-getdevicecaps.md`
- [GetDC](11-windows-icones-dpi/getdc.md) — `https://raw.githubusercontent.com/MicrosoftDocs/sdk-api/docs/sdk-api-src/content/winuser/nf-winuser-getdc.md`
- [ReleaseDC](11-windows-icones-dpi/releasedc.md) — `https://raw.githubusercontent.com/MicrosoftDocs/sdk-api/docs/sdk-api-src/content/winuser/nf-winuser-releasedc.md`
- [GetDpiForSystem](11-windows-icones-dpi/getdpiforsystem.md) — `https://raw.githubusercontent.com/MicrosoftDocs/sdk-api/docs/sdk-api-src/content/winuser/nf-winuser-getdpiforsystem.md`

### 12-mcp-especificacao

- [index.mdx](12-mcp-especificacao/2026-07-28_index.md) — `https://raw.githubusercontent.com/modelcontextprotocol/modelcontextprotocol/main/docs/specification/2026-07-28/index.mdx`
- [changelog.mdx](12-mcp-especificacao/2026-07-28_changelog.md) — `https://raw.githubusercontent.com/modelcontextprotocol/modelcontextprotocol/main/docs/specification/2026-07-28/changelog.mdx`
- [deprecated.mdx](12-mcp-especificacao/2026-07-28_deprecated.md) — `https://raw.githubusercontent.com/modelcontextprotocol/modelcontextprotocol/main/docs/specification/2026-07-28/deprecated.mdx`
- [basic/index.mdx](12-mcp-especificacao/2026-07-28_basic_index.md) — `https://raw.githubusercontent.com/modelcontextprotocol/modelcontextprotocol/main/docs/specification/2026-07-28/basic/index.mdx`
- [basic/versioning.mdx](12-mcp-especificacao/2026-07-28_basic_versioning.md) — `https://raw.githubusercontent.com/modelcontextprotocol/modelcontextprotocol/main/docs/specification/2026-07-28/basic/versioning.mdx`
- [basic/transports/index.mdx](12-mcp-especificacao/2026-07-28_basic_transports_index.md) — `https://raw.githubusercontent.com/modelcontextprotocol/modelcontextprotocol/main/docs/specification/2026-07-28/basic/transports/index.mdx`
- [basic/transports/stdio.mdx](12-mcp-especificacao/2026-07-28_basic_transports_stdio.md) — `https://raw.githubusercontent.com/modelcontextprotocol/modelcontextprotocol/main/docs/specification/2026-07-28/basic/transports/stdio.mdx`
- [basic/patterns/cancellation.mdx](12-mcp-especificacao/2026-07-28_basic_patterns_cancellation.md) — `https://raw.githubusercontent.com/modelcontextprotocol/modelcontextprotocol/main/docs/specification/2026-07-28/basic/patterns/cancellation.mdx`
- [server/index.mdx](12-mcp-especificacao/2026-07-28_server_index.md) — `https://raw.githubusercontent.com/modelcontextprotocol/modelcontextprotocol/main/docs/specification/2026-07-28/server/index.mdx`
- [server/discover.mdx](12-mcp-especificacao/2026-07-28_server_discover.md) — `https://raw.githubusercontent.com/modelcontextprotocol/modelcontextprotocol/main/docs/specification/2026-07-28/server/discover.mdx`
- [server/tools.mdx](12-mcp-especificacao/2026-07-28_server_tools.md) — `https://raw.githubusercontent.com/modelcontextprotocol/modelcontextprotocol/main/docs/specification/2026-07-28/server/tools.mdx`
- [server/utilities/caching.mdx](12-mcp-especificacao/2026-07-28_server_utilities_caching.md) — `https://raw.githubusercontent.com/modelcontextprotocol/modelcontextprotocol/main/docs/specification/2026-07-28/server/utilities/caching.mdx`
- [server/utilities/pagination.mdx](12-mcp-especificacao/2026-07-28_server_utilities_pagination.md) — `https://raw.githubusercontent.com/modelcontextprotocol/modelcontextprotocol/main/docs/specification/2026-07-28/server/utilities/pagination.mdx`
- [schema.ts](12-mcp-especificacao/2026-07-28_schema.ts) — `https://raw.githubusercontent.com/modelcontextprotocol/modelcontextprotocol/main/schema/2026-07-28/schema.ts`
- [changelog.mdx](12-mcp-especificacao/2025-11-25_changelog.md) — `https://raw.githubusercontent.com/modelcontextprotocol/modelcontextprotocol/main/docs/specification/2025-11-25/changelog.mdx`
- [basic/lifecycle.mdx](12-mcp-especificacao/2025-11-25_basic_lifecycle.md) — `https://raw.githubusercontent.com/modelcontextprotocol/modelcontextprotocol/main/docs/specification/2025-11-25/basic/lifecycle.mdx`
- [basic/transports.mdx](12-mcp-especificacao/2025-11-25_basic_transports.md) — `https://raw.githubusercontent.com/modelcontextprotocol/modelcontextprotocol/main/docs/specification/2025-11-25/basic/transports.mdx`
- [server/tools.mdx](12-mcp-especificacao/2025-11-25_server_tools.md) — `https://raw.githubusercontent.com/modelcontextprotocol/modelcontextprotocol/main/docs/specification/2025-11-25/server/tools.mdx`

### 13-claude-desktop

- [MCP local no Claude Desktop](13-claude-desktop/conectar-servidores-locais.md) — `https://raw.githubusercontent.com/modelcontextprotocol/modelcontextprotocol/main/docs/docs/2026-07-28/develop/connect-local-servers.mdx`
- [Help Center](13-claude-desktop/servidores-locais-help-center.html.md) — `https://support.claude.com/en/articles/10949351-getting-started-with-local-mcp-servers-on-claude-desktop`
- [MCPB manifest](13-claude-desktop/mcpb-manifest.md) — `https://raw.githubusercontent.com/modelcontextprotocol/mcpb/main/MANIFEST.md`
- [MCPB](13-claude-desktop/mcpb-readme.md) — `https://raw.githubusercontent.com/modelcontextprotocol/mcpb/main/README.md`
- [Skills](13-claude-desktop/agent-skills-visao-geral.md) — `https://docs.claude.com/en/docs/agents-and-tools/agent-skills/overview.md`
- [Skills: boas práticas](13-claude-desktop/agent-skills-boas-praticas.md) — `https://docs.claude.com/en/docs/agents-and-tools/agent-skills/best-practices.md`
- [Skills no app](13-claude-desktop/usar-skills-no-claude.html.md) — `https://support.claude.com/en/articles/12512180-using-skills-in-claude`

### 14-office-2019

- [Ciclo de vida](14-office-2019/ciclo-de-vida-office-2019.html.md) — `https://learn.microsoft.com/en-us/lifecycle/products/microsoft-office-2019`
- [Novidades Excel 2019](14-office-2019/novidades-excel-2019.html.md) — `https://support.microsoft.com/en-us/office/what-s-new-in-excel-2019-for-windows-5a201203-1155-4055-82a5-82bf0994631f`
