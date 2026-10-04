# Font (Microsoft Forms) — Size e Bold

Fontes: https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/font-object-microsoft-forms
e https://github.com/MicrosoftDocs/VBA-Docs/blob/main/Language/Reference/User-Interface-Help/bold-italic-size-strikethrough-underline-weight-properties.md
(consultado em 19/09/2026)

- Cada controle e formulário tem o próprio objeto **Font**; a propriedade padrão é **Name**
  (texto nulo = fonte padrão do sistema). A fonte do formulário/contêiner é o padrão dos controles novos.
- Sintaxe: `object.Size [= Currency]` — **Size é Currency** ("número que indica o tamanho da fonte").
- Sintaxe: `object.Bold [= Boolean]`. Mudar **Bold** muda **Weight** (True = 700, False = 400).

## Consequência no build (PowerShell 5.1 + COM, conferido no Excel 2019)
- `$ctl.Font.Name = ...` falha pelo PowerShell ("referência de objeto não definida");
  `FontName`/`FontSize`/`FontBold`/`FontItalic` funcionam (o Frame não tem: usar o objeto Font via InvokeMember).
- `FontSize`: Int32 passa; Double/Decimal/Single são recusados; fração vai em
  `System.Runtime.InteropServices.CurrencyWrapper`.
- `VBComponent.Properties(..).Value` (Width, Height, BackColor): passar como **texto** (número em formato invariante).
