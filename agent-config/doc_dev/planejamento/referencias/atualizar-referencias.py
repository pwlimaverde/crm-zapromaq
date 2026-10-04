# -*- coding: utf-8 -*-
"""
atualizar-referencias.py - baixa a documentação oficial das stacks do projeto
e grava em Markdown nesta pasta, para consulta offline.

Fontes, em ordem de preferência:
  1. Markdown original do repositório público da documentação (GitHub raw);
  2. página HTML oficial convertida para Markdown (quando não há repositório público).

Uso (máquina de desenvolvimento, com internet):
    python atualizar-referencias.py            # baixa tudo
    python atualizar-referencias.py mcp vba    # só as pastas cujo nome contém esses termos

Requer: pip install markdownify beautifulsoup4 (só para as páginas HTML).
"""
import os, sys, re, json, datetime, urllib.request, urllib.error, http.cookiejar

# o Learn redireciona gravando cookie; sem guardar cookie o urllib entra em laço de 301
_ABRIDOR = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(http.cookiejar.CookieJar()))

BASE = os.path.dirname(os.path.abspath(__file__))
HOJE = datetime.date.today().isoformat()

GH = 'https://raw.githubusercontent.com/'
VBA = GH + 'MicrosoftDocs/VBA-Docs/main/'
UI = VBA + 'Language/Reference/User-Interface-Help/'
PS = GH + 'MicrosoftDocs/PowerShell-Docs/main/reference/5.1/'
NET = GH + 'dotnet/docs/main/docs/framework/data/adonet/'
SDK = GH + 'MicrosoftDocs/sdk-api/docs/sdk-api-src/content/'
WIN = GH + 'MicrosoftDocs/windows-dev-docs/docs/hub/apps/design/iconography/'
MCP = GH + 'modelcontextprotocol/modelcontextprotocol/main/'
LEARN = 'https://learn.microsoft.com/en-us/'
# a referência do ADO foi movida para 'previous-versions'; o endereço antigo entra em laço de 301
ADO = LEARN + 'previous-versions/sql/ado/reference/ado-api/'

# (pasta, arquivo de destino, url, título curto)
FONTES = [
    # ------------------------------------------------ MSForms (UserForm e controles)
    *[('01-vba-msforms', n + '.md', UI + n + '.md', t) for n, t in [
        ('userform-object', 'UserForm'), ('label-control', 'Label'), ('textbox-control', 'TextBox'),
        ('combobox-control', 'ComboBox'), ('listbox-control', 'ListBox'), ('image-control', 'Image'),
        ('frame-control', 'Frame'), ('scrollbar-control', 'ScrollBar'), ('commandbutton-control', 'CommandButton'),
        ('checkbox-control', 'CheckBox'), ('zoom-property', 'Zoom'), ('picture-property', 'Picture'),
        ('picturesizemode-property', 'PictureSizeMode'), ('backstyle-property-microsoft-forms', 'BackStyle'),
        ('bordercolor-property', 'BorderColor'), ('borderstyle-property', 'BorderStyle'),
        ('specialeffect-property', 'SpecialEffect'), ('mousemove-event', 'MouseMove'),
        ('mousepointer-property', 'MousePointer'), ('add-method-microsoft-forms', 'Controls.Add'),
        ('controls-collection-microsoft-forms', 'Controls'), ('columnwidths-property', 'ColumnWidths'),
        ('show-method', 'Show'), ('change-event', 'Change'), ('autosize-property', 'AutoSize'),
        ('visible-property-microsoft-forms', 'Visible'), ('value-property-microsoft-forms', 'Value')]],
    ('01-vba-msforms', 'figuras-no-controle-image.md', VBA + 'Language/Concepts/Forms/things-you-can-do-with-a-picture-on-an-image-control.md', 'Imagens no Image'),
    # ------------------------------------------------ linguagem VBA
    ('02-vba-linguagem', 'vba-64-bits.md', VBA + 'Language/Concepts/Getting-Started/64-bit-visual-basic-for-applications-overview.md', 'VBA 64 bits'),
    ('02-vba-linguagem', 'ptrsafe.md', UI + 'ptrsafe-keyword.md', 'PtrSafe'),
    ('02-vba-linguagem', 'declare.md', UI + 'declare-statement.md', 'Declare'),
    ('02-vba-linguagem', 'rgb.md', UI + 'rgb-function.md', 'RGB'),
    ('02-vba-linguagem', 'like.md', UI + 'like-operator.md', 'Like'),
    ('02-vba-linguagem', 'iif.md', UI + 'iif-function.md', 'IIf'),
    ('02-vba-linguagem', 'dateadd.md', UI + 'dateadd-function.md', 'DateAdd'),
    ('02-vba-linguagem', 'datediff.md', UI + 'datediff-function.md', 'DateDiff'),
    # ------------------------------------------------ VBE / VBIDE (usado pelo build)
    *[('03-vbe-modelo-extensibilidade', n + '.md', UI + n + '.md', t) for n, t in [
        ('visual-basic-add-in-model-reference', 'Referência'), ('import-method-vba-add-in-object-model', 'Import'),
        ('export-method-vba-add-in-object-model', 'Export'), ('add-method-vba-add-in-object-model', 'Add'),
        ('remove-method-vba-add-in-object-model', 'Remove'), ('addfromstring-method-vba-add-in-object-model', 'AddFromString'),
        ('addfromfile-method-vba-add-in-object-model', 'AddFromFile'), ('insertlines-method-vba-add-in-object-model', 'InsertLines'),
        ('deletelines-method-vba-add-in-object-model', 'DeleteLines'), ('replaceline-method-vba-add-in-object-model', 'ReplaceLine')]],
    *[('03-vbe-modelo-extensibilidade', n.lower() + '.md', VBA + 'Language/Reference/Visual-Basic-Add-in-Model/' + n + '.md', n) for n in [
        'objects-visual-basic-add-in-model', 'collections-visual-basic-add-in-model', 'methods-visual-basic-add-in-model',
        'properties-visual-basic-add-in-model', 'events-visual-basic-add-in-model']],
    ('03-vbe-modelo-extensibilidade', 'excel-workbook-vbproject.md', VBA + 'api/Excel.Workbook.VBProject.md', 'Workbook.VBProject'),
    ('03-vbe-modelo-extensibilidade', 'confiar-acesso-modelo-objeto-vba.html.md',
     'https://support.microsoft.com/en-us/office/enable-or-disable-macros-in-microsoft-365-files-12b036fd-d140-4e74-b45e-16fed1a7e5c6', 'Macros e acesso ao modelo VBA'),
    # ------------------------------------------------ modelo de objetos do Excel
    *[('04-excel-modelo-objetos', n.replace('Excel.', '').lower() + '.md', VBA + 'api/' + n + '.md', n) for n in [
        'Excel.Application.OnTime', 'Excel.Application.EnableEvents', 'Excel.Application.ScreenUpdating',
        'Excel.Application.AutomationSecurity', 'Excel.Workbook.Open', 'Excel.Workbook.BeforeClose',
        'Excel.Workbook.BeforeSave', 'Excel.Workbook.SaveAs', 'Excel.Workbook.SheetBeforeDoubleClick',
        'Excel.Workbook.SheetSelectionChange', 'Excel.Workbook.SheetChange', 'Excel.Workbook.SheetBeforeRightClick',
        'Excel.Worksheet.CodeName', 'Excel.Worksheet.Protect', 'Excel.Shapes.AddShape', 'Excel.Shape.Placement',
        'Excel.Shape.OnAction', 'Excel.FormatConditions.Add', 'Excel.Window.DisplayGridlines']],
    # ------------------------------------------------ faixa de opções (customUI)
    ('05-excel-ribbon-customui', 'visao-geral-ribbon.md', VBA + 'Library-Reference/Concepts/overview-of-the-office-fluent-ribbon.md', 'Ribbon'),
    ('05-excel-ribbon-customui', 'ribbon-por-arquivo-open-xml.md', VBA + 'Library-Reference/Concepts/customize-the-office-fluent-ribbon-by-using-an-open-xml-formats-file.md', 'customUI no pacote'),
    ('05-excel-ribbon-customui', 'iribbonui.md', VBA + 'api/Office.IRibbonUI.md', 'IRibbonUI'),
    ('05-excel-ribbon-customui', 'iribboncontrol.md', VBA + 'api/Office.IRibbonControl.md', 'IRibbonControl'),
    # ------------------------------------------------ SQL do Access (motor ACE)
    *[('06-access-sql-ace', n + '.md', VBA + 'access/Concepts/Structured-Query-Language/' + n + '.md', n) for n in [
        'create-and-delete-tables-and-indexes-using-access-sql', 'define-relationships-between-tables-using-access-sql',
        'modify-a-table-s-design-using-access-sql', 'insert-update-and-delete-records-from-a-table-using-access-sql',
        'retrieve-records-using-access-sql', 'all-distinct-distinctrow-top-predicates-microsoft-access-sql',
        'where-clause-microsoft-access-sql', 'order-by-clause-microsoft-access-sql', 'group-by-clause-microsoft-access-sql',
        'having-clause-microsoft-access-sql', 'from-clause-microsoft-access-sql', 'perform-joins-using-access-sql',
        'like-operator-microsoft-access-sql', 'use-aggregate-functions-to-work-with-values-in-access-sql',
        'use-international-date-formats-in-sql-statements', 'build-sql-statements-that-include-variables-and-controls']],
    ('06-access-sql-ace', 'access-database-engine-2016-redistributable.html.md',
     'https://www.microsoft.com/en-us/download/details.aspx?id=54920', 'ACE redistribuível'),
    # ------------------------------------------------ ADO (front VBA)
    *[('07-ado', n + '.html.md', ADO + n + '?view=sql-server-ver15', n) for n in [
        'connection-object-ado', 'command-object-ado', 'parameter-object', 'createparameter-method-ado',
        'parameters-collection-ado', 'append-method-ado', 'datatypeenum', 'execute-method-ado-command',
        'begintrans-committrans-and-rollbacktrans-methods-ado', 'recordset-object-ado', 'getrows-method-ado',
        'commandtypeenum', 'cursorlocationenum']],
    # ------------------------------------------------ .NET System.Data.OleDb (scripts PowerShell)
    *[('08-dotnet-oledb', n + '.md', NET + n + '.md', n) for n in [
        'ole-db-odbc-and-oracle-connection-pooling', 'configuring-parameters-and-parameter-data-types',
        'ole-db-data-type-mappings', 'local-transactions', 'retrieving-identity-or-autonumber-values']],
    *[('08-dotnet-oledb', n + '.html.md', LEARN + 'dotnet/api/' + n + '?view=netframework-4.8.1', n) for n in [
        'system.data.oledb.oledbtype', 'system.data.oledb.oledbparameter', 'system.data.oledb.oledbconnection.connectionstring',
        'system.data.oledb.oledbtransaction', 'system.data.oledb.oledbenumerator']],
    # ------------------------------------------------ Windows PowerShell 5.1
    ('09-powershell-5.1', 'ConvertTo-Json.md', PS + 'Microsoft.PowerShell.Utility/ConvertTo-Json.md', 'ConvertTo-Json'),
    ('09-powershell-5.1', 'ConvertFrom-Json.md', PS + 'Microsoft.PowerShell.Utility/ConvertFrom-Json.md', 'ConvertFrom-Json'),
    *[('09-powershell-5.1', n + '.md', PS + 'Microsoft.PowerShell.Core/About/' + n + '.md', n) for n in [
        'about_Character_Encoding', 'about_Execution_Policies', 'about_PowerShell_exe', 'about_Scripts', 'about_Scopes',
        'about_Script_Blocks', 'about_Try_Catch_Finally', 'about_Automatic_Variables', 'about_Preference_Variables']],
    # ------------------------------------------------ Edge / Chromium headless
    ('10-edge-headless', 'chrome-headless.html.md', 'https://developer.chrome.com/docs/chromium/headless', 'Headless'),
    # ------------------------------------------------ Windows: ícones e DPI
    ('11-windows-icones-dpi', 'segoe-mdl2-assets.md', WIN + 'segoe-ui-symbol-font.md', 'Segoe MDL2 Assets'),
    ('11-windows-icones-dpi', 'segoe-fluent-icons.md', WIN + 'segoe-fluent-icons-font.md', 'Segoe Fluent Icons'),
    ('11-windows-icones-dpi', 'getdevicecaps.md', SDK + 'wingdi/nf-wingdi-getdevicecaps.md', 'GetDeviceCaps'),
    ('11-windows-icones-dpi', 'getdc.md', SDK + 'winuser/nf-winuser-getdc.md', 'GetDC'),
    ('11-windows-icones-dpi', 'releasedc.md', SDK + 'winuser/nf-winuser-releasedc.md', 'ReleaseDC'),
    ('11-windows-icones-dpi', 'getdpiforsystem.md', SDK + 'winuser/nf-winuser-getdpiforsystem.md', 'GetDpiForSystem'),
    # ------------------------------------------------ MCP - revisão atual (2026-07-28) + legado (2025-11-25)
    *[('12-mcp-especificacao', '2026-07-28_' + n.replace('/', '_').replace('.mdx', '.md'), MCP + 'docs/specification/2026-07-28/' + n, n) for n in [
        'index.mdx', 'changelog.mdx', 'deprecated.mdx', 'basic/index.mdx', 'basic/versioning.mdx', 'basic/transports/index.mdx',
        'basic/transports/stdio.mdx', 'basic/patterns/cancellation.mdx', 'server/index.mdx', 'server/discover.mdx',
        'server/tools.mdx', 'server/utilities/caching.mdx', 'server/utilities/pagination.mdx']],
    ('12-mcp-especificacao', '2026-07-28_schema.ts', MCP + 'schema/2026-07-28/schema.ts', 'schema.ts'),
    *[('12-mcp-especificacao', '2025-11-25_' + n.replace('/', '_').replace('.mdx', '.md'), MCP + 'docs/specification/2025-11-25/' + n, n) for n in [
        'changelog.mdx', 'basic/lifecycle.mdx', 'basic/transports.mdx', 'server/tools.mdx']],
    # ------------------------------------------------ Claude Desktop: MCP local, extensões, skills
    ('13-claude-desktop', 'conectar-servidores-locais.md', MCP + 'docs/docs/2026-07-28/develop/connect-local-servers.mdx', 'MCP local no Claude Desktop'),
    ('13-claude-desktop', 'servidores-locais-help-center.html.md', 'https://support.claude.com/en/articles/10949351-getting-started-with-local-mcp-servers-on-claude-desktop', 'Help Center'),
    ('13-claude-desktop', 'mcpb-manifest.md', GH + 'modelcontextprotocol/mcpb/main/MANIFEST.md', 'MCPB manifest'),
    ('13-claude-desktop', 'mcpb-readme.md', GH + 'modelcontextprotocol/mcpb/main/README.md', 'MCPB'),
    ('13-claude-desktop', 'agent-skills-visao-geral.md', 'https://docs.claude.com/en/docs/agents-and-tools/agent-skills/overview.md', 'Skills'),
    ('13-claude-desktop', 'agent-skills-boas-praticas.md', 'https://docs.claude.com/en/docs/agents-and-tools/agent-skills/best-practices.md', 'Skills: boas práticas'),
    ('13-claude-desktop', 'usar-skills-no-claude.html.md', 'https://support.claude.com/en/articles/12512180-using-skills-in-claude', 'Skills no app'),
    # ------------------------------------------------ Office 2019
    ('14-office-2019', 'ciclo-de-vida-office-2019.html.md', LEARN + 'lifecycle/products/microsoft-office-2019', 'Ciclo de vida'),
    ('14-office-2019', 'novidades-excel-2019.html.md', 'https://support.microsoft.com/en-us/office/what-s-new-in-excel-2019-for-windows-5a201203-1155-4055-82a5-82bf0994631f', 'Novidades Excel 2019'),
]


def baixar(url):
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (crm-zapromaq doc sync)', 'Accept': 'text/markdown, text/html;q=0.9, */*;q=0.8'})
    try:
        with _ABRIDOR.open(req, timeout=60) as r:
            return r.read().decode('utf-8', errors='replace')
    except urllib.error.HTTPError as ex:
        if ex.code not in (301, 302, 303, 307, 308):
            raise
        # algumas páginas do Learn entram em laço de redirecionamento no urllib; o curl resolve
        import subprocess
        r = subprocess.run(['curl', '-sSfL', '--max-redirs', '10', url], capture_output=True, timeout=90)
        if r.returncode != 0:
            raise RuntimeError('curl falhou: ' + r.stderr.decode('utf-8', 'replace').strip())
        return r.stdout.decode('utf-8', errors='replace')


def html_para_md(html):
    from bs4 import BeautifulSoup
    from markdownify import markdownify
    sopa = BeautifulSoup(html, 'html.parser')
    for lixo in sopa.select('script, style, nav, header, footer, aside, form, button, noscript, svg, '
                            '[data-bi-name="feedback-section"], .feedback-section, #ms--additional-resources, '
                            '.page-metadata-container, .breadcrumbs, #article-header, .modular-content-container'):
        lixo.decompose()
    corpo = sopa.select_one('main') or sopa.select_one('article') or sopa.body or sopa
    md = markdownify(str(corpo), heading_style='ATX', bullets='-')
    md = re.sub(r'\n{3,}', '\n\n', md)
    return md.strip() + '\n'


def main(filtros):
    indice, falhas = [], []
    for pasta, arquivo, url, titulo in FONTES:
        if filtros and not any(f.lower() in pasta.lower() for f in filtros):
            continue
        destino = os.path.join(BASE, pasta, arquivo)
        os.makedirs(os.path.dirname(destino), exist_ok=True)
        try:
            bruto = baixar(url)
            e_html = arquivo.endswith('.html.md') or bruto.lstrip()[:15].lower().startswith(('<!doctype', '<html'))
            conteudo = html_para_md(bruto) if e_html else bruto
            if len(conteudo.strip()) < 200:
                raise ValueError('conteúdo vazio ou muito curto')
            cab = ('<!-- fonte: ' + url + ' | obtido em: ' + HOJE +
                   (' | convertido de HTML' if e_html else ' | markdown original') + ' -->\n\n')
            with open(destino, 'w', encoding='utf-8', newline='\n') as f:
                f.write(cab + conteudo)
            indice.append((pasta, arquivo, titulo, url))
            print('OK   ', pasta + '/' + arquivo)
        except Exception as ex:  # segue para o próximo; a falha vai para o relatório
            falhas.append((pasta, arquivo, url, str(ex)))
            print('FALHA', pasta + '/' + arquivo, '->', ex)

    with open(os.path.join(BASE, 'fontes-baixadas.json'), 'w', encoding='utf-8') as f:
        json.dump({'obtido_em': HOJE, 'arquivos': [dict(pasta=p, arquivo=a, titulo=t, fonte=u) for p, a, t, u in indice],
                   'falhas': [dict(pasta=p, arquivo=a, fonte=u, erro=e) for p, a, u, e in falhas]},
                  f, ensure_ascii=False, indent=1)
    print('\n%d arquivos, %d falhas' % (len(indice), len(falhas)))
    return 1 if falhas else 0


if __name__ == '__main__':
    sys.exit(main(sys.argv[1:]))
