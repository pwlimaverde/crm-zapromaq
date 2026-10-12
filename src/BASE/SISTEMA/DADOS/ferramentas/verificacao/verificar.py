# -*- coding: utf-8 -*-
"""
verificar.py - verificação estática de TODO o projeto (máquina de desenvolvimento).

Não substitui a compilação no Excel (o build faz o autoteste na estação),
mas pega antes o que custaria uma ida à estação:

  VBA ........ codificação Windows-1252 + CRLF; SINTAXE por parser formal (ANTLR,
               gramática MS-VBAL - verificar-sintaxe-vba.py); blocos balanceados, chamadas
               modX.Rotina resolvidas, nomes públicos repetidos, declaração fora
               de lugar, nome local encobrindo rotina (lógica do conferir-vba.py)
  Esquema .... todo campo de modSchema existe no DDL (criar-banco.ps1) e na
               consulta da tabela em modCRM; colunas da grade idem
  PowerShell . sintaxe (analisador do Windows PowerShell 5.1) e BOM quando há acento
  .bat ....... CRLF e só ASCII
  JSON ....... sintaxe

Uso:  python verificar.py            (a partir de qualquer pasta)
Saída: 0 sem problemas; 1 com problemas.
"""
import os, re, sys, json, shutil, tempfile, subprocess, importlib.util, hashlib, unicodedata

AQUI = os.path.dirname(os.path.abspath(__file__))
DADOS = os.path.abspath(os.path.join(AQUI, '..', '..'))
BASE = os.path.abspath(os.path.join(DADOS, '..', '..'))
FONTE = os.path.join(DADOS, 'fonte')

problemas = []


def falha(grupo, msg):
    problemas.append((grupo, msg))


def rel(p):
    return os.path.relpath(p, BASE)


# ============================================================ VBA
def fontes_vba():
    saida = []
    for sub, exts in (('modulos', ('.bas',)), ('classes', ('.cls',)), ('formularios', ('.txt',)),
                      ('pasta-de-trabalho', ('.txt',))):
        pasta = os.path.join(FONTE, sub)
        if not os.path.isdir(pasta):
            continue
        for f in sorted(os.listdir(pasta)):
            if f.endswith(exts):
                saida.append(os.path.join(pasta, f))
    return saida


def verificar_codificacao_vba(arquivos):
    for p in arquivos:
        b = open(p, 'rb').read()
        if b.startswith(b'\xef\xbb\xbf'):
            falha('VBA', rel(p) + ': tem BOM UTF-8 - o VBE importa em Windows-1252 e o BOM vira lixo')
        try:
            b.decode('cp1252')
        except UnicodeDecodeError as ex:
            falha('VBA', rel(p) + ': não é Windows-1252 válido (' + str(ex) + ')')
            continue
        sem_crlf = b.replace(b'\r\n', b'')
        if b'\n' in sem_crlf or b'\r' in sem_crlf:
            falha('VBA', rel(p) + ': quebra de linha que não é CRLF')
        # Acento PERDIDO: editor que le o arquivo como UTF-8 troca cada byte
        # cp1252 (ex.: E3 = a com til) por U+FFFD e regrava em UTF-8 (EF BF BD).
        # Em cp1252 isso e texto valido ("ï¿½"), entao as checagens abaixo nao
        # pegavam. Aconteceu com a ferramenta Edit em 10/10/2026.
        if b'\xef\xbf\xbd' in b:
            falha('VBA', rel(p) + ': ' + str(b.count(b'\xef\xbf\xbd')) + ' caractere(s) U+FFFD (EF BF BD) - '
                  'acento perdido, o arquivo foi regravado como UTF-8; restaure do git e edite em cp1252')
        # Arquivo inteiro em UTF-8 (tem acento e decodifica como UTF-8): em cp1252
        # real isso praticamente nao acontece. Pega acento novo gravado em UTF-8.
        elif any(x > 127 for x in b):
            try:
                b.decode('utf-8')
                falha('VBA', rel(p) + ': acentos gravados em UTF-8 - o VBE importa em Windows-1252')
            except UnicodeDecodeError:
                pass
        # UTF-8 gravado por engano aparece como "Ã§", "Ã£" etc. em cp1252
        txt = b.decode('cp1252')
        if re.search('Ã[\u0080-\u00ff]', txt):
            falha('VBA', rel(p) + ': parece UTF-8 lido como Windows-1252 (sequência "Ã?")')
        if p.endswith('.bas') and not txt.startswith('Attribute VB_Name = "'):
            falha('VBA', rel(p) + ': .bas sem "Attribute VB_Name" na primeira linha')


def verificar_estrutura_vba(arquivos):
    spec = importlib.util.spec_from_file_location('conferir_vba', os.path.join(AQUI, 'conferir-vba.py'))
    cv = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(cv)
    tmp = tempfile.mkdtemp(prefix='verif-vba-')
    try:
        for p in arquivos:
            nome = os.path.basename(p)
            if nome.endswith('.cls'):
                nome = nome[:-4] + '.bas'
            shutil.copyfile(p, os.path.join(tmp, nome))
        import io, contextlib
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            rc = cv.main(tmp)
        if rc != 0:
            for linha in buf.getvalue().splitlines():
                l = linha.strip()
                if l.startswith('-') or l.startswith('linha') or 'problema(s)' in l:
                    falha('VBA', l.lstrip('- ').strip())
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def verificar_sintaxe_vba(arquivos):
    """Parser formal (ANTLR, gramática MS-VBAL) - ver verificar-sintaxe-vba.py."""
    try:
        spec = importlib.util.spec_from_file_location('sintaxe_vba', os.path.join(AQUI, 'verificar-sintaxe-vba.py'))
        sv = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(sv)
    except ImportError as ex:
        falha('VBA (sintaxe)', 'analisador sintático indisponível (' + str(ex) + '): pip install antlr4-python3-runtime==4.13.2')
        return
    for p in arquivos:
        for linha, col, msg in sv.analisar(p):
            falha('VBA (sintaxe)', rel(p) + ' linha ' + str(linha) + ', coluna ' + str(col) + ': ' + msg)


# ============================================================ ESQUEMA x SQL x modSchema
def ler_cp1252(p):
    return open(p, encoding='cp1252').read().replace('\r\n', '\n')


def ddl_colunas():
    txt = open(os.path.join(DADOS, 'banco', 'esquema', 'criar-banco.ps1'), encoding='utf-8-sig').read()
    # junta as strings concatenadas do PowerShell: "...' +\n    '..."
    junto = re.sub(r"'\s*\+\s*\n\s*'", '', txt)
    tabelas = {}
    for m in re.finditer(r"CREATE TABLE (\w+) \((.*?)\)'", junto, re.S):
        cols = []
        profundidade, atual = 0, ''
        for ch in m.group(2):
            if ch == '(':
                profundidade += 1
            elif ch == ')':
                profundidade -= 1
            if ch == ',' and profundidade == 0:
                cols.append(atual)
                atual = ''
            else:
                atual += ch
        cols.append(atual)
        tabelas[m.group(1).lower()] = {c.strip().split()[0].lower() for c in cols if c.strip()}
    return tabelas


def bloco_funcao(txt, nome):
    m = re.search(r'Public Function ' + nome + r'\b(.*?)\nEnd Function', txt, re.S)
    return m.group(1) if m else ''


def campos_schema(txt, funcao, tabela):
    corpo = bloco_funcao(txt, funcao)
    m = re.search(r'Case "' + tabela + r'"(.*?)(?:\n\s*Case "|\n\s*End Select)', corpo, re.S)
    if not m:
        return []
    return [x.split('|') for x in re.findall(r'Add c, "([^"]+)"', m.group(1))]


def colunas_select(txt, funcao):
    corpo = bloco_funcao(txt, funcao)
    # junta a string VBA concatenada e isola o SELECT ... FROM
    junto = ''.join(re.findall(r'"([^"]*)"', corpo))
    m = re.search(r'SELECT (.*?) FROM', junto, re.S | re.I)
    if not m:
        return set()
    cols = set()
    for c in m.group(1).split(','):
        c = c.strip()
        am = re.search(r'\bAS\s+(\w+)$', c, re.I)
        cols.add((am.group(1) if am else c.split('.')[-1]).lower())
    return cols


def verificar_esquema():
    ddl = ddl_colunas()
    if not ddl:
        falha('Esquema', 'não consegui ler o DDL de banco/esquema/criar-banco.ps1')
        return
    schema = ler_cp1252(os.path.join(FONTE, 'modulos', 'modSchema.bas'))
    crm = ler_cp1252(os.path.join(FONTE, 'modulos', 'modCRM.bas'))
    consultas = {'clientes': 'SQLClientes', 'contatos': 'SQLContatos', 'oportunidades': 'SQLOportunidades'}
    calculados = {'codigo_visual', 'situacao', 'telefone_atendimento'}
    for tabela, funcao in consultas.items():
        sel = colunas_select(crm, funcao)
        if not sel:
            falha('Esquema', 'modCRM.' + funcao + ': SELECT não encontrado')
            continue
        for p in campos_schema(schema, 'CamposDe', tabela):
            if len(p) != 7:
                falha('Esquema', 'modSchema.CamposDe(' + tabela + '): especificação com ' + str(len(p)) + ' partes: ' + '|'.join(p))
                continue
            campo, tipo, editavel = p[0].lower(), p[2], p[6]
            if editavel == '1' and campo not in ddl.get(tabela, set()):
                falha('Esquema', tabela + '.' + campo + ': campo gravável da ficha não existe no DDL')
            if campo not in sel:
                falha('Esquema', tabela + '.' + campo + ': campo da ficha não vem na consulta modCRM.' + funcao)
        for func_col in ('ColunasDe', 'ColunasPlanilha'):
            for p in campos_schema(schema, func_col, tabela):
                campo = p[0].lower()
                if campo not in sel and campo not in calculados:
                    falha('Esquema', 'modSchema.' + func_col + '(' + tabela + '): coluna "' + campo + '" não vem em modCRM.' + funcao)


# ============================================================ PowerShell, .bat, JSON
def arquivos(ext, raiz=BASE):
    for d, _, fs in os.walk(raiz):
        if os.sep + 'execucao' in d or os.sep + '.git' in d:
            continue
        for f in fs:
            if f.lower().endswith(ext):
                yield os.path.join(d, f)


def verificar_powershell():
    lista = sorted(arquivos('.ps1'))
    for p in lista:
        b = open(p, 'rb').read()
        try:
            b.decode('utf-8')
        except UnicodeDecodeError:
            falha('PowerShell', rel(p) + ': não é UTF-8')
            continue
        tem_acento = any(c > 127 for c in b.replace(b'\xef\xbb\xbf', b''))
        if tem_acento and not b.startswith(b'\xef\xbb\xbf'):
            falha('PowerShell', rel(p) + ': tem acento e não tem BOM (o Windows PowerShell 5.1 lê como ANSI)')
    ps = os.path.join(os.environ.get('SystemRoot', r'C:\Windows'), 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe')
    if not os.path.exists(ps):
        falha('PowerShell', 'powershell.exe 5.1 não encontrado: sintaxe não verificada')
        return
    alvo = json.dumps(lista)
    script = ('$arqs = ConvertFrom-Json @\'\n' + alvo + '\n\'@\n'
              'foreach ($a in $arqs) { $e = $null; $t = $null\n'
              '  [void][System.Management.Automation.Language.Parser]::ParseFile($a, [ref]$t, [ref]$e)\n'
              '  foreach ($x in $e) { [Console]::Out.WriteLine($a + "|" + $x.Extent.StartLineNumber + "|" + $x.Message) } }')
    r = subprocess.run([ps, '-NoProfile', '-NonInteractive', '-Command', '-'], input=script.encode('utf-8'),
                       capture_output=True, timeout=300)
    for linha in r.stdout.decode('utf-8', 'replace').splitlines():
        if '|' in linha:
            a, n, msg = linha.split('|', 2)
            falha('PowerShell', rel(a) + ' linha ' + n + ': ' + msg)


def verificar_bat():
    for p in sorted(arquivos('.bat')):
        if os.sep + 'migracao-3.6' in p:
            continue  # registro histórico, não é executado
        b = open(p, 'rb').read()
        if any(c > 127 for c in b):
            falha('.bat', rel(p) + ': tem caractere fora do ASCII (o cmd.exe usa a página OEM)')
        if b.replace(b'\r\n', b'').count(b'\n'):
            falha('.bat', rel(p) + ': quebra de linha que não é CRLF (o cmd.exe erra rótulos)')


def verificar_json():
    for p in sorted(arquivos('.json')):
        try:
            json.load(open(p, encoding='utf-8'))
        except Exception as ex:
            falha('JSON', rel(p) + ': ' + str(ex))


# ============================================================ formulários: código x layout
# Todo controle citado no código de um formulário (cmdSalvar.Enabled,
# cmdSalvar_Click, bg_seloStatus...) tem de existir no layout JSON:
# o build cria os controles a partir dele, e um nome que falta só
# apareceria como "Variável não definida" na estação.
RE_CONTROLE = re.compile(r'\b((?:bg_|ico_)?(?:cmd|lbl|txt|cbo|chk|fra|lst|grd|ico|selo|opt)[A-Z][A-Za-z0-9]*)')


def nomes_do_layout(d):
    nomes = set(c['nome'] for c in d.get('controles', []))
    for k in d.get('componentes', []):
        nomes.add(k['nome'])
        if k['tipo'] in ('botao', 'selo'):
            nomes.add('bg_' + k['nome'])
        if k['tipo'] == 'botao' and k.get('icone'):
            nomes.add('ico_' + k['nome'])
    return nomes


def verificar_formularios():
    pasta = os.path.join(FONTE, 'layout', 'formularios')
    for arq in sorted(os.listdir(pasta)):
        if not arq.endswith('.json'):
            continue
        d = json.load(open(os.path.join(pasta, arq), encoding='utf-8'))
        cod = os.path.join(FONTE, d['codigo'].replace('/', os.sep))
        if not os.path.exists(cod):
            falha('Formulários', arq + ': código ' + d['codigo'] + ' não existe')
            continue
        nomes = nomes_do_layout(d)
        vistos = set()
        for linha in ler_cp1252(cod).split('\n'):
            linha = re.sub(r'"[^"]*"', '""', linha)          # strings fora
            linha = linha.split("'")[0]                       # comentário fora
            for n in RE_CONTROLE.findall(linha):
                if n not in nomes and n not in vistos:
                    vistos.add(n)
                    falha('Formulários', '%s: "%s" usado no código e ausente do layout' % (d['codigo'], n))


MUNICIPIOS_TOTAL_OFICIAL = 5571          # IBGE localidades, consultado em 12/10/2026 (2024: 5570)
MUNICIPIOS_UFS = {'AC', 'AL', 'AP', 'AM', 'BA', 'CE', 'DF', 'ES', 'GO', 'MA', 'MT', 'MS', 'MG',
                  'PA', 'PB', 'PR', 'PE', 'PI', 'RJ', 'RN', 'RS', 'RO', 'RR', 'SC', 'SP', 'SE', 'TO'}
MUNICIPIOS_CABECALHO = 'cod_ibge;nome;nome_busca;uf;lat;lon'


def nome_busca_de(nome):
    # mesma regra de modValidacao.ChaveLista / Regras.ps1: sem acento, maiúsculas, espaços colapsados
    sem = ''.join(c for c in unicodedata.normalize('NFKD', nome) if not unicodedata.combining(c))
    return ' '.join(sem.upper().split())


def verificar_municipios():
    ref = os.path.join(DADOS, 'banco', 'referencia')
    csv_p = os.path.join(ref, 'municipios.csv')
    fonte_p = os.path.join(ref, 'municipios.fonte.json')
    if not os.path.isfile(csv_p) or not os.path.isfile(fonte_p):
        falha('Municipios', 'municipios.csv ou municipios.fonte.json ausente em banco\\referencia')
        return
    bruto = open(csv_p, 'rb').read()
    if not bruto.startswith(b'\xef\xbb\xbf'):
        falha('Municipios', 'municipios.csv sem BOM UTF-8')
    try:
        fonte = json.loads(open(fonte_p, 'rb').read().decode('utf-8'))
    except (ValueError, UnicodeDecodeError) as e:
        falha('Municipios', 'municipios.fonte.json inválido: ' + str(e))
        return
    if hashlib.sha256(bruto).hexdigest() != fonte.get('csv_sha256'):
        falha('Municipios', 'SHA-256 do municipios.csv diferente do registrado no .fonte.json')
    linhas = bruto.decode('utf-8-sig').replace('\r\n', '\n').split('\n')
    if linhas and linhas[-1] == '':
        linhas.pop()
    if linhas[0] != MUNICIPIOS_CABECALHO:
        falha('Municipios', 'cabeçalho do CSV diferente de ' + MUNICIPIOS_CABECALHO)
    registros = linhas[1:]
    if len(registros) != MUNICIPIOS_TOTAL_OFICIAL:
        falha('Municipios', 'total de linhas %d; esperado %d (IBGE)' % (len(registros), MUNICIPIOS_TOTAL_OFICIAL))
    if fonte.get('total') != len(registros):
        falha('Municipios', '.fonte.json: total %s não bate com %d linhas' % (fonte.get('total'), len(registros)))
    codigos = set()
    ufs = set()
    for n, linha in enumerate(registros, 2):
        c = linha.split(';')
        if len(c) != 6:
            falha('Municipios', 'linha %d com %d campos (esperado 6)' % (n, len(c)))
            continue
        cod, nome, busca, uf, lat, lon = c
        if not re.fullmatch(r'\d{7}', cod):
            falha('Municipios', 'linha %d: cod_ibge não tem 7 dígitos' % n)
        elif cod in codigos:
            falha('Municipios', 'linha %d: cod_ibge %s repetido' % (n, cod))
        codigos.add(cod)
        if busca != nome_busca_de(nome):
            falha('Municipios', 'linha %d: nome_busca diferente da regra para "%s"' % (n, nome))
        if uf not in MUNICIPIOS_UFS:
            falha('Municipios', 'linha %d: UF "%s" fora das 27' % (n, uf))
        ufs.add(uf)
        try:
            la, lo = float(lat), float(lon)
        except ValueError:
            falha('Municipios', 'linha %d: coordenada não numérica' % n)
            continue
        if not (-34 <= la <= 6 and -74 <= lo <= -28):
            falha('Municipios', 'linha %d: coordenada fora do Brasil (%s, %s)' % (n, lat, lon))
    if len(ufs) != 27:
        falha('Municipios', 'UFs presentes: %d; esperado 27' % len(ufs))


def verificar_casos_regra():
    # Cada caso de regra (C01...) é provado nos dois lados: o autoteste do front (VBA) e o
    # teste do conector (PowerShell). Caso só de um lado = regra que pode divergir sem aviso.
    lados = {}
    arquivos_casos = {
        'front': os.path.join(FONTE, 'modulos', 'modAutoteste.bas'),
        'conector': os.path.join(BASE, 'IA', 'teste', 'testar-mcp.ps1'),
    }
    for lado, caminho in arquivos_casos.items():
        if not os.path.isfile(caminho):
            falha('Casos de regra', 'arquivo de teste ausente: ' + rel(caminho))
            return
        txt = ler_cp1252(caminho) if lado == 'front' else open(caminho, encoding='utf-8-sig').read()
        lados[lado] = set()
        for linha in txt.split('\n'):
            if re.search(r'\b(Confere|Conferir)\b', linha):
                lados[lado].update(re.findall(r'\bC(\d\d)\b', linha))
    for n in sorted(lados['front'] ^ lados['conector']):
        so = 'front' if n in lados['front'] else 'conector'
        falha('Casos de regra', 'C%s só tem teste no %s' % (n, so))


def main():
    vba = fontes_vba()
    verificar_casos_regra()
    verificar_formularios()
    verificar_codificacao_vba(vba)
    verificar_estrutura_vba(vba)
    verificar_sintaxe_vba(vba)
    verificar_esquema()
    verificar_powershell()
    verificar_bat()
    verificar_json()
    verificar_municipios()
    print('=== VERIFICAÇÃO ESTÁTICA - CRM Zapromaq ===')
    print('fontes VBA: %d   raiz: %s' % (len(vba), BASE))
    if not problemas:
        print('RESULTADO: sem problemas')
        return 0
    grupo_atual = None
    for g, m in problemas:
        if g != grupo_atual:
            print('\n[' + g + ']')
            grupo_atual = g
        print('  - ' + m)
    print('\nRESULTADO: %d problema(s)' % len(problemas))
    return 1


if __name__ == '__main__':
    sys.exit(main())
