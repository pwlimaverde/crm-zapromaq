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


# Estrutura nova do esquema 1.1 (migração 003). Espelha a spec 003 (Contratos > Tabelas).
ESQUEMA_11_TABELAS = {
    'municipios': ['cod_ibge', 'nome', 'nome_busca', 'uf', 'lat', 'lon', 'fonte_carga'],
    'planejamentos': ['id', 'codigo', 'controle', 'descricao', 'cod_municipio_base', 'raio_km', 'data_inicio',
                      'data_fim', 'responsavel', 'situacao', 'observacoes', 'versao', 'criado_em', 'alterado_em',
                      'alterado_por'],
    'visitas': ['id', 'id_oportunidade', 'id_planejamento', 'situacao', 'data_agendada', 'hora_agendada',
                'responsavel', 'tipo', 'observacoes', 'data_realizada', 'resultado', 'relato', 'motivo_cancelamento',
                'versao', 'criado_em', 'alterado_em', 'alterado_por'],
}
ESQUEMA_11_LISTAS_NOVAS = ['chave', 'encerra', 'resultado', 'quadro', 'exige_familia']
ESQUEMA_11_INDICES = ['ix_mun_uf_busca', 'ix_pln_codigo', 'ix_pln_controle', 'ix_vis_opo_sit', 'ix_vis_pln',
                      'ix_vis_data', 'ix_lis_tipo_chave']


def verificar_migracao_003():
    pasta = os.path.join(DADOS, 'banco', 'migracoes')
    arqs = [f for f in os.listdir(pasta) if f.startswith('003-') and f.endswith('.ps1')] if os.path.isdir(pasta) else []
    if len(arqs) != 1:
        falha('Esquema 1.1', 'migração 003 ausente em banco\\migracoes (003-*.ps1)')
        return
    txt = open(os.path.join(pasta, arqs[0]), encoding='utf-8-sig').read()
    junto = re.sub(r"'\s*\+\s*\n\s*'", '', txt)
    if "EsquemaDe = '1.0'" not in txt or "EsquemaPara = '1.1'" not in txt:
        falha('Esquema 1.1', '003: EsquemaDe/EsquemaPara devem ser 1.0 -> 1.1')
    tabelas = ddl_de_texto(junto)
    for tabela, colunas in ESQUEMA_11_TABELAS.items():
        if tabela not in tabelas:
            falha('Esquema 1.1', '003: tabela ' + tabela + ' não é criada')
            continue
        faltam = [c for c in colunas if c not in tabelas[tabela]]
        if faltam:
            falha('Esquema 1.1', '003: ' + tabela + ' sem colunas ' + ', '.join(faltam))
    for col in ESQUEMA_11_LISTAS_NOVAS:
        if col not in tabelas.get('listas', set()):
            falha('Esquema 1.1', '003: listas sem a coluna ' + col)
    for ix in ESQUEMA_11_INDICES:
        if ix not in junto:
            falha('Esquema 1.1', '003: índice ' + ix + ' não é criado')


def ddl_de_texto(junto):
    # mesma leitura de ddl_colunas, aplicada a um texto; ALTER TABLE ... ADD COLUMN soma colunas
    tabelas = {}
    for m in re.finditer(r"CREATE TABLE (\w+) \((.*?)\)'", junto, re.S):
        cols, prof, atual = [], 0, ''
        for ch in m.group(2):
            prof += (ch == '(') - (ch == ')')
            if ch == ',' and prof == 0:
                cols.append(atual)
                atual = ''
            else:
                atual += ch
        cols.append(atual)
        tabelas.setdefault(m.group(1).lower(), set()).update(c.strip().split()[0].lower() for c in cols if c.strip())
    for m in re.finditer(r"ALTER TABLE (\w+) ADD COLUMN (\w+)", junto, re.I):
        tabelas.setdefault(m.group(1).lower(), set()).add(m.group(2).lower())
    return tabelas


# Nomes de etapa que o programa NAO pode fixar em codigo: o papel vem da tabela listas (spec 003, D-c).
ETAPAS_FIXAS = ['Pedido Fechado', 'Perdido', 'Descartado', 'Contato Inicial', 'Sem Retorno', 'Retorno Agendado',
                'Elaboração da Proposta', 'Negociação', 'Proposta em Análise', 'Proposta em Stand By']
# Onde o nome pode aparecer: teste de regra com dados de exemplo e amostra visual do formulario de calibracao
ETAPAS_EXCECOES = {'modAutoteste.bas', 'frmCalibracao.txt'}


def sem_comentario(linha, comentario):
    # tira o comentario da linha, sem mexer no que esta dentro de aspas
    dentro = False
    for i, ch in enumerate(linha):
        if ch == '"':
            dentro = not dentro
        elif ch == comentario and not dentro:
            return linha[:i]
    return linha


def verificar_textos_de_etapa(vba):
    alvos = []
    for p in vba:
        if os.path.basename(p) not in ETAPAS_EXCECOES:
            alvos.append((p, "'"))
    pasta_ps = os.path.join(BASE, 'IA', 'mcp', 'lib')
    if os.path.isdir(pasta_ps):
        for nome in sorted(os.listdir(pasta_ps)):
            if nome.endswith('.ps1'):
                alvos.append((os.path.join(pasta_ps, nome), '#'))
    for caminho, marca in alvos:
        txt = open(caminho, encoding='cp1252' if caminho.endswith(('.bas', '.cls', '.txt')) else 'utf-8-sig',
                   errors='replace').read()
        for n, linha in enumerate(txt.replace('\r\n', '\n').split('\n'), 1):
            if '2. Negocia' in linha:            # titulo de secao do formulario, nao regra de etapa
                continue
            codigo = sem_comentario(linha, marca)
            for nome in ETAPAS_FIXAS:
                # so literal entre aspas conta ("Perdidos" de um rotulo nao e a etapa Perdido)
                if ('"' + nome + '"') in codigo or ("'" + nome + "'") in codigo:
                    falha('Etapas', rel(caminho) + ':' + str(n) + ': texto fixo de etapa "' + nome + '" no codigo')
                    break


def verificar_contrato_conector():
    # DEV-07: catálogo (Ferramentas.ps1), contrato do teste (testar-mcp.ps1) e o número que o instalador anuncia
    # são o mesmo conjunto. Ferramenta nova sem constar no teste = contrato mudado sem nota de mudança.
    cat = os.path.join(BASE, 'IA', 'mcp', 'lib', 'Ferramentas.ps1')
    tst = os.path.join(BASE, 'IA', 'teste', 'testar-mcp.ps1')
    ins = os.path.join(BASE, 'IA', 'mcp', 'instalar-mcp.ps1')
    if not all(os.path.isfile(p) for p in (cat, tst, ins)):
        falha('Contrato do conector', 'Ferramentas.ps1, testar-mcp.ps1 ou instalar-mcp.ps1 ausente')
        return
    catalogo = set(re.findall(r"name = '([a-z_]+)'; title", open(cat, encoding='utf-8-sig').read()))
    txt = open(tst, encoding='utf-8-sig').read()
    m = re.search(r'\$script:FerramentasEsperadas = @\((.*?)\)', txt, re.S)
    esperadas = set(re.findall(r"'([a-z_]+)'", m.group(1))) if m else set()
    if catalogo != esperadas:
        falha('Contrato do conector', 'catálogo x testar-mcp: só no catálogo %s; só no teste %s' %
              (sorted(catalogo - esperadas), sorted(esperadas - catalogo)))
    q = re.findall(r'\$qtdFerramentas = (\d+)', open(ins, encoding='utf-8-sig').read())
    if len(q) < 2 or int(q[-1]) != len(catalogo):
        falha('Contrato do conector', 'instalar-mcp.ps1 anuncia %s ferramentas; o catálogo tem %d' % (q[-1] if q else '?', len(catalogo)))


def bloco_ps(txt, nome):
    # texto da função PowerShell "function nome" até a chave que fecha (contagem simples de chaves)
    i = txt.find('function ' + nome)
    if i < 0:
        return None
    j = txt.find('{', i)
    prof = 0
    for k in range(j, len(txt)):
        prof += (txt[k] == '{') - (txt[k] == '}')
        if prof == 0:
            return txt[i:k + 1].replace('\r\n', '\n')
    return None


def verificar_espelho_backup():
    # Ciclo 007: o aviso de backup por janela existe em SISTEMA\DADOS\lib\Backup.ps1 e, copiado, no conector
    # (que não depende de SISTEMA em tempo de execução). As cópias têm de ser idênticas.
    lib = os.path.join(DADOS, 'lib', 'Backup.ps1')
    con = os.path.join(BASE, 'IA', 'mcp', 'lib', 'Ferramentas.ps1')
    if not (os.path.isfile(lib) and os.path.isfile(con)):
        return
    a = open(lib, encoding='utf-8-sig').read()
    b = open(con, encoding='utf-8-sig').read()
    for nome in ('Test-DiaUtil', 'Get-InicioJanela', 'Get-CopiasBackup', 'Get-AvisoBackupJanela'):
        x, y = bloco_ps(a, nome), bloco_ps(b, nome)
        if x is None or y is None or x != y:
            falha('Espelho do backup', nome + ': Backup.ps1 e Ferramentas.ps1 diferentes (ou ausente)')
    for const in ('JanelasBackup', 'JanelaEncerraMin', 'JanelaContaMin', 'FimDoDiaUtil'):
        rx = r'^\$script:' + const + r' = (.*)$'
        x = re.findall(rx, a, re.M)
        y = re.findall(rx, b, re.M)
        if not x or x != y:
            falha('Espelho do backup', '$script:' + const + ' diferente entre Backup.ps1 e Ferramentas.ps1')


def tamanhos_ddl():
    # {(tabela, campo): n} dos TEXT(n) do DDL e das migrações (ADD COLUMN / ALTER COLUMN mudam o tamanho)
    textos = [open(os.path.join(DADOS, 'banco', 'esquema', 'criar-banco.ps1'), encoding='utf-8-sig').read()]
    pasta = os.path.join(DADOS, 'banco', 'migracoes')
    for nome in sorted(os.listdir(pasta)) if os.path.isdir(pasta) else []:
        if nome.endswith('.ps1'):
            textos.append(open(os.path.join(pasta, nome), encoding='utf-8-sig').read())
    tam = {}
    for txt in textos:
        junto = re.sub(r"'\s*\+\s*\n\s*'", '', txt)
        for m in re.finditer(r"CREATE TABLE (\w+) \((.*?)\)'", junto, re.S):
            for c in re.finditer(r'(\w+) TEXT\((\d+)\)', m.group(2)):
                tam[(m.group(1).lower(), c.group(1).lower())] = int(c.group(2))
        for m in re.finditer(r"ALTER TABLE (\w+) (?:ADD|ALTER) COLUMN (\w+) TEXT\((\d+)\)", junto, re.I):
            tam[(m.group(1).lower(), m.group(2).lower())] = int(m.group(3))
    return tam


def verificar_matriz_campos():
    # DEV-10 (ciclo 009): a matriz versionada de campos é a referência dos três canais
    arq = os.path.join(DADOS, 'docs', 'matriz-campos.json')
    if not os.path.isfile(arq):
        falha('Matriz de campos', 'docs/matriz-campos.json ausente')
        return
    matriz = json.load(open(arq, encoding='utf-8'))
    ddl = tamanhos_ddl()
    ferr = open(os.path.join(BASE, 'IA', 'mcp', 'lib', 'Ferramentas.ps1'), encoding='utf-8-sig').read()
    m = re.search(r'\$script:TamanhoTexto = @\{(.*?)\n\}', ferr, re.S)
    tam_con = {k: int(v) for k, v in re.findall(r'(\w+) = (\d+)', m.group(1))} if m else {}
    for c in matriz.get('campos', []):
        t = c.get('tamanho')
        if t is None:
            continue
        chave = (c['tabela'], c['campo'])
        if chave in ddl and ddl[chave] != t:
            falha('Matriz de campos', '%s.%s: matriz %d, banco %d' % (c['tabela'], c['campo'], t, ddl[chave]))
        if c['campo'] in tam_con and tam_con[c['campo']] != t:
            falha('Matriz de campos', '%s.%s: matriz %d, conector (TamanhoTexto) %d' % (c['tabela'], c['campo'], t, tam_con[c['campo']]))
    listas = {c['campo'] for c in matriz.get('campos', []) if c.get('normalizacao') == 'lista'}
    reg = open(os.path.join(BASE, 'IA', 'mcp', 'lib', 'Regras.ps1'), encoding='utf-8-sig').read()
    m = re.search(r"\$script:CamposLista = @\((.*?)\)", reg)
    con = set(re.findall(r"'(\w+)'", m.group(1))) if m else set()
    val = ler_cp1252(os.path.join(FONTE, 'modulos', 'modValidacao.bas'))
    m = re.search(r'Case "etapa",(.*?)\n\s*s = ColapsarEspacos', val, re.S)
    fr = set(re.findall(r'"(\w+)"', 'Case "etapa",' + m.group(1))) if m else set()
    for nome, lado in (('conector (CamposLista)', con), ('front (Normalizar)', fr)):
        if lado != listas:
            falha('Matriz de campos', 'campos de lista da matriz x %s: só na matriz %s; só lá %s' % (nome, sorted(listas - lado), sorted(lado - listas)))


def verificar_esquema_atual():
    # Ciclo 010: o esquema exigido é UM só - última migração estrutural, front, conector e montagem
    pasta = os.path.join(DADOS, 'banco', 'migracoes')
    ultimo = ''
    for nome in sorted(os.listdir(pasta)):
        if nome.endswith('.ps1'):
            m = re.search(r"EsquemaPara = '([\d.]+)'", open(os.path.join(pasta, nome), encoding='utf-8-sig').read())
            if m:
                ultimo = m.group(1)
    fontes = {
        'modConfig.VERSAO_ESQUEMA': (ler_cp1252(os.path.join(FONTE, 'modulos', 'modConfig.bas')), r'VERSAO_ESQUEMA As String = "([\d.]+)"'),
        'Banco.ps1 EsquemasSuportados': (open(os.path.join(BASE, 'IA', 'mcp', 'lib', 'Banco.ps1'), encoding='utf-8-sig').read(), r"\$script:EsquemasSuportados = @\('([\d.]+)'\)"),
        'Comum.ps1 EsquemaAtual': (open(os.path.join(DADOS, 'lib', 'Comum.ps1'), encoding='utf-8-sig').read(), r"\$script:EsquemaAtual = '([\d.]+)'"),
    }
    for nome, (txt, rx) in fontes.items():
        m = re.search(rx, txt)
        if not m or m.group(1) != ultimo:
            falha('Esquema atual', '%s = %s; a última migração estrutural leva a %s' % (nome, m.group(1) if m else '?', ultimo))


def main():
    vba = fontes_vba()
    verificar_contrato_conector()
    verificar_esquema_atual()
    verificar_matriz_campos()
    verificar_espelho_backup()
    verificar_textos_de_etapa(vba)
    verificar_migracao_003()
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
