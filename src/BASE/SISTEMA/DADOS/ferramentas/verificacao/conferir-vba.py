# -*- coding: utf-8 -*-
"""
conferir-vba.py - verificacao estatica das fontes VBA.

Nao substitui o Depurar > Compilar do VBE, mas pega antes o que
custaria uma ida e volta: bloco sem fechamento, chamada a rotina
que nao existe, continuacao de linha quebrada, aspas desbalanceadas.

Uso: python3 conferir-vba.py "<pasta 10 - FONTES-VBA>"
"""
import sys, os, re, collections

ABRE = [
    (re.compile(r"^\s*(Public |Private |Friend )?(Static )?Sub\b", re.I), "Sub", re.compile(r"^\s*End Sub\b", re.I)),
    (re.compile(r"^\s*(Public |Private |Friend )?(Static )?Function\b", re.I), "Function", re.compile(r"^\s*End Function\b", re.I)),
    (re.compile(r"^\s*(Public |Private )?Property\b", re.I), "Property", re.compile(r"^\s*End Property\b", re.I)),
]
RE_SELECT = re.compile(r"^\s*Select Case\b", re.I)
RE_ENDSEL = re.compile(r"^\s*End Select\b", re.I)
RE_WITH   = re.compile(r"^\s*With\b", re.I)
RE_ENDW   = re.compile(r"^\s*End With\b", re.I)
RE_FOR    = re.compile(r"^\s*For\b", re.I)
RE_NEXT   = re.compile(r"^\s*Next\b", re.I)
RE_DO     = re.compile(r"^\s*Do\b", re.I)
RE_LOOP   = re.compile(r"^\s*Loop\b", re.I)
RE_IF     = re.compile(r"^\s*If\b.*\bThen\s*(?:'.*)?$", re.I)
RE_ENDIF  = re.compile(r"^\s*End If\b", re.I)
RE_DECL   = re.compile(r"^\s*(?:Public |Private |Friend )?(?:Static )?(Sub|Function|Property (?:Get|Let|Set))\s+(\w+)", re.I)
RE_CHAMADA= re.compile(r"\b(mod\w+)\.(\w+)")

def sem_comentario(linha):
    fora = True
    saida = []
    for ch in linha:
        if ch == '"':
            fora = not fora
        if ch == "'" and fora:
            break
        saida.append(ch)
    return "".join(saida)

def analisar(caminho):
    nome = os.path.basename(caminho)
    txt = open(caminho, encoding="cp1252", errors="replace").read()
    linhas = txt.split("\r\n") if "\r\n" in txt else txt.split("\n")
    # junta continuacao ' _' : sem isso, um If cujo Then cai na linha
    # seguinte parece nao abrir bloco
    juntas, buffer, origem = [], "", 0
    for idx, bruta in enumerate(linhas, 1):
        if not buffer:
            origem = idx
        corpo = sem_comentario(bruta)
        if corpo.rstrip().endswith(" _"):
            buffer += corpo.rstrip()[:-1]
            continue
        juntas.append((origem, buffer + corpo))
        buffer = ""
    if buffer:
        juntas.append((origem, buffer))
    problemas = []
    declaradas = set()
    tem_option = any(l.strip().lower() == "option explicit" for l in linhas)
    if not tem_option:
        problemas.append("sem Option Explicit")

    pilha = []
    cont_sel = cont_with = cont_for = cont_do = cont_if = 0
    dentro = None

    for i, bruta in enumerate(linhas, 1):
        if sem_comentario(bruta).count('"') % 2 == 1 and not bruta.rstrip().endswith(" _"):
            problemas.append("linha %d: aspas desbalanceadas -> %s" % (i, bruta.strip()[:70]))
        if sem_comentario(bruta).rstrip().endswith(" _"):
            if i >= len(linhas) or not linhas[i].strip():
                problemas.append("linha %d: continuacao ' _' sem linha seguinte" % i)

    for i, l in juntas:
        mv = re.match(r"^\s*Public\s+(?:Const\s+)?(\w+)\s*(?:As\b|=)", l, re.I)
        if mv:
            declaradas.add(mv.group(1).lower())
        m = RE_DECL.match(l)
        if m:
            declaradas.add(m.group(2).lower())
            if dentro:
                problemas.append("linha %d: %s comeca antes de fechar %s" % (i, m.group(2), dentro))
            dentro = m.group(2)
            continue

        for rx, tipo, rxfim in ABRE:
            if rxfim.match(l):
                dentro = None
        if RE_SELECT.match(l): cont_sel += 1
        if RE_ENDSEL.match(l): cont_sel -= 1
        if RE_WITH.match(l): cont_with += 1
        if RE_ENDW.match(l): cont_with -= 1
        if RE_FOR.match(l): cont_for += 1
        if RE_NEXT.match(l): cont_for -= 1
        if RE_DO.match(l) and not re.match(r"^\s*Do\s*$", l, re.I) is None or RE_DO.match(l): cont_do += 1
        if RE_LOOP.match(l): cont_do -= 1
        if RE_IF.match(l): cont_if += 1
        if RE_ENDIF.match(l): cont_if -= 1

    # limite do VBA: 25 continuacoes de linha por instrucao, e 1023
    # caracteres por linha fisica. Estourar da erro so na importacao.
    seq = 0
    inicio_seq = 0
    for i, bruta in enumerate(linhas, 1):
        if len(bruta) > 1023:
            problemas.append("linha %d: %d caracteres (limite do VBA: 1023)" % (i, len(bruta)))
        if sem_comentario(bruta).rstrip().endswith(" _"):
            if seq == 0:
                inicio_seq = i
            seq += 1
        else:
            if seq > 24:
                problemas.append("linha %d: instrucao com %d continuacoes de linha (limite do VBA: 25)"
                                 % (inicio_seq, seq))
            seq = 0
    if seq > 24:
        problemas.append("linha %d: instrucao com %d continuacoes de linha (limite do VBA: 25)" % (inicio_seq, seq))

    if dentro: problemas.append("procedimento '%s' nao foi fechado" % dentro)
    for rotulo, saldo in (("Select Case", cont_sel), ("With", cont_with), ("For", cont_for),
                          ("Do", cont_do), ("If (bloco)", cont_if)):
        if saldo != 0:
            problemas.append("%s desbalanceado: saldo %+d" % (rotulo, saldo))
    return nome, declaradas, problemas, txt

def main(pasta):
    arquivos = sorted([os.path.join(pasta, f) for f in os.listdir(pasta)
                       if f.endswith(".bas") or f.endswith(".txt")])
    publicas = {}
    textos = {}
    todos = []
    for caminho in arquivos:
        nome, decl, probs, txt = analisar(caminho)
        mod = nome.rsplit(".", 1)[0]
        publicas[mod.lower()] = decl
        textos[mod] = txt
        todos.append((nome, probs))

    # nome publico repetido entre modulos: o VBA recusa com
    # "Nome repetido encontrado" quando algo chama pelo nome simples
    # (OnAction de botao, Application.Run, chamada sem qualificador)
    # Private nao colide: o nome so existe dentro do proprio modulo.
    # Sem este filtro, cmdFechar_Click de dois formularios aparecia
    # como conflito - e nao e.
    RE_PRIV = re.compile(r"^\s*Private\s+(?:Static\s+)?(?:Sub|Function|Property)\s+(\w+)", re.I | re.M)
    dono = collections.defaultdict(list)
    for mod, decl in publicas.items():
        # publicas usa a chave em minusculas e textos preserva o nome
        fonte = ""
        for k, vtxt in textos.items():
            if k.lower() == mod.lower():
                fonte = vtxt
                break
        # membro de classe ou de formulario nao e global: so existe
        # atraves de uma instancia (obj.Pintar), entao nao colide
        if "VERSION 1.0 CLASS" in fonte or mod.lower().startswith("frm") or mod.lower() == "thisworkbook":
            continue
        privadas = set(m.lower() for m in RE_PRIV.findall(fonte))
        for r in decl:
            if r.lower() in privadas:
                continue
            dono[r].append(mod)
    repetidos = {r: m for r, m in dono.items() if len(m) > 1}

    # chamadas qualificadas
    faltantes = []
    for mod, txt in textos.items():
        for i, bruta in enumerate(txt.replace("\r\n", "\n").split("\n"), 1):
            l = sem_comentario(bruta)
            for alvo, rotina in RE_CHAMADA.findall(l):
                a = alvo.lower()
                if a not in publicas:
                    faltantes.append("%s linha %d: modulo '%s' nao existe" % (mod, i, alvo))
                elif rotina.lower() not in publicas[a]:
                    # pode ser constante publica
                    if not re.search(r"Public Const\s+" + re.escape(rotina) + r"\b", textos.get(alvo, ""), re.I):
                        faltantes.append("%s linha %d: %s.%s nao encontrado" % (mod, i, alvo, rotina))

    repetidos_msg = []
    for r, mods in sorted(repetidos.items()):
        repetidos_msg.append("%s  ->  %s" % (r, ", ".join(sorted(mods))))

    # Declaracao de nivel de modulo tem de vir ANTES de qualquer
    # procedimento. No fim do modulo o VBA nao a reconhece e acusa
    # "Variavel nao definida" no ponto de uso - aconteceu em
    # frmCRM.txt em 17/09/2026, e este script dava OK no arquivo.
    RE_PROC0 = re.compile(r"^(?:Public |Private |Friend )?(?:Static )?(?:Sub|Function|Property)\s", re.I)
    RE_DECL0 = re.compile(r"^(?:Public|Private|Dim|Global)\s+(?:Const\s+)?\w+\s+As\s", re.I)
    fora_de_lugar = []
    for mod, txt in textos.items():
        primeiro = None
        for i, l in enumerate(txt.replace("\r\n", "\n").split("\n"), 1):
            if l[:1] in (" ", "\t"):
                continue
            if primeiro is None and RE_PROC0.match(l):
                primeiro = i
            if primeiro and RE_DECL0.match(l):
                fora_de_lugar.append("%s linha %d: declaracao de modulo depois do procedimento da linha %d"
                                     % (mod, i, primeiro))
                break

    # Parametro ou Dim com o mesmo nome de uma rotina do modulo, usado
    # como chamada dentro dela. O VBA NAO diferencia maiusculas em
    # identificador: dentro da rotina o nome local ganha, e a chamada
    # vira indexacao - "Era esperada uma matriz". Aconteceu em
    # frmCRM.VinculoAtual(campo) chamando Campo() em 18/09/2026.
    RE_CAB = re.compile(r"^(?:Public |Private |Friend )?(?:Static )?(?:Sub|Function|Property\s+\w+)\s+(\w+)\s*\(([^)]*)", re.I)
    RE_PARAM = re.compile(r"(?:ByVal|ByRef|Optional\s+ByVal|Optional\s+ByRef|Optional)\s+(\w+)", re.I)
    sombras = []
    for mod, txt in textos.items():
        linhas_m = txt.replace("\r\n", "\n").split("\n")
        rotinas = set()
        for l in linhas_m:
            m = RE_CAB.match(l)
            if m:
                rotinas.add(m.group(1).lower())
        atual, locais, usados, linha_ini = None, set(), set(), 0

        def fechar():
            for nome in sorted(locais & usados & rotinas):
                if nome != atual:
                    sombras.append("%s linha %d: em %s, o nome local '%s' encobre a rotina %s() do modulo"
                                   % (mod, linha_ini, atual, nome, nome))

        for i, l in enumerate(linhas_m, 1):
            m = RE_CAB.match(l)
            if m:
                if atual:
                    fechar()
                atual, linha_ini = m.group(1).lower(), i
                locais = set(x.lower() for x in RE_PARAM.findall(m.group(2)))
                usados = set()
            elif atual:
                locais |= set(x.lower() for x in re.findall(r"^\s*Dim\s+(\w+)", l))
                usados |= set(x.lower() for x in re.findall(r"\b(\w+)\s*\(", sem_comentario(l)))
        if atual:
            fechar()

    print("=== VERIFICACAO ESTATICA DAS FONTES VBA ===")
    print()
    erros = 0
    for nome, probs in todos:
        marca = "OK" if not probs else "%d problema(s)" % len(probs)
        print("%-22s %s" % (nome, marca))
        for p in probs:
            erros += 1
            print("      -", p)
    print()
    if repetidos_msg:
        print("NOME PUBLICO REPETIDO EM MAIS DE UM MODULO (%d):" % len(repetidos_msg))
        for r in repetidos_msg:
            print("   -", r)
        print("   (o VBA recusa a chamada pelo nome simples: renomeie um deles)")
        erros += len(repetidos_msg)
        print()

    if sombras:
        print("NOME LOCAL ENCOBRINDO ROTINA (%d):" % len(sombras))
        for f in sombras:
            print("   -", f)
        print("   (o VBA ignora maiusculas: renomeie o parametro ou o Dim)")
        erros += len(sombras)
        print()

    if fora_de_lugar:
        print("DECLARACAO DE MODULO FORA DE LUGAR (%d):" % len(fora_de_lugar))
        for f in fora_de_lugar:
            print("   -", f)
        print("   (Const e Dim de modulo vao ANTES do primeiro Sub/Function)")
        erros += len(fora_de_lugar)
        print()

    if faltantes:
        print("CHAMADAS NAO RESOLVIDAS (%d):" % len(faltantes))
        for f in faltantes:
            print("   -", f)
        erros += len(faltantes)
    else:
        print("Todas as chamadas modX.Rotina apontam para rotina existente.")
    print()
    print("RESULTADO:", "sem problemas" if erros == 0 else "%d ponto(s) a conferir" % erros)
    return 0 if erros == 0 else 1

if __name__ == "__main__":
    sys.exit(main(sys.argv[1]))
