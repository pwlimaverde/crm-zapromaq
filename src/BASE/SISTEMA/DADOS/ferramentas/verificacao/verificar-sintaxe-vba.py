# -*- coding: utf-8 -*-
"""
verificar-sintaxe-vba.py - análise SINTÁTICA real das fontes VBA, sem Excel.

Usa um parser gerado pelo ANTLR 4.13.2 a partir da gramática pública
"Visual Basic 7.1" (antlr/grammars-v4, vba/vba7_1), que segue a especificação
oficial da linguagem [MS-VBAL]. Pega o que a verificação por padrões de texto
(verificar.py) não pega: instrução malformada, bloco If/With/Select mal
fechado no meio de expressão, parêntese sobrando, palavra-chave fora de lugar.

Não substitui a compilação (tipos, nomes e referências só o VBE confere - o
build faz isso com o autoteste), mas elimina a ida à estação por erro de digitação.

Requisitos (só desenvolvimento): pip install antlr4-python3-runtime==4.13.2
O parser já gerado fica em antlr-vba/gerado. Para regerar (precisa de Java):
    java -jar antlr-vba/antlr-4.13.2-complete.jar -Dlanguage=Python3 -o antlr-vba/gerado antlr-vba/vbaLexer.g4 antlr-vba/vbaParser.g4

Uso:  python verificar-sintaxe-vba.py [pasta ou arquivos...]   (padrão: fonte/ do projeto)
"""
import os, sys

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(AQUI, 'antlr-vba', 'gerado'))
from antlr4 import InputStream, CommonTokenStream  # noqa: E402
from antlr4.error.ErrorListener import ErrorListener  # noqa: E402
from vbaLexer import vbaLexer  # noqa: E402
from vbaParser import vbaParser  # noqa: E402

FONTE = os.path.abspath(os.path.join(AQUI, '..', '..', 'fonte'))


class Coletor(ErrorListener):
    def __init__(self, deslocamento):
        super().__init__()
        self.erros = []
        self.deslocamento = deslocamento

    def syntaxError(self, recognizer, offendingSymbol, line, column, msg, e):
        self.erros.append((line - self.deslocamento, column, msg))


def texto_do_modulo(caminho):
    txt = open(caminho, 'rb').read().decode('cp1252').replace('\r\n', '\n').replace('\r', '\n')
    desloc = 0
    # código de formulário e da pasta de trabalho vem sem cabeçalho:
    # vira módulo com nome, para a gramática aceitar (linhas deslocadas de 1)
    if not txt.lstrip('\n').startswith('Attribute VB_Name') and not txt.startswith('VERSION'):
        nome = os.path.splitext(os.path.basename(caminho))[0]
        txt = 'Attribute VB_Name = "' + nome + '"\n' + txt
        desloc = 1
    if not txt.endswith('\n'):
        txt += '\n'
    return txt, desloc


def analisar(caminho):
    txt, desloc = texto_do_modulo(caminho)
    lexer = vbaLexer(InputStream(txt))
    coletor = Coletor(desloc)
    lexer.removeErrorListeners()
    lexer.addErrorListener(coletor)
    parser = vbaParser(CommonTokenStream(lexer))
    parser.removeErrorListeners()
    parser.addErrorListener(coletor)
    parser.startRule()
    return coletor.erros


def listar(alvos):
    arqs = []
    for a in alvos:
        if os.path.isdir(a):
            for d, _, fs in os.walk(a):
                for f in sorted(fs):
                    if f.endswith(('.bas', '.cls')) or (f.endswith('.txt') and ('formularios' in d or 'pasta-de-trabalho' in d or 'FONTES' in d)):
                        arqs.append(os.path.join(d, f))
        else:
            arqs.append(a)
    return arqs


def main(argv):
    alvos = argv or [FONTE]
    total = 0
    for p in listar(alvos):
        erros = analisar(p)
        if erros:
            print('%s: %d erro(s) de sintaxe' % (os.path.relpath(p), len(erros)))
            for linha, col, msg in erros[:10]:
                print('    linha %d, coluna %d: %s' % (linha, col, msg))
            total += len(erros)
        else:
            print('%s: ok' % os.path.relpath(p))
    print('\nRESULTADO:', 'sintaxe sem erros' if total == 0 else '%d erro(s) de sintaxe' % total)
    return 0 if total == 0 else 1


if __name__ == '__main__':
    sys.exit(main(sys.argv[1:]))
