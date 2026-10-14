"""
gerar-municipios.py - prepara a base de municípios do IBGE (com coordenadas) que a migração 003 carrega.

Roda só na estação de desenvolvimento (precisa de internet). O resultado é versionado em
src\\BASE\\SISTEMA\\DADOS\\banco\\referencia\\ (municipios.csv + municipios.fonte.json); a montagem da
base só confere e consome esses arquivos, sem baixar nada.

Fontes:
  - IBGE, API de localidades (nome oficial, código de 7 dígitos, UF): dado público.
  - kelvins/municipios-brasileiros (licença MIT): latitude e longitude da sede de cada município.

Regras:
  - Cruzamento pelo código de 7 dígitos. Código presente numa fonte e ausente na outra → para (não chuta).
  - Nome: o do IBGE. Divergência de texto com a outra fonte é registrada no .fonte.json.
  - nome_busca: mesma regra de verificar.py (sem acento, maiúsculas, espaços colapsados).
  - Qualquer checagem que falhe → nada é gravado.

Uso: python gerar-municipios.py [--destino <pasta>]
"""
import argparse
import csv
import datetime
import gzip
import hashlib
import io
import json
import os
import sys
import unicodedata
import urllib.request

URL_IBGE = 'https://servicodados.ibge.gov.br/api/v1/localidades/municipios'
URL_COORDENADAS = 'https://raw.githubusercontent.com/kelvins/municipios-brasileiros/main/csv/municipios.csv'
TOTAL_OFICIAL = 5571          # IBGE localidades, consultado em 12/10/2026 (as duas fontes têm os mesmos 5571 códigos)
UFS = {'AC', 'AL', 'AP', 'AM', 'BA', 'CE', 'DF', 'ES', 'GO', 'MA', 'MT', 'MS', 'MG',
       'PA', 'PB', 'PR', 'PE', 'PI', 'RJ', 'RN', 'RS', 'RO', 'RR', 'SC', 'SP', 'SE', 'TO'}
AQUI = os.path.dirname(os.path.abspath(__file__))
DESTINO_PADRAO = os.path.abspath(os.path.join(AQUI, '..', '..', 'src', 'BASE', 'SISTEMA', 'DADOS', 'banco', 'referencia'))


def nome_busca_de(nome):
    sem = ''.join(c for c in unicodedata.normalize('NFKD', nome) if not unicodedata.combining(c))
    return ' '.join(sem.upper().split())


def baixar(url):
    req = urllib.request.Request(url, headers={'User-Agent': 'crm-zapromaq-gerar-municipios'})
    with urllib.request.urlopen(req, timeout=60) as r:
        bruto = r.read()
    # o servidor do IBGE pode mandar gzip mesmo sem pedido; o hash registrado é o do conteúdo
    if bruto[:2] == b'\x1f\x8b':
        bruto = gzip.decompress(bruto)
    return bruto


def sha256(dados):
    return hashlib.sha256(dados).hexdigest()


def ler_ibge(bruto):
    dados = json.loads(bruto.decode('utf-8'))
    saida = {}
    for m in dados:
        cod = str(m['id'])
        # a API traz a UF por região imediata; microrregião só em parte dos registros
        uf = (m.get('regiao-imediata') or {}).get('regiao-intermediaria', {}).get('UF', {}).get('sigla')
        if not uf and m.get('microrregiao'):
            uf = m['microrregiao']['mesorregiao']['UF']['sigla']
        saida[cod] = {'nome': m['nome'], 'uf': uf}
    return saida


def ler_coordenadas(bruto):
    leitor = csv.DictReader(io.StringIO(bruto.decode('utf-8-sig')))
    saida = {}
    for linha in leitor:
        saida[linha['codigo_ibge'].strip()] = {
            'nome': linha['nome'].strip(),
            'lat': linha['latitude'].strip(),
            'lon': linha['longitude'].strip(),
        }
    return saida


def montar(ibge, coord):
    so_ibge = sorted(set(ibge) - set(coord))
    so_coord = sorted(set(coord) - set(ibge))
    if so_ibge or so_coord:
        raise SystemExit('ABORTADO: códigos só em uma fonte (IBGE: %d, coordenadas: %d). Nada foi gravado.'
                         % (len(so_ibge), len(so_coord)))
    linhas, divergencias = [], []
    for cod in sorted(ibge):
        i, c = ibge[cod], coord[cod]
        if i['nome'] != c['nome']:
            divergencias.append({'cod_ibge': cod, 'ibge': i['nome'], 'coordenadas': c['nome'], 'usado': i['nome']})
        linhas.append((cod, i['nome'], nome_busca_de(i['nome']), i['uf'], c['lat'], c['lon']))
    return linhas, divergencias


def checar(linhas):
    erros = []
    if len(linhas) != TOTAL_OFICIAL:
        erros.append('total %d; esperado %d' % (len(linhas), TOTAL_OFICIAL))
    cods = [l[0] for l in linhas]
    if len(set(cods)) != len(cods):
        erros.append('cod_ibge repetido')
    if any(len(c) != 7 for c in cods):
        erros.append('cod_ibge com tamanho diferente de 7')
    ufs = {l[3] for l in linhas}
    if ufs != UFS:
        erros.append('UFs: %d; esperado 27' % len(ufs))
    for cod, nome, _, uf, lat, lon in linhas:
        try:
            la, lo = float(lat), float(lon)
        except ValueError:
            erros.append('%s: coordenada não numérica' % cod)
            continue
        if not (-34 <= la <= 6 and -74 <= lo <= -28):
            erros.append('%s %s: coordenada fora do Brasil (%s, %s)' % (cod, nome, lat, lon))
    return erros


def gravar(destino, linhas, divergencias, fontes):
    os.makedirs(destino, exist_ok=True)
    caminho_csv = os.path.join(destino, 'municipios.csv')
    buf = io.StringIO()
    buf.write('cod_ibge;nome;nome_busca;uf;lat;lon\r\n')
    for l in linhas:
        buf.write(';'.join(l) + '\r\n')
    dados_csv = ('﻿' + buf.getvalue()).encode('utf-8')
    with open(caminho_csv, 'wb') as f:
        f.write(dados_csv)
    por_uf = {}
    for l in linhas:
        por_uf[l[3]] = por_uf.get(l[3], 0) + 1
    fonte = {
        'gerado_em': datetime.date.today().isoformat(),
        'gerado_por': 'agent-config/ferramentas/gerar-municipios.py',
        'total': len(linhas),
        'total_oficial': TOTAL_OFICIAL,
        'csv_sha256': sha256(dados_csv),
        'por_uf': dict(sorted(por_uf.items())),
        'fontes': fontes,
        'divergencias_resolvidas': divergencias,
        'licencas': {
            'IBGE': 'dado público (API de localidades)',
            'kelvins/municipios-brasileiros': 'MIT (coordenadas da sede)',
        },
    }
    with open(os.path.join(destino, 'municipios.fonte.json'), 'wb') as f:
        f.write((json.dumps(fonte, ensure_ascii=False, indent=2) + '\n').encode('utf-8'))
    return caminho_csv, fonte


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--destino', default=DESTINO_PADRAO)
    args = ap.parse_args()
    agora = datetime.datetime.now().isoformat(timespec='seconds')
    bruto_ibge = baixar(URL_IBGE)
    bruto_coord = baixar(URL_COORDENADAS)
    ibge = ler_ibge(bruto_ibge)
    coord = ler_coordenadas(bruto_coord)
    linhas, divergencias = montar(ibge, coord)
    erros = checar(linhas)
    if erros:
        print('ABORTADO: ' + '; '.join(erros) + '. Nada foi gravado.')
        return 1
    fontes = [
        {'nome': 'IBGE localidades', 'url': URL_IBGE, 'sha256': sha256(bruto_ibge), 'baixado_em': agora},
        {'nome': 'kelvins/municipios-brasileiros', 'url': URL_COORDENADAS, 'sha256': sha256(bruto_coord), 'baixado_em': agora},
    ]
    caminho, fonte = gravar(args.destino, linhas, divergencias, fontes)
    print('OK: %d municípios gravados em %s' % (fonte['total'], caminho))
    print('SHA-256 do CSV: ' + fonte['csv_sha256'])
    print('Divergências de nome resolvidas (usado: IBGE): %d' % len(divergencias))
    return 0


if __name__ == '__main__':
    sys.exit(main())
