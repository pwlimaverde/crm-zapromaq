# -*- coding: utf-8 -*-
"""
conferir-sql-gerado.py - reproduz em Python o SQL que modCRM.Inserir e
modCRM.Atualizar montam a partir do modSchema, e confere:
  - nenhuma coluna repetida
  - numero de '?' igual ao numero de parametros
  - colunas existem no DDL do criar-banco.ps1
"""
import sys, os, re

TIPO_ADO = {"M":"adLongVarWChar","N":"adInteger","C":"adCurrency",
            "D":"adDate","K":"adInteger","T":"adVarWChar","L":"adVarWChar","R":"adVarWChar"}

def campos_do_schema(caminho):
    txt = open(caminho, encoding="cp1252").read()
    # so o corpo de CamposDe: ColunasDe e TituloDe tambem tem Case "clientes"
    ini = txt.index("Public Function CamposDe")
    fim = txt.index("Public Function ColunasDe")
    txt = txt[ini:fim]
    tabelas, atual = {}, None
    for linha in txt.split("\n"):
        m = re.match(r'\s*Case "(\w+)"', linha)
        if m and m.group(1) in ("clientes","contatos","oportunidades"):
            atual = m.group(1); tabelas[atual] = []
            continue
        m = re.match(r'\s*Add c, "([^"]+)"', linha)
        if m and atual:
            p = m.group(1).split("|")
            if len(p) == 7:
                tabelas[atual].append(p)
    return tabelas

def colunas_do_ddl(caminho):
    txt = open(caminho, encoding="utf-8-sig").read()
    tabelas = {}
    for m in re.finditer(r'CREATE TABLE (\w+) \("(.*?)"\)', txt, re.S):
        pass
    for nome in ("clientes","contatos","oportunidades"):
        bloco = re.search(r'CREATE TABLE ' + nome + r' \(" \+(.*?)\)"\)', txt, re.S)
        if not bloco:
            bloco = re.search(r'CREATE TABLE ' + nome + r' \((.*?)\)"\)', txt, re.S)
        cols = set()
        if bloco:
            for c in re.findall(r'"\s*(\w+)\s+[A-Z]', bloco.group(1)):
                cols.add(c.lower())
        tabelas[nome] = cols
    return tabelas

EXTRAS = {
    "clientes":      ["codigo_cliente", "estagio"],
    "contatos":      ["codigo", "controle", "id_cliente"],
    "oportunidades": ["codigo", "controle", "id_contato", "id_cliente"],
}

def main(pasta_fontes, ddl):
    schema = campos_do_schema(os.path.join(pasta_fontes, "modSchema.bas"))
    ddlcols = colunas_do_ddl(ddl)
    problemas = 0

    for tab in ("clientes","contatos","oportunidades"):
        graváveis = [c for c in schema[tab] if c[6] == "1"]
        extras = EXTRAS[tab]
        colunas, marcas, params = [], [], []

        for c in graváveis:
            if c[0] in extras:            # extras tem precedencia
                continue
            colunas.append(c[0]); marcas.append("?")
            params.append(TIPO_ADO.get(c[2], "adVarWChar"))
        for e in extras:
            colunas.append(e); marcas.append("?")
            params.append("adVarWChar" if e in ("codigo","estagio") else "adInteger")
        colunas += ["versao","criado_em","alterado_em","alterado_por"]
        marcas  += ["1","?","?","?"]
        params  += ["adDate","adDate","adVarWChar"]

        sql = "INSERT INTO %s (%s) VALUES (%s)" % (tab, ", ".join(colunas), ",".join(marcas))
        qt = sql.count("?")
        dup = [c for c in set(colunas) if colunas.count(c) > 1]
        fora = [c for c in colunas if ddlcols[tab] and c.lower() not in ddlcols[tab]]

        print("=== %s ===" % tab)
        print("  colunas: %d | '?': %d | parametros: %d" % (len(colunas), qt, len(params)))
        print("  %s" % ("OK" if qt == len(params) else "<<< DESCOMPASSO ? x parametros"))
        if qt != len(params): problemas += 1
        if dup:
            print("  <<< COLUNA REPETIDA: %s" % dup); problemas += 1
        else:
            print("  sem coluna repetida")
        if fora:
            print("  <<< COLUNA QUE NAO EXISTE NO BANCO: %s" % fora); problemas += 1
        elif ddlcols[tab]:
            print("  todas as colunas existem no banco")
        print("  %s..." % sql[:150])
        print()

    print("RESULTADO:", "sem problemas" if problemas == 0 else "%d ponto(s)" % problemas)
    return 0 if problemas == 0 else 1

if __name__ == "__main__":
    sys.exit(main(sys.argv[1], sys.argv[2]))
