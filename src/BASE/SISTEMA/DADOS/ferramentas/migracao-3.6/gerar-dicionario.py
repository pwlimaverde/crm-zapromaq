# -*- coding: utf-8 -*-
"""gerar-dicionario.py - produz o dicionario de dados conferindo contra os cabecalhos reais."""
import sys, os, re, unicodedata, openpyxl
from datetime import datetime

def chave(nome):
    k = unicodedata.normalize("NFKD", str(nome).strip()).encode("ascii","ignore").decode().upper()
    return re.sub(r"[^A-Z0-9]+","_",k).strip("_")

# aba -> [(coluna origem, destino, tipo, regra)]   destino "—" = descarte
MAPA = {
"Clientes": [
 ("CÓD. CLIENTE","codigo_cliente","LONG","Vazio → NULL e estagio = Pré-cliente"),
 ("EMPRESA","empresa","TEXT(120)","MAIÚSCULAS, trim, espaço duplo colapsado. Obrigatório"),
 ("CIDADE","cidade","TEXT(60)","MAIÚSCULAS, trim"),
 ("ESTADO","uf","TEXT(2)","MAIÚSCULAS, 2 caracteres"),
 ("SEGMENTO","segmento","TEXT(60)","Valor da lista, sem transformação"),
 ("TELEFONE GERAL","telefone","TEXT(30)","Só dígitos; máscara na exibição"),
 ("E-MAIL GERAL","email","TEXT(120)","minúsculas, trim"),
 ("CNPJ","cnpj","TEXT(20)","Só dígitos; diferente de 14 → NULL com aviso"),
 ("RESPONSÁVEL","responsavel","TEXT(60)","Valor da lista"),
 ("ADERÊNCIA","aderencia","INTEGER","Só 1, 2 ou 3; qualquer outra coisa → NULL"),
 ("PORTE","porte","INTEGER","Só 1, 2 ou 3"),
 ("QUALIFICAÇÃO","qualificacao","TEXT(40)","Valor da lista"),
 ("OBSERVAÇÕES","observacoes","MEMO","Como digitado, só trim nas pontas"),
 ("PRIORIDADE","—","—","Fórmula na origem; recalculada na exibição a partir de qualificacao"),
 ("VALIDAÇÃO DO CNPJ","—","—","Fórmula na origem; revalidada no front"),
 ("ORDEM","—","—","Descontinuado: 4 valores preenchidos em 855 linhas"),
 ("SITUAÇÃO CADASTRO","ativo","YESNO","Descontinuado como texto; ativo = True para todos na carga"),
],
"Contatos": [
 ("CÓD. CONTATO","codigo","TEXT(15)","Congelado como está. Nomeia vínculo e pasta"),
 ("CONTROLE","controle","LONG","Preservado; novos recebem MAX+1"),
 ("CÓD. CLIENTE","id_cliente","LONG","Resolvido para clientes.id pelo mapa da carga"),
 ("CONTATO","nome","TEXT(120)","Como digitado; vazio → GERAL, registrado como exceção"),
 ("CARGO","cargo","TEXT(60)","Como digitado"),
 ("TELEFONE","telefone","TEXT(30)","Só dígitos"),
 ("E-MAIL","email","TEXT(120)","minúsculas"),
 ("OBSERVAÇÕES","observacoes","MEMO","Como digitado"),
 ("EMPRESA","—","—","Fórmula na origem; vem por JOIN"),
 ("SITUAÇÃO CADASTRO","ativo","YESNO","ativo = True na carga"),
],
"Oportunidades": [
 ("ATENDIMENTO","codigo","TEXT(15)","Congelado. Nomeia a pasta do atendimento"),
 ("CONTROLE","controle","LONG","Preservado; novos recebem MAX+1"),
 ("CÓD. CONTATO","id_contato","LONG","Resolvido para contatos.id"),
 ("CÓD. CLIENTE","id_cliente","LONG","Congelado na criação"),
 ("ATENDIMENTO ANTERIOR","id_atendimento_anterior","LONG","Resolvido pelo código; hoje vazio em todas as 487"),
 ("ORÇAMENTO","orcamento","TEXT(30)","MAIÚSCULAS, trim"),
 ("ETAPA","etapa","TEXT(40)","Valor da lista. Obrigatório"),
 ("RESPONSÁVEL","responsavel","TEXT(60)","Valor da lista"),
 ("MÁQUINA / DESCRIÇÃO","maquina","TEXT(255)","MAIÚSCULAS, trim"),
 ("FAMÍLIA DE MÁQUINA","familia","TEXT(60)","Valor da lista"),
 ("CATEGORIA","categoria","TEXT(60)","Valor da lista"),
 ("TIPO DE VENDA","tipo_venda","TEXT(40)","Como digitado"),
 ("VALOR (R$)","valor","CURRENCY","Vazio permanece NULL — nunca vira zero"),
 ("ORIGEM","origem","TEXT(40)","Valor da lista"),
 ("PRIORIDADE","prioridade","TEXT(20)","Valor da lista"),
 ("DATA DE ENTRADA","dt_entrada","DATETIME","Parâmetro tipado, nunca literal"),
 ("DATA DA PROPOSTA","dt_proposta","DATETIME","Idem"),
 ("ÚLTIMA INTERAÇÃO","ultima_interacao","DATETIME","Idem"),
 ("TENTATIVAS","tentativas","INTEGER","Vazio → NULL"),
 ("PRÓXIMA AÇÃO","prox_acao","TEXT(255)","Como digitado"),
 ("DATA PRÓX. AÇÃO","dt_prox_acao","DATETIME","Idem"),
 ("RETOMAR EM","retomar_em","DATETIME","Idem"),
 ("MOTIVO DE DESFECHO","motivo_desfecho","TEXT(60)","Valor da lista"),
 ("DATA DO DESFECHO","dt_desfecho","DATETIME","Idem"),
 ("OBSERVAÇÕES","observacoes","MEMO","Como digitado"),
 ("SITUAÇÃO","—","—","Calculado na exibição (modCalc)"),
 ("QUADRO","—","—","Calculado na exibição"),
 ("DIAS PARADO","—","—","Calculado na exibição"),
 ("CICLO (DIAS)","—","—","Calculado na exibição"),
 ("MÊS PROPOSTA","—","—","Calculado na exibição"),
 ("MÊS ENTRADA","—","—","Calculado na exibição"),
 ("MÊS DESFECHO","—","—","Calculado na exibição"),
 ("EMPRESA","—","—","JOIN com clientes"),
 ("CIDADE","—","—","JOIN com clientes"),
 ("ESTADO","—","—","JOIN com clientes"),
 ("SEGMENTO","—","—","JOIN com clientes"),
 ("CONTATO","—","—","JOIN com contatos"),
 ("CARGO","—","—","JOIN com contatos"),
 ("TELEFONE","—","—","JOIN com contatos"),
 ("E-MAIL","—","—","JOIN com contatos"),
],
}

def main(xlsm, saida):
    wb = openpyxl.load_workbook(xlsm, read_only=True, data_only=True)
    linhas = ["# Dicionário de dados — migração do CRM 3.6",
              "",
              "**Gerado por `gerar-dicionario.py` em %s**, conferido contra os cabeçalhos reais de `%s`."
              % (datetime.now().strftime("%d/%m/%Y %H:%M"), os.path.basename(xlsm)),
              "",
              "O mapeamento é **por nome de coluna**, nunca por posição. Destino `—` significa descarte declarado.",
              ""]
    problemas = []
    for aba, itens in MAPA.items():
        ws = wb[aba]
        reais = [c for c in next(ws.iter_rows(min_row=3, max_row=3, values_only=True)) if c and str(c).strip()]
        reais_k = {chave(c): str(c).strip() for c in reais}
        declaradas = {chave(o): o for o, *_ in itens}
        faltam = [v for k, v in reais_k.items() if k not in declaradas]
        sobram = [v for k, v in declaradas.items() if k not in reais_k]
        linhas += ["## %s" % aba, "",
                   "%d colunas na origem, %d declaradas." % (len(reais), len(itens)), "",
                   "| Coluna de origem | Destino | Tipo | Regra |", "|---|---|---|---|"]
        for o, d, t, r in itens:
            linhas.append("| `%s` | `%s` | %s | %s |" % (o, d, t, r))
        linhas.append("")
        if faltam:
            problemas.append("%s: coluna sem destino declarado: %s" % (aba, faltam))
        if sobram:
            problemas.append("%s: declarada mas inexistente na origem: %s" % (aba, sobram))

    linhas += ["## Conferência", ""]
    if problemas:
        linhas.append("**PENDENTE — corrigir antes da carga:**")
        linhas += ["- " + p for p in problemas]
    else:
        linhas.append("Todas as colunas das três abas têm destino ou descarte declarado. Nenhuma coluna declarada está ausente da origem.")
    linhas.append("")

    os.makedirs(os.path.dirname(saida), exist_ok=True)
    open(saida, "w", encoding="utf-8").write("\n".join(linhas))
    for p in problemas:
        print("PROBLEMA:", p)
    print("OK" if not problemas else "CONFERIR", "->", os.path.basename(saida))

if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
