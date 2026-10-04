# -*- coding: utf-8 -*-
"""
conferir-calculados.py - valida a reimplementacao dos campos calculados
contra os valores que a propria planilha produziu.

E o passo 7 do roteiro: modCalc so pode ser escrito depois que a logica
reproduz a planilha registro a registro.
"""
import sys, collections
from datetime import date, datetime, timedelta
import openpyxl

ENCERRADAS = ("Pedido Fechado", "Perdido", "Descartado")
PROSPECCAO = ("Contato Inicial", "Sem Retorno", "Retorno Agendado", "Descartado")

def d(v):
    if isinstance(v, datetime): return v.date()
    if isinstance(v, date): return v
    return None

def situacao(cod_contato, etapa, dt_prox_acao, retomar_em, hoje):
    if not cod_contato: return ""
    if not etapa: return "Sem etapa definida"
    if etapa in ENCERRADAS: return "Encerrado"
    if retomar_em and retomar_em <= hoje: return "Retomar hoje"
    if dt_prox_acao and dt_prox_acao < hoje: return "Ação atrasada"
    if not dt_prox_acao: return "Sem próxima ação"
    # fim da semana corrente: hoje - weekday(segunda=1..domingo=7) + 7
    fim_semana = hoje - timedelta(days=(hoje.weekday() + 1)) + timedelta(days=7)
    if dt_prox_acao <= fim_semana: return "Ação nesta semana"
    return "Em dia"

def quadro(etapa):
    if not etapa: return ""
    return "1. Prospecção" if etapa in PROSPECCAO else "2. Funil comercial"

def dias_parado(ultima_interacao, hoje):
    if not ultima_interacao: return ""
    return (hoje - ultima_interacao).days

def ciclo(dt_proposta, dt_desfecho):
    if not dt_proposta or not dt_desfecho: return ""
    return (dt_desfecho - dt_proposta).days

def mes_ref(dt):
    if not dt: return ""
    return dt.year * 100 + dt.month

def main(xlsm, hoje_txt=None):
    hoje = datetime.strptime(hoje_txt, "%Y-%m-%d").date() if hoje_txt else date.today()
    wb = openpyxl.load_workbook(xlsm, read_only=True, data_only=True)
    ws = wb["Oportunidades"]
    # colunas base 0
    C_CONTATO, C_ETAPA = 4, 5
    C_ENTRADA, C_PROPOSTA, C_ULTIMA = 23, 24, 25
    C_PROXACAO, C_RETOMAR, C_DESFECHO = 28, 29, 31
    C_SITUACAO, C_QUADRO = 2, 17
    C_DIAS, C_CICLO, C_MPROP, C_MENTR, C_MDESF = 35, 36, 37, 38, 39

    campos = ["situacao", "quadro", "dias_parado", "ciclo", "mes_proposta", "mes_entrada", "mes_desfecho"]
    ok = collections.Counter()
    dif = collections.Counter()
    exemplos = collections.defaultdict(list)
    total = 0

    for lin in ws.iter_rows(min_row=4, values_only=True):
        if not lin[0]:
            continue
        total += 1
        calc = {
            "situacao":     situacao(lin[C_CONTATO], (lin[C_ETAPA] or "").strip(),
                                     d(lin[C_PROXACAO]), d(lin[C_RETOMAR]), hoje),
            "quadro":       quadro((lin[C_ETAPA] or "").strip()),
            "dias_parado":  dias_parado(d(lin[C_ULTIMA]), hoje),
            "ciclo":        ciclo(d(lin[C_PROPOSTA]), d(lin[C_DESFECHO])),
            "mes_proposta": mes_ref(d(lin[C_PROPOSTA])),
            "mes_entrada":  mes_ref(d(lin[C_ENTRADA])),
            "mes_desfecho": mes_ref(d(lin[C_DESFECHO])),
        }
        planilha = {
            "situacao": lin[C_SITUACAO], "quadro": lin[C_QUADRO],
            "dias_parado": lin[C_DIAS], "ciclo": lin[C_CICLO],
            "mes_proposta": lin[C_MPROP], "mes_entrada": lin[C_MENTR], "mes_desfecho": lin[C_MDESF],
        }
        for c in campos:
            p = planilha[c]
            p = "" if p is None else p
            if isinstance(p, float) and p == int(p):
                p = int(p)
            if str(calc[c]).strip() == str(p).strip():
                ok[c] += 1
            else:
                dif[c] += 1
                if len(exemplos[c]) < 4:
                    exemplos[c].append("%s: calculado=%r planilha=%r (etapa=%r prox=%s retomar=%s)"
                                       % (lin[0], calc[c], p, lin[C_ETAPA],
                                          d(lin[C_PROXACAO]), d(lin[C_RETOMAR])))

    print("Registros conferidos: %d   (TODAY = %s)" % (total, hoje))
    print()
    print("%-16s %8s %8s  %s" % ("campo", "iguais", "difer.", "situacao"))
    for c in campos:
        marca = "OK" if dif[c] == 0 else "CONFERIR"
        print("%-16s %8d %8d  %s" % (c, ok[c], dif[c], marca))
    for c in campos:
        if dif[c]:
            print()
            print("  divergencias em %s:" % c)
            for e in exemplos[c]:
                print("    -", e)
    return 0 if sum(dif.values()) == 0 else 1

if __name__ == "__main__":
    sys.exit(main(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None))
