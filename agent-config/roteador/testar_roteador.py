"""Testes do roteador SEM rede (o Jev não é chamado): mascaramento, atalhos,
política de decisão, modos off/direto/comando, falha sem chave e o contrato do
gancho (stdin JSON -> stdout JSON). Estado gravado numa pasta temporária.

Uso: py -3 agent-config/roteador/testar_roteador.py   (sai 0 se tudo passou)
O verificar-tudo roda este teste no nível -Rapido.
"""

from __future__ import annotations

import io
import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import roteador as r  # noqa: E402

falhas = 0


def conferir(nome: str, ok: bool, detalhe: str = "") -> None:
    global falhas
    if ok:
        print(f"  OK    {nome}")
    else:
        falhas += 1
        print(f"  FALHA {nome}  {detalhe}")


def main() -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    temp = Path(tempfile.mkdtemp(prefix="roteador-teste-"))
    # Estado isolado: nada do teste vai para roteador/estado.
    r.DIR_ESTADO = temp
    r.ARQ_MODO = temp / "modo"
    r.ARQ_LOG = temp / "decisoes.jsonl"
    r.ARQ_STATUS = temp / "status.txt"
    r.ARQ_CATALOGO = temp / "catalogo.json"
    config = r.carregar_config()

    print("--- mascaramento (nada de dado de cliente nem segredo sai da máquina)")
    m = r.mascarar("cliente CNPJ 12.345.678/0001-90 e 12345678000190, fone (41) 99999-0000 e 41999990000")
    conferir("CNPJ formatado", "12.345.678" not in m, m)
    conferir("CNPJ só dígitos", "12345678000190" not in m, m)
    conferir("telefone formatado", "99999-0000" not in m, m)
    conferir("telefone só dígitos", "41999990000" not in m, m)
    m = r.mascarar("CPF 123.456.789-09, senha=abc123, contato@empresa.com.br")
    conferir("CPF", "123.456.789-09" not in m, m)
    conferir("senha", "abc123" not in m, m)
    conferir("e-mail", "empresa.com.br" not in m, m)
    conferir("texto comum intacto", r.mascarar("trocar a senha do usuário na ficha") == "trocar a senha do usuário na ficha")
    conferir("código AT intacto", "AT-0001-0001" in r.mascarar("abre o AT-0001-0001"))

    print("--- catálogo")
    cat = r.catalogo(config)
    conferir("skills do agent-skills lidas", "test-driven-development" in cat["skills"] and len(cat["skills"]) >= 20, str(len(cat["skills"])))
    conferir("agentes do projeto lidos", {"code-reviewer", "security-auditor", "test-engineer"} <= set(cat["especialistas"]))
    conferir("agentes nativos incluídos", {"general-purpose", "Plan"} <= set(cat["especialistas"]))
    conferir("skills_fora_do_jev existem", all(s in cat["skills"] for s in config["skills_fora_do_jev"]))
    conferir("comando_para_skill aponta para skills existentes", all(s in cat["skills"] for s in config["comando_para_skill"].values()))
    conferir("skill_para_especialista coerente",
             all(s in cat["skills"] and e in cat["especialistas"] for s, e in config["skill_para_especialista"].items()))
    conferir("pares concorrentes coerentes",
             all(set(p["par"]) <= set(cat["skills"]) and p["especializada"] in p["par"] for p in config["pares_concorrentes"]))

    print("--- atalhos")
    t, n, s = r.ler_atalhos("#profundo $code-simplification limpa o modGrade", config, cat["skills"])
    conferir("#profundo + $skill", (n, s, t) == ("dificil", "code-simplification", "limpa o modGrade"), f"{n} {s} {t}")
    t, n, s = r.ler_atalhos("$naoexiste faz algo", config, cat["skills"])
    conferir("skill inexistente não é atalho", (n, s) == (None, None) and t.startswith("$naoexiste"))

    print("--- política (respostas sintéticas do Jev)")
    resp = {"nivel": {"choice": "rotina", "confidence": 0.9}, "risco": {"noul": 0.9},
            "skill": {"choice": "test-driven-development", "confidence": 0.8,
                      "probabilities": {"test-driven-development": 0.7, "nenhuma": 0.3}},
            "especialista": {"choice": "test-engineer", "confidence": 0.8}}
    d = r.decidir(config, resp, texto="cria teste", anterior=None)
    conferir("risco alto sobe um nível", d["nivel_final"] == "dificil" and d["escalada"] == "risco", str(d))
    conferir("skill aceita pelo Jev", d["skill"] == "test-driven-development" and d["skill_motivo"] == "jev")
    conferir("modelo do nível", d["modelo"] == "opus")
    resp = {"nivel": {"choice": "simples", "confidence": 0.4}, "skill": {"choice": "nenhuma", "confidence": 0.9, "probabilities": {"nenhuma": 0.9}}}
    d = r.decidir(config, resp, texto="onde fica o modDB?", anterior=None)
    conferir("confiança baixa sobe um nível", d["nivel_final"] == "rotina" and d["escalada"] == "confianca", str(d))
    conferir("nenhuma skill", d["skill"] is None)
    resp = {"skill": {"choice": "test-driven-development", "confidence": 0.3,
                      "probabilities": {"test-driven-development": 0.4, "debugging-and-error-recovery": 0.35, "nenhuma": 0.25}}}
    d = r.decidir(config, resp, texto="a grade quebrou ao filtrar, investiga", anterior=None)
    conferir("par concorrente com gatilho escolhe a especializada", d["skill"] == "debugging-and-error-recovery" and d["skill_motivo"] == "par", str(d))
    anterior = {"nivel_final": "dificil", "especialista": "code-reviewer", "skill": "code-review-and-quality"}
    resp = {"nivel": {"choice": "simples", "confidence": 0.9}, "continuacao": {"noul": 0.9},
            "skill": {"choice": "nenhuma", "confidence": 0.9, "probabilities": {"nenhuma": 0.9}}}
    d = r.decidir(config, resp, texto="sim, continua", anterior=anterior)
    conferir("continuação herda nível, skill e especialista (nunca desce)",
             d["retomar"] and d["nivel_final"] == "dificil" and d["skill"] == "code-review-and-quality" and d["especialista"] == "code-reviewer", str(d))

    print("--- fluxo")
    r.ARQ_MODO.write_text("off\n", encoding="utf-8")
    reg, ctx = r.processar({"prompt": "qualquer coisa", "session_id": "t"}, config, seco=True)
    conferir("modo off não faz nada", reg.get("atalho") == "off" and ctx == "")
    r.ARQ_MODO.write_text("skills\n", encoding="utf-8")
    reg, ctx = r.processar({"prompt": ">> direto", "session_id": "t"}, config, seco=True)
    conferir(">> passa direto", reg.get("atalho") == "direto" and ctx == "")
    reg, ctx = r.processar({"prompt": "/spec campo novo", "session_id": "t"}, config, seco=True)
    conferir("/spec é comando, mapeado para a skill", reg.get("atalho") == "comando" and reg.get("skill") == "spec-driven-development" and ctx == "", str(reg))
    chave = os.environ.pop("TYPESAFE_API_KEY", None)
    try:
        reg, ctx = r.processar({"prompt": "$test-driven-development cria teste do Nz", "session_id": "t"}, config, seco=True)
        conferir("sem chave: falha controlada, skill forçada mantida", reg.get("falha") is None and reg.get("skill") == "test-driven-development", str(reg))
        reg, ctx = r.processar({"prompt": "cria teste do Nz", "session_id": "t"}, config, seco=True)
        conferir("sem chave: decisão de segurança", reg.get("falha") == "chave_ausente" and reg.get("skill") is None, str(reg))
        conferir("status mostra sem chave", r.texto_status(config, reg, "skills") == "roteador: sem chave")
    finally:
        if chave is not None:
            os.environ["TYPESAFE_API_KEY"] = chave
    r.ARQ_MODO.write_text("skills\n", encoding="utf-8")
    reg = {"skill": "code-simplification"}
    ctx = r.contexto_para_claude("skills", reg, False, config)
    conferir("contexto do seletor manda invocar a skill", "ferramenta Skill" in ctx and "code-simplification" in ctx)
    ctx = r.contexto_para_claude("on", {"nivel_final": "rotina", "modelo": "sonnet", "especialista": "test-engineer", "skill": None}, False, config)
    conferir("contexto de delegação cita as regras do projeto", "verificar-tudo" in ctx and "test-engineer" in ctx)

    print("--- contrato do gancho (processo de verdade, sem rede: modo off)")
    env = dict(os.environ, ROTEADOR_ESTADO=str(temp))
    script = str(Path(r.__file__).resolve())
    # `>>` sai antes da rede e não injeta nada; o estado vai para a pasta temporária.
    proc = subprocess.run([sys.executable, script], input=json.dumps({"prompt": ">> oi", "session_id": "teste-contrato"}).encode("utf-8"),
                          capture_output=True, env=env, timeout=20)
    conferir("gancho sai 0 e não escreve nada no stdout para >>", proc.returncode == 0 and proc.stdout == b"", repr(proc.stdout[:200]))
    ultima = r.ler_ultima_decisao("teste-contrato")
    conferir("gancho registrou no estado temporário", ultima is not None and ultima.get("atalho") == "direto", str(ultima))
    conferir("status gravado", r.ARQ_STATUS.read_text(encoding="utf-8").strip().endswith("roteador: direto (>>)"))
    proc = subprocess.run([sys.executable, script], input=b"isto nao e json", capture_output=True, env=env, timeout=20)
    conferir("entrada inválida não derruba a sessão", proc.returncode == 0 and proc.stdout == b"")
    conferir("erro interno aparece no status", "erro interno" in r.ARQ_STATUS.read_text(encoding="utf-8"))

    print()
    print("TUDO CERTO." if falhas == 0 else f"{falhas} falha(s).")
    return 0 if falhas == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
