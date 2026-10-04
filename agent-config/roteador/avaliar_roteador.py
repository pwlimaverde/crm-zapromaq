"""Calibração do roteador em modo seco (não injeta nada, não grava no log).

Lê `avaliar_casos.json` (mensagem, nível e skill esperados; `continuacao`
quando o caso depende da tarefa anterior), consulta o Jev como o gancho faz e
imprime a matriz de nível, o acerto de skill, a continuação e as latências.
Também mede a partida do processo do gancho.

Usa a rede e a chave TYPESAFE_API_KEY (só a mensagem mascarada sai da máquina).
Uso: py -3 agent-config/roteador/avaliar_roteador.py
"""

from __future__ import annotations

import json
import statistics
import subprocess
import sys
import time
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import roteador as r  # noqa: E402

ARQ_CASOS = Path(__file__).resolve().parent / "avaliar_casos.json"

# Tarefa anterior sintética para os casos de continuação.
ANTERIOR = {
    "mensagem": "a ficha abre o registro errado quando a grade está filtrada, descobre por quê",
    "nivel_final": "dificil",
    "especialista": "general-purpose",
    "skill": "debugging-and-error-recovery",
}


def p95(valores: list[int]) -> int:
    if not valores:
        return 0
    ordenados = sorted(valores)
    return ordenados[max(0, int(round(0.95 * len(ordenados))) - 1)]


def medir_partida(vezes: int = 5) -> list[int]:
    """Tempo do processo do gancho num atalho que sai antes da rede."""
    gancho = str(Path(r.__file__).resolve())
    tempos = []
    for _ in range(vezes):
        inicio = time.perf_counter()
        subprocess.run(["py", "-3", gancho, "--cli", "status"], capture_output=True, check=False)
        tempos.append(int((time.perf_counter() - inicio) * 1000))
    return tempos


def main() -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    config = r.carregar_config()
    cat = r.catalogo(config)
    casos = json.loads(ARQ_CASOS.read_text(encoding="utf-8"))
    elegiveis = {s: d for s, d in cat["skills"].items() if s not in config.get("skills_fora_do_jev", [])}
    matriz: Counter = Counter()
    skill_ok = skill_fp = skill_fn = skill_troca = 0
    cont_ok = cont_total = 0
    latencias: list[int] = []
    linhas = []
    for caso in casos:
        anterior = ANTERIOR if "continuacao" in caso else None
        state = {"mensagem": r.mascarar(caso["mensagem"]), "projeto": config["projeto"]}
        if anterior:
            state["tarefa_anterior"] = {
                "mensagem": anterior["mensagem"],
                "nivel": anterior["nivel_final"],
                "especialista": anterior["especialista"],
                "skill": anterior["skill"],
            }
        perguntas = r.montar_perguntas(cat, elegiveis, anterior)
        try:
            respostas, ms = r.consultar_jev(config, state, perguntas)
        except r.FalhaJev as exc:
            print(f"FALHA ({exc.motivo}): {caso['mensagem']}")
            if exc.motivo == "chave_ausente":
                print("Defina TYPESAFE_API_KEY (bloco env do ~/.claude/settings.json).")
                return 1
            continue
        latencias.append(ms)
        d = r.decidir(config, respostas, texto=caso["mensagem"], anterior=anterior)

        esperado_nivel = caso["nivel"] or ANTERIOR["nivel_final"]
        esperado_skill = caso["skill"]
        if caso.get("continuacao") and esperado_skill is None:
            esperado_skill = ANTERIOR["skill"]
        matriz[(esperado_nivel, d["nivel_final"])] += 1

        obtida = d.get("skill")
        if obtida == esperado_skill:
            skill_ok += 1
        elif esperado_skill is None:
            skill_fp += 1
        elif obtida is None:
            skill_fn += 1
        else:
            skill_troca += 1
        if "continuacao" in caso:
            cont_total += 1
            if d["retomar"] == caso["continuacao"]:
                cont_ok += 1
        marca = "ok" if (obtida == esperado_skill and esperado_nivel == d["nivel_final"]) else "XX"
        linhas.append(
            f"{marca} nivel {esperado_nivel:>7}→{d['nivel_final']:<7}"
            f" (jev {d.get('nivel_jev')} {d.get('nivel_confianca') or 0:.2f}, risco "
            f"{d.get('risco') or 0:.2f}{', ↑' + d['escalada'] if d.get('escalada') else ''})"
            f" | skill {esperado_skill or '—'}→{obtida or '—'}"
            f" (cand {d.get('skill_candidata')} {d.get('skill_chance') or 0:.2f}/"
            f"{d.get('skill_confianca') or 0:.2f}) | {d['especialista']} | {caso['mensagem'][:55]}"
        )

    print("\n".join(linhas))
    total = sum(matriz.values())
    print("\nMatriz de nível (linha = esperado, coluna = decidido):")
    print(f"{'':>9}" + "".join(f"{n:>9}" for n in r._ORDEM))
    for e in r._ORDEM:
        print(f"{e:>9}" + "".join(f"{matriz[(e, o)]:>9}" for o in r._ORDEM))
    acerto_nivel = sum(matriz[(n, n)] for n in r._ORDEM)
    abaixo = sum(v for (e, o), v in matriz.items() if r._ORDEM.index(o) < r._ORDEM.index(e))
    print(f"acerto de nível: {acerto_nivel}/{total} · decidido ABAIXO do esperado: {abaixo}")
    print(
        f"skill: acerto {skill_ok}/{total} · falso positivo {skill_fp} · "
        f"falso negativo {skill_fn} · trocada {skill_troca}"
    )
    print(f"continuação: {cont_ok}/{cont_total}")
    partida = medir_partida()
    jev_media = int(statistics.mean(latencias)) if latencias else 0
    print(
        f"latência do Jev: média {jev_media} ms · p95 {p95(latencias)} ms\n"
        f"partida do processo (py -3, sem rede): média {int(statistics.mean(partida))} ms · "
        f"p95 {p95(partida)} ms"
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
