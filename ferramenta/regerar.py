#!/usr/bin/env python3
"""Regera as obras com a MESMA sequência para todas (registro em ferramenta/obras.json).

Uso:  python3 ferramenta/regerar.py [obra ...] [--sem-site]      (Windows: python ferramenta\\regerar.py ...)
      (sem obra = todas; obra = botanico | ed-jardim | torre-catharina)

Para cada obra:
  1. 1ª passada: gera, recalcula (ferramenta/recalc.py, LibreOffice) e audita a planilha de cada cenário.
  2. Gantt (gerador/gantt.py) e curva física (gerador/fisico.py) a partir das planilhas da 1ª passada.
  3. 2ª passada com as abas CURVA FÍSICA, GANTT e ESQUEMÁTICO; cenários com "fixar_2027" recalibram
     cfg.desloc.fator até 2027 bater no valor fixado (diferença vai para o ano seguinte).
  4. JSON do esquemático, dos mapas (só se a obra tiver os insumos) e do site (gerador/viewer.py).
  5. Planilhas: raiz do repositório (planilha_raiz.xlsx) e site/xlsx/<id>.xlsx + .b64.txt.
Depois de todas: make_site → PDFs (pdfs.js) → make_site → conferir_publicacao.py preparar.
A publicação em si segue a rotina do CLAUDE.md (lotes do .pendente.json → reler → conferir).
"""
import base64, json, os, re, shutil, subprocess, sys

if not sys.flags.utf8_mode:   # Windows: força UTF-8 (acentos nos JSON e nomes) também nos processos filhos
    os.environ["PYTHONUTF8"] = "1"; sys.exit(subprocess.run([sys.executable, "-X", "utf8", *sys.argv]).returncode)
F = os.path.dirname(os.path.abspath(__file__)); R = os.path.dirname(F); G = os.path.join(R, "gerador")
# saídas intermediárias: FLUXO_SAIDA permite pô-las num disco local quando o projeto está numa pasta de rede
OUT = os.environ.get("FLUXO_SAIDA") or os.path.join(F, "saida"); REG = json.load(open(os.path.join(F, "obras.json")))


def sh(args, env=None, cwd=R):
    p = subprocess.run(args, env={**os.environ, **(env or {})}, cwd=cwd, capture_output=True, text=True)
    if p.returncode:
        sys.exit(f"ERRO em {' '.join(args)}\n{p.stdout[-3000:]}\n{p.stderr[-3000:]}")
    return p.stdout


def py(*a, env=None):
    return sh([sys.executable, "-W", "ignore", *a], env)


def build(o, c, env):
    """Gera, recalcula (LibreOffice) e audita a planilha de um cenário; devolve o desembolso de 2027."""
    saida = os.path.join(OUT, c["id"]); calc = saida + "_calc.xlsx"
    e = dict(env, DATA_DIR=os.path.join(F, "obras", o["obra"], c["pasta"]), GERADOR=G, **({"CENARIO": c["CENARIO"]} if c.get("CENARIO") else {}))
    py(os.path.join(F, "build_modelo.py"), saida + ".xlsx", env=e)
    shutil.copy(saida + ".xlsx", calc)
    r = json.loads(py(os.path.join(F, "recalc.py"), calc))
    if r.get("total_errors"): sys.exit(f"ERRO: {c['id']} tem {r['total_errors']} células com erro: {r['exemplos']}")
    py(os.path.join(G, "cantos.py"), calc)
    txt = py(os.path.join(F, "audit.py"), calc)
    if "RESUMO: 19 OK de 19" not in txt:
        sys.exit(f"ERRO: auditoria de {c['id']}:\n" + "\n".join(l for l in txt.splitlines() if "FALHA" in l or "RESUMO" in l))
    m = re.search(r"Desembolso 2027.*\|\s*([\d,.]+)\s*/", txt)
    return float(m.group(1).replace(",", "")) if m else None


def calibrar(o, c, env, v):
    """Ajusta cfg.desloc.fator (secante) até 2027 = fixar_2027 (tolerância de meio centavo)."""
    alvo = c["fixar_2027"]; pc = os.path.join(F, "obras", o["obra"], c["pasta"], "cfg.json")
    cfg = json.load(open(pc)); chave = next(iter(cfg["desloc"]["fator"])); k = cfg["desloc"]["fator"][chave]
    pts = [(k, v)]
    for _ in range(6):
        if abs(v - alvo) < 0.005: break
        if len(pts) == 1: k = k * alvo / v
        else: (k1, v1), (k2, v2) = pts[-2:]; k = k2 + (alvo - v2) * (k2 - k1) / (v2 - v1)
        cfg["desloc"]["fator"][chave] = k; json.dump(cfg, open(pc, "w"), ensure_ascii=False, indent=1)
        v = build(o, c, env); pts.append((k, v)); print(f"   {c['id']}: fator {k:.10f} → 2027 = {v:,.2f}")
    if abs(v - alvo) >= 0.005: sys.exit(f"ERRO: {c['id']} não convergiu para 2027 = {alvo:,.2f} (ficou {v:,.2f})")
    return v


def obra(o):
    print(f"== {o['nome_site']}")
    calc = lambda i: os.path.join(OUT, f"{i}_calc.xlsx")
    for c in o["cenarios"]:
        print(f"   1ª passada {c['id']}: 2027 = {build(o, c, {}):,.2f}")
    gantt = {}
    for gd in sorted({c["gantt_de"] for c in o["cenarios"]}):
        gantt[gd] = os.path.join(OUT, f"{gd}_gantt.json"); py(os.path.join(G, "gantt.py"), calc(gd), gantt[gd])
    fis = os.path.join(OUT, f"{o['obra']}_fis.json")
    py(os.path.join(G, "fisico.py"), fis, *[f"{c['cenario']}={calc(c['id'])}" for c in o["cenarios"]], f"--base={o['fis_base']}")
    ex = o.get("extras", {}); E = o.get("esquematico", {})
    for c in o["cenarios"]:
        env = {"ESQ": "1", "FIS_JSON": fis, "GANTT_JSON": gantt[c["gantt_de"]]}
        v = build(o, c, env)
        if c.get("fixar_2027"): v = calibrar(o, c, env, v)
        print(f"   2ª passada {c['id']}: 2027 = {v:,.2f}")
        esq = os.path.join(OUT, f"{c['id']}_esq.json"); a = []
        for pac, ate, rot in E.get("acab", []): a += ["--acab", pac, ate, rot]
        if E.get("cortes"): a += ["--cortes", os.path.join(R, E["cortes"])]
        py(os.path.join(G, "esquematico.py"), "--json", calc(c["id"]), esq, *a)
        vw = {"VW_VALUES": calc(c["id"]), "VW_FORMULAS": os.path.join(OUT, f"{c['id']}.xlsx"), "VW_NOME": o["nome_site"],
              "VW_ID": c["id"], "VW_OBRA": o["obra"], "VW_CEN": c["cenario"], "VW_CEN_DESC": c["descricao"], "VW_CEN_ORD": str(c["ordem"]),
              "VW_FIS": fis, "VW_GANTT": gantt[c["gantt_de"]], "VW_ESQ": esq}
        if ex.get("impermeabilizacao"): vw["VW_IMP"] = os.path.join(R, ex["impermeabilizacao"])
        if ex.get("fachada_tasks"):
            fac = os.path.join(OUT, f"{c['id']}_fac.json")
            py(os.path.join(G, "fachada_mapa.py"), os.path.join(F, "obras", o["obra"], c["pasta"], "data2.json"), os.path.join(R, ex["fachada_tasks"]), fac)
            vw["VW_FAC"] = fac
        print("  ", py(os.path.join(G, "viewer.py"), "--json", os.path.join(R, "site", "obras", f"{c['id']}.json"), env=vw).strip().splitlines()[-1])
        shutil.copy(calc(c["id"]), os.path.join(R, c["planilha_raiz"] + ".xlsx"))
        x = os.path.join(R, "site", "xlsx", f"{c['id']}.xlsx"); shutil.copy(calc(c["id"]), x)
        open(x + ".b64.txt", "w").write(base64.b64encode(open(x, "rb").read()).decode())


def site():
    ms = lambda: py(os.path.join(G, "make_site.py"))
    ms()
    srv = subprocess.Popen([sys.executable, "-m", "http.server", "8766"], cwd=os.path.join(R, "site"),
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    try:
        import time; time.sleep(1); sh(["node", os.path.join(G, "pdfs.js"), "8766"])
    finally:
        srv.terminate()
    ms(); print(sh([sys.executable, os.path.join(G, "conferir_publicacao.py"), "preparar"]).strip().splitlines()[-1])


if __name__ == "__main__":
    try:
        import openpyxl  # noqa: F401
    except ImportError:
        sys.exit("Falta o openpyxl: pip install openpyxl (ou PYTHONPATH apontando para ele)")
    nomes = [a for a in sys.argv[1:] if not a.startswith("--")]
    sel = [o for o in REG["obras"] if not nomes or o["obra"] in nomes]
    if nomes and len(sel) != len(nomes): sys.exit(f"obra desconhecida; use: {[o['obra'] for o in REG['obras']]}")
    os.makedirs(OUT, exist_ok=True)
    for o in sel: obra(o)
    if "--sem-site" not in sys.argv: site()
