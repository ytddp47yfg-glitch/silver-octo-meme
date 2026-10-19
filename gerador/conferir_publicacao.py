"""Confere se o que mudou no site local foi mesmo para a página publicada (artifact do claude.ai).

Rotina obrigatória a cada entrada de modificações (ver CLAUDE.md):

  1. python3 gerador/conferir_publicacao.py preparar
       - impressão digital (sha256) de todos os arquivos que a página usa: index_artifact.html, obras/<id>.json,
         obras/<id>.json.card.json, pdf/<id>.pdf, xlsx/<id>.xlsx.b64.txt
       - conferências locais (corte do cartão = corte dos dados, painéis que sumiram, b64 = .xlsx, .xlsx do site = entregável
         da raiz, data de geração de hoje)
       - compara com site/.publicado.json (último estado CONFIRMADO no ar) e grava site/.pendente.json com o mapa `files`
         para o publish (só o que mudou; em lotes de até 60 MB) e a lista de caminhos a reler
  2. publicar a página com o mapa de site/.pendente.json (Artifact publish, file_path = index_artifact, files = lote)
  3. reler da página publicada os caminhos listados (Artifact read com `paths`, out_dir = pasta de conferência)
  4. python3 gerador/conferir_publicacao.py conferir <pasta de conferência>
       - compara byte a byte cada arquivo relido com o local; só então atualiza site/.publicado.json
       - sai com código 1 (e lista o que falta) se algo não bate ou ainda não foi publicado

Status: python3 gerador/conferir_publicacao.py status   (o que está diferente do último estado confirmado)
Gancho: python3 gerador/conferir_publicacao.py gancho   (Stop hook: código 2 enquanto houver pendência de publicação)
"""
import base64, datetime as dt, hashlib, json, os, sys

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SITE = os.path.join(RAIZ, "site")
PUB, PEND = os.path.join(SITE, ".publicado.json"), os.path.join(SITE, ".pendente.json")
PAGINA = "index_artifact.html"
LOTE = 60 * 1024 * 1024
ENTREGAVEIS = {   # .xlsx do site → entregável na raiz do repositório
    "botanico": "fluxo_caixa_botanico_modelo.xlsx", "botanico-padrao": "fluxo_caixa_botanico_modelo_PADRAO.xlsx",
    "ed-jardim": "fluxo_caixa_ed_jardim_modelo.xlsx", "ed-jardim-padrao": "fluxo_caixa_ed_jardim_modelo_PADRAO.xlsx",
    "ed-jardim-ajustado": "fluxo_caixa_ed_jardim_modelo_PREVISION_AJUSTADO.xlsx",
    "torre-catharina": "fluxo_caixa_torre_catharina_modelo.xlsx", "torre-catharina-padrao": "fluxo_caixa_torre_catharina_modelo_PADRAO.xlsx",
    "torre-catharina-simulacao": "fluxo_caixa_torre_catharina_modelo_SIMULACAO.xlsx"}


def sha(p):
    h = hashlib.sha256()
    with open(p, "rb") as f:
        for b in iter(lambda: f.read(1 << 20), b""): h.update(b)
    return h.hexdigest()


def ids():
    return sorted(f[:-len(".json.card.json")] for f in os.listdir(os.path.join(SITE, "obras")) if f.endswith(".json.card.json"))


def arquivos():
    """{caminho publicado: caminho local} de tudo que a página usa (a página em si fica como PAGINA)."""
    m = {PAGINA: os.path.join(SITE, PAGINA)}
    for i in ids():
        pdfs = sorted(f"pdf/{f}" for f in os.listdir(os.path.join(SITE, "pdf")) if f.startswith(i + "--") and f.endswith(".pdf"))
        for pub in [f"obras/{i}.json", f"obras/{i}.json.card.json", f"xlsx/{i}.xlsx.b64.txt"] + pdfs:
            loc = os.path.join(SITE, pub)
            if os.path.exists(loc): m[pub] = loc
    return m


def manifesto():
    return {pub: dict(sha=sha(loc), kb=round(os.path.getsize(loc) / 1024)) for pub, loc in arquivos().items()}


def conferencias(man, anterior):
    erros, avisos = [], []
    hoje = dt.date.today().isoformat()
    for i in ids():
        card = json.load(open(os.path.join(SITE, f"obras/{i}.json.card.json")))
        dados = json.load(open(os.path.join(SITE, f"obras/{i}.json")))
        if dados.get("DASH", {}).get("corte") != card.get("corte"):
            erros.append(f"{i}: corte do cartão ({card.get('corte')}) ≠ corte dos dados ({dados.get('DASH', {}).get('corte')})")
        if card.get("gerado") != hoje: avisos.append(f"{i}: dados gerados em {card.get('gerado')} (não hoje)")
        paineis = sorted(k for k in ("FIS", "GANTT", "ESQ") if k in dados)
        antes = (anterior.get("_paineis") or {}).get(i)
        if antes is not None and set(antes) - set(paineis):
            erros.append(f"{i}: painel sumiu em relação ao publicado: {sorted(set(antes) - set(paineis))}")
        man.setdefault("_paineis", {})[i] = paineis
        x, b = os.path.join(SITE, f"xlsx/{i}.xlsx"), os.path.join(SITE, f"xlsx/{i}.xlsx.b64.txt")
        if os.path.exists(x):
            if not os.path.exists(b) or base64.b64decode(open(b, "rb").read()) != open(x, "rb").read():
                erros.append(f"{i}: xlsx/{i}.xlsx.b64.txt não corresponde a xlsx/{i}.xlsx (refaça o base64)")
            ent = os.path.join(RAIZ, ENTREGAVEIS.get(i, ""))
            if i in ENTREGAVEIS and os.path.exists(ent) and sha(ent) != sha(x):
                erros.append(f"{i}: site/xlsx/{i}.xlsx ≠ {ENTREGAVEIS[i]} (entregável da raiz)")
        pdfs = [f for f in os.listdir(os.path.join(SITE, "pdf")) if f.startswith(i + "--") and f.endswith(".pdf")]
        npain = 2 + len(paineis)   # DASHBOARD, RETRATO + FIS/GANTT/ESQ
        if len(pdfs) != npain: erros.append(f"{i}: {len(pdfs)} PDF(s) de painel, esperados {npain} (rode gerador/pdfs.js)")
        for f in pdfs:
            if os.path.getmtime(os.path.join(SITE, "pdf", f)) < os.path.getmtime(os.path.join(SITE, f"obras/{i}.json")):
                erros.append(f"{i}: pdf/{f} mais antigo que os dados (rode gerador/pdfs.js)")
    return erros, avisos


def carregar(p):
    return json.load(open(p)) if os.path.exists(p) else {}


def diferencas(man, pub):
    return [k for k in man if not k.startswith("_") and (k not in pub or pub[k]["sha"] != man[k]["sha"])]


def preparar():
    pub = carregar(PUB); man = manifesto()
    erros, avisos = conferencias(man, pub)
    dif = diferencas(man, pub)
    remover = sorted(k for k in pub if not k.startswith("_") and k not in man)
    arq = arquivos()
    lotes, atual, tam = [], {k: None for k in remover}, 0
    for k in [d for d in dif if d != PAGINA]:
        s = os.path.getsize(arq[k])
        if atual and tam + s > LOTE: lotes.append(atual); atual, tam = {}, 0
        atual[k] = arq[k]; tam += s
    if atual or PAGINA in dif or not lotes: lotes.append(atual)
    json.dump(dict(pagina=arq[PAGINA], lotes=lotes, reler=dif, remover=remover, manifesto=man, gerado=dt.datetime.now().isoformat(timespec="seconds")),
              open(PEND, "w"), ensure_ascii=False, indent=1)
    for a in avisos: print("aviso |", a)
    for e in erros: print("ERRO  |", e)
    print(f"{len(dif)} arquivo(s) diferente(s) do último estado confirmado no ar: {', '.join(dif) or '—'}")
    if remover: print(f"{len(remover)} arquivo(s) a remover da página (null no lote 1): {', '.join(remover)}")
    print(f"{len(lotes)} lote(s) de publicação em site/.pendente.json")
    sys.exit(1 if erros else 0)


def conferir(pasta):
    pend = carregar(PEND)
    if not pend: sys.exit("rode 'preparar' antes")
    man = pend["manifesto"]; pub = carregar(PUB)
    ok, falhas = [], []
    for k in pend["reler"]:
        cand = [os.path.join(pasta, k), os.path.join(pasta, os.path.basename(k))]
        if k == PAGINA: cand += [os.path.join(pasta, "index.html")]
        loc = next((c for c in cand if os.path.exists(c)), None)
        if loc is None: falhas.append(f"{k}: não relido da página publicada"); continue
        if k == PAGINA:   # a página publicada pode vir embrulhada: confere se o conteúdo local está contido nela
            local = open(arquivos()[PAGINA], "rb").read().strip(); lido = open(loc, "rb").read()
            (ok if local in lido or sha(loc) == man[k]["sha"] else falhas).append(k if local in lido or sha(loc) == man[k]["sha"] else f"{k}: conteúdo diferente do local")
            continue
        if sha(loc) == man[k]["sha"]: ok.append(k)
        else: falhas.append(f"{k}: publicado ≠ local ({os.path.getsize(loc) // 1024} KB × {man[k]['kb']} KB)")
    for k in ok: pub[k] = man[k]
    for k in pend.get("remover", []):   # removidos: confirmar com a listagem de arquivos da página (scope files)
        if os.path.exists(os.path.join(pasta, k)): falhas.append(f"{k}: deveria ter sido removido e ainda foi relido")
        else: pub.pop(k, None)
    pub["_paineis"] = man.get("_paineis", {}); pub["_conferido"] = dt.datetime.now().isoformat(timespec="seconds")
    json.dump(pub, open(PUB, "w"), ensure_ascii=False, indent=1)
    rest = diferencas(manifesto(), pub)
    print(f"conferidos no ar: {len(ok)} de {len(pend['reler'])}")
    for f in falhas: print("FALHA |", f)
    if rest: print("ainda diferente do publicado:", ", ".join(rest))
    print("OK: página publicada = arquivos locais" if not falhas and not rest else "ATENÇÃO: a página publicada NÃO está igual aos arquivos locais")
    sys.exit(0 if not falhas and not rest else 1)


def status():
    pub = carregar(PUB); man = manifesto()
    dif = diferencas(man, pub)
    print(f"último estado confirmado no ar: {pub.get('_conferido', 'nunca')}")
    print(f"diferente do publicado: {', '.join(dif) or 'nada'}")
    sys.exit(1 if dif else 0)


def gancho():
    """Gancho Stop do Claude Code (.claude/settings.json): bloqueia o fim do turno (código 2) enquanto houver arquivo
    do site diferente do último estado confirmado no ar."""
    if not os.path.exists(os.path.join(SITE, PAGINA)): sys.exit(0)
    dif = diferencas(manifesto(), carregar(PUB))
    if not dif: sys.exit(0)
    print(f"Página publicada desatualizada: {len(dif)} arquivo(s) do site mudaram e não foram conferidos no ar "
          f"({', '.join(dif[:6])}{' …' if len(dif) > 6 else ''}). Siga a rotina do CLAUDE.md: "
          "conferir_publicacao.py preparar → publicar os lotes → reler os caminhos → conferir_publicacao.py conferir <pasta>.",
          file=sys.stderr)
    sys.exit(2)


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "status"
    {"preparar": preparar, "status": status, "gancho": gancho}.get(cmd, lambda: conferir(sys.argv[2]))()
