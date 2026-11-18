"""sheetlock: CLI para duas pessoas (dois PCs) editarem a MESMA planilha numa pasta sincronizada sem se atropelar.

O que faz:
  - TRAVA de rodada: só uma pessoa por vez trabalha (expira sozinha em N horas).
  - DEDUPLICAÇÃO: linha nova que repete nome+cidade, telefone ou site é pulada.
  - SEGURANÇA DE GRAVAÇÃO: recusa gravar se existir cópia em conflito da nuvem, se a planilha estiver aberta no Excel
    ou se ela mudou desde que foi lida; guarda um BACKUP do dia; grava por cópia validada.
  - MAPA DE ÁREAS (CSV): quem está trabalhando em cada área, reservas e "descanso" de áreas que repetiram leads.

Configuração (por PC), em `sheetlock.local.json` (NÃO versionar):  {"quem": "Ana", "pasta": "C:/caminho/da/pasta/sincronizada"}
ou variáveis SHEETLOCK_QUEM e SHEETLOCK_PASTA.

Uso:  python sheetlock.py [--teste PASTA] <comando> ...
  diagnostico                         confere pasta, quem sou eu, Excel aberto, cópias em conflito e trava
  procurar "Nome 1" "Nome 2"          procura em Leads e Descartados (sem acento, por pedaço do nome)
  lead L0012                          mostra a linha de um lead (só colunas preenchidas)
  onde "Coluna=valor" [...]           lista leads por filtro
  novos-leads arquivo.json            acrescenta linhas em Leads (pula duplicata; preenche Responsável)
  descartes arquivo.json              acrescenta linhas em Descartados
  atualizar L0012 "Coluna=valor" ...  muda colunas de um lead
  trava pegar|soltar|ver [area]       trava de rodada (uma por vez entre os dois PCs)
  areas                               lista áreas livres e bloqueadas (com o motivo)
  reservar AREA                       reserva a área para mim (o outro PC não a escolhe)
  descansar AREA [dias]               área que repetiu leads do outro: descansa (padrão 90 dias)
  --teste PASTA                       (antes do comando) trabalha numa cópia em PASTA, sem tocar nos originais
"""
import csv
import datetime
import io
import json
import os
import pathlib
import re
import shutil
import sys
import tempfile
import time
import unicodedata

sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")

AQUI = pathlib.Path(__file__).resolve().parent
LOCAL = AQUI / "sheetlock.local.json"
NOME_XLSX = "planilha-mestre.xlsx"
NOME_CSV = "mapa-de-areas.csv"
TRAVA_HORAS = 4
RESERVA_HORAS = 6
# Cada pessoa tem uma faixa de IDs para não colidirem mesmo com a nuvem atrasada.
FAIXA_ID = {"Bruno": 5000}  # Ana: L0001-L4999, Bruno: L5001+
COLUNAS_NOVAS = ["Responsável"]
CAMPOS_MAPA = ["id", "nome", "status", "reservada_por", "reservada_em", "ultima_por", "descansa_ate"]


def _config() -> dict:
    cfg = {}
    if LOCAL.exists():
        try:
            cfg = json.loads(LOCAL.read_text(encoding="utf-8-sig"))
        except json.JSONDecodeError as e:
            sys.exit(f"ERRO: {LOCAL.name} com JSON inválido ({e}).")
    cfg["quem"] = os.environ.get("SHEETLOCK_QUEM", cfg.get("quem", ""))
    cfg["pasta"] = os.environ.get("SHEETLOCK_PASTA", cfg.get("pasta", str(AQUI / "exemplo" / "dados")))
    return cfg


CFG = _config()
QUEM = str(CFG["quem"]).strip().capitalize()
PASTA = pathlib.Path(CFG["pasta"])
BACKUP = PASTA.parent / (PASTA.name + "-backup")
TESTE = False

args = sys.argv[1:]
if len(args) >= 2 and args[0] == "--teste":
    t = pathlib.Path(args[1])
    t.mkdir(parents=True, exist_ok=True)
    for nome in (NOME_XLSX, NOME_CSV):
        if (PASTA / nome).exists() and not (t / nome).exists():
            shutil.copy2(PASTA / nome, t / nome)
    PASTA, BACKUP, TESTE = t, t / "_backup", True
    args = args[2:]

XLSX = PASTA / NOME_XLSX
CSV = PASTA / NOME_CSV
TRAVA = PASTA / "_trava-rodada.json"


# ------------------------------------------------------------------ utilidades de texto
def norm(s) -> str:
    s = unicodedata.normalize("NFKD", str(s or "")).encode("ascii", "ignore").decode().lower()
    return " ".join("".join(c if c.isalnum() else " " for c in s).split())


def tel(s) -> str:
    d = re.sub(r"\D", "", str(s or ""))
    return d[-10:] if len(d) >= 10 else ""


def dominio(s) -> str:
    s = re.sub(r"^[a-z]+://", "", str(s or "").strip().lower())
    s = re.sub(r"^www\.", "", s).split("/")[0].split("?")[0]
    return s if "." in s else ""


def chaves(r: dict) -> set[str]:
    """Identidades de uma linha: nome+cidade, telefone e domínio do site."""
    out = set()
    if r.get("Empresa"):
        out.add("n:" + norm(r["Empresa"]) + "|" + norm(r.get("Cidade")))
    if tel(r.get("Telefone")):
        out.add("t:" + tel(r["Telefone"]))
    if dominio(r.get("Site")):
        out.add("s:" + dominio(r["Site"]))
    return out


# ------------------------------------------------------------------ segurança da pasta compartilhada
def checar_pasta():
    if not XLSX.exists() or not CSV.exists():
        sys.exit(f"ERRO: não achei a planilha em {PASTA}.\nConfira se o app de sincronização está aberto e logado "
                 "e rode de novo. Nada foi gravado.")


def conflitos() -> list[str]:
    """Cópias em conflito criadas pelo serviço de nuvem quando dois PCs gravam ao mesmo tempo."""
    if TESTE:
        return []
    out = []
    for p in PASTA.iterdir():
        if p.name.startswith("~$") or p in (XLSX, CSV):
            continue
        if re.match(r"(?i)^planilha-mestre.*\.xlsx$", p.name) or re.match(r"(?i)^mapa-de-areas.*\.csv$", p.name):
            out.append(p.name)
    return out


def excel_aberto() -> bool:
    return (PASTA / ("~$" + XLSX.name)).exists()


def antes_de_gravar(alvo: pathlib.Path):
    c = conflitos()
    if c:
        sys.exit("NADA GRAVADO. Existem cópias em conflito: " + ", ".join(c) + ". Junte o que faltar na principal, "
                 "apague a cópia e rode de novo.")
    if alvo == XLSX and excel_aberto():
        sys.exit("NADA GRAVADO. A planilha está aberta no Excel (existe ~$" + XLSX.name + "). Feche e rode de novo.")


CARIMBO: dict = {}


def carimbar(p: pathlib.Path):
    st = p.stat()
    CARIMBO[p] = (st.st_mtime_ns, st.st_size)


def conferir_carimbo(p: pathlib.Path):
    if p in CARIMBO:
        st = p.stat()
        if (st.st_mtime_ns, st.st_size) != CARIMBO[p]:
            sys.exit(f"NADA GRAVADO. {p.name} mudou enquanto eu trabalhava (outra pessoa gravou). "
                     "Rode o mesmo comando de novo: ele relê a versão nova.")


def copia_do_dia(p: pathlib.Path):
    BACKUP.mkdir(parents=True, exist_ok=True)
    alvo = BACKUP / f"{p.stem}-{datetime.date.today():%Y-%m-%d}{p.suffix}"
    if not alvo.exists():
        shutil.copy2(p, alvo)


def gravar_por_copia(p: pathlib.Path, escrever, validar):
    """Escreve num temporário local, valida e só então copia por cima do arquivo compartilhado."""
    antes_de_gravar(p)
    conferir_carimbo(p)
    copia_do_dia(p)
    fd, tmp = tempfile.mkstemp(suffix=p.suffix, dir=BACKUP)
    os.close(fd)
    tmp = pathlib.Path(tmp)
    try:
        escrever(tmp)
        validar(tmp)
        for tentativa in range(4):
            try:
                conferir_carimbo(p)
                shutil.copyfile(tmp, p)
                break
            except PermissionError:
                if tentativa == 3:
                    sys.exit(f"ERRO: {p.name} está travado (Excel aberto ou sincronizando). Nada foi gravado.")
                time.sleep(3)
    finally:
        tmp.unlink(missing_ok=True)
    carimbar(p)


# ------------------------------------------------------------------ planilha e mapa
def abrir():
    import openpyxl
    checar_pasta()
    carimbar(XLSX)
    return openpyxl.load_workbook(XLSX)


def salvar(wb):
    import openpyxl
    gravar_por_copia(XLSX, lambda t: wb.save(t), lambda t: openpyxl.load_workbook(t, read_only=True).close())


def cab(ws) -> list:
    return [c.value for c in ws[1]]


def linhas(ws):
    h = cab(ws)
    for i, r in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if any(v not in (None, "") for v in r):
            yield i, dict(zip(h, r))


def mostrar(d: dict) -> str:
    return " | ".join(f"{k}: {v}" for k, v in d.items() if v not in (None, ""))


def ler_json(p):
    return json.loads(pathlib.Path(p).read_text(encoding="utf-8-sig"))


def ler_csv():
    checar_pasta()
    carimbar(CSV)
    rd = csv.DictReader(io.StringIO(CSV.read_text(encoding="utf-8-sig")), delimiter=";")
    return list(rd.fieldnames), list(rd)


def gravar_csv(campos, rows):
    buf = io.StringIO()
    w = csv.DictWriter(buf, fieldnames=campos, delimiter=";", lineterminator="\r\n", extrasaction="ignore")
    w.writeheader()
    w.writerows(rows)

    def valida(t):
        n = len(list(csv.DictReader(io.StringIO(t.read_text(encoding="utf-8-sig")), delimiter=";")))
        if n != len(rows):
            raise SystemExit("ERRO interno: o mapa validado não tem o mesmo número de áreas. Nada foi gravado.")
    gravar_por_copia(CSV, lambda t: t.write_text(buf.getvalue(), encoding="utf-8-sig"), valida)


def precisa_quem() -> str:
    if not QUEM:
        sys.exit(f"ERRO: este PC não diz quem está rodando. Crie {LOCAL.name} com {{\"quem\": \"Ana\"}}.")
    return QUEM


def proximo_id(ws) -> str:
    ids = [int(m.group(1)) for _, r in linhas(ws) if (m := re.match(r"L(\d+)$", str(r.get("ID") or "")))]
    base = FAIXA_ID.get(QUEM, 0)
    faixa = [i for i in ids if base < i < base + 5000] if base else [i for i in ids if i < 5000]
    return f"L{(max(faixa) if faixa else base) + 1:04d}"


# ------------------------------------------------------------------ trava de rodada
def _trava() -> dict:
    try:
        return json.loads(TRAVA.read_text(encoding="utf-8"))
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def _vigente(t: dict) -> bool:
    inicio = t.get("em")
    try:
        return bool(t) and (datetime.datetime.now() - datetime.datetime.fromisoformat(inicio)).total_seconds() < TRAVA_HORAS * 3600
    except (TypeError, ValueError):
        return False


def cmd_trava(a):
    acao = a[0] if a else "ver"
    t = _trava()
    if acao == "ver":
        if _vigente(t):
            print(f"TRAVA de {t['quem']} desde {t['em']} (área {t.get('area', '-')}); vale {TRAVA_HORAS} h.")
        else:
            print("TRAVA livre.")
    elif acao == "pegar":
        quem = precisa_quem()
        if _vigente(t) and t["quem"] != quem:
            sys.exit(f"RECUSADO: a trava é de {t['quem']} desde {t['em']}. Espere ou peça para soltar.")
        TRAVA.write_text(json.dumps({"quem": quem, "em": datetime.datetime.now().isoformat(timespec="seconds"),
                                     "area": a[1] if len(a) > 1 else ""}, ensure_ascii=False), encoding="utf-8")
        print(f"Trava pega por {quem}.")
    elif acao == "soltar":
        if _vigente(t) and t["quem"] != precisa_quem():
            sys.exit(f"RECUSADO: a trava é de {t['quem']}.")
        TRAVA.unlink(missing_ok=True)
        print("Trava solta.")
    else:
        sys.exit("Use: trava pegar|soltar|ver [area]")


# ------------------------------------------------------------------ comandos
def cmd_diagnostico(_):
    print(f"Pasta: {PASTA} ({'existe' if PASTA.exists() else 'NÃO existe'})")
    print(f"Quem sou eu: {QUEM or '(não configurado)'}")
    print(f"Planilha: {'ok' if XLSX.exists() else 'AUSENTE'} | Mapa: {'ok' if CSV.exists() else 'AUSENTE'}")
    print(f"Excel aberto: {'SIM' if excel_aberto() else 'não'}")
    print(f"Cópias em conflito: {conflitos() or 'nenhuma'}")
    cmd_trava(["ver"])


def cmd_procurar(a):
    wb = abrir()
    alvo = [norm(x) for x in a]
    for aba in ("Leads", "Descartados"):
        for i, r in linhas(wb[aba]):
            if any(x in norm(r.get("Empresa")) for x in alvo):
                print(f"[{aba} linha {i}] {mostrar(r)}")


def cmd_lead(a):
    wb = abrir()
    for _, r in linhas(wb["Leads"]):
        if r.get("ID") == a[0]:
            print(mostrar(r))
            return
    sys.exit(f"Lead {a[0]} não encontrado.")


def cmd_onde(a):
    filtros = [x.split("=", 1) for x in a]
    wb = abrir()
    for _, r in linhas(wb["Leads"]):
        if all(str(r.get(c, "")) == v for c, v in filtros):
            print(mostrar(r))


def _acrescentar(aba: str, arq: str):
    wb = abrir()
    quem = precisa_quem()
    ws = wb[aba]
    h = cab(ws)
    existentes = set()
    for nome in ("Leads", "Descartados"):
        for _, r in linhas(wb[nome]):
            existentes |= chaves(r)
    novos = ler_json(arq)
    gravados, pulados = 0, []
    for r in novos:
        k = chaves(r)
        if k & existentes:
            pulados.append(r.get("Empresa", "?"))
            continue
        r = dict(r)
        if aba == "Leads":
            r.setdefault("ID", proximo_id(ws))
        r.setdefault("Responsável", quem)
        ws.append([r.get(c) for c in h])
        existentes |= k
        gravados += 1
    salvar(wb)
    print(f"{gravados} linha(s) gravada(s) em {aba}. Duplicatas puladas: {pulados or 'nenhuma'}.")


def cmd_atualizar(a):
    wb = abrir()
    ws = wb["Leads"]
    h = cab(ws)
    for i, r in linhas(ws):
        if r.get("ID") == a[0]:
            for par in a[1:]:
                c, v = par.split("=", 1)
                if c not in h:
                    sys.exit(f"Coluna '{c}' não existe.")
                ws.cell(row=i, column=h.index(c) + 1, value=v)
            salvar(wb)
            print(f"{a[0]} atualizado.")
            return
    sys.exit(f"Lead {a[0]} não encontrado.")


def _bloqueio(r: dict, quem: str) -> str:
    hoje = datetime.datetime.now()
    if r.get("descansa_ate"):
        try:
            if datetime.datetime.fromisoformat(r["descansa_ate"]) > hoje:
                return f"descansando até {r['descansa_ate'][:10]}"
        except ValueError:
            pass
    if r.get("reservada_por") and r["reservada_por"] != quem and r.get("reservada_em"):
        try:
            if (hoje - datetime.datetime.fromisoformat(r["reservada_em"])).total_seconds() < RESERVA_HORAS * 3600:
                return f"reservada por {r['reservada_por']}"
        except ValueError:
            pass
    return ""


def cmd_areas(_):
    _, rows = ler_csv()
    for r in rows:
        motivo = _bloqueio(r, QUEM)
        print(f"{'BLOQUEADA' if motivo else 'livre    '} {r['id']:14} {r['nome'][:30]:30} {motivo}")


def cmd_reservar(a):
    quem = precisa_quem()
    campos, rows = ler_csv()
    for r in rows:
        if r["id"] == a[0]:
            motivo = _bloqueio(r, quem)
            if motivo:
                sys.exit(f"RECUSADO: {a[0]} está {motivo}.")
            r["reservada_por"], r["reservada_em"] = quem, datetime.datetime.now().isoformat(timespec="seconds")
            gravar_csv(campos, rows)
            print(f"{a[0]} reservada para {quem}.")
            return
    sys.exit(f"Área {a[0]} não existe.")


def cmd_descansar(a):
    dias = int(a[1]) if len(a) > 1 else 90
    campos, rows = ler_csv()
    for r in rows:
        if r["id"] == a[0]:
            r["descansa_ate"] = (datetime.datetime.now() + datetime.timedelta(days=dias)).isoformat(timespec="seconds")
            gravar_csv(campos, rows)
            print(f"{a[0]} descansa por {dias} dias.")
            return
    sys.exit(f"Área {a[0]} não existe.")


COMANDOS = {"diagnostico": cmd_diagnostico, "procurar": cmd_procurar, "lead": cmd_lead, "onde": cmd_onde,
            "novos-leads": lambda a: _acrescentar("Leads", a[0]), "descartes": lambda a: _acrescentar("Descartados", a[0]),
            "atualizar": cmd_atualizar, "trava": cmd_trava, "areas": cmd_areas, "reservar": cmd_reservar,
            "descansar": cmd_descansar}

if __name__ == "__main__":
    if not args or any(x in ("-h", "--help", "ajuda") for x in args) or args[0] not in COMANDOS:
        print(__doc__)
        sys.exit(0)
    COMANDOS[args[0]](args[1:])
