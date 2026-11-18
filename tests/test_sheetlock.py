"""Testes de ponta a ponta do CLI usando uma cópia temporária dos dados de exemplo."""
import json
import os
import pathlib
import subprocess
import sys

RAIZ = pathlib.Path(__file__).resolve().parents[1]


def rodar(pasta, quem, *args):
    env = dict(os.environ, SHEETLOCK_QUEM=quem, SHEETLOCK_PASTA=str(pasta))
    return subprocess.run([sys.executable, str(RAIZ / "sheetlock.py"), *args], capture_output=True, text=True,
                          encoding="utf-8", env=env)


def preparar(tmp_path):
    subprocess.run([sys.executable, str(RAIZ / "exemplo" / "gerar_exemplo.py")], check=True, capture_output=True)
    import shutil
    dest = tmp_path / "dados"
    shutil.copytree(RAIZ / "exemplo" / "dados", dest)
    return dest


def test_trava_bloqueia_o_outro(tmp_path):
    d = preparar(tmp_path)
    assert "Trava pega" in rodar(d, "Ana", "trava", "pegar", "FL-ORLANDO").stdout
    r = rodar(d, "Bruno", "trava", "pegar")
    assert r.returncode != 0 and "RECUSADO" in r.stderr
    assert rodar(d, "Ana", "trava", "soltar").returncode == 0
    assert rodar(d, "Bruno", "trava", "pegar").returncode == 0


def test_deduplicacao(tmp_path):
    d = preparar(tmp_path)
    novos = [{"Empresa": "Café Aurora", "Cidade": "Orlando"},          # duplicata por nome+cidade
             {"Empresa": "Outro Nome", "Telefone": "+1 555 010 0102"},  # duplicata por telefone
             {"Empresa": "Loja Nova", "Cidade": "Dallas", "Site": "https://www.lojanova.example/x"}]
    arq = tmp_path / "novos.json"
    arq.write_text(json.dumps(novos), encoding="utf-8")
    r = rodar(d, "Ana", "novos-leads", str(arq))
    assert "1 linha(s) gravada(s)" in r.stdout, r.stdout + r.stderr
    assert "Café Aurora" in r.stdout and "Outro Nome" in r.stdout
    # rodar de novo não duplica
    assert "0 linha(s)" in rodar(d, "Ana", "novos-leads", str(arq)).stdout


def test_recusa_gravar_com_copia_em_conflito(tmp_path):
    d = preparar(tmp_path)
    (d / "planilha-mestre (cópia em conflito).xlsx").write_bytes(b"x")
    arq = tmp_path / "n.json"
    arq.write_text(json.dumps([{"Empresa": "Nova Co", "Cidade": "X"}]), encoding="utf-8")
    r = rodar(d, "Ana", "novos-leads", str(arq))
    assert r.returncode != 0 and "NADA GRAVADO" in (r.stderr + r.stdout)


def test_backup_do_dia(tmp_path):
    d = preparar(tmp_path)
    arq = tmp_path / "n.json"
    arq.write_text(json.dumps([{"Empresa": "Nova Co", "Cidade": "X"}]), encoding="utf-8")
    assert rodar(d, "Ana", "novos-leads", str(arq)).returncode == 0
    assert list((tmp_path / "dados-backup").glob("planilha-mestre-*.xlsx"))


def test_reserva_de_area(tmp_path):
    d = preparar(tmp_path)
    assert rodar(d, "Ana", "reservar", "FL-ORLANDO").returncode == 0
    r = rodar(d, "Bruno", "reservar", "FL-ORLANDO")
    assert r.returncode != 0 and "RECUSADO" in r.stderr
