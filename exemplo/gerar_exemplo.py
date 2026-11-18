"""Gera uma planilha e um mapa de áreas FICTÍCIOS em exemplo/dados/ (nenhum dado real)."""
import csv
import pathlib

import openpyxl

SAIDA = pathlib.Path(__file__).resolve().parent / "dados"
SAIDA.mkdir(exist_ok=True)

COLUNAS = ["ID", "Empresa", "Nicho", "Cidade", "UF", "Telefone", "Site", "Status", "Responsável"]
LEADS = [
    ["L0001", "Café Aurora", "Restaurante", "Orlando", "FL", "+1 555 010 0101", "cafeaurora.example", "Novo", "Ana"],
    ["L0002", "Studio Lume", "Pilates", "Boston", "MA", "+1 555 010 0102", "studiolume.example", "Novo", "Ana"],
    ["L0003", "Clínica Vitta", "Saúde", "Miami", "FL", "+1 555 010 0103", "clinicavitta.example", "Contatado", "Ana"],
    ["L0004", "Oficina Central", "Mecânica", "Newark", "NJ", "+1 555 010 0104", "oficinacentral.example", "Novo", "Bruno"],
    ["L0005", "Bela Casa Decor", "Decoração", "Atlanta", "GA", "+1 555 010 0105", "belacasa.example", "Novo", "Bruno"],
]
DESCARTADOS = [
    ["", "Auto Brilho", "Estética automotiva", "Houston", "TX", "+1 555 010 0201", "autobrilho.example", "Descartado", "Ana"],
]
AREAS = [
    {"id": "FL-ORLANDO", "nome": "Orlando e arredores", "status": "livre"},
    {"id": "MA-BOSTON", "nome": "Grande Boston", "status": "livre"},
    {"id": "NJ-NEWARK", "nome": "Newark e arredores", "status": "livre"},
    {"id": "GA-ATLANTA", "nome": "Atlanta metropolitana", "status": "livre"},
]

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Leads"
ws.append(COLUNAS)
for r in LEADS:
    ws.append(r)
wd = wb.create_sheet("Descartados")
wd.append(COLUNAS)
for r in DESCARTADOS:
    wd.append(r)
wb.save(SAIDA / "planilha-mestre.xlsx")

with open(SAIDA / "mapa-de-areas.csv", "w", encoding="utf-8-sig", newline="") as f:
    campos = ["id", "nome", "status", "reservada_por", "reservada_em", "ultima_por", "descansa_ate"]
    w = csv.DictWriter(f, fieldnames=campos, delimiter=";", lineterminator="\r\n")
    w.writeheader()
    for a in AREAS:
        w.writerow({**{c: "" for c in campos}, **a})
print("Exemplo gerado em", SAIDA)
