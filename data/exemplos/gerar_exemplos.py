"""
Gera arquivos CSV de exemplo para os três bancos.
Execute: python data/exemplos/gerar_exemplos.py
"""

import csv
from pathlib import Path

SAIDA = Path(__file__).parent.parent / "extratos"
SAIDA.mkdir(parents=True, exist_ok=True)


# ── Bradesco ──────────────────────────────────────────────────────────────
bradesco_rows = [
    ["Data", "Lançamento", "Dcto", "Crédito", "Débito", "Saldo"],
    ["02/06/2025", "DEPOSITO PIX CLIENTE ABC LTDA",   "00001", "85000,00",  "",          "185000,00"],
    ["02/06/2025", "PAGTO FORN TRIGO SUL COMERCIO",   "00002", "",          "12500,00",  "172500,00"],
    ["02/06/2025", "PAGTO FORN TRIGO SUL COMERCIO",   "00003", "",          "12500,00",  "160000,00"],  # ALERTA: duplo no dia
    ["03/06/2025", "PIX RECEBIDO EMPRESA XYZ S.A.",   "00004", "15000,00",  "",          "175000,00"],  # ALERTA: crédito desconhecido
    ["03/06/2025", "PAGTO FORNECEDOR NOVO LTDA",      "00005", "",          "75000,00",  "100000,00"],  # ALERTA: alto sem histórico + novo
    ["04/06/2025", "DEPOSITO CLIENTE DEF INDUSTRIA",  "00006", "20000,00",  "",          "120000,00"],
    ["04/06/2025", "PAGTO ENERGIA ELETRICA RGE",      "00007", "",          "8900,00",   "111100,00"],
    ["05/06/2025", "PAGTO FORN FARINHA GAUCHA SA",    "00008", "",          "32000,00",  "79100,00"],   # ALERTA: variação 60% (média histórica ~20k)
    ["05/06/2025", "PAGTO FOLHA SALARIOS JUNHO",      "00009", "",          "45000,00",  "34100,00"],
]

with open(SAIDA / "bradesco.csv", "w", newline="", encoding="utf-8") as f:
    writer = csv.writer(f, delimiter=";")
    writer.writerows(bradesco_rows)
print("✔ bradesco.csv gerado")


# ── Sicredi ───────────────────────────────────────────────────────────────
sicredi_rows = [
    ["Data", "Descrição", "Documento", "Valor", "Saldo"],
    ["02/06/2025", "DEPOSITO CLIENTE GHI ALIMENTOS",  "S001", "40000,00",  "140000,00"],
    ["02/06/2025", "PAG FORNECEDOR EMBALAGENS PLUS",  "S002", "-9800,00",  "130200,00"],
    ["03/06/2025", "PAG ALUGUEL GALPAO SANTA MARIA",  "S003", "-5500,00",  "124700,00"],
    ["03/06/2025", "DEPOSITO CLIENTE ABC LTDA",       "S004", "18000,00",  "142700,00"],
    ["04/06/2025", "PAG MANUTENCAO MAQUINARIO",       "S005", "-2300,00",  "140400,00"],
    ["04/06/2025", "PAG FORNECEDOR EMBALAGENS PLUS",  "S006", "-9800,00",  "130600,00"],
    ["05/06/2025", "TRANSFERENCIA RECEBIDA NOVA CORP","S007", "5000,00",   "135600,00"],  # ALERTA: crédito desconhecido
    ["05/06/2025", "PAG SEGURO FROTA",                "S008", "-3200,00",  "132400,00"],
]

with open(SAIDA / "sicredi.csv", "w", newline="", encoding="utf-8") as f:
    writer = csv.writer(f, delimiter=";")
    writer.writerows(sicredi_rows)
print("✔ sicredi.csv gerado")


# ── Banco do Brasil ───────────────────────────────────────────────────────
bb_rows = [
    ["Dt.balancete", "Dt.movimento", "Histórico", "Documento", "Valor", "Saldo"],
    ["02/06/2025", "02/06/2025", "DEPOSITO CLIENTE JKL TRADING",  "B001", "30000,00",  "230000,00"],
    ["02/06/2025", "02/06/2025", "PAG IMPOSTOS FEDERAIS",         "B002", "-22000,00", "208000,00"],
    ["03/06/2025", "03/06/2025", "PAG FORN GRAOS DO SUL LTDA",    "B003", "-18500,00", "189500,00"],
    ["03/06/2025", "03/06/2025", "CREDITO PIX CNPJ DESCONHECIDO", "B004", "8000,00",   "197500,00"],  # ALERTA
    ["04/06/2025", "04/06/2025", "PAG FORN GRAOS DO SUL LTDA",    "B005", "-18500,00", "179000,00"],
    ["04/06/2025", "04/06/2025", "DEPOSITO CLIENTE MNO FARINHA",  "B006", "12000,00",  "191000,00"],
    ["05/06/2025", "05/06/2025", "PAG COMBUSTIVEL POSTO CENTRAL", "B007", "-4500,00",  "186500,00"],
    ["05/06/2025", "05/06/2025", "PAG FRETE TRANSPORTADORA TOP",  "B008", "-6800,00",  "179700,00"],
]

with open(SAIDA / "bb.csv", "w", newline="", encoding="utf-8") as f:
    writer = csv.writer(f, delimiter=";")
    writer.writerows(bb_rows)
print("✔ bb.csv gerado")

print(f"\nArquivos salvos em: {SAIDA}")
