"""
Integração com Google Sheets.
Abas: Dashboard CEO | Extrato Consolidado | Histórico de Alertas | Fluxo de Caixa

Todas as formatações são enviadas em um único batch_update por aba,
evitando rate limit da API.
"""

from __future__ import annotations
import time
from datetime import datetime
from typing import Optional

import gspread
from google.oauth2.service_account import Credentials
import pandas as pd

from config import (
    GOOGLE_CREDENTIALS_FILE, GOOGLE_SHEET_ID, GOOGLE_SHEET_NAME,
    ABA_DASHBOARD, ABA_EXTRATO, ABA_HISTORICO,
)

ABA_FLUXO = "Fluxo de Caixa"

_SCOPES = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

_VERMELHO  = {"red": 0.96, "green": 0.26, "blue": 0.21}
_AMARELO   = {"red": 1.0,  "green": 0.90, "blue": 0.0}
_CINZA     = {"red": 0.24, "green": 0.24, "blue": 0.24}
_BRANCO    = {"red": 1.0,  "green": 1.0,  "blue": 1.0}
_AZUL      = {"red": 0.07, "green": 0.36, "blue": 0.68}
_VERDE_ESC = {"red": 0.0,  "green": 0.39, "blue": 0.24}


def _conectar() -> gspread.Client:
    creds = Credentials.from_service_account_file(str(GOOGLE_CREDENTIALS_FILE), scopes=_SCOPES)
    return gspread.authorize(creds)


def _abrir_ou_criar(client: gspread.Client) -> gspread.Spreadsheet:
    if GOOGLE_SHEET_ID:
        return client.open_by_key(GOOGLE_SHEET_ID)
    try:
        return client.open(GOOGLE_SHEET_NAME)
    except gspread.SpreadsheetNotFound:
        sh = client.create(GOOGLE_SHEET_NAME)
        print(f"  ✔ Nova planilha criada: {sh.url}")
        return sh


def _garantir_aba(sh: gspread.Spreadsheet, nome: str, rows: int = 2000) -> gspread.Worksheet:
    try:
        return sh.worksheet(nome)
    except gspread.WorksheetNotFound:
        return sh.add_worksheet(title=nome, rows=rows, cols=26)


def _req_cor_linha(sheet_id: int, row_idx: int, ncols: int, cor: dict) -> dict:
    """Gera request de coloração para uma linha (0-indexed)."""
    return {"repeatCell": {
        "range": {
            "sheetId": sheet_id,
            "startRowIndex": row_idx, "endRowIndex": row_idx + 1,
            "startColumnIndex": 0, "endColumnIndex": ncols,
        },
        "cell": {"userEnteredFormat": {"backgroundColor": cor}},
        "fields": "userEnteredFormat.backgroundColor",
    }}


def _req_cabecalho(sheet_id: int, ncols: int, cor: dict = None) -> dict:
    cor = cor or _AZUL
    return {"repeatCell": {
        "range": {
            "sheetId": sheet_id,
            "startRowIndex": 0, "endRowIndex": 1,
            "startColumnIndex": 0, "endColumnIndex": ncols,
        },
        "cell": {"userEnteredFormat": {
            "backgroundColor": cor,
            "textFormat": {"bold": True, "foregroundColor": _BRANCO},
            "horizontalAlignment": "CENTER",
        }},
        "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)",
    }}


def _req_autofit(sheet_id: int) -> dict:
    return {"autoResizeDimensions": {
        "dimensions": {"sheetId": sheet_id, "dimension": "COLUMNS",
                       "startIndex": 0, "endIndex": 20}
    }}


def _batch(sh: gspread.Spreadsheet, requests: list[dict]) -> None:
    """Envia requisições em lote com retry em caso de rate limit."""
    if not requests:
        return
    for tentativa in range(4):
        try:
            sh.batch_update({"requests": requests})
            return
        except gspread.exceptions.APIError as e:
            if "429" in str(e) and tentativa < 3:
                espera = 20 * (tentativa + 1)
                print(f"  ⏳ Rate limit — aguardando {espera}s...")
                time.sleep(espera)
            else:
                raise


def _df_para_lista(df: pd.DataFrame) -> list[list]:
    rows = [df.columns.tolist()]
    for _, row in df.iterrows():
        linha = []
        for v in row:
            if pd.isna(v):
                linha.append("")
            elif hasattr(v, "strftime"):
                linha.append(v.strftime("%d/%m/%Y"))
            elif isinstance(v, float):
                linha.append(round(v, 2))
            else:
                linha.append(str(v))
        rows.append(linha)
    return rows


def _escrever_com_retry(ws: gspread.Worksheet, celula: str, dados: list[list]) -> None:
    """Escreve dados com retry em caso de rate limit."""
    for tentativa in range(4):
        try:
            ws.update(celula, dados)
            return
        except gspread.exceptions.APIError as e:
            if "429" in str(e) and tentativa < 3:
                espera = 20 * (tentativa + 1)
                print(f"  ⏳ Rate limit (escrita) — aguardando {espera}s...")
                time.sleep(espera)
            else:
                raise


# ── Dashboard CEO ────────────────────────────────────────────────────────────

def _montar_dashboard(df: pd.DataFrame, alertas: list[dict]) -> list[list]:
    hoje = df["data"].max()
    df_hoje = df[df["data"] == hoje]

    saldos = {}
    for b in ["Bradesco", "Sicredi", "BB"]:
        sub = df[df["banco"] == b]["saldo"]
        saldos[b] = sub.iloc[-1] if not sub.empty else 0.0
    total = sum(saldos.values())

    criticos = sum(1 for a in alertas if a["nivel"] == "CRITICO")
    atencoes = sum(1 for a in alertas if a["nivel"] == "ATENCAO")

    rows = [
        ["DASHBOARD CEO — MOINHO DE TRIGO", "", "", ""],
        [f"Atualizado: {datetime.now().strftime('%d/%m/%Y %H:%M')}", "", "", ""],
        ["", "", "", ""],
        ["SALDOS BANCÁRIOS", "Valor (R$)", "Unidade", ""],
        ["Bradesco  Ag.388 CC.7488-8",       round(saldos["Bradesco"], 2), "Santa Maria", ""],
        ["Sicredi   Coop.434 CC.34799-0",    round(saldos["Sicredi"],  2), "Santa Maria", ""],
        ["BB        Ag.4044 CC.4699-X",       round(saldos["BB"],      2), "Canoas",       ""],
        ["TOTAL CONSOLIDADO",                 round(total, 2),             "",             ""],
        ["", "", "", ""],
        [f"MOVIMENTO — {str(hoje)[:10]}", "", "", ""],
        ["Total entradas",  round(df_hoje["credito"].sum(), 2), "", ""],
        ["Total saídas",    round(df_hoje["debito"].sum(),  2), "", ""],
        ["Resultado",       round(df_hoje["credito"].sum() - df_hoje["debito"].sum(), 2), "", ""],
        ["", "", "", ""],
        [f"ALERTAS — {len(alertas)} total",
         f"Críticos: {criticos}", f"Atenção: {atencoes}", ""],
    ]

    if alertas:
        rows.append(["Nível", "Tipo", "Beneficiário", "Valor (R$)"])
        for a in alertas[:200]:
            rows.append([a["nivel"], a["tipo"], a["beneficiario"], a["valor"]])
    else:
        rows.append(["✔ Nenhum alerta encontrado", "", "", ""])

    return rows


def _montar_fluxo(df_plan: pd.DataFrame) -> list[list]:
    if df_plan.empty:
        return [["Sem dados de planilha"]]
    rows = [["DATA", "UNIDADE", "DESCRIÇÃO", "DÉBITO (R$)", "CRÉDITO (R$)"]]
    col_unidade = "unidade" if "unidade" in df_plan.columns else "banco"
    for _, r in df_plan.sort_values(["data", col_unidade]).iterrows():
        rows.append([
            r["data"].strftime("%d/%m/%Y") if hasattr(r["data"], "strftime") else str(r["data"])[:10],
            str(r.get(col_unidade, "")),
            str(r["descricao"]),
            round(r["debito"], 2) if r["debito"] else "",
            round(r["credito"], 2) if r["credito"] else "",
        ])
    return rows


# ── Função principal ─────────────────────────────────────────────────────────

def atualizar_sheets(
    df: pd.DataFrame,
    alertas: list[dict],
    df_planilha: Optional[pd.DataFrame] = None,
) -> str:
    client = _conectar()
    sh = _abrir_ou_criar(client)
    url = f"https://docs.google.com/spreadsheets/d/{sh.id}"

    # ── Aba 1: Dashboard CEO ─────────────────────────────────────────────
    print("    → Escrevendo Dashboard CEO...")
    ws_dash = _garantir_aba(sh, ABA_DASHBOARD, rows=300)
    ws_dash.clear()
    dados_dash = _montar_dashboard(df, alertas)
    _escrever_com_retry(ws_dash, "A1", dados_dash)

    # Monta todos os requests de formatação do dashboard em um único lote
    reqs_dash = [
        # Título grande
        {"repeatCell": {
            "range": {"sheetId": ws_dash.id, "startRowIndex": 0, "endRowIndex": 1,
                      "startColumnIndex": 0, "endColumnIndex": 4},
            "cell": {"userEnteredFormat": {
                "backgroundColor": _CINZA,
                "textFormat": {"bold": True, "fontSize": 13, "foregroundColor": _BRANCO},
            }},
            "fields": "userEnteredFormat(backgroundColor,textFormat)",
        }},
        # Cabeçalho de saldos (linha 4)
        _req_cabecalho(ws_dash.id, 4),
    ]
    # Colorir alertas (todos de uma vez)
    linha_alerta_inicio = 16  # linha 17 (0-indexed=16)
    for i, a in enumerate(alertas[:200]):
        cor = _VERMELHO if a["nivel"] == "CRITICO" else _AMARELO
        reqs_dash.append(_req_cor_linha(ws_dash.id, linha_alerta_inicio + i, 4, cor))
    _batch(sh, reqs_dash)
    time.sleep(2)

    # ── Aba 2: Extrato Consolidado ────────────────────────────────────────
    print("    → Escrevendo Extrato Consolidado...")
    ws_ext = _garantir_aba(sh, ABA_EXTRATO, rows=max(len(df) + 10, 500))
    ws_ext.clear()
    time.sleep(1)
    _escrever_com_retry(ws_ext, "A1", _df_para_lista(df))

    reqs_ext = [_req_cabecalho(ws_ext.id, len(df.columns))]
    if "alerta" in df.columns:
        for i, alerta in enumerate(df["alerta"].fillna("").tolist()):
            if alerta:
                cor = _VERMELHO if "CRÍTICO" in str(alerta) else _AMARELO
                reqs_ext.append(_req_cor_linha(ws_ext.id, i + 1, len(df.columns), cor))
    _batch(sh, reqs_ext)
    time.sleep(2)

    # ── Aba 3: Histórico de Alertas ───────────────────────────────────────
    print("    → Escrevendo Histórico de Alertas...")
    ws_hist = _garantir_aba(sh, ABA_HISTORICO, rows=5000)
    cab = ["data_alerta", "nivel", "tipo", "descricao", "banco",
           "data_lancamento", "beneficiario", "valor", "documento"]

    existentes = ws_hist.get_all_values()
    if not existentes or existentes[0] != cab:
        ws_hist.clear()
        _escrever_com_retry(ws_hist, "A1", [cab])
        _batch(sh, [_req_cabecalho(ws_hist.id, len(cab))])
        prox = 2
    else:
        prox = len(existentes) + 1

    if alertas:
        novas = [[a["data_alerta"], a["nivel"], a["tipo"], a["descricao"],
                  a["banco"], a["data_lancamento"], a["beneficiario"],
                  a["valor"], a["documento"]] for a in alertas]
        time.sleep(1)
        _escrever_com_retry(ws_hist, f"A{prox}", novas)

        reqs_hist = []
        for i, a in enumerate(alertas):
            cor = _VERMELHO if a["nivel"] == "CRITICO" else _AMARELO
            reqs_hist.append(_req_cor_linha(ws_hist.id, prox - 1 + i, len(cab), cor))
        _batch(sh, reqs_hist)
    time.sleep(2)

    # ── Aba 4: Fluxo de Caixa ────────────────────────────────────────────
    if df_planilha is not None and not df_planilha.empty:
        print("    → Escrevendo Fluxo de Caixa...")
        ws_fluxo = _garantir_aba(sh, ABA_FLUXO, rows=max(len(df_planilha) + 10, 500))
        ws_fluxo.clear()
        time.sleep(1)
        dados_fluxo = _montar_fluxo(df_planilha)
        _escrever_com_retry(ws_fluxo, "A1", dados_fluxo)
        _batch(sh, [_req_cabecalho(ws_fluxo.id, 5, cor=_VERDE_ESC)])

    return url
