"""
Integração com Google Sheets.
Abas: Dashboard CEO | Extrato Consolidado | Histórico de Alertas | Fluxo de Caixa
"""

from __future__ import annotations
from datetime import datetime
from pathlib import Path
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

# Cores (0-1 RGB)
_VERMELHO = {"red": 0.96, "green": 0.26, "blue": 0.21}
_AMARELO  = {"red": 1.0,  "green": 0.90, "blue": 0.0}
_VERDE    = {"red": 0.20, "green": 0.66, "blue": 0.33}
_CINZA    = {"red": 0.24, "green": 0.24, "blue": 0.24}
_BRANCO   = {"red": 1.0,  "green": 1.0,  "blue": 1.0}
_AZUL     = {"red": 0.07, "green": 0.36, "blue": 0.68}
_VERDE_ESC = {"red": 0.0, "green": 0.39, "blue": 0.24}


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


def _formatar_cab(sh: gspread.Spreadsheet, ws: gspread.Worksheet, ncols: int, cor=None) -> None:
    cor = cor or _AZUL
    sh.batch_update({"requests": [{
        "repeatCell": {
            "range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1,
                      "startColumnIndex": 0, "endColumnIndex": ncols},
            "cell": {"userEnteredFormat": {
                "backgroundColor": cor,
                "textFormat": {"bold": True, "foregroundColor": _BRANCO},
                "horizontalAlignment": "CENTER",
            }},
            "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)",
        }
    }]})


def _colorir_linhas(sh: gspread.Spreadsheet, ws: gspread.Worksheet,
                    alertas_col: list[str], offset_row: int = 1) -> None:
    reqs = []
    for i, alerta in enumerate(alertas_col):
        if not alerta:
            continue
        cor = _VERMELHO if "CRÍTICO" in str(alerta) else _AMARELO
        reqs.append({"repeatCell": {
            "range": {"sheetId": ws.id,
                      "startRowIndex": offset_row + i,
                      "endRowIndex":   offset_row + i + 1,
                      "startColumnIndex": 0, "endColumnIndex": 20},
            "cell": {"userEnteredFormat": {"backgroundColor": cor}},
            "fields": "userEnteredFormat.backgroundColor",
        }})
    if reqs:
        sh.batch_update({"requests": reqs})


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


# ── Dashboard CEO ────────────────────────────────────────────────────────────

def _montar_dashboard(df: pd.DataFrame, alertas: list[dict]) -> list[list]:
    hoje = df["data"].max()
    df_hoje = df[df["data"] == hoje]

    bancos = ["Bradesco", "Sicredi", "BB"]
    saldos = {}
    for b in bancos:
        sub = df[df["banco"] == b]["saldo"]
        saldos[b] = sub.iloc[-1] if not sub.empty else 0.0
    total = sum(saldos.values())

    rows = [
        ["DASHBOARD CEO — MOINHO DE TRIGO", "", "", ""],
        [f"Atualizado: {datetime.now().strftime('%d/%m/%Y %H:%M')}", "", "", ""],
        ["", "", "", ""],
        ["SALDOS BANCÁRIOS", "Valor (R$)", "Unidade", ""],
        ["Bradesco  Ag.388 CC.7488-8",  round(saldos["Bradesco"], 2), "Santa Maria", ""],
        ["Sicredi   Coop.434 CC.34799-0", round(saldos["Sicredi"], 2), "Santa Maria", ""],
        ["BB        Ag.4044 CC.4699-X",   round(saldos["BB"], 2),      "Canoas",       ""],
        ["TOTAL CONSOLIDADO",             round(total, 2), "", ""],
        ["", "", "", ""],
        [f"MOVIMENTO DO DIA — {str(hoje)[:10]}", "", "", ""],
        ["Total entradas",  round(df_hoje["credito"].sum(), 2), "", ""],
        ["Total saídas",    round(df_hoje["debito"].sum(), 2),  "", ""],
        ["Resultado",       round(df_hoje["credito"].sum() - df_hoje["debito"].sum(), 2), "", ""],
        ["", "", "", ""],
        [f"ALERTAS — {len(alertas)} total", f"Críticos: {sum(1 for a in alertas if a['nivel']=='CRITICO')}",
         f"Atenção: {sum(1 for a in alertas if a['nivel']=='ATENCAO')}", ""],
    ]

    if alertas:
        rows.append(["Nível", "Tipo", "Beneficiário", "Valor (R$)"])
        for a in alertas[:100]:  # primeiros 100
            rows.append([a["nivel"], a["tipo"], a["beneficiario"], a["valor"]])
    else:
        rows.append(["✔ Nenhum alerta", "", "", ""])
    return rows


# ── Fluxo de caixa da planilha ───────────────────────────────────────────────

def _montar_fluxo(df_plan: pd.DataFrame) -> list[list]:
    if df_plan.empty:
        return [["Sem dados de planilha"]]

    rows = [["DATA", "UNIDADE", "DESCRIÇÃO", "DÉBITO (R$)", "CRÉDITO (R$)"]]
    for _, r in df_plan.sort_values(["data", "unidade" if "unidade" in df_plan.columns else "banco"]).iterrows():
        unidade = r.get("unidade", "") if "unidade" in df_plan.columns else ""
        rows.append([
            r["data"].strftime("%d/%m/%Y") if hasattr(r["data"], "strftime") else str(r["data"])[:10],
            str(unidade),
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

    # ── Aba 1: Dashboard CEO ─────────────────────────────────────────────
    ws_dash = _garantir_aba(sh, ABA_DASHBOARD, rows=500)
    ws_dash.clear()
    dados = _montar_dashboard(df, alertas)
    ws_dash.update("A1", dados)
    # Título em negrito grande
    sh.batch_update({"requests": [{
        "repeatCell": {
            "range": {"sheetId": ws_dash.id, "startRowIndex": 0, "endRowIndex": 1,
                      "startColumnIndex": 0, "endColumnIndex": 4},
            "cell": {"userEnteredFormat": {
                "backgroundColor": _CINZA,
                "textFormat": {"bold": True, "fontSize": 13, "foregroundColor": _BRANCO},
            }},
            "fields": "userEnteredFormat(backgroundColor,textFormat)",
        }
    }]})
    # Pinta alertas
    linha_alertas = 16
    for i, a in enumerate(alertas[:100]):
        cor = _VERMELHO if a["nivel"] == "CRITICO" else _AMARELO
        sh.batch_update({"requests": [{"repeatCell": {
            "range": {"sheetId": ws_dash.id,
                      "startRowIndex": linha_alertas + i,
                      "endRowIndex":   linha_alertas + i + 1,
                      "startColumnIndex": 0, "endColumnIndex": 4},
            "cell": {"userEnteredFormat": {"backgroundColor": cor}},
            "fields": "userEnteredFormat.backgroundColor",
        }}]})

    # ── Aba 2: Extrato Consolidado ────────────────────────────────────────
    ws_ext = _garantir_aba(sh, ABA_EXTRATO, rows=max(len(df) + 10, 1000))
    ws_ext.clear()
    ws_ext.update("A1", _df_para_lista(df))
    _formatar_cab(sh, ws_ext, len(df.columns))
    if "alerta" in df.columns:
        _colorir_linhas(sh, ws_ext, df["alerta"].fillna("").tolist())

    # ── Aba 3: Histórico de Alertas ───────────────────────────────────────
    ws_hist = _garantir_aba(sh, ABA_HISTORICO, rows=5000)
    existentes = ws_hist.get_all_values()
    cab = ["data_alerta", "nivel", "tipo", "descricao", "banco",
           "data_lancamento", "beneficiario", "valor", "documento"]
    if not existentes or existentes == [[]] or existentes[0] != cab:
        ws_hist.clear()
        ws_hist.update("A1", [cab])
        _formatar_cab(sh, ws_hist, len(cab))
        prox = 2
    else:
        prox = len(existentes) + 1

    if alertas:
        novas = [[a["data_alerta"], a["nivel"], a["tipo"], a["descricao"],
                  a["banco"], a["data_lancamento"], a["beneficiario"],
                  a["valor"], a["documento"]] for a in alertas]
        ws_hist.update(f"A{prox}", novas)
        for i, a in enumerate(alertas):
            cor = _VERMELHO if a["nivel"] == "CRITICO" else _AMARELO
            sh.batch_update({"requests": [{"repeatCell": {
                "range": {"sheetId": ws_hist.id,
                          "startRowIndex": prox - 1 + i,
                          "endRowIndex":   prox + i,
                          "startColumnIndex": 0, "endColumnIndex": len(cab)},
                "cell": {"userEnteredFormat": {"backgroundColor": cor}},
                "fields": "userEnteredFormat.backgroundColor",
            }}]})

    # ── Aba 4: Fluxo de Caixa (planilha) ─────────────────────────────────
    if df_planilha is not None and not df_planilha.empty:
        ws_fluxo = _garantir_aba(sh, ABA_FLUXO, rows=max(len(df_planilha) + 10, 1000))
        ws_fluxo.clear()
        dados_fluxo = _montar_fluxo(df_planilha)
        ws_fluxo.update("A1", dados_fluxo)
        _formatar_cab(sh, ws_fluxo, 5, cor=_VERDE_ESC)

    url = f"https://docs.google.com/spreadsheets/d/{sh.id}"
    return url
