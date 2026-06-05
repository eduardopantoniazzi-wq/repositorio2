"""
Integração com Google Sheets.
Cria ou atualiza a planilha com 3 abas: Dashboard CEO, Extrato Consolidado, Histórico de Alertas.
"""

from __future__ import annotations
import json
from datetime import datetime
from pathlib import Path
from typing import Optional

import gspread
from google.oauth2.service_account import Credentials
import pandas as pd

from config import (
    GOOGLE_CREDENTIALS_FILE,
    GOOGLE_SHEET_ID,
    GOOGLE_SHEET_NAME,
    ABA_DASHBOARD,
    ABA_EXTRATO,
    ABA_HISTORICO,
)


_SCOPES = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# Cores para formatação condicional (RGB 0-1)
_VERMELHO  = {"red": 0.96, "green": 0.26, "blue": 0.21}
_AMARELO   = {"red": 1.0,  "green": 0.90, "blue": 0.0}
_VERDE     = {"red": 0.20, "green": 0.66, "blue": 0.33}
_CINZA_BG  = {"red": 0.24, "green": 0.24, "blue": 0.24}
_BRANCO    = {"red": 1.0,  "green": 1.0,  "blue": 1.0}
_AZUL_CAB  = {"red": 0.07, "green": 0.36, "blue": 0.68}


def _conectar() -> gspread.Client:
    creds = Credentials.from_service_account_file(str(GOOGLE_CREDENTIALS_FILE), scopes=_SCOPES)
    return gspread.authorize(creds)


def _abrir_ou_criar_planilha(client: gspread.Client) -> gspread.Spreadsheet:
    if GOOGLE_SHEET_ID:
        return client.open_by_key(GOOGLE_SHEET_ID)
    # Tenta encontrar pelo nome
    try:
        return client.open(GOOGLE_SHEET_NAME)
    except gspread.SpreadsheetNotFound:
        sh = client.create(GOOGLE_SHEET_NAME)
        sh.share(None, perm_type="anyone", role="writer")  # Ajuste conforme necessário
        print(f"  ✔ Nova planilha criada: {sh.url}")
        return sh


def _garantir_aba(sh: gspread.Spreadsheet, nome: str) -> gspread.Worksheet:
    try:
        return sh.worksheet(nome)
    except gspread.WorksheetNotFound:
        return sh.add_worksheet(title=nome, rows=1000, cols=26)


def _formatar_cabecalho(ws: gspread.Worksheet, num_colunas: int) -> None:
    """Aplica cor e negrito no cabeçalho (linha 1)."""
    requests = [{
        "repeatCell": {
            "range": {
                "sheetId": ws.id,
                "startRowIndex": 0,
                "endRowIndex": 1,
                "startColumnIndex": 0,
                "endColumnIndex": num_colunas,
            },
            "cell": {
                "userEnteredFormat": {
                    "backgroundColor": _AZUL_CAB,
                    "textFormat": {"bold": True, "foregroundColor": _BRANCO},
                    "horizontalAlignment": "CENTER",
                }
            },
            "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)",
        }
    }]
    ws.spreadsheet.batch_update({"requests": requests})


def _colorir_alertas(ws: gspread.Worksheet, df: pd.DataFrame, col_alerta: int) -> None:
    """Pinta linhas com alertas de vermelho (CRÍTICO) ou amarelo (ATENÇÃO)."""
    requests = []
    for i, alerta in enumerate(df["alerta"].tolist()):
        if not alerta:
            continue
        row_idx = i + 1  # +1 pelo cabeçalho
        cor = _VERMELHO if "CRÍTICO" in str(alerta) else _AMARELO
        requests.append({
            "repeatCell": {
                "range": {
                    "sheetId": ws.id,
                    "startRowIndex": row_idx,
                    "endRowIndex": row_idx + 1,
                    "startColumnIndex": 0,
                    "endColumnIndex": len(df.columns),
                },
                "cell": {"userEnteredFormat": {"backgroundColor": cor}},
                "fields": "userEnteredFormat.backgroundColor",
            }
        })
    if requests:
        ws.spreadsheet.batch_update({"requests": requests})


def _df_para_lista(df: pd.DataFrame) -> list[list]:
    """Converte DataFrame para lista de listas, tratando datas e NaN."""
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


# ── Dashboard CEO ───────────────────────────────────────────────────────────

def _montar_dashboard(df: pd.DataFrame, alertas: list[dict]) -> list[list]:
    hoje = df["data"].max()
    df_hoje = df[df["data"] == hoje]

    saldo_bradesco = df[df["banco"] == "Bradesco"]["saldo"].iloc[-1] if not df[df["banco"] == "Bradesco"].empty else 0
    saldo_sicredi  = df[df["banco"] == "Sicredi"]["saldo"].iloc[-1] if not df[df["banco"] == "Sicredi"].empty else 0
    saldo_bb       = df[df["banco"] == "BB"]["saldo"].iloc[-1] if not df[df["banco"] == "BB"].empty else 0
    saldo_total    = saldo_bradesco + saldo_sicredi + saldo_bb

    entradas_dia = df_hoje["credito"].sum()
    saidas_dia   = df_hoje["debito"].sum()

    rows = [
        ["DASHBOARD CEO — MOINHO", ""],
        ["Atualizado em", datetime.now().strftime("%d/%m/%Y %H:%M")],
        ["", ""],
        ["SALDOS", ""],
        ["Bradesco", saldo_bradesco],
        ["Sicredi",  saldo_sicredi],
        ["BB",       saldo_bb],
        ["TOTAL CONSOLIDADO", saldo_total],
        ["", ""],
        ["MOVIMENTO DO DIA", str(hoje)[:10]],
        ["Total entradas", entradas_dia],
        ["Total saídas",   saidas_dia],
        ["Resultado",      entradas_dia - saidas_dia],
        ["", ""],
        ["ALERTAS", f"Total: {len(alertas)}"],
    ]

    if alertas:
        rows.append(["Nível", "Tipo", "Data", "Beneficiário", "Valor", "Descrição"])
        for a in alertas:
            rows.append([
                a["nivel"],
                a["tipo"],
                a["data_lancamento"],
                a["beneficiario"],
                a["valor"],
                a["descricao"],
            ])
    else:
        rows.append(["✔ Nenhum alerta encontrado", ""])

    return rows


# ── Função principal ─────────────────────────────────────────────────────────

def atualizar_sheets(df: pd.DataFrame, alertas: list[dict]) -> str:
    """Atualiza as 3 abas da planilha. Retorna a URL."""
    client = _conectar()
    sh = _abrir_ou_criar_planilha(client)

    # ── Aba 1: Dashboard CEO ─────────────────────────────────────────────
    ws_dash = _garantir_aba(sh, ABA_DASHBOARD)
    ws_dash.clear()
    dados_dash = _montar_dashboard(df, alertas)
    ws_dash.update("A1", dados_dash)

    # Formata título e saldo total
    requests_dash = [
        {  # Título em negrito grande
            "repeatCell": {
                "range": {"sheetId": ws_dash.id, "startRowIndex": 0, "endRowIndex": 1,
                           "startColumnIndex": 0, "endColumnIndex": 2},
                "cell": {"userEnteredFormat": {
                    "backgroundColor": _CINZA_BG,
                    "textFormat": {"bold": True, "fontSize": 14, "foregroundColor": _BRANCO},
                }},
                "fields": "userEnteredFormat(backgroundColor,textFormat)",
            }
        }
    ]
    sh.batch_update({"requests": requests_dash})

    # Pinta alertas vermelhos/amarelos no dashboard
    linha_inicio_alertas = 15  # linha onde começam os alertas
    for i, a in enumerate(alertas):
        cor = _VERMELHO if a["nivel"] == "CRITICO" else _AMARELO
        requests_dash2 = [{
            "repeatCell": {
                "range": {
                    "sheetId": ws_dash.id,
                    "startRowIndex": linha_inicio_alertas + i,
                    "endRowIndex": linha_inicio_alertas + i + 1,
                    "startColumnIndex": 0, "endColumnIndex": 6,
                },
                "cell": {"userEnteredFormat": {"backgroundColor": cor}},
                "fields": "userEnteredFormat.backgroundColor",
            }
        }]
        sh.batch_update({"requests": requests_dash2})

    # ── Aba 2: Extrato Consolidado ────────────────────────────────────────
    ws_ext = _garantir_aba(sh, ABA_EXTRATO)
    ws_ext.clear()
    dados_ext = _df_para_lista(df)
    ws_ext.update("A1", dados_ext)
    _formatar_cabecalho(ws_ext, len(df.columns))
    _colorir_alertas(ws_ext, df, col_alerta=df.columns.tolist().index("alerta"))

    # ── Aba 3: Histórico de Alertas ───────────────────────────────────────
    ws_hist = _garantir_aba(sh, ABA_HISTORICO)

    # Mantém histórico anterior e acrescenta novos
    dados_existentes = ws_hist.get_all_values()
    if not dados_existentes or dados_existentes == [[]]:
        cabecalho = [["data_alerta", "nivel", "tipo", "descricao", "banco",
                      "data_lancamento", "beneficiario", "valor", "documento"]]
        ws_hist.update("A1", cabecalho)
        _formatar_cabecalho(ws_hist, 9)
        proxima_linha = 2
    else:
        proxima_linha = len(dados_existentes) + 1

    if alertas:
        novas_linhas = []
        for a in alertas:
            novas_linhas.append([
                a["data_alerta"], a["nivel"], a["tipo"], a["descricao"],
                a["banco"], a["data_lancamento"], a["beneficiario"],
                a["valor"], a["documento"],
            ])
        ws_hist.update(f"A{proxima_linha}", novas_linhas)

        # Pinta novas linhas
        for i, a in enumerate(alertas):
            cor = _VERMELHO if a["nivel"] == "CRITICO" else _AMARELO
            sh.batch_update({"requests": [{
                "repeatCell": {
                    "range": {
                        "sheetId": ws_hist.id,
                        "startRowIndex": proxima_linha - 1 + i,
                        "endRowIndex": proxima_linha + i,
                        "startColumnIndex": 0, "endColumnIndex": 9,
                    },
                    "cell": {"userEnteredFormat": {"backgroundColor": cor}},
                    "fields": "userEnteredFormat.backgroundColor",
                }
            }]})

    url = f"https://docs.google.com/spreadsheets/d/{sh.id}"
    print(f"  ✔ Google Sheets atualizado: {url}")
    return url
