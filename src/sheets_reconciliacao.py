"""
Google Sheets focado em conciliação: previsto x efetivado.
Estrutura:
  Aba 1 — RESUMO CEO         : painel simples com números e alertas críticos
  Aba 2 — NÃO PREVISTOS      : pagamentos feitos sem estar na planilha (risco de desvio)
  Aba 3 — CONCILIAÇÃO        : previstos x pagos lado a lado com status
  Aba 4 — PENDENTES          : previstos que ainda não saíram
  Aba 5 — EXTRATO COMPLETO   : todos os lançamentos bancários brutos
"""

from __future__ import annotations
import time
from datetime import datetime
from typing import Optional

import gspread
from google.oauth2.service_account import Credentials
import pandas as pd

from config import GOOGLE_CREDENTIALS_FILE, GOOGLE_SHEET_ID, GOOGLE_SHEET_NAME

_SCOPES = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# Paleta de cores
_VERMELHO  = {"red": 0.96, "green": 0.26, "blue": 0.21}
_VERMELHO_C= {"red": 1.0,  "green": 0.80, "blue": 0.80}
_AMARELO   = {"red": 1.0,  "green": 0.93, "blue": 0.60}
_VERDE     = {"red": 0.72, "green": 0.93, "blue": 0.74}
_VERDE_ESC = {"red": 0.0,  "green": 0.50, "blue": 0.25}
_AZUL      = {"red": 0.07, "green": 0.36, "blue": 0.68}
_CINZA_ESC = {"red": 0.24, "green": 0.24, "blue": 0.24}
_CINZA_CLA = {"red": 0.93, "green": 0.93, "blue": 0.93}
_LARANJA   = {"red": 1.0,  "green": 0.60, "blue": 0.0}
_BRANCO    = {"red": 1.0,  "green": 1.0,  "blue": 1.0}
_PRETO     = {"red": 0.0,  "green": 0.0,  "blue": 0.0}


def _conectar() -> gspread.Client:
    creds = Credentials.from_service_account_file(str(GOOGLE_CREDENTIALS_FILE), scopes=_SCOPES)
    return gspread.authorize(creds)


def _abrir(client: gspread.Client) -> gspread.Spreadsheet:
    return client.open_by_key(GOOGLE_SHEET_ID)


def _aba(sh: gspread.Spreadsheet, nome: str, rows: int = 2000) -> gspread.Worksheet:
    try:
        ws = sh.worksheet(nome)
        ws.clear()
        return ws
    except gspread.WorksheetNotFound:
        return sh.add_worksheet(title=nome, rows=rows, cols=20)


def _batch(sh: gspread.Spreadsheet, reqs: list[dict]) -> None:
    if not reqs:
        return
    for t in range(4):
        try:
            sh.batch_update({"requests": reqs})
            return
        except gspread.exceptions.APIError as e:
            if "429" in str(e) and t < 3:
                w = 20 * (t + 1)
                print(f"      ⏳ rate limit — aguarda {w}s...")
                time.sleep(w)
            else:
                raise


def _escrever(ws: gspread.Worksheet, dados: list[list], celula: str = "A1") -> None:
    for t in range(4):
        try:
            ws.update(celula, dados, value_input_option="USER_ENTERED")
            return
        except gspread.exceptions.APIError as e:
            if "429" in str(e) and t < 3:
                w = 20 * (t + 1)
                print(f"      ⏳ rate limit — aguarda {w}s...")
                time.sleep(w)
            else:
                raise


def _req_cab(sid: int, ncols: int, cor=None, linha: int = 0) -> dict:
    return {"repeatCell": {
        "range": {"sheetId": sid, "startRowIndex": linha, "endRowIndex": linha + 1,
                  "startColumnIndex": 0, "endColumnIndex": ncols},
        "cell": {"userEnteredFormat": {
            "backgroundColor": cor or _AZUL,
            "textFormat": {"bold": True, "foregroundColor": _BRANCO, "fontSize": 10},
            "horizontalAlignment": "CENTER",
        }},
        "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)",
    }}


def _req_linha(sid: int, row: int, ncols: int, cor: dict) -> dict:
    return {"repeatCell": {
        "range": {"sheetId": sid, "startRowIndex": row, "endRowIndex": row + 1,
                  "startColumnIndex": 0, "endColumnIndex": ncols},
        "cell": {"userEnteredFormat": {"backgroundColor": cor}},
        "fields": "userEnteredFormat.backgroundColor",
    }}


def _req_negrito(sid: int, row: int, ncols: int) -> dict:
    return {"repeatCell": {
        "range": {"sheetId": sid, "startRowIndex": row, "endRowIndex": row + 1,
                  "startColumnIndex": 0, "endColumnIndex": ncols},
        "cell": {"userEnteredFormat": {"textFormat": {"bold": True}}},
        "fields": "userEnteredFormat.textFormat.bold",
    }}


def _req_freeze(sid: int, linhas: int = 1) -> dict:
    return {"updateSheetProperties": {
        "properties": {"sheetId": sid, "gridProperties": {"frozenRowCount": linhas}},
        "fields": "gridProperties.frozenRowCount",
    }}


def _fmt_data(v) -> str:
    if v is None or (isinstance(v, float) and v != v):
        return ""
    if hasattr(v, "strftime"):
        return v.strftime("%d/%m/%Y")
    return str(v)[:10]


def _fmt_r(v) -> str:
    if not v and v != 0:
        return ""
    try:
        return f"R$ {float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(v)


# ════════════════════════════════════════════════════════════════════════════
# ABA 1 — RESUMO CEO
# ════════════════════════════════════════════════════════════════════════════

def _escrever_resumo(sh, df_conc, df_nao_prev, df_banco, periodo: str):
    ws = _aba(sh, "📊 RESUMO CEO", rows=60)
    sid = ws.id

    total_previsto  = df_conc["valor_previsto"].sum()
    total_pago      = df_banco[df_banco["debito"] > 0]["debito"].sum()
    conciliados     = df_conc[df_conc["status"] == "CONCILIADO"]
    divergencias    = df_conc[df_conc["status"].str.startswith("DIVERG", na=False)]
    pendentes       = df_conc[df_conc["status"] == "PENDENTE"]
    nao_prev_alto   = df_nao_prev[df_nao_prev["risco"] == "ALTO"] if not df_nao_prev.empty else pd.DataFrame()

    saldos = {}
    for b in ["Bradesco", "Sicredi", "BB"]:
        sub = df_banco[df_banco["banco"] == b]["saldo"]
        saldos[b] = sub.iloc[-1] if not sub.empty else 0.0

    dados = [
        ["CONTROLE FINANCEIRO — MOINHO DE TRIGO", "", "", ""],
        [f"Período: {periodo}    |    Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}", "", "", ""],
        ["", "", "", ""],

        ["SALDOS BANCÁRIOS", "", "", ""],
        ["Bradesco (Ag.388 / CC.7488-8)", _fmt_r(saldos.get("Bradesco", 0)), "Santa Maria", ""],
        ["Sicredi (Coop.434 / CC.34799-0)", _fmt_r(saldos.get("Sicredi", 0)), "Santa Maria", ""],
        ["BB (Ag.4044 / CC.4699-X)", _fmt_r(saldos.get("BB", 0)), "Canoas", ""],
        ["TOTAL DISPONÍVEL", _fmt_r(sum(saldos.values())), "", ""],
        ["", "", "", ""],

        ["CONCILIAÇÃO DO PERÍODO", "", "", ""],
        ["Total PREVISTO (planilha)", _fmt_r(total_previsto), "", ""],
        ["Total PAGO (extratos)", _fmt_r(total_pago), "", ""],
        ["Diferença", _fmt_r(total_pago - total_previsto), "", ""],
        ["", "", "", ""],

        ["STATUS DOS PAGAMENTOS PREVISTOS", "Qtd", "Valor (R$)", ""],
        ["✅ Conciliados (previsto = pago)", len(conciliados), _fmt_r(conciliados["valor_previsto"].sum()), ""],
        ["⚠️  Divergências (valor/beneficiário diferente)", len(divergencias), _fmt_r(divergencias["valor_previsto"].sum()), ""],
        ["🕐 Pendentes (previstos não pagos ainda)", len(pendentes), _fmt_r(pendentes["valor_previsto"].sum()), ""],
        ["", "", "", ""],

        ["🚨 PAGAMENTOS NÃO PREVISTOS (verificar imediatamente!)", "Qtd", "Valor Total", ""],
        ["Risco ALTO (transferências/PIX sem previsão)", len(nao_prev_alto), _fmt_r(nao_prev_alto["valor"].sum() if not nao_prev_alto.empty else 0), "VER ABA: NÃO PREVISTOS"],
        ["Total não previstos", len(df_nao_prev), _fmt_r(df_nao_prev["valor"].sum() if not df_nao_prev.empty else 0), ""],
        ["", "", "", ""],
    ]

    # Lista dos não previstos de alto risco
    if not nao_prev_alto.empty:
        dados.append(["⛔ PAGAMENTOS DE ALTO RISCO SEM PREVISÃO:", "", "", ""])
        dados.append(["Data", "Banco", "Beneficiário", "Valor"])
        for _, r in nao_prev_alto.iterrows():
            dados.append([_fmt_data(r["data"]), r["banco"], r["descricao"], _fmt_r(r["valor"])])

    _escrever(ws, dados)

    reqs = [
        # Título
        {"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": 0, "endRowIndex": 1,
                      "startColumnIndex": 0, "endColumnIndex": 4},
            "cell": {"userEnteredFormat": {
                "backgroundColor": _CINZA_ESC,
                "textFormat": {"bold": True, "fontSize": 14, "foregroundColor": _BRANCO},
            }},
            "fields": "userEnteredFormat(backgroundColor,textFormat)",
        }},
        _req_cab(sid, 4, _AZUL, 3),       # saldos
        _req_cab(sid, 4, _AZUL, 9),        # conciliação
        _req_cab(sid, 4, _AZUL, 14),       # status
        _req_cab(sid, 4, _VERMELHO, 18),   # não previstos
        _req_linha(sid, 7, 4, _CINZA_CLA), # total disponível
        _req_linha(sid, 12, 4, _AMARELO),  # diferença
        _req_linha(sid, 15, 4, _VERDE),    # conciliados
        _req_linha(sid, 16, 4, _AMARELO),  # divergências
        _req_linha(sid, 17, 4, _CINZA_CLA),# pendentes
        _req_linha(sid, 19, 4, _VERMELHO_C),# risco alto
        _req_linha(sid, 20, 4, _VERMELHO_C),# total nao prev
        _req_freeze(sid, 1),
    ]

    # Pinta linhas de alto risco no final
    if not nao_prev_alto.empty:
        base = 23
        for i in range(len(nao_prev_alto)):
            reqs.append(_req_linha(sid, base + i, 4, _VERMELHO_C))

    _batch(sh, reqs)
    time.sleep(2)
    print("    ✔ Resumo CEO")


# ════════════════════════════════════════════════════════════════════════════
# ABA 2 — NÃO PREVISTOS
# ════════════════════════════════════════════════════════════════════════════

def _escrever_nao_previstos(sh, df_nao_prev):
    ws = _aba(sh, "🚨 NÃO PREVISTOS", rows=max(len(df_nao_prev) + 10, 100))
    sid = ws.id

    cab = ["DATA", "BANCO", "BENEFICIÁRIO / DESCRIÇÃO", "VALOR (R$)", "DOCUMENTO", "RISCO"]
    dados = [cab]

    if df_nao_prev.empty:
        dados.append(["✅ Nenhum pagamento fora da planilha encontrado.", "", "", "", "", ""])
    else:
        # Ordena por risco e valor
        ordem = {"ALTO": 0, "MEDIO": 1, "BAIXO": 2, "OPERACIONAL": 3}
        df_ord = df_nao_prev.copy()
        df_ord["_ord"] = df_ord["risco"].map(ordem).fillna(4)
        df_ord = df_ord.sort_values(["_ord", "valor"], ascending=[True, False])

        for _, r in df_ord.iterrows():
            dados.append([
                _fmt_data(r["data"]),
                r["banco"],
                r["descricao"],
                _fmt_r(r["valor"]),
                r.get("documento", ""),
                r["risco"],
            ])

    _escrever(ws, dados)

    reqs = [_req_cab(sid, 6, _VERMELHO), _req_freeze(sid, 1)]

    if not df_nao_prev.empty:
        df_ord2 = df_nao_prev.copy()
        df_ord2["_ord"] = df_ord2["risco"].map({"ALTO": 0, "MEDIO": 1, "BAIXO": 2, "OPERACIONAL": 3}).fillna(4)
        df_ord2 = df_ord2.sort_values(["_ord", "valor"], ascending=[True, False])
        for i, (_, r) in enumerate(df_ord2.iterrows()):
            cor_map = {"ALTO": _VERMELHO, "MEDIO": _LARANJA, "BAIXO": _AMARELO, "OPERACIONAL": _CINZA_CLA}
            cor = cor_map.get(r["risco"], _CINZA_CLA)
            reqs.append(_req_linha(sid, i + 1, 6, cor))

    _batch(sh, reqs)
    time.sleep(2)
    print(f"    ✔ Não previstos ({len(df_nao_prev)} pagamentos)")


# ════════════════════════════════════════════════════════════════════════════
# ABA 3 — CONCILIAÇÃO COMPLETA
# ════════════════════════════════════════════════════════════════════════════

def _escrever_conciliacao(sh, df_conc):
    ws = _aba(sh, "🔍 CONCILIAÇÃO", rows=max(len(df_conc) + 10, 200))
    sid = ws.id

    cab = [
        "DATA PREVISTA", "UNIDADE", "DESCRIÇÃO PLANILHA", "VALOR PREVISTO",
        "DATA PAGA", "BANCO", "BENEFICIÁRIO EXTRATO", "VALOR PAGO",
        "DIFERENÇA", "STATUS",
    ]
    dados = [cab]

    cor_status = {
        "CONCILIADO":              _VERDE,
        "DIVERGENCIA_VALOR":       _AMARELO,
        "DIVERGENCIA_BENEFICIARIO":_LARANJA,
        "DIVERGENCIA":             _LARANJA,
        "PENDENTE":                _CINZA_CLA,
    }

    if df_conc.empty:
        dados.append(["Sem dados de planilha no período.", "", "", "", "", "", "", "", "", ""])
    else:
        df_ord = df_conc.sort_values(["data_prevista", "unidade"])
        for _, r in df_ord.iterrows():
            dados.append([
                _fmt_data(r["data_prevista"]),
                r.get("unidade", ""),
                r["descricao_planilha"],
                _fmt_r(r["valor_previsto"]),
                _fmt_data(r["data_pagamento"]),
                r["banco"],
                r["descricao_extrato"],
                _fmt_r(r["valor_pago"]) if r["valor_pago"] else "",
                _fmt_r(r["diferenca"]) if r["diferenca"] else "",
                r["status"],
            ])

    _escrever(ws, dados)

    reqs = [_req_cab(sid, 10), _req_freeze(sid, 1)]
    if not df_conc.empty:
        df_ord2 = df_conc.sort_values(["data_prevista", "unidade"])
        for i, (_, r) in enumerate(df_ord2.iterrows()):
            cor = cor_status.get(r["status"], _BRANCO)
            reqs.append(_req_linha(sid, i + 1, 10, cor))

    _batch(sh, reqs)
    time.sleep(2)
    print(f"    ✔ Conciliação ({len(df_conc)} previstos)")


# ════════════════════════════════════════════════════════════════════════════
# ABA 4 — PENDENTES
# ════════════════════════════════════════════════════════════════════════════

def _escrever_pendentes(sh, df_pend):
    ws = _aba(sh, "🕐 PENDENTES", rows=max(len(df_pend) + 10, 100))
    sid = ws.id

    cab = ["DATA PREVISTA", "UNIDADE", "DESCRIÇÃO", "VALOR PREVISTO"]
    dados = [cab]

    if df_pend.empty:
        dados.append(["✅ Nenhum pagamento previsto pendente.", "", "", ""])
    else:
        for _, r in df_pend.sort_values("data_prevista").iterrows():
            dados.append([
                _fmt_data(r["data_prevista"]),
                r.get("unidade", ""),
                r["descricao_planilha"],
                _fmt_r(r["valor_previsto"]),
            ])

    _escrever(ws, dados)
    _batch(sh, [_req_cab(sid, 4, _AZUL), _req_freeze(sid, 1)])
    time.sleep(2)
    print(f"    ✔ Pendentes ({len(df_pend)})")


# ════════════════════════════════════════════════════════════════════════════
# ABA 5 — EXTRATO COMPLETO
# ════════════════════════════════════════════════════════════════════════════

def _escrever_extrato(sh, df_banco):
    ws = _aba(sh, "📄 EXTRATO BANCOS", rows=max(len(df_banco) + 10, 300))
    sid = ws.id

    cab = ["DATA", "BANCO", "DESCRIÇÃO", "CRÉDITO (R$)", "DÉBITO (R$)", "SALDO (R$)", "DOCUMENTO"]
    dados = [cab]

    for _, r in df_banco.sort_values(["data", "banco"]).iterrows():
        dados.append([
            _fmt_data(r["data"]),
            r["banco"],
            r["descricao"],
            _fmt_r(r["credito"]) if r["credito"] else "",
            _fmt_r(r["debito"])  if r["debito"]  else "",
            _fmt_r(r["saldo"])   if r["saldo"]   else "",
            r.get("documento", ""),
        ])

    _escrever(ws, dados)
    _batch(sh, [_req_cab(sid, 7), _req_freeze(sid, 1)])
    time.sleep(2)
    print(f"    ✔ Extrato completo ({len(df_banco)} lançamentos)")


# ════════════════════════════════════════════════════════════════════════════
# FUNÇÃO PRINCIPAL
# ════════════════════════════════════════════════════════════════════════════

def publicar(
    df_banco:    pd.DataFrame,
    df_previsto: pd.DataFrame,
    df_conc:     pd.DataFrame,
    df_nao_prev: pd.DataFrame,
    df_pend:     pd.DataFrame,
    periodo:     str = "",
) -> str:
    client = _conectar()
    sh = _abrir(client)

    print("  → Publicando no Google Sheets...")
    _escrever_resumo(sh, df_conc, df_nao_prev, df_banco, periodo)
    _escrever_nao_previstos(sh, df_nao_prev)
    _escrever_conciliacao(sh, df_conc)
    _escrever_pendentes(sh, df_pend)
    _escrever_extrato(sh, df_banco)

    return f"https://docs.google.com/spreadsheets/d/{sh.id}"
