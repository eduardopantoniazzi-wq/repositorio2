"""
Motor de conciliação: cruza pagamentos PREVISTOS (planilha Excel)
com pagamentos EFETIVADOS (extratos bancários).

Resultado por lançamento:
  CONCILIADO    → previsto e pago, valores batem
  DIVERGENCIA   → previsto e pago, mas valor diferente ou beneficiário suspeito
  NAO_PREVISTO  → pago no banco mas NÃO estava na planilha  ← risco de desvio
  PENDENTE      → estava na planilha mas NÃO foi pago ainda
"""

from __future__ import annotations
import re
from difflib import SequenceMatcher
from typing import Optional

import pandas as pd


# ── Normalização de nomes para comparação ───────────────────────────────────

_STOPWORDS = {
    "nf", "nota", "fiscal", "ltda", "s/a", "sa", "eireli", "me", "epp",
    "pagamento", "de", "boleto", "pag", "pagtf", "pgto", "pagto",
    "transferencia", "transf", "ted", "pix", "enviado", "recebido",
    "ctr", "cta", "nr", "num", "contrato", "parcela",
}


def _normalizar(nome: str) -> str:
    """Extrai palavras significativas do nome para comparação."""
    s = str(nome).upper()
    # Remove números de nota fiscal e documento
    s = re.sub(r"\b(NF|CTR|NR)\s*[:\-]?\s*[\d./]+", "", s)
    # Remove caracteres especiais
    s = re.sub(r"[^\w\s]", " ", s)
    palavras = [p for p in s.split() if len(p) > 2 and p.lower() not in _STOPWORDS]
    return " ".join(palavras[:4])  # primeiras 4 palavras significativas


def _similaridade(a: str, b: str) -> float:
    na, nb = _normalizar(a), _normalizar(b)
    if not na or not nb:
        return 0.0
    # Verifica se qualquer palavra de `a` aparece em `b`
    palavras_a = set(na.split())
    palavras_b = set(nb.split())
    if palavras_a & palavras_b:  # intersecção não vazia
        return max(0.6, SequenceMatcher(None, na, nb).ratio())
    return SequenceMatcher(None, na, nb).ratio()


# ── Conciliador ──────────────────────────────────────────────────────────────

TOLERANCIA_VALOR = 0.02  # 2% de tolerância no valor


def conciliar(
    df_previsto: pd.DataFrame,   # da planilha Excel
    df_extrato:  pd.DataFrame,   # dos extratos bancários (só débitos)
    janela_dias: int = 3,        # quantos dias de diferença aceitar na data
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Retorna três DataFrames:
      1. df_conciliacao: todos os lançamentos previstos com status e match
      2. df_nao_previstos: débitos bancários sem correspondência na planilha
      3. df_pendentes: previstos que não apareceram no extrato
    """
    # Só débitos reais (não tarifas mínimas)
    debitos = df_extrato[df_extrato["debito"] > 0].copy().reset_index(drop=True)
    previstos = df_previsto[df_previsto["debito"] > 0].copy().reset_index(drop=True)

    # Marca de uso (evita casar o mesmo lançamento duas vezes)
    debitos["_usado"] = False
    previstos["_usado"] = False

    resultados = []

    # ── Passo 1: tenta casar cada previsto com um débito real ────────────
    for ip, prev in previstos.iterrows():
        melhor_idx: Optional[int] = None
        melhor_score = 0.0

        candidatos = debitos[
            (~debitos["_usado"]) &
            (debitos["debito"] > 0)
        ]

        for id_, deb in candidatos.iterrows():
            # Diferença de valor
            if prev["debito"] == 0:
                continue
            diff_val = abs(deb["debito"] - prev["debito"]) / prev["debito"]
            if diff_val > 0.50:  # ignora se diferença > 50%
                continue

            # Diferença de data
            try:
                diff_dias = abs((deb["data"] - prev["data"]).days)
            except Exception:
                diff_dias = 999

            if diff_dias > janela_dias:
                continue

            # Score combinado: valor pesa mais
            score_val  = max(0.0, 1.0 - diff_val / 0.50)
            score_data = max(0.0, 1.0 - diff_dias / (janela_dias + 1))
            score_nome = _similaridade(prev["descricao"], deb["descricao"])
            score = score_val * 0.55 + score_nome * 0.30 + score_data * 0.15

            if score > melhor_score:
                melhor_score = score
                melhor_idx = id_

        if melhor_idx is not None and melhor_score >= 0.40:
            deb = debitos.loc[melhor_idx]
            debitos.at[melhor_idx, "_usado"] = True
            previstos.at[ip, "_usado"] = True

            diff_val = abs(deb["debito"] - prev["debito"]) / prev["debito"]
            sim_nome = _similaridade(prev["descricao"], deb["descricao"])

            if diff_val <= TOLERANCIA_VALOR and sim_nome >= 0.4:
                status = "CONCILIADO"
            elif diff_val > TOLERANCIA_VALOR and sim_nome >= 0.4:
                status = "DIVERGENCIA_VALOR"
            elif diff_val <= TOLERANCIA_VALOR and sim_nome < 0.4:
                status = "DIVERGENCIA_BENEFICIARIO"
            else:
                status = "DIVERGENCIA"

            resultados.append({
                "data_prevista":       prev["data"],
                "unidade":             prev.get("unidade", ""),
                "descricao_planilha":  prev["descricao"],
                "valor_previsto":      round(prev["debito"], 2),
                "data_pagamento":      deb["data"],
                "banco":               deb["banco"],
                "descricao_extrato":   deb["descricao"],
                "valor_pago":          round(deb["debito"], 2),
                "diferenca":           round(deb["debito"] - prev["debito"], 2),
                "similaridade_nome":   round(sim_nome, 2),
                "status":              status,
            })
        else:
            resultados.append({
                "data_prevista":       prev["data"],
                "unidade":             prev.get("unidade", ""),
                "descricao_planilha":  prev["descricao"],
                "valor_previsto":      round(prev["debito"], 2),
                "data_pagamento":      None,
                "banco":               "",
                "descricao_extrato":   "",
                "valor_pago":          0.0,
                "diferenca":           -round(prev["debito"], 2),
                "similaridade_nome":   0.0,
                "status":              "PENDENTE",
            })

    df_conciliacao = pd.DataFrame(resultados)

    # ── Passo 2: débitos não casados = suspeitos ─────────────────────────
    nao_previstos_rows = []
    for _, deb in debitos[~debitos["_usado"]].iterrows():
        nao_previstos_rows.append({
            "data":      deb["data"],
            "banco":     deb["banco"],
            "descricao": deb["descricao"],
            "valor":     round(deb["debito"], 2),
            "documento": deb.get("documento", ""),
            "risco":     _classificar_risco(deb),
        })

    df_nao_previstos = pd.DataFrame(nao_previstos_rows)

    # ── Passo 3: pendentes ────────────────────────────────────────────────
    pendentes = df_conciliacao[df_conciliacao["status"] == "PENDENTE"].copy()

    return df_conciliacao, df_nao_previstos, pendentes


def _classificar_risco(deb: pd.Series) -> str:
    """Classifica risco de um pagamento não previsto."""
    desc = str(deb.get("descricao", "")).upper()
    val  = float(deb.get("debito", 0))

    # Itens operacionais normais — baixo risco
    baixo_risco = [
        "TARIFA", "IOF", "SALDO INVEST", "RENDE FACIL", "FACILCRED",
        "RENTAB", "JUROS", "TAXA", "DEBITO SERV", "CUSTAS DE PROTESTO",
        "TARIFA DE PROTESTO", "TAR ", "COBRANCA", "DÉBITO AUTOMÁTICO",
        "FOLHA", "SALARIO", "FGTS", "INSS", "IMPOSTO",
    ]
    for p in baixo_risco:
        if p in desc:
            return "OPERACIONAL"

    # Transferências para pessoas físicas → alto risco
    if val >= 1000 and any(p in desc for p in ["PIX ENVIADO", "TRANSF", "TED"]):
        return "ALTO"

    if val >= 5000:
        return "ALTO"
    elif val >= 500:
        return "MEDIO"
    else:
        return "BAIXO"
