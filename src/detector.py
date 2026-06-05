"""
Motor de detecção de desvios financeiros.
Retorna lista de alertas com nível CRITICO ou ATENCAO.
"""

from __future__ import annotations
from datetime import datetime
from dateutil.relativedelta import relativedelta
from typing import Literal

import pandas as pd

from config import (
    LIMITE_PAGAMENTO_ALTO,
    LIMITE_VARIACAO_FORNECEDOR,
    HORARIO_INICIO_PERMITIDO,
    HORARIO_FIM_PERMITIDO,
    JANELA_HISTORICO_MESES,
)
from .knowledge_base import carregar_fornecedores, carregar_clientes


Nivel = Literal["CRITICO", "ATENCAO"]


def _alerta(nivel: Nivel, tipo: str, descricao: str, linha: dict) -> dict:
    return {
        "data_alerta": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "nivel": nivel,
        "tipo": tipo,
        "descricao": descricao,
        "banco": linha.get("banco", ""),
        "data_lancamento": str(linha.get("data", ""))[:10],
        "beneficiario": linha.get("descricao", ""),
        "valor": linha.get("debito", linha.get("credito", 0)),
        "documento": linha.get("documento", ""),
    }


def rodar_verificacoes(df: pd.DataFrame) -> tuple[pd.DataFrame, list[dict]]:
    """
    Executa todas as verificações de desvio.
    Retorna o DataFrame com coluna 'alerta' preenchida e lista de alertas.
    """
    df = df.copy()
    df["alerta"] = ""
    alertas: list[dict] = []

    fornecedores = carregar_fornecedores()
    clientes     = carregar_clientes()

    hoje = df["data"].max()
    limite_historico = hoje - relativedelta(months=JANELA_HISTORICO_MESES)
    limite_mes_anterior = hoje - relativedelta(months=1)

    # Histórico interno do próprio extrato (últimos 3 meses)
    df_historico = df[df["data"] < hoje.replace(day=1)]

    for idx, row in df.iterrows():
        linha = row.to_dict()
        alertas_linha = []

        # ── 1. Pagamento alto sem histórico do beneficiário ─────────────────
        if row["debito"] > LIMITE_PAGAMENTO_ALTO:
            nome = str(row["descricao"]).strip()
            teve_historico = False

            # Verifica no extrato atual (mês anterior)
            historico_mes = df_historico[
                (df_historico["descricao"] == nome) &
                (df_historico["data"] >= limite_mes_anterior) &
                (df_historico["debito"] > 0)
            ]
            if not historico_mes.empty:
                teve_historico = True

            # Verifica na base de conhecimento
            if nome in fornecedores:
                meses_forncedor = fornecedores[nome].get("meses", [])
                mes_anterior_str = (hoje - relativedelta(months=1)).strftime("%Y-%m")
                if mes_anterior_str in meses_forncedor:
                    teve_historico = True

            if not teve_historico:
                msg = f"Pagamento de R$ {row['debito']:,.2f} para '{nome}' sem histórico no mês anterior"
                alertas.append(_alerta("CRITICO", "PAGAMENTO_ALTO_SEM_HISTORICO", msg, linha))
                alertas_linha.append("CRÍTICO: PAGAMENTO_ALTO_SEM_HISTORICO")

        # ── 2. Mesmo beneficiário recebendo mais de um pagamento no mesmo dia ─
        if row["debito"] > 0:
            duplicados = df[
                (df["descricao"] == row["descricao"]) &
                (df["data"] == row["data"]) &
                (df["debito"] > 0) &
                (df.index != idx)
            ]
            if not duplicados.empty:
                msg = (
                    f"'{row['descricao']}' recebeu {len(duplicados) + 1} pagamentos "
                    f"no mesmo dia {str(row['data'])[:10]}"
                )
                alertas.append(_alerta("ATENCAO", "MULTIPLOS_PAGAMENTOS_DIA", msg, linha))
                alertas_linha.append("ATENÇÃO: MULTIPLOS_PAGAMENTOS_DIA")

        # ── 3. Beneficiário novo (nunca apareceu nos últimos 3 meses) ────────
        if row["debito"] > 0:
            nome = str(row["descricao"]).strip()
            apareceu_antes = False

            hist3m = df_historico[
                (df_historico["descricao"] == nome) &
                (df_historico["data"] >= limite_historico) &
                (df_historico["debito"] > 0)
            ]
            if not hist3m.empty:
                apareceu_antes = True

            if nome in fornecedores:
                apareceu_antes = True

            if not apareceu_antes:
                msg = f"Beneficiário novo: '{nome}' nunca apareceu nos últimos {JANELA_HISTORICO_MESES} meses"
                alertas.append(_alerta("ATENCAO", "BENEFICIARIO_NOVO", msg, linha))
                alertas_linha.append("ATENÇÃO: BENEFICIARIO_NOVO")

        # ── 4. Pagamento em horário fora do padrão ───────────────────────────
        if row.get("horario") and pd.notna(row["horario"]):
            try:
                hora = int(str(row["horario"]).split(":")[0])
                if hora < HORARIO_INICIO_PERMITIDO or hora >= HORARIO_FIM_PERMITIDO:
                    msg = (
                        f"Pagamento às {row['horario']} para '{row['descricao']}' "
                        f"(fora do horário {HORARIO_INICIO_PERMITIDO}h–{HORARIO_FIM_PERMITIDO}h)"
                    )
                    alertas.append(_alerta("CRITICO", "HORARIO_FORA_PADRAO", msg, linha))
                    alertas_linha.append("CRÍTICO: HORARIO_FORA_PADRAO")
            except (ValueError, TypeError):
                pass

        # ── 5. Variação acima de 30% em fornecedor recorrente ────────────────
        if row["debito"] > 0:
            nome = str(row["descricao"]).strip()
            historico_vals = []

            # Valores históricos do extrato
            hist_vals_df = df_historico[
                (df_historico["descricao"] == nome) &
                (df_historico["debito"] > 0)
            ]["debito"].tolist()
            historico_vals.extend(hist_vals_df)

            # Valores da base de conhecimento
            if nome in fornecedores:
                historico_vals.extend(fornecedores[nome].get("valores", []))

            if len(historico_vals) >= 2:
                media = sum(historico_vals) / len(historico_vals)
                if media > 0:
                    variacao = abs(row["debito"] - media) / media
                    if variacao > LIMITE_VARIACAO_FORNECEDOR:
                        msg = (
                            f"Variação de {variacao:.0%} no pagamento para '{nome}': "
                            f"valor atual R$ {row['debito']:,.2f}, média histórica R$ {media:,.2f}"
                        )
                        alertas.append(_alerta("ATENCAO", "VARIACAO_FORNECEDOR", msg, linha))
                        alertas_linha.append("ATENÇÃO: VARIACAO_FORNECEDOR")

        # ── 6. Crédito de CNPJ desconhecido ──────────────────────────────────
        if row["credito"] > 0:
            nome = str(row["descricao"]).strip()
            e_conhecido = nome in clientes

            # Verifica histórico interno
            if not e_conhecido:
                hist_cred = df_historico[
                    (df_historico["descricao"] == nome) &
                    (df_historico["credito"] > 0)
                ]
                if not hist_cred.empty:
                    e_conhecido = True

            if not e_conhecido:
                msg = f"Crédito de R$ {row['credito']:,.2f} de origem desconhecida: '{nome}'"
                alertas.append(_alerta("ATENCAO", "CREDITO_DESCONHECIDO", msg, linha))
                alertas_linha.append("ATENÇÃO: CREDITO_DESCONHECIDO")

        # Marca a coluna no DataFrame
        if alertas_linha:
            df.at[idx, "alerta"] = " | ".join(alertas_linha)

    return df, alertas
