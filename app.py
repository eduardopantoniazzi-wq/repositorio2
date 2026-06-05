"""
Controle Financeiro Anti-Desvio — Moinho de Trigo
App web interativo: compara previstos (planilha FC) x efetivados (extratos)
"""

import io
import sys
import tempfile
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

# Garante que src/ seja encontrado independente de onde o app roda
ROOT = Path(__file__).parent
sys.path.insert(0, str(ROOT))

st.set_page_config(
    page_title="Controle Financeiro — Moinho de Trigo",
    page_icon="🌾",
    layout="wide",
)

_MESES = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho",
          "Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"]


def _detectar_banco(nome: str) -> str:
    n = nome.lower()
    if "bradesco" in n:
        return "Bradesco"
    if "sicredi" in n:
        return "Sicredi"
    if "banrisul" in n:
        return "Banrisul"
    if n.startswith("bb") or "banco_brasil" in n or "bancobrasil" in n or "_bb" in n:
        return "BB"
    return None


def _eh_planilha(nome: str) -> bool:
    n = nome.lower()
    return any(k in n for k in ("receita","despesa","fluxo","caixa","fc")) and n.endswith((".xlsx",".xls"))


def processar_arquivos(extratos_up, planilha_up, meses_sel):
    """Lê todos os arquivos e retorna (df_banco, df_previsto, erros)."""
    from src.readers.bradesco import LeitorBradesco
    from src.readers.sicredi import LeitorSicredi
    from src.readers.bb import LeitorBB
    from src.readers.banrisul import LeitorBanrisul
    from src.readers.planilha import ler_planilha

    LEITORES = {"Bradesco": LeitorBradesco, "Sicredi": LeitorSicredi,
                "BB": LeitorBB, "Banrisul": LeitorBanrisul}

    dfs_banco = []
    erros = []

    for f in extratos_up:
        banco = _detectar_banco(f.name)
        if banco is None:
            erros.append(f"⚠ '{f.name}': não reconhecido como Bradesco, Sicredi ou BB")
            continue
        suffix = Path(f.name).suffix
        with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
            tmp.write(f.read())
            tmp_path = Path(tmp.name)
        try:
            leitor = LEITORES[banco](tmp_path)
            df = leitor.ler()
            df["banco"] = banco
            dfs_banco.append(df)
            st.sidebar.success(f"✔ {f.name} — {len(df)} lançamentos")
        except Exception as e:
            erros.append(f"⚠ '{f.name}': {e}")
        finally:
            tmp_path.unlink(missing_ok=True)

    df_banco = pd.concat(dfs_banco, ignore_index=True) if dfs_banco else pd.DataFrame()

    df_previsto = None
    if planilha_up:
        suffix = Path(planilha_up.name).suffix
        with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
            tmp.write(planilha_up.read())
            tmp_path = Path(tmp.name)
        try:
            df_previsto = ler_planilha(tmp_path, meses=meses_sel or None)
            st.sidebar.success(f"✔ {planilha_up.name} — {len(df_previsto)} lançamentos previstos")
        except Exception as e:
            erros.append(f"⚠ Planilha '{planilha_up.name}': {e}")
        finally:
            tmp_path.unlink(missing_ok=True)

    return df_banco, df_previsto, erros


def colorir_linha(row):
    s = str(row.get("Status", ""))
    if s == "CONCILIADO":
        bg = "background-color:#d4edda;color:#155724"
    elif "DIVERG" in s:
        bg = "background-color:#fff3cd;color:#856404"
    elif "NÃO PREVISTO" in s:
        bg = "background-color:#f8d7da;color:#721c24"
    elif s == "PENDENTE":
        bg = "background-color:#e2e3e5;color:#383d41"
    else:
        bg = ""
    return [bg] * len(row)


def fmt_brl(v):
    try:
        return f"R$ {float(v):,.2f}"
    except Exception:
        return ""


def fmt_data(v):
    try:
        return pd.Timestamp(v).strftime("%d/%m/%Y")
    except Exception:
        return ""


# ── Sidebar ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("📂 Arquivos")
    st.caption("Nomeie os extratos com o nome do banco: bradesco*.pdf, sicredi*.pdf, bb*.pdf")

    extratos_up = st.file_uploader(
        "Extratos bancários (PDF, XLSX, CSV)",
        type=["pdf","xlsx","xls","csv"],
        accept_multiple_files=True,
    )
    planilha_up = st.file_uploader(
        "Planilha FC (receitas_despesas*.xlsx)",
        type=["xlsx","xls"],
    )

    st.subheader("Período")
    mes_atual = _MESES[datetime.now().month - 1]
    mes_ant   = _MESES[datetime.now().month - 2] if datetime.now().month > 1 else "Dezembro"
    meses_sel = st.multiselect("Meses", _MESES, default=[mes_ant, mes_atual])

    janela = st.slider("Janela de datas (dias)", 0, 7, 3)
    rodar  = st.button("▶ Analisar", type="primary", use_container_width=True)

# ── Main ─────────────────────────────────────────────────────────────────────
st.title("🌾 Controle Financeiro Anti-Desvio")
st.caption("Moinho de Trigo — Previstos (planilha FC) × Efetivados (extratos bancários)")

if not rodar:
    st.info("👈 Faça upload dos arquivos na barra lateral e clique em **▶ Analisar**.")
    st.markdown("""
| Cor | Significado |
|-----|-------------|
| 🟢 Verde | Previsto e pago — valores batem |
| 🟡 Amarelo | Pago, mas valor ou beneficiário diferente do previsto |
| 🔴 Vermelho | **Pago sem estar na planilha** — possível desvio |
| ⬜ Cinza | Previsto mas ainda não foi pago |
    """)
    st.stop()

if not extratos_up:
    st.error("Faça upload de ao menos um extrato bancário.")
    st.stop()

with st.spinner("Processando arquivos..."):
    df_banco, df_previsto, erros = processar_arquivos(extratos_up, planilha_up, meses_sel)

for e in erros:
    st.warning(e)

if df_banco.empty:
    st.error("Nenhum extrato pôde ser lido. Verifique os nomes dos arquivos.")
    st.stop()

# ── Saldos ────────────────────────────────────────────────────────────────────
st.header("💰 Saldos Bancários")
cols = st.columns(4)
total = 0.0
for i, banco in enumerate(["Bradesco","Sicredi","BB"]):
    sub = df_banco[df_banco["banco"] == banco]["saldo"]
    saldo = float(sub.iloc[-1]) if not sub.empty else 0.0
    total += saldo
    cols[i].metric(banco, fmt_brl(saldo))
cols[3].metric("TOTAL", fmt_brl(total))

# ── Conciliação ───────────────────────────────────────────────────────────────
if df_previsto is None or df_previsto.empty:
    st.info("Faça upload da planilha FC para ver a conciliação.")
    st.subheader("📄 Extrato Consolidado")
    st.dataframe(df_banco[["data","banco","descricao","credito","debito","saldo"]], use_container_width=True, height=500)
    st.stop()

with st.spinner("Conciliando previsto × efetivado..."):
    from src.reconciler import conciliar
    df_conc, df_nao_prev, df_pend = conciliar(df_previsto, df_banco, janela_dias=janela)

# ── Métricas ──────────────────────────────────────────────────────────────────
st.header("📊 Resumo")
n_conc = (df_conc["status"] == "CONCILIADO").sum()
n_div  = df_conc["status"].str.startswith("DIVERGENCIA").sum()
n_pend = (df_conc["status"] == "PENDENTE").sum()
n_nao  = len(df_nao_prev)
val_nao = float(df_nao_prev["valor"].sum()) if not df_nao_prev.empty else 0.0
val_div = float(df_conc[df_conc["status"].str.startswith("DIVERGENCIA")]["diferenca"].abs().sum())

c1,c2,c3,c4 = st.columns(4)
c1.metric("✅ Conciliados", n_conc)
c2.metric("⚠️ Divergências", n_div,
          delta=fmt_brl(val_div) if n_div else None, delta_color="inverse")
c3.metric("🕐 Pendentes", n_pend)
c4.metric("🚨 Não Previstos", n_nao,
          delta=fmt_brl(val_nao) if n_nao else None, delta_color="inverse")

# ── Tabela principal ──────────────────────────────────────────────────────────
st.header("🔍 Tabela de Conciliação")

# Monta df unificado para exibição
rows_conc = []
for _, r in df_conc.iterrows():
    rows_conc.append({
        "Status":                r["status"],
        "Data Prevista":         fmt_data(r["data_prevista"]),
        "Beneficiário (Planilha)": str(r["descricao_planilha"]),
        "Valor Previsto":        fmt_brl(r["valor_previsto"]),
        "Data Pagamento":        fmt_data(r["data_pagamento"]),
        "Banco":                 str(r["banco"]),
        "Beneficiário (Extrato)": str(r["descricao_extrato"]),
        "Valor Pago":            fmt_brl(r["valor_pago"]) if r["valor_pago"] else "",
        "Diferença":             (f'+{fmt_brl(r["diferenca"])}' if r["diferenca"] > 0
                                  else fmt_brl(r["diferenca"]) if r["diferenca"] != 0 else ""),
    })

rows_nao = []
for _, r in df_nao_prev.iterrows():
    risco = str(r.get("risco",""))
    label = ("🔴 ALTO" if risco=="ALTO" else "🟡 MÉDIO" if risco=="MEDIO"
             else "🟢 BAIXO" if risco=="BAIXO" else risco)
    rows_nao.append({
        "Status":                f"NÃO PREVISTO [{label}]",
        "Data Prevista":         "",
        "Beneficiário (Planilha)": "",
        "Valor Previsto":        "",
        "Data Pagamento":        fmt_data(r["data"]),
        "Banco":                 str(r["banco"]),
        "Beneficiário (Extrato)": str(r["descricao"]),
        "Valor Pago":            fmt_brl(r["valor"]),
        "Diferença":             fmt_brl(r["valor"]),
    })

df_view_all = pd.DataFrame(rows_conc + rows_nao)

# Filtros
cf1, cf2, cf3 = st.columns(3)
status_opts = sorted(df_view_all["Status"].unique().tolist())
# Padrão: mostra tudo menos OPERACIONAL e PENDENTE
status_default = [s for s in status_opts if "OPERACIONAL" not in s]
status_sel = cf1.multiselect("Status", status_opts, default=status_default)
banco_opts = ["Todos"] + sorted(df_banco["banco"].unique().tolist())
banco_sel  = cf2.selectbox("Banco", banco_opts)
busca      = cf3.text_input("Buscar beneficiário", placeholder="Ex: Cooperativa...")

df_view = df_view_all[df_view_all["Status"].isin(status_sel)].copy()
if banco_sel != "Todos":
    df_view = df_view[df_view["Banco"] == banco_sel]
if busca:
    mask = (df_view["Beneficiário (Planilha)"].str.contains(busca, case=False, na=False) |
            df_view["Beneficiário (Extrato)"].str.contains(busca, case=False, na=False))
    df_view = df_view[mask]

st.dataframe(
    df_view.style.apply(colorir_linha, axis=1),
    use_container_width=True,
    height=600,
    column_config={
        "Status":                   st.column_config.TextColumn(width=200),
        "Data Prevista":            st.column_config.TextColumn(width=110),
        "Beneficiário (Planilha)":  st.column_config.TextColumn(width=260),
        "Valor Previsto":           st.column_config.TextColumn(width=130),
        "Data Pagamento":           st.column_config.TextColumn(width=110),
        "Banco":                    st.column_config.TextColumn(width=90),
        "Beneficiário (Extrato)":   st.column_config.TextColumn(width=280),
        "Valor Pago":               st.column_config.TextColumn(width=130),
        "Diferença":                st.column_config.TextColumn(width=120),
    },
)

# ── Download Excel ────────────────────────────────────────────────────────────
buf = io.BytesIO()
with pd.ExcelWriter(buf, engine="openpyxl") as writer:
    df_view_all.to_excel(writer, sheet_name="Conciliação Completa", index=False)
    df_nao_prev.to_excel(writer, sheet_name="Não Previstos", index=False)
    df_pend.to_excel(writer, sheet_name="Pendentes", index=False)
    df_banco.to_excel(writer, sheet_name="Extrato Consolidado", index=False)

st.download_button(
    "📥 Baixar relatório Excel",
    data=buf.getvalue(),
    file_name=f"conciliacao_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)
