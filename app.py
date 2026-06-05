"""
Controle Anti-Desvio — Moinho de Trigo
Tabela: Débitos Previstos (planilha FC) × Débitos Efetivados (extratos)
"""

import io, sys, tempfile
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

sys.path.insert(0, str(Path(__file__).parent))

st.set_page_config(page_title="Controle Anti-Desvio — Moinho de Trigo",
                   page_icon="🌾", layout="wide")

_MESES = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho",
          "Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"]

BANCO_POR_NOME = {
    "bradesco": "Bradesco", "sicredi": "Sicredi",
    "banrisul": "Banrisul", "bb": "BB",
}

def detectar_banco(nome):
    n = nome.lower()
    for k, v in BANCO_POR_NOME.items():
        if k in n:
            return v
    return None

def ler_extrato(nome, conteudo):
    from src.readers.bradesco import LeitorBradesco
    from src.readers.sicredi  import LeitorSicredi
    from src.readers.bb       import LeitorBB
    from src.readers.banrisul import LeitorBanrisul
    LEITORES = {"Bradesco": LeitorBradesco, "Sicredi": LeitorSicredi,
                "BB": LeitorBB, "Banrisul": LeitorBanrisul}
    banco = detectar_banco(nome)
    if not banco:
        return None, f"'{nome}' não reconhecido — renomeie com bradesco/sicredi/bb/banrisul no início"
    suffix = Path(nome).suffix
    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
        tmp.write(conteudo); tmp_path = Path(tmp.name)
    try:
        df = LEITORES[banco](tmp_path).ler()
        df["banco"] = banco
        return df, None
    except Exception as e:
        return None, f"Erro ao ler '{nome}': {e}"
    finally:
        tmp_path.unlink(missing_ok=True)

def ler_planilha(conteudo, meses):
    from src.readers.planilha import ler_planilha as _ler
    suffix = ".xlsx"
    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
        tmp.write(conteudo); tmp_path = Path(tmp.name)
    try:
        return _ler(tmp_path, meses=meses), None
    except Exception as e:
        return None, f"Erro ao ler planilha: {e}"
    finally:
        tmp_path.unlink(missing_ok=True)

def conciliar(df_prev, df_banco):
    """
    Para cada débito PREVISTO, busca o débito bancário mais próximo
    (mesmo valor ±10%, mesma semana, nome parecido).
    Retorna DataFrame pronto para exibição.
    """
    from difflib import SequenceMatcher
    import re
    STOP = {"nf","nota","fiscal","ltda","sa","eireli","me","epp","pag","pgto","ted","pix","de","boleto"}

    def norm(s):
        s = re.sub(r"[^\w\s]", " ", str(s).upper())
        return " ".join(w for w in s.split() if len(w)>2 and w.lower() not in STOP)

    def sim(a, b):
        na, nb = norm(a), norm(b)
        if not na or not nb: return 0.0
        base = SequenceMatcher(None, na, nb).ratio()
        return max(base, 0.6) if set(na.split()) & set(nb.split()) else base

    deb_banco = df_banco[df_banco["debito"] > 0].copy().reset_index(drop=True)
    deb_prev  = df_prev[df_prev["debito"] > 0].copy().reset_index(drop=True)
    deb_banco["_usado"] = False

    linhas = []

    # ── Cruza cada previsto com o melhor candidato bancário ──────────────
    for _, prev in deb_prev.iterrows():
        melhor_idx, melhor_score = None, 0.0
        for idx, deb in deb_banco[~deb_banco["_usado"]].iterrows():
            diff_val = abs(deb["debito"] - prev["debito"]) / max(prev["debito"], 1)
            if diff_val > 0.50: continue
            try:
                diff_dias = abs((deb["data"] - prev["data"]).days)
            except Exception:
                diff_dias = 999
            if diff_dias > 5: continue
            score = (max(0, 1-diff_val/0.5)*0.60 +
                     sim(prev["descricao"], deb["descricao"])*0.25 +
                     max(0, 1-diff_dias/6)*0.15)
            if score > melhor_score:
                melhor_score, melhor_idx = score, idx

        if melhor_idx is not None and melhor_score >= 0.35:
            deb = deb_banco.loc[melhor_idx]
            deb_banco.at[melhor_idx, "_usado"] = True
            diff = round(deb["debito"] - prev["debito"], 2)
            pct  = diff / prev["debito"] * 100 if prev["debito"] else 0
            if abs(diff) <= prev["debito"] * 0.02:
                status = "✅ OK"
            else:
                status = f"⚠️ DIVERGÊNCIA"
            linhas.append({
                "Status":              status,
                "Data Prevista":       prev["data"],
                "Beneficiário Previsto": prev["descricao"],
                "Valor Previsto (R$)": prev["debito"],
                "Data Pago":           deb["data"],
                "Banco":               deb["banco"],
                "Pago Para":           deb["descricao"],
                "Valor Pago (R$)":     deb["debito"],
                "Diferença (R$)":      diff,
                "Diferença (%)":       round(pct, 1),
            })
        else:
            linhas.append({
                "Status":              "🕐 NÃO PAGO",
                "Data Prevista":       prev["data"],
                "Beneficiário Previsto": prev["descricao"],
                "Valor Previsto (R$)": prev["debito"],
                "Data Pago":           None,
                "Banco":               "",
                "Pago Para":           "",
                "Valor Pago (R$)":     None,
                "Diferença (R$)":      -prev["debito"],
                "Diferença (%)":       -100.0,
            })

    # ── Débitos bancários sem correspondência ────────────────────────────
    for _, deb in deb_banco[~deb_banco["_usado"]].iterrows():
        desc = str(deb["descricao"]).upper()
        operacional = any(p in desc for p in ["TARIFA","IOF","JUROS","TAXA","INSS","FGTS"])
        if operacional:
            continue   # ignora tarifas e impostos automáticos
        linhas.append({
            "Status":              "🚨 NÃO PREVISTO",
            "Data Prevista":       None,
            "Beneficiário Previsto": "",
            "Valor Previsto (R$)": None,
            "Data Pago":           deb["data"],
            "Banco":               deb["banco"],
            "Pago Para":           deb["descricao"],
            "Valor Pago (R$)":     deb["debito"],
            "Diferença (R$)":      deb["debito"],
            "Diferença (%)":       None,
        })

    df = pd.DataFrame(linhas)
    return df


def cor_linha(row):
    s = str(row.get("Status",""))
    if s.startswith("✅"):   bg = "#d4edda; color:#155724"
    elif s.startswith("⚠️"): bg = "#fff3cd; color:#856404"
    elif s.startswith("🚨"): bg = "#f8d7da; color:#721c24"
    else:                    bg = "#e2e3e5; color:#383d41"
    return [f"background-color:{bg}"] * len(row)


def fmt_brl(v):
    try:    return f"R$ {float(v):,.2f}"
    except: return ""

def fmt_dt(v):
    try:    return pd.Timestamp(v).strftime("%d/%m/%Y")
    except: return ""

# ─────────────────────────────────────────────────────────────────────────────
st.title("🌾 Débitos Previstos × Efetivados")
st.caption("Moinho de Trigo — compara o que estava previsto na planilha com o que saiu dos bancos")

# Sidebar
with st.sidebar:
    st.header("📂 Arquivos")
    extratos_up = st.file_uploader("Extratos bancários (PDF, XLSX, CSV)",
        type=["pdf","xlsx","xls","csv"], accept_multiple_files=True)
    planilha_up = st.file_uploader("Planilha de previsões (FC)",
        type=["xlsx","xls"])
    st.caption("Nomeie: bradesco_*.pdf · sicredi_*.pdf · bb_*.pdf · banrisul_*.pdf")
    st.divider()
    mes_atual = _MESES[datetime.now().month-1]
    mes_ant   = _MESES[datetime.now().month-2] if datetime.now().month>1 else "Dezembro"
    meses_sel = st.multiselect("Meses", _MESES, default=[mes_ant, mes_atual])
    rodar = st.button("▶ Comparar", type="primary", use_container_width=True)

if not rodar:
    st.markdown("""
### Como funciona
1. Faça upload dos **extratos bancários** do dia (PDF ou Excel)
2. Faça upload da **planilha de previsões** (FC)
3. Selecione o **mês** e clique em **▶ Comparar**

A tabela mostra linha por linha:

| Cor | O que significa |
|---|---|
| 🟢 Verde | Pago conforme previsto |
| 🟡 Amarelo | Pago, mas valor diferente do previsto |
| 🔴 Vermelho | **Saiu do banco sem estar na planilha** |
| ⬜ Cinza | Previsto mas ainda não pago |
""")
    st.stop()

if not extratos_up:
    st.error("Faça upload dos extratos bancários.")
    st.stop()
if not planilha_up:
    st.error("Faça upload da planilha de previsões.")
    st.stop()

# Lê extratos
with st.spinner("Lendo extratos..."):
    dfs = []
    for f in extratos_up:
        df, err = ler_extrato(f.name, f.read())
        if err:
            st.warning(err)
        else:
            dfs.append(df)
            st.sidebar.success(f"✔ {f.name} ({len(df)} lançamentos)")

if not dfs:
    st.error("Nenhum extrato lido. Verifique os nomes dos arquivos.")
    st.stop()

df_banco = pd.concat(dfs, ignore_index=True)

# Lê planilha
with st.spinner("Lendo planilha de previsões..."):
    df_prev, err = ler_planilha(planilha_up.read(), meses=meses_sel or None)
    if err:
        st.error(err); st.stop()
    st.sidebar.success(f"✔ {planilha_up.name} ({len(df_prev)} lançamentos)")

# Saldos
st.subheader("💰 Saldos dos Bancos")
cols = st.columns(5)
total = 0.0
bancos = ["Bradesco","Sicredi","BB","Banrisul"]
for i, b in enumerate(bancos):
    sub = df_banco[df_banco["banco"]==b]["saldo"]
    s = float(sub.iloc[-1]) if not sub.empty else 0.0
    total += s
    cols[i].metric(b, fmt_brl(s))
cols[4].metric("**TOTAL**", fmt_brl(total))

# Concilia
with st.spinner("Comparando previstos × efetivados..."):
    df_res = conciliar(df_prev, df_banco)

# Métricas
st.subheader("📊 Resumo")
n_ok   = df_res["Status"].str.startswith("✅").sum()
n_div  = df_res["Status"].str.startswith("⚠️").sum()
n_nao  = df_res["Status"].str.startswith("🚨").sum()
n_pend = df_res["Status"].str.startswith("🕐").sum()
v_div  = df_res[df_res["Status"].str.startswith("⚠️")]["Diferença (R$)"].abs().sum()
v_nao  = df_res[df_res["Status"].str.startswith("🚨")]["Valor Pago (R$)"].sum()

c1,c2,c3,c4 = st.columns(4)
c1.metric("✅ Conforme previsto", n_ok)
c2.metric("⚠️ Valor diferente",  n_div, delta=fmt_brl(v_div) if n_div else None, delta_color="inverse")
c3.metric("🚨 Fora da planilha", n_nao, delta=fmt_brl(v_nao) if n_nao else None, delta_color="inverse")
c4.metric("🕐 Ainda não pago",   n_pend)

# Filtros
st.subheader("📋 Tabela Comparativa")
cf1, cf2, cf3 = st.columns(3)
status_opts = sorted(df_res["Status"].unique())
status_sel  = cf1.multiselect("Filtrar por status", status_opts, default=status_opts)
banco_opts  = ["Todos"] + sorted(df_banco["banco"].unique())
banco_sel   = cf2.selectbox("Banco", banco_opts)
busca       = cf3.text_input("Buscar nome", placeholder="ex: Cooperativa, Fornecedor...")

df_view = df_res[df_res["Status"].isin(status_sel)].copy()
if banco_sel != "Todos":
    df_view = df_view[df_view["Banco"] == banco_sel]
if busca:
    mask = (df_view["Beneficiário Previsto"].str.contains(busca, case=False, na=False) |
            df_view["Pago Para"].str.contains(busca, case=False, na=False))
    df_view = df_view[mask]

# Formata para exibição
df_show = df_view.copy()
df_show["Data Prevista"]  = df_show["Data Prevista"].apply(fmt_dt)
df_show["Data Pago"]      = df_show["Data Pago"].apply(fmt_dt)
df_show["Valor Previsto (R$)"] = df_show["Valor Previsto (R$)"].apply(fmt_brl)
df_show["Valor Pago (R$)"]     = df_show["Valor Pago (R$)"].apply(fmt_brl)
df_show["Diferença (R$)"] = df_show["Diferença (R$)"].apply(
    lambda v: f"+{fmt_brl(v)}" if isinstance(v,(int,float)) and v>0 else
              (fmt_brl(v) if isinstance(v,(int,float)) else ""))
df_show["Diferença (%)"]  = df_show["Diferença (%)"].apply(
    lambda v: f"{v:+.1f}%" if isinstance(v,(int,float)) else "")

st.dataframe(
    df_show.style.apply(cor_linha, axis=1),
    use_container_width=True,
    height=620,
    column_config={
        "Status":                  st.column_config.TextColumn(width=160),
        "Data Prevista":           st.column_config.TextColumn("Data Prev.", width=100),
        "Beneficiário Previsto":   st.column_config.TextColumn("Previsto Para", width=240),
        "Valor Previsto (R$)":     st.column_config.TextColumn("Vlr Previsto", width=130),
        "Data Pago":               st.column_config.TextColumn("Data Pago", width=100),
        "Banco":                   st.column_config.TextColumn(width=90),
        "Pago Para":               st.column_config.TextColumn(width=250),
        "Valor Pago (R$)":         st.column_config.TextColumn("Vlr Pago", width=130),
        "Diferença (R$)":          st.column_config.TextColumn("Diferença R$", width=120),
        "Diferença (%)":           st.column_config.TextColumn("Diferença %", width=100),
    },
)

# Download
buf = io.BytesIO()
with pd.ExcelWriter(buf, engine="openpyxl") as w:
    df_view.to_excel(w, sheet_name="Comparativo", index=False)
    df_res[df_res["Status"].str.startswith("🚨")].to_excel(w, sheet_name="Fora da Planilha", index=False)
    df_res[df_res["Status"].str.startswith("⚠️")].to_excel(w, sheet_name="Divergências de Valor", index=False)
    df_banco.to_excel(w, sheet_name="Extrato Consolidado", index=False)

st.download_button("📥 Baixar Excel",
    data=buf.getvalue(),
    file_name=f"comparativo_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True)
