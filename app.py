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
_MESES_NUM = {m: f"{i+1:02d}" for i, m in enumerate(_MESES)}

BANCO_POR_NOME = {
    "bradesco": "Bradesco", "sicredi": "Sicredi",
    "banrisul": "Banrisul",
    "bb_alimentos": "BB Alimentos", "bbalimentos": "BB Alimentos",
    "bb": "BB",
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
                "BB": LeitorBB, "BB Alimentos": LeitorBB, "Banrisul": LeitorBanrisul}
    banco = detectar_banco(nome)
    if not banco:
        return None, f"'{nome}' não reconhecido — renomeie com bradesco/sicredi/bb/bb_alimentos/banrisul no início"
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


def _mes_nome(numero: int) -> str:
    return _MESES[numero - 1]

def conciliar(df_prev, df_banco, limite_alerta: float = 1_500.0):
    """
    Casamento global: monta a matriz de scores de todos os pares possíveis
    e atribui o melhor par disponível globalmente (não guloso por linha).
    Isso evita que dois previstos com valor parecido troquem de beneficiário.
    """
    from difflib import SequenceMatcher
    import re

    STOP = {"nf","nota","fiscal","ltda","sa","eireli","me","epp","pag","pgto",
            "ted","pix","de","boleto","pagamento","transferencia","transf"}

    def norm(s):
        s = re.sub(r"[^\w\s]", " ", str(s).upper())
        return " ".join(w for w in s.split() if len(w) > 2 and w.lower() not in STOP)

    def sim_nome(a, b):
        na, nb = norm(a), norm(b)
        if not na or not nb:
            return 0.0
        palavras_a = set(na.split())
        palavras_b = set(nb.split())
        seq = SequenceMatcher(None, na, nb).ratio()
        # Qualquer palavra em comum já eleva bastante o score
        comuns = palavras_a & palavras_b
        if comuns:
            return max(seq, 0.55 + 0.1 * min(len(comuns), 3))
        return seq

    deb_banco = df_banco[df_banco["debito"] > 0].copy().reset_index(drop=True)
    deb_prev  = df_prev[df_prev["debito"] > 0].copy().reset_index(drop=True)

    nP = len(deb_prev)
    nB = len(deb_banco)

    # ── Monta matriz de scores (nP × nB) ────────────────────────────────
    scores = {}   # (ip, ib) -> score
    for ip, prev in deb_prev.iterrows():
        for ib, deb in deb_banco.iterrows():
            diff_val = abs(deb["debito"] - prev["debito"]) / max(prev["debito"], 1)
            if diff_val > 0.60:
                continue
            try:
                diff_dias = abs((deb["data"] - prev["data"]).days)
            except Exception:
                diff_dias = 999
            if diff_dias > 5:
                continue

            s_val  = max(0.0, 1 - diff_val / 0.60)
            s_nome = sim_nome(prev["descricao"], deb["descricao"])
            s_data = max(0.0, 1 - diff_dias / 6)

            # Nome tem peso maior que valor — evita trocar beneficiários
            score = s_nome * 0.50 + s_val * 0.35 + s_data * 0.15

            # Descarta pares sem nenhuma afinidade de nome E valor muito diferente
            if s_nome < 0.20 and diff_val > 0.10:
                continue

            scores[(ip, ib)] = score

    # ── Casamento global: atribui pares por ordem de melhor score ────────
    pares_ordenados = sorted(scores.items(), key=lambda x: x[1], reverse=True)
    usados_prev  = set()
    usados_banco = set()
    atribuicoes  = {}   # ip -> (ib, score)

    for (ip, ib), score in pares_ordenados:
        if ip in usados_prev or ib in usados_banco:
            continue
        if score < 0.30:
            break
        atribuicoes[ip] = (ib, score)
        usados_prev.add(ip)
        usados_banco.add(ib)

    LIMITE_ALERTA = float(limite_alerta)

    # ── Monta linhas da tabela ───────────────────────────────────────────
    linhas = []

    for ip, prev in deb_prev.iterrows():
        if ip in atribuicoes:
            ib, score = atribuicoes[ip]
            deb  = deb_banco.loc[ib]
            diff = round(deb["debito"] - prev["debito"], 2)
            pct  = diff / prev["debito"] * 100 if prev["debito"] else 0
            s_nome = sim_nome(prev["descricao"], deb["descricao"])

            if abs(diff) <= prev["debito"] * 0.02 and s_nome >= 0.40:
                status = "✅ OK"
            elif abs(diff) > prev["debito"] * 0.02 and s_nome >= 0.40:
                status = "⚠️ VALOR DIFERENTE"
            elif abs(diff) <= prev["debito"] * 0.02 and s_nome < 0.40:
                status = "⚠️ BENEFICIÁRIO DIFERENTE"
            else:
                status = "⚠️ DIVERGÊNCIA"

            # Alerta somente quando há divergência E o valor da diferença passa o limite
            if ("VALOR" in status or "DIVERGÊNCIA" in status) and abs(diff) > LIMITE_ALERTA:
                alerta = f"🔴 ATENÇÃO — diferença de R$ {abs(diff):,.2f}"
            elif "BENEFICIÁRIO" in status and deb["debito"] > LIMITE_ALERTA:
                alerta = f"🔴 ATENÇÃO — beneficiário divergente R$ {deb['debito']:,.2f}"
            else:
                alerta = ""

            linhas.append({
                "Status":                  status,
                "🔴 Alerta":               alerta,
                "Data Prevista":           prev["data"],
                "Beneficiário Previsto":   prev["descricao"],
                "Valor Previsto (R$)":     prev["debito"],
                "Data Pago":               deb["data"],
                "Banco":                   deb["banco"],
                "Pago Para":               deb["descricao"],
                "Valor Pago (R$)":         deb["debito"],
                "Diferença (R$)":          diff,
                "Diferença (%)":           round(pct, 1),
            })
        else:
            linhas.append({
                "Status":                  "🕐 NÃO PAGO",
                "🔴 Alerta":               "",
                "Data Prevista":           prev["data"],
                "Beneficiário Previsto":   prev["descricao"],
                "Valor Previsto (R$)":     prev["debito"],
                "Data Pago":               None,
                "Banco":                   "",
                "Pago Para":               "",
                "Valor Pago (R$)":         None,
                "Diferença (R$)":          -prev["debito"],
                "Diferença (%)":           -100.0,
            })

    # ── Débitos bancários sem par = não previstos ────────────────────────
    OPERACIONAL = ["TARIFA","IOF","JUROS","TAXA","INSS","FGTS","SALDO","RENTAB",
                   "FACILCRED","RENDE FACIL","DEBITO SERV"]
    for ib, deb in deb_banco.iterrows():
        if ib in usados_banco:
            continue
        desc = str(deb["descricao"]).upper()
        if any(p in desc for p in OPERACIONAL):
            continue
        linhas.append({
            "Status":                  "🚨 NÃO PREVISTO",
            "🔴 Alerta":               "",
            "Data Prevista":           None,
            "Beneficiário Previsto":   "",
            "Valor Previsto (R$)":     None,
            "Data Pago":               deb["data"],
            "Banco":                   deb["banco"],
            "Pago Para":               deb["descricao"],
            "Valor Pago (R$)":         deb["debito"],
            "Diferença (R$)":          deb["debito"],
            "Diferença (%)":           None,
        })

    return pd.DataFrame(linhas)



def cor_linha(row):
    s = str(row.get("Status",""))
    if s.startswith("✅"):
        bg = "#d4edda; color:#155724"
    elif "BENEFICIÁRIO" in s:
        bg = "#f8d7da; color:#721c24"   # vermelho — nome trocado é mais grave
    elif s.startswith("⚠️"):
        bg = "#fff3cd; color:#856404"   # amarelo — só valor diferente
    elif s.startswith("🚨"):
        bg = "#f8d7da; color:#721c24"
    else:
        bg = "#e2e3e5; color:#383d41"
    return [f"background-color:{bg}"] * len(row)


def fmt_brl(v):
    try:    return f"R$ {float(v):,.2f}"
    except: return ""

def fmt_dt(v):
    try:    return pd.Timestamp(v).strftime("%d/%m/%Y")
    except: return ""

# ─────────────────────────────────────────────────────────────────────────────
st.title("🌾 Débitos Previstos × Efetivados")
st.caption("Moinho de Trigo — comparação diária: o que estava previsto na planilha vs o que saiu dos bancos")

# Sidebar
with st.sidebar:
    st.header("📂 Arquivos")
    extratos_up = st.file_uploader("Extratos bancários (PDF, XLSX, CSV)",
        type=["pdf","xlsx","xls","csv"], accept_multiple_files=True)
    planilha_up = st.file_uploader("Planilha de previsões (FC)",
        type=["xlsx","xls"])
    st.caption("Nomeie: bradesco_*.pdf · sicredi_*.pdf · bb_*.pdf · bb_alimentos_*.pdf · banrisul_*.pdf")
    st.divider()
    data_sel = st.date_input("📅 Data do extrato", value=datetime.now().date())
    mostrar_periodo = st.toggle("Ver período (mais de um dia)", value=False)
    if mostrar_periodo:
        data_fim = st.date_input("Até", value=datetime.now().date())
    else:
        data_fim = data_sel
    limite_alerta = st.number_input("🔴 Alertar diferenças acima de (R$)",
                                    min_value=0, value=1500, step=500)
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

# Determina meses necessários a partir das datas selecionadas
meses_necessarios = list({_mes_nome(d.month)
                          for d in pd.date_range(data_sel, data_fim, freq="MS").date.tolist()
                          or [data_sel]})
if not meses_necessarios:
    meses_necessarios = [_mes_nome(data_sel.month)]

# Lê planilha
with st.spinner("Lendo planilha de previsões..."):
    df_prev, err = ler_planilha(planilha_up.read(), meses=meses_necessarios)
    if err:
        st.error(err); st.stop()
    st.sidebar.success(f"✔ {planilha_up.name} ({len(df_prev)} lançamentos)")

# Filtra pelo intervalo de datas
data_ini_ts = pd.Timestamp(data_sel)
data_fim_ts = pd.Timestamp(data_fim)

df_banco = df_banco[(df_banco["data"] >= data_ini_ts) & (df_banco["data"] <= data_fim_ts)].copy()
df_prev  = df_prev[(df_prev["data"]  >= data_ini_ts) & (df_prev["data"]  <= data_fim_ts)].copy()

if df_banco.empty and df_prev.empty:
    st.warning(f"Nenhum lançamento encontrado para {data_sel.strftime('%d/%m/%Y')}. "
               "Verifique se a data bate com os extratos enviados.")
    st.stop()

periodo_label = (data_sel.strftime("%d/%m/%Y") if data_sel == data_fim
                 else f"{data_sel.strftime('%d/%m/%Y')} a {data_fim.strftime('%d/%m/%Y')}")
st.info(f"📅 Mostrando: **{periodo_label}** — "
        f"{len(df_banco[df_banco['debito']>0])} débitos bancários · "
        f"{len(df_prev[df_prev['debito']>0])} previstos")

# Saldos
st.subheader("💰 Saldos dos Bancos")
bancos_presentes = [b for b in ["Bradesco","Sicredi","BB","BB Alimentos","Banrisul"]
                    if not df_banco[df_banco["banco"] == b].empty]
cols = st.columns(len(bancos_presentes) + 1)
for i, b in enumerate(bancos_presentes):
    sub = df_banco[df_banco["banco"] == b]["saldo"]
    sub_nz = sub[sub != 0]
    s = float(sub_nz.iloc[-1]) if not sub_nz.empty else 0.0
    cols[i].metric(b, fmt_brl(s))

# Saldo FC da planilha para o dia selecionado
from src.readers.planilha import ler_saldo_fc as _ler_saldo_fc
saldo_fc = None
if planilha_up:
    try:
        import tempfile as _tmp
        with _tmp.NamedTemporaryFile(suffix=".xlsx", delete=False) as _t:
            planilha_up.seek(0); _t.write(planilha_up.read()); _tp = Path(_t.name)
        saldo_fc = _ler_saldo_fc(_tp, pd.Timestamp(data_sel))
        _tp.unlink(missing_ok=True)
    except Exception:
        pass

cols[-1].metric("📊 Saldo FC do Dia", fmt_brl(saldo_fc) if saldo_fc is not None else "—")

# Concilia
with st.spinner("Comparando previstos × efetivados..."):
    df_res = conciliar(df_prev, df_banco, limite_alerta=limite_alerta)

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
        "Status":                  st.column_config.TextColumn(width=200),
        "🔴 Alerta":               st.column_config.TextColumn(width=280),
        "Data Prevista":           st.column_config.TextColumn("Data Prev.", width=100),
        "Beneficiário Previsto":   st.column_config.TextColumn("Previsto Para", width=230),
        "Valor Previsto (R$)":     st.column_config.TextColumn("Vlr Previsto", width=130),
        "Data Pago":               st.column_config.TextColumn("Data Pago", width=100),
        "Banco":                   st.column_config.TextColumn(width=90),
        "Pago Para":               st.column_config.TextColumn(width=230),
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
