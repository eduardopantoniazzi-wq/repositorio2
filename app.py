"""
Controle Financeiro Anti-Desvio — Moinho de Trigo
App web interativo: compara previstos (planilha FC) x efetivados (extratos)
"""

import io
import re
import tempfile
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Controle Financeiro — Moinho de Trigo",
    page_icon="🌾",
    layout="wide",
)

# ── Helpers de leitura (inline, sem depender de src/) ────────────────────────

def _limpar_valor(v) -> float:
    s = re.sub(r"[^\d,.\-]", "", str(v or ""))
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except Exception:
        return 0.0


def _re_data():
    return re.compile(r"\d{2}/\d{2}/\d{4}")


def _parse_bradesco_pdf(path: str) -> pd.DataFrame:
    import pdfplumber
    RE_VALOR = re.compile(r"-?[\d]{1,3}(?:\.\d{3})*,\d{2}")
    registros = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            texto = page.extract_text(layout=True) or ""
            linhas = texto.split("\n")
            buf_desc = []
            data_atual = ""
            for linha in linhas:
                vals = RE_VALOR.findall(linha)
                datas = _re_data().findall(linha)
                if datas and len(vals) >= 2:
                    if buf_desc and data_atual:
                        pass
                    data_atual = datas[0]
                    saldo = _limpar_valor(vals[-1])
                    valor = _limpar_valor(vals[-2])
                    desc = re.sub(r"[\d]{2}/[\d]{2}/[\d]{4}", "", linha)
                    desc = RE_VALOR.sub("", desc).strip()
                    buf_desc = [desc]
                    credito = valor if valor >= 0 else 0.0
                    debito  = abs(valor) if valor < 0 else 0.0
                    registros.append({
                        "data": data_atual, "descricao": desc.strip(),
                        "credito": credito, "debito": debito, "saldo": saldo,
                    })
                elif linha.strip() and not datas and buf_desc:
                    if registros:
                        registros[-1]["descricao"] += " " + linha.strip()
    return pd.DataFrame(registros)


def _parse_sicredi_pdf(path: str) -> pd.DataFrame:
    import pdfplumber
    registros = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            tbl = page.extract_table()
            if not tbl:
                continue
            for row in tbl:
                if not row or not _re_data().search(str(row[0] or "")):
                    continue
                while len(row) < 5:
                    row.append(None)
                data  = _re_data().search(str(row[0])).group()
                desc  = str(row[1] or "").strip()
                doc   = str(row[2] or "").strip()
                val   = _limpar_valor(row[3])
                saldo = _limpar_valor(row[4])
                registros.append({
                    "data": data, "descricao": desc, "documento": doc,
                    "credito": val if val >= 0 else 0.0,
                    "debito": abs(val) if val < 0 else 0.0,
                    "saldo": saldo,
                })
    return pd.DataFrame(registros)


def _parse_bb_pdf(path: str) -> pd.DataFrame:
    import pdfplumber
    RE_VAL_CD = re.compile(r"([\d]{1,3}(?:\.\d{3})*,\d{2})\s*\n?\s*([CD])")
    RE_VAL_NUM = re.compile(r"-?[\d]{1,3}(?:\.\d{3})*,\d{2}")
    registros = []
    with pdfplumber.open(path) as pdf:
        all_rows = []
        for page in pdf.pages:
            tbl = page.extract_table()
            if not tbl:
                continue
            start = 0
            for i, row in enumerate(tbl):
                if row and any("Hist" in str(c or "") or "balancete" in str(c or "").lower() for c in row):
                    start = i + 1
                    break
            all_rows.extend(tbl[start:])
    i = 0
    while i < len(all_rows):
        row = list(all_rows[i])
        while len(row) < 8:
            row.append(None)
        dt_bal = str(row[0] or "").strip()
        dt_mov = str(row[1] or "").strip()
        historico = str(row[4] or "").strip()
        val_raw = str(row[6] or "").strip()
        sld_raw = str(row[7] or "").strip()
        tem_data = _re_data().search(dt_bal) or _re_data().search(dt_mov)
        if not tem_data and not RE_VAL_CD.search(val_raw) and not historico:
            i += 1
            continue
        data = ""
        m = _re_data().search(dt_mov) or _re_data().search(dt_bal)
        if m:
            data = m.group()
        hist_limpo = re.sub(r"^\d{3,5}\s+", "", historico).strip()
        hist_limpo = re.sub(r"^\d{3,5}\s+", "", hist_limpo).strip()
        beneficiario = ""
        if i + 1 < len(all_rows):
            prox = list(all_rows[i + 1])
            while len(prox) < 8:
                prox.append(None)
            prox_dt = str(prox[0] or "").strip()
            prox_hist = str(prox[4] or "").strip()
            prox_val = str(prox[6] or "").strip()
            if not _re_data().search(prox_dt) and prox_hist and not RE_VAL_CD.search(prox_val) and not RE_VAL_NUM.search(prox_val):
                beneficiario = prox_hist
                i += 1
        descricao = f"{hist_limpo} / {beneficiario}".strip(" /") if beneficiario else hist_limpo
        m_cd = RE_VAL_CD.search(val_raw)
        if m_cd:
            v = _limpar_valor(m_cd.group(1))
            cred = v if m_cd.group(2) == "C" else 0.0
            deb  = v if m_cd.group(2) == "D" else 0.0
        else:
            m_n = RE_VAL_NUM.search(val_raw)
            v = _limpar_valor(m_n.group() if m_n else "0")
            cred = v if v >= 0 else 0.0
            deb  = abs(v) if v < 0 else 0.0
        saldo_m = RE_VAL_NUM.search(sld_raw)
        saldo = _limpar_valor(saldo_m.group() if saldo_m else "0")
        if data:
            registros.append({"data": data, "descricao": descricao, "credito": cred, "debito": deb, "saldo": saldo})
        i += 1
    return pd.DataFrame(registros)


def _parse_excel_banco(path: str, banco: str) -> pd.DataFrame:
    for hr in range(0, 10):
        try:
            df = pd.read_excel(path, header=hr, dtype=str)
            cols = [str(c).lower() for c in df.columns]
            if any("data" in c or "dt" in c for c in cols):
                break
        except Exception:
            continue
    mapa = {}
    for col in df.columns:
        cl = str(col).strip().lower()
        if re.search(r"data|dt\b", cl):
            mapa[col] = "data"
        elif re.search(r"hist|desc|lançamento|lancamento", cl):
            mapa[col] = "descricao"
        elif re.search(r"crédito|credito|entrada", cl):
            mapa[col] = "credito"
        elif re.search(r"déb|deb|saída|saida", cl):
            mapa[col] = "debito"
        elif re.search(r"saldo", cl):
            mapa[col] = "saldo"
        elif re.search(r"doc|document", cl):
            mapa[col] = "documento"
        elif re.search(r"valor", cl):
            mapa[col] = "valor"
    df = df.rename(columns=mapa)
    if "valor" in df.columns and "credito" not in df.columns:
        df["valor_n"] = df["valor"].apply(_limpar_valor)
        df["credito"] = df["valor_n"].apply(lambda v: v if v > 0 else 0.0)
        df["debito"]  = df["valor_n"].apply(lambda v: abs(v) if v < 0 else 0.0)
    for c in ["credito", "debito", "saldo"]:
        if c in df.columns:
            df[c] = df[c].apply(_limpar_valor)
        else:
            df[c] = 0.0
    return df


def ler_extrato(nome: str, conteudo: bytes) -> pd.DataFrame:
    """Lê extrato bancário de qualquer formato suportado."""
    nome_lower = nome.lower()
    if "bradesco" in nome_lower:
        banco = "Bradesco"
    elif "sicredi" in nome_lower:
        banco = "Sicredi"
    elif nome_lower.startswith("bb") or "banco_brasil" in nome_lower or "bancobrasil" in nome_lower:
        banco = "BB"
    else:
        banco = "Banco"

    with tempfile.NamedTemporaryFile(suffix=Path(nome).suffix, delete=False) as tmp:
        tmp.write(conteudo)
        tmp_path = tmp.name

    try:
        if nome_lower.endswith(".pdf"):
            if banco == "Bradesco":
                df = _parse_bradesco_pdf(tmp_path)
            elif banco == "Sicredi":
                df = _parse_sicredi_pdf(tmp_path)
            else:
                df = _parse_bb_pdf(tmp_path)
        elif nome_lower.endswith((".xlsx", ".xls")):
            df = _parse_excel_banco(tmp_path, banco)
        elif nome_lower.endswith(".csv"):
            df = pd.read_csv(io.BytesIO(conteudo), sep=None, engine="python", dtype=str, on_bad_lines="skip")
        else:
            return pd.DataFrame()
    finally:
        Path(tmp_path).unlink(missing_ok=True)

    df["banco"] = banco
    if "data" in df.columns:
        df["data"] = pd.to_datetime(df["data"], dayfirst=True, errors="coerce")
    for c in ["credito", "debito", "saldo"]:
        if c not in df.columns:
            df[c] = 0.0
    if "descricao" not in df.columns:
        df["descricao"] = ""
    return df.dropna(subset=["data"])


def ler_planilha_fc(conteudo: bytes, meses: list[str]) -> pd.DataFrame:
    """Lê planilha de fluxo de caixa no formato matricial."""
    sys_path_backup = None
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp.write(conteudo)
        tmp_path = tmp.name

    try:
        import sys
        sys.path.insert(0, "/home/user/repositorio2")
        from src.readers.planilha import ler_planilha
        df = ler_planilha(Path(tmp_path), meses=meses)
    except Exception:
        df = pd.DataFrame()
    finally:
        Path(tmp_path).unlink(missing_ok=True)

    return df


def conciliar_simples(df_previsto: pd.DataFrame, df_banco: pd.DataFrame, janela_dias: int = 3) -> pd.DataFrame:
    """Versão simplificada da conciliação para exibição interativa."""
    from difflib import SequenceMatcher
    STOPWORDS = {"nf","nota","fiscal","ltda","sa","eireli","me","epp","pagamento","de","boleto","pag","pgto","ted","pix"}

    def normalizar(s):
        s = re.sub(r"[^\w\s]", " ", str(s).upper())
        return " ".join(w for w in s.split() if len(w) > 2 and w.lower() not in STOPWORDS)

    def similar(a, b):
        na, nb = normalizar(a), normalizar(b)
        if not na or not nb:
            return 0.0
        wa, wb = set(na.split()), set(nb.split())
        base = SequenceMatcher(None, na, nb).ratio()
        return max(base, 0.6) if wa & wb else base

    debitos = df_banco[df_banco["debito"] > 0].copy().reset_index(drop=True)
    previstos = df_previsto[df_previsto["debito"] > 0].copy().reset_index(drop=True)
    debitos["_usado"] = False

    rows = []
    for _, prev in previstos.iterrows():
        melhor_idx, melhor_score = None, 0.0
        for id_, deb in debitos[~debitos["_usado"]].iterrows():
            if prev["debito"] == 0:
                continue
            diff_val = abs(deb["debito"] - prev["debito"]) / prev["debito"]
            if diff_val > 0.50:
                continue
            try:
                diff_dias = abs((deb["data"] - prev["data"]).days)
            except Exception:
                diff_dias = 999
            if diff_dias > janela_dias:
                continue
            score = (max(0.0, 1 - diff_val/0.5) * 0.55 +
                     similar(prev["descricao"], deb["descricao"]) * 0.30 +
                     max(0.0, 1 - diff_dias/(janela_dias+1)) * 0.15)
            if score > melhor_score:
                melhor_score, melhor_idx = score, id_

        if melhor_idx is not None and melhor_score >= 0.40:
            deb = debitos.loc[melhor_idx]
            debitos.at[melhor_idx, "_usado"] = True
            diff_val = abs(deb["debito"] - prev["debito"]) / prev["debito"]
            sim = similar(prev["descricao"], deb["descricao"])
            if diff_val <= 0.02 and sim >= 0.4:
                status = "CONCILIADO"
            elif diff_val > 0.02:
                status = "DIVERGÊNCIA DE VALOR"
            else:
                status = "DIVERGÊNCIA DE BENEFICIÁRIO"
            rows.append({
                "status":            status,
                "data_prevista":     prev["data"],
                "beneficiário_planilha": prev["descricao"],
                "valor_previsto":    prev["debito"],
                "data_pagamento":    deb["data"],
                "banco":             deb["banco"],
                "beneficiário_extrato": deb["descricao"],
                "valor_pago":        deb["debito"],
                "diferença_R$":      round(deb["debito"] - prev["debito"], 2),
            })
        else:
            rows.append({
                "status":            "PENDENTE",
                "data_prevista":     prev["data"],
                "beneficiário_planilha": prev["descricao"],
                "valor_previsto":    prev["debito"],
                "data_pagamento":    None,
                "banco":             "",
                "beneficiário_extrato": "",
                "valor_pago":        0.0,
                "diferença_R$":      -prev["debito"],
            })

    # Não previstos
    nao_prev = []
    for _, deb in debitos[~debitos["_usado"]].iterrows():
        val = deb["debito"]
        desc = str(deb["descricao"]).upper()
        baixo = any(p in desc for p in ["TARIFA","IOF","JUROS","TAXA","INSS","FGTS","SALARIO","IMPOSTO"])
        if baixo:
            risco = "OPERACIONAL"
        elif val >= 5000 or (val >= 1000 and any(p in desc for p in ["PIX","TRANSF","TED"])):
            risco = "🔴 ALTO"
        elif val >= 500:
            risco = "🟡 MÉDIO"
        else:
            risco = "BAIXO"
        nao_prev.append({
            "status":            "NÃO PREVISTO",
            "data_prevista":     None,
            "beneficiário_planilha": f"[{risco}]",
            "valor_previsto":    0.0,
            "data_pagamento":    deb["data"],
            "banco":             deb["banco"],
            "beneficiário_extrato": deb["descricao"],
            "valor_pago":        deb["debito"],
            "diferença_R$":      deb["debito"],
        })

    return pd.DataFrame(rows + nao_prev)


# ── Estilos de cor por status ────────────────────────────────────────────────

def colorir_linha(row):
    s = row["status"]
    if s == "CONCILIADO":
        bg = "background-color: #d4edda; color: #155724"
    elif "DIVERGÊNCIA" in s:
        bg = "background-color: #fff3cd; color: #856404"
    elif s == "NÃO PREVISTO":
        bg = "background-color: #f8d7da; color: #721c24"
    elif s == "PENDENTE":
        bg = "background-color: #e2e3e5; color: #383d41"
    else:
        bg = ""
    return [bg] * len(row)


# ── Layout principal ─────────────────────────────────────────────────────────

st.title("🌾 Controle Financeiro Anti-Desvio")
st.caption("Moinho de Trigo — Compara pagamentos previstos (planilha FC) x efetivados (extratos bancários)")

# ── Sidebar: uploads ─────────────────────────────────────────────────────────
with st.sidebar:
    st.header("📂 Upload de Arquivos")

    st.subheader("Extratos Bancários")
    st.caption("Bradesco, Sicredi ou BB — PDF, XLSX ou CSV")
    extratos_up = st.file_uploader(
        "Selecione os extratos", type=["pdf", "xlsx", "xls", "csv"],
        accept_multiple_files=True, key="extratos",
    )

    st.subheader("Planilha de Previsão (FC)")
    st.caption("receitas_despesas*.xlsx ou fluxo*.xlsx")
    planilha_up = st.file_uploader(
        "Selecione a planilha", type=["xlsx", "xls"], key="planilha",
    )

    st.subheader("Período")
    _MESES = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho",
              "Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"]
    mes_atual = _MESES[datetime.now().month - 1]
    mes_ant   = _MESES[datetime.now().month - 2] if datetime.now().month > 1 else "Dezembro"
    meses_sel = st.multiselect("Meses da planilha", _MESES, default=[mes_ant, mes_atual])

    janela = st.slider("Janela de datas (dias)", 0, 7, 3,
                       help="Quantos dias de diferença aceitar ao casar previsto x efetivado")

    rodar = st.button("▶ Analisar", type="primary", use_container_width=True)

# ── Processamento ─────────────────────────────────────────────────────────────
if rodar:
    if not extratos_up:
        st.error("Faça upload de ao menos um extrato bancário.")
        st.stop()

    with st.spinner("Lendo extratos bancários..."):
        dfs_banco = []
        for f in extratos_up:
            try:
                df = ler_extrato(f.name, f.read())
                if not df.empty:
                    dfs_banco.append(df)
                    st.sidebar.success(f"✔ {f.name} ({len(df)} lançamentos)")
            except Exception as e:
                st.sidebar.warning(f"⚠ {f.name}: {e}")

        if not dfs_banco:
            st.error("Nenhum extrato pôde ser lido.")
            st.stop()

        df_banco = pd.concat(dfs_banco, ignore_index=True)
        df_banco_real = df_banco[df_banco["banco"] != "Planilha"].copy()

    df_previsto = None
    if planilha_up:
        with st.spinner("Lendo planilha de previsão..."):
            try:
                df_previsto = ler_planilha_fc(planilha_up.read(), meses=meses_sel or None)
                if df_previsto is not None and not df_previsto.empty:
                    st.sidebar.success(f"✔ {planilha_up.name} ({len(df_previsto)} lançamentos)")
            except Exception as e:
                st.sidebar.warning(f"⚠ Planilha: {e}")

    # ── Saldos ───────────────────────────────────────────────────────────────
    st.header("💰 Saldos Bancários")
    cols = st.columns(4)
    total = 0.0
    for i, banco in enumerate(["Bradesco", "Sicredi", "BB"]):
        sub = df_banco_real[df_banco_real["banco"] == banco]["saldo"]
        saldo = sub.iloc[-1] if not sub.empty else 0.0
        total += saldo
        cols[i].metric(banco, f"R$ {saldo:,.2f}")
    cols[3].metric("**TOTAL**", f"R$ {total:,.2f}")

    # ── Conciliação ──────────────────────────────────────────────────────────
    if df_previsto is not None and not df_previsto.empty:
        with st.spinner("Conciliando previsto x efetivado..."):
            df_conc = conciliar_simples(df_previsto, df_banco_real, janela_dias=janela)

        # Métricas resumo
        st.header("📊 Resumo da Conciliação")
        n_conc  = (df_conc["status"] == "CONCILIADO").sum()
        n_div   = df_conc["status"].str.startswith("DIVERGÊNCIA").sum()
        n_pend  = (df_conc["status"] == "PENDENTE").sum()
        n_nao   = (df_conc["status"] == "NÃO PREVISTO").sum()
        val_nao = df_conc[df_conc["status"] == "NÃO PREVISTO"]["valor_pago"].sum()
        val_div = df_conc[df_conc["status"].str.startswith("DIVERGÊNCIA")]["diferença_R$"].abs().sum()

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("✅ Conciliados",      n_conc)
        c2.metric("⚠️ Divergências",     n_div, delta=f"R$ {val_div:,.2f}" if n_div else None, delta_color="inverse")
        c3.metric("🕐 Pendentes",        n_pend)
        c4.metric("🚨 Não Previstos",    n_nao, delta=f"R$ {val_nao:,.2f}" if n_nao else None, delta_color="inverse")

        # ── Filtros ──────────────────────────────────────────────────────────
        st.header("🔍 Tabela de Conciliação")
        col_f1, col_f2, col_f3 = st.columns(3)

        status_opts = sorted(df_conc["status"].unique().tolist())
        status_sel  = col_f1.multiselect("Status", status_opts, default=status_opts)

        banco_opts = ["Todos"] + sorted(df_banco_real["banco"].unique().tolist())
        banco_sel  = col_f2.selectbox("Banco", banco_opts)

        busca = col_f3.text_input("Buscar beneficiário", placeholder="Ex: Cooperativa, Fornecedor...")

        df_view = df_conc[df_conc["status"].isin(status_sel)].copy()
        if banco_sel != "Todos":
            df_view = df_view[df_view["banco"] == banco_sel]
        if busca:
            mask = (
                df_view["beneficiário_planilha"].str.contains(busca, case=False, na=False) |
                df_view["beneficiário_extrato"].str.contains(busca, case=False, na=False)
            )
            df_view = df_view[mask]

        # Formata datas e valores para exibição
        def fmt_data(v):
            try:
                return pd.Timestamp(v).strftime("%d/%m/%Y")
            except Exception:
                return ""

        df_show = df_view.copy()
        df_show["data_prevista"]  = df_show["data_prevista"].apply(fmt_data)
        df_show["data_pagamento"] = df_show["data_pagamento"].apply(fmt_data)
        df_show["valor_previsto"] = df_show["valor_previsto"].apply(lambda v: f"R$ {v:,.2f}" if v else "")
        df_show["valor_pago"]     = df_show["valor_pago"].apply(lambda v: f"R$ {v:,.2f}" if v else "")
        df_show["diferença_R$"]   = df_show["diferença_R$"].apply(
            lambda v: f"+R$ {v:,.2f}" if v > 0 else (f"-R$ {abs(v):,.2f}" if v < 0 else "")
        )

        st.dataframe(
            df_show.style.apply(colorir_linha, axis=1),
            use_container_width=True,
            height=550,
            column_config={
                "status":                   st.column_config.TextColumn("Status", width=180),
                "data_prevista":            st.column_config.TextColumn("Data Prevista", width=110),
                "beneficiário_planilha":    st.column_config.TextColumn("Beneficiário (Planilha)", width=250),
                "valor_previsto":           st.column_config.TextColumn("Valor Previsto", width=120),
                "data_pagamento":           st.column_config.TextColumn("Data Pag.", width=110),
                "banco":                    st.column_config.TextColumn("Banco", width=90),
                "beneficiário_extrato":     st.column_config.TextColumn("Beneficiário (Extrato)", width=280),
                "valor_pago":               st.column_config.TextColumn("Valor Pago", width=120),
                "diferença_R$":             st.column_config.TextColumn("Diferença", width=110),
            },
        )

        # ── Download ─────────────────────────────────────────────────────────
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df_conc.to_excel(writer, sheet_name="Conciliação", index=False)
            df_conc[df_conc["status"] == "NÃO PREVISTO"].to_excel(writer, sheet_name="Não Previstos", index=False)
            df_conc[df_conc["status"] == "PENDENTE"].to_excel(writer, sheet_name="Pendentes", index=False)
            df_banco_real.to_excel(writer, sheet_name="Extrato Consolidado", index=False)
        st.download_button(
            "📥 Baixar relatório Excel",
            data=buf.getvalue(),
            file_name=f"conciliacao_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    else:
        # Sem planilha: mostra só o extrato consolidado
        st.info("Faça upload da planilha de previsão (FC) para ver a conciliação.")
        st.header("📄 Extrato Consolidado")
        st.dataframe(df_banco_real[["data","banco","descricao","credito","debito","saldo"]],
                     use_container_width=True, height=500)

else:
    st.info("👈 Faça upload dos arquivos na barra lateral e clique em **▶ Analisar**.")
    st.markdown("""
    ### Como usar
    1. **Extratos bancários**: faça upload dos PDFs ou Excel do Bradesco, Sicredi e BB
    2. **Planilha FC**: faça upload do arquivo `receitas_despesas_2026.xlsx`
    3. **Meses**: selecione o(s) mês(es) a analisar
    4. Clique em **▶ Analisar**

    ### O que a tabela mostra
    | Cor | Significado |
    |-----|-------------|
    | 🟢 Verde | Conciliado — previsto e pago, valores batem |
    | 🟡 Amarelo | Divergência — valor ou beneficiário diferente do previsto |
    | 🔴 Vermelho | **Não previsto** — pago no banco mas não estava na planilha |
    | ⬜ Cinza | Pendente — previsto mas ainda não pago |
    """)
