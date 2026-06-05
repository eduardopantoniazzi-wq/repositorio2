#!/usr/bin/env python3
"""
Controle Financeiro Anti-Desvio — Moinho de Trigo
===================================================
Compara o que foi PREVISTO na planilha com o que foi EFETIVAMENTE PAGO
nos extratos bancários, destacando qualquer pagamento não autorizado.

Uso:
  python atualizar.py                         # modo completo
  python atualizar.py --sem-sheets            # sem Google Sheets (apenas local)
  python atualizar.py --dir PASTA_EXTRATOS    # pasta diferente
  python atualizar.py --planilha ARQUIVO.xlsx # planilha separada
  python atualizar.py --meses Junho Julho     # meses da planilha

Nomeação dos arquivos de extrato:
  bradesco*.pdf/xlsx/csv  →  Bradesco
  sicredi*.pdf/xlsx/csv   →  Sicredi
  bb*.pdf/xlsx/csv        →  BB
  receitas*.xlsx / despesas*.xlsx / fluxo*.xlsx  →  planilha de previsão
"""

from __future__ import annotations
import sys
import argparse
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from colorama import Fore, Style, init
init(autoreset=True)

_MESES_PT = [
    "Janeiro","Fevereiro","Março","Abril","Maio","Junho",
    "Julho","Agosto","Setembro","Outubro","Novembro","Dezembro",
]


def _linha(char="═", n=62):
    return char * n


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--sem-sheets", action="store_true")
    parser.add_argument("--dir",       type=str, default=None)
    parser.add_argument("--planilha",  type=str, default=None)
    parser.add_argument("--meses",     nargs="*", default=None)
    args = parser.parse_args()

    from config import EXTRATOS_DIR, GOOGLE_CREDENTIALS_FILE

    diretorio = Path(args.dir) if args.dir else EXTRATOS_DIR
    planilha  = Path(args.planilha) if args.planilha else None

    if args.meses:
        meses = args.meses
    else:
        mes_atual    = _MESES_PT[datetime.now().month - 1]
        mes_anterior = _MESES_PT[datetime.now().month - 2] if datetime.now().month > 1 else "Dezembro"
        meses = [mes_anterior, mes_atual]

    periodo = f"{meses[0]} a {meses[-1]} de {datetime.now().year}"

    print(f"\n{Style.BRIGHT}{_linha()}")
    print("  CONTROLE FINANCEIRO ANTI-DESVIO — MOINHO DE TRIGO")
    print(f"  {datetime.now().strftime('%d/%m/%Y %H:%M')} | Período: {periodo}")
    print(f"{_linha()}{Style.RESET_ALL}\n")

    # ── 1. Lê extratos bancários ─────────────────────────────────────────
    print(f"{Style.BRIGHT}[1/4] Lendo extratos bancários...{Style.RESET_ALL}")
    from src.consolidator import consolidar

    try:
        df_banco = consolidar(diretorio=diretorio)
    except Exception as e:
        print(f"{Fore.RED}ERRO ao ler extratos: {e}{Style.RESET_ALL}\n")
        sys.exit(1)

    df_banco_real = df_banco[df_banco["banco"] != "Planilha"].copy()
    print(f"  {Fore.GREEN}✔ {len(df_banco_real)} lançamentos bancários "
          f"({', '.join(df_banco_real['banco'].unique())}){Style.RESET_ALL}")

    # Saldos
    for b in ["Bradesco", "Sicredi", "BB"]:
        sub = df_banco_real[df_banco_real["banco"] == b]["saldo"]
        if not sub.empty:
            print(f"    {b}: R$ {sub.iloc[-1]:>15,.2f}")
    total = sum(
        df_banco_real[df_banco_real["banco"] == b]["saldo"].iloc[-1]
        for b in ["Bradesco", "Sicredi", "BB"]
        if not df_banco_real[df_banco_real["banco"] == b].empty
    )
    print(f"    {'TOTAL':10s}: {Fore.CYAN}R$ {total:>15,.2f}{Style.RESET_ALL}")

    # ── 2. Lê planilha de previsão ───────────────────────────────────────
    print(f"\n{Style.BRIGHT}[2/4] Lendo planilha de previsão...{Style.RESET_ALL}")
    from src.readers.planilha import ler_planilha

    # Detecta planilha automaticamente se não informada
    if planilha is None:
        candidatos = [
            p for p in diretorio.iterdir()
            if p.suffix.lower() in (".xlsx", ".xls")
            and any(k in p.stem.lower() for k in ("receita", "despesa", "fluxo", "caixa", "fc"))
        ]
        planilha = candidatos[0] if candidatos else None

    df_previsto = None
    if planilha and planilha.exists():
        try:
            df_previsto = ler_planilha(planilha, meses=meses)
            print(f"  {Fore.GREEN}✔ {len(df_previsto)} lançamentos previstos "
                  f"de {planilha.name}{Style.RESET_ALL}")
        except Exception as e:
            print(f"  {Fore.YELLOW}⚠ Não foi possível ler a planilha: {e}{Style.RESET_ALL}")
    else:
        print(f"  {Fore.YELLOW}⚠ Planilha de previsão não encontrada em {diretorio}{Style.RESET_ALL}")
        print(f"    Coloque um arquivo receitas*.xlsx ou despesas*.xlsx na pasta de extratos.")

    # ── 3. Conciliação ───────────────────────────────────────────────────
    print(f"\n{Style.BRIGHT}[3/4] Conciliando previsto x efetivado...{Style.RESET_ALL}")
    from src.reconciler import conciliar

    if df_previsto is not None and not df_previsto.empty:
        df_conc, df_nao_prev, df_pend = conciliar(df_previsto, df_banco_real)
    else:
        import pandas as pd
        df_conc     = pd.DataFrame()
        df_nao_prev = pd.DataFrame()
        df_pend     = pd.DataFrame()

    # Resumo no terminal
    if not df_conc.empty:
        n_conc  = (df_conc["status"] == "CONCILIADO").sum()
        n_div   = df_conc["status"].str.startswith("DIVERG", na=False).sum()
        n_pend  = (df_conc["status"] == "PENDENTE").sum()
        print(f"  {Fore.GREEN}✅ Conciliados:  {n_conc:>4}{Style.RESET_ALL}")
        if n_div:
            print(f"  {Fore.YELLOW}⚠  Divergências: {n_div:>4}{Style.RESET_ALL}")
        if n_pend:
            print(f"  {Fore.WHITE}🕐 Pendentes:    {n_pend:>4}{Style.RESET_ALL}")

    if not df_nao_prev.empty:
        n_alto  = (df_nao_prev["risco"] == "ALTO").sum()
        n_med   = (df_nao_prev["risco"] == "MEDIO").sum()
        n_total = len(df_nao_prev)
        val_nao = df_nao_prev[df_nao_prev["risco"].isin(["ALTO","MEDIO"])]["valor"].sum()
        print(f"\n  {Fore.RED}{'━'*50}")
        print(f"  🚨 PAGAMENTOS NÃO PREVISTOS: {n_total} "
              f"({n_alto} alto risco, {n_med} médio)")
        print(f"  Valor em risco: R$ {val_nao:,.2f}")
        print(f"  {'━'*50}{Style.RESET_ALL}")

        nao_prev_risco = df_nao_prev[df_nao_prev["risco"].isin(["ALTO","MEDIO"])]
        for _, r in nao_prev_risco.head(10).iterrows():
            cor = Fore.RED if r["risco"] == "ALTO" else Fore.YELLOW
            print(f"  {cor}[{r['risco']:6s}] {str(r['data'])[:10]}  "
                  f"R$ {r['valor']:>12,.2f}  {r['descricao'][:55]}{Style.RESET_ALL}")
    else:
        print(f"\n  {Fore.GREEN}✅ Nenhum pagamento não previsto encontrado.{Style.RESET_ALL}")

    # ── 4. Google Sheets ─────────────────────────────────────────────────
    usar_sheets = not args.sem_sheets
    if usar_sheets and not GOOGLE_CREDENTIALS_FILE.exists():
        print(f"\n{Fore.YELLOW}[4/4] Google Sheets pulado "
              f"(sem credenciais em {GOOGLE_CREDENTIALS_FILE}){Style.RESET_ALL}\n")
        usar_sheets = False

    if usar_sheets:
        print(f"\n{Style.BRIGHT}[4/4] Publicando no Google Sheets...{Style.RESET_ALL}")
        try:
            from src.sheets_reconciliacao import publicar
            url = publicar(
                df_banco    = df_banco_real,
                df_previsto = df_previsto if df_previsto is not None else pd.DataFrame(),
                df_conc     = df_conc,
                df_nao_prev = df_nao_prev,
                df_pend     = df_pend,
                periodo     = periodo,
            )
            print(f"\n  {Fore.GREEN}✔ Planilha atualizada: {url}{Style.RESET_ALL}")
        except Exception as e:
            print(f"\n  {Fore.RED}ERRO Google Sheets: {e}{Style.RESET_ALL}")
    else:
        print(f"\n[4/4] Google Sheets pulado (--sem-sheets).")

    # ── Salva local ───────────────────────────────────────────────────────
    saida_dir = Path("data")
    saida_dir.mkdir(exist_ok=True)

    df_banco_real.to_excel(saida_dir / "extrato_consolidado.xlsx", index=False)
    if not df_conc.empty:
        df_conc.to_excel(saida_dir / "conciliacao.xlsx", index=False)
    if not df_nao_prev.empty:
        df_nao_prev.to_excel(saida_dir / "nao_previstos.xlsx", index=False)

    print(f"\n  Arquivos locais salvos em data/")
    print(f"\n{_linha('─')}\n")


if __name__ == "__main__":
    main()
