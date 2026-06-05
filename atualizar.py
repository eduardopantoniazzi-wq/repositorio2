#!/usr/bin/env python3
"""
Sistema de Controle Financeiro Anti-Desvio — Moinho de Trigo
=============================================================
Uso:
  python atualizar.py                         # modo completo (com Google Sheets)
  python atualizar.py --sem-sheets            # sem Google Sheets
  python atualizar.py --dir PASTA_EXTRATOS    # pasta diferente
  python atualizar.py --planilha ARQUIVO.xlsx # planilha separada
  python atualizar.py --meses Junho Julho     # meses específicos da planilha

Nomeação esperada dos arquivos de extrato na pasta:
  bradesco*.pdf / bradesco*.xlsx / bradesco*.csv
  sicredi*.pdf  / sicredi*.xlsx  / sicredi*.csv
  bb*.pdf       / bb*.xlsx       / bb*.csv
  receitas*.xlsx / despesas*.xlsx / fluxo*.xlsx  → planilha de receitas/despesas
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


def main() -> None:
    parser = argparse.ArgumentParser(description="Controle Financeiro Anti-Desvio")
    parser.add_argument("--sem-sheets", action="store_true")
    parser.add_argument("--dir",       type=str, default=None,
                        help="Pasta com os extratos bancários (padrão: data/extratos/)")
    parser.add_argument("--planilha",  type=str, default=None,
                        help="Caminho explícito da planilha de receitas/despesas")
    parser.add_argument("--meses",     nargs="*", default=None,
                        help="Meses da planilha a incluir (ex: Junho Julho). "
                             "Padrão: mês atual + anterior")
    args = parser.parse_args()

    from config import EXTRATOS_DIR, GOOGLE_CREDENTIALS_FILE

    diretorio = Path(args.dir) if args.dir else EXTRATOS_DIR
    planilha  = Path(args.planilha) if args.planilha else None

    # Meses padrão: mês atual + mês anterior
    if args.meses:
        meses = args.meses
    else:
        mes_atual    = _MESES_PT[datetime.now().month - 1]
        mes_anterior = _MESES_PT[datetime.now().month - 2] if datetime.now().month > 1 else "Dezembro"
        meses = [mes_anterior, mes_atual]

    print(f"\n{Style.BRIGHT}{'═'*60}")
    print("  CONTROLE FINANCEIRO ANTI-DESVIO — MOINHO DE TRIGO")
    print(f"{'═'*60}{Style.RESET_ALL}")
    print(f"  Data: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    print(f"  Extratos: {diretorio}")
    print(f"  Meses da planilha: {', '.join(meses)}\n")

    # ── 1. Leitura e consolidação ────────────────────────────────────────
    print(f"{Style.BRIGHT}[1/4] Lendo extratos...{Style.RESET_ALL}")
    from src.consolidator import consolidar
    try:
        df = consolidar(
            diretorio=diretorio,
            arquivo_planilha=planilha,
            meses_planilha=meses,
        )
    except (FileNotFoundError, ValueError) as e:
        print(f"{Fore.RED}ERRO: {e}{Style.RESET_ALL}\n")
        sys.exit(1)

    bancos_reais = df[df["banco"] != "Planilha"]["banco"].unique()
    total = len(df)
    print(f"  {Fore.GREEN}✔ {total} lançamentos | Bancos: {', '.join(bancos_reais)}{Style.RESET_ALL}")

    # ── 2. Detecção de desvios ───────────────────────────────────────────
    print(f"\n{Style.BRIGHT}[2/4] Verificações anti-desvio...{Style.RESET_ALL}")
    # Só roda detector nos lançamentos bancários reais (não na planilha)
    df_bancos = df[df["banco"] != "Planilha"].copy()
    from src.detector import rodar_verificacoes
    df_bancos_alertas, alertas = rodar_verificacoes(df_bancos)

    criticos = sum(1 for a in alertas if a["nivel"] == "CRITICO")
    atencoes = sum(1 for a in alertas if a["nivel"] == "ATENCAO")

    if criticos:
        print(f"  {Fore.RED}⚠ {criticos} alerta(s) CRÍTICO(s){Style.RESET_ALL}")
    if atencoes:
        print(f"  {Fore.YELLOW}⚠ {atencoes} alerta(s) de ATENÇÃO{Style.RESET_ALL}")
    if not alertas:
        print(f"  {Fore.GREEN}✔ Nenhum desvio detectado{Style.RESET_ALL}")

    # Propaga alertas para o df completo
    df.loc[df["banco"] != "Planilha", "alerta"] = df_bancos_alertas["alerta"].values

    # ── 3. Base de conhecimento ──────────────────────────────────────────
    print(f"\n{Style.BRIGHT}[3/4] Atualizando base de conhecimento...{Style.RESET_ALL}")
    from src.knowledge_base import atualizar_base
    atualizar_base(df_bancos_alertas)

    # ── 4. Relatório no terminal ─────────────────────────────────────────
    from src.report import imprimir_resumo
    imprimir_resumo(df_bancos_alertas, alertas)

    # ── 5. Google Sheets ─────────────────────────────────────────────────
    usar_sheets = not args.sem_sheets
    if usar_sheets and not GOOGLE_CREDENTIALS_FILE.exists():
        print(f"{Fore.YELLOW}[5/5] Google Sheets pulado "
              f"(credenciais não encontradas em {GOOGLE_CREDENTIALS_FILE}){Style.RESET_ALL}")
        print(f"       → Veja o README.md para configurar as credenciais.\n")
        usar_sheets = False

    if usar_sheets:
        print(f"{Style.BRIGHT}[5/5] Atualizando Google Sheets...{Style.RESET_ALL}")
        try:
            from src.sheets import atualizar_sheets
            url = atualizar_sheets(df_bancos_alertas, alertas, df_planilha=df[df["banco"] == "Planilha"])
            print(f"  {Fore.GREEN}✔ Planilha: {url}{Style.RESET_ALL}\n")
        except Exception as e:
            print(f"  {Fore.RED}ERRO Google Sheets: {e}{Style.RESET_ALL}\n")
    else:
        print(f"[5/5] Google Sheets pulado (--sem-sheets).\n")

    # ── Salva extrato consolidado localmente ──────────────────────────────
    saida = Path("data") / "extrato_consolidado.xlsx"
    df.to_excel(saida, index=False)
    print(f"  Extrato salvo em: {saida}")

    saida_alertas = Path("data") / "alertas_ultimo.xlsx"
    if alertas:
        import pandas as pd_
        pd_.DataFrame(alertas).to_excel(saida_alertas, index=False)
        print(f"  Alertas salvos em: {saida_alertas}\n")


if __name__ == "__main__":
    main()
