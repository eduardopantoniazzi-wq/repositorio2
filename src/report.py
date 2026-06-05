"""
Relatório local em terminal (colorido) para execução sem Google Sheets.
"""

from __future__ import annotations
from colorama import Fore, Style, init

import pandas as pd

init(autoreset=True)


def imprimir_resumo(df: pd.DataFrame, alertas: list[dict]) -> None:
    hoje = df["data"].max()
    df_hoje = df[df["data"] == hoje]

    print("\n" + "═" * 60)
    print(f"  RELATÓRIO FINANCEIRO — {str(hoje)[:10]}")
    print("═" * 60)

    # Saldos por banco
    print(f"\n{Style.BRIGHT}SALDOS POR BANCO:{Style.RESET_ALL}")
    saldo_total = 0
    for banco in df["banco"].unique():
        ultimo = df[df["banco"] == banco]["saldo"].iloc[-1]
        saldo_total += ultimo
        cor = Fore.GREEN if ultimo >= 0 else Fore.RED
        print(f"  {banco:<12} {cor}R$ {ultimo:>15,.2f}{Style.RESET_ALL}")
    print(f"  {'TOTAL':<12} {Fore.CYAN}R$ {saldo_total:>15,.2f}{Style.RESET_ALL}")

    # Movimento do dia
    entradas = df_hoje["credito"].sum()
    saidas   = df_hoje["debito"].sum()
    print(f"\n{Style.BRIGHT}MOVIMENTO DO DIA ({str(hoje)[:10]}):{Style.RESET_ALL}")
    print(f"  Entradas: {Fore.GREEN}R$ {entradas:>12,.2f}{Style.RESET_ALL}")
    print(f"  Saídas:   {Fore.RED}R$ {saidas:>12,.2f}{Style.RESET_ALL}")
    print(f"  Resultado:{Fore.CYAN} R$ {entradas - saidas:>12,.2f}{Style.RESET_ALL}")

    # Alertas
    criticos = [a for a in alertas if a["nivel"] == "CRITICO"]
    atencoes = [a for a in alertas if a["nivel"] == "ATENCAO"]

    print(f"\n{Style.BRIGHT}ALERTAS: {len(alertas)} total "
          f"({len(criticos)} críticos, {len(atencoes)} atenção){Style.RESET_ALL}")

    for a in criticos:
        print(f"  {Fore.RED}[CRÍTICO] {a['tipo']}{Style.RESET_ALL}")
        print(f"           {a['descricao']}")

    for a in atencoes:
        print(f"  {Fore.YELLOW}[ATENÇÃO] {a['tipo']}{Style.RESET_ALL}")
        print(f"           {a['descricao']}")

    if not alertas:
        print(f"  {Fore.GREEN}✔ Nenhum alerta encontrado.{Style.RESET_ALL}")

    print("═" * 60 + "\n")
