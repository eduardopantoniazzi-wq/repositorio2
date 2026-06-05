#!/usr/bin/env python3
"""
Sistema de Controle Financeiro Anti-Desvio — Moinho de Trigo
=============================================================
Uso: python atualizar.py [--sem-sheets] [--dir PASTA_EXTRATOS]

  --sem-sheets    Roda as verificações mas não atualiza o Google Sheets
                  (útil para testar sem credenciais configuradas)
  --dir PASTA     Pasta com os extratos (padrão: data/extratos/)
"""

from __future__ import annotations
import sys
import argparse
from pathlib import Path

# Garante que o diretório raiz está no path
sys.path.insert(0, str(Path(__file__).parent))

from colorama import Fore, Style, init
init(autoreset=True)


def main() -> None:
    parser = argparse.ArgumentParser(description="Controle Financeiro Anti-Desvio")
    parser.add_argument("--sem-sheets", action="store_true",
                        help="Pula a atualização do Google Sheets")
    parser.add_argument("--dir", type=str, default=None,
                        help="Pasta com os arquivos de extrato")
    args = parser.parse_args()

    from config import EXTRATOS_DIR, GOOGLE_CREDENTIALS_FILE

    diretorio = Path(args.dir) if args.dir else EXTRATOS_DIR

    print(f"\n{Style.BRIGHT}{'═'*60}")
    print("  CONTROLE FINANCEIRO ANTI-DESVIO — MOINHO DE TRIGO")
    print(f"{'═'*60}{Style.RESET_ALL}\n")

    # ── 1. Leitura e consolidação dos extratos ───────────────────────────
    print(f"{Style.BRIGHT}[1/4] Lendo extratos em: {diretorio}{Style.RESET_ALL}")
    from src.consolidator import consolidar
    try:
        df = consolidar(diretorio=diretorio)
    except FileNotFoundError:
        print(f"{Fore.RED}ERRO: Nenhum arquivo encontrado em {diretorio}")
        print(f"Gere os arquivos de exemplo com:{Style.RESET_ALL}")
        print("  python data/exemplos/gerar_exemplos.py\n")
        sys.exit(1)
    except Exception as e:
        print(f"{Fore.RED}ERRO ao ler extratos: {e}{Style.RESET_ALL}")
        sys.exit(1)

    print(f"  {Fore.GREEN}✔ {len(df)} lançamentos consolidados de {df['banco'].nunique()} banco(s){Style.RESET_ALL}")

    # ── 2. Verificações de desvio ────────────────────────────────────────
    print(f"\n{Style.BRIGHT}[2/4] Rodando verificações de desvio...{Style.RESET_ALL}")
    from src.detector import rodar_verificacoes
    df_com_alertas, alertas = rodar_verificacoes(df)

    criticos = sum(1 for a in alertas if a["nivel"] == "CRITICO")
    atencoes = sum(1 for a in alertas if a["nivel"] == "ATENCAO")

    if criticos:
        print(f"  {Fore.RED}⚠ {criticos} alerta(s) CRÍTICO(s){Style.RESET_ALL}")
    if atencoes:
        print(f"  {Fore.YELLOW}⚠ {atencoes} alerta(s) de ATENÇÃO{Style.RESET_ALL}")
    if not alertas:
        print(f"  {Fore.GREEN}✔ Nenhum desvio detectado{Style.RESET_ALL}")

    # ── 3. Atualiza base de conhecimento ─────────────────────────────────
    print(f"\n{Style.BRIGHT}[3/4] Atualizando base de fornecedores/clientes...{Style.RESET_ALL}")
    from src.knowledge_base import atualizar_base
    atualizar_base(df_com_alertas)

    # ── 4. Relatório ─────────────────────────────────────────────────────
    from src.report import imprimir_resumo
    imprimir_resumo(df_com_alertas, alertas)

    # ── 5. Google Sheets ─────────────────────────────────────────────────
    usar_sheets = not args.sem_sheets
    if usar_sheets and not GOOGLE_CREDENTIALS_FILE.exists():
        print(f"{Fore.YELLOW}[4/4] Google Sheets pulado: credenciais não encontradas em "
              f"{GOOGLE_CREDENTIALS_FILE}{Style.RESET_ALL}")
        print(f"       Rode com --sem-sheets para suprimir este aviso, ou siga o README para configurar.\n")
        usar_sheets = False

    if usar_sheets:
        print(f"{Style.BRIGHT}[4/4] Atualizando Google Sheets...{Style.RESET_ALL}")
        try:
            from src.sheets import atualizar_sheets
            url = atualizar_sheets(df_com_alertas, alertas)
            print(f"  {Fore.GREEN}✔ Planilha disponível em: {url}{Style.RESET_ALL}\n")
        except Exception as e:
            print(f"  {Fore.RED}ERRO no Google Sheets: {e}{Style.RESET_ALL}")
            print(f"  Verifique as credenciais e o ID da planilha em config.py\n")
    else:
        print(f"[4/4] Google Sheets pulado (--sem-sheets).\n")

    # ── Salva extrato consolidado localmente ──────────────────────────────
    saida = Path("data") / "extrato_consolidado.xlsx"
    df_com_alertas.to_excel(saida, index=False)
    print(f"  Extrato salvo localmente em: {saida}\n")


if __name__ == "__main__":
    main()
