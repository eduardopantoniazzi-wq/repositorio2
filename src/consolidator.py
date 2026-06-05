"""
Consolida extratos dos três bancos + planilha Excel em um único DataFrame padronizado.
"""

from __future__ import annotations
from pathlib import Path
from typing import Optional

import pandas as pd

from .readers.bradesco import LeitorBradesco
from .readers.sicredi import LeitorSicredi
from .readers.bb import LeitorBB
from .readers.planilha import ler_planilha


_LEITORES = {
    "bradesco": LeitorBradesco,
    "sicredi":  LeitorSicredi,
    "bb":       LeitorBB,
    "bancodobrasil": LeitorBB,
    "banco_do_brasil": LeitorBB,
}

_SUFIXOS_BANCO = {".pdf", ".csv", ".xlsx", ".xls"}


def _detectar_banco(caminho: Path) -> Optional[str]:
    import re
    nome = re.sub(r"[^a-z0-9]", "", caminho.stem.lower())
    for chave in _LEITORES:
        if chave in nome:
            return chave
    return None


def _eh_planilha_receitas(caminho: Path) -> bool:
    """Heurística: Excel com nome contendo 'receita' ou 'despesa'."""
    nome = caminho.stem.lower()
    return any(p in nome for p in ("receita", "despesa", "fluxo", "caixa"))


def consolidar(
    diretorio: Optional[Path] = None,
    arquivos_banco: Optional[list[Path]] = None,
    arquivo_planilha: Optional[Path] = None,
    meses_planilha: Optional[list[str]] = None,
) -> pd.DataFrame:
    """
    Lê extratos bancários e/ou planilha Excel de despesas.
    Retorna DataFrame consolidado com colunas padrão.
    """
    frames: list[pd.DataFrame] = []
    erros: list[str] = []

    # ── Extratos bancários ───────────────────────────────────────────────
    if arquivos_banco is None and diretorio is not None:
        arquivos_banco = [
            p for p in Path(diretorio).iterdir()
            if p.suffix.lower() in _SUFIXOS_BANCO and not _eh_planilha_receitas(p)
        ]

    if arquivos_banco:
        for arq in arquivos_banco:
            chave = _detectar_banco(arq)
            if chave:
                Leitor = _LEITORES[chave]
                try:
                    df = Leitor(arq).ler()
                    frames.append(df)
                    print(f"  ✔ {arq.name} → {Leitor.BANCO} ({len(df)} lançamentos)")
                except Exception as e:
                    erros.append(f"{arq.name}: {e}")
            else:
                # Tenta cada leitor
                lido = False
                for chave2, Leitor in _LEITORES.items():
                    try:
                        df = Leitor(arq).ler()
                        frames.append(df)
                        print(f"  ✔ {arq.name} → {Leitor.BANCO} (auto-detectado)")
                        lido = True
                        break
                    except Exception:
                        continue
                if not lido:
                    erros.append(f"{arq.name}: banco não reconhecido")

    # ── Planilha de receitas/despesas ────────────────────────────────────
    if arquivo_planilha is None and diretorio is not None:
        candidatos = [
            p for p in Path(diretorio).iterdir()
            if p.suffix.lower() in (".xlsx", ".xls") and _eh_planilha_receitas(p)
        ]
        if candidatos:
            arquivo_planilha = candidatos[0]

    if arquivo_planilha and arquivo_planilha.exists():
        try:
            df_plan = ler_planilha(arquivo_planilha, meses=meses_planilha)
            if not df_plan.empty:
                # Garante colunas padrão
                from .readers.base import COLUNAS_PADRAO
                extra_cols = [c for c in df_plan.columns if c not in COLUNAS_PADRAO]
                df_plan_std = df_plan[COLUNAS_PADRAO + extra_cols] if extra_cols else df_plan[COLUNAS_PADRAO]
                frames.append(df_plan_std)
                print(f"  ✔ Planilha {arquivo_planilha.name} → {len(df_plan)} lançamentos")
        except Exception as e:
            erros.append(f"Planilha {arquivo_planilha.name}: {e}")

    if erros:
        for e in erros:
            print(f"  ⚠ {e}")

    if not frames:
        raise ValueError("Nenhum dado foi lido com sucesso.")

    df_total = pd.concat(frames, ignore_index=True)
    df_total = df_total.sort_values(["data", "banco"]).reset_index(drop=True)
    return df_total
