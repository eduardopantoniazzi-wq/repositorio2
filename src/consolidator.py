"""
Consolida extratos de múltiplos bancos em um único DataFrame padronizado.
Detecta automaticamente o banco a partir do nome do arquivo.
"""

from __future__ import annotations
from pathlib import Path
from typing import Optional
import re

import pandas as pd

from .readers.bradesco import LeitorBradesco
from .readers.sicredi import LeitorSicredi
from .readers.bb import LeitorBB


_LEITORES = {
    "bradesco": LeitorBradesco,
    "sicredi":  LeitorSicredi,
    "bb":       LeitorBB,
    "bancodobrasil": LeitorBB,
    "banco_do_brasil": LeitorBB,
}

_SUFIXOS_ACEITOS = {".pdf", ".csv", ".xlsx", ".xls"}


def _detectar_banco(caminho: Path) -> Optional[str]:
    nome = caminho.stem.lower()
    nome = re.sub(r"[^a-z0-9]", "", nome)
    for chave in _LEITORES:
        if chave in nome:
            return chave
    return None


def consolidar(
    diretorio: Optional[Path] = None,
    arquivos: Optional[list[Path]] = None,
) -> pd.DataFrame:
    """
    Lê todos os extratos encontrados e retorna DataFrame consolidado.
    Passa `diretorio` para varrer uma pasta, ou `arquivos` para lista explícita.
    """
    if arquivos is None:
        if diretorio is None:
            raise ValueError("Informe `diretorio` ou `arquivos`.")
        arquivos = [
            p for p in Path(diretorio).iterdir()
            if p.suffix.lower() in _SUFIXOS_ACEITOS
        ]

    if not arquivos:
        raise FileNotFoundError(f"Nenhum arquivo de extrato encontrado em {diretorio}")

    frames = []
    erros = []

    for arq in arquivos:
        chave_banco = _detectar_banco(arq)
        if chave_banco is None:
            # Tenta por conteúdo — força leitura com cada leitor
            for chave, Leitor in _LEITORES.items():
                try:
                    df = Leitor(arq).ler()
                    frames.append(df)
                    print(f"  ✔ {arq.name} → {Leitor.BANCO} (detecção por conteúdo)")
                    break
                except Exception:
                    continue
            else:
                erros.append(arq.name)
        else:
            Leitor = _LEITORES[chave_banco]
            try:
                df = Leitor(arq).ler()
                frames.append(df)
                print(f"  ✔ {arq.name} → {Leitor.BANCO}")
            except Exception as e:
                erros.append(f"{arq.name} ({e})")

    if erros:
        print(f"\n  ⚠ Arquivos ignorados: {', '.join(erros)}")

    if not frames:
        raise ValueError("Nenhum extrato foi lido com sucesso.")

    df_total = pd.concat(frames, ignore_index=True)
    df_total = df_total.sort_values(["data", "banco"]).reset_index(drop=True)
    return df_total
