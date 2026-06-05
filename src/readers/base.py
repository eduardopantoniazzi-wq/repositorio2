"""
Classe base para leitura de extratos bancários.
Cada banco herda desta classe e implementa _parse().
"""

from __future__ import annotations
import re
from pathlib import Path
from typing import Union

import pandas as pd


COLUNAS_PADRAO = ["data", "banco", "descricao", "credito", "debito", "saldo", "horario", "documento"]


def _limpar_valor(v) -> float:
    """Converte strings como '1.234,56' ou '-1234.56' para float."""
    if pd.isna(v) or v == "" or v is None:
        return 0.0
    s = str(v).strip().replace("\xa0", "").replace(" ", "")
    # Remove símbolo de moeda
    s = re.sub(r"[R$]", "", s)
    # Formato brasileiro: 1.234,56
    if re.search(r"\d\.\d{3},", s):
        s = s.replace(".", "").replace(",", ".")
    # Formato com vírgula como decimal: 1234,56
    elif "," in s and "." not in s:
        s = s.replace(",", ".")
    # Remove pontos remanescentes que sejam separadores de milhar
    # (ex: 1.234.567 → 1234567)
    elif s.count(".") > 1:
        s = s.replace(".", "")
    s = s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return 0.0


def _ler_pdf_tabelas(caminho: Path) -> list[pd.DataFrame]:
    """Extrai todas as tabelas de um PDF como lista de DataFrames."""
    try:
        import pdfplumber
    except ImportError:
        raise ImportError(
            "pdfplumber não disponível. Instale com: pip install pdfplumber\n"
            "Para arquivos CSV/Excel o pdfplumber não é necessário."
        )
    tabelas = []
    with pdfplumber.open(caminho) as pdf:
        for page in pdf.pages:
            for tabela in page.extract_tables():
                if tabela:
                    df = pd.DataFrame(tabela[1:], columns=tabela[0])
                    tabelas.append(df)
    return tabelas


class LeitorBase:
    BANCO: str = ""

    def __init__(self, caminho: Union[str, Path]):
        self.caminho = Path(caminho)
        self.sufixo = self.caminho.suffix.lower()

    def ler(self) -> pd.DataFrame:
        if self.sufixo == ".pdf":
            df = self._parse_pdf()
        elif self.sufixo in (".xlsx", ".xls"):
            df = self._parse_excel()
        elif self.sufixo == ".csv":
            df = self._parse_csv()
        else:
            raise ValueError(f"Formato não suportado: {self.sufixo}")

        df = self._normalizar(df)
        return df

    # ── Métodos a implementar por subclasse ─────────────────────────────────
    def _parse_pdf(self) -> pd.DataFrame:
        raise NotImplementedError

    def _parse_excel(self) -> pd.DataFrame:
        raise NotImplementedError

    def _parse_csv(self) -> pd.DataFrame:
        raise NotImplementedError

    # ── Normalização comum ──────────────────────────────────────────────────
    def _normalizar(self, df: pd.DataFrame) -> pd.DataFrame:
        for col in COLUNAS_PADRAO:
            if col not in df.columns:
                df[col] = None if col in ("horario", "documento") else (0.0 if col in ("credito", "debito", "saldo") else "")

        df["banco"] = self.BANCO
        df["credito"] = df["credito"].apply(_limpar_valor)
        df["debito"]  = df["debito"].apply(_limpar_valor)
        df["saldo"]   = df["saldo"].apply(_limpar_valor)
        df["data"]    = pd.to_datetime(df["data"], dayfirst=True, errors="coerce")
        df["descricao"] = df["descricao"].astype(str).str.strip()

        # Remove linhas sem data válida ou sem movimentação
        df = df.dropna(subset=["data"])
        df = df[~((df["credito"] == 0) & (df["debito"] == 0) & (df["saldo"] == 0))]

        return df[COLUNAS_PADRAO].reset_index(drop=True)
